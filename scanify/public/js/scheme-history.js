/**
 * scheme-history.js - Reusable tabbed history modal for the scheme flow.
 * Exposes:
 *   openSchemeHistory({ doctor_code, stockist_code, hq, products })
 *   openSchemeHistoryFromForm()  - builds context from the scheme create form DOM
 * Requires templates/includes/scheme_history_modal.html to be included on the page.
 * All helpers are prefixed `sh` to avoid clashing with page-level functions.
 */

let shCtx = {};
// Per-tab filter state (reset whenever the modal is (re)opened).
let shSchemePeriod = 3;          // trailing months for Doctor Scheme History
let shSalesEndMonth = '';        // YYYY-MM end month for Sales History (blank = latest on record)
let shSalesMode = 'after_deduction';

function shFmt(v) {
    return new Intl.NumberFormat('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(v || 0);
}
function shInt(v) {
    return new Intl.NumberFormat('en-IN').format(Math.round(v || 0));
}
function shMoney(v) { return '₹ ' + shFmt(v); }
/** Compact number: up to 2 decimals, trailing zeros stripped. */
function shNum(v) {
    var n = Math.round((Number(v) || 0) * 100) / 100;
    return new Intl.NumberFormat('en-IN', { maximumFractionDigits: 2 }).format(n);
}
/** Same as shNum but renders an empty cell when the value is ~zero (matches the client sheet). */
function shNumBlank(v) {
    var n = Number(v) || 0;
    if (Math.abs(n) < 0.005) return '';
    return shNum(n);
}
function shEsc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
        return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
}

function shPost(method, payload) {
    return fetch('/api/method/scanify.api.' + method, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'X-Frappe-CSRF-Token': frappe.csrf_token },
        body: JSON.stringify(payload || {}),
    }).then(function (r) { return r.json(); });
}

function shDivision() {
    try {
        const btn = document.querySelector('.division-name');
        if (btn && btn.textContent.trim()) return btn.textContent.trim();
    } catch (e) { }
    const match = document.cookie.match(/(?:^|;\s*)division=([^;]*)/);
    if (match) return decodeURIComponent(match[1]);
    return '';
}

const SH_STATUS = { Approved: 'success', Pending: 'warning', Rejected: 'danger', Rerouted: 'info', Deducted: 'primary' };

// ===================== SHARED UI HELPERS =====================

function shSpinner() {
    return '<div class="text-center py-4"><i class="fa fa-spinner fa-spin fa-2x text-muted"></i></div>';
}
function shEmpty(msg) {
    return '<div class="alert alert-light text-center text-muted mb-0">' + (msg || 'No data') + '</div>';
}
/** cards: [{label, value, sub, color}] -> a clean row of stat cards */
function shStatCards(cards) {
    let html = '<div class="row sh-stats mb-3">';
    cards.forEach(function (c) {
        html += `<div class="col"><div class="sh-stat border rounded p-2 h-100">
            <div class="sh-stat-label text-muted">${c.label}</div>
            <div class="sh-stat-value ${c.color ? 'text-' + c.color : ''}">${c.value}</div>
            ${c.sub ? '<div class="sh-stat-sub text-muted">' + c.sub + '</div>' : ''}
        </div></div>`;
    });
    html += '</div>';
    return html;
}
function shBadge(status) {
    return `<span class="badge badge-${SH_STATUS[status] || 'secondary'}">${status || '-'}</span>`;
}
function shScrollTable(head, body, maxh) {
    return `<div class="table-responsive" style="max-height:${maxh || 320}px;overflow-y:auto;">
        <table class="table table-sm table-bordered table-hover mb-0">
            <thead class="bg-light">${head}</thead><tbody>${body}</tbody></table></div>`;
}
/** trend: [{month, a, b}] with column labels -> small trend table */
function shTrendTable(trend, labelA, labelB, fmtA, fmtB) {
    if (!trend || !trend.length) return '';
    let body = '';
    trend.forEach(function (t) {
        body += `<tr><td>${t.month || '-'}</td><td class="text-right">${fmtA(t.a)}</td><td class="text-right">${fmtB(t.b)}</td></tr>`;
    });
    return '<h6 class="mt-3 text-muted"><i class="fa fa-chart-line"></i> Monthly Trend</h6>' +
        shScrollTable(`<tr><th>Month</th><th class="text-right">${labelA}</th><th class="text-right">${labelB}</th></tr>`, body, 200);
}

// ===================== ENTRY POINTS =====================

function openSchemeHistory(ctx) {
    shCtx = ctx || {};
    shSchemePeriod = 3;
    shSalesEndMonth = '';
    shSalesMode = 'after_deduction';
    ['sh-scheme', 'sh-secondary', 'sh-product', 'sh-doctor'].forEach(function (id) {
        document.getElementById(id).innerHTML = shSpinner();
    });
    $('#schemeHistoryModal').modal('show');

    shLoadSchemeTab();
    shLoadSalesHistoryTab();
    shLoadDoctorTab();
    shLoadProductTab();
}

/** Build context from the shared scheme creation form (scheme-new / scheme-repeat). */
function openSchemeHistoryFromForm() {
    const doctor_code = (document.getElementById('doctorCode') || {}).value || '';
    const stockist_code = (document.getElementById('stockistSelect') || {}).value || '';
    const hq = (document.getElementById('hqSelect') || {}).value || '';
    const products = [];
    document.querySelectorAll('#itemsTbody select.product-select').forEach(function (sel) {
        if (sel.value) {
            const opt = sel.options[sel.selectedIndex];
            products.push({ product_code: sel.value, product_name: opt ? opt.getAttribute('data-name') || opt.textContent.trim() : sel.value });
        }
    });
    openSchemeHistory({ doctor_code: doctor_code, stockist_code: stockist_code, hq: hq, products: products });
}

// ===================== SCHEME HISTORY (Doctor Scheme History — client Format 1) =====================
// Product-line detail of the doctor's scheme requests:
//   App Date | P Code | PTS | Qty (bx/units) | Free (bx/units) | Spl Rate | Value

function shSetSchemePeriod(v) { shSchemePeriod = v; shLoadSchemeTab(); }

function shPeriodFilter() {
    var opts = [['3', 'Last 3 Months'], ['6', 'Last 6 Months'], ['12', 'Last 12 Months'], ['0', 'All']];
    var sel = '';
    opts.forEach(function (o) {
        sel += '<option value="' + o[0] + '"' + (String(shSchemePeriod) === o[0] ? ' selected' : '') + '>' + o[1] + '</option>';
    });
    return '<div class="d-flex align-items-center mb-3" style="gap:8px;">'
        + '<span class="text-muted" style="font-size:.8rem;font-weight:600;">Period</span>'
        + '<select class="form-control form-control-sm" style="width:auto;" onchange="shSetSchemePeriod(this.value)">' + sel + '</select>'
        + '</div>';
}

function shLoadSchemeTab() {
    const pane = document.getElementById('sh-scheme');
    if (!shCtx.doctor_code) {
        pane.innerHTML = shEmpty('Select a doctor to see scheme history');
        return;
    }
    pane.innerHTML = shSpinner();
    shPost('get_doctor_scheme_history', {
        doctor_code: shCtx.doctor_code, hq: shCtx.hq, division: shDivision(), period_months: shSchemePeriod
    }).then(function (r) {
        const d = (r.message && r.message.success) ? r.message : null;
        const rows = d ? (d.rows || []) : [];
        const t = d ? (d.totals || {}) : {};
        const header = '<div class="mb-2"><strong><i class="fa fa-user-md text-muted"></i> Doctor Name: </strong>'
            + '<span class="text-danger font-weight-bold">' + shEsc(d && d.doctor_name ? d.doctor_name : '-') + '</span></div>';

        if (!rows.length) { pane.innerHTML = header + shPeriodFilter() + shEmpty('No scheme history for this period'); return; }

        let body = '';
        rows.forEach(function (s) {
            body += `<tr>
                <td>${shEsc(s.application_date) || '-'}</td>
                <td><strong>${shEsc(s.product_code) || '-'}</strong></td>
                <td class="text-right">${shNum(s.pts)}</td>
                <td class="text-right">${shNumBlank(s.quantity)}</td>
                <td class="text-right">${shNumBlank(s.free_quantity)}</td>
                <td class="text-right">${shNumBlank(s.special_rate)}</td>
                <td class="text-right font-weight-bold">${shFmt(s.value)}</td>
            </tr>`;
        });

        pane.innerHTML = header + shPeriodFilter()
            + shStatCards([
                { label: 'Scheme Requests', value: t.requests || 0 },
                { label: 'Total Lines', value: t.lines || 0 },
                { label: 'Total Free Qty', value: shNum(t.free_qty) },
                { label: 'Total Value', value: shMoney(t.value), color: 'primary' },
            ])
            + shScrollTable(
                '<tr><th>App Date</th><th>P Code</th><th class="text-right">PTS</th>'
                + '<th class="text-right">Qty<br>bx/units</th><th class="text-right">Free<br>bx/units</th>'
                + '<th class="text-right">Spl rate</th><th class="text-right">Value</th></tr>',
                body, 400)
            + '<small class="text-muted d-block mt-2">Value = Qty &times; (Special Rate or PTS). Free goods are not valued.</small>';
    }).catch(function () { pane.innerHTML = shEmpty('Error loading scheme history'); });
}

// ===================== SALES HISTORY — PAST 3 MONTHS (HQ, client Format 2) =====================
// Product x month secondary-sales matrix (box qty) + latest closing, with an
// H.Q Value header row in Rs. Lakhs.

function shSetSalesEndMonth(v) { shSalesEndMonth = v; shLoadSalesHistoryTab(); }
function shSetSalesMode(v) { shSalesMode = v; shLoadSalesHistoryTab(); }

function shSalesFilter(resolvedEnd) {
    return '<div class="d-flex align-items-center flex-wrap mb-3" style="gap:14px;">'
        + '<div class="d-flex align-items-center" style="gap:8px;">'
        + '<span class="text-muted" style="font-size:.8rem;font-weight:600;">Up to Month</span>'
        + '<input type="month" class="form-control form-control-sm" style="width:auto;" value="' + shEsc(resolvedEnd || shSalesEndMonth) + '" onchange="shSetSalesEndMonth(this.value)">'
        + '</div>'
        + '<div class="d-flex align-items-center" style="gap:8px;">'
        + '<span class="text-muted" style="font-size:.8rem;font-weight:600;">Sales</span>'
        + '<select class="form-control form-control-sm" style="width:auto;" onchange="shSetSalesMode(this.value)">'
        + '<option value="after_deduction"' + (shSalesMode === 'after_deduction' ? ' selected' : '') + '>After Deduction</option>'
        + '<option value="before_deduction"' + (shSalesMode === 'before_deduction' ? ' selected' : '') + '>Before Deduction</option>'
        + '</select></div></div>';
}

function shLoadSalesHistoryTab() {
    const pane = document.getElementById('sh-secondary');
    if (!shCtx.hq) { pane.innerHTML = shEmpty('An HQ is required to show sales history'); return; }
    pane.innerHTML = shSpinner();
    shPost('get_hq_sales_history_3m', {
        hq: shCtx.hq, division: shDivision(), end_month: shSalesEndMonth, sales_mode: shSalesMode, months: 3
    }).then(function (r) {
        const d = (r.message && r.message.success) ? r.message : null;
        const months = d ? (d.months || []) : [];
        const products = d ? (d.products || []) : [];
        const hqVal = d ? (d.hq_value || {}) : {};

        const header = '<div class="mb-2"><strong><i class="fa fa-hospital text-muted"></i> H.Q : </strong>'
            + '<span class="text-primary font-weight-bold">' + shEsc(d && d.hq_name ? d.hq_name : '-') + '</span>'
            + (d && d.period_label ? ' <span class="text-muted">· Past 3 Months (' + shEsc(d.period_label) + ')</span>' : '') + '</div>';

        if (!months.length || !products.length) { pane.innerHTML = header + shSalesFilter(shSalesEndMonth) + shEmpty('No sales history for this HQ / period'); return; }

        const resolvedEnd = shSalesEndMonth || months[months.length - 1].key;
        const lastLbl = months[months.length - 1].label;
        let head = '<tr><th>P Code</th>';
        months.forEach(function (m) { head += '<th class="text-right">' + shEsc(m.label) + '<br>Sales</th>'; });
        head += '<th class="text-right">' + shEsc(lastLbl) + '<br>Closing</th></tr>';

        // H.Q Value row (Rs. Lakhs, fixed 2 decimals) — pinned first, matching the client sheet.
        let hqRow = '<tr style="background:#eef2ff;font-weight:700;">'
            + '<td>H.Q Value</td>';
        (hqVal.monthly || []).forEach(function (v) { hqRow += '<td class="text-right">' + shFmt(v) + '</td>'; });
        hqRow += '<td class="text-right">' + shFmt(hqVal.closing) + '</td></tr>';

        let body = hqRow;
        products.forEach(function (p) {
            body += '<tr><td><strong>' + shEsc(p.product_code) + '</strong></td>';
            (p.monthly || []).forEach(function (v) { body += '<td class="text-right">' + shNumBlank(v) + '</td>'; });
            body += '<td class="text-right">' + shNum(p.closing) + '</td></tr>';
        });

        pane.innerHTML = header + shSalesFilter(resolvedEnd)
            + shScrollTable(head, body, 420)
            + '<small class="text-muted d-block mt-2">H.Q Value row in <strong>&#8377; Lakhs</strong>; product rows are secondary sales in <strong>boxes</strong>. '
            + 'Closing = ' + shEsc(lastLbl) + ' closing stock.</small>';
    }).catch(function () { pane.innerHTML = shEmpty('Error loading sales history'); });
}

// ===================== DOCTOR HISTORY =====================

function shLoadDoctorTab() {
    const pane = document.getElementById('sh-doctor');
    if (!shCtx.doctor_code) { pane.innerHTML = shEmpty('Select a doctor to see doctor history'); return; }
    shPost('get_doctor_history_for_scheme', { doctor_code: shCtx.doctor_code, hq: shCtx.hq })
        .then(function (r) {
            const d = r.message;
            if (!d || !d.success) { pane.innerHTML = shEmpty('No doctor history'); return; }

            let html = `<div class="mb-2"><strong>${d.doctor_name || '-'}</strong>
                <span class="text-muted">${d.place ? '· ' + d.place : ''} ${d.specialization ? '· ' + d.specialization : ''} ${d.hq ? '· HQ ' + d.hq : ''}</span></div>`;

            html += shStatCards([
                { label: 'Total Schemes', value: d.total_schemes || 0 },
                { label: 'Approved', value: d.total_approved || 0, color: 'success' },
                { label: 'Pending', value: d.total_pending || 0, color: 'warning' },
                { label: 'Rejected', value: d.total_rejected || 0, color: 'danger' },
                { label: 'Total Value', value: shMoney(d.total_value), sub: d.last_scheme_date ? 'Last: ' + d.last_scheme_date : '' },
            ]);

            // Top products
            let ps = '';
            (d.product_summary || []).forEach(function (p) {
                ps += `<tr><td><strong>${p.product_name || '-'}</strong><br><small class="text-muted">${p.product_code}</small></td>
                    <td class="text-right">${shInt(p.total_quantity)}</td>
                    <td class="text-right">${shInt(p.total_free_quantity)}</td>
                    <td class="text-right">${shMoney(p.total_value)}</td>
                    <td class="text-right">${p.scheme_count || 0}</td></tr>`;
            });
            if (ps) {
                html += '<h6 class="mt-3 text-muted"><i class="fa fa-cubes"></i> Top Products (Approved)</h6>' +
                    shScrollTable('<tr><th>Product</th><th class="text-right">Qty</th><th class="text-right">Free</th><th class="text-right">Value</th><th class="text-right">Schemes</th></tr>', ps, 200);
            }

            // Monthly trend
            html += shTrendTable(
                (d.chart_data || []).map(function (t) { return { month: t.month, a: t.scheme_count, b: t.total_value }; }),
                'Schemes', 'Value', shInt, shMoney);

            // Recent requests
            let rs = '';
            (d.recent_schemes || []).forEach(function (s) {
                rs += `<tr>
                    <td><a href="/portal/scheme-detail?name=${s.name}" target="_blank">${s.name}</a></td>
                    <td>${s.application_date || '-'}</td>
                    <td>${s.stockist_name || '-'}</td>
                    <td class="text-right">${s.product_count || 0}</td>
                    <td class="text-right">${shMoney(s.total_scheme_value)}</td>
                    <td>${shBadge(s.approval_status)}</td></tr>`;
            });
            if (!rs) rs = '<tr><td colspan="6" class="text-center text-muted">No recent schemes</td></tr>';
            html += '<h6 class="mt-3 text-muted"><i class="fa fa-list"></i> Recent Requests</h6>' +
                shScrollTable('<tr><th>Scheme</th><th>Date</th><th>Stockist</th><th class="text-right">Products</th><th class="text-right">Value</th><th>Status</th></tr>', rs, 240);

            pane.innerHTML = html;
        }).catch(function () { pane.innerHTML = shEmpty('Error loading doctor history'); });
}

// ===================== PRODUCT HISTORY =====================

function shLoadProductTab() {
    const pane = document.getElementById('sh-product');
    const products = shCtx.products || [];
    if (!products.length) { pane.innerHTML = shEmpty('Add a product to see product history'); return; }

    let picker = '';
    if (products.length > 1) {
        picker = '<div class="form-group"><label class="text-muted">Product</label><select class="form-control form-control-sm" id="sh-product-picker" onchange="shLoadProductDetail(this.value)">';
        products.forEach(function (p) { picker += `<option value="${p.product_code}">${p.product_name || p.product_code}</option>`; });
        picker += '</select></div>';
    }
    pane.innerHTML = picker + '<div id="sh-product-detail">' + shSpinner() + '</div>';
    shLoadProductDetail(products[0].product_code);
}

function shLoadProductDetail(productCode) {
    const box = document.getElementById('sh-product-detail');
    if (!box) return;
    box.innerHTML = shSpinner();
    shPost('get_product_history_for_scheme', { product_code: productCode, doctor_code: shCtx.doctor_code, hq: shCtx.hq })
        .then(function (r) {
            const d = r.message;
            if (!d || !d.success) { box.innerHTML = shEmpty('No product history'); return; }

            let html = `<div class="mb-2"><strong>${d.product_name || '-'}</strong>
                <span class="text-muted">${d.pack ? '· ' + d.pack : ''} · PTS ${shMoney(d.pts)}</span></div>`;

            html += shStatCards([
                { label: 'Total Schemes', value: d.total_schemes || 0 },
                { label: 'Total Quantity', value: shInt(d.total_quantity) },
                { label: 'Total Value', value: shMoney(d.total_value) },
                { label: 'Last Ordered', value: d.last_order_date || 'Never' },
            ]);

            // Monthly trend
            html += shTrendTable(
                (d.chart_data || []).map(function (t) { return { month: t.month, a: t.quantity, b: t.value }; }),
                'Qty', 'Value', shInt, shMoney);

            // Recent schemes
            let rs = '';
            (d.recent_schemes || []).forEach(function (s) {
                rs += `<tr>
                    <td><a href="/portal/scheme-detail?name=${s.name}" target="_blank">${s.name}</a></td>
                    <td>${s.application_date || '-'}</td>
                    <td>${s.doctor_name || '-'}</td>
                    <td class="text-right">${shInt(s.quantity)}</td>
                    <td class="text-right">${shMoney(s.product_value)}</td>
                    <td>${shBadge(s.approval_status)}</td></tr>`;
            });
            if (!rs) rs = '<tr><td colspan="6" class="text-center text-muted">No recent schemes</td></tr>';
            html += '<h6 class="mt-3 text-muted"><i class="fa fa-list"></i> Recent Scheme Requests</h6>' +
                shScrollTable('<tr><th>Scheme</th><th>Date</th><th>Doctor</th><th class="text-right">Qty</th><th class="text-right">Value</th><th>Status</th></tr>', rs, 240);

            box.innerHTML = html;
        }).catch(function () { box.innerHTML = shEmpty('Error loading product history'); });
}

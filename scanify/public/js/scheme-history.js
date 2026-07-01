/**
 * scheme-history.js - Reusable tabbed history modal for the scheme flow.
 * Exposes:
 *   openSchemeHistory({ doctor_code, stockist_code, hq, products })
 *   openSchemeHistoryFromForm()  - builds context from the scheme create form DOM
 * Requires templates/includes/scheme_history_modal.html to be included on the page.
 * All helpers are prefixed `sh` to avoid clashing with page-level functions.
 */

let shCtx = {};

function shFmt(v) {
    return new Intl.NumberFormat('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(v || 0);
}
function shInt(v) {
    return new Intl.NumberFormat('en-IN').format(Math.round(v || 0));
}
function shMoney(v) { return '₹ ' + shFmt(v); }

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
    ['sh-scheme', 'sh-secondary', 'sh-product', 'sh-doctor'].forEach(function (id) {
        document.getElementById(id).innerHTML = shSpinner();
    });
    $('#schemeHistoryModal').modal('show');

    shLoadSchemeTab();
    shLoadSecondaryTab();
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

// ===================== SCHEME HISTORY =====================

function shLoadSchemeTab() {
    const pane = document.getElementById('sh-scheme');
    if (!shCtx.doctor_code && !shCtx.stockist_code) {
        pane.innerHTML = shEmpty('Select a doctor or stockist to see scheme history');
        return;
    }
    shPost('get_scheme_history_portal', {
        doctor_code: shCtx.doctor_code, stockist_code: shCtx.stockist_code, hq: shCtx.hq, division: shDivision()
    }).then(function (r) {
        const rows = (r.message && r.message.success) ? (r.message.data || []) : [];
        if (!rows.length) { pane.innerHTML = shEmpty('No scheme history'); return; }

        const totalValue = rows.reduce(function (a, s) { return a + (s.total_scheme_value || 0); }, 0);
        const totalFree = rows.reduce(function (a, s) { return a + (s.total_free_qty || 0); }, 0);
        const approved = rows.filter(function (s) { return s.approval_status === 'Approved'; }).length;
        const deducted = rows.filter(function (s) { return s.approval_status === 'Deducted'; }).length;

        let body = '';
        rows.forEach(function (s) {
            body += `<tr>
                <td><a href="/portal/scheme-detail?name=${s.name}" target="_blank">${s.name}</a>${s.repeated_request ? ' <span class="badge badge-light">R</span>' : ''}</td>
                <td>${s.application_date || '-'}</td>
                <td>${s.doctor_name || '-'}</td>
                <td>${s.stockist_name || '-'}</td>
                <td>${s.hq_name || '-'}</td>
                <td class="text-right">${s.product_count || 0}</td>
                <td class="text-right">${shInt(s.total_free_qty)}</td>
                <td class="text-right">${shMoney(s.total_scheme_value)}</td>
                <td>${shBadge(s.approval_status)}</td>
            </tr>`;
        });

        pane.innerHTML =
            shStatCards([
                { label: 'Total Schemes', value: rows.length },
                { label: 'Approved', value: approved, color: 'success' },
                { label: 'Deducted', value: deducted, color: 'primary' },
                { label: 'Total Free Qty', value: shInt(totalFree) },
                { label: 'Total Order Value', value: shMoney(totalValue) },
            ]) +
            shScrollTable(
                '<tr><th>Scheme</th><th>Date</th><th>Doctor</th><th>Stockist</th><th>HQ</th><th class="text-right">Products</th><th class="text-right">Free Qty</th><th class="text-right">Order Value</th><th>Status</th></tr>',
                body, 380);
    }).catch(function () { pane.innerHTML = shEmpty('Error loading scheme history'); });
}

// ===================== SECONDARY SALES (statement level) =====================

function shLoadSecondaryTab() {
    const pane = document.getElementById('sh-secondary');
    if (!shCtx.stockist_code) { pane.innerHTML = shEmpty('Select a stockist to see secondary sales history'); return; }
    shPost('get_stockist_statement_history', { stockist_code: shCtx.stockist_code, division: shDivision() })
        .then(function (r) {
            const rows = (r.message && r.message.success) ? (r.message.data || []) : [];
            if (!rows.length) { pane.innerHTML = shEmpty('No statements found for this stockist'); return; }

            const latest = rows[0];
            const avgSales = rows.reduce(function (a, s) { return a + (s.total_sales_value_pts || 0); }, 0) / rows.length;

            let body = '';
            rows.forEach(function (s) {
                body += `<tr>
                    <td><strong>${s.month_label || '-'}</strong></td>
                    <td class="text-right">${shInt(s.sales_qty)}</td>
                    <td class="text-right">${shInt(s.free_qty)}</td>
                    <td class="text-right">${shInt(s.free_qty_scheme)}</td>
                    <td class="text-right">${shMoney(s.total_opening_value)}</td>
                    <td class="text-right">${shMoney(s.total_purchase_value)}</td>
                    <td class="text-right">${shMoney(s.total_sales_value_pts)}</td>
                    <td class="text-right">${shMoney(s.total_free_value)}</td>
                    <td class="text-right">${shMoney(s.total_closing_value)}</td>
                    <td class="text-center"><a href="/portal/statement-view?name=${encodeURIComponent(s.name)}" target="_blank" title="Open statement"><i class="fa fa-external-link-alt"></i></a></td>
                </tr>`;
            });

            pane.innerHTML =
                shStatCards([
                    { label: 'Statements', value: rows.length },
                    { label: 'Latest Month', value: latest.month_label || '-' },
                    { label: 'Latest Closing', value: shMoney(latest.total_closing_value), color: 'primary' },
                    { label: 'Avg Sales (PTS)', value: shMoney(avgSales) },
                ]) +
                shScrollTable(
                    '<tr><th>Month</th><th class="text-right">Sales Qty</th><th class="text-right">Free Qty</th><th class="text-right">Scheme Free</th><th class="text-right">Opening</th><th class="text-right">Purchase</th><th class="text-right">Sales (PTS)</th><th class="text-right">Free Val</th><th class="text-right">Closing</th><th></th></tr>',
                    body, 380) +
                '<small class="text-muted d-block mt-2">Scheme Free = free goods added by scheme deductions. Values in ₹.</small>';
        }).catch(function () { pane.innerHTML = shEmpty('Error loading secondary sales'); });
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

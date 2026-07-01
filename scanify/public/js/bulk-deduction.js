/**
 * bulk-deduction.js - Filter-driven bulk scheme deduction.
 * Loads approved schemes for a scheme month, each mapped to its stockist's
 * statement for the target statement month, and deducts the selected ones.
 * Depends on ALL_REGIONS / ALL_TEAMS / ALL_HQS injected by the page template.
 */

let candidates = [];   // full candidate rows from backend

function getActiveDivision() {
    try {
        const btn = document.querySelector('.division-name');
        if (btn && btn.textContent.trim()) return btn.textContent.trim();
    } catch (e) { }
    const match = document.cookie.match(/(?:^|;\s*)division=([^;]*)/);
    if (match) return decodeURIComponent(match[1]);
    return document.getElementById('divisionLabel') ? document.getElementById('divisionLabel').textContent.trim() : 'Prima';
}

$(document).ready(function () {
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('f-deduction-date').value = today;
    // Default both month pickers to the current month
    const ym = today.slice(0, 7);
    document.getElementById('f-scheme-month').value = ym;
    document.getElementById('f-statement-month').value = ym;
    cascadeFromZone();
});

// ===================== CASCADING FILTERS =====================

function cascadeFromZone() {
    const zone = document.getElementById('f-zone').value;
    const regionSelect = document.getElementById('f-region');
    const current = regionSelect.value;
    regionSelect.innerHTML = '<option value="">All Regions</option>';
    ALL_REGIONS.forEach(function (r) {
        if (!zone || r.zone === zone) {
            const opt = document.createElement('option');
            opt.value = r.name;
            opt.textContent = r.region_name || r.name;
            regionSelect.appendChild(opt);
        }
    });
    if (current) regionSelect.value = current;
    cascadeFromRegion();
}

function cascadeFromRegion() {
    const region = document.getElementById('f-region').value;
    const teamSelect = document.getElementById('f-team');
    const current = teamSelect.value;
    teamSelect.innerHTML = '<option value="">All Teams</option>';
    ALL_TEAMS.forEach(function (t) {
        if (!region || t.region === region) {
            const opt = document.createElement('option');
            opt.value = t.name;
            opt.textContent = t.team_name || t.name;
            teamSelect.appendChild(opt);
        }
    });
    if (current) teamSelect.value = current;
    cascadeFromTeam();
}

function cascadeFromTeam() {
    const region = document.getElementById('f-region').value;
    const team = document.getElementById('f-team').value;
    const hqSelect = document.getElementById('f-hq');
    const current = hqSelect.value;
    // Teams that belong to the chosen region (when no explicit team picked)
    const regionTeams = new Set();
    if (region) ALL_TEAMS.forEach(function (t) { if (t.region === region) regionTeams.add(t.name); });
    hqSelect.innerHTML = '<option value="">All HQs</option>';
    ALL_HQS.forEach(function (h) {
        let show = true;
        if (team) show = h.team === team;
        else if (region) show = regionTeams.has(h.team);
        if (show) {
            const opt = document.createElement('option');
            opt.value = h.name;
            opt.textContent = h.hq_name || h.name;
            hqSelect.appendChild(opt);
        }
    });
    if (current) hqSelect.value = current;
}

// ===================== LOAD CANDIDATES =====================

function loadCandidates() {
    const schemeMonth = document.getElementById('f-scheme-month').value;
    const statementMonth = document.getElementById('f-statement-month').value;
    if (!schemeMonth || !statementMonth) {
        showAlert('Please choose both a scheme month and a statement month', 'warning');
        return;
    }

    $('#resultsCard').show();
    $('#loading').show();
    $('#candTable').hide();
    $('#emptyState').hide();
    $('#deductBtn').prop('disabled', true);

    $.ajax({
        url: '/api/method/scanify.api.get_bulk_deduction_candidates',
        type: 'POST',
        contentType: 'application/json',
        headers: { 'X-Frappe-CSRF-Token': frappe.csrf_token },
        data: JSON.stringify({
            division: getActiveDivision(),
            zone: document.getElementById('f-zone').value,
            region: document.getElementById('f-region').value,
            team: document.getElementById('f-team').value,
            hq: document.getElementById('f-hq').value,
            scheme_month: schemeMonth,
            statement_month: statementMonth,
        }),
        success: function (r) {
            $('#loading').hide();
            if (!(r.message && r.message.success)) {
                showAlert((r.message && r.message.message) || 'Failed to load schemes', 'danger');
                return;
            }
            candidates = r.message.data || [];
            renderCandidates();
        },
        error: function (xhr) {
            $('#loading').hide();
            console.error(xhr.responseText);
            showAlert('Error loading schemes', 'danger');
        }
    });
}

function renderCandidates() {
    $('#countBadge').text(candidates.length);
    if (!candidates.length) {
        $('#candTable').hide();
        $('#emptyState').show();
        updateSelCount();
        return;
    }

    let html = '';
    candidates.forEach(function (c, idx) {
        const disabled = !c.can_deduct;
        const stmtCell = c.can_deduct
            ? '<span class="text-success">Matched</span>'
            : `<span class="text-muted" title="${c.reason || ''}">${c.reason || 'Not available'}</span>`;
        html += `<tr class="${disabled ? 'disabled-row' : ''}">
            <td><input type="checkbox" class="cand-check" data-idx="${idx}" ${disabled ? 'disabled' : 'checked'} onchange="updateSelCount()"></td>
            <td class="cand-view" onclick="openDetail(${idx})"><strong>${c.scheme_request}</strong>${c.has_discount ? ' <span class="badge badge-warning">Discount</span>' : ''}</td>
            <td>${c.doctor_name || '-'}</td>
            <td>${c.stockist_name || '-'}</td>
            <td>${c.hq_name || '-'}</td>
            <td>${c.application_date || '-'}</td>
            <td class="text-right">${c.total_free_qty || 0}</td>
            <td class="text-right">&#8377; ${formatCurrency(c.total_value || 0)}</td>
            <td>${stmtCell}</td>
            <td><button class="btn btn-sm btn-outline-secondary" onclick="openDetail(${idx})"><i class="fa fa-eye"></i></button></td>
        </tr>`;
    });
    $('#candBody').html(html);
    $('#candTable').show();
    $('#emptyState').hide();
    updateSelCount();
}

function toggleSelectAll(checked) {
    $('.cand-check:not(:disabled)').prop('checked', checked);
    updateSelCount();
}

function updateSelCount() {
    const n = $('.cand-check:checked').length;
    $('#selCount').text(n);
    $('#deductBtn').prop('disabled', n === 0);
    const selectable = $('.cand-check:not(:disabled)').length;
    $('#headCheck').prop('checked', selectable > 0 && n === selectable);
}

// ===================== DETAIL POPUP =====================

function openDetail(idx) {
    const c = candidates[idx];
    let rows = '';
    (c.items || []).forEach(function (it) {
        rows += `<tr>
            <td><strong>${it.product_name || it.product_code}</strong><br><small class="text-muted">${it.product_code || ''}</small></td>
            <td>${it.pack || '-'}</td>
            <td class="text-right">${it.scheme_free_qty || 0}</td>
            <td class="text-right">${it.current_sales_qty || 0}</td>
            <td class="text-right">${it.deduct_qty || 0}</td>
            <td class="text-right">&#8377; ${formatCurrency(it.pts || 0)}</td>
            <td class="text-center">${(it.special_rate || 0) > 0
                ? `<span class="badge badge-warning">&#8377; ${formatCurrency(it.current_pts || it.pts || 0)} &rarr; &#8377; ${formatCurrency(it.special_rate)}</span>`
                : '<span class="text-muted">&mdash;</span>'}</td>
            <td class="text-right">&#8377; ${formatCurrency(it.deducted_value || 0)}</td>
        </tr>`;
    });
    if (!rows) rows = '<tr><td colspan="8" class="text-center text-muted">No deductible products</td></tr>';

    $('#detailBody').html(`
        <div class="mb-2">
            <strong>Scheme:</strong> ${c.scheme_request}
            <span class="mx-2">|</span><strong>Doctor:</strong> ${c.doctor_name || '-'}
            <span class="mx-2">|</span><strong>Stockist:</strong> ${c.stockist_name || '-'}
        </div>
        ${c.can_deduct ? '' : `<div class="alert alert-warning py-2">${c.reason || 'Not deductible'}</div>`}
        <div class="table-responsive">
            <table class="table table-sm table-bordered mb-0">
                <thead class="bg-light">
                    <tr><th>Product</th><th>Pack</th><th class="text-right">Scheme Free</th><th class="text-right">Stmt Sales</th><th class="text-right">Deduct</th><th class="text-right">PTS</th><th class="text-center">Discount</th><th class="text-right">Value</th></tr>
                </thead>
                <tbody>${rows}</tbody>
                <tfoot>
                    <tr class="table-active">
                        <td colspan="4" class="text-right"><strong>Totals</strong></td>
                        <td class="text-right"><strong>${c.total_free_qty || 0}</strong></td>
                        <td></td>
                        <td></td>
                        <td class="text-right"><strong>&#8377; ${formatCurrency(c.total_value || 0)}</strong></td>
                    </tr>
                </tfoot>
            </table>
        </div>
    `);
    $('#detailModal').modal('show');
}

// ===================== SUBMIT =====================

function submitBulk() {
    const deductionDate = document.getElementById('f-deduction-date').value;
    if (!deductionDate) {
        showAlert('Please choose a deduction date', 'warning');
        return;
    }

    const selected = [];
    $('.cand-check:checked').each(function () {
        const c = candidates[parseInt($(this).data('idx'), 10)];
        if (c && c.can_deduct) {
            selected.push({
                scheme_request: c.scheme_request,
                stockist_statement: c.stockist_statement,
                items: c.items,
            });
        }
    });

    if (!selected.length) {
        showAlert('No deductible schemes selected', 'warning');
        return;
    }

    const $btn = $('#deductBtn');
    $btn.prop('disabled', true).html('<i class="fa fa-spinner fa-spin"></i> Deducting...');

    $.ajax({
        url: '/api/method/scanify.api.create_bulk_scheme_deductions_portal',
        type: 'POST',
        contentType: 'application/json',
        headers: { 'X-Frappe-CSRF-Token': frappe.csrf_token },
        data: JSON.stringify({
            deductions: JSON.stringify(selected),
            deduction_date: deductionDate,
            division: getActiveDivision(),
        }),
        success: function (r) {
            $btn.prop('disabled', false).html('<i class="fa fa-check-double"></i> Deduct Selected (<span id="selCount">0</span>)');
            if (r.message && r.message.success) {
                showAlert(r.message.message || 'Deductions created', 'success');
                setTimeout(function () { window.location.href = '/portal/scheme-deduction-list'; }, 1200);
            } else {
                showAlert((r.message && r.message.message) || 'Bulk deduction failed', 'danger');
            }
        },
        error: function (xhr) {
            $btn.prop('disabled', false).html('<i class="fa fa-check-double"></i> Deduct Selected (<span id="selCount">0</span>)');
            console.error(xhr.responseText);
            showAlert('Error creating deductions', 'danger');
        }
    });
}

// ===================== UTILITIES =====================

function formatCurrency(v) {
    return new Intl.NumberFormat('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(v || 0);
}

function showAlert(msg, type) {
    const a = $(`<div class="alert alert-${type} alert-dismissible fade show" role="alert"
        style="position:fixed;top:90px;right:24px;z-index:9999;min-width:350px;">
        <strong>${msg}</strong>
        <button type="button" class="close" data-dismiss="alert">&times;</button></div>`);
    $('body').append(a);
    setTimeout(() => a.fadeOut(() => a.remove()), 4500);
}

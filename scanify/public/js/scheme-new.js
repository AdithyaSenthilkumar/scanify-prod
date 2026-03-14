/**
 * scheme-new.js - Full logic for New Scheme Request portal page
 */

let itemCounter = 0;
let allProducts = [];
let selectedDoctor = null;
let doctorProductLimits = {}; // product_code -> count_this_month
let hqDivisionMap = {};       // hq_name -> { team, region }
let doctorSearchTimeout = null;

/** Read the active division from the portal's navbar switcher or DOM */
function getActiveDivision() {
    // Try the division-name span in the navbar
    try {
        const btn = document.querySelector('.division-name');
        if (btn && btn.textContent.trim()) return btn.textContent.trim();
    } catch (e) { }
    // Try hidden input on the page
    const hidden = document.getElementById('division');
    if (hidden && hidden.value) return hidden.value;
    // Try cookie
    const match = document.cookie.match(/(?:^|;\s*)division=([^;]*)/);
    if (match) return decodeURIComponent(match[1]);
    return 'Prima';
}

$(document).ready(function () {
    // Set today as default application date
    const today = new Date().toISOString().split('T')[0];
    $('#applicationDate').val(today);

    // Load HQs and products
    loadHQs();
    loadProducts();

    // Doctor dropdown change
    $('#doctorSelect').on('change', function () {
        const selected = $(this).val();
        if (selected) {
            const $opt = $(this).find('option:selected');
            selectDoctorFromDropdown($opt);
        } else {
            clearDoctorSelection();
        }
    });

    // HQ change → fill region/team, load doctors dropdown, load stockists by region
    $('#hqSelect').on('change', function () {
        const hq = $(this).val();
        if (hq) {
            const $opt = $(this).find('option:selected');
            const region = $opt.attr('data-region') || '';
            const regionName = $opt.attr('data-region-name') || region;
            const team = $opt.attr('data-team') || '';
            const teamName = $opt.attr('data-team-name') || team;
            $('#regionDisplay').val(regionName);
            $('#regionValue').val(region);
            $('#teamDisplay').val(teamName);

            // Load doctors dropdown for this HQ
            loadDoctorsForHQ(hq);

            // Load stockists from entire region
            loadStockistsByRegion(region);
        } else {
            $('#regionDisplay').val('');
            $('#regionValue').val('');
            $('#teamDisplay').val('');
            $('#doctorSelect').prop('disabled', true).html('<option value="">-- Select HQ first --</option>');
            $('#stockistSelect').html('<option value="">-- Select HQ first --</option>');
        }

        // Clear doctor selection when HQ changes
        clearDoctorSelection();
    });

    // Form submit
    $('#schemeForm').on('submit', function (e) {
        e.preventDefault();
        submitSchemeRequest();
    });

    // Add first row
    addItemRow();
});

// ===================== MASTER DATA =====================

function loadHQs() {
    const division = getActiveDivision();
    $.ajax({
        url: '/api/method/scanify.api.get_user_hqs',
        type: 'POST',
        contentType: 'application/json',
        headers: { 'X-Frappe-CSRF-Token': frappe.csrf_token },
        data: JSON.stringify({ division: division }),
        success: function (r) {
            if (r.message && r.message.length > 0) {
                let html = '<option value="">-- Select HQ --</option>';
                r.message.forEach(function (hq) {
                    html += `<option value="${hq.name}" data-team="${hq.team || ''}" data-team-name="${hq.team_name || hq.team || ''}" data-region="${hq.region || ''}" data-region-name="${hq.region_name || hq.region || ''}">${hq.hq_name || hq.name}</option>`;
                    hqDivisionMap[hq.name] = { team: hq.team, region: hq.region, team_name: hq.team_name, region_name: hq.region_name };
                });
                $('#hqSelect').html(html);
            } else {
                $('#hqSelect').html('<option value="">No HQs found for this division</option>');
            }
        },
        error: function (xhr) { console.error('HQ load error:', xhr.responseText); }
    });
}

function loadStockistsByRegion(region) {
    if (!region) {
        $('#stockistSelect').html('<option value="">-- Select HQ first --</option>');
        return;
    }
    $('#stockistSelect').html('<option value="">Loading...</option>');
    const division = getActiveDivision();
    $.ajax({
        url: '/api/method/scanify.api.get_stockists_by_region',
        type: 'POST',
        contentType: 'application/json',
        headers: {
            'X-Frappe-CSRF-Token': frappe.csrf_token
        },
        data: JSON.stringify({ region: region, division: division }),
        success: function (r) {
            let html = '<option value="">-- Select Stockist --</option>';
            if (r.message && r.message.length > 0) {
                r.message.forEach(function (s) {
                    const hqLabel = s.hq_name ? ` (HQ: ${s.hq_name})` : '';
                    html += `<option value="${s.name}" data-name="${s.stockist_name}">${s.stockist_name}${hqLabel} — ${s.stockist_code || s.name}</option>`;
                });
            } else {
                html = '<option value="">No stockists found in this region</option>';
            }
            $('#stockistSelect').html(html);
        },
        error: function (xhr) {
            console.error(xhr.responseText);
            $('#stockistSelect').html('<option value="">Error loading stockists</option>');
        }
    });
}

function clearDoctorSelection() {
    selectedDoctor = null;
    $('#doctorCode').val('');
    $('#doctorName, #doctorPlace, #doctorSpecialization, #doctorHospital').val('');
    $('#doctorFieldsRow').hide();
    $('#doctorInfoPanel').hide();
    $('#doctorLimitInfo').hide();
    doctorProductLimits = {};
    updateProductLimitWarnings();
}

function loadProducts() {
    const division = getActiveDivision();
    $.ajax({
        url: '/api/method/scanify.api.get_active_products',
        type: 'POST',
        contentType: 'application/json',
        headers: { 'X-Frappe-CSRF-Token': frappe.csrf_token },
        data: JSON.stringify({ division: division }),
        success: function (r) {
            if (r.message) {
                allProducts = r.message;
                // Refresh existing rows
                $('#itemsTbody tr').each(function () {
                    const rowId = $(this).data('row-id');
                    if (rowId) refreshProductDropdown(rowId);
                });
            }
        },
        error: function (xhr) { console.error('Products load error:', xhr.responseText); }
    });
}

// ===================== DOCTOR DROPDOWN =====================

function loadDoctorsForHQ(hq) {
    const division = getActiveDivision();
    $('#doctorSelect').prop('disabled', true).html('<option value="">Loading doctors...</option>');

    $.ajax({
        url: '/api/method/scanify.api.get_doctors_for_hq',
        type: 'POST',
        contentType: 'application/json',
        headers: { 'X-Frappe-CSRF-Token': frappe.csrf_token },
        data: JSON.stringify({ hq: hq, division: division }),
        success: function (r) {
            let html = '<option value="">-- Select Doctor --</option>';
            if (r.message && r.message.length > 0) {
                r.message.forEach(function (d) {
                    html += `<option value="${d.name}"
                        data-code="${d.name}"
                        data-doctor-code="${d.doctor_code || ''}"
                        data-name="${d.doctor_name}"
                        data-place="${d.place || ''}"
                        data-specialization="${d.specialization || ''}"
                        data-hospital="${d.hospital_address || ''}"
                        data-hq="${d.hq || ''}"
                        data-team="${d.team || ''}"
                        data-region="${d.region || ''}">
                        ${d.doctor_name} (${d.doctor_code || d.name})${d.place ? ' - ' + d.place : ''}
                    </option>`;
                });
            } else {
                html = '<option value="">No doctors found for this HQ</option>';
            }
            $('#doctorSelect').html(html).prop('disabled', false);
        },
        error: function (xhr) {
            console.error('Doctor load error:', xhr.responseText);
            $('#doctorSelect').html('<option value="">Error loading doctors</option>').prop('disabled', false);
        }
    });
}

function selectDoctorFromDropdown($opt) {
    selectedDoctor = {
        code: $opt.data('code'),
        doctor_code: $opt.data('doctor-code'),
        name: $opt.data('name'),
        place: $opt.data('place'),
        specialization: $opt.data('specialization'),
        hospital: $opt.data('hospital'),
        hq: $opt.data('hq'),
        team: $opt.data('team'),
        region: $opt.data('region'),
    };

    $('#doctorCode').val(selectedDoctor.code);

    // Fill doctor fields
    $('#doctorName').val(selectedDoctor.name);
    $('#doctorPlace').val(selectedDoctor.place);
    $('#doctorSpecialization').val(selectedDoctor.specialization);
    $('#doctorHospital').val(selectedDoctor.hospital);
    $('#doctorFieldsRow').show();

    // Show info panel
    $('#doctorInfoContent').html(`
        <strong>${selectedDoctor.name}</strong> (${selectedDoctor.doctor_code})<br>
        <small>${selectedDoctor.place} | ${selectedDoctor.specialization || 'General'}</small>
        ${selectedDoctor.hospital ? `<br><small>${selectedDoctor.hospital}</small>` : ''}
    `);
    $('#doctorInfoPanel').show();

    // Load monthly limit info
    loadDoctorMonthlyLimit(selectedDoctor.code);
}

function loadDoctorMonthlyLimit(doctorCode) {
    const appDate = $('#applicationDate').val() || new Date().toISOString().split('T')[0];
    $.ajax({
        url: '/api/method/scanify.api.get_doctor_monthly_limit_info',
        type: 'POST',
        contentType: 'application/json',
        headers: {
            'X-Frappe-CSRF-Token': frappe.csrf_token
        },
        data: JSON.stringify({ doctor_code: doctorCode, application_date: appDate }),
        success: function (r) {
            if (r.message && r.message.success) {
                const data = r.message;
                doctorProductLimits = data.product_counts || {};
                const total = data.total_requests || 0;
                const month = data.month;

                let badgeClass = total >= 3 ? 'badge-danger' : total >= 2 ? 'badge-warning' : 'badge-success';
                let msg = `${total} scheme(s) already this ${month}. Max 3 per product per month.`;
                if (total === 0) msg = `No schemes yet in ${month}. Max 3 per product per month.`;

                $('#limitBadgeText').text(msg).removeClass('badge-success badge-warning badge-danger').addClass(badgeClass);
                $('#doctorLimitInfo').show();

                // Update product limit warnings in table
                updateProductLimitWarnings();
            }
        },
        error: function (xhr) {
            console.error(xhr.responseText);
        }
    });
}

// ===================== ITEMS TABLE =====================

function addItemRow() {
    itemCounter++;
    const rowId = itemCounter;
    const row = `<tr id="item-row-${rowId}" data-row-id="${rowId}">
        <td>
            <select class="form-control product-select" id="product-${rowId}" onchange="onProductChange(${rowId})">
                <option value="">-- Select Product --</option>
            </select>
            <div id="limit-warning-${rowId}" class="text-danger small mt-1" style="display:none;"></div>
        </td>
        <td><input type="text" class="form-control" id="pack-${rowId}" readonly></td>
        <td><input type="number" class="form-control" id="qty-${rowId}" min="1" value="1" oninput="calculateRow(${rowId})"></td>
        <td><input type="number" class="form-control" id="freeqty-${rowId}" min="0" value="0" oninput="calculateRow(${rowId})"></td>
        <td><input type="number" class="form-control" id="rate-${rowId}" step="0.01" readonly></td>
        <td><input type="number" class="form-control" id="specialrate-${rowId}" step="0.01" placeholder="Optional" oninput="calculateRow(${rowId})"></td>
        <td><input type="text" class="form-control bg-light" id="schemepct-${rowId}" readonly></td>
        <td><input type="text" class="form-control bg-light text-right" id="value-${rowId}" readonly></td>
        <td class="text-center">
            <button type="button" class="btn btn-sm btn-outline-danger" onclick="removeRow(${rowId})">
                <i class="fa fa-trash"></i>
            </button>
        </td>
    </tr>`;

    $('#itemsTbody').append(row);
    refreshProductDropdown(rowId);
}

function refreshProductDropdown(rowId) {
    const $sel = $(`#product-${rowId}`);
    const currentVal = $sel.val();
    let html = '<option value="">-- Select Product --</option>';
    allProducts.forEach(function (p) {
        const used = doctorProductLimits[p.name] || 0;
        const limitReached = used >= 3;
        html += `<option value="${p.name}" data-name="${p.product_name}" data-pack="${p.pack || ''}" data-pts="${p.pts || 0}" ${limitReached ? 'class="text-danger"' : ''}>${p.product_name} (${p.product_code})</option>`;
    });
    $sel.html(html);
    if (currentVal) $sel.val(currentVal);
}

function updateProductLimitWarnings() {
    $('#itemsTbody tr').each(function () {
        const rowId = $(this).data('row-id');
        if (!rowId) return;
        const productCode = $(`#product-${rowId}`).val();
        if (productCode) checkProductLimit(rowId, productCode);
    });
    refreshProductDropdowns();
}

function refreshProductDropdowns() {
    $('#itemsTbody tr').each(function () {
        const rowId = $(this).data('row-id');
        if (rowId) refreshProductDropdown(rowId);
    });
}

function onProductChange(rowId) {
    const $sel = $(`#product-${rowId}`);
    const productCode = $sel.val();
    if (!productCode) {
        $(`#pack-${rowId}`).val('');
        $(`#rate-${rowId}`).val('');
        $(`#pack-${rowId}`).val('');
        $(`#value-${rowId}`).val('');
        $(`#schemepct-${rowId}`).val('');
        $(`#limit-warning-${rowId}`).hide();
        return;
    }

    const opt = $sel.find('option:selected');
    $(`#pack-${rowId}`).val(opt.data('pack') || '');
    $(`#rate-${rowId}`).val(parseFloat(opt.data('pts') || 0).toFixed(2));
    calculateRow(rowId);
    checkProductLimit(rowId, productCode);
}

function checkProductLimit(rowId, productCode) {
    const used = doctorProductLimits[productCode] || 0;
    const $warn = $(`#limit-warning-${rowId}`);
    if (used >= 3) {
        $warn.text(`Limit reached: ${used}/3 requests this month for this product`).removeClass('text-warning').addClass('text-danger').show();
    } else if (used >= 2) {
        $warn.text(`${used}/3 requests this month - approaching limit`).removeClass('text-danger').addClass('text-warning').show();
    } else {
        $warn.hide();
    }
}

function getConversionFactor(packStr) {
    if (!packStr) return 1;
    packStr = String(packStr).trim().toUpperCase();
    var match = packStr.match(/(\d+)\s*[xX]\s*(\d+)/);
    if (match) return parseFloat(match[1]) || 1;
    if (/UNIT|BOX|ML|GM|MG|'S/.test(packStr)) return 1;
    return 1;
}

function calculateRow(rowId) {
    const qty = parseFloat($(`#qty-${rowId}`).val()) || 0;
    const freeQty = parseFloat($(`#freeqty-${rowId}`).val()) || 0;
    const rate = parseFloat($(`#rate-${rowId}`).val()) || 0;
    const specialRate = parseFloat($(`#specialrate-${rowId}`).val()) || 0;
    const pack = $(`#pack-${rowId}`).val() || '';
    const conversionFactor = getConversionFactor(pack);

    let schemePct = 0;
    if (specialRate > 0 && rate > 0) {
        schemePct = ((rate - specialRate) / rate) * 100;
    } else if (freeQty > 0 && qty > 0) {
        schemePct = (freeQty / qty) * 100;
    }

    let value = 0;
    if (specialRate > 0) {
        value = qty * specialRate;
    } else if (freeQty > 0) {
        value = (freeQty / conversionFactor) * rate;
    } else {
        value = qty * rate;
    }

    $(`#schemepct-${rowId}`).val(schemePct.toFixed(2) + '%');
    $(`#value-${rowId}`).val('\u20B9 ' + formatCurrency(value));

    calculateTotal();
}

function calculateTotal() {
    let total = 0;
    $('#itemsTbody tr').each(function () {
        const rowId = $(this).data('row-id');
        if (!rowId) return;
        const valStr = $(`#value-${rowId}`).val().replace(/[^0-9.-]/g, '').trim();
        total += parseFloat(valStr) || 0;
    });
    $('#totalValue').text('\u20B9 ' + formatCurrency(total));
}

function removeRow(rowId) {
    $(`#item-row-${rowId}`).remove();
    calculateTotal();
}

// ===================== SUBMISSION =====================

function submitSchemeRequest() {
    // Validate
    const doctorCode = $('#doctorCode').val();
    if (!doctorCode) {
        showAlert('Please select a doctor', 'warning');
        $('#doctorSelect').focus();
        return;
    }

    const hq = $('#hqSelect').val();
    if (!hq) {
        showAlert('Please select HQ', 'warning');
        return;
    }

    const stockistCode = $('#stockistSelect').val();
    if (!stockistCode) {
        showAlert('Please select a stockist', 'warning');
        return;
    }

    // Collect items
    const items = [];
    let hasProduct = false;
    let hasError = false;

    $('#itemsTbody tr').each(function () {
        const rowId = $(this).data('row-id');
        if (!rowId) return;
        const productCode = $(`#product-${rowId}`).val();
        if (!productCode) return;

        const qty = parseFloat($(`#qty-${rowId}`).val()) || 0;
        if (qty <= 0) {
            showAlert('Order quantity must be greater than 0 for all products', 'warning');
            hasError = true;
            return false;
        }

        // Check limit
        const used = doctorProductLimits[productCode] || 0;
        if (used >= 3) {
            const pName = $(`#product-${rowId} option:selected`).data('name') || productCode;
            showAlert(`Limit reached for ${pName}: already ${used}/3 requests this month`, 'danger');
            hasError = true;
            return false;
        }

        hasProduct = true;
        const $sel = $(`#product-${rowId} option:selected`);
        items.push({
            product_code: productCode,
            product_name: $sel.data('name') || '',
            pack: $(`#pack-${rowId}`).val(),
            quantity: qty,
            free_quantity: parseFloat($(`#freeqty-${rowId}`).val()) || 0,
            product_rate: parseFloat($(`#rate-${rowId}`).val()) || 0,
            special_rate: parseFloat($(`#specialrate-${rowId}`).val()) || 0,
        });
    });

    if (hasError) return;
    if (!hasProduct) {
        showAlert('Please add at least one product', 'warning');
        return;
    }

    const $btn = $('#submitBtn');
    $btn.prop('disabled', true).html('<i class="fa fa-spinner fa-spin"></i> Submitting...');

    // Get stockist name
    const stockistName = $('#stockistSelect option:selected').data('name') || '';
    const hqData = hqDivisionMap[$('#hqSelect').val()] || {};

    const data = {
        application_date: $('#applicationDate').val(),
        doctor_code: selectedDoctor ? selectedDoctor.code : doctorCode,
        doctor_name: selectedDoctor ? selectedDoctor.name : $('#doctorName').val(),
        doctor_place: selectedDoctor ? selectedDoctor.place : '',
        specialization: selectedDoctor ? selectedDoctor.specialization : '',
        hospital_address: selectedDoctor ? selectedDoctor.hospital : '',
        hq: hq,
        team: hqData.team || (selectedDoctor ? selectedDoctor.team : ''),
        region: hqData.region || (selectedDoctor ? selectedDoctor.region : ''),
        stockist_code: stockistCode,
        stockist_name: stockistName,
        scheme_notes: $('#schemeNotes').val(),
        items: items,
    };

    $.ajax({
        url: '/api/method/scanify.api.create_scheme_request_v2',
        type: 'POST',
        contentType: 'application/json',
        headers: {
            'X-Frappe-CSRF-Token': frappe.csrf_token
        },
        data: JSON.stringify({ data: JSON.stringify(data) }),
        success: function (r) {
            $btn.prop('disabled', false).html('<i class="fa fa-save"></i> Submit Scheme Request');
            if (r.message && r.message.success) {
                showAlert('Scheme request created: ' + r.message.name, 'success');
                setTimeout(() => window.location.href = '/portal/scheme-list', 1500);
            } else {
                const msg = (r.message && r.message.message) || 'Failed to create scheme request';
                showAlert(msg, 'danger');
            }
        },
        error: function (xhr) {
            $btn.prop('disabled', false).html('<i class="fa fa-save"></i> Submit Scheme Request');
            console.error(xhr.responseText);
            showAlert('Error submitting scheme request. Please try again.', 'danger');
        }
    });
}

// ===================== UTILITIES =====================

function formatCurrency(v) {
    return new Intl.NumberFormat('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(v || 0);
}

function showAlert(msg, type) {
    let a = $(`<div class="alert alert-${type} alert-dismissible fade show" role="alert" 
        style="position:fixed;top:70px;right:20px;z-index:9999;min-width:350px;box-shadow:0 4px 6px rgba(0,0,0,.15);max-width:500px;">
        <strong>${msg}</strong>
        <button type="button" class="close" data-dismiss="alert">&times;</button></div>`);
    $('body').append(a);
    setTimeout(() => a.fadeOut(() => a.remove()), 5000);
}

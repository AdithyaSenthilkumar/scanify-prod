/**
 * scheme-repeat.js - Logic for Repeat Scheme Request portal page
 * Same form as new scheme request, but:
 *   - Doctor dropdown shows ONLY doctors with approved schemes
 *   - Product dropdown shows ONLY approved products for the selected doctor
 *   - Submitted with repeated_request = 1
 *   - Limit of 3 per product per month is enforced
 */

let itemCounter = 0;
let approvedProducts = [];   // products from approved schemes for the selected doctor
let selectedDoctor = null;
let doctorProductLimits = {};
let hqDivisionMap = {};

function getActiveDivision() {
    try {
        const btn = document.querySelector('.division-name');
        if (btn && btn.textContent.trim()) return btn.textContent.trim();
    } catch (e) { }
    const hidden = document.getElementById('division');
    if (hidden && hidden.value) return hidden.value;
    const match = document.cookie.match(/(?:^|;\s*)division=([^;]*)/);
    if (match) return decodeURIComponent(match[1]);
    return 'Prima';
}

$(document).ready(function () {
    const today = new Date().toISOString().split('T')[0];
    $('#applicationDate').val(today);

    // Load HQs on page load
    loadHQs();

    // HQ change → fill region/team, load approved doctors, load stockists by region
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

            // Load ONLY approved doctors for this HQ
            loadApprovedDoctorsForHQ(hq);

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

    // Form submit
    $('#repeatSchemeForm').on('submit', function (e) {
        e.preventDefault();
        submitRepeatRequest();
    });
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
                    html += '<option value="' + hq.name + '" data-team="' + (hq.team || '') + '" data-team-name="' + (hq.team_name || hq.team || '') + '" data-region="' + (hq.region || '') + '" data-region-name="' + (hq.region_name || hq.region || '') + '">' + (hq.hq_name || hq.name) + '</option>';
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

function loadApprovedDoctorsForHQ(hq) {
    const division = getActiveDivision();
    $('#doctorSelect').prop('disabled', true).html('<option value="">Loading approved doctors...</option>');

    $.ajax({
        url: '/api/method/scanify.api.get_approved_doctors_for_hq',
        type: 'POST',
        contentType: 'application/json',
        headers: { 'X-Frappe-CSRF-Token': frappe.csrf_token },
        data: JSON.stringify({ hq: hq, division: division }),
        success: function (r) {
            let html = '<option value="">-- Select Doctor --</option>';
            if (r.message && r.message.length > 0) {
                r.message.forEach(function (d) {
                    html += '<option value="' + d.name + '"'
                        + ' data-code="' + d.name + '"'
                        + ' data-doctor-code="' + (d.doctor_code || '') + '"'
                        + ' data-name="' + (d.doctor_name || '') + '"'
                        + ' data-place="' + (d.place || '') + '"'
                        + ' data-specialization="' + (d.specialization || '') + '"'
                        + ' data-hospital="' + (d.hospital_address || '') + '"'
                        + ' data-hq="' + (d.hq || '') + '"'
                        + ' data-team="' + (d.team || '') + '"'
                        + ' data-region="' + (d.region || '') + '">'
                        + (d.doctor_name || '') + ' (' + (d.doctor_code || d.name) + ')' + (d.place ? ' - ' + d.place : '')
                        + '</option>';
                });
            } else {
                html = '<option value="">No doctors with approved schemes in this HQ</option>';
            }
            $('#doctorSelect').html(html).prop('disabled', false);
        },
        error: function (xhr) {
            console.error('Doctor load error:', xhr.responseText);
            $('#doctorSelect').html('<option value="">Error loading doctors</option>').prop('disabled', false);
        }
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
        headers: { 'X-Frappe-CSRF-Token': frappe.csrf_token },
        data: JSON.stringify({ region: region, division: division }),
        success: function (r) {
            let html = '<option value="">-- Select Stockist --</option>';
            if (r.message && r.message.length > 0) {
                r.message.forEach(function (s) {
                    var hqLabel = s.hq_name ? ' (HQ: ' + s.hq_name + ')' : '';
                    html += '<option value="' + s.name + '" data-name="' + (s.stockist_name || '') + '">' + (s.stockist_name || '') + hqLabel + ' — ' + (s.stockist_code || s.name) + '</option>';
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

// ===================== DOCTOR SELECTION =====================

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
    $('#doctorInfoContent').html(
        '<strong>' + selectedDoctor.name + '</strong> (' + selectedDoctor.doctor_code + ')<br>' +
        '<small>' + selectedDoctor.place + ' | ' + (selectedDoctor.specialization || 'General') + '</small>' +
        (selectedDoctor.hospital ? '<br><small>' + selectedDoctor.hospital + '</small>' : '')
    );
    $('#doctorInfoPanel').show();

    // Load monthly limit info
    loadDoctorMonthlyLimit(selectedDoctor.code);

    // Load approved products for this doctor, then add first row
    loadApprovedProductsForDoctor(selectedDoctor.code, function () {
        // Clear existing product rows and add a fresh one
        $('#itemsTbody').html('');
        itemCounter = 0;
        $('#addProductBtn').prop('disabled', false);
        addItemRow();
        calculateTotal();
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
    approvedProducts = [];
    // Clear products table
    $('#itemsTbody').html('');
    itemCounter = 0;
    $('#addProductBtn').prop('disabled', true);
    calculateTotal();
    updateProductLimitWarnings();
}

// ===================== APPROVED PRODUCTS =====================

function loadApprovedProductsForDoctor(doctorCode, callback) {
    const division = getActiveDivision();
    approvedProducts = [];

    $.ajax({
        url: '/api/method/scanify.api.get_approved_products_for_doctor',
        type: 'POST',
        contentType: 'application/json',
        headers: { 'X-Frappe-CSRF-Token': frappe.csrf_token },
        data: JSON.stringify({ doctor_code: doctorCode, division: division }),
        success: function (r) {
            if (r.message) {
                approvedProducts = r.message;
            }
            if (callback) callback();
        },
        error: function (xhr) {
            console.error('Approved products load error:', xhr.responseText);
            if (callback) callback();
        }
    });
}

// ===================== DOCTOR MONTHLY LIMIT =====================

function loadDoctorMonthlyLimit(doctorCode) {
    const appDate = $('#applicationDate').val() || new Date().toISOString().split('T')[0];
    $.ajax({
        url: '/api/method/scanify.api.get_doctor_monthly_limit_info',
        type: 'POST',
        contentType: 'application/json',
        headers: { 'X-Frappe-CSRF-Token': frappe.csrf_token },
        data: JSON.stringify({ doctor_code: doctorCode, application_date: appDate }),
        success: function (r) {
            if (r.message && r.message.success) {
                var data = r.message;
                doctorProductLimits = data.product_counts || {};
                var total = data.total_requests || 0;
                var month = data.month;

                var badgeClass = total >= 3 ? 'badge-danger' : total >= 2 ? 'badge-warning' : 'badge-success';
                var msg = total + ' scheme(s) already this ' + month + '. Max 3 per product per month.';
                if (total === 0) msg = 'No schemes yet in ' + month + '. Max 3 per product per month.';

                $('#limitBadgeText').text(msg).removeClass('badge-success badge-warning badge-danger').addClass(badgeClass);
                $('#doctorLimitInfo').show();

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
    var rowId = itemCounter;
    var row = '<tr id="item-row-' + rowId + '" data-row-id="' + rowId + '">' +
        '<td>' +
            '<select class="form-control product-select" id="product-' + rowId + '" onchange="onProductChange(' + rowId + ')">' +
                '<option value="">-- Select Product --</option>' +
            '</select>' +
            '<div id="limit-warning-' + rowId + '" class="text-danger small mt-1" style="display:none;"></div>' +
        '</td>' +
        '<td><input type="text" class="form-control" id="pack-' + rowId + '" readonly></td>' +
        '<td><input type="number" class="form-control" id="qty-' + rowId + '" min="1" value="1" oninput="calculateRow(' + rowId + ')"></td>' +
        '<td><input type="number" class="form-control" id="freeqty-' + rowId + '" min="0" value="0" oninput="calculateRow(' + rowId + ')"></td>' +
        '<td><input type="number" class="form-control" id="rate-' + rowId + '" step="0.01" readonly></td>' +
        '<td><input type="number" class="form-control" id="specialrate-' + rowId + '" step="0.01" placeholder="Optional" oninput="calculateRow(' + rowId + ')"></td>' +
        '<td><input type="text" class="form-control bg-light" id="schemepct-' + rowId + '" readonly></td>' +
        '<td><input type="text" class="form-control bg-light text-right" id="value-' + rowId + '" readonly></td>' +
        '<td class="text-center">' +
            '<button type="button" class="btn btn-sm btn-outline-danger" onclick="removeRow(' + rowId + ')">' +
                '<i class="fa fa-trash"></i>' +
            '</button>' +
        '</td>' +
    '</tr>';

    $('#itemsTbody').append(row);
    refreshProductDropdown(rowId);
}

function refreshProductDropdown(rowId) {
    var $sel = $('#product-' + rowId);
    var currentVal = $sel.val();
    var html = '<option value="">-- Select Product --</option>';
    approvedProducts.forEach(function (p) {
        var used = doctorProductLimits[p.name] || 0;
        var limitReached = used >= 3;
        html += '<option value="' + p.name + '" data-name="' + (p.product_name || '') + '" data-pack="' + (p.pack || '') + '" data-pts="' + (p.pts || 0) + '"' + (limitReached ? ' class="text-danger"' : '') + '>' + (p.product_name || '') + ' (' + (p.product_code || p.name) + ')</option>';
    });
    $sel.html(html);
    if (currentVal) $sel.val(currentVal);
}

function updateProductLimitWarnings() {
    $('#itemsTbody tr').each(function () {
        var rowId = $(this).data('row-id');
        if (!rowId) return;
        var productCode = $('#product-' + rowId).val();
        if (productCode) checkProductLimit(rowId, productCode);
    });
    refreshProductDropdowns();
}

function refreshProductDropdowns() {
    $('#itemsTbody tr').each(function () {
        var rowId = $(this).data('row-id');
        if (rowId) refreshProductDropdown(rowId);
    });
}

function onProductChange(rowId) {
    var $sel = $('#product-' + rowId);
    var productCode = $sel.val();
    if (!productCode) {
        $('#pack-' + rowId).val('');
        $('#rate-' + rowId).val('');
        $('#value-' + rowId).val('');
        $('#schemepct-' + rowId).val('');
        $('#limit-warning-' + rowId).hide();
        return;
    }

    var opt = $sel.find('option:selected');
    $('#pack-' + rowId).val(opt.data('pack') || '');
    $('#rate-' + rowId).val(parseFloat(opt.data('pts') || 0).toFixed(2));
    calculateRow(rowId);
    checkProductLimit(rowId, productCode);
}

function checkProductLimit(rowId, productCode) {
    var used = doctorProductLimits[productCode] || 0;
    var $warn = $('#limit-warning-' + rowId);
    if (used >= 3) {
        $warn.text('Limit reached: ' + used + '/3 requests this month for this product').removeClass('text-warning').addClass('text-danger').show();
    } else if (used >= 2) {
        $warn.text(used + '/3 requests this month - approaching limit').removeClass('text-danger').addClass('text-warning').show();
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
    var qty = parseFloat($('#qty-' + rowId).val()) || 0;
    var freeQty = parseFloat($('#freeqty-' + rowId).val()) || 0;
    var rate = parseFloat($('#rate-' + rowId).val()) || 0;
    var specialRate = parseFloat($('#specialrate-' + rowId).val()) || 0;
    var pack = $('#pack-' + rowId).val() || '';
    var conversionFactor = getConversionFactor(pack);

    var schemePct = 0;
    if (specialRate > 0 && rate > 0) {
        schemePct = ((rate - specialRate) / rate) * 100;
    } else if (freeQty > 0 && qty > 0) {
        schemePct = (freeQty / qty) * 100;
    }

    var value = 0;
    if (specialRate > 0) {
        value = qty * specialRate;
    } else if (freeQty > 0) {
        value = (freeQty / conversionFactor) * rate;
    } else {
        value = qty * rate;
    }

    $('#schemepct-' + rowId).val(schemePct.toFixed(2) + '%');
    $('#value-' + rowId).val('\u20B9 ' + formatCurrency(value));

    calculateTotal();
}

function calculateTotal() {
    var total = 0;
    $('#itemsTbody tr').each(function () {
        var rowId = $(this).data('row-id');
        if (!rowId) return;
        var valStr = $('#value-' + rowId).val().replace(/[^0-9.-]/g, '').trim();
        total += parseFloat(valStr) || 0;
    });
    $('#totalValue').text('\u20B9 ' + formatCurrency(total));
}

function removeRow(rowId) {
    $('#item-row-' + rowId).remove();
    calculateTotal();
}

// ===================== SUBMISSION =====================

function submitRepeatRequest() {
    var hq = $('#hqSelect').val();
    if (!hq) {
        showAlert('Please select an HQ', 'warning');
        return;
    }

    var doctorCode = $('#doctorCode').val();
    if (!doctorCode) {
        showAlert('Please select a doctor', 'warning');
        $('#doctorSelect').focus();
        return;
    }

    var stockistCode = $('#stockistSelect').val();
    if (!stockistCode) {
        showAlert('Please select a stockist', 'warning');
        return;
    }

    // Collect items
    var items = [];
    var hasProduct = false;
    var hasError = false;

    $('#itemsTbody tr').each(function () {
        var rowId = $(this).data('row-id');
        if (!rowId) return;
        var productCode = $('#product-' + rowId).val();
        if (!productCode) return;

        var qty = parseFloat($('#qty-' + rowId).val()) || 0;
        if (qty <= 0) {
            showAlert('Order quantity must be greater than 0 for all products', 'warning');
            hasError = true;
            return false;
        }

        // Check limit
        var used = doctorProductLimits[productCode] || 0;
        if (used >= 3) {
            var pName = $('#product-' + rowId + ' option:selected').data('name') || productCode;
            showAlert('Limit reached for ' + pName + ': already ' + used + '/3 requests this month', 'danger');
            hasError = true;
            return false;
        }

        hasProduct = true;
        var $sel = $('#product-' + rowId + ' option:selected');
        items.push({
            product_code: productCode,
            product_name: $sel.data('name') || '',
            pack: $('#pack-' + rowId).val(),
            quantity: qty,
            free_quantity: parseFloat($('#freeqty-' + rowId).val()) || 0,
            product_rate: parseFloat($('#rate-' + rowId).val()) || 0,
            special_rate: parseFloat($('#specialrate-' + rowId).val()) || 0,
        });
    });

    if (hasError) return;
    if (!hasProduct) {
        showAlert('Please add at least one product', 'warning');
        return;
    }

    var $btn = $('#submitBtn');
    $btn.prop('disabled', true).html('<i class="fa fa-spinner fa-spin"></i> Submitting...');

    var stockistName = $('#stockistSelect option:selected').data('name') || '';
    var hqData = hqDivisionMap[hq] || {};

    var data = {
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
        repeated_request: 1,
        items: items,
    };

    $.ajax({
        url: '/api/method/scanify.api.create_scheme_request_v2',
        type: 'POST',
        contentType: 'application/json',
        headers: { 'X-Frappe-CSRF-Token': frappe.csrf_token },
        data: JSON.stringify({ data: JSON.stringify(data) }),
        success: function (r) {
            $btn.prop('disabled', false).html('<i class="fa fa-redo"></i> Submit Repeat Request');
            if (r.message && r.message.success) {
                showAlert('Repeat scheme request created: ' + r.message.name, 'success');
                setTimeout(function () { window.location.href = '/portal/scheme-list'; }, 1500);
            } else {
                var msg = (r.message && r.message.message) || 'Failed to create repeat request';
                showAlert(msg, 'danger');
            }
        },
        error: function (xhr) {
            $btn.prop('disabled', false).html('<i class="fa fa-redo"></i> Submit Repeat Request');
            console.error(xhr.responseText);
            showAlert('Error submitting repeat request. Please try again.', 'danger');
        }
    });
}

// ===================== UTILITIES =====================

function formatCurrency(v) {
    return new Intl.NumberFormat('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(v || 0);
}

function showAlert(msg, type) {
    var a = $('<div class="alert alert-' + type + ' alert-dismissible fade show" role="alert" style="position:fixed;top:70px;right:20px;z-index:9999;min-width:350px;box-shadow:0 4px 6px rgba(0,0,0,.15);max-width:500px;"><strong>' + msg + '</strong><button type="button" class="close" data-dismiss="alert">&times;</button></div>');
    $('body').append(a);
    setTimeout(function () { a.fadeOut(function () { a.remove(); }); }, 5000);
}

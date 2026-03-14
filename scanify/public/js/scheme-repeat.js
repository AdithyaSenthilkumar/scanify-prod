/**
 * scheme-repeat.js - Logic for Repeat Scheme Request portal page
 * Products are editable; doctor, HQ, stockist are locked from approved scheme.
 * Only approved products for the selected doctor appear in the product dropdowns.
 */

let itemCounter = 0;
let approvedProducts = [];   // products from approved schemes for the doctor
let selectedSchemeData = null;
let doctorProductLimits = {};
let approvedSchemes = [];

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

    loadApprovedSchemes();

    $('#schemeSelect').on('change', function () {
        const val = $(this).val();
        $('#loadSchemeBtn').prop('disabled', !val);
    });

    $('#repeatSchemeForm').on('submit', function (e) {
        e.preventDefault();
        submitRepeatRequest();
    });
});

// ===================== LOAD APPROVED SCHEMES =====================

function loadApprovedSchemes() {
    const filters = { status: 'Approved' };

    $('#schemeSelect').html('<option value="">Loading...</option>');
    $('#loadSchemeBtn').prop('disabled', true);

    $.ajax({
        url: '/api/method/scanify.api.get_scheme_list_portal',
        type: 'POST',
        contentType: 'application/json',
        headers: { 'X-Frappe-CSRF-Token': frappe.csrf_token },
        data: JSON.stringify({ division: getActiveDivision(), filters: JSON.stringify(filters) }),
        success: function (r) {
            if (!(r.message && r.message.success)) {
                $('#schemeSelect').html('<option value="">Error loading schemes</option>');
                return;
            }

            approvedSchemes = r.message.data || [];
            if (!approvedSchemes.length) {
                $('#schemeSelect').html('<option value="">No approved schemes found</option>');
                return;
            }

            let html = '<option value="">-- Select Approved Scheme --</option>';
            approvedSchemes.forEach(function (s) {
                html += '<option value="' + s.name + '">' + s.name + ' | ' + (s.doctor_name || '-') + ' | ' + (s.stockist_name || '-') + ' | \u20B9' + formatCurrency(s.total_scheme_value || 0) + '</option>';
            });
            $('#schemeSelect').html(html);
        },
        error: function (xhr) {
            console.error(xhr.responseText);
            $('#schemeSelect').html('<option value="">Error loading schemes</option>');
        }
    });
}

// ===================== LOAD SCHEME DETAILS =====================

function loadSchemeDetails() {
    const schemeName = $('#schemeSelect').val();
    if (!schemeName) return;

    $('#loadSchemeBtn').prop('disabled', true).html('<i class="fa fa-spinner fa-spin"></i> Loading...');

    $.ajax({
        url: '/api/method/scanify.api.get_scheme_detail',
        type: 'POST',
        contentType: 'application/json',
        headers: { 'X-Frappe-CSRF-Token': frappe.csrf_token },
        data: JSON.stringify({ scheme_name: schemeName }),
        success: function (r) {
            $('#loadSchemeBtn').prop('disabled', false).html('<i class="fa fa-download"></i> Load Scheme');
            if (r.message && r.message.success) {
                selectedSchemeData = r.message;
                populateFormFromScheme(r.message);
            } else {
                showAlert((r.message && r.message.message) || 'Failed to load scheme details', 'danger');
            }
        },
        error: function (xhr) {
            $('#loadSchemeBtn').prop('disabled', false).html('<i class="fa fa-download"></i> Load Scheme');
            console.error(xhr.responseText);
            showAlert('Error loading scheme details', 'danger');
        }
    });
}

function populateFormFromScheme(d) {
    // Source scheme
    $('#sourceSchemeDisplay').val(d.name);

    // HQ & location (read-only)
    $('#hqDisplay').val(d.hq || '-');
    $('#hqValue').val(d.hq || '');
    $('#regionDisplay').val(d.region || '-');
    $('#regionValue').val(d.region || '');
    $('#teamDisplay').val(d.team || '-');
    $('#teamValue').val(d.team || '');

    // Doctor (read-only)
    $('#doctorCode').val(d.doctor_code || '');
    $('#doctorName').val(d.doctor_name || '');
    $('#doctorPlace').val(d.doctor_place || '');
    $('#doctorSpecialization').val(d.specialization || '');
    $('#doctorHospital').val(d.hospital_address || '');

    // Stockist (read-only)
    $('#stockistDisplay').val((d.stockist_name || '-') + ' (' + (d.stockist_code || '') + ')');
    $('#stockistCode').val(d.stockist_code || '');
    $('#stockistName').val(d.stockist_name || '');

    // Scheme notes
    $('#schemeNotes').val('Repeated from ' + d.name);

    // Show the form
    $('#formSection').show();

    // Load approved products for this doctor, then populate items
    loadApprovedProductsForDoctor(d.doctor_code, function () {
        // Clear existing rows
        $('#itemsTbody').html('');
        itemCounter = 0;

        // Pre-fill items from source scheme
        if (d.items && d.items.length > 0) {
            d.items.forEach(function (item) {
                addItemRowWithData(item);
            });
        } else {
            addItemRow();
        }

        calculateTotal();
    });

    // Load doctor monthly limits
    loadDoctorMonthlyLimit(d.doctor_code);
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
                const data = r.message;
                doctorProductLimits = data.product_counts || {};
                const total = data.total_requests || 0;
                const month = data.month;

                let badgeClass = total >= 3 ? 'badge-danger' : total >= 2 ? 'badge-warning' : 'badge-success';
                let msg = total + ' scheme(s) already this ' + month + '. Max 3 per product per month.';
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
    const rowId = itemCounter;
    const row = '<tr id="item-row-' + rowId + '" data-row-id="' + rowId + '">' +
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

function addItemRowWithData(item) {
    itemCounter++;
    const rowId = itemCounter;
    const row = '<tr id="item-row-' + rowId + '" data-row-id="' + rowId + '">' +
        '<td>' +
            '<select class="form-control product-select" id="product-' + rowId + '" onchange="onProductChange(' + rowId + ')">' +
                '<option value="">-- Select Product --</option>' +
            '</select>' +
            '<div id="limit-warning-' + rowId + '" class="text-danger small mt-1" style="display:none;"></div>' +
        '</td>' +
        '<td><input type="text" class="form-control" id="pack-' + rowId + '" readonly></td>' +
        '<td><input type="number" class="form-control" id="qty-' + rowId + '" min="1" value="' + (item.quantity || 1) + '" oninput="calculateRow(' + rowId + ')"></td>' +
        '<td><input type="number" class="form-control" id="freeqty-' + rowId + '" min="0" value="' + (item.free_quantity || 0) + '" oninput="calculateRow(' + rowId + ')"></td>' +
        '<td><input type="number" class="form-control" id="rate-' + rowId + '" step="0.01" readonly></td>' +
        '<td><input type="number" class="form-control" id="specialrate-' + rowId + '" step="0.01" placeholder="Optional" value="' + (item.special_rate || '') + '" oninput="calculateRow(' + rowId + ')"></td>' +
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

    // Set the product value and trigger change
    if (item.product_code) {
        $('#product-' + rowId).val(item.product_code);
        onProductChange(rowId);
    }
}

function refreshProductDropdown(rowId) {
    const $sel = $('#product-' + rowId);
    const currentVal = $sel.val();
    let html = '<option value="">-- Select Product --</option>';
    approvedProducts.forEach(function (p) {
        const used = doctorProductLimits[p.name] || 0;
        const limitReached = used >= 3;
        html += '<option value="' + p.name + '" data-name="' + (p.product_name || '') + '" data-pack="' + (p.pack || '') + '" data-pts="' + (p.pts || 0) + '"' + (limitReached ? ' class="text-danger"' : '') + '>' + (p.product_name || '') + ' (' + (p.product_code || p.name) + ')</option>';
    });
    $sel.html(html);
    if (currentVal) $sel.val(currentVal);
}

function updateProductLimitWarnings() {
    $('#itemsTbody tr').each(function () {
        const rowId = $(this).data('row-id');
        if (!rowId) return;
        const productCode = $('#product-' + rowId).val();
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
    const $sel = $('#product-' + rowId);
    const productCode = $sel.val();
    if (!productCode) {
        $('#pack-' + rowId).val('');
        $('#rate-' + rowId).val('');
        $('#value-' + rowId).val('');
        $('#schemepct-' + rowId).val('');
        $('#limit-warning-' + rowId).hide();
        return;
    }

    const opt = $sel.find('option:selected');
    $('#pack-' + rowId).val(opt.data('pack') || '');
    $('#rate-' + rowId).val(parseFloat(opt.data('pts') || 0).toFixed(2));
    calculateRow(rowId);
    checkProductLimit(rowId, productCode);
}

function checkProductLimit(rowId, productCode) {
    const used = doctorProductLimits[productCode] || 0;
    const $warn = $('#limit-warning-' + rowId);
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
    const qty = parseFloat($('#qty-' + rowId).val()) || 0;
    const freeQty = parseFloat($('#freeqty-' + rowId).val()) || 0;
    const rate = parseFloat($('#rate-' + rowId).val()) || 0;
    const specialRate = parseFloat($('#specialrate-' + rowId).val()) || 0;
    const pack = $('#pack-' + rowId).val() || '';
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

    $('#schemepct-' + rowId).val(schemePct.toFixed(2) + '%');
    $('#value-' + rowId).val('\u20B9 ' + formatCurrency(value));

    calculateTotal();
}

function calculateTotal() {
    let total = 0;
    $('#itemsTbody tr').each(function () {
        const rowId = $(this).data('row-id');
        if (!rowId) return;
        const valStr = $('#value-' + rowId).val().replace(/[^0-9.-]/g, '').trim();
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
    if (!selectedSchemeData) {
        showAlert('Please select and load an approved scheme first', 'warning');
        return;
    }

    const doctorCode = $('#doctorCode').val();
    if (!doctorCode) {
        showAlert('Doctor information is missing', 'warning');
        return;
    }

    const stockistCode = $('#stockistCode').val();
    if (!stockistCode) {
        showAlert('Stockist information is missing', 'warning');
        return;
    }

    // Collect items
    const items = [];
    let hasProduct = false;
    let hasError = false;

    $('#itemsTbody tr').each(function () {
        const rowId = $(this).data('row-id');
        if (!rowId) return;
        const productCode = $('#product-' + rowId).val();
        if (!productCode) return;

        const qty = parseFloat($('#qty-' + rowId).val()) || 0;
        if (qty <= 0) {
            showAlert('Order quantity must be greater than 0 for all products', 'warning');
            hasError = true;
            return false;
        }

        // Check limit
        const used = doctorProductLimits[productCode] || 0;
        if (used >= 3) {
            const pName = $('#product-' + rowId + ' option:selected').data('name') || productCode;
            showAlert('Limit reached for ' + pName + ': already ' + used + '/3 requests this month', 'danger');
            hasError = true;
            return false;
        }

        hasProduct = true;
        const $sel = $('#product-' + rowId + ' option:selected');
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

    const $btn = $('#submitBtn');
    $btn.prop('disabled', true).html('<i class="fa fa-spinner fa-spin"></i> Submitting...');

    const data = {
        application_date: $('#applicationDate').val(),
        doctor_code: doctorCode,
        doctor_name: $('#doctorName').val(),
        doctor_place: $('#doctorPlace').val(),
        specialization: $('#doctorSpecialization').val(),
        hospital_address: $('#doctorHospital').val(),
        hq: $('#hqValue').val(),
        team: $('#teamValue').val(),
        region: $('#regionValue').val(),
        stockist_code: stockistCode,
        stockist_name: $('#stockistName').val(),
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
                const msg = (r.message && r.message.message) || 'Failed to create repeat request';
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

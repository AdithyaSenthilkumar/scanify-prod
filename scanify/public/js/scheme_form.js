let itemCounter = 0;
let products = [];

$(document).ready(function () {
    loadMasters();
    addItemRow();

    // Doctor search
    $('#doctor_search').on('input', function () {
        let term = $(this).val();
        if (term.length >= 2) {
            searchDoctors(term);
        } else {
            $('#doctor-results').removeClass('show');
        }
    });

    // HQ change
    $('#hq').on('change', function () {
        loadStockists($(this).val());
    });

    // Form submission
    $('#scheme-request-form').on('submit', function (e) {
        e.preventDefault();
        submitSchemeRequest();
    });
});

function loadMasters() {
    // Load HQs
    $.ajax({
        url: '/api/method/scanify.api.get_user_hqs',
        type: 'POST',
        contentType: 'application/json',
        headers: {
            'X-Frappe-CSRF-Token': frappe.csrf_token
        },
        success: function (r) {
            if (r.message) {
                let html = '<option value="">Select HQ</option>';
                r.message.forEach(function (hq) {
                    html += `<option value="${hq.name}">${hq.hq_name || hq.hqname || hq.name}</option>`;
                });
                $('#hq').html(html);
            }
        }
    });

    // Load products
    $.ajax({
        url: '/api/method/scanify.api.get_active_products',
        type: 'POST',
        contentType: 'application/json',
        headers: {
            'X-Frappe-CSRF-Token': frappe.csrf_token
        },
        success: function (r) {
            if (r.message) {
                products = r.message;
            }
        }
    });
}

function loadStockists(hq) {
    if (!hq) {
        $('#stockistcode').html('<option value="">Select HQ first</option>');
        return;
    }

    $.ajax({
        url: '/api/method/scanify.api.get_stockists_by_hq',
        type: 'POST',
        contentType: 'application/json',
        headers: {
            'X-Frappe-CSRF-Token': frappe.csrf_token
        },
        data: JSON.stringify({ hq: hq }),
        success: function (r) {
            if (r.message) {
                let html = '<option value="">Select Stockist</option>';
                r.message.forEach(function (stockist) {
                    html += `<option value="${stockist.name}">${stockist.stockistname}</option>`;
                });
                $('#stockistcode').html(html);
            }
        }
    });
}

function searchDoctors(term) {
    $.ajax({
        url: '/api/method/scanify.api.search_doctors',
        type: 'POST',
        contentType: 'application/json',
        headers: {
            'X-Frappe-CSRF-Token': frappe.csrf_token
        },
        data: JSON.stringify({ searchterm: term }),
        success: function (r) {
            if (r.message && r.message.length > 0) {
                let html = '';
                r.message.forEach(function (doctor) {
                    html += `<a class="dropdown-item" href="#" data-code="${doctor.name}" 
                             data-name="${doctor.doctorname}" data-place="${doctor.place || ''}"
                             onclick="selectDoctor(this); return false;">
                        <strong>${doctor.doctorname}</strong><br>
                        <small>${doctor.doctorcode} - ${doctor.place || ''} - ${doctor.specialization || ''}</small>
                    </a>`;
                });
                $('#doctor-results').html(html).addClass('show');
            } else {
                $('#doctor-results').html('<div class="dropdown-item">No doctors found</div>').addClass('show');
            }
        }
    });
}

function selectDoctor(el) {
    let code = $(el).data('code');
    let name = $(el).data('name');
    let place = $(el).data('place');

    $('#doctorcode').val(code);
    $('#doctor_search').val(name);
    $('#doctor-info').text(`${code} - ${place}`);
    $('#doctor-results').removeClass('show');
}

function addItemRow() {
    itemCounter++;
    let html = `<tr id="row-${itemCounter}">
        <td>
            <select class="form-control product-select" name="items[${itemCounter}][productcode]" 
                    onchange="onProductChange(${itemCounter})" required>
                <option value="">Select Product</option>
            </select>
        </td>
        <td><input type="text" class="form-control" name="items[${itemCounter}][pack]" readonly></td>
        <td><input type="number" class="form-control" name="items[${itemCounter}][quantity]" 
                   min="1" required onchange="calculateRow(${itemCounter})"></td>
        <td><input type="number" class="form-control" name="items[${itemCounter}][freequantity]" 
                   min="0" value="0" onchange="calculateRow(${itemCounter})"></td>
        <td><input type="number" class="form-control" name="items[${itemCounter}][productrate]" 
                   step="0.01" readonly></td>
        <td><input type="number" class="form-control" name="items[${itemCounter}][specialrate]" 
                   step="0.01" onchange="calculateRow(${itemCounter})"></td>
        <td><input type="text" class="form-control" name="items[${itemCounter}][schemepercentage]" readonly></td>
        <td><input type="number" class="form-control" name="items[${itemCounter}][productvalue]" 
                   step="0.01" readonly></td>
        <td><button type="button" class="btn btn-sm btn-danger" onclick="removeRow(${itemCounter})">
            <i class="fa fa-trash"></i></button></td>
    </tr>`;

    $('#items-tbody').append(html);

    // Populate product dropdown
    let productHtml = '<option value="">Select Product</option>';
    products.forEach(function (product) {
        productHtml += `<option value="${product.name}">${product.productname} - ${product.productcode}</option>`;
    });
    $(`#row-${itemCounter} .product-select`).html(productHtml);
}

function removeRow(id) {
    $(`#row-${id}`).remove();
    calculateTotal();
}

function onProductChange(rowId) {
    let productCode = $(`#row-${rowId} select[name*="productcode"]`).val();
    if (!productCode) return;

    let product = products.find(p => p.name === productCode);
    if (product) {
        $(`#row-${rowId} input[name*="pack"]`).val(product.pack);
        $(`#row-${rowId} input[name*="productrate"]`).val(product.pts);
        calculateRow(rowId);
    }
}

function calculateRow(rowId) {
    let qty = parseFloat($(`#row-${rowId} input[name*="quantity"]`).val()) || 0;
    let freeQty = parseFloat($(`#row-${rowId} input[name*="freequantity"]`).val()) || 0;
    let rate = parseFloat($(`#row-${rowId} input[name*="productrate"]`).val()) || 0;
    let specialRate = parseFloat($(`#row-${rowId} input[name*="specialrate"]`).val()) || 0;

    let schemePercent = 0;
    if (specialRate > 0 && rate > 0) {
        schemePercent = ((rate - specialRate) / rate) * 100;
    } else if (freeQty > 0 && qty > 0) {
        schemePercent = (freeQty / qty) * 100;
    }

    let value = qty * (specialRate > 0 ? specialRate : rate);

    $(`#row-${rowId} input[name*="schemepercentage"]`).val(schemePercent.toFixed(2) + '%');
    $(`#row-${rowId} input[name*="productvalue"]`).val(value.toFixed(2));

    calculateTotal();
}

function calculateTotal() {
    let total = 0;
    $('input[name*="productvalue"]').each(function () {
        total += parseFloat($(this).val()) || 0;
    });
    $('#total-value').val(format_currency(total));
}

function submitSchemeRequest() {
    let formData = new FormData($('#scheme-request-form')[0]);

    // Validate
    if (!$('#doctorcode').val()) {
        frappe.msgprint('Please select a doctor');
        return;
    }

    if ($('#items-tbody tr').length === 0) {
        frappe.msgprint('Please add at least one product');
        return;
    }

    $('#submit-btn').prop('disabled', true).html('<i class="fa fa-spinner fa-spin"></i> Submitting...');

    $.ajax({
        url: '/api/method/scanify.api.create_scheme_request',
        type: 'POST',
        contentType: 'application/json',
        headers: {
            'X-Frappe-CSRF-Token': frappe.csrf_token
        },
        data: JSON.stringify({ data: getFormData() }),
        success: function (r) {
            $('#submit-btn').prop('disabled', false).html('<i class="fa fa-save"></i> Submit Scheme Request');

            if (r.message && r.message.success) {
                if (typeof frappe.msgprint !== 'undefined') {
                    frappe.msgprint({
                        title: 'Success',
                        message: 'Scheme request created successfully',
                        indicator: 'green'
                    });
                } else {
                    alert('Scheme request created successfully');
                }
                setTimeout(function () {
                    window.location.href = '/portal/schemes/' + r.message.name;
                }, 1500);
            } else {
                let msg = (r.message && r.message.message) || 'Failed to create scheme request';
                if (typeof frappe.msgprint !== 'undefined') {
                    frappe.msgprint({
                        title: 'Error',
                        message: msg,
                        indicator: 'red'
                    });
                } else {
                    alert('Error: ' + msg);
                }
            }
        },
        error: function (xhr) {
            $('#submit-btn').prop('disabled', false).html('<i class="fa fa-save"></i> Submit Scheme Request');
            console.error(xhr.responseText);
            if (typeof frappe.msgprint !== 'undefined') {
                frappe.msgprint({
                    title: 'Error',
                    message: 'Failed to create scheme request',
                    indicator: 'red'
                });
            } else {
                alert('Failed to create scheme request');
            }
        }
    });
}

function getFormData() {
    let data = {
        applicationdate: $('input[name="applicationdate"]').val(),
        hq: $('#hq').val(),
        doctorcode: $('#doctorcode').val(),
        stockistcode: $('#stockistcode').val(),
        chemist: $('input[name="chemist"]').val(),
        schemenotes: $('textarea[name="schemenotes"]').val(),
        items: []
    };

    $('#items-tbody tr').each(function () {
        let item = {
            productcode: $(this).find('select[name*="productcode"]').val(),
            quantity: parseInt($(this).find('input[name*="quantity"]').val()) || 0,
            freequantity: parseInt($(this).find('input[name*="freequantity"]').val()) || 0,
            specialrate: parseFloat($(this).find('input[name*="specialrate"]').val()) || 0
        };
        if (item.productcode) {
            data.items.push(item);
        }
    });

    return data;
}

let currentMasterType = 'hq';
let currentDivision = '';
let allRecords = [];
let filteredRecords = [];

// Master configurations
const masterConfigs = {
    hq: {
        title: 'HQ Master',
        doctype: 'HQ Master',
        fields: [
            { name: 'name', label: 'HQ Code', type: 'text', required: true, readonly_on_edit: true },
            { name: 'hq_name', label: 'HQ Name', type: 'text', required: true },
            { name: 'team', label: 'Team', type: 'link', options: 'Team Master', required: true },
            { name: 'region', label: 'Region', type: 'link', options: 'Region Master', required: true },
            { name: 'status', label: 'Status', type: 'select', options: ['Active', 'Inactive'], required: true }
        ],
        columns: ['name', 'hq_name', 'team', 'region', 'division', 'status'],
        searchFields: ['name', 'hq_name', 'team', 'region'],
        excelColumns: ['HQ Code', 'HQ Name', 'Team', 'Region', 'Division', 'Status'],
        excelSample: ['HQ001', 'Chennai Central', 'Team A', 'South', 'Prima', 'Active']
    },
    stockist: {
        title: 'Stockist Master',
        doctype: 'Stockist Master',
        fields: [
            { name: 'stockist_code', label: 'Stockist Code', type: 'text', required: true, readonly_on_edit: true },
            { name: 'stockist_name', label: 'Stockist Name', type: 'text', required: true },
            { name: 'hq', label: 'HQ', type: 'link', options: 'HQ Master', required: true },
            { name: 'address', label: 'Address', type: 'textarea', required: false },
            { name: 'contact', label: 'Contact Number', type: 'text', required: false },
            { name: 'email', label: 'Email', type: 'email', required: false },
            { name: 'status', label: 'Status', type: 'select', options: ['Active', 'Inactive'], required: true }
        ],
        columns: ['stockist_code', 'stockist_name', 'hq', 'contact', 'status'],
        searchFields: ['stockist_code', 'stockist_name', 'hq'],
        excelColumns: ['Stockist Code', 'Stockist Name', 'HQ', 'Address', 'Contact', 'Email', 'Status'],
        excelSample: ['STK001', 'ABC Pharma Distributors', 'HQ001', '123 Main St, Chennai', '9876543210', 'abc@example.com', 'Active']
    },
    product: {
        title: 'Product Master',
        doctype: 'Product Master',
        fields: [
            { name: 'product_code', label: 'Product Code', type: 'text', required: true, readonly_on_edit: true },
            { name: 'product_name', label: 'Product Name', type: 'text', required: true },
            { name: 'pack', label: 'Pack Size', type: 'text', required: true },
            { name: 'pts', label: 'PTS Rate', type: 'number', required: true },
            { name: 'ptr', label: 'PTR Rate', type: 'number', required: false },
            { name: 'product_type', label: 'Product Type', type: 'select', options: ['Prima', 'Vektra'], required: false },
            { name: 'status', label: 'Status', type: 'select', options: ['Active', 'Inactive'], required: true }
        ],
        columns: ['product_code', 'product_name', 'pack', 'pts', 'division', 'status'],
        searchFields: ['product_code', 'product_name'],
        excelColumns: ['Product Code', 'Product Name', 'Pack', 'PTS', 'PTR', 'Division', 'Product Type', 'Status'],
        excelSample: ['PROD001', 'Paracetamol 500mg', '10x10', '50.00', '55.00', 'Prima', 'Prima', 'Active']
    },
    doctor: {
        title: 'Doctor Master',
        doctype: 'Doctor Master',
        fields: [
            { name: 'doctor_code', label: 'Doctor Code', type: 'text', required: true, readonly_on_edit: true },
            { name: 'doctor_name', label: 'Doctor Name', type: 'text', required: true },
            { name: 'place', label: 'Place', type: 'text', required: true },
            { name: 'specialization', label: 'Specialization', type: 'text', required: false },
            { name: 'hospital_address', label: 'Hospital Address', type: 'textarea', required: false },
            { name: 'hq', label: 'HQ', type: 'link', options: 'HQ Master', required: true },
            { name: 'team', label: 'Team', type: 'link', options: 'Team Master', required: true },
            { name: 'region', label: 'Region', type: 'link', options: 'Region Master', required: true },
            { name: 'citypool', label: 'City Pool', type: 'text', required: false },
            { name: 'status', label: 'Status', type: 'select', options: ['Active', 'Inactive'], required: true }
        ],
        columns: ['doctor_code', 'doctor_name', 'place', 'specialization', 'hq', 'team', 'status'],
        searchFields: ['doctor_code', 'doctor_name', 'place', 'hq'],
        excelColumns: ['Doctor Code', 'Doctor Name', 'Place', 'Specialization', 'Hospital Address', 'HQ', 'Team', 'Region', 'City Pool', 'Status'],
        excelSample: ['DOC001', 'Dr. Sharma', 'Chennai', 'Cardiology', 'Apollo Hospital, Chennai', 'HQ001', 'Team A', 'South', 'Chennai', 'Active']
    },
    team: {
        title: 'Team Master',
        doctype: 'Team Master',
        fields: [
            { name: 'name', label: 'Team Code', type: 'text', required: true, readonly_on_edit: true },
            { name: 'team_name', label: 'Team Name', type: 'text', required: true },
            { name: 'region', label: 'Region', type: 'link', options: 'Region Master', required: true },
            { name: 'status', label: 'Status', type: 'select', options: ['Active', 'Inactive'], required: true }
        ],
        columns: ['name', 'team_name', 'region', 'division', 'status'],
        searchFields: ['name', 'team_name', 'region'],
        excelColumns: ['Team Code', 'Team Name', 'Region', 'Division', 'Status'],
        excelSample: ['TEAM001', 'Team Alpha', 'South', 'Prima', 'Active']
    },
    region: {
        title: 'Region Master',
        doctype: 'Region Master',
        fields: [
            { name: 'name', label: 'Region Code', type: 'text', required: true, readonly_on_edit: true },
            { name: 'region_name', label: 'Region Name', type: 'text', required: true },
            { name: 'status', label: 'Status', type: 'select', options: ['Active', 'Inactive'], required: true }
        ],
        columns: ['name', 'region_name', 'division', 'status'],
        searchFields: ['name', 'region_name'],
        excelColumns: ['Region Code', 'Region Name', 'Division', 'Status'],
        excelSample: ['RGN001', 'South Region', 'Prima', 'Active']
    }
};

window.addEventListener('load', function() {
    // Get current division from the page
    const divisionText = document.querySelector('#divisionMenuButton .division-name');
    currentDivision = divisionText ? divisionText.textContent.trim() : 'Prima';

    // Load initial data
    loadMasterData();

    // Event listeners
    document.getElementById('master-type').addEventListener('change', function() {
        currentMasterType = this.value;
        loadMasterData();
    });

    document.getElementById('search-input').addEventListener('input', debounce(function() {
        filterRecords();
    }, 300));
});

// REPLACE loadMasterData to use $.ajax instead of frappe.call
function loadMasterData() {
    const config = masterConfigs[currentMasterType];
    
    // Update UI
    $('#table-title').text(config.title);
    buildTableHeader(config);
    buildFilterControls(config);
    
    // Show loading
    $('#table-body').html(`
        <tr>
            <td colspan="${config.columns.length + 1}" class="text-center p-5">
                <i class="fa fa-spinner fa-spin fa-3x text-muted"></i>
                <p class="mt-3 text-muted">Loading ${config.title}...</p>
            </td>
        </tr>
    `);
    
    // Fetch data with CSRF token
    $.ajax({
        url: '/api/method/scanify.api.get_master_data',
        type: 'POST',
        contentType: 'application/json',
        headers: {
        'X-Frappe-CSRF-Token': window.csrf_token
        },
        data: JSON.stringify({
            doctype: config.doctype,
            division: currentDivision
        }),
        success: function(response) {
            console.log('API Response:', response);
            if (response.message && response.message.success) {
                allRecords = response.message.data;
                filteredRecords = [...allRecords];
                renderTable();
            } else {
                const errorMsg = response.message ? response.message.message : 'Failed to load data';
                showError(errorMsg);
                $('#table-body').html(`
                    <tr>
                        <td colspan="${config.columns.length + 1}" class="text-center p-5">
                            <i class="fa fa-exclamation-triangle fa-3x text-danger"></i>
                            <p class="mt-3 text-danger">${errorMsg}</p>
                        </td>
                    </tr>
                `);
            }
        },
        error: function(xhr, status, error) {
            console.error('AJAX Error:', status, error, xhr.responseText);
            showError('Error loading data: ' + error);
            $('#table-body').html(`
                <tr>
                    <td colspan="${config.columns.length + 1}" class="text-center p-5">
                        <i class="fa fa-exclamation-triangle fa-3x text-danger"></i>
                        <p class="mt-3 text-danger">Error loading data</p>
                        <button class="btn btn-sm btn-primary" onclick="loadMasterData()">Retry</button>
                    </td>
                </tr>
            `);
        }
    });
}


// Build table header
function buildTableHeader(config) {
    let html = '<tr>';
    config.columns.forEach(col => {
        const field = config.fields.find(f => f.name === col);
        html += `<th>${field ? field.label : col}</th>`;
    });
    html += '<th class="text-center">Actions</th>';
    html += '</tr>';
    $('#table-header').html(html);
}

// Build filter controls
function buildFilterControls(config) {
    let html = '';
    
    // Status filter (common for all)
    html += `
        <div class="col-md-3">
            <label>Status</label>
            <select class="form-control filter-control" data-field="status">
                <option value="">All</option>
                <option value="Active">Active</option>
                <option value="Inactive">Inactive</option>
            </select>
        </div>
    `;
    
    // Division filter (for applicable masters)
    if (['hq', 'product', 'team', 'region'].includes(currentMasterType)) {
        html += `
            <div class="col-md-3">
                <label>Division</label>
                <select class="form-control filter-control" data-field="division">
                    <option value="">All</option>
                    <option value="Prima">Prima</option>
                    <option value="Vektra">Vektra</option>
                    <option value="Both">Both</option>
                </select>
            </div>
        `;
    }
    
    $('#filter-controls').html(html);
    
    // Attach filter listeners
    $('.filter-control').on('change', function() {
        filterRecords();
    });
}

// Render table
function renderTable() {
    const config = masterConfigs[currentMasterType];
    
    if (filteredRecords.length === 0) {
        $('#table-body').html(`
            <tr>
                <td colspan="${config.columns.length + 1}" class="text-center p-5">
                    <i class="fa fa-inbox fa-3x text-muted"></i>
                    <p class="mt-3 text-muted">No records found</p>
                </td>
            </tr>
        `);
        $('#record-count').text('0');
        return;
    }
    
    let html = '';
    filteredRecords.forEach(record => {
        html += '<tr>';
        config.columns.forEach(col => {
            html += `<td>${record[col] || '-'}</td>`;
        });
        html += `
            <td class="text-center action-buttons">
                <button class="btn btn-sm btn-primary" onclick="editRecord('${record.name}')" title="Edit">
                    <i class="fa fa-edit"></i>
                </button>
                <button class="btn btn-sm btn-danger" onclick="deleteRecord('${record.name}')" title="Delete">
                    <i class="fa fa-trash"></i>
                </button>
            </td>
        `;
        html += '</tr>';
    });
    
    $('#table-body').html(html);
    $('#record-count').text(filteredRecords.length);
}

// Filter records
function filterRecords() {
    const config = masterConfigs[currentMasterType];
    const searchTerm = $('#search-input').val().toLowerCase();
    
    filteredRecords = allRecords.filter(record => {
        // Search filter
        if (searchTerm) {
            const matches = config.searchFields.some(field => {
                const value = record[field];
                return value && value.toString().toLowerCase().includes(searchTerm);
            });
            if (!matches) return false;
        }
        
        // Other filters
        let passFilters = true;
        $('.filter-control').each(function() {
            const field = $(this).data('field');
            const value = $(this).val();
            if (value && record[field] !== value) {
                passFilters = false;
            }
        });
        
        return passFilters;
    });
    
    renderTable();
}

// Reset filters
function resetFilters() {
    $('#search-input').val('');
    $('.filter-control').val('');
    filterRecords();
}

// Show add modal
function showAddModal() {
    const config = masterConfigs[currentMasterType];
    $('#modalTitle').text(`Add New ${config.title}`);
    $('#record-id').val('');
    buildForm(config, null);
    $('#editModal').modal('show');
}

// Edit record
function editRecord(name) {
    const config = masterConfigs[currentMasterType];
    const record = allRecords.find(r => r.name === name);
    
    $('#modalTitle').text(`Edit ${config.title}`);
    $('#record-id').val(name);
    buildForm(config, record);
    $('#editModal').modal('show');
}

function buildForm(config, data) {
    let html = '<div class="row">';
    
    config.fields.forEach((field, index) => {
        // Skip rendering the code field for new records (it's auto-generated)
        const isCodeField = ['name', 'stockistcode', 'productcode', 'doctorcode'].includes(field.name);
        if (!data && isCodeField) {
            return; // Don't show code fields for new records
        }
        
        // Skip division field (it's session-based)
        if (field.name === 'division') {
            return; // DON'T SHOW DIVISION FIELD
        }
        
        const value = data ? (data[field.name] || '') : 
                     (field.name === 'status' ? 'Active' : '');
        
        const readonly = (data && field.readonlyonedit) || (isCodeField && data) ? 'readonly' : '';
        
        html += `<div class="col-md-6 mb-3">`;
        html += `<label>${field.label}${field.required ? '<span class="text-danger">*</span>' : ''}</label>`;
        
        if (field.type === 'select') {
            html += `<select class="form-control" name="${field.name}" ${field.required ? 'required' : ''} ${readonly}>`;
            if (Array.isArray(field.options)) {
                field.options.forEach(opt => {
                    html += `<option value="${opt}" ${value == opt ? 'selected' : ''}>${opt}</option>`;
                });
            }
            html += `</select>`;
        } else if (field.type === 'textarea') {
            html += `<textarea class="form-control" name="${field.name}" rows="2" ${field.required ? 'required' : ''}>${value}</textarea>`;
        } else if (field.type === 'link') {
            html += `<input type="text" class="form-control link-field" name="${field.name}" data-doctype="${field.options}" value="${value}" ${readonly} ${field.required ? 'required' : ''}>`;
        } else {
            html += `<input type="${field.type}" class="form-control" name="${field.name}" value="${value}" ${readonly} ${field.required ? 'required' : ''} ${field.type === 'number' ? 'step="0.01"' : ''}>`;
        }
        
        html += `</div>`;
        html += `</div>`;
    });
    
    $('#form-fields').html(html);
}
// Save record
function saveRecord() {
    const config = masterConfigs[currentMasterType];
    const recordId = $('#record-id').val();
    
    // Collect form data
    const data = {};
    config.fields.forEach(field => {
        // Skip division field
        if (field.name === 'division') {
            return;
        }
        
        // Skip code fields for new records (auto-generated)
        const isCodeField = ['name', 'stockistcode', 'productcode', 'doctorcode'].includes(field.name);
        if (!recordId && isCodeField) {
            return;
        }
        
        const value = $(`[name="${field.name}"]`).val();
        if (value) {
            data[field.name] = value;
        }
    });
    
    // Validate required fields
    let isValid = true;
    config.fields.forEach(field => {
        if (field.required && !data[field.name] && field.name !== 'division') {
            showError(`${field.label} is required`);
            isValid = false;
        }
    });
    
    if (!isValid) return;
    
    showLoadingOverlay('Saving...');
    
    // AJAX with CSRF token
    $.ajax({
        url: '/api/method/scanify.api.save_master_record',
        type: 'POST',
        contentType: 'application/json',
        headers: {
            'X-Frappe-CSRF-Token': frappe.csrf_token  // ADD CSRF TOKEN
        },
        data: JSON.stringify({
            doctype: config.doctype,
            name: recordId,
            data: data
        }),
        success: function(response) {
            hideLoadingOverlay();
            if (response.message && response.message.success) {
                showAlert('Record saved successfully!', 'success');
                $('#editModal').modal('hide');
                loadMasterData();
            } else {
                showError(response.message ? response.message.message : 'Failed to save record');
            }
        },
        error: function(xhr, status, error) {
            hideLoadingOverlay();
            showError('Error saving record: ' + error);
        }
    });
}


function deleteRecord(name) {
    if (!confirm('Are you sure you want to delete this record?')) {
        return;
    }
    
    const config = masterConfigs[currentMasterType];
    showLoadingOverlay('Deleting...');
    
    $.ajax({
        url: '/api/method/scanify.api.delete_master_record',
        type: 'POST',
        contentType: 'application/json',
        headers: {
        'X-Frappe-CSRF-Token': window.csrf_token
        },
        data: JSON.stringify({
            doctype: config.doctype,
            name: name
        }),
        success: function(response) {
            hideLoadingOverlay();
            if (response.message && response.message.success) {
                showAlert('Record deleted successfully!', 'success');
                loadMasterData();
            } else {
                showError(response.message?.message || 'Failed to delete record');
            }
        },
        error: function(xhr) {
            hideLoadingOverlay();
            showError('Error deleting record');
        }
    });
}


// Show bulk import modal
function showBulkImportModal() {
    const config = masterConfigs[currentMasterType];
    $('#import-master-type').text(config.title);
    
    // Build sample table
    let sampleHtml = `
        <table class="table table-bordered sample-excel-table">
            <thead><tr>`;
    
    config.excelColumns.forEach(col => {
        sampleHtml += `<th>${col}</th>`;
    });
    
    sampleHtml += `</tr></thead><tbody><tr>`;
    
    config.excelSample.forEach(val => {
        sampleHtml += `<td>${val}</td>`;
    });
    
    sampleHtml += `</tr></tbody></table>`;
    
    $('#sample-table-container').html(sampleHtml);
    $('#import-file').val('');
    $('#import-progress').hide();
    $('#import-results').hide();
    
    $('#bulkImportModal').modal('show');
}

function processImport() {
    const file = $('#import-file')[0].files[0];
    if (!file) {
        showError('Please select a file');
        return;
    }
    
    const config = masterConfigs[currentMasterType];
    const formData = new FormData();
    formData.append('file', file);
    formData.append('doctype', config.doctype);
    formData.append('division', currentDivision);
    
    // Show progress
    $('#import-progress').show();
    $('.progress-bar').css('width', '10%');
    
    // Upload file
    $.ajax({
        url: '/api/method/scanify.api.import_master_data',
        type: 'POST',
        data: formData,
        processData: false,
        contentType: false,
        headers: {
        'X-Frappe-CSRF-Token': window.csrf_token
        },
        xhr: function() {
            const xhr = new window.XMLHttpRequest();
            xhr.upload.addEventListener('progress', function(e) {
                if (e.lengthComputable) {
                    const percent = (e.loaded / e.total) * 100;
                    $('.progress-bar').css('width', percent + '%');
                }
            });
            return xhr;
        },
        success: function(response) {
            $('.progress-bar').css('width', '100%');
            if (response.message && response.message.success) {
                $('#import-alert')
                    .removeClass('alert-danger')
                    .addClass('alert-success')
                    .html(`<strong>Import Successful!</strong><br>
                           <i class="fa fa-check-circle"></i> ${response.message.imported} records imported<br>
                           ${response.message.failed > 0 ? `<i class="fa fa-exclamation-triangle"></i> ${response.message.failed} records failed` : ''}`);
                
                setTimeout(function() {
                    $('#bulkImportModal').modal('hide');
                    loadMasterData();
                }, 2000);
            } else {
                $('#import-alert')
                    .removeClass('alert-success')
                    .addClass('alert-danger')
                    .html(`<strong>Import Failed!</strong><br>${response.message ? response.message.message : 'Unknown error'}`);
            }
            $('#import-results').show();
        },
        error: function() {
            $('.progress-bar').css('width', '100%').addClass('bg-danger');
            $('#import-alert')
                .removeClass('alert-success')
                .addClass('alert-danger')
                .html('<strong>Error!</strong> Failed to upload file');
            $('#import-results').show();
        }
    });
}

// Utility functions
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

function showLoadingOverlay(message) {
    const overlay = `
        <div id="loading-overlay" style="position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.7);z-index:99999;display:flex;align-items:center;justify-content:center;">
            <div style="background:white;padding:20px 40px;border-radius:8px;text-align:center;">
                <div class="spinner-border text-primary mb-2" role="status"></div>
                <div style="font-size:16px;color:#333;">${message}</div>
            </div>
        </div>
    `;
    $('body').append(overlay);
}

function hideLoadingOverlay() {
    $('#loading-overlay').remove();
}

function showAlert(message, type) {
    const alert = `
        <div class="alert alert-${type} alert-dismissible fade show" role="alert" style="position:fixed;top:70px;right:20px;z-index:9999;min-width:300px;box-shadow:0 4px 6px rgba(0,0,0,0.1);">
            <strong>${message}</strong>
            <button type="button" class="close" data-dismiss="alert">&times;</button>
        </div>
    `;
    $('body').append(alert);
    
    setTimeout(function() {
        $('.alert').fadeOut(function() {
            $(this).remove();
        });
    }, 3000);
}

function showError(message) {
    showAlert(message, 'danger');
}

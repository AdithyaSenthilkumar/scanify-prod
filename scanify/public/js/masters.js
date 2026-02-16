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
                { name: 'hq_name', label: 'HQ Name', type: 'text', required: true },  // FIXED: removed code field, only hq_name
                { name: 'team', label: 'Team', type: 'link', options: 'Team Master', required: true },
                { name: 'region', label: 'Region', type: 'link', options: 'Region Master', required: true },
                { name: 'zone', label: 'Zone', type: 'text' },
                { name: 'status', label: 'Status', type: 'select', options: ['Active', 'Inactive'], required: true }
            ],
            columns: ['hq_name', 'team', 'region', 'zone', 'status'],
            searchFields: ['hq_name', 'team', 'region', 'zone'],
            excelColumns: ['HQ Name', 'Team', 'Region', 'Zone', 'Status'],
            excelSample: ['Chennai Central', 'Team A', 'South', 'Zone 123 Main St, Chennai', 'Active']
        },
        stockist: {
            title: 'Stockist Master',
            doctype: 'Stockist Master',
            fields: [
                { name: 'stockist_code', label: 'Stockist Code', type: 'text', readonly_on_edit: true },
                { name: 'stockist_name', label: 'Stockist Name', type: 'text', required: true },
                // Hierarchy
                { name: 'division', label: 'Division', type: 'link', options: 'Division' },
                { name: 'hq', label: 'HQ', type: 'link', options: 'HQ Master', required: true },
                { name: 'team', label: 'Team', type: 'link', options: 'Team Master', required: true },
                { name: 'region', label: 'Region', type: 'link', options: 'Region Master', required: true },
                { name: 'zone', label: 'Zone', type: 'text' }, // NOT LINK

                // Address
                { name: 'address', label: 'Address', type: 'textarea' },

                // Contact
                { name: 'contact_person', label: 'Contact Person', type: 'text' },
                { name: 'phone', label: 'Phone', type: 'text' },
                { name: 'email', label: 'Email', type: 'email' },

                { name: 'status', label: 'Status', type: 'select', options: ['Active','Inactive'], required: true }
            ],
            columns: ['stockist_code','stockist_name','hq','team','region','zone','phone','status'],
            searchFields: ['stockist_code','stockist_name','city','phone'],
            excelColumns: ['Stockist Code', 'Stockist Name', 'HQ', 'Team', 'Region', 'Zone', 'Address', 'Contact Person', 'Phone', 'Email', 'Status'],
            excelSample: ['STK001', 'ABC Pharma Distributors', 'HQ001', 'Team A', 'South','North Zone', 'Zone 123 Main St, Chennai','John Smith', '9876543210', 'abc@example.com', 'Active']
                },
        product: {
    title: 'Product Master',
    doctype: 'Product Master',
    fields: [
        // Client-entered code (DocType autoname is field:product_code)
        { name: 'product_code', label: 'Product Code', type: 'text', required: true, readonly_on_edit: true }, // keep readonly on edit unless backend rename is implemented
        { name: 'product_name', label: 'Product Name', type: 'text', required: true },

        // Dropdowns as per DocType JSON
        {
            name: 'product_group',
            label: 'Product Group',
            type: 'select',
            required: true,
            options: ['Dentist', 'Derma', 'Contus', 'Xptum', 'Amino', 'Ortho', 'Drez', 'Gynae', 'Jusdee', 'Dygerm', 'Others']
        },
        {
            name: 'category',
            label: 'Category',
            type: 'select',
            options: ['Main Product', 'Hospital Product', 'New Product']
        },

        { name: 'pack', label: 'Pack', type: 'text' },
        // DocType fieldtype is Data (often values like "10's", "Unit", etc.)
        { name: 'pack_conversion', label: 'Pack Conversion', type: 'text' },

        // Pricing (DocType: Currency/Percent)
        { name: 'mrp', label: 'MRP', type: 'number' },
        { name: 'ptr', label: 'PTR', type: 'number' },
        { name: 'pts', label: 'PTS', type: 'number' },
        { name: 'gst_rate', label: 'GST Rate (%)', type: 'number', default: 5 },

        { name: 'status', label: 'Status', type: 'select', options: ['Active', 'Inactive'], required: true }
    ],

    // Columns: remove product_type, include category/product_group, use gst_rate
    columns: ['product_code', 'product_name', 'product_group', 'category', 'pack', 'pack_conversion', 'pts', 'ptr', 'mrp', 'gst_rate', 'status'],

    searchFields: ['product_code', 'product_name', 'product_group', 'category', 'pack'],

    // Excel: remove Product Type, add Category, fix GST field naming
    excelColumns: [
        'Product Code', 'Product Name', 'Product Group', 'Category',
        'Pack', 'Pack Conversion', 'PTS', 'PTR', 'MRP', 'GST Rate (%)', 'Status'
    ],
    excelSample: [
        'PROD001', 'Paracetamol 500mg', 'Others', 'Main Product',
        '10x10', "10's", '10.00', '55.00', '60.00', '5', 'Active'
    ]
},
           doctor: {
        title: 'Doctor Master',
        doctype: 'Doctor Master',
        fields: [
            { name: 'doctor_code', label: 'Doctor Code', type: 'text', required: true, readonly_on_edit: true },
            { name: 'doctor_name', label: 'Doctor Name', type: 'text', required: true },
            { name: 'qualification', label: 'Qualification', type: 'text' },
            { name: 'doctor_category', label: 'Doctor Category', type: 'select', options: ['CRM', 'ASM KBL', 'SO – TOP GYANE'] },
            { name: 'specialization', label: 'Specialization', type: 'text' },
            { name: 'phone', label: 'Phone', type: 'text' },
            { name: 'place', label: 'Place', type: 'text' },
            { name: 'hospital_address', label: 'Hospital Address', type: 'textarea' },
            { name: 'house_address', label: 'House Address', type: 'textarea' },
            { name: 'division', label: 'Division', type: 'link', options: 'Division' },
            { name: 'hq', label: 'HQ', type: 'link', options: 'HQ Master' },
            { name: 'team', label: 'Team', type: 'link', options: 'Team Master' },
            { name: 'region', label: 'Region', type: 'link', options: 'Region Master' },
            { name: 'state_name', label: 'State', type: 'text' },
            { name: 'zone', label: 'Zone', type: 'text' },
            { name: 'chemist_name', label: 'Chemist Name', type: 'text' },
            { name: 'status', label: 'Status', type: 'select', options: ['Active', 'Inactive'], required: true }
        ],
        columns: ['doctor_code', 'doctor_name', 'place', 'specialization', 'hq', 'phone', 'status'],
        searchFields: ['doctor_code', 'doctor_name', 'place', 'phone', 'chemist_name'],
        excelColumns: [
            'Doctor Code', 'Doctor Name', 'Qualification', 'Doctor Category', 
            'Specialization', 'Phone', 'Place', 'Hospital Address', 'House Address', 
            'Division', 'HQ', 'Team', 'Region', 'State', 'Zone', 'Chemist Name', 'Status'
        ],
        excelSample: [
            'D0001', 'Dr. Sharma', 'MBBS', 'CRM', 'Cardiology', '9876543210', 
            'Chennai', 'Apollo Hospital', '12th Street, Adyar', 'Prima', 'HQ001', 
            'Team A', 'South', 'Tamil Nadu', 'Zone 1', 'Apollo Pharmacy', 'Active'
        ]
    },

        team: {
            title: 'Team Master',
            doctype: 'Team Master',
            fields: [
                { name: 'team_name', label: 'Team Name', type: 'text', required: true },
                { name: 'region', label: 'Region', type: 'link', options: 'Region Master', required: true },
                { name: 'division', label: 'Division', type: 'link', options: 'Division Master', required: true },
                { name: 'sanctioned_strength', label: 'Sanctioned Strength', type: 'number', required: false },
                { name: 'status', label: 'Status', type: 'select', options: ['Active','Inactive'], required: true }
            ],
            columns: ['team_name','region','division','sanctioned_strength','status'],
            searchFields: ['team_name', 'region'],
            excelColumns: ['Team Name', 'Region', 'Division', 'Sanctioned Strength', 'Status'],
            excelSample: ['Team Alpha', 'South', 'Prima', 10, 'Active']
        },
        region: {
            title: 'Region Master',
            doctype: 'Region Master',
            fields: [
                { name: 'region_name', label: 'Region Name', type: 'text', required: true },  // FIXED: only region_name
                { name: 'zone', label: 'Zone', type: 'text' },
                { name: 'state', label: 'State', type: 'text' },
                { name: 'status', label: 'Status', type: 'select', options: ['Active', 'Inactive'], required: true }
            ],
            columns: ['region_name','zone','state','division','status'],
            searchFields: ['region_name','zone','state'],
            excelColumns: ['Region Name', 'Zone', 'State', 'Division', 'Status'],
            excelSample: ['South Region', 'Zone A', 'Tamil Nadu', 'Prima', 'Active']
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

        // Dropdown filters (status, division, etc.)
        let passFilters = true;

        $('.filter-control').each(function () {
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

    // Update the buildForm function to handle readonly name display
    function buildForm(config, data) {
        let html = '<div class="row">';
        
        config.fields.forEach((field, index) => {
            // Skip column break fields
            if (field.type === 'column_break' || field.fieldtype === 'Column Break') {
                return;
            }
            
            // Skip division field (it's session-based)
            if (field.name === 'division') {
                return;
            }
            
            // For edit mode, show the code as readonly info (for HQ, Team, Region)
            if (data && field.name === 'name' && ['hq', 'team', 'region'].includes(currentMasterType)) {
                html += `<div class="col-md-6 mb-3">`;
                html += `<label>Code</label>`;
                html += `<input type="text" class="form-control" value="${data.name}" readonly>`;
                html += `</div>`;
                return;
            }
            
            // Skip 'name' field for new records (autoname will generate it)
            if (!data && field.name === 'name') {
                return;
            }
            
            // Skip readonly code fields for new records (stockist, product, doctor)
            const isCodeField = ['stockist_code'].includes(field.name);
            if (!data && isCodeField && field.readonly_on_edit) {
                return;
            }
            
            const value = data ? (data[field.name] || '') : 
                        (field.name === 'status' ? 'Active' : '');
            
            const readonly = (data && field.readonly_on_edit) || (data && isCodeField) ? 'readonly' : '';
            
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
        });
        
        html += '</div>';
$('#form-fields').html(html);
setTimeout(() => hydrateLinkLabels(), 50);
    }
    async function hydrateLinkLabels() {
    const fields = document.querySelectorAll('.link-field');

    for (const el of fields) {
        const val = el.value;
        const doctype = el.getAttribute('data-doctype');
        if (!val || !doctype) continue;

        try {
            const res = await fetch('/api/method/frappe.client.get', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'X-Frappe-CSRF-Token': window.csrf_token
                },
                body: JSON.stringify({
                    doctype,
                    name: val
                })
            });

            const data = await res.json();
            const doc = data.message;

            const label =
                doc?.hq_name ||
                doc?.team_name ||
                doc?.region_name ||
                doc?.stockist_name ||
                doc?.doctor_name ||
                val;

            el.value = label;
            el.setAttribute("data-value", val);
        } catch (e) {
            console.warn("Hydrate failed", e);
        }
    }
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
            
const input = $(`[name="${field.name}"]`);
let value = input.val();

if (field.type === 'link') {
    const real = input.attr("data-value");
    if (real) value = real;
}

if (value) data[field.name] = value;

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
$(document).ready(function () {

    $(document).on('input', '.link-field', debounce(async function () {

        const el = this;
        const doctype = el.getAttribute('data-doctype');
        const search = el.value;

        if (!doctype || !search) return;

        const res = await fetch('/api/method/scanify.api.portal_link_search', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'X-Frappe-CSRF-Token': window.csrf_token
            },
            body: JSON.stringify({ doctype, search })
        });

        const data = await res.json();
        const list = data.message || [];

        $('.link-dropdown').remove();
        if (!list.length) return;

        const menu = $('<div class="link-dropdown"></div>');

        // ⭐ ADD THIS BLOCK (POSITION FIX)
        const rect = el.getBoundingClientRect();
        const parent = $(el).parent();

        menu.css({
            position: 'absolute',
            top: el.offsetTop + el.offsetHeight + 4,
            left: el.offsetLeft,
            width: el.offsetWidth,
            zIndex: 9999
        });
        // ⭐ END FIX

        list.forEach(r => {
            const item = $(`<div>${r.label}</div>`);

            item.on('click', function () {
                el.value = r.label;
                el.setAttribute("data-value", r.value);
                menu.remove();
            });

            menu.append(item);
        });

        // Attach inside same container (important for modals)
        $(el).parent().css('position', 'relative');
        $(el).after(menu);

    }, 300));

});


$(document).ready(function () {

    // 1. Autocomplete Logic for Link Fields
    $(document).on('input', '.link-field', debounce(async function () {
        const el = this;
        const doctype = el.getAttribute('data-doctype');
        const search = el.value;

        // Clear data-value if user clears input or types new text
        // This ensures invalid manual entries don't keep old IDs
        if (search === '') {
            el.setAttribute("data-value", "");
        }

        if (!doctype || !search) return;

        try {
            const res = await fetch('/api/method/scanify.api.portal_link_search', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'X-Frappe-CSRF-Token': window.csrf_token
                },
                body: JSON.stringify({ doctype, search })
            });

            const data = await res.json();
            const list = data.message || [];

            $('.link-dropdown').remove();
            if (!list.length) return;

            const menu = $('<div class="link-dropdown"></div>');

            // Position the dropdown correctly
            const rect = el.getBoundingClientRect();
            // We append to parent to move with scroll, but use absolute positioning relative to input
            menu.css({
                position: 'absolute',
                top: el.offsetTop + el.offsetHeight + 4,
                left: el.offsetLeft,
                width: el.offsetWidth,
                zIndex: 9999,
                background: 'white',
                border: '1px solid #ddd',
                borderRadius: '4px',
                boxShadow: '0 4px 6px rgba(0,0,0,0.1)',
                maxHeight: '200px',
                overflowY: 'auto'
            });

            list.forEach(r => {
                const item = $(`<div style="padding:8px 12px;cursor:pointer;border-bottom:1px solid #eee;">${r.label}</div>`);
                
                // Hover effect
                item.hover(
                    function() { $(this).css('background-color', '#f8f9fa'); }, 
                    function() { $(this).css('background-color', 'transparent'); }
                );

                item.on('click', function (e) {
                    e.preventDefault();
                    e.stopPropagation(); // Prevent bubbling causing issues
                    
                    // 1. Set Display Value
                    el.value = r.label;
                    
                    // 2. Set Actual ID
                    el.setAttribute("data-value", r.value);
                    
                    // 3. Remove Menu
                    menu.remove();
                    
                    // 4. CRITICAL: Trigger 'change' event so listeners know value updated
                    // This triggers the auto-fill logic immediately
                    $(el).trigger('change'); 
                });

                menu.append(item);
            });

            // Ensure parent is relative so absolute positioning works
            $(el).parent().css('position', 'relative');
            $(el).after(menu);

        } catch (e) {
            console.error("Link search error:", e);
        }

    }, 300));

    // Close dropdown when clicking outside
    $(document).on('click', function (e) {
        if (!$(e.target).closest('.link-field, .link-dropdown').length) {
            $('.link-dropdown').remove();
        }
    });

    // 2. Auto-Fill Logic (Region -> Zone/State)
    // Changed from 'blur' to 'change' to support Dropdown Selection
    $(document).on('change', '[name="region"]', async function() {
        const region = this.getAttribute("data-value");
        
        // If cleared, clear dependent fields
        if (!region) {
            $('[name="zone"]').val('');
            $('[name="state"]').val('');
            return;
        }

        try {
            const res = await fetch('/api/method/frappe.client.get', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'X-Frappe-CSRF-Token': window.csrf_token
                },
                body: JSON.stringify({
                    doctype: 'Region Master',
                    name: region
                })
            });

            const data = await res.json();
            const r = data.message;

            if (r) {
                // Auto-fill Zone
                if ($('[name="zone"]').length) {
                     $('[name="zone"]').val(r.zone || '');
                }
                
                // Auto-fill State (if field exists)
                if ($('[name="state"]').length) {
                    $('[name="state"]').val(r.state || '');
                }
            }
        } catch (e) {
            console.warn("Failed to fetch region details", e);
        }
    });

    // Add other field dependencies here if needed (e.g. Team -> HQ)
    /*
    $(document).on('change', '[name="team"]', async function() {
        // ... similar logic for Team ...
    });
    */
});



$(document).on('click', function (e) {
    if (!$(e.target).closest('.link-field, .link-dropdown').length) {
        $('.link-dropdown').remove();
    }
});

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

$(document).on('blur','[name="region"]', async function(){

    const region = this.getAttribute("data-value");
    if (!region) return;

    const res = await fetch('/api/method/frappe.client.get', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'X-Frappe-CSRF-Token': window.csrf_token
        },
        body: JSON.stringify({
            doctype: 'Region Master',
            name: region
        })
    });

    const data = await res.json();
    const r = data.message;

    if (r) {
        $('[name="zone"]').val(r.zone || '');
        $('[name="state"]').val(r.state || '');
    }
});




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
    return function (...args) {
        const context = this;
        clearTimeout(timeout);
        timeout = setTimeout(() => {
            func.apply(context, args);
        }, wait);
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


let currentMasterType = 'hq';
let currentDivision = '';
let allRecords = [];
let filteredRecords = [];
let currentPage = 1;
const PAGE_SIZE = 25;

// Single master mode: when this page only shows one master type (e.g. Division Master page)
let singleMasterMode = (typeof window.MASTERS_SINGLE_MODE !== 'undefined') ? window.MASTERS_SINGLE_MODE : null;

// Master configurations
const masterConfigs = {
    hq: {
        title: 'HQ Master',
        doctype: 'HQ Master',
        fields: [
            { name: 'hq_code', label: 'HQ Code', type: 'text', system_generated: true, readonly_on_edit: true },
            { name: 'hq_name', label: 'HQ Name', type: 'text', required: true },
            // team_select: dropdown filtered by division; selecting team auto-fills region & zone
            { name: 'team', label: 'Team', type: 'team_select', required: true },
            { name: 'region', label: 'Region', type: 'region_select', required: true, readonly_always: true, help: 'Auto-filled from Team' },
            { name: 'zone', label: 'Zone', type: 'text', readonly_always: true, help: 'Auto-filled from Team\'s Region' },
            { name: 'per_capita', label: 'Per Capita', type: 'number', required: false },
            { name: 'status', label: 'Status', type: 'select', options: ['Active', 'Inactive'], required: true },
            { name: 'team_label', label: 'Team', type: 'display_only' },
            { name: 'region_label', label: 'Region', type: 'display_only' },
            { name: 'zone_label', label: 'Zone', type: 'display_only' }
        ],
        columns: ['hq_code', 'hq_name', 'team_label', 'region_label', 'zone_label', 'per_capita', 'status'],
        searchFields: ['hq_code', 'hq_name', 'team_label', 'region_label', 'zone_label'],
        excelColumns: ['HQ Name', 'Team', 'Region', 'Zone', 'Per Capita', 'Status'],
        excelSample: ['Chennai Central', 'Team A', 'South', 'Zone A', 2, 'Active']
    },
    stockist: {
        title: 'Stockist Master',
        doctype: 'Stockist Master',
        fields: [
            { name: 'stockist_code', label: 'Stockist Code', type: 'text', readonly_on_edit: true, system_generated: true },
            { name: 'stockist_name', label: 'Stockist Name', type: 'text', required: true },
            // Hierarchy - HQ is a dropdown, others auto-fill
            { name: 'hq', label: 'HQ', type: 'hq_select', required: true },
            { name: 'team', label: 'Team', type: 'text', readonly_always: true, help: 'Auto-filled from HQ' },
            { name: 'region', label: 'Region', type: 'text', readonly_always: true, help: 'Auto-filled from HQ' },
            { name: 'zone', label: 'Zone', type: 'text', readonly_always: true, help: 'Auto-filled from HQ' },

            // Address
            { name: 'address', label: 'Address', type: 'textarea' },

            // Contact
            { name: 'contact_person', label: 'Contact Person', type: 'text' },
            { name: 'phone', label: 'Phone', type: 'text' },
            { name: 'email', label: 'Email', type: 'email' },

            { name: 'status', label: 'Status', type: 'select', options: ['Active', 'Inactive'], required: true },
            { name: 'hq_label', label: 'HQ', type: 'display_only' },
            { name: 'team_label', label: 'Team', type: 'display_only' },
            { name: 'region_label', label: 'Region', type: 'display_only' },
            { name: 'zone_label', label: 'Zone', type: 'display_only' }
        ],
        columns: ['stockist_code', 'stockist_name', 'hq_label', 'team_label', 'region_label', 'zone_label', 'phone', 'status'],
        searchFields: ['stockist_code', 'stockist_name', 'city', 'phone'],
        excelColumns: ['Stockist Name', 'HQ', 'Team', 'Region', 'Zone', 'Address', 'Contact Person', 'Phone', 'Email', 'Status'],
        excelSample: ['ABC Pharma Distributors', 'Chennai Central', 'Team A', 'South', 'Zone 1', '123 Main St', 'John Smith', '9876543210', 'abc@example.com', 'Active']
    },
    product: {
        title: 'Product Master',
        doctype: 'Product Master',
        fields: [
            { name: 'product_code', label: 'Product Code', type: 'text', required: true, readonly_on_edit: true },
            { name: 'product_name', label: 'Product Name', type: 'text', required: true },
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
            { name: 'pack_conversion', label: 'Pack Conversion', type: 'text' },
            { name: 'mrp', label: 'MRP', type: 'number' },
            { name: 'ptr', label: 'PTR', type: 'number' },
            { name: 'pts', label: 'PTS', type: 'number' },
            { name: 'gst_rate', label: 'GST Rate (%)', type: 'number', default: 5 },
            { name: 'status', label: 'Status', type: 'select', options: ['Active', 'Inactive'], required: true }
        ],
        columns: ['product_code', 'product_name', 'product_group', 'category', 'pack', 'pack_conversion', 'pts', 'ptr', 'mrp', 'gst_rate', 'status'],
        searchFields: ['product_code', 'product_name', 'product_group', 'category', 'pack'],
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
            // doctor_code is system-generated: shown readonly in edit, hidden in new
            { name: 'doctor_code', label: 'Doctor Code', type: 'text', system_generated: true, readonly_on_edit: true },
            { name: 'doctor_name', label: 'Doctor Name', type: 'text', required: true },
            { name: 'qualification', label: 'Qualification', type: 'text' },
            { name: 'doctor_category', label: 'Doctor Category', type: 'select', options: ['CRM', 'ASM KBL', 'SO – TOP GYANE'] },
            { name: 'specialization', label: 'Specialization', type: 'text' },
            { name: 'phone', label: 'Phone', type: 'text' },
            { name: 'place', label: 'Place', type: 'text' },
            { name: 'hospital_address', label: 'Hospital Address', type: 'textarea' },
            { name: 'house_address', label: 'House Address', type: 'textarea' },
            { name: 'hq', label: 'HQ', type: 'hq_select' },
            { name: 'team', label: 'Team', type: 'text', readonly_always: true, help: 'Auto-filled from HQ' },
            { name: 'region', label: 'Region', type: 'text', readonly_always: true, help: 'Auto-filled from HQ' },
            { name: 'state', label: 'State', type: 'state_select' },
            { name: 'zone', label: 'Zone', type: 'text', readonly_always: true, help: 'Auto-filled from HQ' },
            { name: 'chemist_name', label: 'Chemist Name', type: 'text' },
            { name: 'status', label: 'Status', type: 'select', options: ['Active', 'Inactive'], required: true },
            { name: 'hq_label', label: 'HQ', type: 'display_only' },
            { name: 'state_label', label: 'State', type: 'display_only' }
        ],
        columns: ['doctor_code', 'doctor_name', 'place', 'specialization', 'hq_label', 'state_label', 'phone', 'status'],
        searchFields: ['doctor_code', 'doctor_name', 'place', 'phone', 'chemist_name', 'state_label'],
        excelColumns: [
            'Doctor Name', 'Qualification', 'Doctor Category',
            'Specialization', 'Phone', 'Place', 'Hospital Address', 'House Address',
            'Division', 'HQ', 'Team', 'Region', 'State', 'Zone', 'Chemist Name', 'Status'
        ],
        excelSample: [
            'Dr. Sharma', 'MBBS', 'CRM', 'Cardiology', '9876543210',
            'Chennai', 'Apollo Hospital', '12th Street, Adyar', 'Prima', 'Chennai Central',
            'Team A', 'South', 'Tamil Nadu', 'Zone 1', 'Apollo Pharmacy', 'Active'
        ]
    },

    team: {
        title: 'Team Master',
        doctype: 'Team Master',
        fields: [
            { name: 'team_code', label: 'Team Code', type: 'text', system_generated: true, readonly_on_edit: true },
            { name: 'team_name', label: 'Team Name', type: 'text', required: true },
            { name: 'region', label: 'Region', type: 'region_select', required: true },
            // sanctioned_strength is auto-calculated from HQ per_capita — always readonly
            { name: 'sanctioned_strength', label: 'Sanctioned Strength', type: 'number', readonly_always: true, help: 'Auto-calculated from sum of HQ Per Capita' },
            { name: 'status', label: 'Status', type: 'select', options: ['Active', 'Inactive'], required: true },
            { name: 'region_label', label: 'Region', type: 'display_only' }
        ],
        columns: ['team_code', 'team_name', 'division', 'region_label', 'sanctioned_strength', 'status'],
        searchFields: ['team_code', 'team_name', 'region_label'],
        excelColumns: ['Team Name', 'Region', 'Division', 'Status'],
        excelSample: ['Team Alpha', 'R0001', 'Prima', 'Active']
    },
    region: {
        title: 'Region Master',
        doctype: 'Region Master',
        fields: [
            { name: 'region_code', label: 'Region Code', type: 'text', system_generated: true, readonly_on_edit: true },
            { name: 'region_name', label: 'Region Name', type: 'text', required: true },
            { name: 'division', label: 'Division', type: 'link', options: 'Division', required: true },
            { name: 'zone', label: 'Zone', type: 'zone_select' },
            { name: 'state', label: 'State', type: 'state_select' },
            { name: 'status', label: 'Status', type: 'select', options: ['Active', 'Inactive'], required: true },
            // Virtual display-only fields (resolved by API, not in form)
            { name: 'zone_label', label: 'Zone', type: 'display_only' },
            { name: 'state_label', label: 'State', type: 'display_only' }
        ],
        // name = auto-code R0001; title_field means links display region_name
        columns: ['region_code', 'region_name', 'division', 'zone_label', 'state_label', 'status'],
        searchFields: ['region_code', 'region_name', 'zone_label', 'state_label'],
        excelColumns: ['Region Name', 'Division', 'Zone', 'State', 'Status'],
        excelSample: ['South Region', 'Prima', 'Z0001', 'ST0001', 'Active']
    },
    division: {
        title: 'Division Master',
        doctype: 'Division',
        fields: [
            { name: 'division_name', label: 'Division Name', type: 'text', required: true }
        ],
        columns: ['division_name'],
        searchFields: ['division_name'],
        excelColumns: ['Division Name'],
        excelSample: ['Prima']
    },
    zone: {
        title: 'Zone Master',
        doctype: 'Zone Master',
        fields: [
            { name: 'zone_code', label: 'Zone Code', type: 'text', system_generated: true, readonly_on_edit: true },
            { name: 'zone_name', label: 'Zone Name', type: 'text', required: true },
            { name: 'status', label: 'Status', type: 'select', options: ['Active', 'Inactive'], required: true }
        ],
        columns: ['zone_code', 'zone_name', 'division', 'status'],
        searchFields: ['zone_code', 'zone_name'],
        excelColumns: ['Zone Name', 'Division', 'Status'],
        excelSample: ['Zone A', 'Prima', 'Active']
    },
    state: {
        title: 'State Master',
        doctype: 'State Master',
        fields: [
            { name: 'state_code', label: 'State Code', type: 'text', system_generated: true, readonly_on_edit: true },
            { name: 'state_name', label: 'State Name', type: 'text', required: true },
            { name: 'status', label: 'Status', type: 'select', options: ['Active', 'Inactive'], required: true }
        ],
        columns: ['state_code', 'state_name', 'division', 'status'],
        searchFields: ['state_code', 'state_name'],
        excelColumns: ['State Name', 'Division', 'Status'],
        excelSample: ['Maharashtra', 'Prima', 'Active']
    }
};

// HQ data cache for dropdown (loaded per division)
let hqCache = [];

// Region data cache for dropdown (loaded per division)
let regionCache = [];

// Team data cache for dropdown (loaded per division)
let teamCache = [];

// Zone data cache (per division — zone_code is the value, zone_name is the label)
let zoneCache = [];

// State data cache (per division — state_code is the value, state_name is the label)
let stateCache = [];

window.addEventListener('load', function () {
    // Get current division from the page
    const divisionText = document.querySelector('#divisionMenuButton .division-name');
    currentDivision = divisionText ? divisionText.textContent.trim() : 'Prima';

    // Single master mode: a page can set window.MASTERS_SINGLE_MODE before this script loads
    if (singleMasterMode) {
        currentMasterType = singleMasterMode;
        // Hide the master type selector card if it exists
        const masterTypeSelect = document.getElementById('master-type');
        if (masterTypeSelect) {
            const card = masterTypeSelect.closest('.card');
            if (card) card.style.display = 'none';
        }
    }

    // Pre-load all dropdown caches
    loadHQCache();
    loadRegionCache();
    loadTeamCache();
    loadZoneCache();
    loadStateCache();

    // Load initial data
    loadMasterData();

    // Event listeners (safe — elements may not exist in single master mode)
    const masterTypeEl = document.getElementById('master-type');
    if (masterTypeEl) {
        masterTypeEl.addEventListener('change', function () {
            currentMasterType = this.value;
            loadMasterData();
        });
    }

    const searchEl = document.getElementById('search-input');
    if (searchEl) {
        searchEl.addEventListener('input', debounce(function () {
            filterRecords();
        }, 300));
    }
});

function loadHQCache() {
    $.ajax({
        url: '/api/method/scanify.api.get_hq_list',
        type: 'POST',
        contentType: 'application/json',
        headers: { 'X-Frappe-CSRF-Token': window.csrf_token },
        data: JSON.stringify({ division: currentDivision }),
        success: function (r) {
            if (r.message && r.message.success) {
                hqCache = r.message.data;
            }
        }
    });
}

function loadRegionCache() {
    $.ajax({
        url: '/api/method/scanify.api.get_region_list',
        type: 'POST',
        contentType: 'application/json',
        headers: { 'X-Frappe-CSRF-Token': window.csrf_token },
        data: JSON.stringify({ division: currentDivision }),
        success: function (r) {
            if (r.message && r.message.success) {
                regionCache = r.message.data;
                populateRegionDropdown(null);
            }
        }
    });
}

function loadTeamCache() {
    $.ajax({
        url: '/api/method/scanify.api.get_team_list',
        type: 'POST',
        contentType: 'application/json',
        headers: { 'X-Frappe-CSRF-Token': window.csrf_token },
        data: JSON.stringify({ division: currentDivision }),
        success: function (r) {
            if (r.message && r.message.success) {
                teamCache = r.message.data;
                populateTeamDropdown(null);
            }
        }
    });
}

function loadZoneCache() {
    $.ajax({
        url: '/api/method/scanify.api.get_zone_list',
        type: 'POST',
        contentType: 'application/json',
        headers: { 'X-Frappe-CSRF-Token': window.csrf_token },
        data: JSON.stringify({ division: currentDivision }),
        success: function (r) {
            if (r.message && r.message.success) {
                zoneCache = r.message.data;
            }
        }
    });
}

function loadStateCache() {
    $.ajax({
        url: '/api/method/scanify.api.get_state_list',
        type: 'POST',
        contentType: 'application/json',
        headers: { 'X-Frappe-CSRF-Token': window.csrf_token },
        data: JSON.stringify({ division: currentDivision }),
        success: function (r) {
            if (r.message && r.message.success) {
                stateCache = r.message.data;
            }
        }
    });
}

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
        success: function (response) {
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
        error: function (xhr, status, error) {
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
    // Filters removed: division is session-controlled and search is sufficient.
    $('#filter-controls').empty().hide();
}

// Render table with pagination
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
        renderPagination(0);
        return;
    }

    const totalRecords = filteredRecords.length;
    const totalPages = Math.ceil(totalRecords / PAGE_SIZE);
    if (currentPage > totalPages) currentPage = totalPages;
    if (currentPage < 1) currentPage = 1;
    const start = (currentPage - 1) * PAGE_SIZE;
    const pageRecords = filteredRecords.slice(start, start + PAGE_SIZE);

    let html = '';
    pageRecords.forEach(record => {
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
    $('#record-count').text(totalRecords);
    renderPagination(totalRecords);
}

function renderPagination(totalRecords) {
    let container = document.getElementById('masters-pagination');
    if (!container) {
        container = document.createElement('div');
        container.id = 'masters-pagination';
        container.className = 'card-footer d-flex justify-content-between align-items-center';
        const cardBody = document.querySelector('#data-table');
        if (cardBody && cardBody.parentNode) cardBody.parentNode.appendChild(container);
    }
    if (totalRecords === 0) { container.innerHTML = ''; return; }
    const totalPages = Math.ceil(totalRecords / PAGE_SIZE);
    const start = (currentPage - 1) * PAGE_SIZE + 1;
    const end = Math.min(currentPage * PAGE_SIZE, totalRecords);
    let html = `<small class="text-muted">Showing ${start}-${end} of ${totalRecords}</small>`;
    html += '<nav><ul class="pagination pagination-sm mb-0">';
    html += `<li class="page-item ${currentPage === 1 ? 'disabled' : ''}"><a class="page-link" href="#" onclick="goToMastersPage(${currentPage - 1});return false;">&laquo;</a></li>`;
    let startPage = Math.max(1, currentPage - 2);
    let endPage = Math.min(totalPages, startPage + 4);
    if (endPage - startPage < 4) startPage = Math.max(1, endPage - 4);
    for (let i = startPage; i <= endPage; i++) {
        html += `<li class="page-item ${i === currentPage ? 'active' : ''}"><a class="page-link" href="#" onclick="goToMastersPage(${i});return false;">${i}</a></li>`;
    }
    html += `<li class="page-item ${currentPage === totalPages ? 'disabled' : ''}"><a class="page-link" href="#" onclick="goToMastersPage(${currentPage + 1});return false;">&raquo;</a></li>`;
    html += '</ul></nav>';
    container.innerHTML = html;
}

function goToMastersPage(page) {
    const totalPages = Math.ceil(filteredRecords.length / PAGE_SIZE);
    if (page < 1 || page > totalPages) return;
    currentPage = page;
    renderTable();
    document.getElementById('data-table').scrollIntoView({ behavior: 'smooth', block: 'start' });
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

        return true;
    });

    currentPage = 1;
    renderTable();
}


// Reset filters
function resetFilters() {
    $('#search-input').val('');
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

// Build form — handles hq_select type for stockist/doctor HQ, and system_generated code fields
function buildForm(config, data) {
    let html = '<div class="row">';

    config.fields.forEach((field, index) => {
        // Skip column break fields
        if (field.type === 'column_break' || field.fieldtype === 'Column Break') {
            return;
        }

        // Skip virtual display-only fields (used only for list column headers)
        if (field.type === 'display_only') {
            return;
        }

        // Skip division field (it's session-based)
        if (field.name === 'division') {
            return;
        }

        // System-generated code fields: show readonly info on edit, skip entirely on new
        if (field.system_generated) {
            if (data) {
                // Edit mode: show current value as readonly
                html += `<div class="col-md-6 mb-3">`;
                html += `<label>${field.label} <small class="text-muted">(system generated)</small></label>`;
                html += `<input type="text" class="form-control" value="${data[field.name] || ''}" readonly style="background:#f8f9fa;color:#6c757d;">`;
                html += `</div>`;
            }
            // For new records, skip entirely — backend auto-generates
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

        // Skip 'name' field for new records when it is auto-generated.
        // Division master uses the name field as the user-entered primary field, so do not skip there.
        if (!data && field.name === 'name' && currentMasterType !== 'division') {
            return;
        }

        const value = data ? (data[field.name] || '') :
            (field.name === 'status' ? 'Active' : (field.default || ''));

        const isReadonly = field.readonly_always || (data && field.readonly_on_edit);
        const readonly = isReadonly ? 'readonly' : '';
        const readonlyStyle = isReadonly ? 'style="background:#f8f9fa;color:#6c757d;"' : '';

        html += `<div class="col-md-6 mb-3">`;
        const helpText = field.help ? `<small class="text-info"><i class="fa fa-magic"></i> ${field.help}</small>` : '';
        html += `<label>${field.label}${field.required ? '<span class="text-danger">*</span>' : ''}</label>`;

        if (field.type === 'hq_select') {
            // HQ dropdown — auto-fills team/region/zone (Stockist/Doctor forms)
            html += `<select class="form-control hq-select-field" name="${field.name}" data-value="${value}" ${field.required ? 'required' : ''}>`;
            html += `<option value="">-- Select HQ --</option>`;
            html += `</select>`;
            if (helpText) html += helpText;
        } else if (field.type === 'team_select') {
            // Team dropdown — filtered by division; selecting auto-fills Region + Zone (HQ form)
            html += `<select class="form-control team-select-field" name="${field.name}" data-value="${value}" ${field.required ? 'required' : ''}>`;
            html += `<option value="">-- Select Team --</option>`;
            html += `</select>`;
            if (helpText) html += helpText;
        } else if (field.type === 'region_select') {
            // Region dropdown — shows region_name, stores composite key; disabled when readonly
            const disabledAttr = field.readonly_always ? 'disabled style="background:#f8f9fa;color:#6c757d;pointer-events:none;"' : '';
            html += `<select class="form-control region-select-field" name="${field.name}" data-value="${value}" ${field.required && !field.readonly_always ? 'required' : ''} ${disabledAttr}>`;
            html += `<option value="">-- Select Region --</option>`;
            html += `</select>`;
            if (helpText) html += helpText;
        } else if (field.type === 'zone_select') {
            // Zone dropdown — sourced from Zone Master
            html += `<select class="form-control zone-select-field" name="${field.name}" data-value="${value}" ${field.required ? 'required' : ''}>`;
            html += `<option value="">-- Select Zone --</option>`;
            html += `</select>`;
            if (helpText) html += helpText;
        } else if (field.type === 'state_select') {
            // State dropdown — sourced from State Master
            html += `<select class="form-control state-select-field" name="${field.name}" data-value="${value}" ${field.required ? 'required' : ''}>`;
            html += `<option value="">-- Select State --</option>`;
            html += `</select>`;
            if (helpText) html += helpText;
        } else if (field.type === 'select') {
            html += `<select class="form-control" name="${field.name}" ${field.required ? 'required' : ''} ${readonly} ${readonlyStyle}>`;
            if (Array.isArray(field.options)) {
                field.options.forEach(opt => {
                    html += `<option value="${opt}" ${value == opt ? 'selected' : ''}>${opt}</option>`;
                });
            }
            html += `</select>`;
        } else if (field.type === 'textarea') {
            html += `<textarea class="form-control" name="${field.name}" rows="2" ${field.required ? 'required' : ''} ${readonly} ${readonlyStyle}>${value}</textarea>`;
        } else if (field.type === 'link') {
            html += `<input type="text" class="form-control link-field" name="${field.name}" data-doctype="${field.options}" value="${value}" ${readonly} ${readonlyStyle} ${field.required ? 'required' : ''}>`;
        } else {
            html += `<input type="${field.type === 'number' ? 'number' : 'text'}" class="form-control" name="${field.name}" value="${value}" ${readonly} ${readonlyStyle} ${field.required ? 'required' : ''} ${field.type === 'number' ? 'step="1"' : ''}>`;
        }
        if (helpText && field.type !== 'hq_select' && field.type !== 'team_select' && field.type !== 'region_select' && field.type !== 'zone_select' && field.type !== 'state_select') html += helpText;

        html += `</div>`;
    });

    html += '</div>';
    $('#form-fields').html(html);

    // Populate all dropdowns
    populateHQDropdown(data);
    populateRegionDropdown(data);
    populateTeamDropdown(data);
    populateZoneDropdown(data);
    populateStateDropdown(data);

    setTimeout(() => hydrateLinkLabels(), 50);
}

function populateHQDropdown(data) {
    const hqSelects = document.querySelectorAll('.hq-select-field');
    if (!hqSelects.length) return;

    // Use cached HQ list
    hqSelects.forEach(select => {
        const currentVal = select.getAttribute('data-value') || (data ? data[select.name] : '');
        select.innerHTML = '<option value="">-- Select HQ --</option>';
        hqCache.forEach(hq => {
            const opt = document.createElement('option');
            opt.value = hq.name;
            opt.textContent = hq.hq_name;
            if (hq.name === currentVal || hq.hq_name === currentVal) {
                opt.selected = true;
            }
            select.appendChild(opt);
        });

        // Trigger auto-fill if editing
        if (currentVal) {
            const hqData = hqCache.find(h => h.name === currentVal || h.hq_name === currentVal);
            if (hqData) {
                fillHQRelatedFields(hqData);
            }
        }
    });

    // Also reload HQ if cache is empty (first load)
    if (!hqCache.length) {
        $.ajax({
            url: '/api/method/scanify.api.get_hq_list',
            type: 'POST',
            contentType: 'application/json',
            headers: { 'X-Frappe-CSRF-Token': window.csrf_token },
            data: JSON.stringify({ division: currentDivision }),
            success: function (r) {
                if (r.message && r.message.success) {
                    hqCache = r.message.data;
                    hqSelects.forEach(select => {
                        const currentVal = select.getAttribute('data-value') || '';
                        select.innerHTML = '<option value="">-- Select HQ --</option>';
                        hqCache.forEach(hq => {
                            const opt = document.createElement('option');
                            opt.value = hq.name;
                            opt.textContent = hq.hq_name;
                            if (hq.name === currentVal || hq.hq_name === currentVal) {
                                opt.selected = true;
                            }
                            select.appendChild(opt);
                        });
                        if (currentVal) {
                            const hqData = hqCache.find(h => h.name === currentVal || h.hq_name === currentVal);
                            if (hqData) fillHQRelatedFields(hqData);
                        }
                    });
                }
            }
        });
    }
}

function populateRegionDropdown(data) {
    const regionSelects = document.querySelectorAll('.region-select-field');
    if (!regionSelects.length) return;

    function fillOneSelect(select) {
        const currentVal = select.getAttribute('data-value') || (data ? data[select.name] : '');
        select.innerHTML = '<option value="">-- Select Region --</option>';
        regionCache.forEach(region => {
            const opt = document.createElement('option');
            opt.value = region.name;            // auto-code: "R0001"
            opt.textContent = region.region_name; // display: "South Region"
            if (region.name === currentVal || region.region_name === currentVal) {
                opt.selected = true;
            }
            select.appendChild(opt);
        });
    }

    if (regionCache.length) {
        regionSelects.forEach(fillOneSelect);
    } else {
        // Cache not ready yet — fetch and populate
        $.ajax({
            url: '/api/method/scanify.api.get_region_list',
            type: 'POST',
            contentType: 'application/json',
            headers: { 'X-Frappe-CSRF-Token': window.csrf_token },
            data: JSON.stringify({ division: currentDivision }),
            success: function (r) {
                if (r.message && r.message.success) {
                    regionCache = r.message.data;
                    regionSelects.forEach(fillOneSelect);
                }
            }
        });
    }
}

function populateTeamDropdown(data) {
    const teamSelects = document.querySelectorAll('.team-select-field');
    if (!teamSelects.length) return;

    function fillOneTeamSelect(select) {
        const currentVal = select.getAttribute('data-value') || (data ? data[select.name] : '');
        select.innerHTML = '<option value="">-- Select Team --</option>';
        teamCache.forEach(team => {
            const opt = document.createElement('option');
            opt.value = team.name;          // auto-code: "T0001"
            opt.textContent = team.team_name; // display: "Team Alpha"
            // Match by auto-code OR by team_name
            if (team.name === currentVal || team.team_name === currentVal) {
                opt.selected = true;
            }
            select.appendChild(opt);
        });

        // If editing and a team is pre-selected, auto-fill region + zone
        if (currentVal) {
            const teamData = teamCache.find(t => t.name === currentVal || t.team_name === currentVal);
            if (teamData) fillTeamRelatedFields(teamData);
        }
    }

    if (teamCache.length) {
        teamSelects.forEach(fillOneTeamSelect);
    } else {
        // Cache not ready — fetch inline
        $.ajax({
            url: '/api/method/scanify.api.get_team_list',
            type: 'POST',
            contentType: 'application/json',
            headers: { 'X-Frappe-CSRF-Token': window.csrf_token },
            data: JSON.stringify({ division: currentDivision }),
            success: function (r) {
                if (r.message && r.message.success) {
                    teamCache = r.message.data;
                    teamSelects.forEach(fillOneTeamSelect);
                }
            }
        });
    }
}

function populateZoneDropdown(data) {
    const zoneSelects = document.querySelectorAll('.zone-select-field');
    if (!zoneSelects.length) return;

    function fillOneZoneSelect(select) {
        const currentVal = select.getAttribute('data-value') || (data ? data[select.getAttribute('name')] : '');
        select.innerHTML = '<option value="">-- Select Zone --</option>';
        zoneCache.forEach(zone => {
            const opt = document.createElement('option');
            opt.value = zone.name;
            opt.textContent = zone.zone_name;
            if (zone.name === currentVal || zone.zone_name === currentVal) {
                opt.selected = true;
            }
            select.appendChild(opt);
        });
    }

    if (zoneCache.length) {
        zoneSelects.forEach(fillOneZoneSelect);
    } else {
        $.ajax({
            url: '/api/method/scanify.api.get_zone_list',
            type: 'POST',
            contentType: 'application/json',
            headers: { 'X-Frappe-CSRF-Token': window.csrf_token },
            data: JSON.stringify({ division: currentDivision }),
            success: function (r) {
                if (r.message && r.message.success) {
                    zoneCache = r.message.data;
                    zoneSelects.forEach(fillOneZoneSelect);
                }
            }
        });
    }
}

function populateStateDropdown(data) {
    const stateSelects = document.querySelectorAll('.state-select-field');
    if (!stateSelects.length) return;

    function fillOneStateSelect(select) {
        const currentVal = select.getAttribute('data-value') || (data ? data[select.getAttribute('name')] : '');
        select.innerHTML = '<option value="">-- Select State --</option>';
        stateCache.forEach(state => {
            const opt = document.createElement('option');
            opt.value = state.name;
            opt.textContent = state.state_name;
            if (state.name === currentVal || state.state_name === currentVal) {
                opt.selected = true;
            }
            select.appendChild(opt);
        });
    }

    if (stateCache.length) {
        stateSelects.forEach(fillOneStateSelect);
    } else {
        $.ajax({
            url: '/api/method/scanify.api.get_state_list',
            type: 'POST',
            contentType: 'application/json',
            headers: { 'X-Frappe-CSRF-Token': window.csrf_token },
            data: JSON.stringify({ division: currentDivision }),
            success: function (r) {
                if (r.message && r.message.success) {
                    stateCache = r.message.data;
                    stateSelects.forEach(fillOneStateSelect);
                }
            }
        });
    }
}

function fillTeamRelatedFields(teamData) {
    // Auto-fill Region select from team's linked region (value = R0001, label = region_name)
    const regionSelect = $('select.region-select-field[name="region"]');
    if (regionSelect.length && teamData.region) {
        regionSelect.val(teamData.region); // auto-code R0001, option text shows region_name
    }

    // Auto-fill Zone:
    //   zone_select dropdown  → use zone code (Z0001) as value
    //   Data text input       → use zone_name (human readable) for display
    const zoneSelect = $('select.zone-select-field[name="zone"]');
    const zoneInput = $('input[name="zone"]');
    if (zoneSelect.length) {
        zoneSelect.val(teamData.zone || '');    // Z0001
    } else if (zoneInput.length) {
        zoneInput.val(teamData.zone_name || teamData.zone || '');
    }
}

function fillHQRelatedFields(hqData) {
    // AUTO-FILL: team — now a team_select dropdown
    const teamSelect = $('select.team-select-field[name="team"]');
    const teamInput = $('input[name="team"]');

    if (teamSelect.length) {
        teamSelect.val(hqData.team || '');
    } else if (teamInput.length) {
        teamInput.val(hqData.team || '');
        teamInput.attr('data-value', hqData.team || '');
        if (hqData.team && teamInput.is('[readonly]')) {
            fetch('/api/method/frappe.client.get', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json', 'X-Frappe-CSRF-Token': window.csrf_token },
                body: JSON.stringify({ doctype: 'Team Master', name: hqData.team })
            }).then(r => r.json()).then(d => {
                if (d.message && d.message.team_name) teamInput.val(d.message.team_name);
            }).catch(() => { });
        }
    }

    // AUTO-FILL: region — region_select dropdown or readonly text field
    const regionSelect = $('select.region-select-field[name="region"]');
    const regionInput = $('input[name="region"]');

    if (regionSelect.length) {
        regionSelect.val(hqData.region || '');
    } else if (regionInput.length) {
        regionInput.val(hqData.region || '');
        regionInput.attr('data-value', hqData.region || '');
        if (hqData.region) {
            fetch('/api/method/frappe.client.get', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json', 'X-Frappe-CSRF-Token': window.csrf_token },
                body: JSON.stringify({ doctype: 'Region Master', name: hqData.region })
            }).then(r => r.json()).then(d => {
                if (d.message && d.message.region_name) regionInput.val(d.message.region_name);
            }).catch(() => { });
        }
    }

    // AUTO-FILL: zone
    if ($('[name="zone"]').length) $('[name="zone"]').val(hqData.zone || '');
}

// Team select change handler — auto-fills Region and Zone
$(document).on('change', '.team-select-field', function () {
    const teamId = this.value;
    if (!teamId) {
        // Clear downstream fields
        $('select.region-select-field').val('');
        $('[name="zone"]').val('');
        return;
    }
    // Find in cache first (enriched with region + zone)
    const teamData = teamCache.find(t => t.name === teamId);
    if (teamData) {
        fillTeamRelatedFields(teamData);
        return;
    }
    // Fallback: fetch from API
    $.ajax({
        url: '/api/method/scanify.api.get_team_details',
        type: 'POST',
        contentType: 'application/json',
        headers: { 'X-Frappe-CSRF-Token': window.csrf_token },
        data: JSON.stringify({ team_name: teamId }),
        success: function (r) {
            if (r.message && r.message.success) {
                fillTeamRelatedFields(r.message.data);
            }
        }
    });
});

// HQ select change handler
$(document).on('change', '.hq-select-field', function () {
    const hqId = this.value;
    if (!hqId) {
        $('[name="team"]').val('');
        $('[name="region"]').val('');
        $('[name="zone"]').val('');
        return;
    }
    // Find in cache first
    const hqData = hqCache.find(h => h.name === hqId);
    if (hqData) {
        fillHQRelatedFields(hqData);
        return;
    }
    // Fallback: fetch from API
    $.ajax({
        url: '/api/method/scanify.api.get_hq_details',
        type: 'POST',
        contentType: 'application/json',
        headers: { 'X-Frappe-CSRF-Token': window.csrf_token },
        data: JSON.stringify({ hq_name: hqId }),
        success: function (r) {
            if (r.message && r.message.success) {
                const d = r.message.data;
                fillHQRelatedFields(d);
            }
        }
    });
});

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
        // Skip display_only virtual fields (not real form inputs)
        if (field.type === 'display_only') {
            return;
        }

        // Skip division field (it's session-based) — except for Division master itself
        if (field.name === 'division' && currentMasterType !== 'division') {
            return;
        }

        // Skip system-generated code fields (backend handles them)
        if (field.system_generated && !recordId) {
            return;
        }

        // Skip readonly_always fields, BUT capture their value for the save payload
        if (field.readonly_always) {
            const el = $(`[name="${field.name}"]`);
            let val;
            if (field.type === 'region_select' || field.type === 'team_select') {
                // Disabled <select>: jQuery .val() still works even when disabled
                val = el.val();
            } else {
                // Text input: prefer data-value (real doc name) over display value
                val = el.attr('data-value') || el.val();
            }
            if (val) data[field.name] = val;
            return;
        }

        // Skip sanctioned_strength entirely — it's auto-computed
        if (field.name === 'sanctioned_strength') {
            return;
        }

        // Skip code fields for new records (auto-generated)
        // Exception: Division master uses 'name' as the user-provided key
        const isCodeField = ['name'].includes(field.name);
        if (!recordId && isCodeField && currentMasterType !== 'division') {
            return;
        }

        let input, value;

        if (field.type === 'hq_select' || field.type === 'region_select' || field.type === 'team_select' || field.type === 'zone_select' || field.type === 'state_select') {
            // Select dropdowns: value is already the internal doc name
            input = $(`[name="${field.name}"]`);
            value = input.val();
        } else {
            input = $(`[name="${field.name}"]`);
            value = input.val();

            if (field.type === 'link') {
                const real = input.attr("data-value");
                if (real) value = real;
            }
        }

        if (value) data[field.name] = value;
    });

    // Validate required fields
    let isValid = true;
    config.fields.forEach(field => {
        if (field.required && !data[field.name] && field.name !== 'division' && !field.system_generated && !field.readonly_always) {
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
            'X-Frappe-CSRF-Token': frappe.csrf_token
        },
        data: JSON.stringify({
            doctype: config.doctype,
            name: recordId,
            data: data
        }),
        success: function (response) {
            hideLoadingOverlay();
            if (response.message && response.message.success) {
                showAlert('Record saved successfully!', 'success');
                $('#editModal').modal('hide');
                loadMasterData();
                // Refresh cache state so next time dropdowns are opened they fetch fresh data
                if (currentMasterType === 'hq') {
                    hqCache = [];
                    // Fallback to loadHQCache if it exists for backwards compatibility
                    if (typeof loadHQCache === 'function') loadHQCache();
                } else if (currentMasterType === 'zone') {
                    zoneCache = [];
                } else if (currentMasterType === 'state') {
                    stateCache = [];
                } else if (currentMasterType === 'region') {
                    regionCache = [];
                } else if (currentMasterType === 'team') {
                    teamCache = [];
                }
            } else {
                showError(response.message ? response.message.message : 'Failed to save record');
            }
        },
        error: function (xhr, status, error) {
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
                    function () { $(this).css('background-color', '#f8f9fa'); },
                    function () { $(this).css('background-color', 'transparent'); }
                );

                item.on('click', function (e) {
                    e.preventDefault();
                    e.stopPropagation();
                    el.value = r.label;
                    el.setAttribute("data-value", r.value);
                    menu.remove();
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
    $(document).on('change', '[name="region"]', async function () {
        const region = this.getAttribute("data-value");

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
                if ($('[name="zone"]').length && !$('[name="zone"]').is('[readonly]')) {
                    $('[name="zone"]').val(r.zone || '');
                }
                if ($('[name="state"]').length) {
                    $('[name="state"]').val(r.state || '');
                }
            }
        } catch (e) {
            console.warn("Failed to fetch region details", e);
        }
    });
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
        success: function (response) {
            hideLoadingOverlay();
            if (response.message && response.message.success) {
                showAlert('Record deleted successfully!', 'success');
                loadMasterData();
                // Refresh HQ cache if HQ was deleted (affects sanctioned strength)
                if (currentMasterType === 'hq') {
                    loadHQCache();
                }
            } else {
                showError(response.message?.message || 'Failed to delete record');
            }
        },
        error: function (xhr) {
            hideLoadingOverlay();
            showError('Error deleting record');
        }
    });
}

$(document).on('blur', '[name="region"]', async function () {

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
        if (!$('[name="zone"]').is('[readonly]')) $('[name="zone"]').val(r.zone || '');
        $('[name="state"]').val(r.state || '');
    }
});




// Show bulk import modal
function showBulkImportModal() {
    const config = masterConfigs[currentMasterType];
    $('#import-master-type').text(config.title);

    // Build sample table — using a div-based approach for Firefox compatibility
    let sampleHtml = `<div class="table-responsive" style="overflow-x:auto;-webkit-overflow-scrolling:touch;">`;
    sampleHtml += `<table class="table table-bordered table-sm sample-excel-table" style="min-width:100%;border-collapse:collapse;">`;
    sampleHtml += `<thead><tr>`;

    config.excelColumns.forEach(col => {
        sampleHtml += `<th style="background:#217346;color:white;font-weight:600;padding:8px 10px;white-space:nowrap;">${col}</th>`;
    });

    sampleHtml += `</tr></thead><tbody><tr>`;

    config.excelSample.forEach(val => {
        sampleHtml += `<td style="padding:8px 10px;font-family:monospace;white-space:nowrap;border:1px solid #ddd;">${val}</td>`;
    });

    sampleHtml += `</tr></tbody></table></div>`;

    $('#sample-table-container').html(sampleHtml);
    $('#import-file').val('');
    $('#import-progress').hide();
    $('#import-results').hide();

    $('#bulkImportModal').modal('show');
}

function downloadSampleTemplate() {
    const config = masterConfigs[currentMasterType];

    // Build CSV content
    const headers = config.excelColumns.join(',');
    const sample = config.excelSample.map(v => {
        // Wrap values with commas or quotes in double quotes
        const str = String(v);
        if (str.includes(',') || str.includes('"') || str.includes('\n')) {
            return `"${str.replace(/"/g, '""')}"`;
        }
        return str;
    }).join(',');

    const csvContent = `${headers}\n${sample}`;
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', `${config.title.replace(/\s+/g, '_')}_template.csv`);
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);

    showAlert(`Downloaded ${config.title} template`, 'success');
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
        xhr: function () {
            const xhr = new window.XMLHttpRequest();
            xhr.upload.addEventListener('progress', function (e) {
                if (e.lengthComputable) {
                    const percent = (e.loaded / e.total) * 100;
                    $('.progress-bar').css('width', percent + '%');
                }
            });
            return xhr;
        },
        success: function (response) {
            $('.progress-bar').css('width', '100%');
            if (response.message && response.message.success) {
                $('#import-alert')
                    .removeClass('alert-danger')
                    .addClass('alert-success')
                    .html(`<strong>Import Successful!</strong><br>
                            <i class="fa fa-check-circle"></i> ${response.message.imported} records imported<br>
                            ${response.message.failed > 0 ? `<i class="fa fa-exclamation-triangle"></i> ${response.message.failed} records failed` : ''}`);

                setTimeout(function () {
                    $('#bulkImportModal').modal('hide');
                    loadMasterData();
                    // Refresh cache state so next time dropdowns are opened they fetch fresh data
                    if (currentMasterType === 'hq') {
                        hqCache = [];
                        if (typeof loadHQCache === 'function') loadHQCache();
                    } else if (currentMasterType === 'zone') {
                        zoneCache = [];
                    } else if (currentMasterType === 'state') {
                        stateCache = [];
                    } else if (currentMasterType === 'region') {
                        regionCache = [];
                    } else if (currentMasterType === 'team') {
                        teamCache = [];
                    }
                }, 2000);
            } else {
                $('#import-alert')
                    .removeClass('alert-success')
                    .addClass('alert-danger')
                    .html(`<strong>Import Failed!</strong><br>${response.message ? response.message.message : 'Unknown error'}`);
            }
            $('#import-results').show();
        },
        error: function () {
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
            <div class="alert custom-toast-alert alert-${type} alert-dismissible fade show" role="alert" style="position:fixed;top:70px;right:20px;z-index:9999;min-width:300px;box-shadow:0 4px 6px rgba(0,0,0,0.1);">
                <strong>${message}</strong>
                <button type="button" class="close" data-dismiss="alert">&times;</button>
            </div>
        `;
    const $alert = $(alert);
    $('body').append($alert);

    setTimeout(function () {
        $alert.fadeOut(function () {
            $(this).remove();
        });
    }, 3000);
}

function showError(message) {
    showAlert(message, 'danger');
}




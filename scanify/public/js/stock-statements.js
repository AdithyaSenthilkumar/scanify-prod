// scanify/public/js/stock-statements.js

let extracted_data = [];
let statement_doc = null;
let uploaded_file_url = null;
let current_zoom = 1;

$(document).ready(function () {
  initialize_page();
});

function get_csrf_token() {
  return window.csrftoken || window.csrf_token || (window.frappe && frappe.csrf_token) || '';
}

/** Get the active division from the navbar division-switcher or a cookie */
function get_active_division() {
  // The division switcher stores the selected division in a cookie or DOM attribute
  try {
    const btn = document.querySelector('.division-name');
    if (btn && btn.textContent.trim()) return btn.textContent.trim();
  } catch (e) { }
  // Fallback – read cookie set by portal_base
  const match = document.cookie.match(/(?:^|;\s*)division=([^;]*)/);
  if (match) return decodeURIComponent(match[1]);
  return '';
}

function initialize_page() {
  const today = new Date();
  $('#statement-month').val(today.toISOString().slice(0, 7));

  init_stockist_search();
  init_file_upload();

  // Check for duplicate when month changes
  $('#statement-month').on('change', function () {
    check_duplicate_statement();
  });

  $('#statement-form').on('submit', async function (e) {
    e.preventDefault();
    await handle_extraction();
  });

  $('#browse-btn').on('click', function (e) {
    e.stopPropagation();
    document.getElementById('statement-file').click();
  });

  $('#clear-file-btn').on('click', function (e) {
    e.stopPropagation();
    clear_file();
  });

  $(document).on('change', '#qc-tbody input, #full-tbody input', function () {
    recalc_row(this);
  });
}

/* ------------ Loader step helpers ------------ */
function set_step(id, state) {
  const el = document.getElementById(id);
  if (!el) return;
  el.classList.remove('active', 'done');
  el.classList.add(state);
}

/* ---------------- Stockist search ---------------- */
function init_stockist_search() {
  let search_timeout;

  $('#stockist-search').on('input focus click', function (e) {
    const term = $(this).val().trim();
    clearTimeout(search_timeout);

    if (e.type === 'input') {
      $('#stockist-code').val('');
      $('#hq,#team,#region,#zone').val('');
      $('#stockist-info').text('');
      $('#stockist-meta-row').hide();
    }

    search_timeout = setTimeout(() => {
      search_stockists(term);
    }, 250);
  });

  $(document).on('click', function (e) {
    if (!$(e.target).closest('#stockist-search, .link-dropdown').length) {
      $('.link-dropdown').remove();
    }
  });
}

function search_stockists(term) {
  $.ajax({
    url: '/api/method/scanify.api.searchstockists',
    type: 'POST',
    contentType: 'application/json',
    headers: { 'X-Frappe-CSRF-Token': get_csrf_token() },
    data: JSON.stringify({
      searchterm: term,
      division: get_active_division()
    }),
    success: function (r) {
      render_stockist_results(r.message || []);
    },
    error: function () {
      $('.link-dropdown').remove();
      show_alert('Failed to search stockists', 'danger');
    }
  });
}

function render_stockist_results(stockists) {
  $('.link-dropdown').remove();
  if (!stockists.length) return;

  const $input = $('#stockist-search');
  const $menu = $('<div class="link-dropdown"></div>');

  stockists.forEach(st => {
    const label = `${st.stockist_code} - ${st.stockist_name}`;
    const meta = [st.hq, st.team, st.region, st.zone].filter(Boolean).join(' | ');

    const $item = $(`
      <div class="autocomplete-item">
        <div class="font-weight-bold">${escape_html(st.stockist_code)} — ${escape_html(st.stockist_name)}</div>
        <small class="text-muted">${escape_html(meta)}</small>
      </div>
    `);

    $item.on('click', async function () {
      $('#stockist-search').val(label);
      $('#stockist-code').val(st.stockist_code);
      $('.link-dropdown').remove();
      await populate_stockist_details(st.stockist_code);
      check_duplicate_statement();
    });

    $menu.append($item);
  });

  const offset = $input.offset();
  $menu.css({
    position: 'absolute',
    top: (offset.top + $input.outerHeight() + 2) + 'px',
    left: offset.left + 'px',
    width: $input.outerWidth() + 'px',
    zIndex: 99999,
    background: '#fff',
    border: '1px solid #ced4da',
    borderRadius: '8px',
    boxShadow: '0 4px 16px rgba(0,0,0,0.12)',
    maxHeight: '260px',
    overflowY: 'auto'
  });

  $('body').append($menu);
}

async function populate_stockist_details(stockist_code) {
  try {
    const details = await call_api('scanify.api.get_stockist_details', { stockist_code });
    $('#hq').val(details.hq || '');
    $('#team').val(details.team || '');
    $('#region').val(details.region || '');
    $('#zone').val(details.zone || '');
    $('#stockist-info').text(`Selected: ${details.stockist_name || stockist_code}`);
    $('#stockist-meta-row').show();
  } catch (e) {
    $('#stockist-info').text('');
    show_alert('Unable to fetch stockist details', 'warning');
  }
}

/* ---------------- Duplicate statement check ---------------- */
function check_duplicate_statement() {
  const stockist_code = $('#stockist-code').val();
  const month = $('#statement-month').val();

  // Clear previous warning
  $('#duplicate-warning').remove();
  $('#extract-btn').prop('disabled', false);

  if (!stockist_code || !month) return;

  $.ajax({
    url: '/api/method/scanify.api.check_statement_exists',
    type: 'POST',
    contentType: 'application/json',
    headers: { 'X-Frappe-CSRF-Token': get_csrf_token() },
    data: JSON.stringify({ stockist_code: stockist_code, statement_month: month }),
    success: function (r) {
      if (r.message && r.message.exists) {
        const warn = `<div id="duplicate-warning" class="alert alert-danger mt-3" role="alert">
          <i class="fa fa-exclamation-triangle"></i>
          <strong>Duplicate Statement!</strong> A statement already exists for this stockist in the selected month:
          <strong>${r.message.statement_name}</strong>.
          Please choose a different stockist or month.
        </div>`;
        $('#stockist-meta-row').after(warn);
        $('#extract-btn').prop('disabled', true);
      }
    }
  });
}

/* ---------------- File upload UI ---------------- */
function init_file_upload() {
  const dropzone = document.getElementById('file-drop-zone');
  const fileInput = document.getElementById('statement-file');

  dropzone.addEventListener('dragover', function (e) {
    e.preventDefault();
    dropzone.classList.add('drag-over');
  });
  dropzone.addEventListener('dragleave', function () {
    dropzone.classList.remove('drag-over');
  });
  dropzone.addEventListener('drop', function (e) {
    e.preventDefault();
    dropzone.classList.remove('drag-over');
    const files = e.dataTransfer.files;
    if (!files || !files.length) return;
    fileInput.files = files;
    show_file_preview(files[0]);
  });

  fileInput.addEventListener('change', function () {
    if (this.files && this.files.length) show_file_preview(this.files[0]);
  });
}

function show_file_preview(file) {
  $('#file-name').text(file.name);
  $('#file-size').text((file.size / 1024).toFixed(2) + ' KB');
  $('#upload-placeholder').hide();
  $('#file-preview').show();
}

function clear_file() {
  document.getElementById('statement-file').value = '';
  $('#file-preview').hide();
  $('#upload-placeholder').show();
}

/* ---------------- Extraction flow ---------------- */
async function handle_extraction() {
  const stockist_code = $('#stockist-code').val();
  const month = $('#statement-month').val();
  const file = document.getElementById('statement-file').files[0];

  if (!stockist_code) return show_alert('Please select a stockist', 'warning');
  if (!month) return show_alert('Please select a statement month', 'warning');
  if (!file) return show_alert('Please upload a statement file', 'warning');

  // Safety guard: check for duplicate before proceeding
  try {
    const dupCheck = await call_api('scanify.api.check_statement_exists', {
      stockist_code: stockist_code,
      statement_month: month
    });
    if (dupCheck && dupCheck.exists) {
      show_alert(`A statement already exists for this stockist in this month: ${dupCheck.statement_name}. Cannot upload duplicate.`, 'danger');
      return;
    }
  } catch (e) {
    console.error('Duplicate check failed:', e);
  }

  $('#entry-card').fadeOut(300);
  setTimeout(() => { $('#extraction-loader').fadeIn(300); }, 300);

  try {
    // Step 1: Upload
    set_step('step-upload', 'active');
    uploaded_file_url = await upload_file(file);
    set_step('step-upload', 'done');

    // Step 2: Create doc
    set_step('step-create', 'active');
    statement_doc = await create_statement_doc(stockist_code, month, uploaded_file_url);
    set_step('step-create', 'done');

    // Step 3: Kick off AI extraction (runs in background thread on server)
    set_step('step-ai', 'active');
    const result = await extract_statement(statement_doc, uploaded_file_url);
    if (!result || !result.success) throw new Error((result && result.message) || 'Extraction failed to start');
    set_step('step-ai', 'done');

    // Extraction is running in the background — redirect to list immediately.
    // The list page shows "In Progress" badge; user can refresh or click View once done.
    $('#extraction-loader').fadeOut(300, () => {
      show_alert('Statement created! AI extraction is processing in the background. Check the list for status.', 'success');
      setTimeout(() => {
        window.location.href = '/portal/stock-statements-list';
      }, 1500);
    });

  } catch (err) {
    console.error(err);
    $('#extraction-loader').fadeOut(300, () => {
      $('#entry-card').fadeIn(300);
    });
    show_alert('Extraction failed: ' + (err.message || err), 'danger');
  }
}

async function upload_file(file) {
  const formData = new FormData();
  formData.append('file', file);
  formData.append('is_private', 1);

  return new Promise((resolve, reject) => {
    $.ajax({
      url: '/api/method/upload_file',
      type: 'POST',
      headers: { 'X-Frappe-CSRF-Token': get_csrf_token() },
      data: formData,
      processData: false,
      contentType: false,
      success: function (r) {
        const url = r.message && r.message.file_url;
        if (!url) return reject(new Error('Upload succeeded but file_url missing'));
        resolve(url);
      },
      error: function () { reject(new Error('File upload failed')); }
    });
  });
}

async function create_statement_doc(stockist_code, month, file_url) {
  const statement_month = month + '-01';

  return new Promise((resolve, reject) => {
    $.ajax({
      url: '/api/method/frappe.client.insert',
      type: 'POST',
      contentType: 'application/json',
      headers: { 'X-Frappe-CSRF-Token': get_csrf_token() },
      data: JSON.stringify({
        doc: {
          doctype: 'Stockist Statement',
          stockist_code: stockist_code,
          statement_month: statement_month,
          uploaded_file: file_url,
          extracted_data_status: 'Pending',
          docstatus: 0
        }
      }),
      success: function (r) {
        if (r.message && r.message.name) resolve(r.message.name);
        else reject(new Error('Failed to create statement'));
      },
      error: function () { reject(new Error('Document creation failed')); }
    });
  });
}

async function extract_statement(doc_name, file_url) {
  // Kick off extraction — returns immediately, extraction runs in background thread
  return call_api('scanify.api.extract_stockist_statement', { doc_name, file_url });
}

/* ---------------- Core Calculation Engine ---------------- */
function get_conversion_factor(pack_str) {
  if (!pack_str) return 1;
  pack_str = String(pack_str).trim().toUpperCase();
  const match = pack_str.match(/(\d+)\s*[xX]\s*(\d+)/);
  if (match) return parseFloat(match[1]);
  return 1;
}

async function enrich_data_with_master_info() {
  const promises = extracted_data.map(async (row) => {
    if (row.productcode) {
      try {
        const product = await call_api('frappe.client.get_value', {
          doctype: 'Product Master',
          fieldname: ['product_name', 'pts', 'ptr', 'pack'],
          filters: { name: row.productcode }
        });
        if (product) {
          row.productname = product.product_name;
          row.pts = parseFloat(product.pts || 0);
          row.ptr = parseFloat(product.ptr || 0);
          row.pack = product.pack;
          row.conversion_factor = get_conversion_factor(row.pack);
        } else {
          row.pts = 0; row.ptr = 0; row.conversion_factor = 1;
        }
      } catch (e) {
        console.error('Product fetch error:', row.productcode);
        row.pts = 0; row.conversion_factor = 1;
      }
    } else {
      row.pts = 0; row.conversion_factor = 1;
    }

    row.openingqty = parseFloat(row.openingqty || 0);
    row.purchaseqty = parseFloat(row.purchaseqty || 0);
    row.salesqty = parseFloat(row.salesqty || 0);
    row.freeqty = parseFloat(row.freeqty || 0);
    row.freeqtyscheme = parseFloat(row.freeqtyscheme || 0);
    row.returnqty = parseFloat(row.returnqty || 0);
    row.miscoutqty = parseFloat(row.miscoutqty || 0);

    do_row_calc(row);
  });
  await Promise.all(promises);
}

function do_row_calc(row) {
  // Called once after OCR enrichment — calculates derived values without
  // overwriting the closing qty/value that the OCR already extracted.
  const cf = row.conversion_factor || 1;
  const pts = row.pts || 0;
  const ptr = row.ptr || 0;

  row.openingvalue  = (row.openingqty  / cf) * pts;
  row.purchasevalue = (row.purchaseqty / cf) * pts;

  // Sales value is based on sales qty only.
  // Free-scheme deductions are handled outside QC by the scheme approval flow.
  row.salesvaluepts = (row.salesqty / cf) * pts;
  row.salesvalueptr = (row.salesqty / cf) * ptr;

  // schemedeductedqty is NOT computed during QC — it is set only when a
  // scheme deduction is applied and approved. Leave whatever OCR gave (0).

  // Closing: OCR may have extracted qty, value, or both.
  // Priority: if qty is present use it to derive value;
  //           else if value is present derive qty from it.
  // Never overwrite with a formula derived from other movements.
  if (row.closingqty > 0) {
    row.closingvalue = (row.closingqty / cf) * pts;
  } else if (row.closingvalue > 0 && pts > 0) {
    row.closingqty = (row.closingvalue / pts) * cf;
  }
}

// Per-field recalculation used during QC editing.
// Only recalculates the value DIRECTLY derived from the changed field.
// Closing qty/value are isolated — only touched when closingqty itself is edited.
// schemedeductedqty is NEVER touched here — it is set only by scheme approval.
function calc_row_field(row, changed_field) {
  const cf = row.conversion_factor || 1;
  const pts = row.pts || 0;
  const ptr = row.ptr || 0;

  switch (changed_field) {
    case 'openingqty':
      row.openingvalue = (row.openingqty / cf) * pts;
      break;
    case 'purchaseqty':
      row.purchasevalue = (row.purchaseqty / cf) * pts;
      break;
    case 'salesqty':
      // Only sales qty drives the sales value columns.
      row.salesvaluepts = (row.salesqty / cf) * pts;
      row.salesvalueptr = (row.salesqty / cf) * ptr;
      break;
    case 'freeqty':
    case 'returnqty':
    case 'miscoutqty':
      // These fields record physical movements but do not independently drive
      // any value column in the QC phase — leave all calculated fields as-is.
      break;
    case 'closingqty':
      // User explicitly corrected the closing qty — recalculate its value only.
      row.closingvalue = (row.closingqty / cf) * pts;
      break;
    default:
      break;
  }
}

function display_results() {
  update_summary_cards();
  render_preview_table();
  $('#extraction-info').text(`${extracted_data.length} products extracted — ${$('#stockist-search').val()}`);
  $('#fs-title').text($('#stockist-search').val());
  $('#results-section').fadeIn(400);
  $('html, body').animate({ scrollTop: $('#results-section').offset().top - 80 }, 400);
}

function update_summary_cards() {
  let op = 0, sa = 0, cl = 0;
  extracted_data.forEach(row => {
    op += parseFloat(row.openingqty || 0);
    sa += parseFloat(row.salesqty || 0);
    cl += parseFloat(row.closingqty || 0);
  });
  $('#total-products').text(extracted_data.length);
  $('#total-opening').text(op.toFixed(0));
  $('#total-sales').text(sa.toFixed(0));
  $('#total-closing').text(cl.toFixed(0));
}

function render_preview_table() {
  const $tbody = $('#preview-tbody');
  $tbody.empty();
  extracted_data.slice(0, 30).forEach(row => {
    $tbody.append(`
      <tr>
        <td><strong>${escape_html(row.productcode || '')}</strong></td>
        <td>${escape_html(row.productname || '')}</td>
        <td>${escape_html(row.pack || '')}</td>
        <td class="text-right">${fmt(row.openingqty)}</td>
        <td class="text-right">${fmt(row.purchaseqty)}</td>
        <td class="text-right">${fmt(row.salesqty)}</td>
        <td class="text-right">${fmt(row.freeqty)}</td>
        <td class="text-right">${fmt(row.closingqty)}</td>
        <td class="text-right">${fmt(row.closingvalue)}</td>
      </tr>
    `);
  });
}

function fmt(v, currency) {
  const n = parseFloat(v || 0);
  if (isNaN(n)) return currency ? '0.00' : '0';
  if (currency) return n.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return Math.round(n).toString();
}

function escape_html(s) {
  return String(s || '').replace(/[&<>"']/g, m => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;'
  }[m]));
}

function show_alert(message, type) {
  if (window.showAlert) {
    window.showAlert(message, type);
  } else {
    // Fallback: simple toast
    const map = { success: '#059669', danger: '#dc2626', warning: '#d97706', info: '#2563eb' };
    const color = map[type] || '#2563eb';
    const $t = $(`<div style="position:fixed;bottom:24px;right:24px;z-index:99999;background:${color};color:white;padding:12px 20px;border-radius:10px;font-size:14px;font-weight:600;box-shadow:0 4px 12px rgba(0,0,0,0.2);">${escape_html(message)}</div>`);
    $('body').append($t);
    setTimeout(() => $t.fadeOut(() => $t.remove()), 3500);
  }
}

function call_api(method, args) {
  return new Promise((resolve, reject) => {
    $.ajax({
      url: '/api/method/' + method,
      type: 'POST',
      contentType: 'application/json',
      headers: { 'X-Frappe-CSRF-Token': get_csrf_token() },
      data: JSON.stringify(args || {}),
      success: function (r) {
        if (r.message === undefined && r.docs === undefined) return reject(new Error('Empty response'));
        resolve(r.message || r.docs);
      },
      error: function (xhr) {
        reject(new Error((xhr.responseJSON && xhr.responseJSON.message) || 'API error'));
      }
    });
  });
}

/* ------------ Table & Modal rendering ------------ */

const col_configs = [
  { id: 'productcode',       label: 'Code',         readonly: true },
  { id: 'productname',       label: 'Product Name', readonly: true },
  { id: 'pack',              label: 'Pack',         readonly: true },
  { id: 'pts',               label: 'PTS',          readonly: true, curr: true },
  { id: 'conversion_factor', label: 'Conv',         readonly: true },
  { id: 'openingqty',        label: 'Opening' },
  { id: 'purchaseqty',       label: 'Purchase' },
  { id: 'salesqty',          label: 'Sales' },
  { id: 'freeqty',           label: 'Free' },
  { id: 'returnqty',         label: 'Return' },
  { id: 'closingqty',        label: 'Closing' },       // editable — OCR-extracted, user can correct
  { id: 'miscoutqty',        label: 'Misc Out' },
  { id: 'freeqtyscheme',     label: 'Free(Sch)',    readonly: true }, // scheme-managed, not QC-editable
  { id: 'schemedeductedqty', label: 'Sch Ded',      readonly: true }, // set only on scheme approval
  { id: 'openingvalue',      label: 'Open Val',     readonly: true, curr: true },
  { id: 'purchasevalue',     label: 'Purch Val',    readonly: true, curr: true },
  { id: 'salesvaluepts',     label: 'Sales(PTS)',   readonly: true, curr: true },
  { id: 'salesvalueptr',     label: 'Sales(PTR)',   readonly: true, curr: true },
  { id: 'closingvalue',      label: 'Cls Val',      readonly: true, curr: true }
];

window.open_fullscreen_table = function () {
  render_full_edit_table();
  $('#fullscreen-modal').modal('show');
};

window.open_qc_screen = function () {
  render_qc_table();
  // Reset pan/zoom state
  qc_viewer_reset();

  // Hide all viewers first
  const $canvas = $('#viewer-canvas');
  $canvas.removeClass('pdf-active');
  $('#pdf-canvas-container').hide().empty();
  $('#img-viewer').hide().attr('src', '');
  $('#spreadsheet-viewer').hide().empty();

  // Show modal first, then initialise viewer once visible so dimensions are available
  $('#qc-modal').off('shown.bs.modal.qcviewer').one('shown.bs.modal.qcviewer', function () {
    if (!uploaded_file_url) {
      show_alert('Original file not available for QC view', 'warning');
      return;
    }
    const ext = uploaded_file_url.split('.').pop().toLowerCase();
    if (ext === 'pdf') {
      render_pdf_to_canvas(uploaded_file_url, 'pdf-canvas-container', 'viewer-canvas', 'document-viewer');
    } else if (['jpg', 'jpeg', 'png'].includes(ext)) {
      const $img = $('#img-viewer');
      $img.off('load').on('load', function () {
        // Size canvas to image natural size
        $canvas.css({ width: this.naturalWidth + 'px', height: this.naturalHeight + 'px' });
        // Auto fit to viewer width
        const viewer = document.getElementById('document-viewer');
        const fitZoom = viewer.clientWidth / this.naturalWidth;
        current_zoom = Math.min(fitZoom, 1);
        pan_x = 0; pan_y = 0;
        apply_viewer_transform();
        init_qc_viewer_events();
      });
      $img.attr('src', uploaded_file_url).show();
    } else if (['xls', 'xlsx', 'csv', 'txt'].includes(ext)) {
      // Render spreadsheet inside the draggable canvas
      var $sv = $('#spreadsheet-viewer');
      $sv.empty().show().html(
        '<div class="text-center py-4"><i class="fa fa-spinner fa-spin fa-2x text-muted"></i><br><small class="text-muted mt-2">Loading spreadsheet…</small></div>'
      );
      call_api('scanify.api.render_spreadsheet_preview', {
        file_url: uploaded_file_url
      }).then(r => {
        if (r && r.html) {
          $sv.html(r.html);
        } else {
          $sv.html('<div class="text-center text-muted py-4"><i class="fa fa-file-alt fa-2x mb-2"></i><br>Could not render preview</div>');
        }
        // Size canvas to the rendered content and enable pan/zoom
        setTimeout(function () {
          var w = $sv[0].scrollWidth || $sv[0].offsetWidth;
          var h = $sv[0].scrollHeight || $sv[0].offsetHeight;
          $canvas.css({ width: w + 'px', height: h + 'px' });
          var viewer = document.getElementById('document-viewer');
          var fitZoom = viewer.clientWidth / w;
          current_zoom = Math.min(fitZoom, 1);
          pan_x = 0; pan_y = 0;
          apply_viewer_transform();
          init_qc_viewer_events();
        }, 50);
      }).catch(() => {
        $sv.html('<div class="text-center text-muted py-4"><i class="fa fa-file-alt fa-2x mb-2"></i><br>Could not render preview</div>');
      });
    }
  });

  $('#qc-modal').modal('show');
};

function render_qc_table() {
  const $thead = $('#qc-thead');
  const $tbody = $('#qc-tbody');
  $thead.empty(); $tbody.empty();

  const active_cols = [];
  $('#column-checkboxes input:checked').each(function () {
    active_cols.push($(this).val());
  });

  let thr = '<tr>';
  col_configs.forEach(c => {
    if (active_cols.includes(c.id)) thr += `<th>${c.label}</th>`;
  });
  thr += '</tr>';
  $thead.html(thr);

  const textCols = ['productcode', 'productname', 'pack'];
  extracted_data.forEach((row, i) => {
    let tr = `<tr data-idx="${i}">`;
    col_configs.forEach(c => {
      if (!active_cols.includes(c.id)) return;
      if (c.readonly) {
        const isText = textCols.includes(c.id);
        const display = c.curr ? fmt(row[c.id], true) : (isText ? escape_html(row[c.id] != null ? row[c.id] : '') : (row[c.id] != null ? row[c.id] : ''));
        const calcAttr = !isText ? ` data-calcfield="${c.id}"` : '';
        tr += `<td class="${isText ? '' : 'text-right'}"${calcAttr}>${display}</td>`;
      } else {
        tr += `<td class="p-0"><input type="number" class="ss-edit-input qc-input" data-col="${c.id}" value="${row[c.id]}" min="0" step="any"></td>`;
      }
    });
    tr += '</tr>';
    $tbody.append(tr);
  });
}

function render_full_edit_table() {
  const $tbody = $('#full-tbody');
  $tbody.empty();

  const ftTextCols = ['productcode', 'productname', 'pack'];
  extracted_data.forEach((row, i) => {
    let tr = `<tr data-idx="${i}">`;
    col_configs.forEach(c => {
      if (c.readonly) {
        const isText = ftTextCols.includes(c.id);
        const display = c.curr ? fmt(row[c.id], true) : escape_html(row[c.id] != null ? row[c.id] : '');
        const calcAttr = !isText ? ` data-calcfield="${c.id}"` : '';
        tr += `<td class="${isText ? '' : 'text-right'}"${calcAttr}>${display}</td>`;
      } else {
        tr += `<td class="p-0"><input type="number" class="ss-edit-input qc-input" data-col="${c.id}" value="${row[c.id]}" min="0" step="any"></td>`;
      }
    });
    tr += '</tr>';
    $tbody.append(tr);
  });
  calc_grand_totals();
}

function recalc_row(input_el) {
  const $tr = $(input_el).closest('tr');
  const idx = $tr.data('idx');
  const row = extracted_data[idx];

  const col = $(input_el).data('col');
  row[col] = parseFloat($(input_el).val() || 0);

  // Recalculate only the derived values affected by this specific field change.
  // Closing qty/value are NEVER recalculated from other movements.
  calc_row_field(row, col);

  // Only refresh the calculated display cells in THIS row that carry
  // a data-calcfield attribute — never touch cells unrelated to the edit.
  $tr.find('[data-calcfield]').each(function () {
    const calcField = $(this).data('calcfield');
    const cfg = col_configs.find(function (c) { return c.id === calcField; });
    $(this).text(fmt(row[calcField], cfg && cfg.curr));
  });

  // Also refresh the matching row in the other table (full-table ↔ QC)
  const otherSelector = $tr.closest('#qc-tbody').length
    ? '#full-tbody tr[data-idx="' + idx + '"]'
    : '#qc-tbody tr[data-idx="' + idx + '"]';
  $(otherSelector).find('[data-calcfield]').each(function () {
    const calcField = $(this).data('calcfield');
    const cfg = col_configs.find(function (c) { return c.id === calcField; });
    $(this).text(fmt(row[calcField], cfg && cfg.curr));
  });

  update_summary_cards();
  calc_grand_totals();
}

function calc_grand_totals() {
  const totals = {};
  col_configs.forEach(c => totals[c.id] = 0);

  extracted_data.forEach(r => {
    col_configs.forEach(c => {
      if (!['productcode', 'productname', 'pack'].includes(c.id)) {
        totals[c.id] += (parseFloat(r[c.id]) || 0);
      }
    });
  });

  let trHtml = '<td colspan="3" class="text-right font-weight-bold">TOTALS</td>';
  col_configs.slice(3).forEach(c => {
    const v = totals[c.id] || 0;
    trHtml += `<td class="text-right">${fmt(v, c.curr)}</td>`;
  });
  $('#full-totals').html(trHtml);
}

$(document).on('change', '#column-checkboxes input', function () {
  const count = $('#column-checkboxes input:checked').length;
  if (count > 8) {
    $(this).prop('checked', false);
    show_alert('Maximum 8 columns allowed for QC view.', 'warning');
    return;
  }
  render_qc_table();
});

/* ============================================================
   POWERFUL PAN / ZOOM / DRAG VIEWER (QC screen)
   Supports: images (jpg/png) and PDF iframes
   - Mouse wheel zoom with cursor-anchored scaling
   - Click-and-drag panning
   - Touch pinch-to-zoom and two-finger pan
   - Zoom buttons, fit-to-width, reset
   - Smooth transforms, no jank
   ============================================================ */

/* ---- PDF.js Canvas Renderer ---- */
function render_pdf_to_canvas(url, containerId, canvasId, viewerId) {
  const $container = $('#' + containerId);
  const $viewerCanvas = $('#' + canvasId);
  const viewer = document.getElementById(viewerId);
  $container.empty().show();
  $container.html('<div style="text-align:center;padding:2rem;color:#94a3b8;"><i class="fa fa-spinner fa-spin fa-2x"></i><br><small>Rendering PDF…</small></div>');

  const PDF_SCALE = 1.5; // render resolution multiplier for sharpness

  pdfjsLib.getDocument(url).promise.then(function (pdf) {
    $container.empty();
    var totalHeight = 0;
    var maxWidth = 0;
    var renderChain = Promise.resolve();

    for (var p = 1; p <= pdf.numPages; p++) {
      (function (pageNum) {
        renderChain = renderChain.then(function () {
          return pdf.getPage(pageNum).then(function (page) {
            var vp = page.getViewport({ scale: PDF_SCALE });
            var c = document.createElement('canvas');
            c.width = vp.width;
            c.height = vp.height;
            c.style.width = (vp.width / PDF_SCALE) + 'px';
            c.style.height = (vp.height / PDF_SCALE) + 'px';
            $container.append(c);
            totalHeight += (vp.height / PDF_SCALE) + 2;
            maxWidth = Math.max(maxWidth, vp.width / PDF_SCALE);
            return page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
          });
        });
      })(p);
    }

    renderChain.then(function () {
      // Size the wrapper canvas div to fit all pages
      $viewerCanvas.css({ width: maxWidth + 'px', height: totalHeight + 'px' });
      // Auto fit-to-width
      var fitZoom = viewer.clientWidth / maxWidth;
      current_zoom = Math.min(fitZoom, 1);
      pan_x = 0; pan_y = 0;
      apply_viewer_transform();
      init_qc_viewer_events();
    });
  }).catch(function (err) {
    $container.html('<div style="text-align:center;padding:2rem;color:#ef4444;"><i class="fa fa-exclamation-triangle fa-2x"></i><br>Failed to load PDF</div>');
  });
}

let pan_x = 0, pan_y = 0;
let is_dragging = false, drag_start_x = 0, drag_start_y = 0;
let _qc_viewer_bound = false;
let _qc_touch_state = null;

function qc_viewer_reset() {
  current_zoom = 1;
  pan_x = 0;
  pan_y = 0;
  _qc_viewer_bound = false;
  _qc_touch_state = null;
  const canvas = document.getElementById('viewer-canvas');
  if (canvas) {
    canvas.style.transform = '';
    canvas.style.display = '';
  }
  update_zoom_display();
}

function apply_viewer_transform() {
  const canvas = document.getElementById('viewer-canvas');
  if (!canvas) return;
  canvas.style.transform = `translate(${pan_x}px, ${pan_y}px) scale(${current_zoom})`;
  update_zoom_display();
}

function update_zoom_display() {
  const el = document.getElementById('zoom-level-display');
  if (el) el.textContent = Math.round(current_zoom * 100) + '%';
}

window.zoom_in = function () {
  const viewer = document.getElementById('document-viewer');
  if (!viewer) return;
  zoom_at_center(0.25);
};

window.zoom_out = function () {
  const viewer = document.getElementById('document-viewer');
  if (!viewer) return;
  zoom_at_center(-0.25);
};

window.reset_zoom = function () {
  current_zoom = 1;
  pan_x = 0;
  pan_y = 0;
  apply_viewer_transform();
};

window.fit_to_width = function () {
  const viewer = document.getElementById('document-viewer');
  const canvas = document.getElementById('viewer-canvas');
  if (!viewer || !canvas) return;
  const viewerW = viewer.clientWidth;
  const canvasW = parseInt(canvas.style.width) || canvas.scrollWidth || viewerW;
  current_zoom = viewerW / canvasW;
  pan_x = 0;
  pan_y = 0;
  apply_viewer_transform();
};

function zoom_at_center(delta) {
  const viewer = document.getElementById('document-viewer');
  if (!viewer) return;
  const rect = viewer.getBoundingClientRect();
  const cx = rect.width / 2;
  const cy = rect.height / 2;
  zoom_at_point(cx, cy, delta);
}

function zoom_at_point(viewX, viewY, delta) {
  const oldZoom = current_zoom;
  const newZoom = Math.min(Math.max(current_zoom + delta, 0.1), 8);
  if (newZoom === oldZoom) return;

  // Anchor: the point under the cursor stays fixed
  const scale = newZoom / oldZoom;
  pan_x = viewX - scale * (viewX - pan_x);
  pan_y = viewY - scale * (viewY - pan_y);
  current_zoom = newZoom;
  apply_viewer_transform();
}

function init_qc_viewer_events() {
  if (_qc_viewer_bound) return;
  _qc_viewer_bound = true;

  const viewer = document.getElementById('document-viewer');
  if (!viewer) return;

  // ---- MOUSE WHEEL ZOOM ----
  viewer.addEventListener('wheel', function (e) {
    e.preventDefault();
    const rect = viewer.getBoundingClientRect();
    const mx = e.clientX - rect.left;
    const my = e.clientY - rect.top;
    const delta = e.deltaY > 0 ? -0.15 : 0.15;
    zoom_at_point(mx, my, delta);
  }, { passive: false });

  // ---- MOUSE DRAG PAN ----
  viewer.addEventListener('mousedown', function (e) {
    if (e.button !== 0) return;
    // If clicking on iframe (PDF), allow PDF interaction at zoom 1
    if (e.target.tagName === 'IFRAME' && current_zoom <= 1.05) return;
    is_dragging = true;
    drag_start_x = e.clientX - pan_x;
    drag_start_y = e.clientY - pan_y;
    viewer.classList.add('is-dragging');
    e.preventDefault();
  });

  document.addEventListener('mousemove', function (e) {
    if (!is_dragging) return;
    pan_x = e.clientX - drag_start_x;
    pan_y = e.clientY - drag_start_y;
    apply_viewer_transform();
  });

  document.addEventListener('mouseup', function () {
    if (!is_dragging) return;
    is_dragging = false;
    viewer.classList.remove('is-dragging');
  });

  // ---- TOUCH SUPPORT (pinch-to-zoom + drag) ----
  viewer.addEventListener('touchstart', function (e) {
    if (e.touches.length === 1) {
      is_dragging = true;
      drag_start_x = e.touches[0].clientX - pan_x;
      drag_start_y = e.touches[0].clientY - pan_y;
      _qc_touch_state = null;
    } else if (e.touches.length === 2) {
      is_dragging = false;
      const dx = e.touches[0].clientX - e.touches[1].clientX;
      const dy = e.touches[0].clientY - e.touches[1].clientY;
      const rect = viewer.getBoundingClientRect();
      _qc_touch_state = {
        dist: Math.sqrt(dx * dx + dy * dy),
        zoom: current_zoom,
        cx: ((e.touches[0].clientX + e.touches[1].clientX) / 2) - rect.left,
        cy: ((e.touches[0].clientY + e.touches[1].clientY) / 2) - rect.top,
        panX: pan_x,
        panY: pan_y
      };
    }
    e.preventDefault();
  }, { passive: false });

  viewer.addEventListener('touchmove', function (e) {
    if (e.touches.length === 1 && is_dragging) {
      pan_x = e.touches[0].clientX - drag_start_x;
      pan_y = e.touches[0].clientY - drag_start_y;
      apply_viewer_transform();
    } else if (e.touches.length === 2 && _qc_touch_state) {
      const dx = e.touches[0].clientX - e.touches[1].clientX;
      const dy = e.touches[0].clientY - e.touches[1].clientY;
      const newDist = Math.sqrt(dx * dx + dy * dy);
      const scale = newDist / _qc_touch_state.dist;
      const newZoom = Math.min(Math.max(_qc_touch_state.zoom * scale, 0.1), 8);
      const zoomRatio = newZoom / _qc_touch_state.zoom;
      pan_x = _qc_touch_state.cx - zoomRatio * (_qc_touch_state.cx - _qc_touch_state.panX);
      pan_y = _qc_touch_state.cy - zoomRatio * (_qc_touch_state.cy - _qc_touch_state.panY);
      current_zoom = newZoom;
      apply_viewer_transform();
    }
    e.preventDefault();
  }, { passive: false });

  viewer.addEventListener('touchend', function () {
    is_dragging = false;
    _qc_touch_state = null;
  });

  // ---- DOUBLE-CLICK TO ZOOM ----
  viewer.addEventListener('dblclick', function (e) {
    const rect = viewer.getBoundingClientRect();
    const mx = e.clientX - rect.left;
    const my = e.clientY - rect.top;
    if (current_zoom > 1.5) {
      // Reset
      current_zoom = 1; pan_x = 0; pan_y = 0;
      apply_viewer_transform();
    } else {
      zoom_at_point(mx, my, 1.0);
    }
  });
}

// Clean up viewer events when modal closes
$(document).on('hidden.bs.modal', '#qc-modal', function () {
  _qc_viewer_bound = false;
  is_dragging = false;
  _qc_touch_state = null;
});

/* ============================================================
   SAVE
   ============================================================ */
function do_save(btn_el) {
  if (!statement_doc) {
    show_alert('No statement document to submit.', 'warning');
    return;
  }
  const $btn = $(btn_el);
  $btn.prop('disabled', true).html('<i class="fa fa-spinner fa-spin"></i> Submitting…');

  call_api('scanify.api.save_extracted_statement', {
    doc_name: statement_doc,
    data: extracted_data
  }).then(r => {
    if (r && r.success) {
      show_alert('Statement submitted successfully!', 'success');
      $btn.html('<i class="fa fa-check"></i> Submitted');
      setTimeout(() => {
        window.location.href = '/portal/stock-statements-list';
      }, 1500);
    } else {
      $btn.prop('disabled', false).html('<i class="fa fa-paper-plane"></i> Submit & Finalise');
      show_alert('Submission failed: ' + ((r && r.message) || 'Unknown error'), 'danger');
    }
  }).catch(e => {
    $btn.prop('disabled', false).html('<i class="fa fa-paper-plane"></i> Submit & Finalise');
    show_alert('Failed to submit: ' + e.message, 'danger');
  });
}

window.save_statement = function (event) {
  do_save(event && event.currentTarget ? event.currentTarget : document.getElementById('btn-save'));
};

window.save_statement_from_qc = function (event) {
  $('#qc-modal').modal('hide');
  do_save(event && event.currentTarget ? event.currentTarget : document.getElementById('btn-save'));
};

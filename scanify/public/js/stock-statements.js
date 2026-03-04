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

  $(document).on('input', '#qc-tbody input, #full-tbody input', function () {
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

    // Step 3: AI extract
    set_step('step-ai', 'active');
    const result = await extract_statement(statement_doc, uploaded_file_url);
    if (!result || !result.success) throw new Error((result && result.message) || 'Extraction failed');
    set_step('step-ai', 'done');

    // Step 4: Enrich
    set_step('step-enrich', 'active');
    const doc_data = await call_api('frappe.client.get', { doctype: 'Stockist Statement', name: statement_doc });

    extracted_data = (doc_data.items || []).map(item => ({
      productcode: item.product_code,
      openingqty: item.opening_qty || 0,
      purchaseqty: item.purchase_qty || 0,
      salesqty: item.sales_qty || 0,
      freeqty: item.free_qty || 0,
      freeqtyscheme: 0,
      returnqty: item.return_qty || 0,
      miscoutqty: item.misc_out_qty || 0,
      closingqty: item.closing_qty || 0,
      closingvalue: item.closing_value || 0
    }));

    await enrich_data_with_master_info();
    set_step('step-enrich', 'done');

    $('#extraction-loader').fadeOut(300, () => {
      display_results();
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
  const cf = row.conversion_factor || 1;
  const pts = row.pts || 0;
  const ptr = row.ptr || 0;

  const c_op = row.openingqty / cf;
  const c_pu = row.purchaseqty / cf;
  const c_sa = row.salesqty / cf;
  const c_fr = row.freeqty / cf;
  const c_frs = row.freeqtyscheme / cf;
  const c_rt = row.returnqty / cf;
  const c_mo = row.miscoutqty / cf;

  row.openingvalue = c_op * pts;
  row.purchasevalue = c_pu * pts;

  const closing_base = c_op + c_pu - c_sa - c_fr - c_frs - c_rt - c_mo;
  row.closingqty = closing_base * cf;
  row.closingvalue = closing_base * pts;

  const deducted = (row.salesqty + row.freeqty - row.freeqtyscheme) / cf;
  row.salesvaluepts = deducted * pts;
  row.salesvalueptr = deducted * ptr;
  row.schemedeductedqty = deducted * cf;
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

function fmt(v) {
  const n = parseFloat(v || 0);
  return isNaN(n) ? '0.00' : n.toFixed(2);
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
  { id: 'productcode', label: 'Code', readonly: true },
  { id: 'productname', label: 'Product Name', readonly: true },
  { id: 'pack', label: 'Pack', readonly: true },
  { id: 'pts', label: 'PTS', readonly: true, curr: true },
  { id: 'conversion_factor', label: 'Conv', readonly: true },
  { id: 'openingqty', label: 'Opening' },
  { id: 'purchaseqty', label: 'Purchase' },
  { id: 'salesqty', label: 'Sales' },
  { id: 'freeqty', label: 'Free' },
  { id: 'freeqtyscheme', label: 'Appr Free' },
  { id: 'returnqty', label: 'Return' },
  { id: 'schemedeductedqty', label: 'Sch Ded', readonly: true },
  { id: 'closingqty', label: 'Closing', readonly: true },
  { id: 'miscoutqty', label: 'Misc Out' },
  { id: 'openingvalue', label: 'Open Val', readonly: true, curr: true },
  { id: 'purchasevalue', label: 'Purch Val', readonly: true, curr: true },
  { id: 'salesvaluepts', label: 'Sales(PTS)', readonly: true, curr: true },
  { id: 'salesvalueptr', label: 'Sales(PTR)', readonly: true, curr: true },
  { id: 'closingvalue', label: 'Cls Val', readonly: true, curr: true }
];

window.open_fullscreen_table = function () {
  render_full_edit_table();
  $('#fullscreen-modal').modal('show');
};

window.open_qc_screen = function () {
  render_qc_table();

  if (!uploaded_file_url) {
    show_alert('Original file not available for QC view', 'warning');
  } else {
    const ext = uploaded_file_url.split('.').pop().toLowerCase();
    if (ext === 'pdf') {
      $('#pdf-viewer').attr('src', uploaded_file_url).show();
      $('#img-viewer').hide();
    } else if (['jpg', 'jpeg', 'png'].includes(ext)) {
      $('#img-viewer').attr('src', uploaded_file_url).show();
      $('#pdf-viewer').hide();
    }
  }

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

  extracted_data.forEach((row, i) => {
    let tr = `<tr data-idx="${i}">`;
    col_configs.forEach(c => {
      if (!active_cols.includes(c.id)) return;
      if (c.readonly) {
        const v = c.curr ? fmt(row[c.id]) : (row[c.id] || '');
        tr += `<td class="text-right" id="cell-${i}-${c.id}">${v}</td>`;
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

  extracted_data.forEach((row, i) => {
    let tr = `<tr data-idx="${i}">`;
    col_configs.forEach(c => {
      if (c.readonly) {
        const v = c.curr ? fmt(row[c.id]) : escape_html(row[c.id] || '');
        tr += `<td class="text-right" id="fcell-${i}-${c.id}">${v}</td>`;
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
  do_row_calc(row);

  col_configs.forEach(c => {
    if (c.readonly) {
      const v = c.curr ? fmt(row[c.id]) : row[c.id];
      $tr.find(`#cell-${idx}-${c.id}`).text(v);
      $(`#fcell-${idx}-${c.id}`).text(v);
    }
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
    trHtml += `<td class="text-right">${c.curr ? fmt(v) : v.toFixed(2)}</td>`;
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
   ZOOM helpers (QC screen)
   ============================================================ */
window.zoom_in = function () { current_zoom = Math.min(current_zoom + 0.2, 3); apply_zoom(); };
window.zoom_out = function () { current_zoom = Math.max(current_zoom - 0.2, 0.4); apply_zoom(); };
window.reset_zoom = function () { current_zoom = 1; apply_zoom(); };

function apply_zoom() {
  $('#document-viewer img, #document-viewer iframe').css('transform', `scale(${current_zoom})`);
}

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

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

function initialize_page() {
  const today = new Date();
  $('#statement-month').val(today.toISOString().slice(0, 7));

  init_stockist_search();
  init_file_upload();

  $('#statement-form').on('submit', async function (e) {
    e.preventDefault();
    await handle_extraction();
  });

  $('#browse-btn').on('click', function () {
    document.getElementById('statement-file').click();
  });

  $('#clear-file-btn').on('click', function () {
    clear_file();
  });
}

/* ---------------- Stockist search ---------------- */

function init_stockist_search() {
  let search_timeout;

  $('#stockist-search').on('input', function () {
    const term = $(this).val().trim();
    clearTimeout(search_timeout);

    $('#stockist-code').val('');
    $('#hq,#team,#region,#zone').val('');
    $('#stockist-info').text('');

    if (term.length < 2) {
      $('#stockist-results').hide().empty();
      return;
    }

    search_timeout = setTimeout(() => {
      search_stockists(term);
    }, 250);
  });

  $(document).on('click', function (e) {
    if (!$(e.target).closest('#stockist-search, #stockist-results').length) {
      $('#stockist-results').hide();
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
      division: $('#division').val()
    }),
    success: function (r) {
      const list = (r.message || []);
      render_stockist_results(list);
    },
    error: function () {
      $('#stockist-results').hide().empty();
      show_alert('Failed to search stockists', 'danger');
    }
  });
}

function render_stockist_results(stockists) {
  if (!stockists.length) {
    $('#stockist-results').hide().empty();
    return;
  }

  const $wrap = $('#stockist-results');
  $wrap.empty();

  stockists.forEach(st => {
    const label = `${st.stockist_code} - ${st.stockist_name}`;
    const meta = [st.hq, st.team, st.region, st.zone].filter(Boolean).join(' | ');

    const $item = $(`
      <div class="autocomplete-item">
        <div class="font-weight-bold">${escape_html(label)}</div>
        <small class="text-muted">${escape_html(meta)}</small>
      </div>
    `);

    $item.on('click', async function () {
      $('#stockist-search').val(label);
      $('#stockist-code').val(st.stockist_code);
      $wrap.hide().empty();

      await populate_stockist_details(st.stockist_code);
    });

    $wrap.append($item);
  });

  $wrap.show();
}

async function populate_stockist_details(stockist_code) {
  try {
    const details = await call_api('scanify.api.get_stockist_details', { stockist_code });
    $('#hq').val(details.hq || '');
    $('#team').val(details.team || '');
    $('#region').val(details.region || '');
    $('#zone').val(details.zone || '');
    $('#stockist-info').text(`Selected: ${details.stockist_name || stockist_code}`);
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

  dropzone.addEventListener('click', function (e) {
    // prevent clicking inside preview buttons from re-opening picker
    if ($(e.target).closest('#clear-file-btn, #browse-btn').length) return;
    fileInput.click();
  });
}

function show_file_preview(file) {
  $('#file-name').text(file.name);
  $('#file-size').text((file.size / 1024).toFixed(2) + ' KB');
  $('.upload-placeholder').hide();
  $('#file-preview').show();
}

function clear_file() {
  const fileInput = document.getElementById('statement-file');
  fileInput.value = '';
  $('#file-preview').hide();
  $('.upload-placeholder').show();
}

/* ---------------- Extraction flow ---------------- */

async function handle_extraction() {
  const stockist_code = $('#stockist-code').val();
  const month = $('#statement-month').val();
  const file = document.getElementById('statement-file').files[0];

  if (!stockist_code) return show_alert('Please select a stockist', 'warning');
  if (!month) return show_alert('Please select statement month', 'warning');
  if (!file) return show_alert('Please upload a statement file', 'warning');

  $('#entry-card').fadeOut();
  $('#extraction-loader').fadeIn();

  try {
    uploaded_file_url = await upload_file(file);
    statement_doc = await create_statement_doc(stockist_code, month, uploaded_file_url);
    const result = await extract_statement(statement_doc, uploaded_file_url);

    if (!result.success) throw new Error(result.message || 'Extraction failed');

    extracted_data = result.data || [];
    $('#extraction-loader').fadeOut();
    display_results();
  } catch (error) {
    console.error(error);
    $('#extraction-loader').fadeOut();
    $('#entry-card').fadeIn();
    show_alert('Extraction failed: ' + (error.message || error), 'danger');
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
        const url = r.message && (r.message.file_url || r.message.file_url);
        if (!url) return reject(new Error('Upload succeeded but file_url missing'));
        resolve(url);
      },
      error: function () {
        reject(new Error('File upload failed'));
      }
    });
  });
}

async function create_statement_doc(stockist_code, month, file_url) {
  // input type=month gives YYYY-MM; store as YYYY-MM-01
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
          extraction_data_status: 'Pending',
          docstatus: 0
        }
      }),
      success: function (r) {
        if (r.message && r.message.name) resolve(r.message.name);
        else reject(new Error('Failed to create statement'));
      },
      error: function () {
        reject(new Error('Document creation failed'));
      }
    });
  });
}

async function extract_statement(doc_name, file_url) {
  // Backend returns {success, data, message}
  return call_api('scanify.api.extract_stockist_statement', {
    doc_name: doc_name,
    file_url: file_url
  });
}

function display_results() {
  const total_products = extracted_data.length;

  let total_opening = 0, total_sales = 0, total_closing = 0;
  extracted_data.forEach(row => {
    total_opening += parseFloat(row.openingqty || 0);
    total_sales += parseFloat(row.salesqty || 0);
    total_closing += parseFloat(row.closingqty || 0);
  });

  $('#total-products').text(total_products);
  $('#total-opening').text(total_opening.toFixed(0));
  $('#total-sales').text(total_sales.toFixed(0));
  $('#total-closing').text(total_closing.toFixed(0));

  $('#extraction-info').text(`${total_products} products extracted successfully`);
  render_preview_table();
  $('#results-section').fadeIn();
}

function render_preview_table() {
  const $tbody = $('#preview-tbody');
  $tbody.empty();

  extracted_data.slice(0, 30).forEach(row => {
    $tbody.append(`
      <tr>
        <td>${escape_html(row.productcode || '')}</td>
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
  // keep your existing alert UI if you already have one
  alert(message);
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
        if (r.message === undefined) return reject(new Error('Empty response'));
        resolve(r.message);
      },
      error: function (xhr) {
        reject(new Error(xhr.responseJSON?.message || 'API error'));
      }
    });
  });
}

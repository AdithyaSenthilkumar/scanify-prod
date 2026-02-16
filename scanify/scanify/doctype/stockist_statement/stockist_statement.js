frappe.ui.form.on('Stockist Statement', {
    refresh: function(frm) {
        // Extract Data button
        if (frm.doc.uploaded_file && frm.doc.extracted_data_status === 'Pending' && !frm.doc.__islocal) {
            frm.add_custom_button('Extract Data with AI', function() {
                extract_statement_data(frm);
            });
        }
        
        // Fetch Previous Month button
        if (!frm.doc.__islocal && frm.doc.docstatus === 0) {
            frm.add_custom_button('Fetch Previous Month Closing', function() {
                fetch_previous_closing(frm);
            });
        }
        
        // View uploaded file
        if (frm.doc.uploaded_file) {
            frm.add_custom_button('View Uploaded File', function() {
                window.open(frm.doc.uploaded_file, '_blank');
            });
        }
    },
    
    statement_month: function(frm) {
        // Month is automatically set - no need for from_date and to_date
        if (frm.doc.statement_month) {
            // Auto-populate opening balances from previous month
            if (frm.doc.stockist_code && !frm.doc.__islocal) {
                fetch_previous_closing(frm);
            }
        }
    }
});
frappe.ui.form.on('Stockist Statement', {
    refresh: function(frm) {
        if (!frm.doc.__islocal) {
            frm.add_custom_button('Open in Excel View', function() {
                frappe.set_route('List', 'Stockist Statement Item', 'Report', {
                    'parent': frm.doc.name
                });
            }).addClass('btn-primary');
        }
    }
});
frappe.ui.form.on('Stockist Statement', {
    refresh: function(frm) {
        // Add "Toggle Detailed View" button
        if (!frm.doc.__islocal && frm.doc.items && frm.doc.items.length > 0) {
            frm.add_custom_button('View Fullscreen Table', function() {
                show_detailed_view_dialog(frm);
            }).addClass('btn-primary');
        }
    }
});

function show_detailed_view_dialog(frm) {
    let dialog = new frappe.ui.Dialog({
        title: `Detailed Statement: ${frm.doc.stockist_name || frm.doc.stockist_code}`,
        size: 'extra-large',
        fields: [{ fieldname: 'table_html', fieldtype: 'HTML' }],
        primary_action_label: 'Close',
        primary_action: () => dialog.hide()
    });

    let html = build_detailed_table(frm);
    dialog.fields_dict.table_html.$wrapper.html(html);

    const $wrapper = dialog.$wrapper;

    // TRULY FULL WIDTH - KILL ALL MARGINS
    $wrapper.find('.modal-dialog').css({
        'width': '100vw',
        'max-width': '100vw',
        'margin': '0',
        'height': '100vh'
    });

    $wrapper.find('.modal-content').css({
        'height': '100vh',
        'border': 'none',
        'border-radius': '0'
    });

    // Remove the "gutter" on the right (Frappe's default padding)
    $wrapper.find('.modal-body').css({
        'padding': '0px',
        'margin': '0px',
        'overflow': 'hidden' 
    });

    $wrapper.find('.container').css({
        'width': '100%',
        'max-width': '100%',
        'padding': '0'
    });

    // Target the specific Frappe row/column classes that add padding
    $wrapper.find('.form-layout, .form-page, .form-section, .section-body').css({
        'padding': '0 !important',
        'margin': '0 !important'
    });

    dialog.show();
}

function build_detailed_table(frm) {
    let items = frm.doc.items || [];
    
    // Adjusted widths to be tighter so more fits on screen
    const columns = [
        { fieldname: 'product_code', label: 'Product Code', width: '90px' },
        { fieldname: 'product_name', label: 'Product Name', width: '180px' },
        { fieldname: 'pack', label: 'Pack', width: '70px' },
        { fieldname: 'pts', label: 'PTS', width: '70px', is_currency: true },
        { fieldname: 'conversion_factor', label: 'Conv', width: '60px' },
        { fieldname: 'opening_qty', label: 'Opening Qty', width: '90px' },
        { fieldname: 'purchase_qty', label: 'Purchase Qty', width: '90px' },
        { fieldname: 'sales_qty', label: 'Sales Qty', width: '90px' },
        { fieldname: 'free_qty', label: 'Free Qty', width: '80px' },
        { fieldname: 'free_qty_scheme', label: 'Appr Free (Scheme)', width: '110px' },
        { fieldname: 'sales_return_qty', label: 'Return Qty', width: '85px' },
        { fieldname: 'scheme_deducted_qty_calc', label: 'Scheme Ded', width: '90px' },
        { fieldname: 'closing_qty', label: 'Closing Qty', width: '90px' },
        { fieldname: 'misc_out_trans_out', label: 'Misc Out', width: '80px' },
        { fieldname: 'opening_value', label: 'Opening Val', width: '100px', is_currency: true },
        { fieldname: 'purchase_value', label: 'Purchase Val', width: '100px', is_currency: true },
        { fieldname: 'sales_value_pts', label: 'Sales (PTS)', width: '100px', is_currency: true },
        { fieldname: 'sales_value_ptr', label: 'Sales (PTR)', width: '100px', is_currency: true },
        { fieldname: 'closing_value', label: 'Closing Val', width: '100px', is_currency: true }
    ];

    let html = `
        <style>
            .table-containe r {
                width: 100vw;
                height: calc(100vh - 120px);
                overflow: auto;
                margin: 0;
                padding: 0;
                display: block;
            }
            .detailed-view-table {
                width: 100%;
                table-layout: fixed; /* Forces columns to stay within assigned widths */
                border-collapse: collapse;
                font-size: 11px;
                background: white;
            }
            .detailed-view-table thead th {
                position: sticky;
                top: 0;
                background: #2c3e50;
                color: white;
                padding: 8px 4px;
                z-index: 20;
                border: 1px solid #1a252f;
                text-align: left;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
            }
            .detailed-view-table td {
                padding: 6px 4px;
                border: 1px solid #d1d8dd;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
            }
            .summary-row {
                position: sticky;
                bottom: 0;
                background: #d4edda !important;
                font-weight: bold;
                z-index: 10;
            }
            .text-right { text-align: right; }
        </style>
        
        <div class="table-container">
            <table class="detailed-view-table">
                <thead>
                    <tr>
    `;

    columns.forEach(col => {
        html += `<th style="width: ${col.width};" title="${col.label}">${col.label}</th>`;
    });
    html += `</tr></thead><tbody>`;

    items.forEach(item => {
        html += '<tr>';
        columns.forEach(col => {
            let val = item[col.fieldname] || 0;
            let display = col.is_currency ? format_currency(val) : (typeof val === 'number' ? val.toFixed(2) : val);
            let align = (col.is_currency || typeof val === 'number') ? 'text-right' : '';
            html += `<td class="${align}" title="${display}">${display}</td>`;
        });
        html += '</tr>';
    });

    // Totals logic
    let t_op = 0, t_pu = 0, t_sa = 0, t_cl = 0, t_ov = 0, t_cv = 0;
    items.forEach(i => {
        t_op += flt(i.opening_qty); t_pu += flt(i.purchase_qty); t_sa += flt(i.sales_qty);
        t_cl += flt(i.closing_qty); t_ov += flt(i.opening_value); t_cv += flt(i.closing_value);
    });

    html += `
        <tr class="summary-row">
            <td colspan="4" class="text-right">TOTALS</td>
            <td class="text-right">${t_op.toFixed(2)}</td>
            <td class="text-right">${t_pu.toFixed(2)}</td>
            <td class="text-right">${t_sa.toFixed(2)}</td>
            <td colspan="3"></td>
            <td class="text-right">${t_cl.toFixed(2)}</td>
            <td class="text-right">${format_currency(t_ov)}</td>
            <td></td>
            <td class="text-right">${format_currency(t_cv)}</td>
        </tr>
    </tbody></table></div>`;

    return html;
}

function format_currency(value) {
    if (!value) return '0.00';
    return parseFloat(value).toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2});
}

function set_compact_view(frm) {
    // Show only essential columns
    const compact_fields = [
        'product_code', 'product_name', 'pack',
        'opening_qty', 'purchase_qty', 'sales_qty',
        'closing_qty', 'closing_value'
    ];
    
    if (frm.fields_dict.items && frm.fields_dict.items.grid) {
        // Hide all first
        const all_fields = [
            'opening_qty', 'purchase_qty', 'sales_qty', 'free_qty',
            'free_qty_scheme', 'return_qty', 'misc_out_qty', 'closing_qty',
            'pts', 'closing_value', 'conversion_factor'
        ];
        
        all_fields.forEach(field => {
            frm.fields_dict.items.grid.update_docfield_property(
                field, 'in_list_view', compact_fields.includes(field) ? 1 : 0
            );
        });
        
        frm.fields_dict.items.grid.refresh();
    }
}

function set_detailed_view(frm) {
    // Show all columns
    const detailed_fields = [
        'product_code', 'product_name', 'pack',
        'opening_qty', 'purchase_qty', 'sales_qty', 'free_qty',
        'free_qty_scheme', 'return_qty', 'misc_out_qty',
        'closing_qty', 'closing_value'
    ];
    
    if (frm.fields_dict.items && frm.fields_dict.items.grid) {
        detailed_fields.forEach(field => {
            frm.fields_dict.items.grid.update_docfield_property(
                field, 'in_list_view', 1
            );
        });
        
        frm.fields_dict.items.grid.refresh();
    }
}

function toggle_detailed_view(frm) {
    if (!frm._detailed_view_active) {
        set_detailed_view(frm);
        frm._detailed_view_active = true;
        frappe.show_alert({message: 'Detailed View Enabled', indicator: 'blue'});
    } else {
        set_compact_view(frm);
        frm._detailed_view_active = false;
        frappe.show_alert({message: 'Compact View Enabled', indicator: 'green'});
    }
}


frappe.ui.form.on('Stockist Statement Item', {
    product_code: function(frm, cdt, cdn) {
        let row = locals[cdt][cdn];
        if (row.product_code) {
            frappe.db.get_value('Product Master', row.product_code, ['product_name', 'pack', 'pts'], (r) => {
                frappe.model.set_value(cdt, cdn, 'product_name', r.product_name);
                frappe.model.set_value(cdt, cdn, 'pack', r.pack);
                frappe.model.set_value(cdt, cdn, 'pts', r.pts);
                
                // Calculate conversion factor from pack
                let conversion_factor = get_conversion_factor(r.pack);
                frappe.model.set_value(cdt, cdn, 'conversion_factor', conversion_factor);
                
                calculate_item_closing(frm, cdt, cdn);
            });
        }
    },
    
    opening_qty: function(frm, cdt, cdn) { calculate_item_closing(frm, cdt, cdn); },
    purchase_qty: function(frm, cdt, cdn) { calculate_item_closing(frm, cdt, cdn); },
    sales_qty: function(frm, cdt, cdn) { calculate_item_closing(frm, cdt, cdn); },
    free_qty: function(frm, cdt, cdn) { calculate_item_closing(frm, cdt, cdn); },
    free_qty_scheme: function(frm, cdt, cdn) { calculate_item_closing(frm, cdt, cdn); },
    return_qty: function(frm, cdt, cdn) { calculate_item_closing(frm, cdt, cdn); },
    misc_out_qty: function(frm, cdt, cdn) { calculate_item_closing(frm, cdt, cdn); }
});

function calculate_item_closing(frm, cdt, cdn) {
    let row = locals[cdt][cdn];
    
    let conversion_factor = row.conversion_factor || 1;
    
    let converted_opening = flt(row.opening_qty) / conversion_factor;
    let converted_purchase = flt(row.purchase_qty) / conversion_factor;
    let converted_sales = flt(row.sales_qty) / conversion_factor;
    let converted_free = flt(row.free_qty) / conversion_factor;
    let converted_free_scheme = flt(row.free_qty_scheme) / conversion_factor;
    let converted_return = flt(row.return_qty) / conversion_factor;
    let converted_misc_out = flt(row.misc_out_qty) / conversion_factor;
    
    let pts = flt(row.pts || 0);
    
    frappe.model.set_value(cdt, cdn, 'opening_value', converted_opening * pts);
    frappe.model.set_value(cdt, cdn, 'purchase_value', converted_purchase * pts);

    let closing_qty_base = converted_opening + converted_purchase 
                         - converted_sales - converted_free - converted_free_scheme
                         - converted_return - converted_misc_out;
    
    // Convert back to display units
    frappe.model.set_value(cdt, cdn, 'closing_qty', closing_qty_base * conversion_factor);
    
    // Closing Value = Closing Qty (base) * PTS
    if (!row.closing_value || row.closing_value == 0) {
        frappe.model.set_value(cdt, cdn, 'closing_value', closing_qty_base * pts);
    }
    // Recalculate totals (only for fields JS can calculate)
    calculate_totals(frm);
}

function calculate_totals(frm) {
    let total_opening = 0;
    let total_purchase = 0;
    let total_free = 0;
    let total_closing = 0;
    let total_sales_pts = 0;
    let total_sales_ptr = 0;
    
    if (frm.doc.items) {
        frm.doc.items.forEach(item => {
            let conversion_factor = flt(item.conversion_factor) || 1;
            
            // Opening and Purchase: divide by conversion_factor first
            total_opening += (flt(item.opening_qty) / conversion_factor) * flt(item.pts);
            total_purchase += (flt(item.purchase_qty) / conversion_factor) * flt(item.pts);
            total_free += (flt(item.free_qty) / conversion_factor) * flt(item.pts);
            
            // Closing value is already calculated correctly
            total_closing += flt(item.closing_value);
            
            // Sales values - use what Python calculated (already in the row)
            total_sales_pts += flt(item.sales_value_pts);
            total_sales_ptr += flt(item.sales_value_ptr);
        });
    }
    
    frm.set_value('total_opening_value', total_opening);
    frm.set_value('total_purchase_value', total_purchase);
    frm.set_value('total_free_value', total_free);
    frm.set_value('total_closing_value', total_closing);
    frm.set_value('total_sales_value_pts', total_sales_pts);
    frm.set_value('total_sales_value_ptr', total_sales_ptr);
}

function extract_statement_data(frm) {
    frappe.confirm(
        'Extract data from uploaded file using AI?',
        function() {
            frm.set_value('extracted_data_status', 'In Progress');
            frm.save().then(() => {
                frappe.call({
                    method: 'scanify.api.extract_stockist_statement',
                    args: {
                        doc_name: frm.doc.name,
                        file_url: frm.doc.uploaded_file
                    },
                    freeze: true,
                    freeze_message: 'Extracting data with AI...',
                    callback: function(r) {
                        // Handle successful API call
                        if (r.message && r.message.success) {
                            frappe.show_alert({
                                message: 'Data extracted successfully!',
                                indicator: 'green'
                            });
                            frm.reload_doc();
                        } else {
                            // API call succeeded but extraction failed
                            frappe.msgprint({
                                title: 'Extraction Failed',
                                message: r.message ? r.message.message : 'Unknown error occurred',
                                indicator: 'red'
                            });
                            frm.reload_doc();
                        }
                    },
                    error: function(r) {
                        // Handle API call errors (network, permissions, etc.)
                        frappe.msgprint({
                            title: 'Error',
                            message: 'Failed to extract data. Check extraction notes for details.',
                            indicator: 'red'
                        });
                        frm.reload_doc();
                    }
                });
            });
        }
    );
}


function fetch_previous_closing(frm) {
    if (!frm.doc.stockist_code || !frm.doc.statement_month) {
        frappe.msgprint('Please select stockist and month first');
        return;
    }
    
    frappe.call({
        method: 'scanify.api.fetch_previous_month_closing',
        args: {
            stockist_code: frm.doc.stockist_code,
            current_month: frm.doc.statement_month
        },
        callback: function(r) {
            if (r.message && r.message.length > 0) {
                frm.clear_table('items');
                r.message.forEach(item => {
                    let row = frm.add_child('items');
                    row.product_code = item.product_code;
                    row.product_name = item.product_name;
                    row.pack = item.pack;
                    row.opening_qty = item.closing_qty;
                    row.pts = item.pts;
                });
                frm.refresh_field('items');
                calculate_totals(frm);
                frappe.msgprint('Previous month closing fetched successfully');
            } else {
                frappe.msgprint('No previous month data found', 'Info');
            }
        }
    });
}

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
        // Add toggle button
        if (!frm.doc.__islocal) {
            frm.add_custom_button(__('Toggle Detailed View'), function() {
                toggle_detailed_view(frm);
            });
        }
        
        // Start with compact view
        set_compact_view(frm);
    }
});

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
    frappe.model.set_value(cdt, cdn, 'closing_value', closing_qty_base * pts);
    
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

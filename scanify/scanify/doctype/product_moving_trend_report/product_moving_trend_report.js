frappe.ui.form.on('Product Moving Trend Report', {
    refresh: function(frm) {
        if (!frm.is_new()) {
            frm.add_custom_button(__('Generate Report'), function() {
                frappe.call({
                    method: 'scanify.scanify.doctype.product_moving_trend_report.product_moving_trend_report.generate_report',
                    args: {
                        doc_name: frm.doc.name
                    },
                    callback: function(r) {
                        if (r.message && r.message.success) {
                            frappe.msgprint({
                                title: __('Success'),
                                message: __('Report generated successfully'),
                                indicator: 'green'
                            });
                            frm.reload_doc();
                        }
                    }
                });
            }).addClass('btn-primary');
            
            frm.add_custom_button(__('Export to Excel'), function() {
                frappe.call({
                    method: 'scanify.scanify.doctype.product_moving_trend_report.product_moving_trend_report.export_to_excel',
                    args: {
                        doc_name: frm.doc.name
                    },
                    callback: function(r) {
                        if (r.message && r.message.success) {
                            frappe.msgprint({
                                title: __('Export Successful'),
                                message: __('Excel file exported successfully'),
                                indicator: 'green'
                            });
                            window.open(r.message.file_url);
                        }
                    }
                });
            }).addClass('btn-success');
            
            frm.add_custom_button(__('Export to PDF'), function() {
                frappe.call({
                    method: 'scanify.scanify.doctype.product_moving_trend_report.product_moving_trend_report.export_to_pdf',
                    args: {
                        doc_name: frm.doc.name
                    },
                    callback: function(r) {
                        if (r.message && r.message.success) {
                            frappe.msgprint({
                                title: __('Export Successful'),
                                message: __('PDF exported successfully'),
                                indicator: 'green'
                            });
                            window.open(r.message.file_url);
                        }
                    }
                });
            }).addClass('btn-info');
        }
    },
    
    product_category: function(frm) {
        // If category selected, clear specific product
        if (frm.doc.product_category && frm.doc.product_category != 'All Products') {
            frm.set_value('product_code', null);
        }
    },
    
    product_code: function(frm) {
        // If specific product selected, clear category
        if (frm.doc.product_code) {
            frm.set_value('product_category', null);
        }
    }
});

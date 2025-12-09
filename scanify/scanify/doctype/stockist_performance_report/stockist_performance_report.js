// Copyright (c) 2025, Stedman Pharmaceuticals and contributors
// For license information, please see license.txt

// frappe.ui.form.on("Stockist Performance Report", {
// 	refresh(frm) {

// 	},
// });
frappe.ui.form.on('Stockist Performance Report', {
    refresh: function(frm) {
        if (!frm.is_new()) {
            frm.add_custom_button(__('Generate Report'), function() {
                frappe.call({
                    method: 'scanify.scanify.doctype.stockist_performance_report.stockist_performance_report.generate_report',
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
                    method: 'scanify.scanify.doctype.stockist_performance_report.stockist_performance_report.export_to_excel',
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
                    method: 'scanify.scanify.doctype.stockist_performance_report.stockist_performance_report.export_to_pdf',
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
    
    onload: function(frm) {
        if (frm.is_new()) {
            // Set default dates - current month
            let today = frappe.datetime.get_today();
            frm.set_value('from_date', frappe.datetime.month_start(today));
            frm.set_value('to_date', frappe.datetime.month_end(today));
        }
    }
});

frappe.ui.form.on('Secondary Sales Report', {
    refresh: function(frm) {
        // Add custom buttons
        if (!frm.is_new()) {
            frm.add_custom_button(__('Generate Report'), function() {
                frappe.call({
                    method: 'scanify.scanify.doctype.secondary_sales_report.secondary_sales_report.generate_report',
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
            
            frm.add_custom_button(__('Export to PDF'), function() {
                frappe.call({
                    method: 'scanify.scanify.doctype.secondary_sales_report.secondary_sales_report.export_to_pdf',
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
            
            frm.add_custom_button(__('Export to Excel'), function() {
                frappe.call({
                    method: 'scanify.scanify.doctype.secondary_sales_report.secondary_sales_report.export_to_excel',
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
        }
    },
    
    report_type: function(frm) {
        // Reset filters based on report type
        if (frm.doc.report_type === 'HQ Wise') {
            frm.set_value('region', null);
            frm.set_value('team', null);
            frm.set_value('stockist', null);
        } else if (frm.doc.report_type === 'Team Wise') {
            frm.set_value('hq', null);
            frm.set_value('stockist', null);
        } else if (frm.doc.report_type === 'Region Wise') {
            frm.set_value('hq', null);
            frm.set_value('team', null);
            frm.set_value('stockist', null);
        } else if (frm.doc.report_type === 'Stockist Wise') {
            frm.set_value('region', null);
            frm.set_value('team', null);
            frm.set_value('hq', null);
        }
    },
    
    include_scheme_deduction: function(frm) {
        frm.refresh_field('scheme_deduction_value');
    }
});

frappe.ui.form.on('Scheme Not Reflected Report', {
    refresh: function(frm) {
        if (!frm.is_new()) {
            frm.add_custom_button(__('Generate Report'), function() {
                frappe.call({
                    method: 'scanify.scanify.doctype.scheme_not_reflected_report.scheme_not_reflected_report.generate_report',
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
                    method: 'scanify.scanify.doctype.scheme_not_reflected_report.scheme_not_reflected_report.export_to_excel',
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
                    method: 'scanify.scanify.doctype.scheme_not_reflected_report.scheme_not_reflected_report.export_to_pdf',
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
        // Set default dates
        if (frm.is_new()) {
            let today = frappe.datetime.get_today();
            
            // Statement period - current month
            frm.set_value('from_date', frappe.datetime.month_start(today));
            frm.set_value('to_date', frappe.datetime.month_end(today));
            
            // Scheme approval period - last 3 months
            let three_months_ago = frappe.datetime.add_months(today, -3);
            frm.set_value('scheme_approval_from_date', three_months_ago);
            frm.set_value('scheme_approval_to_date', today);
        }
    }
});

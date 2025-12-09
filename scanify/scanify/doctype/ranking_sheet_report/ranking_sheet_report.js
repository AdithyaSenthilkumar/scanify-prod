frappe.ui.form.on('Ranking Sheet Report', {
    refresh: function(frm) {
        if (!frm.is_new()) {
            frm.add_custom_button(__('Generate Report'), function() {
                frappe.call({
                    method: 'scanify.scanify.doctype.ranking_sheet_report.ranking_sheet_report.generate_report',
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
                    method: 'scanify.scanify.doctype.ranking_sheet_report.ranking_sheet_report.export_to_excel',
                    args: {
                        doc_name: frm.doc.name
                    },
                    callback: function(r) {
                        if (r.message && r.message.success) {
                            frappe.msgprint({
                                title: __('Export Successful'),
                                message: __('Excel file exported with rankings'),
                                indicator: 'green'
                            });
                            window.open(r.message.file_url);
                        }
                    }
                });
            }).addClass('btn-success');
            
            frm.add_custom_button(__('Export to PDF'), function() {
                frappe.call({
                    method: 'scanify.scanify.doctype.ranking_sheet_report.ranking_sheet_report.export_to_pdf',
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
    
    period_type: function(frm) {
        if (frm.doc.period_type === 'Monthly') {
            let today = frappe.datetime.get_today();
            frm.set_value('from_date', frappe.datetime.month_start(today));
            frm.set_value('to_date', frappe.datetime.month_end(today));
        } else if (frm.doc.period_type === 'Yearly') {
            let today = frappe.datetime.get_today();
            frm.set_value('from_date', frappe.datetime.year_start(today));
            frm.set_value('to_date', frappe.datetime.year_end(today));
        }
    },
    
    quarter: function(frm) {
        if (frm.doc.quarter && frm.doc.period_type === 'Quarterly') {
            frm.save();
        }
    },
    
    ranking_type: function(frm) {
        // Show/hide product code based on ranking type
        frm.set_df_property('product_code', 'hidden', 
            frm.doc.ranking_type !== 'Product-wise');
    }
});

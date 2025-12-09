frappe.ui.form.on('Incentive Calculation Report', {
    refresh: function(frm) {
        if (!frm.is_new()) {
            frm.add_custom_button(__('Generate Report'), function() {
                frappe.call({
                    method: 'scanify.scanify.doctype.incentive_calculation_report.incentive_calculation_report.generate_report',
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
                    method: 'scanify.scanify.doctype.incentive_calculation_report.incentive_calculation_report.export_to_excel',
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
                    method: 'scanify.scanify.doctype.incentive_calculation_report.incentive_calculation_report.export_to_pdf',
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
        // Show/hide date fields based on period type
        frm.set_df_property('quarter', 'reqd', frm.doc.period_type === 'Quarterly');
        frm.set_df_property('from_date', 'reqd', frm.doc.period_type === 'Custom Date Range');
        frm.set_df_property('to_date', 'reqd', frm.doc.period_type === 'Custom Date Range');
        
        if (frm.doc.period_type === 'Monthly') {
            // Set current month
            let today = frappe.datetime.get_today();
            frm.set_value('from_date', frappe.datetime.month_start(today));
            frm.set_value('to_date', frappe.datetime.month_end(today));
        }
    },
    
    quarter: function(frm) {
        if (frm.doc.quarter && frm.doc.period_type === 'Quarterly') {
            // Dates will be set automatically by Python validate
            frm.save();
        }
    },
    
    calculation_type: function(frm) {
        // Show appropriate rate fields
        let show_unit = frm.doc.calculation_type === 'Product-wise' || frm.doc.calculation_type === 'Both';
        let show_rupee = frm.doc.calculation_type === 'Rupee-wise' || frm.doc.calculation_type === 'Both';
        
        frm.set_df_property('incentive_rate_per_unit', 'hidden', !show_unit);
        frm.set_df_property('incentive_rate_per_rupee', 'hidden', !show_rupee);
    }
});

// bulk_statement_upload_list.js

frappe.listview_settings['Bulk Statement Upload'] = {
    add_fields: ['status', 'statement_month'],
    get_indicator: function (doc) {
        if (doc.status === 'Completed') {
            return [__('Completed'), 'green', 'status,=,Completed'];
        } else if (doc.status === 'In Progress') {
            return [__('In Progress'), 'orange', 'status,=,In Progress'];
        } else if (doc.status === 'Failed') {
            return [__('Failed'), 'red', 'status,=,Failed'];
        } else {
            return [__('Pending'), 'gray', 'status,=,Pending'];
        }
    },

    onload: function (listview) {
        // Button to create a new bulk job quickly
        listview.page.add_inner_button(__('New Bulk Job'), function () {
            frappe.new_doc('Bulk Statement Upload');
        });

        // Button to re-run extraction for selected failed jobs
        listview.page.add_inner_button(__('Re-run Failed Jobs'), function () {
            let selected = listview.get_checked_items();

            if (!selected.length) {
                frappe.msgprint(__('Please select at least one failed job.'));
                return;
            }

            // Filter only failed jobs
            selected = selected.filter(d => d.status === 'Failed');

            if (!selected.length) {
                frappe.msgprint(__('No failed jobs selected.'));
                return;
            }

            frappe.confirm(
                __('Re-run extraction for {0} failed job(s)?', [selected.length]),
                () => {
                    selected.forEach(doc => {
                        frappe.call({
                            method: 'scanify.api.bulk_extract_statements',
                            args: {
                                month: doc.statement_month,
                                zip_file_url: doc.zip_file
                            },
                            callback: function (r) {
                                // No-op; results are stored in the doc itself
                            }
                        });
                    });
                    frappe.show_alert({
                        message: __('Re-run started for selected jobs'),
                        indicator: 'orange'
                    });
                    setTimeout(() => listview.refresh(), 2000);
                }
            );
        });
    }
};

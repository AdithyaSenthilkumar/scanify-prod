frappe.ui.form.on('Bulk Statement Upload', {
    refresh: function(frm) {
        // Show extract button only for new/pending docs
        if (frm.doc.docstatus === 0 && frm.doc.status === 'Pending' && frm.doc.zipfile && !frm.doc.__islocal) {
            frm.add_custom_button(__('Start Extraction'), function() {
                frappe.confirm(
                    'This will start background extraction. You can close this form and check status later. Continue?',
                    function() {
                        frappe.call({
                            method: 'scanify.api.bulk_extract_statements_async',
                            args: {
                                docname: frm.doc.name
                            },
                            freeze: true,
                            freeze_message: __('Queuing extraction job...'),
                            callback: function(r) {
                                if (r.message && r.message.success) {
                                    frappe.show_alert({
                                        message: __('Extraction job queued successfully!'),
                                        indicator: 'green'
                                    });
                                    frm.reload_doc();
                                }
                            }
                        });
                    }
                );
            }).addClass('btn-primary');
        }
        
        // Refresh button for in-progress jobs
        if (frm.doc.status === 'In Progress' || frm.doc.status === 'Queued') {
            frm.add_custom_button(__('Refresh Status'), function() {
                frm.reload_doc();
            });
            
            // Auto-refresh every 5 seconds
            if (!frm.doc.__auto_refresh_interval) {
                frm.doc.__auto_refresh_interval = setInterval(function() {
                    frm.reload_doc();
                }, 5000);
            }
        } else {
            // Clear auto-refresh if status changed
            if (frm.doc.__auto_refresh_interval) {
                clearInterval(frm.doc.__auto_refresh_interval);
                frm.doc.__auto_refresh_interval = null;
            }
        }
        
        // View results button
        if (frm.doc.extraction_log && (frm.doc.status === 'Completed' || frm.doc.status === 'Partially Completed')) {
            frm.add_custom_button(__('View Results'), function() {
                let log = JSON.parse(frm.doc.extraction_log);
                show_extraction_results(log);
            });
        }
    }
});

function show_extraction_results(results) {
    let html = '<table class="table table-bordered"><thead><tr>' +
        '<th>File</th><th>Status</th><th>Stockist</th><th>Items</th><th>Statement</th></tr></thead><tbody>';
    
    results.forEach(r => {
        let status_indicator = r.status === 'Success' ? 'green' : (r.status === 'Skipped' ? 'orange' : 'red');
        html += `<tr>
            <td>${r.file}</td>
            <td><span class="indicator ${status_indicator}">${r.status}</span></td>
            <td>${r.stockist || 'N/A'}</td>
            <td>${r.items_extracted || 0}</td>
            <td>${r.statement ? `<a href="/app/stockist-statement/${r.statement}">${r.statement}</a>` : (r.message || '-')}</td>
        </tr>`;
    });
    
    html += '</tbody></table>';
    
    frappe.msgprint({
        title: __('Extraction Results'),
        message: html,
        indicator: 'blue',
        wide: true
    });
}

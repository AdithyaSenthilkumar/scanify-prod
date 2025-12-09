frappe.ui.form.on('Scheme Request', {
    refresh: function(frm) {
        // Set status indicator
        if (frm.doc.approval_status) {
            let color = {
                'Approved': 'green',
                'Rejected': 'red',
                'Rerouted': 'orange',
                'Pending': 'blue'
            }[frm.doc.approval_status] || 'gray';
            
            frm.dashboard.set_headline_alert(
                `Status: ${frm.doc.approval_status}`, 
                color
            );
        }
        
        // Set query for stockist based on HQ
        frm.set_query('stockist_code', function() {
            if (!frm.doc.hq) {
                frappe.msgprint(__('Please select HQ first'));
                return { filters: { 'name': '' } };
            }
            return {
                filters: {
                    'hq': frm.doc.hq,
                    'status': 'Active'
                }
            };
        });
        if (frm.doc.approval_status === 'Approved' && !frm.is_new()) {
            frm.add_custom_button(__('Repeat Request'), function() {
                repeat_scheme_request(frm);
            }).css({
                'background-color': '#17a2b8', 
                'color': 'white', 
                'font-weight': 'bold'
            });
        }
        
        // Manager action buttons
        if (frm.doc.docstatus === 0 && 
            frm.doc.approval_status === 'Pending' && 
            !frm.is_new() &&
            (frappe.user.has_role('Sales Manager') || frappe.user.has_role('System Manager'))) {
            
            frm.add_custom_button(__('Approve'), function() {
                approve_scheme(frm);
            }, __('Actions')).css({'background-color': '#28a745', 'color': 'white'});
            
            frm.add_custom_button(__('Reject'), function() {
                reject_scheme(frm);
            }, __('Actions')).css({'background-color': '#dc3545', 'color': 'white'});
            
            frm.add_custom_button(__('Reroute'), function() {
                reroute_scheme(frm);
            }, __('Actions')).css({'background-color': '#fd7e14', 'color': 'white'});
        }
        
        // View attachments button
        if (frm.doc.proof_attachment_1 || frm.doc.proof_attachment_2 || 
            frm.doc.proof_attachment_3 || frm.doc.proof_attachment_4) {
            frm.add_custom_button(__('View Documents'), function() {
                show_all_attachments(frm);
            });
        }
        
        // Add custom CSS for history buttons in grid
        add_grid_history_buttons(frm);
    },
    
    onload: function(frm) {
        // Set query for stockist on form load as well
        frm.set_query('stockist_code', function() {
            if (!frm.doc.hq) {
                return { filters: { 'name': '' } };
            }
            return {
                filters: {
                    'hq': frm.doc.hq,
                    'status': 'Active'
                }
            };
        });
    },
    
    search_doctor: function(frm) {
        if (frm.doc.search_doctor && frm.doc.search_doctor.length >= 2) {
            show_doctor_search_dialog(frm);
        }
    },
    
    doctor_code: function(frm) {
        if (frm.doc.doctor_code) {
            frappe.db.get_value('Doctor Master', frm.doc.doctor_code, 
                ['doctor_name', 'place', 'city_pool', 'team', 'region', 'specialization', 'hospital_clinic'], 
                (r) => {
                    if (r) {
                        frm.set_value('doctor_name', r.doctor_name);
                        frm.set_value('doctor_place', r.place);
                        frm.set_value('city_pool', r.city_pool);
                        frm.set_value('team', r.team);
                        frm.set_value('region', r.region);
                        frm.set_value('specialization', r.specialization);
                        frm.set_value('hospital_clinic', r.hospital_clinic);
                    }
            });
                        check_doctor_monthly_limit(frm);

        }
    },
        application_date: function(frm) {
        if (frm.doc.doctor_code && frm.doc.application_date) {
            check_doctor_monthly_limit(frm);
        }
    },
    
    hq: function(frm) {
        // Clear stockist when HQ changes
        if (frm.doc.stockist_code) {
            frm.set_value('stockist_code', '');
            frm.set_value('stockist_name', '');
        }
        
        if (frm.doc.hq) {
            // Refresh the stockist field query
            frm.fields_dict['stockist_code'].get_query = function() {
                return {
                    filters: {
                        'hq': frm.doc.hq,
                        'status': 'Active'
                    }
                };
            };
            
            // Show count of available stockists
            frappe.call({
                method: 'frappe.client.get_count',
                args: {
                    doctype: 'Stockist Master',
                    filters: {
                        'hq': frm.doc.hq,
                        'status': 'Active'
                    }
                },
                callback: function(r) {
                    if (r.message) {
                        frappe.show_alert({
                            message: __('{0} stockists available in {1}', [r.message, frm.doc.hq]),
                            indicator: 'blue'
                        }, 5);
                    }
                }
            });
        }
    },
    
    stockist_code: function(frm) {
        if (frm.doc.stockist_code) {
            // Verify stockist belongs to selected HQ
            frappe.db.get_value('Stockist Master', frm.doc.stockist_code, 'hq', (r) => {
                if (r && r.hq !== frm.doc.hq) {
                    frappe.msgprint({
                        title: __('Warning'),
                        message: __('Selected stockist does not belong to HQ: {0}', [frm.doc.hq]),
                        indicator: 'orange'
                    });
                    frm.set_value('stockist_code', '');
                }
            });
        }
    }
});

frappe.ui.form.on('Scheme Request Item', {
    quantity: function(frm, cdt, cdn) {
        calculate_item_value(frm, cdt, cdn);
    },
    
    special_rate: function(frm, cdt, cdn) {
        calculate_item_value(frm, cdt, cdn);
    },
    
    product_code: function(frm, cdt, cdn) {
        let row = locals[cdt][cdn];
        if (row.product_code) {
            frappe.db.get_value('Product Master', row.product_code, 
                ['product_name', 'pack', 'pts'], (r) => {
                    frappe.model.set_value(cdt, cdn, 'product_name', r.product_name);
                    frappe.model.set_value(cdt, cdn, 'pack', r.pack);
                    frappe.model.set_value(cdt, cdn, 'product_rate', r.pts);
                    
                    // Refresh grid to show history button
                    frm.refresh_field('items');
            });
        }
    },
    
    // Add history button when row is rendered
    items_add: function(frm, cdt, cdn) {
        setTimeout(() => {
            add_grid_history_buttons(frm);
        }, 300);
    }
});
function check_doctor_monthly_limit(frm) {
    if (!frm.doc.doctor_code) return;
    
    frappe.call({
        method: 'scanify.scanify.doctype.scheme_request.scheme_request.get_doctor_monthly_count',
        args: {
            doctor_code: frm.doc.doctor_code,
            application_date: frm.doc.application_date || frappe.datetime.nowdate()
        },
        callback: function(r) {
            if (r.message && r.message.success) {
                let data = r.message;
                let indicator = 'blue';
                let message = `${data.count} of ${data.limit} requests used for ${data.month}`;
                
                if (data.remaining === 0) {
                    indicator = 'red';
                    message = `⚠️ Limit reached: ${data.count}/${data.limit} requests for ${data.month}`;
                } else if (data.remaining === 1) {
                    indicator = 'orange';
                    message = `⚠️ ${data.remaining} slot remaining for ${data.month} (${data.count}/${data.limit} used)`;
                } else {
                    indicator = 'green';
                    message = `${data.remaining} slots remaining for ${data.month} (${data.count}/${data.limit} used)`;
                }
                
                frm.dashboard.set_headline_alert(message, indicator);
                
                // Show persistent message if limit reached
                if (data.remaining === 0 && frm.is_new()) {
                    frappe.msgprint({
                        title: __('Monthly Limit Reached'),
                        message: __('Doctor <b>{0}</b> already has {1} scheme requests in {2}. Maximum 3 requests allowed per doctor per month.', 
                            [frm.doc.doctor_name, data.count, data.month]),
                        indicator: 'red',
                        primary_action: {
                            label: 'Clear Doctor',
                            action: function() {
                                frm.set_value('doctor_code', '');
                                frm.set_value('doctor_name', '');
                            }
                        }
                    });
                }
            }
        }
    });
}

// ==================== NEW: REPEAT REQUEST ====================

function repeat_scheme_request(frm) {
    frappe.confirm(
        __('Are you sure you want to create a new scheme request based on <b>{0}</b>?<br><br>' +
           'This will create a new request with:<br>' +
           '• Same doctor, HQ, stockist, and products<br>' +
           '• Today\'s date as application date<br>' +
           '• Pending approval status', [frm.doc.name]),
        function() {
            // Yes
            frappe.call({
                method: 'scanify.scanify.doctype.scheme_request.scheme_request.repeat_scheme_request',
                args: {
                    source_name: frm.doc.name
                },
                freeze: true,
                freeze_message: __('Creating new scheme request...'),
                callback: function(r) {
                    if (r.message && r.message.success) {
                        frappe.show_alert({
                            message: __('New scheme request created successfully'),
                            indicator: 'green'
                        }, 5);
                        
                        // Redirect to new document
                        frappe.set_route('Form', 'Scheme Request', r.message.doc_name);
                    }
                },
                error: function(r) {
                    frappe.msgprint({
                        title: __('Error'),
                        message: r.message || 'Failed to create repeat request',
                        indicator: 'red'
                    });
                }
            });
        },
        function() {
            // No - do nothing
        }
    );
}


// NEW: Add history buttons to all product rows
function add_grid_history_buttons(frm) {
    // Only for managers during approval
    if (!frappe.user.has_role('Sales Manager') && !frappe.user.has_role('System Manager')) {
        return;
    }
    
    if (!frm.doc.items || frm.doc.items.length === 0) {
        return;
    }
    
    // Add buttons to each grid row
    frm.doc.items.forEach((item, idx) => {
        if (!item.product_code) return;
        
        let grid_row = frm.fields_dict.items.grid.grid_rows[idx];
        if (!grid_row) return;
        
        // Check if button already exists
        if (grid_row.wrapper.find('.btn-view-product-history').length > 0) {
            return;
        }
        
        // Find the row element
        let $row = grid_row.wrapper.find('.grid-row');
        
        // Add button container at the end of the row
        let $btn_container = $('<div class="col grid-static-col" style="padding: 5px;"></div>');
        let $btn = $(`
            <button class="btn btn-xs btn-default btn-view-product-history" 
                    style="white-space: nowrap;" 
                    title="View Product History">
                <i class="fa fa-history"></i> History
            </button>
        `);
        
        $btn.on('click', function(e) {
            e.preventDefault();
            e.stopPropagation();
            show_product_history_dialog(frm, item.product_code);
        });
        
        $btn_container.append($btn);
        $row.append($btn_container);
    });
}

// ALTERNATIVE: Add a main button above the grid to view any product history
frappe.ui.form.on('Scheme Request', {
    refresh: function(frm) {
        // ... existing refresh code ...
        
        // Add "View Product History" button for managers
        if ((frappe.user.has_role('Sales Manager') || frappe.user.has_role('System Manager')) 
            && frm.doc.items && frm.doc.items.length > 0) {
            
            frm.add_custom_button(__('View Product History'), function() {
                show_product_selection_dialog(frm);
            }, __('Reports'));
        }
    }
});

// NEW: Dialog to select which product history to view
function show_product_selection_dialog(frm) {
    if (!frm.doc.items || frm.doc.items.length === 0) {
        frappe.msgprint(__('No products added yet'));
        return;
    }
    
    let products = frm.doc.items.map(item => {
        return {
            label: `${item.product_name || item.product_code} (${item.pack || '-'})`,
            value: item.product_code
        };
    });
    
    frappe.prompt([
        {
            fieldname: 'product_code',
            fieldtype: 'Select',
            label: 'Select Product',
            options: products,
            reqd: 1
        }
    ], (values) => {
        show_product_history_dialog(frm, values.product_code);
    }, __('Select Product to View History'), __('View History'));
}

// Show product historical data with charts
function show_product_history_dialog(frm, product_code) {
    if (!product_code) {
        frappe.msgprint(__('Please select a product first'));
        return;
    }
    
    let d = new frappe.ui.Dialog({
        title: __('Product History: {0}', [product_code]),
        size: 'extra-large',
        fields: [
            {
                fieldname: 'loading',
                fieldtype: 'HTML'
            }
        ]
    });
    
    d.fields_dict.loading.$wrapper.html(`
        <div class="text-center" style="padding: 40px;">
            <i class="fa fa-spinner fa-spin fa-3x text-muted"></i>
            <p style="margin-top: 20px; font-size: 14px;">Loading historical data...</p>
        </div>
    `);
    d.show();
    
    // Fetch product history
    frappe.call({
        method: 'scanify.api.get_product_history_for_scheme',
        args: {
            product_code: product_code,
            doctor_code: frm.doc.doctor_code,
            hq: frm.doc.hq
        },
        callback: function(r) {
            if (r.message && r.message.success) {
                let data = r.message;
                let html = generate_product_history_html(data);
                d.fields_dict.loading.$wrapper.html(html);
                
                // Render charts if data available
                if (data.chart_data && data.chart_data.length > 0) {
                    setTimeout(() => {
                        render_product_history_chart(data.chart_data, 'product-history-chart');
                    }, 100);
                }
            } else {
                d.fields_dict.loading.$wrapper.html(
                    '<div class="alert alert-warning">No historical data found for this product</div>'
                );
            }
        },
        error: function(err) {
            d.fields_dict.loading.$wrapper.html(
                '<div class="alert alert-danger">Error loading data. Please try again.</div>'
            );
        }
    });
}

// Generate HTML for product history
function generate_product_history_html(data) {
    let html = `
        <div class="product-history-container" style="padding: 15px;">
            <div class="row">
                <div class="col-md-6">
                    <h5><i class="fa fa-cube"></i> Product Details</h5>
                    <table class="table table-bordered">
                        <tr><th width="40%">Product Code</th><td>${data.product_code || '-'}</td></tr>
                        <tr><th>Product Name</th><td><strong>${data.product_name || '-'}</strong></td></tr>
                        <tr><th>Pack</th><td>${data.pack || '-'}</td></tr>
                        <tr><th>PTS Rate</th><td><strong>₹${flt(data.pts).toFixed(2)}</strong></td></tr>
                    </table>
                </div>
                <div class="col-md-6">
                    <h5><i class="fa fa-bar-chart"></i> Historical Summary (Last 6 Months)</h5>
                    <table class="table table-bordered">
                        <tr><th width="50%">Total Past Schemes</th><td><strong>${data.total_schemes || 0}</strong></td></tr>
                        <tr><th>Total Quantity Ordered</th><td><strong>${data.total_quantity || 0}</strong></td></tr>
                        <tr><th>Total Value</th><td><strong>₹${flt(data.total_value).toFixed(2)}</strong></td></tr>
                        <tr><th>Last Ordered</th><td>${data.last_order_date || 'Never'}</td></tr>
                    </table>
                </div>
            </div>
            
            <div class="row" style="margin-top: 20px;">
                <div class="col-md-12">
                    <h5><i class="fa fa-line-chart"></i> Scheme History Trend</h5>
                    <div id="product-history-chart" style="height: 300px; border: 1px solid #ddd; border-radius: 4px; padding: 10px;"></div>
                </div>
            </div>
            
            <div class="row" style="margin-top: 20px;">
                <div class="col-md-12">
                    <h5><i class="fa fa-list"></i> Recent Scheme Requests</h5>
                    <div style="max-height: 300px; overflow-y: auto;">
                        <table class="table table-bordered table-striped table-hover">
                            <thead style="background-color: #f5f5f5;">
                                <tr>
                                    <th>Scheme ID</th>
                                    <th>Date</th>
                                    <th>Doctor</th>
                                    <th>Quantity</th>
                                    <th>Value</th>
                                    <th>Status</th>
                                </tr>
                            </thead>
                            <tbody>
    `;
    
    if (data.recent_schemes && data.recent_schemes.length > 0) {
        data.recent_schemes.forEach(scheme => {
            html += `
                <tr>
                    <td><a href="/app/scheme-request/${scheme.name}" target="_blank">${scheme.name}</a></td>
                    <td>${frappe.datetime.str_to_user(scheme.application_date)}</td>
                    <td>${scheme.doctor_name || '-'}</td>
                    <td>${scheme.quantity}</td>
                    <td>₹${flt(scheme.product_value).toFixed(2)}</td>
                    <td><span class="indicator ${get_status_color(scheme.approval_status)}">${scheme.approval_status}</span></td>
                </tr>
            `;
        });
    } else {
        html += '<tr><td colspan="6" class="text-center text-muted">No recent schemes found</td></tr>';
    }
    
    html += `
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    return html;
}

// Render chart using Frappe Charts
function render_product_history_chart(chart_data, container_id) {
    if (!chart_data || chart_data.length === 0) {
        $(`#${container_id}`).html('<p class="text-center text-muted" style="padding: 40px;">No chart data available</p>');
        return;
    }
    
    let labels = chart_data.map(d => d.month);
    let quantities = chart_data.map(d => d.quantity);
    let values = chart_data.map(d => d.value);
    
    new frappe.Chart(`#${container_id}`, {
        title: "Monthly Trend",
        data: {
            labels: labels,
            datasets: [
                {
                    name: "Quantity",
                    values: quantities,
                    chartType: 'bar'
                },
                {
                    name: "Value (₹)",
                    values: values,
                    chartType: 'line'
                }
            ]
        },
        type: 'axis-mixed',
        height: 280,
        colors: ['#5e64ff', '#28a745']
    });
}

function get_status_color(status) {
    const colors = {
        'Approved': 'green',
        'Pending': 'orange',
        'Rejected': 'red',
        'Rerouted': 'blue'
    };
    return colors[status] || 'gray';
}

function show_doctor_search_dialog(frm) {
    let d = new frappe.ui.Dialog({
        title: 'Search Doctor',
        fields: [
            {
                fieldname: 'search_results',
                fieldtype: 'HTML'
            }
        ],
        primary_action_label: 'Close',
        primary_action: function() {
            d.hide();
        }
    });
    
    frappe.call({
        method: 'scanify.api.search_doctors',
        args: { search_term: frm.doc.search_doctor },
        callback: function(r) {
            if (r.message && r.message.length > 0) {
                let html = `
                    <div style="max-height: 400px; overflow-y: auto;">
                        <table class="table table-bordered">
                            <thead>
                                <tr>
                                    <th>Doctor</th>
                                    <th>Place</th>
                                    <th>Specialization</th>
                                    <th>Action</th>
                                </tr>
                            </thead>
                            <tbody>
                `;
                
                r.message.forEach(function(doctor) {
                    html += `
                        <tr>
                            <td><strong>${doctor.doctor_name}</strong><br><small>${doctor.doctor_code}</small></td>
                            <td>${doctor.place || '-'}</td>
                            <td>${doctor.specialization || 'General'}</td>
                            <td>
                                <button class="btn btn-xs btn-primary select-doctor" data-code="${doctor.name}">
                                    Select
                                </button>
                            </td>
                        </tr>
                    `;
                });
                
                html += '</tbody></table></div>';
                d.fields_dict.search_results.$wrapper.html(html);
                
                d.fields_dict.search_results.$wrapper.find('.select-doctor').on('click', function() {
                    let code = $(this).data('code');
                    frm.set_value('doctor_code', code);
                    d.hide();
                });
            } else {
                d.fields_dict.search_results.$wrapper.html(
                    '<div class="alert alert-warning">No doctors found</div>'
                );
            }
        }
    });
    
    d.show();
}

function calculate_item_value(frm, cdt, cdn) {
    let row = locals[cdt][cdn];
    let rate = row.special_rate ? flt(row.special_rate) : flt(row.product_rate);
    let value = flt(row.quantity) * rate;
    frappe.model.set_value(cdt, cdn, 'product_value', value);
    calculate_total(frm);
}

function calculate_total(frm) {
    let total = 0;
    (frm.doc.items || []).forEach(item => {
        total += flt(item.product_value);
    });
    frm.set_value('total_scheme_value', total);
}

function approve_scheme(frm) {
    frappe.prompt([
        {
            fieldname: 'comments',
            fieldtype: 'Small Text',
            label: 'Approval Comments',
            reqd: 1
        }
    ], (values) => {
        frappe.call({
            method: 'scanify.api.approve_scheme_request',
            args: {
                doc_name: frm.doc.name,
                comments: values.comments
            },
            callback: function(r) {
                if (r.message) {
                    frm.reload_doc();
                    frappe.show_alert({
                        message: 'Scheme approved successfully',
                        indicator: 'green'
                    }, 5);
                }
            }
        });
    }, __('Approve Scheme'), __('Approve'));
}

function reject_scheme(frm) {
    frappe.prompt([
        {
            fieldname: 'comments',
            fieldtype: 'Small Text',
            label: 'Rejection Reason',
            reqd: 1
        }
    ], (values) => {
        frappe.call({
            method: 'scanify.api.reject_scheme_request',
            args: {
                doc_name: frm.doc.name,
                comments: values.comments
            },
            callback: function(r) {
                if (r.message) {
                    frm.reload_doc();
                    frappe.show_alert({
                        message: 'Scheme rejected',
                        indicator: 'red'
                    }, 5);
                }
            }
        });
    }, __('Reject Scheme'), __('Reject'));
}

function reroute_scheme(frm) {
    frappe.prompt([
        {
            fieldname: 'comments',
            fieldtype: 'Small Text',
            label: 'Reroute Reason',
            reqd: 1
        }
    ], (values) => {
        frappe.call({
            method: 'scanify.api.reroute_scheme_request',
            args: {
                doc_name: frm.doc.name,
                comments: values.comments
            },
            callback: function(r) {
                if (r.message) {
                    frm.reload_doc();
                    frappe.show_alert({
                        message: 'Scheme rerouted for revision',
                        indicator: 'orange'
                    }, 5);
                }
            }
        });
    }, __('Reroute Scheme'), __('Reroute'));
}

function show_all_attachments(frm) {
    let attachments = [];
    [1, 2, 3, 4].forEach(i => {
        let att = frm.doc[`proof_attachment_${i}`];
        if (att) attachments.push({ num: i, url: att });
    });
    
    let html = '<div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 15px;">';
    attachments.forEach(att => {
        let isImage = att.url.match(/\.(jpg|jpeg|png|gif)$/i);
        html += `
            <div style="border: 2px solid #ddd; border-radius: 8px; padding: 10px; text-align: center;">
                <h5>Document ${att.num}</h5>
                ${isImage ? 
                    `<img src="${att.url}" style="width: 100%; cursor: pointer;" 
                        onclick="window.open('${att.url}', '_blank')"/>` :
                    `<p><i class="fa fa-file"></i> ${att.url.split('/').pop()}</p>`
                }
                <button class="btn btn-sm btn-default" style="margin-top: 8px;" 
                    onclick="window.open('${att.url}', '_blank')">
                    <i class="fa fa-external-link"></i> View/Download
                </button>
            </div>
        `;
    });
    html += '</div>';
    
    frappe.msgprint({
        title: __('Attached Documents'),
        message: html,
        wide: true
    });
}

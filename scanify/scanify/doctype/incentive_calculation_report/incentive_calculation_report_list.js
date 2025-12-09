frappe.listview_settings['Incentive Calculation Report'] = {
    add_fields: ["division", "calculation_type", "period_type", "total_incentive_amount"],
    filters: [
        ["docstatus", "=", 1]
    ],
    get_indicator: function(doc) {
        if (doc.docstatus === 1) {
            return [__("Submitted"), "green"];
        } else if (doc.docstatus === 0) {
            return [__("Draft"), "blue"];
        }
    },
    onload: function(listview) {
        listview.page.add_inner_button(__("Generate New Report"), function() {
            frappe.set_route('Form', 'Incentive Calculation Report', 'new');
        });
    }
};

frappe.listview_settings['Secondary Sales Report'] = {
    add_fields: ["report_type", "from_date", "to_date", "generated_by"],
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
            frappe.set_route('Form', 'Secondary Sales Report', 'new');
        });
    }
};

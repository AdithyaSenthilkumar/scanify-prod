frappe.listview_settings['Product Moving Trend Report'] = {
    add_fields: ["from_date", "to_date", "product_category", "generated_by"],
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
            frappe.set_route('Form', 'Product Moving Trend Report', 'new');
        });
    }
};

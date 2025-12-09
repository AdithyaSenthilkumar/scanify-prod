frappe.listview_settings['Stockist Performance Report'] = {
    add_fields: ["from_date", "to_date", "total_stockists", "total_secondary_value"],
    filters: [
        ["docstatus", "=", 1]
    ],
    get_indicator: function(doc) {
        if (doc.docstatus === 1) {
            return [__("Submitted"), "green"];
        } else {
            return [__("Draft"), "blue"];
        }
    }
};

frappe.listview_settings['Ranking Sheet Report'] = {
    add_fields: ["ranking_type", "division", "period_type", "total_participants"],
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

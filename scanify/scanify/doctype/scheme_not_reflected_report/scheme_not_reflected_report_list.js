frappe.listview_settings['Scheme Not Reflected Report'] = {
    add_fields: ["from_date", "to_date", "total_unreflected_schemes", "reflection_percentage"],
    filters: [
        ["docstatus", "=", 1]
    ],
    get_indicator: function(doc) {
        if (doc.reflection_percentage >= 80) {
            return [__("Good"), "green"];
        } else if (doc.reflection_percentage >= 50) {
            return [__("Moderate"), "orange"];
        } else {
            return [__("Action Needed"), "red"];
        }
    }
};

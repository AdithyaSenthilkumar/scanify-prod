frappe.listview_settings['Doctor Scheme Summary Report'] = {
    add_fields: ["from_date", "to_date", "total_doctors", "total_schemes", "average_reflection_rate"],
    filters: [
        ["docstatus", "=", 1]
    ],
    get_indicator: function(doc) {
        if (doc.average_reflection_rate >= 80) {
            return [__("Excellent"), "green"];
        } else if (doc.average_reflection_rate >= 60) {
            return [__("Good"), "blue"];
        } else if (doc.average_reflection_rate >= 40) {
            return [__("Average"), "orange"];
        } else {
            return [__("Poor"), "red"];
        }
    }
};

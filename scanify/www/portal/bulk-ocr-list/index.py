import frappe

def get_context(context):
    if frappe.session.user == "Guest":
        frappe.throw("You must be logged in", frappe.PermissionError)
    context.no_cache = 1
    user_division = frappe.db.get_value("User", frappe.session.user, "division") or "Prima"
    if hasattr(frappe.session, "user_division") and frappe.session.user_division:
        user_division = frappe.session.user_division
    context.division = user_division

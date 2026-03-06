import frappe


def get_context(context):
    if frappe.session.user == "Guest":
        frappe.throw("Please login to continue", frappe.PermissionError)

    context.no_cache = 1
    user = frappe.session.user

    # Division
    division = None
    if hasattr(frappe.session, "user_division") and frappe.session.user_division:
        division = frappe.session.user_division
    if not division:
        division = frappe.db.get_value("User", user, "division")
    context.division = division or "Prima"


    return context

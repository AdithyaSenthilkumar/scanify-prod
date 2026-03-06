import frappe


def get_context(context):
    if frappe.session.user == "Guest":
        frappe.throw("Please login to continue", frappe.PermissionError)

    context.no_cache = 1
    user = frappe.session.user

    division = None
    if hasattr(frappe.session, "user_division") and frappe.session.user_division:
        division = frappe.session.user_division
    if not division:
        division = frappe.db.get_value("User", user, "division")
    context.division = division or "Prima"

    roles = frappe.get_roles(user)
    if "System Manager" in roles:
        context.user_role = "System Manager"
    elif "Sales Manager" in roles:
        context.user_role = "Sales Manager"
    else:
        context.user_role = "User"

    return context

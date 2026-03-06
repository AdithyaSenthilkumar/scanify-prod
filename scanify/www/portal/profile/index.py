import frappe

def get_context(context):
    if frappe.session.user == "Guest":
        frappe.throw("Please login to continue", frappe.PermissionError)

    context.no_cache = 1
    user = frappe.session.user

    # Get division
    user_division = None
    if hasattr(frappe.session, "user_division") and frappe.session.user_division:
        user_division = frappe.session.user_division
    if not user_division:
        user_division = frappe.db.get_value("User", user, "division")
    if not user_division:
        user_division = "Prima"

    context.division = user_division

    # Get user role
    roles = frappe.get_roles(user)
    if "System Manager" in roles:
        context.user_role = "System Manager"
    elif "Sales Manager" in roles:
        context.user_role = "Sales Manager"
    else:
        context.user_role = "User"

    return context

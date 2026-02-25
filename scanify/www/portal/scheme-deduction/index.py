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

    # Only managers/system managers can perform deductions
    roles = frappe.get_roles(user)
    if "System Manager" not in roles and "Sales Manager" not in roles:
        frappe.throw("You do not have permission to access this page", frappe.PermissionError)

    return context

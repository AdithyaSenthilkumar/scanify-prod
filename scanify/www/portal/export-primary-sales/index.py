import frappe
from scanify.api import get_user_division


def get_context(context):
    if frappe.session.user == "Guest":
        frappe.throw("Please login to continue", frappe.PermissionError)

    context.no_cache = 1
    context.division = get_user_division() or "Prima"

    roles = frappe.get_roles(frappe.session.user)
    if "System Manager" in roles:
        context.user_role = "System Manager"
    elif "Sales Manager" in roles:
        context.user_role = "Sales Manager"
    else:
        context.user_role = "User"

    # Get distinct months available
    context.available_months = frappe.db.sql("""
        SELECT DISTINCT upload_month
        FROM `tabPrimary Sales Data`
        WHERE division = %(division)s
        ORDER BY upload_month DESC
    """, {"division": context.division}, as_dict=1)

    return context

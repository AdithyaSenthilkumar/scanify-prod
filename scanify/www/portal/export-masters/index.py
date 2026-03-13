import frappe
from scanify.api import get_user_division

def get_context(context):
    context.no_cache = 1

    # Check authentication
    if frappe.session.user == 'Guest':
        frappe.throw("Please login to continue", frappe.PermissionError)

    # Get user division
    context.division = get_user_division()

    # Get user role
    roles = frappe.get_roles(frappe.session.user)
    if 'System Manager' in roles:
        context.user_role = 'System Manager'
    elif 'Sales Manager' in roles:
        context.user_role = 'Sales Manager'
    else:
        context.user_role = 'User'

    return context

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
    context.user_role = get_user_role(frappe.session.user)
    
    return context

def get_user_role(user):
    roles = frappe.get_roles(user)
    if 'System Manager' in roles:
        return 'System Manager'
    elif 'Sales Manager' in roles:
        return 'Sales Manager'
    return 'User'

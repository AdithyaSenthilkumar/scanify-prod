import frappe
from scanify.api import get_user_division

def get_context(context):
    context.no_cache = 1
    
    # Check authentication
    if frappe.session.user == 'Guest':
        frappe.throw('Please login to continue', frappe.PermissionError)
    
    # Get user division
    context.division = get_user_division()
    context.user_role = get_user_role(frappe.session.user)
    
    return context

def get_user_division():
    """Get user's division from HQ Master or User"""
    try:
        hq_division = frappe.db.get_value('HQ Master', {'email': frappe.session.user}, 'division')
        if hq_division:
            return hq_division
        return frappe.db.get_value('User', frappe.session.user, 'division') or 'Prima'
    except:
        return 'Prima'

def get_user_role(user):
    """Determine user's primary role"""
    roles = frappe.get_roles(user)
    if 'System Manager' in roles:
        return 'System Manager'
    elif 'Sales Manager' in roles:
        return 'Sales Manager'
    return 'User'

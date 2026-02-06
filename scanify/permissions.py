import frappe
from scanify.api import get_user_division
def has_scheme_permission(doc, user):
    """Check if user has permission to access scheme"""
    if frappe.session.user == "Administrator":
        return True
    
    # User can access their own schemes
    if doc.requestedby == user:
        return True
    
    # Managers can access schemes in their division
    if "Sales Manager" in frappe.get_roles(user):
        user_division = get_user_division()
        return doc.division == user_division
    
    return False

def has_statement_permission(doc, user):
    """Check if user has permission to access statement"""
    if frappe.session.user == "Administrator":
        return True
    
    # Check division access
    user_division = get_user_division()
    return doc.division == user_division

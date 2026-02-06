import frappe
from frappe.utils import get_first_day, nowdate
from scanify.api import get_user_division

def get_context(context):
    context.no_cache = 1
    
    # Check authentication
    if frappe.session.user == "Guest":
        frappe.throw("Please login to continue", frappe.PermissionError)
    
    user = frappe.session.user
    
    # Get user division (from session or database)
    context.division = get_user_division()
    
    # Get user role
    context.user_role = get_user_role(user)
    
    # Get dashboard stats filtered by division
    context.stats = get_dashboard_stats(user, context.division)
    
    return context


def get_user_role(user):
    """Get user's primary role"""
    roles = frappe.get_roles(user)
    if "Sales Manager" in roles:
        return "Sales Manager"
    elif "System Manager" in roles:
        return "System Manager"
    else:
        return "Sales Officer"


def get_dashboard_stats(user, division):
    """Get dashboard statistics filtered by division"""
    stats = {}
    
    try:
        # Pending schemes count (filtered by division)
        stats["pending_schemes"] = frappe.db.count(
            "Scheme Request",
            {"approval_status": "Pending", "docstatus": 0, "division": division}
        )
        
        # Approved this month (filtered by division)
        stats["approved_this_month"] = frappe.db.sql("""
            SELECT COUNT(DISTINCT sr.name)
            FROM `tabScheme Request` sr
            INNER JOIN `tabScheme Approval Log` sal ON sal.parent = sr.name
            WHERE sr.division = %s
            AND sal.action = 'Approved'
            AND sal.action_date >= %s
        """, (division, get_first_day(nowdate())))[0][0]
        
        # Pending statements (filtered by division if stockist has division)
        stats["pending_statements"] = frappe.db.count("Stockist Statement", {"docstatus": 0})
        
        # Total active doctors
        stats["total_doctors"] = frappe.db.count("Doctor Master", {"status": "Active"})
        
    except Exception as e:
        frappe.log_error(f"Error getting dashboard stats: {str(e)}")
        stats = {
            "pending_schemes": 0,
            "approved_this_month": 0,
            "pending_statements": 0,
            "total_doctors": 0
        }
    
    return stats

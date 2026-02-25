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
    
    # Get recent scheme requests for this division
    context.recent_requests = get_recent_requests(context.division)
    
    # Helpful greeting based on time of day
    from datetime import datetime
    hour = datetime.now().hour
    if hour < 12:
        context.greeting = "Good Morning"
    elif hour < 17:
        context.greeting = "Good Afternoon"
    else:
        context.greeting = "Good Evening"
        
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
        
        # Active HQ Count for this division
        stats["active_hqs"] = frappe.db.count("HQ Master", {"division": division, "status": "Active"})
        
        # Chart 1 & 2 shared: build last 14 day date series
        from frappe.utils import add_days, getdate
        start_date = add_days(nowdate(), -13)

        # Scheme requests per day (last 14 days)
        scheme_activity = frappe.db.sql("""
            SELECT DATE(creation) as date, COUNT(*) as count
            FROM `tabScheme Request`
            WHERE division = %s AND creation >= %s
            GROUP BY DATE(creation)
            ORDER BY DATE(creation) ASC
        """, (division, start_date), as_dict=1)
        for d in scheme_activity:
            if d.get('date'): d['date'] = str(d['date'])
        stats["scheme_activity"] = scheme_activity

        # Stock statements per day (last 14 days)
        statement_activity = frappe.db.sql("""
            SELECT DATE(creation) as date, COUNT(*) as count
            FROM `tabStockist Statement`
            WHERE creation >= %s
            GROUP BY DATE(creation)
            ORDER BY DATE(creation) ASC
        """, (start_date,), as_dict=1)
        for d in statement_activity:
            if d.get('date'): d['date'] = str(d['date'])
        stats["statement_activity"] = statement_activity

        # Chart 2: Monthly approval funnel — last 6 months breakdown
        monthly_funnel = frappe.db.sql("""
            SELECT
                DATE_FORMAT(creation, '%%b %%Y') as month,
                DATE_FORMAT(creation, '%%Y-%%m') as sort_key,
                SUM(CASE WHEN approval_status = 'Approved' THEN 1 ELSE 0 END) as approved,
                SUM(CASE WHEN approval_status = 'Pending' THEN 1 ELSE 0 END) as pending,
                SUM(CASE WHEN approval_status = 'Rejected' THEN 1 ELSE 0 END) as rejected
            FROM `tabScheme Request`
            WHERE division = %s AND creation >= DATE_SUB(NOW(), INTERVAL 6 MONTH)
            GROUP BY DATE_FORMAT(creation, '%%Y-%%m'), DATE_FORMAT(creation, '%%b %%Y')
            ORDER BY sort_key ASC
        """, (division,), as_dict=1)
        stats["monthly_funnel"] = monthly_funnel

    except Exception as e:
        frappe.log_error(f"Error getting dashboard stats: {str(e)}")
        stats = {
            "pending_schemes": 0,
            "approved_this_month": 0,
            "pending_statements": 0,
            "active_hqs": 0,
            "scheme_activity": [],
            "statement_activity": [],
            "monthly_funnel": []
        }
    
    return stats


def get_recent_requests(division, limit=5):
    """Get recent scheme requests for the division"""
    return frappe.get_all(
        "Scheme Request",
        fields=["name", "creation", "doctor_name", "approval_status", "total_scheme_value"],
        filters={"division": division},
        order_by="creation desc",
        limit=limit
    )

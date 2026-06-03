import frappe
from scanify.api import get_user_division, get_stockist_report_filter_options


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

    # Cascading org filter options (zones, regions, teams, hqs)
    opts = get_stockist_report_filter_options(context.division)
    context.zones = opts.get("zones", [])
    context.regions = opts.get("regions", [])
    context.teams = opts.get("teams", [])
    context.hqs = opts.get("hqs", [])

    # Distinct months with secondary sales data (draft + submitted)
    context.available_months = frappe.db.sql("""
        SELECT DISTINCT DATE_FORMAT(statement_month, '%%Y-%%m') AS month
        FROM `tabStockist Statement`
        WHERE division = %(division)s AND docstatus IN (0, 1)
        ORDER BY month DESC
    """, {"division": context.division}, as_dict=1)

    return context

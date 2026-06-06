import frappe
from scanify.api import get_user_division, get_stockist_report_filter_options

def get_context(context):
    context.no_cache = 1

    if frappe.session.user == 'Guest':
        frappe.throw('Please login to continue', frappe.PermissionError)

    division = get_user_division()
    context.division = division
    context.user_role = get_user_role(frappe.session.user)

    opts = get_stockist_report_filter_options(division)
    context.zones = opts.get("zones", [])
    context.regions = opts.get("regions", [])
    context.teams = opts.get("teams", [])
    context.hqs = opts.get("hqs", [])
    context.stockists = opts.get("stockists", [])

    return context

def get_user_role(user):
    roles = frappe.get_roles(user)
    if 'System Manager' in roles:
        return 'System Manager'
    elif 'Sales Manager' in roles:
        return 'Sales Manager'
    return 'User'

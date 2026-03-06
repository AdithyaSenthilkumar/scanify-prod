import frappe
from scanify.api import get_user_division, get_insights_filter_options

def get_context(context):
    if frappe.session.user == "Guest":
        frappe.throw("Please login to continue", frappe.PermissionError)

    context.no_cache = 1
    division = get_user_division()
    context.division = division

    # Pre-load filter options for dropdowns
    filter_opts = get_insights_filter_options(division)
    context.regions = filter_opts.get("regions", [])
    context.teams = filter_opts.get("teams", [])
    context.hqs = filter_opts.get("hqs", [])
    context.financial_years = filter_opts.get("financial_years", [])

    return context

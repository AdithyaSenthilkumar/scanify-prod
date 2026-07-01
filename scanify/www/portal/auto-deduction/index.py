import frappe
from scanify.api import get_user_division, get_scheme_report_filter_options


def get_context(context):
    if frappe.session.user == "Guest":
        frappe.throw("Please login to continue", frappe.PermissionError)

    context.no_cache = 1
    division = get_user_division() or "Prima"
    context.division = division

    # Hierarchy filter options (Zone -> Region -> Team -> HQ), cascaded client-side
    opts = get_scheme_report_filter_options(division)
    context.zones = opts.get("zones", [])
    context.regions = opts.get("regions", [])
    context.teams = opts.get("teams", [])
    context.hqs = opts.get("hqs", [])

    return context

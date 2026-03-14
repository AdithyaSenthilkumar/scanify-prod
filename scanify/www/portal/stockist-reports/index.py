import frappe
from scanify.api import get_user_division, get_stockist_report_filter_options


def get_context(context):
    if frappe.session.user == "Guest":
        frappe.throw("You must be logged in.", frappe.PermissionError)

    context.no_cache = 1
    division = get_user_division() or "Prima"
    context.division = division

    opts = get_stockist_report_filter_options(division)
    context.regions = opts.get("regions", [])
    context.teams = opts.get("teams", [])
    context.hqs = opts.get("hqs", [])
    context.stockists = opts.get("stockists", [])
    context.months = opts.get("months", [])
    context.statement_months = opts.get("statement_months", [])

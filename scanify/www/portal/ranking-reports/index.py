import frappe
from scanify.api import get_user_division, get_ranking_report_filter_options


def get_context(context):
    if frappe.session.user == "Guest":
        frappe.throw("You must be logged in.", frappe.PermissionError)

    context.no_cache = 1
    division = get_user_division() or "Prima"
    context.division = division

    opts = get_ranking_report_filter_options(division)
    context.zones = opts.get("zones", [])
    context.regions = opts.get("regions", [])
    context.teams = opts.get("teams", [])
    context.hqs = opts.get("hqs", [])
    context.stockists = opts.get("stockists", [])
    context.products = opts.get("products", [])
    context.doctors = opts.get("doctors", [])

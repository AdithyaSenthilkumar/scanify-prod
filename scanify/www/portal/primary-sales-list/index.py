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

    # Hierarchical filter options (code + name + parent links)
    opts = get_stockist_report_filter_options(context.division)
    context.zones = opts.get("zones", [])
    context.regions = opts.get("regions", [])
    context.teams = opts.get("teams", [])
    context.hqs = opts.get("hqs", [])

    context.product_groups = [
        "DREZ GROUP", "AMINORICH GROUP", "OTHER PRODUCTS", "JUSDEE GROUP",
        "DREZ 10% GROUP", "HOSPITAL PRODUCTS", "CONTUS GROUP", "XPTUM GROUP",
        "DYGERM GROUP", "ORTHO GROUP", "DENTIST GROUP", "GYNAE GROUP"
    ]

    # Get distinct months available
    context.available_months = frappe.db.sql("""
        SELECT DISTINCT upload_month
        FROM `tabPrimary Sales Data`
        WHERE division = %(division)s
        ORDER BY upload_month DESC
    """, {"division": context.division}, as_dict=1)

    return context

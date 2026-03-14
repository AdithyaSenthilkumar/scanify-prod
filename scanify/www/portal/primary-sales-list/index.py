import frappe
from scanify.api import get_user_division


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

    # Filter options for dropdowns
    div_filter = ["in", [context.division, "Both"]]

    context.zones = frappe.get_all(
        "Zone Master",
        filters={"division": div_filter, "status": "Active"},
        fields=["name", "zone_name"],
        order_by="zone_name asc"
    )
    context.regions = frappe.get_all(
        "Region Master",
        filters={"division": div_filter, "status": "Active"},
        fields=["name", "region_name"],
        order_by="region_name asc"
    )
    context.teams = frappe.get_all(
        "Team Master",
        filters={"division": div_filter, "status": "Active"},
        fields=["name", "team_name"],
        order_by="team_name asc"
    )
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

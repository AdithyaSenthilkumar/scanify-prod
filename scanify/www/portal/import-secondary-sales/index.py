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

    # Recent backfilled statements for this division (most recent first)
    stockists = frappe.get_all(
        "Stockist Master",
        filters={"division": ["in", [context.division, "Both"]]},
        pluck="name",
    )
    context.recent_statements = []
    if stockists:
        context.recent_statements = frappe.get_all(
            "Stockist Statement",
            filters={"stockist_code": ["in", stockists]},
            fields=["name", "stockist_code", "stockist_name", "statement_month",
                    "docstatus", "qc_confidence", "creation"],
            order_by="creation desc",
            limit=10,
        )

    return context

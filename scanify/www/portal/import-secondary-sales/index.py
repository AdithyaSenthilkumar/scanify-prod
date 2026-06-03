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

    # Recent import runs (persistent log) for this division
    context.recent_uploads = frappe.get_all(
        "Secondary Sales Upload",
        filters={"division": context.division},
        fields=["name", "upload_month", "status", "upload_date", "uploaded_by",
                "total_data_rows", "statements_created", "items_created",
                "skipped_existing", "unmatched_stockists", "inactive_stockists",
                "unmapped_products", "create_errors"],
        order_by="creation desc",
        limit=10,
    )

    return context

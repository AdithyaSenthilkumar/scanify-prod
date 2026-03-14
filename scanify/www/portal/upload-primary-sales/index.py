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

    # Get recent uploads for this division
    context.recent_uploads = frappe.get_all(
        "Primary Sales Upload",
        filters={"division": context.division},
        fields=["name", "upload_month", "status", "total_rows", "success_count", "error_count", "upload_date", "uploaded_by"],
        order_by="creation desc",
        limit=10
    )

    return context

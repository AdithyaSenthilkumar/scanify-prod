import frappe

def get_context(context):
    if frappe.session.user == "Guest":
        frappe.throw("Please login to continue", frappe.PermissionError)

    context.no_cache = 1
    user = frappe.session.user

    scheme_name = frappe.form_dict.get("name")
    if not scheme_name:
        frappe.throw("Scheme name is required")

    # Get division
    user_division = None
    if hasattr(frappe.session, "user_division") and frappe.session.user_division:
        user_division = frappe.session.user_division
    if not user_division:
        user_division = frappe.db.get_value("User", user, "division")
    if not user_division:
        user_division = "Prima"

    context.division = user_division
    context.scheme_name = scheme_name

    # Approval is a manager function: portal Admin/HO (or a Frappe Sales/System Manager).
    from scanify.permissions import get_portal_role
    roles = frappe.get_roles(user)
    portal_role = get_portal_role(user)
    context.portal_role = portal_role
    context.user_role = portal_role
    context.is_manager = (
        portal_role in ("Admin", "HO")
        or "System Manager" in roles
        or "Sales Manager" in roles
    )

    return context

import frappe


def get_context(context):
    if frappe.session.user == "Guest":
        frappe.throw("Please login to continue", frappe.PermissionError)

    context.no_cache = 1

    # The page fetches everything it needs from scanify.api.get_my_profile; the
    # context values below are only used for the very first server-rendered paint.
    from scanify.permissions import get_portal_role
    from scanify.api import get_user_division

    context.user_role = get_portal_role(frappe.session.user)
    context.division = get_user_division() or ""
    return context

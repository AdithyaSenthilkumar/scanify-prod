import frappe

def get_context(context):
    if frappe.session.user == "Guest":
        frappe.throw("Please login to continue", frappe.PermissionError)

    context.no_cache = 1
    user = frappe.session.user

    # Get division
    user_division = None
    if hasattr(frappe.session, "user_division") and frappe.session.user_division:
        user_division = frappe.session.user_division
    if not user_division:
        user_division = frappe.db.get_value("User", user, "division")
    if not user_division:
        user_division = "Prima"

    context.division = user_division

    # Regional roles work inside their own regions, so this page reads as a status view
    # for them ("Scheme Status") rather than the full request list.
    from scanify.permissions import scheme_list_label, is_regional_role
    context.page_title = scheme_list_label(user)
    context.page_subtitle = ("Status of scheme requests in your regions for"
                             if is_regional_role(user) else "All scheme requests for")

    # Get user role
    roles = frappe.get_roles(user)
    if "System Manager" in roles:
        context.user_role = "System Manager"
        context.is_manager = True
    elif "Sales Manager" in roles:
        context.user_role = "Sales Manager"
        context.is_manager = True
    else:
        context.user_role = "User"
        context.is_manager = False

    return context

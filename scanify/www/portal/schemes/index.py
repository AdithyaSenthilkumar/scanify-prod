import frappe

def get_context(context):
    context.no_cache = 1
    
    if frappe.session.user == "Guest":
        frappe.throw("Please login to continue", frappe.PermissionError)
    
    user = frappe.session.user
    context.division = get_user_division(user)
    context.user_role = get_user_role(user)
    
    return context

def get_user_division(user):
    try:
        hq_division = frappe.db.get_value("HQ Master", {"email": user}, "division")
        if hq_division:
            return hq_division
        return frappe.db.get_value("User", user, "division") or "Prima"
    except:
        return "Prima"

def get_user_role(user):
    roles = frappe.get_roles(user)
    if "System Manager" in roles:
        return "System Manager"
    elif "Sales Manager" in roles:
        return "Sales Manager"
    return "User"

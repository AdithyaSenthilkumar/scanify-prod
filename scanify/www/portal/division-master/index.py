import frappe

def get_context(context):
    context.no_cache = 1
    division = frappe.cache().hget("user_division", frappe.session.user) or "Prima"
    context.division = division
    return context

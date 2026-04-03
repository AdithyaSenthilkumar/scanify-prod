import frappe
from scanify.api import get_user_division

def get_context(context):
    if frappe.session.user == "Guest":
        frappe.throw("Please login to continue", frappe.PermissionError)

    # Check if chatbot is enabled in settings
    chatbot_enabled = frappe.db.get_single_value("Scanify Settings", "enable_chatbot")
    if not chatbot_enabled:
        frappe.throw("Chatbot is currently disabled. Contact your administrator to enable it.", frappe.PermissionError)

    context.no_cache = 1
    division = get_user_division()
    context.division = division
    context.user_fullname = frappe.session.user_fullname or frappe.session.user

    return context

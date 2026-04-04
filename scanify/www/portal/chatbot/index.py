import frappe
from scanify.api import get_user_division

def get_context(context):
    if frappe.session.user == "Guest":
        frappe.throw("Please login to continue", frappe.PermissionError)

    context.no_cache = 1

    # Check if chatbot is enabled in settings (bypass cache)
    chatbot_enabled = frappe.db.get_value("Scanify Settings", "Scanify Settings", "enable_chatbot")
    context.chatbot_disabled = not chatbot_enabled

    if context.chatbot_disabled:
        return context

    division = get_user_division()
    context.division = division
    context.user_fullname = frappe.session.user_fullname or frappe.session.user

    return context

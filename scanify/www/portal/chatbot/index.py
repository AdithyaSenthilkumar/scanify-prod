import frappe
from scanify.api import get_user_division

def get_context(context):
    if frappe.session.user == "Guest":
        frappe.throw("Please login to continue", frappe.PermissionError)

    context.no_cache = 1

    # Check if chatbot is enabled - use same lookup pattern as get_gemini_settings
    settings_name = frappe.db.get_value("Scanify Settings", {"company_name": "Stedman Pharmaceuticals"}, "name")
    if not settings_name:
        settings_name = "Scanify Settings"
    chatbot_enabled = frappe.db.get_value("Scanify Settings", settings_name, "enable_chatbot")
    context.chatbot_disabled = not chatbot_enabled

    if context.chatbot_disabled:
        return context

    division = get_user_division()
    context.division = division
    context.user_fullname = frappe.session.user_fullname or frappe.session.user

    return context

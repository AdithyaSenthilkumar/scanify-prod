import frappe

def get_context(context):
    if frappe.session.user == "Guest":
        frappe.throw("Please login to continue", frappe.PermissionError)

    context.no_cache = 1

    doc_name = frappe.form_dict.get("name") or frappe.form_dict.get("id")
    if not doc_name:
        frappe.throw("Statement name is required")

    context.doc_name = doc_name

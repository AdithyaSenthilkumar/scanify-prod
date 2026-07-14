import frappe


def execute():
    """Add per-user scheme-email routing fields to the User doctype.

    Approved-scheme emails are addressed using the To/CC configured on the scheme
    requestor's profile (combined with the Team Master's To/CC). These are custom
    fields so they survive framework upgrades. Idempotent.
    """
    fields = [
        {
            "fieldname": "scheme_to_email",
            "label": "Scheme Email - To",
            "fieldtype": "Data",
            "options": "Email",
            "insert_after": "division",
            "description": "Primary recipient for this user's approved-scheme emails. One email only.",
        },
        {
            "fieldname": "scheme_cc_emails",
            "label": "Scheme Email - CC",
            "fieldtype": "Small Text",
            "insert_after": "scheme_to_email",
            "description": "CC recipients for this user's approved-scheme emails. Separate multiple with comma or new line.",
        },
    ]

    for f in fields:
        cf_name = f"User-{f['fieldname']}"
        if frappe.db.exists("Custom Field", cf_name):
            print(f"• Custom Field {cf_name} already exists")
            continue
        frappe.get_doc({
            "doctype": "Custom Field",
            "dt": "User",
            "translatable": 0,
            "no_copy": 0,
            "reqd": 0,
            "unique": 0,
            "read_only": 0,
            **f,
        }).insert(ignore_permissions=True)
        print(f"✓ Added Custom Field {cf_name}")

    frappe.db.commit()

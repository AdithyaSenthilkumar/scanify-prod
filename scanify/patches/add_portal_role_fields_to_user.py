import frappe


def execute():
    """Add portal role + division/region mapping fields to User, and migrate existing
    portal users to the Admin role so the rollout doesn't lock anyone out. Idempotent."""
    fields = [
        {
            "fieldname": "portal_role",
            "label": "Portal Role",
            "fieldtype": "Select",
            "options": "\nAdmin\nHO\nRegional User (Future)\nRegional User",
            "insert_after": "division",
            "description": "Single portal role. Drives sidebar/page access. Admin = full portal.",
        },
        {
            "fieldname": "allowed_divisions",
            "label": "Allowed Divisions",
            "fieldtype": "Small Text",
            "insert_after": "portal_role",
            "description": "Divisions this user may access (comma or newline separated). Ignored for Admin.",
        },
        {
            "fieldname": "allowed_regions",
            "label": "Allowed Regions",
            "fieldtype": "Small Text",
            "insert_after": "allowed_divisions",
            "description": "Region codes this user may access (comma or newline separated). Ignored for Admin.",
        },
    ]

    for f in fields:
        cf_name = f"User-{f['fieldname']}"
        if frappe.db.exists("Custom Field", cf_name):
            print(f"• Custom Field {cf_name} already exists")
            continue
        frappe.get_doc({
            "doctype": "Custom Field", "dt": "User",
            "translatable": 0, "no_copy": 0, "reqd": 0, "unique": 0, "read_only": 0,
            **f,
        }).insert(ignore_permissions=True)
        print(f"✓ Added Custom Field {cf_name}")

    frappe.db.commit()

    # Migrate existing portal users → Admin (continuity). Only those without a role set
    # who currently have access via UATadmin or System Manager.
    users = frappe.get_all("User", filters={"enabled": 1, "name": ["not in", ["Guest"]]},
                           pluck="name")
    migrated = 0
    for u in users:
        if frappe.db.get_value("User", u, "portal_role"):
            continue
        roles = frappe.get_roles(u)
        if "UATadmin" in roles or "System Manager" in roles or u == "Administrator":
            frappe.db.set_value("User", u, "portal_role", "Admin", update_modified=False)
            migrated += 1
    frappe.db.commit()
    print(f"✓ Migrated {migrated} existing user(s) to portal_role=Admin")

import frappe


def execute():
    """Grant each portal user the underlying Frappe roles their portal_role needs, so
    portal operations (create scheme, edit masters, submit, deductions) don't fail on
    Frappe doctype permissions. Add-only (never strips). Idempotent."""
    from scanify.permissions import sync_frappe_roles

    users = frappe.get_all("User", filters={"enabled": 1, "name": ["not in", ["Guest"]]},
                           pluck="name")
    synced = 0
    for u in users:
        if u == "Administrator":
            continue
        if not frappe.db.get_value("User", u, "portal_role"):
            continue
        before = set(frappe.get_roles(u))
        sync_frappe_roles(u, prune=False)
        if set(frappe.get_roles(u)) != before:
            synced += 1
    frappe.db.commit()
    print(f"✓ Synced Frappe roles for {synced} portal user(s)")

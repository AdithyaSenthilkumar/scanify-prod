import frappe


def execute():
    """Repair portal_role <-> Frappe-role drift.

    Some portal users (created before the sync existed, or set to a portal role via the
    Desk / an import / a direct db write) ended up as portal Admin / HO / Regional without
    the underlying Frappe roles their role needs (System Manager, Sales Manager, Sales
    User, …). The portal treats them by portal_role, so sidebar + page guards pass, but
    every real write (save a master, delete a scheme) then fails Frappe's doctype
    permission check with 'does not have doctype access via role permission'.

    Re-grant each enabled portal user the roles their portal_role needs. Add-only
    (prune=False, never strips a legacy role like UATadmin), idempotent."""
    from scanify.permissions import sync_frappe_roles

    users = frappe.get_all(
        "User",
        filters={
            "enabled": 1,
            "portal_role": ["is", "set"],
            "name": ["not in", ["Guest", "Administrator"]],
        },
        pluck="name",
    )
    fixed = 0
    for u in users:
        before = set(frappe.get_roles(u))
        sync_frappe_roles(u, prune=False)
        if set(frappe.get_roles(u)) != before:
            fixed += 1
    frappe.db.commit()
    print(f"✓ Re-synced Frappe roles for {len(users)} portal user(s); {fixed} updated")

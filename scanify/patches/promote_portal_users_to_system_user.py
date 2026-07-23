import frappe


def execute():
    """Fix A — make every portal account a System User with the Frappe roles its
    portal_role needs.

    Portal access is gated by User.portal_role, but the real doctype permissions
    (read/write/submit, and private-file download for scheme attachments) come from the
    underlying Frappe roles synced by scanify.permissions.sync_frappe_roles. Frappe only
    lets a *System User* hold desk-access roles (Sales User/Manager, System Manager,
    Stock Manager) — a Website User silently drops them on every save. Historical portal
    accounts created as Website Users therefore pass the portal gate but fail every real
    write and can't download scheme attachments ('does not have doctype access via role
    permission').

    Re-syncing the roles (add-only) both grants the permissions and — because those roles
    carry desk_access — auto-promotes the account to System User via Frappe's own
    set_system_user(). Add-only (never strips a legacy role) and idempotent, so it is safe
    to re-run. Disabled users are included too, so re-enabling one later Just Works."""
    from scanify.permissions import sync_frappe_roles

    users = frappe.get_all(
        "User",
        filters={
            "portal_role": ["is", "set"],
            "name": ["not in", ["Guest", "Administrator"]],
        },
        fields=["name", "user_type"],
    )
    synced = promoted = 0
    for u in users:
        before = set(frappe.get_roles(u.name))
        try:
            sync_frappe_roles(u.name, prune=False)
        except Exception:
            frappe.log_error(frappe.get_traceback(), f"promote_portal_users: {u.name}")
            continue
        if set(frappe.get_roles(u.name)) != before:
            synced += 1
        if u.user_type == "Website User" and \
                frappe.db.get_value("User", u.name, "user_type") == "System User":
            promoted += 1
    frappe.db.commit()
    print(f"✓ Fix A: processed {len(users)} portal user(s); {synced} role-synced, "
          f"{promoted} promoted Website User → System User")

import frappe
import functools

# ─────────────────────────────────────────────────────────────────────────────
# Portal roles & access control
#
# A user has exactly ONE portal role (User.portal_role). This module is the single
# source of truth for: which processes each role may use, which route maps to which
# process, and how non-admin users are scoped to their mapped divisions/regions.
# ─────────────────────────────────────────────────────────────────────────────

ROLE_ADMIN = "Admin"
ROLE_HO = "HO"
ROLE_RF = "Regional User (Future)"
ROLE_R = "Regional User"
PORTAL_ROLES = [ROLE_ADMIN, ROLE_HO, ROLE_RF, ROLE_R]
_ALL = set(PORTAL_ROLES)

# process key -> set of roles allowed to use it
PROCESS_ROLES = {
    "dashboard":       _ALL,
    "masters":         {ROLE_ADMIN},
    "primary_upload":  {ROLE_ADMIN},
    "primary_view":    {ROLE_ADMIN},
    "secondary":       {ROLE_ADMIN, ROLE_HO},        # OCR / statements
    "secondary_admin": {ROLE_ADMIN},                 # reload / delete statement
    "scheme_list":     _ALL,
    "scheme_new":      {ROLE_ADMIN, ROLE_HO, ROLE_RF},
    "scheme_repeat":   _ALL,
    "scheme_email":    {ROLE_ADMIN, ROLE_HO},
    "deductions":      {ROLE_ADMIN},
    "scheme_delete":   {ROLE_ADMIN},
    "sales_targets":   {ROLE_ADMIN},
    "chatbot":         {ROLE_ADMIN},
    "reports":         _ALL,
    "audit":           {ROLE_ADMIN},
    "users":           {ROLE_ADMIN},
}

# portal route (segment after /portal/) -> process key
ROUTE_PROCESS = {
    "portal": "dashboard", "": "dashboard", "profile": "dashboard",
    "masters": "masters", "division-master": "masters", "export-masters": "masters",
    "upload-primary-sales": "primary_upload",
    "primary-sales-list": "primary_view", "export-primary-sales": "primary_view",
    "stock-statements-list": "secondary", "stock-statements": "secondary",
    "manual-statement-entry": "secondary", "bulk-ocr-list": "secondary",
    "bulk-ocr-new": "secondary", "bulk-ocr-view": "secondary", "statement-view": "secondary",
    "import-secondary-sales": "secondary", "export-secondary-sales": "secondary",
    "reload-statements": "secondary_admin", "delete-statement": "secondary_admin",
    "scheme-list": "scheme_list", "scheme-detail": "scheme_list",
    "scheme-new": "scheme_new",
    "scheme-repeat": "scheme_repeat",
    "scheme-email": "scheme_email",
    "scheme-deduction": "deductions", "auto-deduction": "deductions",
    "scheme-deduction-list": "deductions",
    "scheme-delete-revert": "scheme_delete",
    "sales-targets": "sales_targets", "sales-targets-list": "sales_targets",
    "chatbot": "chatbot",
    "insights": "reports", "stockist-reports": "reports", "scheme-reports": "reports",
    "ranking-reports": "reports", "year-wise-report": "reports",
    "audit-trail": "audit",
    "users": "users",
}


# Underlying Frappe roles granted per portal role. The portal (sidebar + page guard)
# does the ACCESS control by portal_role; these Frappe roles only supply the doctype
# PERMISSIONS the allowed operations need (create scheme, edit masters, submit, …).
# Admin gets the full app set; scheme-only roles get Sales User (create) etc.
PORTAL_FRAPPE_ROLES = {
    ROLE_ADMIN: ["System Manager", "Sales Manager", "Sales User", "Stock Manager"],
    ROLE_HO:    ["Sales Manager", "Sales User"],
    ROLE_RF:    ["Sales User"],
    ROLE_R:     ["Sales User"],
}
# The set of Frappe roles this module manages on portal users (safe to add/prune).
APP_MANAGED_ROLES = {"System Manager", "Sales Manager", "Sales User", "Stock Manager"}


def sync_frappe_roles(user, portal_role=None, prune=True):
    """Grant the Frappe roles a portal user's role needs (and optionally strip the
    app-managed roles it should no longer have). Never touches Administrator, and
    never strips the acting user's own System Manager (avoids self-lockout).

    A portal user's Frappe roles are managed here from portal_role, so any Role Profile
    bound to the account is DETACHED first: Frappe's populate_role_profile_roles pins a
    user's roles to exactly the profile's set on every save, which would otherwise strip
    the desk-access roles we grant and — because user_type is derived from whether any
    role has desk access — silently keep the account a Website User (no doctype perms,
    no scheme-attachment download). The whole update is one save so validation sees the
    detached profile and the granted roles together."""
    if user == "Administrator":
        return
    portal_role = portal_role or get_portal_role(user)
    desired = [r for r in PORTAL_FRAPPE_ROLES.get(portal_role, []) if frappe.db.exists("Role", r)]
    have = set(frappe.get_roles(user))
    doc = frappe.get_doc("User", user)
    dirty = False

    # Detach any Role Profile — it competes with portal-managed roles (see docstring).
    if doc.get("role_profiles") or doc.get("role_profile_name"):
        doc.role_profiles = []
        doc.role_profile_name = None
        dirty = True

    for r in [r for r in desired if r not in have]:
        doc.append("roles", {"role": r})
        dirty = True

    if prune:
        to_remove = {r for r in (APP_MANAGED_ROLES - set(desired)) if r in have}
        if user == frappe.session.user:
            to_remove.discard("System Manager")  # never strip the acting admin's own SM
        if to_remove:
            doc.roles = [r for r in doc.roles if r.role not in to_remove]
            dirty = True

    if dirty:
        doc.flags.ignore_permissions = True
        doc.save()


def sync_user_frappe_roles(doc, method=None):
    """User doc_event (after_insert / on_update): keep the underlying Frappe roles
    aligned with the user's portal_role, so portal access is ALWAYS backed by the real
    doctype permissions — no matter how portal_role got set: the portal Users page, the
    Frappe Desk User form, a data import, or a patch.

    Without this, portal_role and Frappe roles drift: a user can be a portal Admin
    (sidebar + page guards pass) yet lack System Manager / Sales roles, so every real
    write (save a master, delete a scheme) fails Frappe's doctype permission check with
    'does not have doctype access via role permission'.

    Recursion-guarded because sync_frappe_roles grants/strips roles via
    User.add_roles / remove_roles, each of which saves the User and re-fires on_update."""
    user = getattr(doc, "name", None)
    if user in ("Administrator", "Guest", "", None):
        return
    if not doc.get("portal_role"):
        return
    if frappe.flags.get("in_sync_user_frappe_roles"):
        return
    # Normally act only when the portal role actually changed (or on first insert), so
    # unrelated profile saves (name, image, password, scheme emails, …) don't re-sync
    # every time. But ALSO self-heal: if a portal user is missing the Frappe roles their
    # role needs — e.g. an account created as a Website User before the sync existed, or
    # one whose desk roles were stripped — re-sync on any save so it converges. Because a
    # desk-access role flips user_type, the account then auto-becomes a System User; a
    # Website User silently drops these roles (and so can't write or download attachments).
    prune = True
    if method == "on_update" and not doc.has_value_changed("portal_role"):
        desired = set(PORTAL_FRAPPE_ROLES.get(doc.portal_role, []))
        if desired.issubset(set(frappe.get_roles(user))):
            return  # already correct — nothing to do
        prune = False  # self-heal is add-only; don't strip legacy roles on an unrelated save
    frappe.flags.in_sync_user_frappe_roles = True
    try:
        sync_frappe_roles(user, doc.portal_role, prune=prune)
    except Exception:
        # A role-sync problem must never block the User save itself (login-critical).
        frappe.log_error(frappe.get_traceback(), "sync_user_frappe_roles failed")
    finally:
        frappe.flags.in_sync_user_frappe_roles = False


def get_portal_role(user=None):
    """The user's single portal role. Administrator / any Frappe System Manager is Admin.
    Unmapped portal users default to the most restrictive role."""
    user = user or frappe.session.user
    if user == "Administrator":
        return ROLE_ADMIN
    roles = frappe.get_roles(user)
    if "System Manager" in roles:
        return ROLE_ADMIN
    return frappe.db.get_value("User", user, "portal_role") or ROLE_R


def is_portal_admin(user=None):
    return get_portal_role(user) == ROLE_ADMIN


def is_regional_role(user=None):
    """True for the two region-mapped field roles (Regional User and Regional User
    (Future)) — i.e. users who work within their own regions rather than across the
    whole division."""
    return get_portal_role(user) in (ROLE_R, ROLE_RF)


def scheme_list_label(user=None):
    """Heading for the scheme list page and its sidebar entry. Regional roles track the
    status of requests in their own regions rather than browsing every request, so the
    page reads as 'Scheme Status' for them. Exposed to Jinja for portal_base.html."""
    return "Scheme Status" if is_regional_role(user) else "Scheme Requests"


def has_app_access(user=None):
    """Apps-screen visibility for the Scanify Portal entry — any logged-in user."""
    return frappe.session.user not in (None, "", "Guest")


def can_access(process, user=None):
    return get_portal_role(user) in PROCESS_ROLES.get(process, {ROLE_ADMIN})


def is_manager(user=None):
    """Scheme approvers = portal Admin/HO (or a legacy Frappe Sales/System Manager)."""
    if get_portal_role(user) in (ROLE_ADMIN, ROLE_HO):
        return True
    roles = frappe.get_roles(user or frappe.session.user)
    return "System Manager" in roles or "Sales Manager" in roles


def require_manager(user=None):
    if not is_manager(user):
        frappe.throw("Only HO/Admin can approve or route scheme requests.", frappe.PermissionError)


def nav_access(user=None):
    """process -> bool map, used by the sidebar template to hide disallowed items."""
    role = get_portal_role(user)
    return frappe._dict({p: (role in roles) for p, roles in PROCESS_ROLES.items()})


def require(process, user=None):
    """Server-side guard for whitelisted endpoints. Raises 403 if not allowed."""
    if not can_access(process, user):
        frappe.throw("You are not permitted to perform this action.", frappe.PermissionError)


def require_process(process):
    """Decorator for whitelisted endpoints: enforce that the caller's portal role
    may use `process` (per PROCESS_ROLES) before the body runs. Apply it *below*
    @frappe.whitelist() so the registered/callable object carries the guard:

        @frappe.whitelist()
        @require_process("secondary_admin")
        def reload_stockist_statements(...): ...

    functools.wraps preserves the wrapped function's signature, so Frappe's argument
    dispatch (inspect.signature -> get_newargs) still binds form_dict kwargs correctly.
    This only adds ACCESS control (portal role); it does not replace the doctype-level
    permission checks the body may still perform."""
    def decorator(fn):
        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            require(process)
            return fn(*args, **kwargs)
        return wrapper
    return decorator


def _page_from_path(path):
    path = (path or "").strip("/")
    if not path.startswith("portal"):
        return None
    parts = path.split("/")
    return parts[1] if len(parts) > 1 else "portal"


# The Frappe Desk (framework UI) is served under these prefixes. /app and /apps
# redirect to /desk, and the desk SPA shell renders from /desk[/...].
_DESK_EXACT = {"/app", "/apps", "/desk"}

# Framework landing pages a portal-managed user should never land on: the bare site
# root (which resolves to the desk for System Managers, or Frappe's generic /me page
# otherwise) and the built-in /me profile. These aren't "desk paths", so they'd slip
# past _is_desk_path — send them to /portal too.
_HOME_EXACT = {"/", "/me"}


def _is_desk_path(path):
    if not path:
        return False
    p = (path.split("?", 1)[0]).rstrip("/") or "/"
    return p in _DESK_EXACT or p.startswith("/app/") or p.startswith("/desk/")


def _is_home_path(path):
    if not path:
        return False
    p = (path.split("?", 1)[0]).rstrip("/") or "/"
    return p in _HOME_EXACT


def _desk_blocked(user=None):
    """True if this account is a portal-managed user that must be kept out of the
    Frappe Desk. Administrator always keeps Desk access (escape hatch); so does any
    account with no portal_role set (clear the field to grant a user raw Desk access)."""
    user = user or frappe.session.user
    if user in ("Administrator", "Guest", "", None):
        return False
    return bool(frappe.db.get_value("User", user, "portal_role"))


def enforce_portal_access(context=None):
    """`update_website_context` hook — runs for every rendered web page.
    (1) Portal-managed users are redirected away from the Frappe Desk (/app, /desk)
        and the framework landing pages (site root, /me) to the portal — they get
        the portal, never the framework UI.
    (2) For /portal/* routes, users whose role can't access that page are sent back
        to the dashboard. Unmapped portal pages are left open to avoid lockouts."""
    try:
        request = getattr(frappe.local, "request", None)
        path = request.path if request else ""
    except Exception:
        return
    if frappe.session.user == "Guest":
        return  # login flow handles auth
    if _desk_blocked() and (_is_desk_path(path) or _is_home_path(path)):
        frappe.local.flags.redirect_location = "/portal"
        raise frappe.Redirect(http_status_code=302)
    page = _page_from_path(path)
    if page is None:
        return
    process = ROUTE_PROCESS.get(page)
    if process is None:
        return  # unmapped portal page → allow
    if not can_access(process):
        frappe.local.flags.redirect_location = "/portal?denied=1"
        raise frappe.Redirect(http_status_code=302)


# ─────────────────────────────────────────────────────────────────────────────
# Division / Region scoping (non-admins see only their mapped divisions & regions)
# ─────────────────────────────────────────────────────────────────────────────

def _split(raw):
    if not raw:
        return []
    return [x.strip() for x in str(raw).replace("\n", ",").split(",") if x.strip()]


def get_allowed_divisions(user=None):
    """List of divisions a user may use. Admin → all divisions."""
    user = user or frappe.session.user
    if is_portal_admin(user):
        return frappe.get_all("Division", pluck="name")
    divs = _split(frappe.db.get_value("User", user, "allowed_divisions"))
    if not divs:
        d = frappe.db.get_value("User", user, "division")
        if d:
            divs = [d]
    return divs


def get_allowed_region_codes(user=None, division=None):
    """Region codes a non-admin user may see (optionally filtered to a division).
    Returns None for admins (meaning: no restriction / all regions)."""
    user = user or frappe.session.user
    if is_portal_admin(user):
        return None
    codes = _split(frappe.db.get_value("User", user, "allowed_regions"))
    if division and codes:
        valid = set(frappe.get_all(
            "Region Master", filters={"division": ["in", [division, "Both"]]}, pluck="name"))
        codes = [c for c in codes if c in valid]
    return codes


def clamp_region_codes(requested_region=None, division=None, user=None):
    """Resolve which region codes a query must be limited to.
    - Admin: [requested_region] if one was chosen, else None (= all regions).
    - Non-admin: intersection of the request with the user's allowed regions.
      Returns [] when the user is allowed nothing (locked out of data)."""
    allowed = get_allowed_region_codes(user, division)
    if allowed is None:  # admin
        return [requested_region] if requested_region else None
    if requested_region:
        return [requested_region] if requested_region in allowed else []
    return allowed


@frappe.whitelist()
def get_user_scope(division=None):
    """Everything the front-end needs to lock itself down: role, nav map, allowed
    divisions, and the regions (with names) the user may pick — auto-select when one."""
    from scanify.api import get_user_division
    user = frappe.session.user
    role = get_portal_role(user)
    admin = role == ROLE_ADMIN
    division = division or get_user_division()

    if admin:
        regions = frappe.get_all(
            "Region Master",
            filters={"status": "Active", "division": ["in", [division, "Both"]]},
            fields=["name", "region_name"], order_by="region_name")
    else:
        codes = get_allowed_region_codes(user, division)
        regions = frappe.get_all(
            "Region Master", filters={"name": ["in", codes or [""]]},
            fields=["name", "region_name"], order_by="region_name") if codes else []

    return {
        "role": role,
        "is_admin": admin,
        "divisions": get_allowed_divisions(user),
        "active_division": division,
        "regions": regions,
        "region_codes": [r["name"] for r in regions],
        "nav": nav_access(user),
    }


# ─────────────────────────────────────────────────────────────────────────────
# Website document permissions (existing hooks)
# ─────────────────────────────────────────────────────────────────────────────

def has_scheme_permission(doc, user):
    """Check if user has permission to access scheme"""
    if frappe.session.user == "Administrator":
        return True
    if getattr(doc, "requested_by", None) == user:
        return True
    if is_portal_admin(user):
        return True
    from scanify.api import get_user_division
    return doc.division == get_user_division()


def has_statement_permission(doc, user):
    """Check if user has permission to access statement"""
    if frappe.session.user == "Administrator":
        return True
    from scanify.api import get_user_division
    return doc.division == get_user_division()

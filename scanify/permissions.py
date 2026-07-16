import frappe

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
    never strips the acting user's own System Manager (avoids self-lockout)."""
    if user == "Administrator":
        return
    portal_role = portal_role or get_portal_role(user)
    desired = [r for r in PORTAL_FRAPPE_ROLES.get(portal_role, []) if frappe.db.exists("Role", r)]
    have = set(frappe.get_roles(user))
    doc = frappe.get_doc("User", user)
    to_add = [r for r in desired if r not in have]
    if to_add:
        doc.add_roles(*to_add)
    if prune:
        to_remove = [r for r in (APP_MANAGED_ROLES - set(desired)) if r in have]
        if user == frappe.session.user:
            to_remove = [r for r in to_remove if r != "System Manager"]
        if to_remove:
            doc.remove_roles(*to_remove)


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


def _page_from_path(path):
    path = (path or "").strip("/")
    if not path.startswith("portal"):
        return None
    parts = path.split("/")
    return parts[1] if len(parts) > 1 else "portal"


# The Frappe Desk (framework UI) is served under these prefixes. /app and /apps
# redirect to /desk, and the desk SPA shell renders from /desk[/...].
_DESK_EXACT = {"/app", "/apps", "/desk"}


def _is_desk_path(path):
    if not path:
        return False
    p = (path.split("?", 1)[0]).rstrip("/") or "/"
    return p in _DESK_EXACT or p.startswith("/app/") or p.startswith("/desk/")


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
        to the portal — they get the portal, never the framework UI.
    (2) For /portal/* routes, users whose role can't access that page are sent back
        to the dashboard. Unmapped portal pages are left open to avoid lockouts."""
    try:
        request = getattr(frappe.local, "request", None)
        path = request.path if request else ""
    except Exception:
        return
    if frappe.session.user == "Guest":
        return  # login flow handles auth
    if _is_desk_path(path) and _desk_blocked():
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

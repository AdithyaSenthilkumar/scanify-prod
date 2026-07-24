import frappe

no_cache = 1


def get_context(context):
    """Server-side logout for /logout — overrides Frappe's own www/logout page.

    Frappe's version renders an HTML shell whose only job is to JS-post to the logout
    endpoint; on this site that render lands the user on a 404 instead. Since the
    portal is the whole UI here, there is no reason for a logout *page* at all: end
    the session and bounce to /login. Nothing is ever rendered — logout.html exists
    only so the template resolver finds this route (scanify's www folder is searched
    before frappe's, so this wins).
    """
    try:
        if frappe.session.user != "Guest":
            frappe.local.login_manager.logout()
            frappe.db.commit()
    except Exception:
        # A session that can't be torn down is already unusable — still send the
        # user to /login rather than surfacing a framework error page.
        frappe.log_error(frappe.get_traceback(), "Portal Logout Error")

    # 302, not the default 301: a permanently-cached redirect on /logout would be
    # baked into the browser forever.
    frappe.local.flags.redirect_location = "/login"
    raise frappe.Redirect(http_status_code=302)

import frappe


def execute():
    """Make post-login land on /portal for everyone by pointing System Settings
    default_app at the scanify app (whose add_to_apps_screen route is /portal).
    get_default_path() then returns /portal for both System and Website users."""
    try:
        frappe.db.set_single_value("System Settings", "default_app", "scanify")
        frappe.db.commit()
        print("✓ System Settings default_app = scanify (post-login -> /portal)")
    except Exception as e:
        # default_app is a Select bound to installed apps; if not settable, skip quietly.
        print(f"• Could not set default_app: {e}")

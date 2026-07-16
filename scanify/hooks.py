app_name = "scanify"
app_title = "Scanify"
app_publisher = "Stedman Pharmaceuticals"
app_description = "Stockist Entry and Scheme Management System"
app_email = "admin@stedman.com"
app_license = "mit"

# Include CSS and JS
app_include_js = "/assets/scanify/js/scanify.js"
app_include_css = "/assets/scanify/css/scanify.css"
#web_include_js = [
#    "scanify/public/js/login_redirect.js"
#]
# Workspace and branding
website_context = {
    "brand_html": "<img src='/assets/scanify/images/stedman_logo.png' style='max-height: 40px;' />"
}
# Portal settings
has_website_permission = {
    "Scheme Request": "scanify.permissions.has_scheme_permission",
    "Stockist Statement": "scanify.permissions.has_statement_permission"
}

# Central role-based page lock: runs for every rendered web page and redirects
# users away from portal pages their role can't access.
update_website_context = [
    "scanify.permissions.enforce_portal_access"
]

# Expose access helpers to Jinja so portal_base.html can hide disallowed sidebar items.
jinja = {
    "methods": [
        "scanify.permissions.nav_access",
        "scanify.permissions.get_portal_role",
        "scanify.permissions.get_allowed_divisions",
    ]
}

# Web form list
web_form_list_context = {
    "Scheme Request": "scanify.portal.get_scheme_list_context"
}


# Redirect after login

# Document hooks
doc_events = {
    "Stockist Statement": {
        "validate": "scanify.scanify.doctype.stockist_statement.stockist_statement.validate_closing_balance",
    },
    "Scheme Request": {
        "on_submit": "scanify.scanify.doctype.scheme_request.scheme_request.create_stock_adjustment"
    }
}

# Register the portal as an "app" so post-login lands on /portal for BOTH System and
# Website users (via default_app -> get_default_path). Set System Settings default_app
# = "scanify" (done by patch set_default_app_to_portal).
add_to_apps_screen = [
    {
        "name": "scanify",
        "logo": "/assets/scanify/images/scanify_logo.jpg",
        "title": "Scanify Portal",
        "route": "/portal",
        "has_permission": "scanify.permissions.has_app_access",
    }
]

# Boot session
boot_session = "scanify.boot.boot_session"

fixtures = [
]


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

# Boot session
boot_session = "scanify.boot.boot_session"

fixtures = [
]


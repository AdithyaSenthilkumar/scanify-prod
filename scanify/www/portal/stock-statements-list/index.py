import frappe
from scanify.api import get_user_division, get_stockist_report_filter_options


def get_context(context):
    if frappe.session.user == "Guest":
        frappe.throw("You must be logged in to view stock statements", frappe.PermissionError)

    context.no_cache = 1

    # Priority 1: Check session first (for current page load)
    user_division = None
    if hasattr(frappe.session, "user_division") and frappe.session.user_division:
        user_division = frappe.session.user_division

    # Priority 2: Check User document
    if not user_division:
        user_division = frappe.db.get_value("User", frappe.session.user, "division")

    if not user_division:
        user_division = "Prima"

    context.division = user_division

    # Load filter options (zones, regions, teams, hqs) for the division
    opts = get_stockist_report_filter_options(user_division)
    context.zones = opts.get("zones", [])
    context.regions = opts.get("regions", [])
    context.teams = opts.get("teams", [])
    context.hqs = opts.get("hqs", [])

    # Stockist Statement carries its own `division` field, kept in sync with the stockist on
    # every save (set_division_from_stockist), so we filter on it directly. This is more robust
    # than the old approach of re-deriving the stockist set with a status="Active" filter, which
    # silently hid statements whose stockist had since been deactivated.
    #
    # We also load ALL matching statements (no page cap). Filtering/search on this page happens
    # client-side over the rows rendered into the DOM, so any statement not fetched here is
    # invisible to every filter AND the name search. The previous limit_page_length=300 meant
    # that once a division exceeded 300 statements, the older ones (e.g. a freshly QC-reviewed
    # batch for a single region) simply could not be found by any means.
    statements = frappe.get_all(
        "Stockist Statement",
        filters={"docstatus": ["in", [0, 1]], "division": ["in", [user_division, "Both"]]},
        fields=["name", "stockist_code", "statement_month", "extracted_data_status", "docstatus",
                "creation", "qc_confidence", "confidence_score", "hq", "team", "region", "zone"],
        order_by="creation desc",
        limit_page_length=0,
    )

    # Enrich with stockist names + HQ display names
    if statements:
        found_stockist_codes = list({s.stockist_code for s in statements})
        stockist_names = {
            row.name: row.stockist_name
            for row in frappe.get_all("Stockist Master", {"name": ["in", found_stockist_codes]}, ["name", "stockist_name"])
        }
        hq_codes = list({s.hq for s in statements if s.hq})
        hq_names = {}
        if hq_codes:
            hq_names = {
                row.name: row.hq_name
                for row in frappe.get_all("HQ Master", {"name": ["in", hq_codes]}, ["name", "hq_name"])
            }
        for stmt in statements:
            stmt.stockist_name = stockist_names.get(stmt.stockist_code, stmt.stockist_code)
            stmt.hq_name = hq_names.get(stmt.hq, stmt.hq or "")

    context.statements = statements

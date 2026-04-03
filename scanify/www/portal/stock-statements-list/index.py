import frappe

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
    
    # We need to filter statements by the user's division. Stockist Statement doesn't have division directly.
    # It has stockist_code. Let's fetch the stockists belonging to the user's division or "Both".
    stockists = frappe.get_all("Stockist Master", {"division": ["in", [user_division, "Both"]], "status": "Active"}, pluck="name")
    
    if not stockists:
        context.statements = []
        return
        
    # Get all statements for these stockists (draft + submitted)
    statements = frappe.get_all(
        "Stockist Statement",
        filters={"docstatus": ["in", [0, 1]], "stockist_code": ["in", stockists]},
        fields=["name", "stockist_code", "statement_month", "extracted_data_status", "docstatus", "creation", "qc_confidence", "confidence_score"],
        order_by="creation desc",
        limit_page_length=300
    )
    
    # Enrich with stockist names
    if statements:
        # We fetch the exact names again just in case some stockists weren't fetched in the pluck above (if status active had changed, though unlikely)
        found_stockist_codes = list({s.stockist_code for s in statements})
        stockist_names = {
            row.name: row.stockist_name 
            for row in frappe.get_all("Stockist Master", {"name": ["in", found_stockist_codes]}, ["name", "stockist_name"])
        }
        for stmt in statements:
            stmt.stockist_name = stockist_names.get(stmt.stockist_code, stmt.stockist_code)
            
    context.statements = statements

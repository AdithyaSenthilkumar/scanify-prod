import frappe
from frappe.model.document import Document
from frappe.utils import flt

class SchemeDeduction(Document):
    def validate(self):
        """Validate before saving"""
        self.validate_scheme_and_statement()
        self.validate_stockist_match()
        self.validate_products_exist_in_statement()
        self.calculate_totals()
    
    def validate_scheme_and_statement(self):
        """Ensure scheme and statement are submitted"""
        if self.scheme_request:
            scheme = frappe.get_doc("Scheme Request", self.scheme_request)
            if scheme.docstatus != 1:
                frappe.throw(f"Scheme Request {self.scheme_request} must be submitted first")
        
        if self.stockist_statement:
            statement = frappe.get_doc("Stockist Statement", self.stockist_statement)
            if statement.docstatus == 2:
                frappe.throw(f"Cannot deduct from cancelled Stockist Statement {self.stockist_statement}")
    
    def validate_stockist_match(self):
        """Ensure scheme and statement belong to same stockist"""
        if self.scheme_request and self.stockist_statement:
            scheme = frappe.get_doc("Scheme Request", self.scheme_request)
            statement = frappe.get_doc("Stockist Statement", self.stockist_statement)
            
            if scheme.stockist_code != statement.stockist_code:
                frappe.throw(
                    f"Stockist mismatch: Scheme ({scheme.stockist_code}) != Statement ({statement.stockist_code})"
                )
    
    def validate_products_exist_in_statement(self):
        """Validate that all products exist in the stockist statement"""
        if not self.stockist_statement or not self.items:
            return
        
        statement = frappe.get_doc("Stockist Statement", self.stockist_statement)
        
        # Build a set of product codes in the statement
        statement_products = {item.product_code for item in statement.items}
        
        # Check each deduction item
        missing_products = []
        for item in self.items:
            if item.product_code not in statement_products:
                missing_products.append(item.product_code)
        
        if missing_products:
            frappe.throw(
                f"Cannot apply deduction. The following products are not in the stockist statement: {', '.join(missing_products)}<br><br>"
                f"Please remove these products from the deduction or add them to the statement first."
            )
    
    def calculate_totals(self):
        """Calculate total deducted qty and value"""
        self.total_deducted_qty = 0
        self.total_deducted_value = 0
        
        for item in self.items:
            self.total_deducted_qty += flt(item.deduct_qty)
            item.deducted_value = flt(item.deduct_qty) * flt(item.pts)
            self.total_deducted_value += item.deducted_value
    
    def on_submit(self):
        """Apply deduction to stockist statement"""
        self.apply_deduction()
        self.status = "Applied"
        self.db_update()
    
    def on_cancel(self):
        """Reverse deduction from stockist statement"""
        self.reverse_deduction()
        self.status = "Cancelled"
        self.db_update()
    
    def apply_deduction(self):
        """
        Deduct free qty from stockist statement
        CORRECTED LOGIC:
        - Deduct from sales_qty (not free_qty)
        - Populate free_qty_scheme field
        """
        statement = frappe.get_doc("Stockist Statement", self.stockist_statement)
        
        products_updated = []
        
        for deduction_item in self.items:
            # Find matching product in statement
            for stmt_item in statement.items:
                if stmt_item.product_code == deduction_item.product_code:
                    # CORRECT DEDUCTION LOGIC:
                    # 1. Add to free_qty_scheme (this represents scheme free goods)
                    stmt_item.free_qty_scheme = flt(stmt_item.free_qty_scheme) + flt(deduction_item.deduct_qty)
                    
                    
                    products_updated.append(deduction_item.product_code)
                    break
        
        # Recalculate closing balances with new logic
        statement.calculate_closing_and_totals()
        statement.save(ignore_permissions=True)
        
        frappe.msgprint(
            f"Deduction applied to {self.stockist_statement}<br>"
            f"Updated {len(products_updated)} products: <br>"
            + "<br>".join(products_updated) + "<br>"
            f"Closing balances recalculated.",
            alert=True,
            indicator="green"
        )
    
    def reverse_deduction(self):
        """Reverse the deduction on cancel"""
        statement = frappe.get_doc("Stockist Statement", self.stockist_statement)
        
        for deduction_item in self.items:
            for stmt_item in statement.items:
                if stmt_item.product_code == deduction_item.product_code:
                    # REVERSE THE DEDUCTION:
                    # 1. Subtract from free_qty_scheme
                    stmt_item.free_qty_scheme = flt(stmt_item.free_qty_scheme) - flt(deduction_item.deduct_qty)
                    
                    # 2. Add back to sales_qty
                    stmt_item.sales_qty = flt(stmt_item.sales_qty) + flt(deduction_item.deduct_qty)
                    break
        
        # Recalculate closing balances
        statement.calculate_closing_and_totals()
        statement.save(ignore_permissions=True)
        
        frappe.msgprint(
            f"Deduction reversed from {self.stockist_statement}",
            alert=True,
            indicator="orange"
        )

@frappe.whitelist()
@frappe.validate_and_sanitize_search_inputs
def get_scheme_requests(doctype, txt, searchfield, start, page_len, filters):
    """Custom search query for Scheme Request with date and doctor name"""
    return frappe.db.sql("""
        SELECT 
            sr.name,
            CONCAT(
                sr.name, 
                ' | ', 
                DATE_FORMAT(sr.application_date, '%%d-%%b-%%Y'),
                ' | Dr. ',
                COALESCE(dm.doctor_name, sr.doctor_code)
            ) as label
        FROM 
            `tabScheme Request` sr
        LEFT JOIN 
            `tabDoctor Master` dm ON sr.doctor_code = dm.name
        WHERE 
            sr.docstatus = 1
            AND (
                sr.name LIKE %(txt)s
                OR sr.doctor_code LIKE %(txt)s
                OR dm.doctor_name LIKE %(txt)s
            )
        ORDER BY 
            sr.application_date DESC
        LIMIT %(start)s, %(page_len)s
    """, {
        'txt': f"%{txt}%",
        'start': start,
        'page_len': page_len
    })

@frappe.whitelist()
@frappe.validate_and_sanitize_search_inputs
def get_stockist_statements(doctype, txt, searchfield, start, page_len, filters):
    """Custom search query for Stockist Statement with month"""
    stockist_code = filters.get('stockist_code')
    
    conditions = []
    if stockist_code:
        conditions.append(f"ss.stockist_code = '{stockist_code}'")
    if txt:
        conditions.append(f"(ss.name LIKE '%{txt}%' OR ss.statement_month LIKE '%{txt}%')")
    
    where_clause = " AND ".join(conditions) if conditions else "1=1"
    
    return frappe.db.sql(f"""
        SELECT 
            ss.name,
            CONCAT(
                ss.name,
                ' | ',
                DATE_FORMAT(ss.statement_month, '%%b-%%Y'),
                ' | ',
                COALESCE(sm.stockist_name, ss.stockist_code)
            ) as label
        FROM 
            `tabStockist Statement` ss
        LEFT JOIN 
            `tabStockist Master` sm ON ss.stockist_code = sm.name
        WHERE 
            ss.docstatus != 2
            AND {where_clause}
        ORDER BY 
            ss.statement_month DESC
        LIMIT {start}, {page_len}
    """)

@frappe.whitelist()
def fetch_and_populate_items(scheme_request, stockist_statement):
    """
    Fetch items from scheme and match with stockist statement
    Returns items with both scheme qty and current statement qty
    ONLY returns products that exist in BOTH scheme and statement
    """
    if not scheme_request or not stockist_statement:
        return []
    
    # Get scheme items
    scheme = frappe.get_doc("Scheme Request", scheme_request)
    
    # Get statement items
    statement = frappe.get_doc("Stockist Statement", stockist_statement)
    
    # Build a map of current free qty in statement
    statement_items_map = {}
    for stmt_item in statement.items:
        statement_items_map[stmt_item.product_code] = {
            "free_qty": flt(stmt_item.free_qty),
            "exists": True
        }
    
    # Build result items - ONLY for products that exist in statement
    items = []
    skipped_products = []
    
    for scheme_item in scheme.items:
        # Check if product exists in statement
        if scheme_item.product_code not in statement_items_map:
            skipped_products.append(scheme_item.product_code)
            continue
        
        product = frappe.get_doc("Product Master", scheme_item.product_code)
        
        current_free_qty = statement_items_map[scheme_item.product_code]["free_qty"]
        scheme_free_qty = flt(scheme_item.free_quantity)
        
        items.append({
            "product_code": scheme_item.product_code,
            "product_name": product.product_name,
            "pack": product.pack,
            "scheme_free_qty": scheme_free_qty,
            "current_free_qty": current_free_qty,
            "deduct_qty": scheme_free_qty,  # Default to full scheme qty
            "pts": flt(product.pts),
            "deducted_value": scheme_free_qty * flt(product.pts)
        })
    
    # Show warning if some products were skipped
    if skipped_products:
        frappe.msgprint(
            f"⚠️ Warning: {len(skipped_products)} product(s) from the scheme are not in this stockist statement and were skipped:<br>"
            f"<b>{', '.join(skipped_products)}</b><br><br>"
            f"Only products present in the statement have been loaded.",
            title="Some Products Skipped",
            indicator="orange"
        )
    
    return items

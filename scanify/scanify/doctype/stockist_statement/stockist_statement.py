import frappe
from frappe.model.document import Document
from frappe.utils import flt, add_months, get_first_day, get_last_day

class StockistStatement(Document):
    def validate(self):
        """Validate and calculate closing balance"""
        self.calculate_closing_and_totals()
    
    def calculate_closing_and_totals(self):
        """Calculate closing quantities, values and totals including free qty"""
        total_opening = 0
        total_purchase = 0
        total_sales = 0
        total_free = 0
        total_closing = 0
        
        for item in self.items:
            if not item.product_code:
                continue
            
            # Closing = Opening + Purchase - Sales - Free - Return - Misc Out
            item.closing_qty = (
                flt(item.opening_qty) +
                flt(item.purchase_qty) -
                flt(item.sales_qty) -
                flt(item.free_qty) -
                flt(item.return_qty) -
                flt(item.misc_out_qty)
            )
            
            # Calculate values
            pts = flt(item.pts)
            item.closing_value = flt(item.closing_qty) * pts
            
            # Add to totals
            total_opening += flt(item.opening_qty) * pts
            total_purchase += flt(item.purchase_qty) * pts
            total_sales += flt(item.sales_qty) * pts
            total_free += flt(item.free_qty) * pts
            total_closing += flt(item.closing_value)
        
        self.total_opening_value = total_opening
        self.total_purchase_value = total_purchase
        self.total_sales_value = total_sales
        self.total_free_value = total_free
        self.total_closing_value = total_closing

def validate_closing_balance(doc, method):
    """Hook to validate closing balance"""
    doc.calculate_closing_and_totals()

def update_next_month_opening(doc, method):
    """Update next month's opening balance after submission"""
    try:
        # Calculate next month
        next_month = add_months(doc.statement_month, 1)
        next_month_first = get_first_day(next_month)
        
        # Check if next month's statement exists
        next_statement = frappe.db.exists("Stockist Statement", {
            "stockist_code": doc.stockist_code,
            "statement_month": next_month_first,
            "docstatus": 0
        })
        
        if next_statement:
            next_doc = frappe.get_doc("Stockist Statement", next_statement)
            
            # Update opening quantities
            for item in doc.items:
                if not item.product_code:
                    continue
                for next_item in next_doc.items:
                    if next_item.product_code == item.product_code:
                        next_item.opening_qty = flt(item.closing_qty) or 0
                        break
            
            next_doc.calculate_closing_and_totals()
            next_doc.save()
            frappe.msgprint(f"Next month's opening balance updated for {next_doc.name}")
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Update Next Month Opening Error")

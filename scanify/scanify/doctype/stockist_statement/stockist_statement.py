import frappe
import re
from frappe.model.document import Document
from frappe.utils import flt, add_months, get_first_day, get_last_day

class StockistStatement(Document):
    def validate(self):
        self.set_division_from_stockist()
        self.calculate_closing_and_totals()   # your existing method

    def set_division_from_stockist(self):
        if not self.stockist_code:
            return

        division = frappe.db.get_value("Stockist Master", self.stockist_code, "division")
        if not division:
            division = frappe.db.get_value("HQ Master", self.hq, "division")
        if not division:
            frappe.throw(f"Stockist {self.stockist_code} has no Division set")

        self.division = division
    def _get_approved_scheme_qty_map(self):
        """Return {productcode: approved_free_qty} for THIS statement month only (statement units)."""
        approved_map = {}

        if not self.stockist_code or not self.statement_month:
            return approved_map

        start_date = get_first_day(self.statement_month)
        end_date   = get_last_day(self.statement_month)

        rows = frappe.db.sql("""
            SELECT
                sri.product_code AS product_code,
                SUM(COALESCE(sri.free_quantity, 0)) AS approved_free_qty
            FROM `tabScheme Request` sr
            INNER JOIN `tabScheme Request Item` sri ON sri.parent = sr.name
            WHERE sr.stockist_code = %s
            AND sr.docstatus = 1
            AND sr.application_date BETWEEN %s AND %s
            GROUP BY sri.product_code
        """, (self.stockist_code, start_date, end_date), as_dict=True)
        for r in rows:
            approved_map[r.product_code] = flt(r.approved_free_qty)

        return approved_map
    
    def calculate_closing_and_totals(self):
        """Calculate closing qty and value totals with pack-to-strip conversion"""

        total_sales_qty = 0
        total_sales_value_pts = 0
        total_sales_value_ptr = 0
        total_opening_value = 0
        total_purchase_value = 0
        total_closing_value = 0
        approved_scheme_map = self._get_approved_scheme_qty_map()


        for item in self.items:
            if not item.product_code:
                continue

            # -------- FETCH PRODUCT MASTER --------
            product = frappe.db.get_value(
                "Product Master",
                item.product_code,
                ["pts", "ptr", "pack"],
                as_dict=True
            )
            if not product:
                continue

            pts = flt(product.pts or 0)
            ptr = flt(product.ptr or 0)

            # -------- UNIT CONVERSION (BOX ➜ STRIP) --------
            conversion_factor = flt(self.get_conversion_factor(product.pack)) or 1
            item.conversion_factor = conversion_factor 

            opening_qty_base = flt(item.opening_qty) / conversion_factor
            purchase_qty_base = flt(item.purchase_qty) / conversion_factor
            sales_qty_base = flt(item.sales_qty) / conversion_factor
            free_qty_base = flt(item.free_qty) / conversion_factor
            scheme_free_qty_base = flt(item.free_qty_scheme) / conversion_factor
            return_qty_base = flt(item.return_qty) / conversion_factor
            misc_out_qty_base = flt(item.misc_out_qty) / conversion_factor
            closing_qty_base = flt(item.closing_qty) / conversion_factor

            if item.closing_qty is None or item.closing_qty == 0:
                closing_for_calc = (
                    opening_qty_base + purchase_qty_base 
                    - sales_qty_base - free_qty_base - scheme_free_qty_base
                    - return_qty_base - misc_out_qty_base
                )
                # Store UNCONVERTED closing (multiply back)
                item.closing_qty = closing_for_calc * conversion_factor
            else:
                # Use stockist's reported closing
                closing_for_calc = flt(item.closing_qty) / conversion_factor
            
            approved_scheme_qty = flt(approved_scheme_map.get(item.product_code, 0))
            item.scheme_deducted_qty_calc = flt(item.sales_qty) + flt(item.free_qty) - approved_scheme_qty 
            
            scheme_deducted_qty_base = flt(item.scheme_deducted_qty_calc) / conversion_factor


            # -------- VALUE CALCULATIONS (STRIP LEVEL) --------
            item.opening_value = opening_qty_base * pts
            item.purchase_value = purchase_qty_base * pts
            item.sales_value_pts = scheme_deducted_qty_base  * pts
            item.sales_value_ptr = scheme_deducted_qty_base  * ptr
            item.closing_value = closing_qty_base * pts

            # -------- TOTALS --------
            total_sales_qty += flt(item.sales_qty)
            total_sales_value_pts += item.sales_value_pts
            total_sales_value_ptr += item.sales_value_ptr
            total_opening_value += item.opening_value
            total_purchase_value += item.purchase_value
            total_closing_value += item.closing_value

        # -------- DOCUMENT TOTALS --------
        self.total_sales_qty = total_sales_qty
        self.total_sales_value_pts = total_sales_value_pts
        self.total_sales_value_ptr = total_sales_value_ptr
        self.total_opening_value = total_opening_value
        self.total_purchase_value = total_purchase_value
        self.total_closing_value = total_closing_value

    
    def get_conversion_factor(self, pack_str):
        """
        Extract conversion factor from pack field
        Examples:
        - "10x6" -> 10
        - "1x10" -> 1
        - "10's" -> 1
        - "Unit" -> 1
        - "10ml" -> 1
        - "10gms" -> 1
        
        Returns: conversion factor (denominator for division)
        """
        if not pack_str:
            return 1
        
        pack_str = str(pack_str).strip().upper()
        
        # Pattern 1: "AxB" format (e.g., "10x6", "1x10")
        match = re.match(r'(\d+)\s*[xX]\s*(\d+)', pack_str)
        if match:
            return flt(match.group(1))  # Return the first number (before 'x')
        
        # Pattern 2: Check for unit/box indicators
        if any(indicator in pack_str for indicator in ['UNIT', 'BOX', 'ML', 'GM', 'MG', "'S"]):
            return 1
        
        # Default: no conversion
        return 1

def validate_closing_balance(doc, method):
    """Hook to validate closing balance"""
    doc.calculate_closing_and_totals()

def update_next_month_opening(doc, method):
    """Update next month's opening balance after submission"""
    try:
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
            
            # Update opening quantities from current closing
            for item in doc.items:
                if not item.product_code:
                    continue
                
                for next_item in next_doc.items:
                    if next_item.product_code == item.product_code:
                        next_item.opening_qty = flt(item.closing_qty or 0)
                        break
            
            next_doc.calculate_closing_and_totals()
            next_doc.save()
            frappe.msgprint(f"Next month's opening balance updated for {next_doc.name}")
    
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Update Next Month Opening Error")

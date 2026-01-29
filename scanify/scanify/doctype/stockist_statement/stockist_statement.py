import frappe
import re
from frappe.model.document import Document
from frappe.utils import flt, add_months, get_first_day, get_last_day

class StockistStatement(Document):
    def validate(self):
        """Validate and calculate closing balance"""
        self.calculate_closing_and_totals()
    
    def calculate_closing_and_totals(self):
        """
        Calculate values using pack conversion internally
        BUT keep original quantities in the table for display
        """
        total_opening = 0
        total_purchase = 0
        total_sales = 0
        total_free = 0
        total_free_scheme = 0
        total_closing = 0
        
        for item in self.items:
            if not item.product_code:
                continue
            
            # Get product details
            product = frappe.db.get_value(
                "Product Master",
                item.product_code,
                ["pts", "pack"],
                as_dict=True
            )
            
            if not product:
                continue
            
            # Get conversion factor (store for reference, but don't modify quantities)
            item.conversion_factor = self.get_conversion_factor(product.pack)
            conversion_factor = flt(item.conversion_factor or 1)
            
            # Use ORIGINAL quantities from stockist statement (don't overwrite them)
            # Apply conversion ONLY for calculation
            opening_for_calc = flt(item.opening_qty) / conversion_factor
            purchase_for_calc = flt(item.purchase_qty) / conversion_factor
            sales_for_calc = flt(item.sales_qty) / conversion_factor
            free_for_calc = flt(item.free_qty) / conversion_factor
            free_scheme_for_calc = flt(item.free_qty_scheme) / conversion_factor
            return_for_calc = flt(item.return_qty) / conversion_factor
            misc_out_for_calc = flt(item.misc_out_qty) / conversion_factor
            
            # Calculate closing (if not provided by stockist)
            # If closing_qty already extracted from OCR, use it; otherwise calculate
            if item.closing_qty is None or item.closing_qty == 0:
                closing_for_calc = (
                    opening_for_calc + purchase_for_calc 
                    - sales_for_calc - free_for_calc - free_scheme_for_calc
                    - return_for_calc - misc_out_for_calc
                )
                # Store UNCONVERTED closing (multiply back)
                item.closing_qty = closing_for_calc * conversion_factor
            else:
                # Use stockist's reported closing
                closing_for_calc = flt(item.closing_qty) / conversion_factor
            
            # Get PTS from Product Master
            pts = flt(product.pts)
            
            # Calculate VALUES using converted quantities
            # Store values, not converted quantities
            opening_value = opening_for_calc * pts
            purchase_value = purchase_for_calc * pts
            sales_value = sales_for_calc * pts
            free_value = free_for_calc * pts
            free_scheme_value = free_scheme_for_calc * pts
            closing_value = closing_for_calc * pts
            
            # Store ONLY values in hidden fields (optional, for detailed tracking)
            item.opening_value = opening_value
            item.purchase_value = purchase_value
            item.sales_value = sales_value
            item.free_value = free_value
            item.free_scheme_value = free_scheme_value
            item.closing_value = closing_value
            
            # Accumulate totals
            total_opening += opening_value
            total_purchase += purchase_value
            total_sales += sales_value
            total_free += free_value
            total_free_scheme += free_scheme_value
            total_closing += closing_value
        
        # Set document totals (values only)
        self.total_opening_value = total_opening
        self.total_purchase_value = total_purchase
        self.total_sales_value = total_sales
        self.total_free_value = total_free
        self.total_free_scheme_value = total_free_scheme
        self.total_closing_value = total_closing

    
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

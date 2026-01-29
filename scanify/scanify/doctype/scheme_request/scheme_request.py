# scanify/scanify/doctype/scheme_request/scheme_request.py

import frappe
from frappe.model.document import Document
from frappe.utils import flt, nowdate, getdate, get_first_day, get_last_day
from datetime import datetime

class SchemeRequest(Document):
    def validate(self):
        self.calculate_total_scheme_value()
        # REMOVED: self.validate_attachments() - No longer mandatory
        self.validate_monthly_doctor_limit()  # NEW
    
    def calculate_total_scheme_value(self):
        total = 0
        if not self.items:
            self.total_scheme_value = 0
            return
        
        for item in self.items:
            rate = flt(item.special_rate or 0) if item.special_rate else flt(item.product_rate or 0)
            quantity = flt(item.quantity or 0)
            item.product_value = quantity * rate
            total += item.product_value
        
        self.total_scheme_value = total
    
    # REMOVED: validate_attachments method - attachments are now optional
    
    def validate_monthly_doctor_limit(self):
        """Validate that a doctor can have maximum 3 requests per product per month"""
        if not self.doctor_code or not self.application_date or not self.items:
            return
        
        app_date = getdate(self.application_date)
        first_day = get_first_day(app_date)
        last_day = get_last_day(app_date)
        month_name = app_date.strftime("%B %Y")

        # Get all submitted scheme requests for this doctor in current month
        existing_requests = frappe.db.sql("""
            SELECT sri.product_code, COUNT(DISTINCT sr.name) as request_count
            FROM `tabScheme Request` sr
            INNER JOIN `tabScheme Request Item` sri ON sr.name = sri.parent
            WHERE sr.doctor_code = %(doctor_code)s
            AND sr.application_date BETWEEN %(first_day)s AND %(last_day)s
            AND sr.docstatus != 2
            {exclude_current}
            GROUP BY sri.product_code
        """.format(
            exclude_current="AND sr.name != %(current_name)s" if not self.is_new() else ""
        ), {
            "doctor_code": self.doctor_code,
            "first_day": first_day,
            "last_day": last_day,
            "current_name": self.name if not self.is_new() else ""
        }, as_dict=True)
        
        # Build a dict of product_code -> count
        product_request_counts = {row.product_code: row.request_count for row in existing_requests}
        
        # Check each product in current request
        violations = []
        for item in self.items:
            if not item.product_code:
                continue
            
            current_count = product_request_counts.get(item.product_code, 0)
            
            # If this product already has 3+ requests this month, reject
            if current_count >= 3:
                product_name = frappe.db.get_value("Product Master", item.product_code, "product_name") or item.product_code
                violations.append(
                    f"• <b>{product_name} ({item.product_code})</b>: Already has {current_count} requests this month"
                )
        
        if violations:
            frappe.throw(
                f"<b>Product-wise Monthly Limit Exceeded for Dr. {self.doctor_name} ({self.doctor_code})</b><br><br>"
                f"The following products already have 3 or more requests in <b>{month_name}</b>:<br><br>"
                + "<br>".join(violations) +
                f"<br><br><i>Maximum 3 requests per product per doctor per month allowed.</i>",
                title="Monthly Product Limit Exceeded"
            )
    
    
    def on_submit(self):
        if self.approval_status == "Approved":
            self.create_stock_adjustment()
        else:
            frappe.throw("Cannot submit scheme request without approval")
    
    def create_stock_adjustment(self):
        """Create stock adjustment after scheme approval"""
        try:
            self.append("approval_log", {
                "approver": frappe.session.user,
                "approval_level": "Final",
                "action": "Approved",
                "action_date": nowdate(),
                "comments": "Scheme approved and submitted"
            })
            self.save()
            frappe.msgprint(
                f"Scheme request approved. Total value: ₹{flt(self.total_scheme_value or 0):.2f}",
                indicator="green"
            )
        except Exception as e:
            frappe.log_error(frappe.get_traceback(), "Create Stock Adjustment Error")

    def before_save(self):
        """Calculate scheme percentages before saving"""
        for item in self.items:
            scheme_pct = 0
            
            # If special rate is provided, calculate based on rate difference
            if item.special_rate and item.product_rate:
                discount = flt(item.product_rate) - flt(item.special_rate)
                if item.product_rate > 0:
                    scheme_pct = (discount / flt(item.product_rate)) * 100
            # Otherwise calculate based on free quantity
            elif item.free_quantity and item.quantity:
                scheme_pct = (flt(item.free_quantity) / flt(item.quantity)) * 100
            
            item.scheme_percentage = scheme_pct
            
            # Recalculate product value
            rate = item.special_rate if item.special_rate else item.product_rate
            item.product_value = flt(item.quantity) * flt(rate)


@frappe.whitelist()
def create_stock_adjustment(doc, method=None):
    """Create stock adjustment entry when scheme is approved"""
    # This is called via hooks on_submit
    # Add your stock adjustment logic here if needed
    # For now, just pass to avoid the error
    pass
# NEW: Repeat Request Functionality
@frappe.whitelist()
def repeat_scheme_request(source_name):
    """
    Create a new scheme request by duplicating an approved request
    
    Args:
        source_name: Name of the source Scheme Request document
    
    Returns:
        dict: Success status and new document name
    """
    try:
        # Get source document
        source_doc = frappe.get_doc("Scheme Request", source_name)
        
        # Validate that source is approved
        if source_doc.approval_status != "Approved":
            frappe.throw(
                "Only approved scheme requests can be repeated",
                title="Invalid Status"
            )
        
        
        # Create new document
        new_doc = frappe.new_doc("Scheme Request")
        
        # Copy header fields
        new_doc.application_date = nowdate()
        new_doc.requested_by = frappe.session.user
        new_doc.team = source_doc.team
        new_doc.region = source_doc.region
        new_doc.hq = source_doc.hq
        new_doc.stockist_code = source_doc.stockist_code
        new_doc.stockist_name = source_doc.stockist_name
        new_doc.doctor_code = source_doc.doctor_code
        new_doc.doctor_name = source_doc.doctor_name
        new_doc.doctor_place = source_doc.doctor_place
        new_doc.specialization = source_doc.specialization
        new_doc.hospital_clinic = source_doc.hospital_clinic
        new_doc.scheme_notes = f"Repeated from {source_doc.name}"
        
        # Set default status
        new_doc.approval_status = "Pending"
        
        # Copy items
        for item in source_doc.items:
            new_doc.append("items", {
                "product_code": item.product_code,
                "product_name": item.product_name,
                "pack": item.pack,
                "quantity": item.quantity,
                "free_quantity": item.free_quantity,
                "product_rate": item.product_rate,
                "special_rate": item.special_rate,
                "product_value": item.product_value
            })
        
        # Save the new document
        new_doc.insert()
        frappe.db.commit()
        
        return {
            "success": True,
            "message": f"New scheme request created successfully",
            "doc_name": new_doc.name,
            "doc_url": f"/app/scheme-request/{new_doc.name}"
        }
        
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Repeat Scheme Request Error")
        frappe.throw(str(e))


# NEW: Get monthly request count for doctor
@frappe.whitelist()
def get_doctor_monthly_count(doctor_code, application_date=None):
    """
    Get the count of scheme requests for a doctor in a given month
    
    Args:
        doctor_code: Doctor code
        application_date: Date to check (defaults to today)
    
    Returns:
        dict: Count and remaining slots
    """
    try:
        if not application_date:
            application_date = nowdate()
        
        app_date = getdate(application_date)
        first_day = get_first_day(app_date)
        last_day = get_last_day(app_date)
        
        count = frappe.db.count("Scheme Request", filters={
            "doctor_code": doctor_code,
            "application_date": ["between", [first_day, last_day]],
            "docstatus": ["!=", 2]
        })
        
        remaining = max(0, 3 - count)
        month_name = app_date.strftime("%B %Y")
        
        return {
            "success": True,
            "doctor_code": doctor_code,
            "month": month_name,
            "count": count,
            "remaining": remaining,
            "limit": 3,
            "can_create": remaining > 0
        }
        
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Get Doctor Monthly Count Error")
        return {
            "success": False,
            "message": str(e)
        }

import frappe
from frappe.model.document import Document
from frappe.utils import flt, now

class HQYearlyTarget(Document):
    def validate(self):
        """Validate the target data before saving"""
        self.validate_dates()
        self.validate_hq_duplicates()
        self.calculate_totals()

    def validate_dates(self):
        """Ensure start date is before end date"""
        if self.start_date and self.end_date:
            if self.start_date > self.end_date:
                frappe.throw("Start Date cannot be after End Date")

    def validate_hq_duplicates(self):
        """Check for duplicate HQs in the targets table"""
        hq_list = []
        for item in self.hq_targets:
            if item.hq in hq_list:
                frappe.throw(f"Duplicate HQ found: {item.hq}. Each HQ can appear only once.")
            hq_list.append(item.hq)

    def before_save(self):
        """Calculate totals before saving"""
        self.calculate_totals()

        # Set upload metadata
        if not self.upload_date:
            self.upload_date = now()
        if not self.uploaded_by:
            self.uploaded_by = frappe.session.user

    def calculate_totals(self):
        """Calculate quarterly, yearly, and summary totals"""
        total_amount = 0
        total_hqs = len(self.hq_targets)

        for item in self.hq_targets:
            # Calculate quarterly totals
            item.q1_total = flt(item.apr) + flt(item.may) + flt(item.jun)
            item.q2_total = flt(item.jul) + flt(item.aug) + flt(item.sep)
            item.q3_total = flt(item.oct) + flt(item.nov) + flt(item.dec)
            item.q4_total = flt(item.jan) + flt(item.feb) + flt(item.mar)

            # Calculate yearly total
            item.yearly_total = (
                flt(item.q1_total) + 
                flt(item.q2_total) + 
                flt(item.q3_total) + 
                flt(item.q4_total)
            )

            total_amount += flt(item.yearly_total)

        # Set summary fields
        self.total_target_amount = total_amount
        self.total_hqs = total_hqs

@frappe.whitelist()
def get_region_hqs(region, division=None):
    """Get all HQs for a given region and optional division"""
    filters = {
        "region": region,
        "status": "Active"
    }

    if division:
        filters["division"] = division

    hqs = frappe.get_all(
        "HQ Master",
        filters=filters,
        fields=["name", "hq_name", "team"],
        order_by="name"
    )

    return hqs
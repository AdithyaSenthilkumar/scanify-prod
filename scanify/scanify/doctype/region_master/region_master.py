import frappe
from frappe.model.document import Document

class RegionMaster(Document):
    def before_save(self):
        # Sync region_code with the auto-generated name (R0001, R0002, ...)
        if self.name and not self.region_code:
            self.region_code = self.name

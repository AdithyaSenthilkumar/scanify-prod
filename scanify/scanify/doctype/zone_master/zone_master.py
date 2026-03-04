import frappe
from frappe.model.document import Document

class ZoneMaster(Document):
    def before_save(self):
        # Sync zone_code with the auto-generated name (Z0001, Z0002, ...)
        if self.name and not self.zone_code:
            self.zone_code = self.name

import frappe
from frappe.model.document import Document

class ZoneMaster(Document):
    def validate(self):
        self.check_duplicate_in_division()

    def before_save(self):
        # Sync zone_code with the auto-generated name (Z0001, Z0002, ...)
        if self.name and not self.zone_code:
            self.zone_code = self.name

    def check_duplicate_in_division(self):
        if not self.zone_name or not self.division:
            return
        filters = {"zone_name": self.zone_name, "division": self.division}
        if not self.is_new():
            filters["name"] = ["!=", self.name]
        if frappe.db.exists("Zone Master", filters):
            frappe.throw(f"Zone '{self.zone_name}' already exists in division '{self.division}'")

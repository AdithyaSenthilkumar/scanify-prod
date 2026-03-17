import frappe
from frappe.model.document import Document

class RegionMaster(Document):
    def validate(self):
        self.check_duplicate_in_division()

    def before_save(self):
        # Sync region_code with the auto-generated name (R0001, R0002, ...)
        if self.name and not self.region_code:
            self.region_code = self.name

    def check_duplicate_in_division(self):
        if not self.region_name or not self.division:
            return
        filters = {"region_name": self.region_name, "division": self.division}
        if not self.is_new():
            filters["name"] = ["!=", self.name]
        if frappe.db.exists("Region Master", filters):
            frappe.throw(f"Region '{self.region_name}' already exists in division '{self.division}'")

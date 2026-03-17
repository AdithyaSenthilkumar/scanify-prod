import frappe
from frappe.model.document import Document

class StateMaster(Document):
    def validate(self):
        self.check_duplicate_in_division()

    def before_save(self):
        # Sync state_code with the auto-generated name (ST0001, ST0002, ...)
        if self.name and not self.state_code:
            self.state_code = self.name

    def check_duplicate_in_division(self):
        if not self.state_name or not self.division:
            return
        filters = {"state_name": self.state_name, "division": self.division}
        if not self.is_new():
            filters["name"] = ["!=", self.name]
        if frappe.db.exists("State Master", filters):
            frappe.throw(f"State '{self.state_name}' already exists in division '{self.division}'")

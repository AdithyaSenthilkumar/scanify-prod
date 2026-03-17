import frappe
from frappe.model.document import Document

class TeamMaster(Document):
    def validate(self):
        self.check_duplicate_in_division()

    def before_save(self):
        # Sync team_code with the auto-generated name (T0001, T0002, ...)
        if self.name and not self.team_code:
            self.team_code = self.name

    def check_duplicate_in_division(self):
        if not self.team_name or not self.division:
            return
        filters = {"team_name": self.team_name, "division": self.division}
        if not self.is_new():
            filters["name"] = ["!=", self.name]
        if frappe.db.exists("Team Master", filters):
            frappe.throw(f"Team '{self.team_name}' already exists in division '{self.division}'")

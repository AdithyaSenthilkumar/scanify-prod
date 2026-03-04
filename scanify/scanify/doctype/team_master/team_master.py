import frappe
from frappe.model.document import Document

class TeamMaster(Document):
    def before_save(self):
        # Sync team_code with the auto-generated name (T0001, T0002, ...)
        if self.name and not self.team_code:
            self.team_code = self.name

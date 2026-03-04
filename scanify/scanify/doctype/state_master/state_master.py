import frappe
from frappe.model.document import Document

class StateMaster(Document):
    def before_save(self):
        # Sync state_code with the auto-generated name (ST0001, ST0002, ...)
        if self.name and not self.state_code:
            self.state_code = self.name

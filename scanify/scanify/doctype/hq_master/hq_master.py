import frappe
from frappe.model.document import Document

class HQMaster(Document):
	def validate(self):
		self.check_duplicate_in_division()

	def before_save(self):
		# Sync hq_code with the auto-generated name (HQ0001, HQ0002, ...)
		if self.name and not self.hq_code:
			self.hq_code = self.name

	def check_duplicate_in_division(self):
		if not self.hq_name or not self.division:
			return
		filters = {"hq_name": self.hq_name, "division": self.division}
		if not self.is_new():
			filters["name"] = ["!=", self.name]
		if frappe.db.exists("HQ Master", filters):
			frappe.throw(f"HQ '{self.hq_name}' already exists in division '{self.division}'")

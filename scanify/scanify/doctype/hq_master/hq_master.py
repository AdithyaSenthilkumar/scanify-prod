import frappe
from frappe.model.document import Document

class HQMaster(Document):
	def before_save(self):
		# Sync hq_code with the auto-generated name (HQ0001, HQ0002, ...)
		if self.name and not self.hq_code:
			self.hq_code = self.name

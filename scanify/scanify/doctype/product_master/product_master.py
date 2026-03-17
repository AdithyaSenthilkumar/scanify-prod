import frappe
from frappe.model.document import Document

class ProductMaster(Document):
	def validate(self):
		self.check_duplicate_in_division()

	def check_duplicate_in_division(self):
		if not self.product_code or not self.division:
			return
		filters = {"product_code": self.product_code, "division": self.division}
		if not self.is_new():
			filters["name"] = ["!=", self.name]
		if frappe.db.exists("Product Master", filters):
			frappe.throw(f"Product '{self.product_code}' already exists in division '{self.division}'")

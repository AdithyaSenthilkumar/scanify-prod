import frappe
from frappe.model.document import Document

class ProductMaster(Document):
	def validate(self):
		self.check_duplicate_in_division()
		self.set_excluded_region_codes()

	def set_excluded_region_codes(self):
		"""Mirror the selected excluded regions into a read-only comma-separated
		code string. Region codes (e.g. R0001) are what statements store and match
		against, so we surface them explicitly even though selection is by name."""
		seen = []
		for row in (self.excluded_regions or []):
			code = (row.region or "").strip()
			if code and code not in seen:
				seen.append(code)
		self.excluded_region_codes = ", ".join(seen)

	def check_duplicate_in_division(self):
		"""Product Code must be unique WITHIN a division (but the same code may be
		reused across different divisions). The autoname is a series (PRD-####) so
		the code is NOT the primary key; references link by the id and display the
		editable code. Legacy documents keep name == product_code."""
		if not self.product_code or not self.division:
			return
		filters = {"product_code": self.product_code, "division": self.division}
		if not self.is_new():
			filters["name"] = ["!=", self.name]
		if frappe.db.exists("Product Master", filters):
			frappe.throw(f"Product '{self.product_code}' already exists in division '{self.division}'")

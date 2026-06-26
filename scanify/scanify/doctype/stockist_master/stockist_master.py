import frappe
from frappe.model.document import Document

class StockistMaster(Document):
    def validate(self):
        self.set_division_from_hq()
        self.check_duplicate_in_division()
        self.check_duplicate_code_in_division()

    def before_save(self):
        # autoname (S0001) is ready here
        if not self.stockist_code:
            self.stockist_code = self.name

    def set_division_from_hq(self):
        if not self.hq:
            return

        division = frappe.db.get_value("HQ Master", self.hq, "division")
        if not division:
            frappe.throw(f"HQ {self.hq} does not have a Division set")

        self.division = division

    def check_duplicate_in_division(self):
        if not self.stockist_name or not self.division:
            return
        filters = {"stockist_name": self.stockist_name, "division": self.division, "hq": self.hq or ""}
        if not self.is_new():
            filters["name"] = ["!=", self.name]
        if frappe.db.exists("Stockist Master", filters):
            frappe.throw(f"Stockist '{self.stockist_name}' already exists in division '{self.division}' under the same HQ")

    def check_duplicate_code_in_division(self):
        """Stockist Code must be unique WITHIN a division (but the same code may be
        reused across different divisions). This replaces the old global unique
        constraint on stockist_code, which blocked legitimate cross-division reuse
        when masters for multiple divisions are bulk-uploaded."""
        if not self.stockist_code or not self.division:
            return
        filters = {"stockist_code": self.stockist_code, "division": self.division}
        if not self.is_new():
            filters["name"] = ["!=", self.name]
        if frappe.db.exists("Stockist Master", filters):
            frappe.throw(
                f"Stockist Code '{self.stockist_code}' already exists in division '{self.division}'"
            )


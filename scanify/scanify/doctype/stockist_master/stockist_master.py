import frappe
from frappe.model.document import Document

class StockistMaster(Document):
    def validate(self):
        self.set_division_from_hq()
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


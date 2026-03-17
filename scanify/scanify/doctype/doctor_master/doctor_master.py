import frappe
from frappe.model.document import Document

class DoctorMaster(Document):
    def before_save(self):
        # autoname (D0001) is available by the time before_save runs
        if not self.doctor_code:
            self.doctor_code = self.name

    def validate(self):
        self.set_division_from_hq()
        self.check_duplicate_in_division()

    def set_division_from_hq(self):
        if not self.hq:
            return

        division = frappe.db.get_value("HQ Master", self.hq, "division")
        if not division:
            frappe.throw(f"HQ {self.hq} does not have a Division set")

        self.division = division

    def check_duplicate_in_division(self):
        if not self.doctor_name or not self.division:
            return
        filters = {"doctor_name": self.doctor_name, "division": self.division, "hq": self.hq or ""}
        if not self.is_new():
            filters["name"] = ["!=", self.name]
        if frappe.db.exists("Doctor Master", filters):
            frappe.throw(f"Doctor '{self.doctor_name}' already exists in division '{self.division}' under the same HQ")
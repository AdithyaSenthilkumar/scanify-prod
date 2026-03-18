import frappe
from frappe.model.document import Document


class StockistProductCorrection(Document):
    def validate(self):
        # Normalize raw_product_name to uppercase/stripped for consistent matching
        if self.raw_product_name:
            self.raw_product_name = self.raw_product_name.strip().upper()

        # Enforce uniqueness of (stockist_code, raw_product_name)
        existing = frappe.db.get_value(
            "Stockist Product Correction",
            {
                "stockist_code": self.stockist_code,
                "raw_product_name": self.raw_product_name,
                "name": ["!=", self.name],
            },
            "name",
        )
        if existing:
            frappe.throw(
                f"A correction already exists for stockist {self.stockist_code} "
                f"and raw name '{self.raw_product_name}' ({existing}). "
                "Please update the existing record instead."
            )

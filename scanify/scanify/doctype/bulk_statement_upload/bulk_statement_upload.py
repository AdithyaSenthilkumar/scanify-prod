import frappe
from frappe.model.document import Document
import json

class BulkStatementUpload(Document):
    def after_insert(self):
        """Start bulk extraction after insert"""
        self.status = "In Progress"
        self.save()
        frappe.db.commit()
        
        # Call bulk extraction
        try:
            from scanify.api import bulk_extract_statements_async
            result = bulk_extract_statements_async(self.statement_month, self.zip_file)
            
            if result.get('success'):
                self.status = "Completed"
                self.extraction_log = json.dumps(result, indent=2)
            else:
                self.status = "Failed"
                self.extraction_log = result.get('message', 'Unknown error')
            
            self.save()
            frappe.db.commit()
            
        except Exception as e:
            self.status = "Failed"
            self.extraction_log = str(e)
            self.save()
            frappe.db.commit()
            frappe.log_error(frappe.get_traceback(), "Bulk Upload Error")

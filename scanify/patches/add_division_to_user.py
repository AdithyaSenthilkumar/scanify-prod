import frappe

def execute():
    """Add division field to User doctype"""
    
    try:
        # Check if field already exists
        if frappe.db.exists("Custom Field", "User-division"):
            print("Division field already exists in User doctype")
            return
        
        # Create custom field
        custom_field = frappe.get_doc({
            "doctype": "Custom Field",
            "dt": "User",
            "label": "Division",
            "fieldname": "division",
            "fieldtype": "Select",
            "options": "Prima\nVektra",
            "insert_after": "email",
            "allow_on_submit": 0,
            "translatable": 0,
            "in_list_view": 0,
            "in_standard_filter": 1,
            "bold": 0,
            "collapsible": 0,
            "ignore_user_permissions": 0,
            "ignore_xss_filter": 0,
            "no_copy": 0,
            "permlevel": 0,
            "print_hide": 0,
            "read_only": 0,
            "report_hide": 0,
            "reqd": 0,
            "search_index": 0,
            "unique": 0
        })
        
        custom_field.insert(ignore_permissions=True)
        frappe.db.commit()
        
        print("✓ Division field added to User doctype successfully")
        
    except Exception as e:
        frappe.log_error(f"Error adding division field: {str(e)}")
        print(f"✗ Error adding division field: {str(e)}")

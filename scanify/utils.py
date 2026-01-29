import frappe
from frappe.utils import flt, getdate, add_months

def import_scheme_master_data(file_path):
	"""Import scheme master data from Excel"""
	import pandas as pd
	
	df = pd.read_excel(file_path)
	
	for idx, row in df.iterrows():
		# Check if doctor exists
		if not frappe.db.exists("Doctor Master", row['doc_code']):
			doctor = frappe.get_doc({
				"doctype": "Doctor Master",
				"doctor_code": row['doc_code'],
				"doctor_name": row['doc_name'],
				"place": row['doc_place'],
				"team": row['team'],
				"region": row['region']
			})
			doctor.insert(ignore_permissions=True)
		
		# Check if stockist exists
		if not frappe.db.exists("Stockist Master", row['stc_code']):
			stockist = frappe.get_doc({
				"doctype": "Stockist Master",
				"stockist_code": row['stc_code'],
				"stockist_name": row['stc_name'],
				"hq": get_hq_from_team(row['team'])
			})
			stockist.insert(ignore_permissions=True)
		
		# Create scheme request
		scheme = frappe.get_doc({
			"doctype": "Scheme Request",
			"entry_date": getdate(row['entry_date']),
			"application_date": getdate(row['app_date']),
			"doctor_code": row['doc_code'],
			"stockist_code": row['stc_code'],
			"requested_by": "Administrator"
		})
		
		scheme.append("items", {
			"product_code": row['prod_code'],
			"quantity": flt(row['prod_qty']),
			"free_quantity": flt(row['prod_free_qty']),
			"special_rate": flt(row['prod_spl_rate']) if row['prod_spl_rate'] else flt(row['prod_rate'])
		})
		
		scheme.insert(ignore_permissions=True)
		frappe.db.commit()
		
		if idx % 100 == 0:
			print(f"Processed {idx} records")

def get_hq_from_team(team_name):
	"""Get first HQ for a team"""
	hq = frappe.db.get_value("HQ Master", {"team": team_name}, "name")
	return hq

def generate_monthly_statements_template(stockist_code, month):
	"""Generate template for monthly statement entry"""
	products = frappe.get_all("Product Master", 
		filters={"status": "Active"},
		fields=["product_code", "product_name", "pack", "pts"])
	
	# Get previous month closing as opening
	prev_month = add_months(getdate(month), -1)
	prev_statement = frappe.db.get_value("Stockist Statement", {
		"stockist_code": stockist_code,
		"statement_month": prev_month,
		"docstatus": 1
	}, "name")
	
	opening_balances = {}
	if prev_statement:
		items = frappe.get_all("Stockist Statement Item",
			filters={"parent": prev_statement},
			fields=["product_code", "closing_qty"])
		opening_balances = {item.product_code: item.closing_qty for item in items}
	
	# Create new statement
	statement = frappe.get_doc({
		"doctype": "Stockist Statement",
		"stockist_code": stockist_code,
		"statement_month": month,
		"from_date": frappe.utils.get_first_day_of_the_month(month),
		"to_date": frappe.utils.get_last_day_of_the_month(month)
	})
	
	for product in products:
		statement.append("items", {
			"product_code": product.product_code,
			"opening_qty": opening_balances.get(product.product_code, 0)
		})
	
	return statement

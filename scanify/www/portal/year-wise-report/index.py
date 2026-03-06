import frappe
from scanify.api import get_user_division

def get_context(context):
    if frappe.session.user == "Guest":
        frappe.throw("You must be logged in.", frappe.PermissionError)

    context.no_cache = 1
    user_division = get_user_division() or "Prima"
    context.division = user_division
    
    financial_year = frappe.form_dict.get('financial_year')
    context.financial_year = financial_year
    
    # Get available financial years from targets
    f_years = frappe.db.sql("""
        SELECT DISTINCT financial_year 
        FROM `tabHQ Yearly Target` 
        WHERE docstatus = 1 AND division = %s
        ORDER BY financial_year DESC
    """, (user_division,), as_dict=True)
    
    context.available_years = [fy.financial_year for fy in f_years if fy.financial_year]
    
    if not financial_year and context.available_years:
        financial_year = context.available_years[0]
        context.financial_year = financial_year
        
    context.report_data = []
    
    if financial_year:
        # Fetch targets grouped by HQ, joined with HQ Master to get accurate Team and Region maps
        raw_data = frappe.db.sql("""
            SELECT 
                ti.hq,
                COALESCE(ti.hq_name, hm.hq_name, ti.hq) as hq_name,
                COALESCE(hm.team, ti.team) as team,
                hm.region as region,
                SUM(ti.apr) as apr, SUM(ti.may) as may, SUM(ti.jun) as jun,
                SUM(ti.jul) as jul, SUM(ti.aug) as aug, SUM(ti.sep) as sep,
                SUM(ti.oct) as oct, SUM(ti.nov) as nov, SUM(ti.`dec`) as `dec`,
                SUM(ti.jan) as jan, SUM(ti.feb) as feb, SUM(ti.mar) as mar,
                SUM(ti.yearly_total) as yearly_total
            FROM `tabHQ Yearly Target` yt
            INNER JOIN `tabHQ Target Item` ti ON ti.parent = yt.name AND ti.parenttype = 'HQ Yearly Target'
            LEFT JOIN `tabHQ Master` hm ON hm.name = ti.hq
            WHERE yt.docstatus = 1 
              AND yt.financial_year = %s 
              AND yt.division = %s
            GROUP BY ti.hq, COALESCE(ti.hq_name, hm.hq_name, ti.hq), COALESCE(hm.team, ti.team), hm.region
            ORDER BY region, team, hq_name
        """, (financial_year, user_division), as_dict=True)
        
        # Organize data by Region -> Team -> HQ
        organized = {}
        for row in raw_data:
            r = row.region or "Unknown Region"
            t = row.team or "Unknown Team"
            
            if r not in organized:
                organized[r] = {
                    "teams": {},
                    "totals": {"apr": 0, "may": 0, "jun": 0, "q1": 0, "jul": 0, "aug": 0, "sep": 0, "q2": 0, "oct": 0, "nov": 0, "dec": 0, "q3": 0, "jan": 0, "feb": 0, "mar": 0, "q4": 0, "total": 0}
                }
                
            if t not in organized[r]["teams"]:
                organized[r]["teams"][t] = {
                    "hqs": [],
                    "totals": {"apr": 0, "may": 0, "jun": 0, "q1": 0, "jul": 0, "aug": 0, "sep": 0, "q2": 0, "oct": 0, "nov": 0, "dec": 0, "q3": 0, "jan": 0, "feb": 0, "mar": 0, "q4": 0, "total": 0}
                }
                
            q1 = (row.apr or 0) + (row.may or 0) + (row.jun or 0)
            q2 = (row.jul or 0) + (row.aug or 0) + (row.sep or 0)
            q3 = (row.oct or 0) + (row.nov or 0) + (row.dec or 0)
            q4 = (row.jan or 0) + (row.feb or 0) + (row.mar or 0)
            
            hq_row = {
                "hq_name": row.hq_name,
                "apr": row.apr or 0, "may": row.may or 0, "jun": row.jun or 0, "q1": q1,
                "jul": row.jul or 0, "aug": row.aug or 0, "sep": row.sep or 0, "q2": q2,
                "oct": row.oct or 0, "nov": row.nov or 0, "dec": row.dec or 0, "q3": q3,
                "jan": row.jan or 0, "feb": row.feb or 0, "mar": row.mar or 0, "q4": q4,
                "total": row.yearly_total or 0
            }
            organized[r]["teams"][t]["hqs"].append(hq_row)
            
            # Add to Team totals
            for k in hq_row:
                if k != "hq_name":
                    organized[r]["teams"][t]["totals"][k] += hq_row[k]
                    organized[r]["totals"][k] += hq_row[k]
                    
        context.report_data = organized

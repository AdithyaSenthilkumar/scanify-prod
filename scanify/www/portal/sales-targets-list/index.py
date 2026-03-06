import frappe
from scanify.api import get_user_division

def get_context(context):
    if frappe.session.user == "Guest":
        frappe.throw("You must be logged in to view sales targets", frappe.PermissionError)

    context.no_cache = 1

    user_division = get_user_division() or "Prima"
    context.division = user_division
    
    selected_region = frappe.form_dict.get('region') or ''
    selected_team = frappe.form_dict.get('team') or ''
    
    context.selected_region = selected_region
    context.selected_team = selected_team

    # Fetch filters for division
    context.regions = frappe.get_all("Region Master", filters={"division": ["in", [user_division, "Both"]], "status": "Active"}, fields=["name", "region_name"], order_by="region_name asc")
    context.teams = frappe.get_all("Team Master", filters={"division": ["in", [user_division, "Both"]], "status": "Active"}, fields=["name", "team_name"], order_by="team_name asc")

    conds = ""
    if selected_region:
        conds += " AND t.region = %(selected_region)s"
    if selected_team:
        conds += " AND ti.team = %(selected_team)s"

    context.target_rows = frappe.db.sql(
        f"""
        SELECT
            t.name AS target_id,
            t.financial_year,
            t.start_date,
            t.end_date,
            t.status,
            t.docstatus,
            t.creation,
            ti.hq,
            COALESCE(ti.hq_name, hq.hq_name, ti.hq) AS hq_name,
            ti.team,
            ti.yearly_total
        FROM `tabHQ Yearly Target` t
        INNER JOIN `tabHQ Target Item` ti
            ON ti.parent = t.name
            AND ti.parenttype = 'HQ Yearly Target'
            AND ti.parentfield = 'hq_targets'
        LEFT JOIN `tabHQ Master` hq ON hq.name = ti.hq
        WHERE t.docstatus < 2
          AND t.division = %(division)s
          {conds}
        ORDER BY t.creation DESC, ti.idx ASC
        LIMIT 500
        """,
        {
            "division": user_division,
            "selected_region": selected_region,
            "selected_team": selected_team
        },
        as_dict=True,
    )

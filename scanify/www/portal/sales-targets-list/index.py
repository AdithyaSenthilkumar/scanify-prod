import frappe
from scanify.api import get_user_division

def get_context(context):
    if frappe.session.user == "Guest":
        frappe.throw("You must be logged in to view sales targets", frappe.PermissionError)

    context.no_cache = 1

    user_division = get_user_division() or "Prima"
    context.division = user_division

    context.target_rows = frappe.db.sql(
        """
        SELECT
            t.name AS target_id,
            t.financial_year,
            t.start_date,
            t.end_date,
            t.status,
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
        ORDER BY t.creation DESC, ti.idx ASC
        LIMIT 500
        """,
        {"division": user_division},
        as_dict=True,
    )

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

    # Available financial years — include Draft targets (docstatus 0) so a target
    # can be reported before it is approved, not just submitted (docstatus 1).
    f_years = frappe.db.sql("""
        SELECT DISTINCT financial_year
        FROM `tabHQ Yearly Target`
        WHERE docstatus < 2 AND division = %s
        ORDER BY financial_year DESC
    """, (user_division,), as_dict=True)

    context.available_years = [fy.financial_year for fy in f_years if fy.financial_year]

    if not financial_year and context.available_years:
        financial_year = context.available_years[0]
        context.financial_year = financial_year

    context.report_data = []
    context.report_status = ""
    context.has_draft = False

    if financial_year:
        # Pull every HQ target row for the FY from BOTH Draft and Approved targets.
        # Ordered so the highest-priority document per HQ comes first: Approved
        # (docstatus 1) before Draft (0), then most recently modified. We then keep
        # only the first row per HQ, so a Draft + Approved pair for the same FY never
        # double-counts and the Approved figure always wins.
        # Non-admins only see HQs inside the regions they are mapped to.
        from scanify.permissions import get_allowed_region_codes
        allowed_regions = get_allowed_region_codes(division=user_division)
        region_cond, region_params = "", []
        if allowed_regions is not None:
            if allowed_regions:
                region_cond = " AND hm.region IN ({})".format(
                    ", ".join(["%s"] * len(allowed_regions)))
                region_params = list(allowed_regions)
            else:
                region_cond = " AND 1=0"

        # Resolve Region/Team/HQ codes to their display names here, so the report shows
        # readable names (e.g. "Chennai" / "Team South") instead of R0xxx / T0xxx codes.
        rows = frappe.db.sql("""
            SELECT
                ti.hq,
                COALESCE(ti.hq_name, hm.hq_name, ti.hq) AS hq_name,
                COALESCE(tm.team_name, hm.team, ti.team) AS team,
                COALESCE(rm.region_name, hm.region) AS region,
                ti.apr, ti.may, ti.jun, ti.jul, ti.aug, ti.sep,
                ti.oct, ti.nov, ti.`dec`, ti.jan, ti.feb, ti.mar,
                ti.yearly_total,
                yt.docstatus AS docstatus,
                yt.modified AS modified
            FROM `tabHQ Yearly Target` yt
            INNER JOIN `tabHQ Target Item` ti
                ON ti.parent = yt.name AND ti.parenttype = 'HQ Yearly Target'
            LEFT JOIN `tabHQ Master` hm ON hm.name = ti.hq
            LEFT JOIN `tabTeam Master` tm ON tm.name = COALESCE(hm.team, ti.team)
            LEFT JOIN `tabRegion Master` rm ON rm.name = hm.region
            WHERE yt.docstatus < 2
              AND yt.financial_year = %s
              AND yt.division = %s
        """ + region_cond + """
            ORDER BY ti.hq, yt.docstatus DESC, yt.modified DESC
        """, tuple([financial_year, user_division] + region_params), as_dict=True)

        # Deduplicate per HQ (first wins = Approved > latest Draft).
        seen = set()
        raw_data = []
        has_draft = False
        for row in rows:
            if row.hq in seen:
                continue
            seen.add(row.hq)
            raw_data.append(row)
            if row.docstatus == 0:
                has_draft = True

        context.has_draft = has_draft
        if raw_data:
            context.report_status = "Draft" if has_draft else "Approved"

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

            # Add to Team + Region totals
            for k in hq_row:
                if k != "hq_name":
                    organized[r]["teams"][t]["totals"][k] += hq_row[k]
                    organized[r]["totals"][k] += hq_row[k]

        context.report_data = organized

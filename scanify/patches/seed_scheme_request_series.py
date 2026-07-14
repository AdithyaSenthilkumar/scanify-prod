import re

import frappe


def execute():
    """Seed `tabSeries` for the per-year Scheme Request naming series.

    Scheme Request's autoname changed from `format:SCH-{YYYY}-{#####}` to the dotted
    naming series `SCH-.YYYY.-.#####`. The old `format:` path resolved `{#####}`
    independently of the rest of the string, so every scheme drew from a single global
    (empty-key) counter that never reset per year. The dotted series instead keys the
    counter by the fully-resolved prefix `SCH-<year>-`, so each new year restarts at 1.

    To avoid colliding with names already issued this/previous years, seed each existing
    year's counter to the highest serial already used, so the next document continues
    from there. Idempotent: safe to re-run.
    """
    rows = frappe.db.sql(
        "SELECT name FROM `tabScheme Request` WHERE name LIKE 'SCH-%'",
        as_dict=True,
    )

    # year prefix (e.g. 'SCH-2026-') -> highest serial seen
    max_by_prefix = {}
    pattern = re.compile(r"^(SCH-\d{4}-)(\d+)$")
    for r in rows:
        m = pattern.match(r["name"] or "")
        if not m:
            continue
        prefix, serial = m.group(1), int(m.group(2))
        if serial > max_by_prefix.get(prefix, 0):
            max_by_prefix[prefix] = serial

    for prefix, current in max_by_prefix.items():
        # `tabSeries` has only name/current columns, so query it directly rather than
        # via frappe.db.get_value (which would add an ORDER BY creation and fail).
        row = frappe.db.sql("SELECT current FROM `tabSeries` WHERE name = %s", (prefix,))
        existing = row[0][0] if row else None
        if existing is None:
            frappe.db.sql(
                "INSERT INTO `tabSeries` (name, current) VALUES (%s, %s)",
                (prefix, current),
            )
            print(f"✓ Seeded Series `{prefix}` at {current}")
        elif int(existing) < current:
            frappe.db.sql(
                "UPDATE `tabSeries` SET current = %s WHERE name = %s",
                (current, prefix),
            )
            print(f"✓ Advanced Series `{prefix}` {existing} -> {current}")
        else:
            print(f"• Series `{prefix}` already at {existing} (>= {current}), left as-is")

    if not max_by_prefix:
        print("No existing SCH-YYYY-##### names found; nothing to seed.")

    frappe.db.commit()

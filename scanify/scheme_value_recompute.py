"""One-off maintenance: recompute Scheme Request Item.product_value to Qty x Rate.

Background
----------
A scheme line's order value is `quantity * (special_rate or product_rate)`. Order
qty is captured in box/units and both rates are per box/unit, so no pack
conversion applies — this matches all 30,091 historical lines in the client's
migration sheets (column "prod_qty in box/ units").

Two earlier formulas stored wrong values for portal-created requests:
  * before 2026-07-14 — free-goods lines were valued off the FREE qty:
        (free_qty / strips_per_box) * pts
  * 2026-07-14 .. 2026-07-24 — every line was divided by the pack factor:
        (quantity / strips_per_box) * rate      (understated NxM packs by 10x etc.)

This script rewrites product_value (and the parent's total_scheme_value) for
those rows.

Migrated requests are NEVER touched. Their product_value is the client's own
authoritative figure, and their product_rate was fetched from today's Product
Master (not the historical rate on the sheet), so recomputing would silently
restate history. They are identified by the backfill's markers on the parent.

Usage
-----
    bench --site <site> execute scanify.scheme_value_recompute.run
    bench --site <site> execute scanify.scheme_value_recompute.run \
        --kwargs "{'mode': 'commit'}"
    bench --site <site> execute scanify.scheme_value_recompute.run \
        --kwargs "{'mode': 'revert', 'log_file': '/path/to/scheme_value_recompute_<ts>.log'}"

'dryrun' (the default) changes nothing and prints the full impact breakdown.
'commit' writes a tab-separated undo log next to the site (item name, old value,
new value) which 'revert' replays in reverse.
"""

import datetime
import os

import frappe
from frappe.utils import flt

# Markers written by the 2026-07-22 historical backfill (see _backfill_tmp/scheme_backfill.py)
BACKFILL_NOTE = "Historical scheme — migrated"
EMAIL_MARKER = "Backfill (not emailed)"

MIGRATED_COND = (
    "(IFNULL(sr.scheme_notes, '') LIKE %(bf_note)s"
    " OR IFNULL(sr.email_sent_to, '') = %(bf_email)s)"
)


def _log_path(site_path=None):
    # Absolute: frappe.get_site_path() is relative to the bench `sites/` dir, and the
    # revert run may be launched from somewhere else.
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    base = os.path.abspath(site_path or os.getcwd())
    return os.path.join(base, f"scheme_value_recompute_{stamp}.log")


def run(mode="dryrun", since=None, log_file=None):
    """Recompute product_value = quantity x (special_rate or product_rate).

    mode:  'dryrun' (default) | 'commit' | 'revert'
    since: optional 'YYYY-MM-DD'; only requests created on/after this date.
    """
    if mode == "revert":
        return _revert(log_file)

    if mode not in ("dryrun", "commit"):
        frappe.throw("mode must be one of: dryrun, commit, revert")

    params = {"bf_note": BACKFILL_NOTE + "%", "bf_email": EMAIL_MARKER}
    conds = [f"NOT {MIGRATED_COND}", "sr.docstatus != 2"]
    if since:
        conds.append("sr.creation >= %(since)s")
        params["since"] = since

    rows = frappe.db.sql(
        f"""
        SELECT sri.name AS item, sri.parent AS req, sr.creation AS created,
               sr.application_date, sri.product_code, sri.pack,
               IFNULL(sri.quantity, 0)      AS quantity,
               IFNULL(sri.free_quantity, 0) AS free_quantity,
               IFNULL(sri.product_rate, 0)  AS product_rate,
               IFNULL(sri.special_rate, 0)  AS special_rate,
               IFNULL(sri.product_value, 0) AS product_value
          FROM `tabScheme Request Item` sri
    INNER JOIN `tabScheme Request` sr ON sr.name = sri.parent
         WHERE {' AND '.join(conds)}
      ORDER BY sr.creation, sri.idx
        """,
        params,
        as_dict=True,
    )

    changes = []          # (item, req, old, new)
    unchanged = 0
    for r in rows:
        rate = flt(r.special_rate) if flt(r.special_rate) > 0 else flt(r.product_rate)
        new_value = flt(r.quantity) * rate
        if abs(new_value - flt(r.product_value)) < 0.005:
            unchanged += 1
            continue
        changes.append((r.item, r.req, flt(r.product_value), new_value, r))

    total_old = sum(c[2] for c in changes)
    total_new = sum(c[3] for c in changes)
    reqs = sorted({c[1] for c in changes})

    print(f"Scope: {len(rows)} non-migrated scheme lines"
          + (f" created on/after {since}" if since else " (all dates)"))
    print(f"  already correct : {unchanged}")
    print(f"  to be corrected : {len(changes)} lines across {len(reqs)} requests")
    print(f"  value {total_old:,.2f} -> {total_new:,.2f}  (delta {total_new - total_old:+,.2f})")

    if changes:
        print("\n  sample (first 15):")
        print("   {:<28} {:<10} {:>9} {:>10} {:>10} {:>13} {:>13}".format(
            "REQUEST", "PRODUCT", "QTY", "RATE", "SPL", "OLD VALUE", "NEW VALUE"))
        for item, req, old, new, r in changes[:15]:
            print("   {:<28} {:<10} {:>9.2f} {:>10.2f} {:>10.2f} {:>13.2f} {:>13.2f}".format(
                req, str(r.product_code)[:10], flt(r.quantity), flt(r.product_rate),
                flt(r.special_rate), old, new))

    if mode == "dryrun":
        print("\nDRY RUN - nothing written. Re-run with mode='commit' to apply.")
        return {"scope": len(rows), "changes": len(changes), "requests": len(reqs)}

    if not changes:
        print("\nNothing to do.")
        return {"scope": len(rows), "changes": 0, "requests": 0}

    path = log_file or _log_path(frappe.get_site_path())
    with open(path, "w", encoding="utf-8") as f:
        f.write("# item\told_value\tnew_value\n")
        for item, req, old, new, _r in changes:
            f.write(f"{item}\t{old!r}\t{new!r}\n")

    for item, _req, _old, new, _r in changes:
        frappe.db.sql(
            "UPDATE `tabScheme Request Item` SET product_value = %s WHERE name = %s",
            (new, item))

    _resync_totals(reqs)
    frappe.db.commit()
    print(f"\nCOMMITTED {len(changes)} lines across {len(reqs)} requests.")
    print(f"Undo log: {path}")
    print(f"Revert with: mode='revert', log_file='{path}'")
    return {"scope": len(rows), "changes": len(changes), "requests": len(reqs), "log": path}


def _resync_totals(req_names):
    """Set each parent's total_scheme_value to the sum of its line values."""
    for i in range(0, len(req_names), 500):
        chunk = req_names[i:i + 500]
        ph = ", ".join(["%s"] * len(chunk))
        frappe.db.sql(
            f"""UPDATE `tabScheme Request` sr
                   SET sr.total_scheme_value = (
                        SELECT IFNULL(SUM(sri.product_value), 0)
                          FROM `tabScheme Request Item` sri
                         WHERE sri.parent = sr.name)
                 WHERE sr.name IN ({ph})""",
            chunk)


def _revert(log_file):
    """Restore product_value from an undo log written by mode='commit'."""
    if not log_file or not os.path.exists(log_file):
        frappe.throw(f"log_file not found: {log_file}")

    restored = 0
    reqs = set()
    with open(log_file, encoding="utf-8") as f:
        for line in f:
            line = line.rstrip("\n")
            if not line or line.startswith("#"):
                continue
            item, old, _new = line.split("\t")
            frappe.db.sql(
                "UPDATE `tabScheme Request Item` SET product_value = %s WHERE name = %s",
                (float(old), item))
            parent = frappe.db.get_value("Scheme Request Item", item, "parent")
            if parent:
                reqs.add(parent)
            restored += 1

    _resync_totals(sorted(reqs))
    frappe.db.commit()
    print(f"Reverted {restored} lines across {len(reqs)} requests from {log_file}")
    return {"reverted": restored, "requests": len(reqs)}

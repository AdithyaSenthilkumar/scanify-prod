"""
Secondary Sales Import – April 2026  (PRODUCTION COPY)
=======================================================

DEPLOYMENT STEPS (do these on the production server):
  1. Copy this file to:
         <bench>/apps/scanify/scanify/import_april_secondary_prod.py

  2. Copy the Excel file to the server at a known path.
     Update EXCEL_PATH below to match that path.

  3. Update SITE_NAME below to your production site name.

  4. Open a terminal in the bench folder and run dry-run first:
         bench --site <site_name> execute scanify.import_april_secondary_prod.run --kwargs '{"mode": "dry_run"}'

  5. Review the log file written at LOG_FILE. If everything looks correct:
         bench --site <site_name> execute scanify.import_april_secondary_prod.run

LOGIC:
  - Stockist name matched EXACTLY only (fuzzy matches are SKIPPED and logged)
  - Product matched by code first, then fuzzy name (threshold 60%)
  - purchase_qty = clstock - opstock + quantity + freeqty
  - All value computations handled by StockistStatement.validate()
  - Duplicate statements (same stockist + month) are skipped safely
  - All results written to LOG_FILE
"""

import frappe
import openpyxl
from frappe.utils import flt, now_datetime
from difflib import SequenceMatcher

# ═══════════════════════════════════════════════════════════════════
#  CONFIGURATION  ←  UPDATE THESE BEFORE RUNNING ON PRODUCTION
# ═══════════════════════════════════════════════════════════════════

EXCEL_PATH      = "/home/adithya/secondary_upload_april.xlsx"
#                  ^ Copy the Excel file to the server and set the path here.
#                    On Windows paths use forward slashes:
#                    e.g. "C:/Users/You/Desktop/secondary upload april.xlsx"

STATEMENT_MONTH = "2026-04-01"
#                  ^ First day of the month being imported.

LOG_FILE        = "/home/adithya/migration_log_april_2026.txt"
#                  ^ Full path for the output log. Must be writable.
#                    On Windows: "C:/Users/You/Desktop/migration_log_april_2026.txt"

FUZZY_THRESHOLD = 0.55   # Detect near-matches for the log (stockist is still SKIPPED)
PRODUCT_FUZZY   = 0.60   # Min similarity to auto-accept product name match

# ═══════════════════════════════════════════════════════════════════


def _ratio(a, b):
    return SequenceMatcher(None, a.lower().strip(), b.lower().strip()).ratio()


def _best_match(query, candidates, key_fn, threshold):
    best, best_score = None, 0.0
    for c in candidates:
        score = _ratio(query, key_fn(c))
        if score > best_score:
            best, best_score = c, score
    if best_score >= threshold:
        return best, best_score
    return None, best_score


def load_masters():
    stockists = frappe.get_all(
        "Stockist Master",
        fields=["name", "stockist_name", "stockist_code", "hq", "division"],
        filters={"status": "Active"},
        order_by="stockist_name",
    )
    products = frappe.get_all(
        "Product Master",
        fields=["name", "product_code", "product_name", "pts", "pack"],
        filters={"status": "Active"},
        order_by="product_code",
    )
    product_by_code = {p.product_code.strip().upper(): p for p in products}
    return stockists, products, product_by_code


def match_stockist(old_code, old_name, stockists):
    """
    Returns (match, score, method).
    method = "exact"    → safe to import
    method = "fuzzy"    → SKIP; logged for manual correction
    method = "no_match" → SKIP; stockist not in master at all
    """
    old_clean = old_name.strip().lower()
    for s in stockists:
        if s.stockist_name.strip().lower() == old_clean:
            return s, 1.0, "exact"

    best, score = _best_match(old_name, stockists, lambda s: s.stockist_name, FUZZY_THRESHOLD)
    if best:
        return best, score, "fuzzy"
    return None, 0.0, "no_match"


def match_product(pcode, pname, product_by_code, products):
    code_upper = pcode.strip().upper()
    if code_upper in product_by_code:
        return product_by_code[code_upper], 1.0, "exact_code"

    best, score = _best_match(pname, products, lambda p: p.product_name, PRODUCT_FUZZY)
    if best:
        return best, score, "fuzzy_name"
    return None, 0.0, "no_match"


def read_excel():
    wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    headers = [str(h).strip().lower() if h else "" for h in rows[0]]
    col = {h: i for i, h in enumerate(headers)}

    required = ["stockistcode", "stockistname", "pcode", "product",
                "quantity", "clstock", "freeqty", "opstock", "pts"]
    missing = [r for r in required if r not in col]
    if missing:
        frappe.throw(f"Required columns not found in Excel: {missing}\nFound columns: {headers}")

    groups = {}
    blank_rows = 0
    for row in rows[1:]:
        stockist_code = str(row[col["stockistcode"]] or "").strip()
        stockist_name = str(row[col["stockistname"]] or "").strip()
        pcode         = str(row[col["pcode"]] or "").strip()
        pname         = str(row[col["product"]] or "").strip()

        if not stockist_code and not stockist_name:
            blank_rows += 1
            continue
        if not pcode and not pname:
            blank_rows += 1
            continue

        key = (stockist_code, stockist_name)
        if key not in groups:
            groups[key] = []
        groups[key].append({
            "pcode"  : pcode,
            "pname"  : pname,
            "pack"   : str(row[col["pack"]] or "").strip() if "pack" in col else "",
            "qty"    : flt(row[col["quantity"]]),
            "clstock": flt(row[col["clstock"]]),
            "freeqty": flt(row[col["freeqty"]]),
            "opstock": flt(row[col["opstock"]]),
            "pts"    : flt(row[col["pts"]]),
        })

    print(f"Excel: {len(groups)} stockists, {sum(len(v) for v in groups.values())} product rows, {blank_rows} blank rows skipped")
    return groups


def run(mode="live"):
    """
    mode='dry_run'  → report only, nothing is saved
    mode='live'     → create Stockist Statement drafts
    """
    dry = (mode == "dry_run")
    lines = []

    def log(msg=""):
        print(msg)
        lines.append(str(msg))

    log("=" * 70)
    log(f"  Secondary Import – April 2026  |  {'DRY RUN' if dry else 'LIVE'}")
    log(f"  Site: {frappe.local.site}")
    log(f"  Run at: {now_datetime()}")
    log(f"  Excel: {EXCEL_PATH}")
    log("=" * 70)
    log()

    stockists, products, product_by_code = load_masters()
    log(f"  Masters loaded: {len(stockists)} stockists, {len(products)} products")
    log()

    groups = read_excel()
    log()

    summary = {
        "created"              : 0,
        "skipped_exists"       : [],   # (doc_name, stockist_name)
        "stockist_no_match"    : [],   # (old_name, old_code)
        "stockist_fuzzy_skipped": [],  # (old_name, old_code, nearest_name, score)
        "product_no_match"     : [],   # (pcode, pname, stockist_name)
    }

    for (old_code, old_name), product_rows in sorted(groups.items(), key=lambda x: x[0][1]):

        # ── Match stockist ──────────────────────────────────────────────────
        s_match, s_score, s_method = match_stockist(old_code, old_name, stockists)

        if s_method == "fuzzy":
            log(f"  ~ FUZZY SKIP : [{old_code}] '{old_name}'")
            log(f"                 nearest: '{s_match.stockist_name}' ({s_score:.0%})  – add to master or fix name")
            summary["stockist_fuzzy_skipped"].append((old_name, old_code, s_match.stockist_name, s_score))
            continue

        if s_method == "no_match":
            log(f"  X NO MATCH   : [{old_code}] '{old_name}' – not in Stockist Master")
            summary["stockist_no_match"].append((old_name, old_code))
            continue

        stockist_code = s_match.name

        # ── Skip if already exists ──────────────────────────────────────────
        existing = frappe.db.exists("Stockist Statement", {
            "stockist_code"  : stockist_code,
            "statement_month": STATEMENT_MONTH,
        })
        if existing:
            summary["skipped_exists"].append((existing, s_match.stockist_name))
            continue

        # ── Build items list ────────────────────────────────────────────────
        items = []
        for pr in product_rows:
            p_match, p_score, p_method = match_product(
                pr["pcode"], pr["pname"], product_by_code, products
            )
            purchase_qty = max(0, flt(pr["clstock"]) - flt(pr["opstock"]) + flt(pr["qty"]) + flt(pr["freeqty"]))

            if not p_match:
                log(f"    X NO PRODUCT: [{pr['pcode']}] '{pr['pname']}' in '{old_name}'")
                summary["product_no_match"].append((pr["pcode"], pr["pname"], old_name))
                items.append({
                    "product_code"   : None,
                    "raw_product_name": f"{pr['pcode']} - {pr['pname']}",
                    "mapping_status" : "unmapped",
                    "opening_qty"    : pr["opstock"],
                    "purchase_qty"   : purchase_qty,
                    "sales_qty"      : pr["qty"],
                    "free_qty"       : pr["freeqty"],
                    "closing_qty"    : pr["clstock"],
                    "pts"            : pr["pts"],
                    "row_type"       : "product",
                })
                continue

            item = {
                "product_code"    : p_match.product_code,
                "raw_product_name": pr["pname"],
                "mapping_status"  : "matched" if p_method == "exact_code" else "auto_mapped",
                "opening_qty"     : pr["opstock"],
                "purchase_qty"    : purchase_qty,
                "sales_qty"       : pr["qty"],
                "free_qty"        : pr["freeqty"],
                "closing_qty"     : pr["clstock"],
                "row_type"        : "product",
            }
            if pr["pts"] > 0:
                item["pts"] = pr["pts"]
            items.append(item)

        # ── Dry run ─────────────────────────────────────────────────────────
        if dry:
            log(f"  + WOULD CREATE: {s_match.stockist_name} ({stockist_code}) – {len(items)} products")
            summary["created"] += 1
            continue

        # ── Create document ─────────────────────────────────────────────────
        try:
            doc = frappe.new_doc("Stockist Statement")
            doc.stockist_code         = stockist_code
            doc.statement_month       = STATEMENT_MONTH
            doc.extracted_data_status = "Completed"
            doc.extraction_notes      = f"Imported from secondary upload Excel (old code: {old_code})"
            doc.qc_confidence         = "Verification Needed"
            for it in items:
                doc.append("items", it)
            doc.insert(ignore_permissions=True)
            frappe.db.commit()
            log(f"  + CREATED: {doc.name} | {s_match.stockist_name} | {len(items)} rows")
            summary["created"] += 1
        except Exception as e:
            frappe.db.rollback()
            log(f"  ! ERROR: {s_match.stockist_name} ({stockist_code}) — {e}")
            frappe.log_error(frappe.get_traceback(), f"Import April 2026 – {old_name}")

    # ── Print summary ────────────────────────────────────────────────────────
    log()
    log("=" * 70)
    log("  SUMMARY")
    log("=" * 70)
    log(f"  {'Would create' if dry else 'Created'}       : {summary['created']} statements")
    log(f"  Already exist    : {len(summary['skipped_exists'])} skipped")
    log(f"  Fuzzy skipped    : {len(summary['stockist_fuzzy_skipped'])} (need manual mapping)")
    log(f"  No match         : {len(summary['stockist_no_match'])} (not in master)")
    log(f"  Unmapped products: {len(summary['product_no_match'])} rows")

    if summary["skipped_exists"]:
        log()
        log("  Already-existing statements (skipped):")
        for doc_name, sname in summary["skipped_exists"]:
            log(f"    {doc_name:<35}  {sname}")

    if summary["stockist_fuzzy_skipped"]:
        log()
        log("  Fuzzy-skipped stockists  (add correct entry to Stockist Master, then re-run):")
        log(f"  {'OLD NAME IN EXCEL':<55} {'OLD CODE':<10} {'NEAREST MASTER':<55} SCORE")
        log("  " + "-" * 130)
        for old, code, nearest, score in sorted(summary["stockist_fuzzy_skipped"], key=lambda x: -x[3]):
            log(f"  {old:<55} {code:<10} {nearest:<55} {score:.0%}")

    if summary["stockist_no_match"]:
        log()
        log("  No-match stockists  (add to Stockist Master, then re-run):")
        for name, code in summary["stockist_no_match"]:
            log(f"    [{code}]  {name}")

    if summary["product_no_match"]:
        log()
        log("  Unmapped products  (add to Product Master, then re-run):")
        seen = set()
        for code, name, stk in summary["product_no_match"]:
            key = (code, name)
            if key not in seen:
                seen.add(key)
                log(f"    [{code}]  {name}")

    log()
    log("=" * 70)

    # ── Write log file ───────────────────────────────────────────────────────
    try:
        with open(LOG_FILE, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        print(f"\nLog written to: {LOG_FILE}")
    except Exception as e:
        print(f"WARNING: could not write log file — {e}")

    return {
        "created"        : summary["created"],
        "skipped_exists" : len(summary["skipped_exists"]),
        "fuzzy_skipped"  : len(summary["stockist_fuzzy_skipped"]),
        "no_match"       : len(summary["stockist_no_match"]),
        "unmapped_products": len(summary["product_no_match"]),
    }

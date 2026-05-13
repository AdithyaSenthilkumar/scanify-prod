import frappe
import json
from frappe import _
from frappe.utils import flt, nowdate, add_months, get_first_day
import requests
import os
import base64
import mimetypes
from frappe import _
from google import genai as genai_sdk
from google.genai import types as genai_types
from PIL import Image
import io
import zipfile
import tempfile
import time
from frappe.utils import flt, cstr
from frappe.utils.background_jobs import enqueue
import re
from difflib import SequenceMatcher

def get_gemini_settings():
    """
    Fetch Gemini API settings from Scanify Settings DocType
    Returns: tuple (api_key, model_name, is_enabled)
    """
    try:
        # Try to get settings - single doctype
        settings_name = frappe.db.get_value("Scanify Settings", {"company_name": "Stedman Pharmaceuticals"}, "name")
        
        if not settings_name:
            # Fallback: try getting the single record directly
            settings_name = "Scanify Settings"
        
        # Fetch settings data
        settings_data = frappe.db.get_value(
            "Scanify Settings",
            settings_name,
            ["enable_gemini", "gemini_model_name"],
            as_dict=True
        )
        
        if not settings_data:
            frappe.throw(_("Scanify Settings not found. Please create it first."))
        
        # Check if Gemini is enabled
        if not settings_data.get("enable_gemini"):
            frappe.throw(_("Gemini AI extraction is not enabled in Scanify Settings"))
        
        # Get API key
        api_key = frappe.utils.password.get_decrypted_password(
            "Scanify Settings",
            settings_name,
            "gemini_api_key"
        )
        
        if not api_key:
            frappe.throw(_("Gemini API key not configured in Scanify Settings"))
        
        # Get model name with fallback
        model_name = settings_data.get("gemini_model_name") or "gemini-2.5-flash"
        
        frappe.logger().info(f"✅ Gemini settings loaded: Model={model_name}")
        
        return api_key, model_name, True
        
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Gemini Settings Error")
        frappe.throw(_("Error fetching Gemini settings: {0}").format(str(e)))

        
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Gemini Settings Error")
        frappe.throw(_("Error fetching Gemini settings: {0}").format(str(e)))

def build_product_catalog_for_prompt():
    """
    Build comprehensive product catalog with all matching hints for Gemini
    Returns formatted string for prompt inclusion
    """
    products = frappe.get_all(
        "Product Master",
        filters={"status": "Active"},
        fields=[
            "product_code",
            "product_name",
            "pack",
            "pack_conversion",
            "division",
            "product_group",
            "pts",
            "ptr",
            "mrp"
        ],
        order_by="division, product_group, product_name"
    )
    
    if not products:
        frappe.throw("No active products found in Product Master")
    
    # Group by division for better organization
    catalog_text = "\n=== PRODUCT MASTER CATALOG ===\n"
    catalog_text += f"Total Products: {len(products)}\n\n"
    
    current_division = None
    current_group = None
    
    for p in products:
        # Division header
        if p.get("division") != current_division:
            current_division = p.get("division")
            catalog_text += f"\n--- {current_division} Division ---\n"
        
        # Group header
        if p.get("product_group") != current_group:
            current_group = p.get("product_group")
            catalog_text += f"\n  [{current_group} Group]\n"
        
        # Product entry with all matching hints
        catalog_text += f"  • Code: {p['product_code']}\n"
        catalog_text += f"    Name: {p['product_name']}\n"
        catalog_text += f"    Pack: {p.get('pack', 'N/A')}\n"
        catalog_text += f"    Conversion: {p.get('pack_conversion', 'N/A')}\n"
        catalog_text += f"    PTS: {p.get('pts', 0)}\n\n"
    
    return catalog_text, products

@frappe.whitelist()
def extract_stockist_statement(doc_name, file_url):
    """
    Extract stockist statement data using Gemini AI.
    Runs the heavy Gemini call in a background thread to avoid nginx 504 timeouts.
    Frontend should poll check_extraction_status() for completion.
    """
    import threading

    try:
        doc = frappe.get_doc("Stockist Statement", doc_name)

        if doc.extracted_data_status == "In Progress":
            return {"success": True, "message": "Extraction already in progress", "async": True}

        if not file_url:
            return {"success": False, "message": "No file uploaded"}

        # Validate file exists before spawning thread
        from frappe.utils.file_manager import get_file_path
        file_path = get_file_path(file_url)
        if not os.path.exists(file_path):
            return {"success": False, "message": f"File not found: {file_url}"}

        # Mark as In Progress immediately
        doc.extracted_data_status = "In Progress"
        doc.extraction_notes = ""
        doc.save()
        frappe.db.commit()

        # Capture context for background thread
        site = frappe.local.site

        def _run_extraction():
            try:
                frappe.init(site=site)
                frappe.connect()
                _do_extract(doc_name, file_url)
            except Exception as thread_err:
                try:
                    frappe.init(site=site)
                    frappe.connect()
                    _doc = frappe.get_doc("Stockist Statement", doc_name)
                    _doc.extracted_data_status = "Failed"
                    _doc.extraction_notes = f"Extraction failed: {str(thread_err)}"
                    _doc.save()
                    frappe.db.commit()
                except Exception:
                    pass
                frappe.log_error(frappe.get_traceback(), "Gemini Extraction Thread Error")
            finally:
                try:
                    frappe.destroy()
                except Exception:
                    pass

        t = threading.Thread(target=_run_extraction, daemon=True, name=f"extract_{doc_name}")
        t.start()

        return {"success": True, "message": "Extraction started", "async": True}

    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Gemini Extraction Error")
        return {"success": False, "message": f"Extraction failed: {str(e)}"}


def _build_correction_map(stockist_code):
    """Build {RAW_PRODUCT_NAME: product_code} map from Stockist Product Correction for a stockist."""
    corrections = frappe.get_all(
        "Stockist Product Correction",
        filters={"stockist_code": stockist_code, "status": "Active"},
        fields=["raw_product_name", "mapped_product_code"],
    )
    return {c["raw_product_name"].strip().upper(): c["mapped_product_code"] for c in corrections}


def _build_correction_prompt(stockist_code):
    """Build stockist-specific correction hints for Gemini without overriding its final decision."""
    correction_map = _build_correction_map(stockist_code)
    if not correction_map:
        return ""

    lines = ["\n=== STOCKIST-SPECIFIC PRODUCT CORRECTION HINTS ==="]
    lines.append(
        "These are past QC-approved mappings for this stockist. Use them as strong hints only when the visible"
        " product text and pack still agree with the mapped product."
    )
    for raw_name, product_code in sorted(correction_map.items()):
        lines.append(f"- {raw_name} -> {product_code}")

    return "\n".join(lines) + "\n"


def _get_first_present_value(data, *keys):
    """Return the first non-empty value for the given keys."""
    for key in keys:
        if key in data and data.get(key) not in (None, ""):
            return data.get(key)
    return 0


def _parse_numeric_value(value):
    """Parse extracted numeric text while preserving negatives and parenthesized values."""
    text = cstr(value).strip()
    if not text:
        return 0

    if text.startswith("(") and text.endswith(")"):
        text = f"-{text[1:-1]}"

    text = text.replace(",", "")
    return flt(text)


def _build_gemini_generation_config(model_name):
    """Tune Gemini thinking per model family for extraction requests."""
    model_key = cstr(model_name).strip().lower()
    thinking_config = None

    if "gemini-3" in model_key:
        thinking_config = genai_types.ThinkingConfig(thinking_level="high")
    else:
        # Gemini 2.5 models use thinkingBudget. Bump it modestly above the previous fixed 1024 cap.
        thinking_config = genai_types.ThinkingConfig(thinking_budget=2048)

    return genai_types.GenerateContentConfig(thinking_config=thinking_config)


def _normalize_row_type(row_type, raw_product_name):
    """Normalize Gemini row type output and backfill known special rows from the raw label."""
    normalized = cstr(row_type).strip().lower().replace(" ", "_")
    if normalized in {"product", "others", "branch_transfer"}:
        return normalized

    raw_name = cstr(raw_product_name).strip().upper()
    if raw_name == "OTHERS" or raw_name.startswith("OTHERS "):
        return "others"
    if raw_name == "BRANCH TRANSFER" or raw_name.startswith("BRANCH TRANSFER"):
        return "branch_transfer"

    return "product"


def _build_division_product_codes(statement_division, products_list=None):
    """Build the valid product-code set for a statement division."""
    division_product_codes = set()
    if not statement_division:
        return division_product_codes

    source_products = products_list
    if source_products is None:
        source_products = frappe.get_all(
            "Product Master",
            filters={"status": "Active"},
            fields=["product_code", "division"],
        )

    for product in source_products:
        if product.get("division") in (statement_division, "Both"):
            division_product_codes.add(product["product_code"])

    return division_product_codes


def _build_statement_item_row(item_data, statement_division=None, division_product_codes=None):
    """Normalize one Gemini row into a Stockist Statement child row plus bookkeeping metadata."""
    raw_name = cstr(item_data.get("raw_product_name")).strip().upper()
    row_type = _normalize_row_type(item_data.get("row_type"), raw_name)
    mapping_basis = cstr(item_data.get("mapping_basis")).strip().lower()

    is_unmapped = bool(item_data.get("unmapped", False))
    product_code = cstr(item_data.get("product_code")).strip() or None
    if row_type != "product":
        product_code = None
        is_unmapped = False

    if product_code and statement_division and division_product_codes and product_code not in division_product_codes:
        return None, {
            "unmapped": False,
            "auto_mapped": False,
            "special_row": False,
            "skipped_division": True,
        }

    mapping_status = "matched"
    if row_type == "product":
        if is_unmapped or not product_code:
            mapping_status = "unmapped"
        elif mapping_basis == "stockist_correction_hint":
            mapping_status = "auto_mapped"

    row_confidence = min(max(_parse_numeric_value(item_data.get("confidence")), 0), 100)
    operational_sales_qty = _parse_numeric_value(item_data.get("operational_sales_qty"))
    product_sales_qty = _parse_numeric_value(item_data.get("sales_qty"))

    if row_type != "product":
        if not operational_sales_qty:
            operational_sales_qty = product_sales_qty
        product_sales_qty = 0

    row = {
        "raw_product_name": raw_name,
        "row_type": row_type,
        "mapping_status": mapping_status,
        "opening_qty": _parse_numeric_value(item_data.get("opening_qty")),
        "purchase_qty": _parse_numeric_value(item_data.get("purchase_qty")),
        "sales_qty": product_sales_qty,
        "operational_sales_qty": operational_sales_qty,
        "free_qty": _parse_numeric_value(item_data.get("free_qty")),
        "return_qty": _parse_numeric_value(item_data.get("return_qty")),
        "misc_out_qty": _parse_numeric_value(item_data.get("misc_out_qty")),
        "closing_qty": _parse_numeric_value(item_data.get("closing_qty")),
        "closing_value": _parse_numeric_value(item_data.get("closing_value")),
        "row_confidence": round(row_confidence, 1),
        "math_check": "N/A",
    }
    if product_code:
        row["product_code"] = product_code

    return row, {
        "unmapped": mapping_status == "unmapped",
        "auto_mapped": mapping_status == "auto_mapped",
        "special_row": row_type != "product",
        "skipped_division": False,
    }


def _calculate_confidence_score(statement_rows):
    """Average row-level extraction confidence without mixing in QC penalties."""
    confidence_values = [flt(row.get("row_confidence")) for row in statement_rows]
    return round(sum(confidence_values) / len(confidence_values), 1) if confidence_values else 0


def _build_extraction_notes(items_added, confidence_score, unmapped_count=0, auto_mapped_count=0, special_row_count=0,
    skipped_division_count=0, statement_division=None):
    """Build consistent extraction notes for single and bulk flows."""
    notes_parts = [f"Successfully extracted {items_added} rows using AI with product catalog"]
    notes_parts.append(f"Extraction confidence: {confidence_score}%")
    if unmapped_count:
        notes_parts.append(f"{unmapped_count} product rows need QC mapping")
    if auto_mapped_count:
        notes_parts.append(f"{auto_mapped_count} correction-assisted rows need review")
    if special_row_count:
        notes_parts.append(f"{special_row_count} special rows were preserved outside product mapping")
    if skipped_division_count:
        notes_parts.append(f"{skipped_division_count} rows skipped (not in {statement_division} division)")
    return ". ".join(notes_parts)


def _build_statement_rows(extracted_data, statement_division=None, products_list=None):
    """Convert Gemini response rows into statement child rows and summary counts."""
    division_product_codes = _build_division_product_codes(statement_division, products_list)
    statement_rows = []
    counts = {
        "unmapped_count": 0,
        "auto_mapped_count": 0,
        "special_row_count": 0,
        "skipped_division_count": 0,
    }

    for item_data in extracted_data:
        row, metadata = _build_statement_item_row(
            item_data,
            statement_division=statement_division,
            division_product_codes=division_product_codes,
        )
        if metadata.get("skipped_division"):
            counts["skipped_division_count"] += 1
            continue

        if metadata.get("unmapped"):
            counts["unmapped_count"] += 1
        if metadata.get("auto_mapped"):
            counts["auto_mapped_count"] += 1
        if metadata.get("special_row"):
            counts["special_row_count"] += 1

        statement_rows.append(row)

    return statement_rows, counts


def _replace_statement_items(doc, statement_rows):
    """Replace a statement's items with normalized extraction rows."""
    doc.items = []
    for row in statement_rows:
        doc.append("items", row)


def _do_extract(doc_name, file_url):
    """
    Actual extraction logic, runs inside a background thread with its own DB connection.
    """
    api_key, model_name, is_enabled = get_gemini_settings()
    genai_client = genai_sdk.Client(api_key=api_key)

    doc = frappe.get_doc("Stockist Statement", doc_name)

    from frappe.utils.file_manager import get_file_path
    file_path = get_file_path(file_url)

    # Build product catalog
    product_catalog, products_list = build_product_catalog_for_prompt()

    # Extract with enhanced prompt
    extracted_data = call_gemini_extraction_with_catalog(
        file_path,
        doc.stockist_code,
        product_catalog,
        products_list,
        model_name,
        genai_client
    )

    if not extracted_data or len(extracted_data) == 0:
        doc.extracted_data_status = "Failed"
        doc.extraction_notes = "No data extracted - AI returned empty results"
        doc.save()
        frappe.db.commit()
        return

    statement_rows, counts = _build_statement_rows(
        extracted_data,
        statement_division=doc.division,
        products_list=products_list,
    )
    _replace_statement_items(doc, statement_rows)
    items_added = len(statement_rows)

    if items_added == 0:
        doc.extracted_data_status = "Failed"
        doc.extraction_notes = "No valid products found in extracted data"
        doc.save()
        frappe.db.commit()
        return

    doc.confidence_score = _calculate_confidence_score(statement_rows)

    doc.extracted_data_status = "Completed"
    doc.extraction_notes = _build_extraction_notes(
        items_added,
        doc.confidence_score,
        unmapped_count=counts["unmapped_count"],
        auto_mapped_count=counts["auto_mapped_count"],
        special_row_count=counts["special_row_count"],
        skipped_division_count=counts["skipped_division_count"],
        statement_division=doc.division,
    )
    doc.populate_previous_month_closing()
    doc.calculate_closing_and_totals()
    doc.calculate_qc_confidence()
    doc.save()
    frappe.db.commit()


@frappe.whitelist()
def check_extraction_status(doc_name):
    """Poll endpoint: returns current extraction status for a Stockist Statement."""
    status = frappe.db.get_value(
        "Stockist Statement", doc_name,
        ["extracted_data_status", "extraction_notes"],
        as_dict=True
    )
    if not status:
        return {"success": False, "message": "Document not found"}
    return {
        "success": True,
        "status": status.extracted_data_status,
        "notes": status.extraction_notes or ""
    }


@frappe.whitelist()
def save_extracted_statement(doc_name, data):
    """
    Save the extracted and QC-verified data, then immediately submit the statement
    (docstatus → 1). This is a one-way finalisation — no edits after this point.
    Called from the portal after extraction + optional QC edits.
    data: JSON array of product rows (same structure as extracted_data in JS)
    """
    try:
        if isinstance(data, str):
            data = json.loads(data)

        doc = frappe.get_doc("Stockist Statement", doc_name)

        if doc.docstatus == 1:
            frappe.throw(_("This statement has already been submitted and cannot be changed."))

        # Replace items in doc
        doc.items = []
        for row in data:
            product_code = row.get("productcode") or row.get("product_code") or None
            mapping_status = row.get("mapping_status") or row.get("mappingstatus") or "matched"
            raw_product_name = row.get("raw_product_name") or row.get("rawproductname") or ""
            row_type = _normalize_row_type(row.get("rowtype") or row.get("row_type"), raw_product_name)
            sales_qty = _parse_numeric_value(_get_first_present_value(row, "salesqty", "sales_qty"))
            operational_sales_qty = _parse_numeric_value(_get_first_present_value(row, "operationalsalesqty", "operational_sales_qty"))

            if row_type != "product":
                product_code = None
                mapping_status = "matched"
                if not operational_sales_qty:
                    operational_sales_qty = sales_qty
                sales_qty = 0

            # Unmapped items are kept in the statement but won't contribute to totals
            item_dict = {
                "raw_product_name": raw_product_name,
                "row_type": row_type,
                "mapping_status": mapping_status,
                "opening_qty": _parse_numeric_value(_get_first_present_value(row, "openingqty", "opening_qty")),
                "purchase_qty": _parse_numeric_value(_get_first_present_value(row, "purchaseqty", "purchase_qty")),
                "sales_qty": sales_qty,
                "operational_sales_qty": operational_sales_qty,
                "free_qty": _parse_numeric_value(_get_first_present_value(row, "freeqty", "free_qty")),
                "free_qty_scheme": _parse_numeric_value(_get_first_present_value(row, "freeqtyscheme", "free_qty_scheme")),
                "return_qty": _parse_numeric_value(_get_first_present_value(row, "returnqty", "return_qty")),
                "misc_out_qty": _parse_numeric_value(_get_first_present_value(row, "miscoutqty", "misc_out_qty")),
                "closing_qty": _parse_numeric_value(_get_first_present_value(row, "closingqty", "closing_qty")),
                "closing_value": _parse_numeric_value(_get_first_present_value(row, "closingvalue", "closing_value")),
                "row_confidence": _parse_numeric_value(_get_first_present_value(row, "row_confidence", "confidence")),
            }
            if product_code:
                item_dict["product_code"] = product_code
            doc.append("items", item_dict)

        doc.extracted_data_status = "Completed"
        doc.calculate_closing_and_totals()
        # Save first (persist items)
        doc.save(ignore_permissions=True)
        frappe.db.commit()
        # Submit — sets docstatus = 1, locks the document permanently
        doc.submit()
        frappe.db.commit()

        return {
            "success": True,
            "message": f"Statement submitted with {len(doc.items)} items.",
            "doc_name": doc.name
        }

    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Save Extracted Statement Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def save_draft_statement(doc_name, data):
    """
    Persist the current items of a draft Stockist Statement without submitting.
    Called automatically from the portal QC screen on every edit (debounced).
    """
    try:
        if isinstance(data, str):
            data = json.loads(data)

        doc = frappe.get_doc("Stockist Statement", doc_name)

        if doc.docstatus == 1:
            return {"success": False, "message": "Statement already submitted."}

        doc.items = []
        for row in data:
            product_code = row.get("productcode") or row.get("product_code") or None
            mapping_status = row.get("mapping_status") or row.get("mappingstatus") or "matched"
            raw_product_name = row.get("raw_product_name") or row.get("rawproductname") or ""
            row_type = _normalize_row_type(row.get("rowtype") or row.get("row_type"), raw_product_name)
            sales_qty = _parse_numeric_value(_get_first_present_value(row, "salesqty", "sales_qty"))
            operational_sales_qty = _parse_numeric_value(_get_first_present_value(row, "operationalsalesqty", "operational_sales_qty"))

            if row_type != "product":
                product_code = None
                mapping_status = "matched"
                if not operational_sales_qty:
                    operational_sales_qty = sales_qty
                sales_qty = 0

            item_dict = {
                "raw_product_name": raw_product_name,
                "row_type": row_type,
                "mapping_status": mapping_status,
                "opening_qty": _parse_numeric_value(_get_first_present_value(row, "openingqty", "opening_qty")),
                "purchase_qty": _parse_numeric_value(_get_first_present_value(row, "purchaseqty", "purchase_qty")),
                "sales_qty": sales_qty,
                "operational_sales_qty": operational_sales_qty,
                "free_qty": _parse_numeric_value(_get_first_present_value(row, "freeqty", "free_qty")),
                "free_qty_scheme": _parse_numeric_value(_get_first_present_value(row, "freeqtyscheme", "free_qty_scheme")),
                "return_qty": _parse_numeric_value(_get_first_present_value(row, "returnqty", "return_qty")),
                "misc_out_qty": _parse_numeric_value(_get_first_present_value(row, "miscoutqty", "misc_out_qty")),
                "closing_qty": _parse_numeric_value(_get_first_present_value(row, "closingqty", "closing_qty")),
                "closing_value": _parse_numeric_value(_get_first_present_value(row, "closingvalue", "closing_value")),
                "row_confidence": _parse_numeric_value(_get_first_present_value(row, "row_confidence", "confidence")),
            }
            if product_code:
                item_dict["product_code"] = product_code
            doc.append("items", item_dict)

        doc.calculate_closing_and_totals()
        doc.save(ignore_permissions=True)
        frappe.db.commit()

        return {"success": True, "message": f"Draft saved ({len(doc.items)} items)."}

    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Save Draft Statement Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def get_statement_for_view(doc_name):
    """
    Return full statement data for the portal view page.
    """
    try:
        doc = frappe.get_doc("Stockist Statement", doc_name)
        # Build items list enriched with product master info
        items = []
        for item in doc.items:
            product = {}
            row_type = item.row_type or _normalize_row_type(None, item.raw_product_name)
            operational_sales_qty = flt(item.operational_sales_qty or (row_type != "product" and item.sales_qty or 0))
            sales_qty = 0 if row_type != "product" else flt(item.sales_qty)
            if item.product_code:
                product = frappe.db.get_value(
                    "Product Master", item.product_code,
                    ["product_name", "pts", "ptr", "pack"], as_dict=True
                ) or {}
            items.append({
                "productcode": item.product_code or "",
                "productname": item.product_name or product.get("product_name", ""),
                "rawproductname": item.raw_product_name or "",
                "rowtype": row_type,
                "mappingstatus": item.mapping_status or "matched",
                "pack": item.pack or product.get("pack", ""),
                "pts": flt(item.pts or product.get("pts", 0)),
                "ptr": flt(product.get("ptr", 0)),
                "conversion_factor": flt(item.conversion_factor or 1),
                "openingqty": flt(item.opening_qty),
                "purchaseqty": flt(item.purchase_qty),
                "salesqty": sales_qty,
                "operationalsalesqty": operational_sales_qty,
                "freeqty": flt(item.free_qty),
                "freeqtyscheme": flt(item.free_qty_scheme),
                "returnqty": flt(item.return_qty),
                "miscoutqty": flt(item.misc_out_qty),
                "closingqty": flt(item.closing_qty),
                "closingvalue": flt(item.closing_value),
                "openingvalue": flt(item.opening_value),
                "purchasevalue": flt(item.purchase_value),
                "salesvaluepts": flt(item.sales_value_pts),
                "salesvalueptr": flt(item.sales_value_ptr),
                "schemedeductedqty": flt(item.scheme_deducted_qty_calc),
                "row_confidence": flt(item.row_confidence),
                "math_check": item.math_check or "N/A",
            })

        return {
            "success": True,
            "doc": {
                "name": doc.name,
                "stockist_code": doc.stockist_code,
                "stockist_name": doc.stockist_name,
                "statement_month": str(doc.statement_month),
                "hq": doc.hq,
                "team": doc.team,
                "region": doc.region,
                "zone": doc.zone,
                "extracted_data_status": doc.extracted_data_status,
                "uploaded_file": doc.uploaded_file,
                "docstatus": doc.docstatus,
                "qc_confidence": doc.qc_confidence or "",
                "confidence_score": flt(doc.confidence_score),
                "total_opening_value": flt(doc.total_opening_value),
                "total_purchase_value": flt(doc.total_purchase_value),
                "total_operational_sales_qty": flt(doc.total_operational_sales_qty),
                "total_closing_value": flt(doc.total_closing_value),
                "total_sales_value_pts": flt(doc.total_sales_value_pts),
                "total_sales_value_ptr": flt(doc.total_sales_value_ptr),
                "division": doc.division or "",
            },
            "items": items
        }
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Get Statement For View Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def get_primary_sales_for_stockist(stockist_code, statement_month):
    """
    Fetch Primary Sales Data rows for a given stockist and month.
    statement_month can be 'YYYY-MM-DD' or 'YYYY-MM'; we extract YYYY-MM.
    """
    if not stockist_code or not statement_month:
        return {"success": False, "message": "Stockist code and month are required."}

    # Normalise to YYYY-MM then derive first/last day of that month
    month_str = str(statement_month)[:7]          # "2026-02"
    from frappe.utils import getdate, get_last_day
    first_day = getdate(month_str + "-01")
    last_day = get_last_day(first_day)

    rows = frappe.get_all(
        "Primary Sales Data",
        filters={
            "stockist_code": stockist_code,
            "invoicedate": ["between", [first_day, last_day]],
            "iscancelled": 0,
        },
        fields=[
            "invoiceno", "invoicedate", "pcode", "product", "pack",
            "batchno", "quantity", "freeqty", "pts", "ptsvalue",
            "ptr", "ptrvalue",
        ],
        order_by="invoicedate asc, pcode asc",
        limit_page_length=0,
    )

    return {"success": True, "rows": rows, "month": month_str}


def call_gemini_extraction_with_catalog(file_path, stockist_code, product_catalog, products_list, model_name=None, genai_client=None):
    """
    Enhanced extraction that sends the full product catalog to Gemini
    Gemini does the matching directly using product codes
    """
    if not genai_client:
        api_key, model_name, is_enabled = get_gemini_settings()
        genai_client = genai_sdk.Client(api_key=api_key)
    
    try:
        # Determine file type
        mime_type, _ = mimetypes.guess_type(file_path)
        file_ext = os.path.splitext(file_path)[1].lower()
        correction_prompt = _build_correction_prompt(stockist_code)
        correction_map = _build_correction_map(stockist_code)
        generation_config = _build_gemini_generation_config(model_name)
        
        # Enhanced prompt with product catalog
        prompt = f"""You are extracting pharmaceutical stockist statement data for STEDMAN PHARMACEUTICALS.

{product_catalog}
    {correction_prompt}

=== EXTRACTION RULES ===

1. COLUMN DETECTION (DO THIS FIRST — CRITICAL):
   - FIRST: Read ALL column headers in the document from left to right.
   - Map each header to one of the standard fields below:
     . Opening Qty: "Op.Qty", "OPSTK", "Opening", "Open.Qty", "QpnStk", "Op.Stk"
     . Purchase Qty: "Purch.Qty", "PURCH", "Receipt", "Pr.Qty", "Pur", "Recv"
     . Sales Qty: "Sales", "Sale", "Sl", "Sold", "Sales Qty"
     . Free Qty: "Free Qty", "Free", "Scheme Qty", "Fre"
     . Return Qty: "Return", "Ret", "Sales Ret", "SR", "Sal.Ret"
     . Misc Out: "Misc.Out", "M.Out", "Transfer", "Trans", "Adj"
     . Closing Qty: "Closing", "Cls", "Cl.Bal", "ClsStk", "Closing Qty", "Balance", "Bal"
     . Closing Value: "Closing Value", "Closing Val", "Cl.Value", "Closing Amount", "Cls.Val"
   - IMPORTANT: Distinguish QUANTITY columns (integers) from VALUE/AMOUNT columns (decimals with currency).
     Do NOT extract value columns as quantities. Value columns typically have large decimal numbers.
   - IGNORE previous-month or historical columns: PMSalesQty, Prev.Cls, LM Sales, Last Month, PM Cls, etc.
   - IMPORTANT: "AVSL" is NOT an Opening column. It typically refers to a scheme/free goods allocation or a sales benchmark — do NOT map it to opening_qty.
   - VERIFY: For a few sample rows, check that Opening + Purchase - Sales - Free - Return - MiscOut ≈ Closing.
     If this does NOT add up, re-examine your column assignments — you may have the wrong column mapped.

2. PRODUCT MATCHING (CRITICAL — STRICT MATCHING ONLY):
   - Match each product ONLY when you are confident. Use ALL of these together:
     a) Product NAME must closely match a catalog entry
     b) Pack SIZE must be consistent with the catalog
     c) Pack CONVERSION must be consistent
   - Be aware of common OCR misreads: G↔C, O↔0, I↔1↔L, 5↔S, 8↔B, rn↔m, cl↔d
     Example: "GALCIGEN" might be OCR misread of "CALCIGEN" — use pack size to disambiguate.
   - If a product name is ambiguous or could match multiple catalog entries, mark it UNMAPPED rather than guessing.
   - Return ALL products from the statement, including those you CANNOT match.
   - For matched products: set "product_code" to the EXACT catalog code and "unmapped" to false
   - For unmatched products: set "product_code" to null and "unmapped" to true
    - If you rely on a stockist-specific correction hint and it agrees with the visible row text and pack, set
      "mapping_basis" to "stockist_correction_hint".

2A. SPECIAL ROWS (CRITICAL):
    - Preserve operational rows such as "OTHERS" and "BRANCH TRANSFER".
    - Do NOT force these rows into Product Master mapping.
    - Set "row_type" to one of: "product", "others", "branch_transfer".
    - For special rows, set "product_code" to null, "unmapped" to false, and "mapping_basis" to "special_row".
    - For special rows, copy the printed movement quantity into "operational_sales_qty" and set "sales_qty" to 0.
    - For normal product rows, set "operational_sales_qty" to 0.

3. QUANTITY EXTRACTION (CRITICAL):
   - Extract quantities for all identified columns AND closing value if present.
   - ONLY extract values for columns that ACTUALLY EXIST in the statement. Do NOT calculate or infer missing columns.
   - If a column does not exist in the statement, set its value to 0. Do NOT compute it from other columns.
   - If Closing Qty column exists in the statement, extract the printed value directly.
   - If Closing Qty column does NOT exist in the statement, set closing_qty to 0. Do NOT calculate it as Opening + Purchase - Sales - etc.
   - If Closing Value column exists in the statement, extract the printed value directly.
   - If Closing Value column does NOT exist in the statement, set closing_value to 0. Do NOT calculate it.
    - Preserve negative quantities exactly as printed.
    - If a quantity is shown in parentheses, output it as a negative number.
   - SELF-CHECK each row: For columns that DO exist, verify Opening + Purchase - Sales - Free - Return - MiscOut should ≈ Closing.
     If a row's math is way off, double-check which column values you assigned.

4. EXTRACTION CONFIDENCE (CRITICAL — per row):
   - For EACH row, assign a "confidence" integer from 0 to 100 representing how confident you are that the extracted data is correct.
   - Score 100: You can clearly read ALL fields in the row, the values make sense, and math sanity checks out. USE 100 LIBERALLY — if you can read it, it's 100.
   - Score 70-99: One or more values are ambiguous — a blurry digit, faded text, overlapping columns, or the math doesn't add up (likely an OCR misread).
   - Score 30-69: Several fields are hard to read, the row layout is confusing, or quantities seem very unusual.
   - Score 0-29: Very poor readability, heavy guessing, or the row is heavily damaged/truncated.
   - For unmapped products: still score based on how well you could READ the quantities (mapping is separate).
   - Sanity check: if Opening + Purchase - Sales - Free - Return - MiscOut does NOT equal Closing and the difference is significant, REDUCE confidence below 100 (the OCR likely got a digit wrong).
   - If some columns are genuinely absent from the statement (e.g. no Free or Return column), that is NOT a reason to reduce confidence — only reduce when you are uncertain about what you read.
   - IMPORTANT: Do NOT artificially lower confidence. If you can read a row clearly and the math checks out, it MUST be 100.

5. OUTPUT FORMAT:
   - Return ONLY a valid JSON array — no markdown, no explanation, no extra text.
   - Include ONLY quantities (NO price/value columns except closing_value).
   - Include ALL products from the statement, matched and unmatched.

EXPECTED JSON FORMAT:
[
  {{
    "product_code": "ARC",
    "raw_product_name": "ARCALION 200MG",
        "row_type": "product",
        "mapping_basis": "catalog_exact",
    "unmapped": false,
    "confidence": 98,
    "opening_qty": 88,
    "purchase_qty": 29,
    "sales_qty": 59,
    "operational_sales_qty": 0,
    "free_qty": 0,
    "return_qty": 0,
    "misc_out_qty": 0,
    "closing_qty": 58,
    "closing_value": 500.00
  }},
  {{
    "product_code": null,
        "raw_product_name": "BRANCH TRANSFER",
        "row_type": "branch_transfer",
        "mapping_basis": "special_row",
        "unmapped": false,
        "confidence": 92,
    "opening_qty": 10,
    "purchase_qty": 5,
    "sales_qty": 0,
    "operational_sales_qty": 8,
    "free_qty": 0,
    "return_qty": 0,
    "misc_out_qty": 0,
    "closing_qty": 7,
    "closing_value": 0
  }}
]

IMPORTANT:
- NO values/prices in output (except closing_value)
- NO markdown formatting
- ONLY valid JSON array
- Always include row_type and mapping_basis for every row
- Use EXACT product codes from catalog for matched products
- Set product_code to null for unmatched products
- Always include raw_product_name for ALL products
- Always include confidence (0-100) for ALL products
- Always include operational_sales_qty for every row
"""
        
        frappe.logger().info(f"Using Gemini model: {model_name}")

        # Retry logic for rate limiting (429) and unavailable (503)
        max_retries = 3
        base_delay = 2
        response = None
        current_model = model_name
        # Fallback model when primary is overloaded (503)
        fallback_model = "gemini-3-flash-preview"
        attempt = 0
        used_fallback = False

        while attempt < max_retries:
            try:
                if file_ext in [".pdf", ".jpg", ".jpeg", ".png"]:
                    with open(file_path, "rb") as f:
                        file_data = f.read()

                    if file_ext == ".pdf":
                        file_part = genai_types.Part.from_bytes(
                            data=file_data, mime_type="application/pdf"
                        )
                    else:
                        file_part = genai_types.Part.from_bytes(
                            data=file_data, mime_type=mime_type or "image/jpeg"
                        )

                    response = genai_client.models.generate_content(
                        model=current_model,
                        contents=[prompt, file_part],
                        config=generation_config
                    )

                elif file_ext in [".csv", ".txt"]:
                    with open(file_path, "r", encoding="utf-8") as f:
                        file_content = f.read()
                    response = genai_client.models.generate_content(
                        model=current_model,
                        contents=f"{prompt}\n\nCONTENT:\n{file_content}",
                        config=generation_config
                    )

                elif file_ext in [".xls", ".xlsx"]:
                    import pandas as pd
                    df = pd.read_excel(file_path)
                    file_content = df.to_string()
                    response = genai_client.models.generate_content(
                        model=current_model,
                        contents=f"{prompt}\n\nCONTENT:\n{file_content}",
                        config=generation_config
                    )

                else:
                    frappe.throw(f"Unsupported file type: {file_ext}")

                break  # Success

            except Exception as e:
                err_str = str(e).lower()
                is_rate_limit = "quota" in err_str or "rate" in err_str or "429" in err_str or "resource_exhausted" in err_str
                is_unavailable = "503" in err_str or "unavailable" in err_str or "overloaded" in err_str or "high demand" in err_str

                if (is_rate_limit or is_unavailable) and attempt < max_retries - 1:
                    delay = base_delay * (2 ** attempt)
                    frappe.logger().warning(
                        f"{'503 Unavailable' if is_unavailable else 'Rate limit'} on {current_model} "
                        f"(attempt {attempt + 1}/{max_retries}), retrying in {delay}s..."
                    )
                    time.sleep(delay)
                    attempt += 1
                elif is_unavailable and not used_fallback:
                    # Primary model exhausted retries — switch to fallback
                    frappe.logger().warning(
                        f"{current_model} unavailable after {max_retries} attempts, "
                        f"falling back to {fallback_model}"
                    )
                    current_model = fallback_model
                    used_fallback = True
                    attempt = 0
                    max_retries = 2
                    time.sleep(base_delay)
                elif is_rate_limit or is_unavailable:
                    frappe.throw(
                        f"Gemini API unavailable after retries (tried {model_name}"
                        f"{' and ' + fallback_model if used_fallback else ''}). "
                        f"Please try again later."
                    )
                else:
                    raise e
        
        # Parse response
        try:
            frappe.logger().info("=== GEMINI RAW RESPONSE START ===")
            frappe.logger().info(response.text)
            frappe.logger().info("=== GEMINI RAW RESPONSE END ===")
        except Exception as log_err:
            frappe.logger().error(f"Failed to log Gemini response: {log_err}")
        
        response_text = response.text.strip()
        
        # Clean response
        if response_text.startswith("```"):
            response_text = response_text.split("```", 1)[1]
        if response_text.startswith("json"):
            response_text = response_text[4:]
        if response_text.endswith("```"):
            response_text = response_text.rsplit("```", 1)[0]
        
        response_text = response_text.strip()
        
        frappe.logger().info("=== Cleaned Response Text ===")
        frappe.logger().info(response_text)
        
        # Robust JSON parsing with repair for common Gemini output issues
        extracted_items = None
        try:
            extracted_items = json.loads(response_text)
        except json.JSONDecodeError:
            # Attempt repairs: trailing commas, truncated JSON
            repaired = response_text
            # Remove trailing commas before ] or }
            repaired = re.sub(r',\s*([\]\}])', r'\1', repaired)
            # If JSON array is truncated (no closing ]), try to close it
            if repaired.lstrip().startswith('[') and not repaired.rstrip().endswith(']'):
                # Find last complete object (ending with })
                last_brace = repaired.rfind('}')
                if last_brace > 0:
                    repaired = repaired[:last_brace + 1] + ']'
            try:
                extracted_items = json.loads(repaired)
                frappe.logger().info("JSON parsed after repair")
            except json.JSONDecodeError as je:
                frappe.logger().error(f"JSON repair also failed: {je}")
                frappe.throw(f"Failed to parse AI response as JSON: {je}")
        
        extracted_items = extracted_items or []
        frappe.logger().info(f"Parsed Items Count: {len(extracted_items)}")
        
        # Validate product codes and tag mapping status
        valid_codes = {p["product_code"] for p in products_list}
        validated_items = []
        
        for item in extracted_items:
            raw_name = (item.get("raw_product_name") or "").strip()
            row_type = _normalize_row_type(item.get("row_type"), raw_name)
            mapping_basis = cstr(item.get("mapping_basis")).strip().lower()
            pc = (item.get("product_code") or "").strip() or None
            is_unmapped = item.get("unmapped", False)

            if row_type != "product":
                item["product_code"] = None
                item["raw_product_name"] = raw_name
                item["row_type"] = row_type
                item["mapping_basis"] = mapping_basis or "special_row"
                item["unmapped"] = False
                item["operational_sales_qty"] = _parse_numeric_value(item.get("operational_sales_qty") or item.get("sales_qty"))
                item["sales_qty"] = 0
                validated_items.append(item)
                frappe.logger().info(f"Special row kept: {raw_name} (type={row_type})")
                continue

            if not mapping_basis and raw_name.strip().upper() in correction_map and correction_map[raw_name.strip().upper()] == pc:
                mapping_basis = "stockist_correction_hint"

            if pc and pc in valid_codes:
                # Gemini matched to a valid catalog product
                item["product_code"] = pc
                item["raw_product_name"] = raw_name
                item["row_type"] = row_type
                item["mapping_basis"] = mapping_basis or "catalog_match"
                item["unmapped"] = False
                item["operational_sales_qty"] = _parse_numeric_value(item.get("operational_sales_qty"))
                validated_items.append(item)
            else:
                # Unmatched — keep the item but mark unmapped
                item["product_code"] = None
                item["raw_product_name"] = raw_name
                item["row_type"] = row_type
                item["mapping_basis"] = mapping_basis or "unmapped"
                item["unmapped"] = True
                item["operational_sales_qty"] = _parse_numeric_value(item.get("operational_sales_qty"))
                validated_items.append(item)
                frappe.logger().info(f"Unmapped product kept: {raw_name} (code={pc})")
        
        frappe.logger().info(f"Final Items: {len(validated_items)} (matched: {sum(1 for i in validated_items if not i.get('unmapped'))}, unmapped: {sum(1 for i in validated_items if i.get('unmapped'))})")
        return validated_items
        
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Gemini API Call Error - Enhanced")
        frappe.throw(f"Extraction failed: {str(e)}")


@frappe.whitelist()
def map_filenames_to_stockists_via_gemini(filenames):
    """
    Use Gemini to map a list of filenames to stockist codes in a single call.
    Much more accurate than pure fuzzy matching.
    filenames: JSON string or list of filename strings
    Returns: dict mapping filename -> stockist_code (or null)
    """
    try:
        if isinstance(filenames, str):
            filenames = json.loads(filenames)

        if not filenames:
            return {"success": True, "mapping": {}}

        # Fetch stockist catalog (name + code only to minimize tokens)
        stockists = frappe.get_all(
            "Stockist Master",
            filters={"status": "Active"},
            fields=["name", "stockist_name"],
            order_by="stockist_name asc"
        )

        if not stockists:
            return {"success": False, "message": "No active stockists found"}

        # Build compact catalog
        catalog_lines = [f"{s['name']}|{s['stockist_name']}" for s in stockists]
        catalog_text = "\n".join(catalog_lines)

        filenames_text = "\n".join([f"{i+1}. {f}" for i, f in enumerate(filenames)])

        prompt = f"""You are matching pharmaceutical stockist statement filenames to stockist records.

STOCKIST CATALOG (format: CODE|NAME):
{catalog_text}

FILENAMES TO MATCH:
{filenames_text}

TASK: Match each filename to the most likely stockist code from the catalog above.
- The filename may contain part of the stockist name (possibly misspelled or abbreviated)
- Ignore date parts, month names, years, numbers, extensions
- Look for the distinctive name portion

Return a JSON object where keys are the exact filenames and values are the matched stockist CODE (or null if no good match):
{{
  "filename1.pdf": "CODE1",
  "filename2.jpg": null,
  ...
}}

IMPORTANT:
- Return ONLY valid JSON, no explanation
- Use exact filenames as keys (including extension)
- Use exact stockist CODEs from the catalog as values
- Return null if confidence is low
"""

        api_key, model_name, _ = get_gemini_settings()
        client = genai_sdk.Client(api_key=api_key)

        response = client.models.generate_content(
            model=model_name,
            contents=prompt,
            config=genai_types.GenerateContentConfig(
                thinking_config=genai_types.ThinkingConfig(thinking_budget=0)
            )
        )

        resp_text = response.text.strip()
        if resp_text.startswith("```"):
            resp_text = resp_text.split("```", 1)[1]
        if resp_text.lower().startswith("json"):
            resp_text = resp_text[4:]
        if resp_text.endswith("```"):
            resp_text = resp_text.rsplit("```", 1)[0]
        resp_text = resp_text.strip()

        mapping = json.loads(resp_text)

        # Validate that returned codes actually exist
        valid_codes = {s["name"] for s in stockists}
        validated = {}
        for fname, code in mapping.items():
            if code and code in valid_codes:
                validated[fname] = code
            else:
                validated[fname] = None

        return {"success": True, "mapping": validated}

    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Filename→Stockist Gemini Mapping Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def start_bulk_ocr_job(docname):
    """
    Portal entry point: validates the doc, then runs extraction in a background thread.
    Uses threading instead of RQ enqueue to avoid os.fork() deadlocks in WSL/gunicorn environments.
    The thread gets its own Frappe DB connection via frappe.init()/connect() so it is fully
    independent; progress is written to DB and polled by the frontend every 5 seconds.
    """
    import threading

    try:
        doc = frappe.get_doc("Bulk Statement Upload", docname)

        if doc.status in ("In Progress",):
            return {"success": False, "message": f"Job is already {doc.status}. Cannot restart."}

        if not doc.zip_file:
            return {"success": False, "message": "No ZIP file attached to this job."}

        # Capture all context-specific values before the request context is torn down
        site = frappe.local.site
        zip_file_url = doc.zip_file
        month = str(doc.statement_month)

        # Immediately mark as Queued so the UI updates
        doc.status = "Queued"
        doc.save(ignore_permissions=True)
        frappe.db.commit()

        def run_in_thread():
            """Thread target: initialises its own Frappe connection, runs extraction, then destroys."""
            try:
                frappe.init(site=site)
                frappe.connect()
                process_bulk_extraction(docname=docname, month=month, zip_file_url=zip_file_url)
            except Exception as thread_err:
                # Best-effort: mark doc as Failed with error info
                try:
                    frappe.init(site=site)
                    frappe.connect()
                    _doc = frappe.get_doc("Bulk Statement Upload", docname)
                    _doc.status = "Failed"
                    _doc.extraction_log = json.dumps([
                        {"file": "—", "status": "Failed", "message": str(thread_err)}
                    ])
                    _doc.save(ignore_permissions=True)
                    frappe.db.commit()
                except Exception:
                    pass
            finally:
                try:
                    frappe.destroy()
                except Exception:
                    pass

        t = threading.Thread(target=run_in_thread, daemon=True, name=f"bulk_ocr_{docname}")
        t.start()

        return {"success": True, "message": "Bulk OCR job started in background thread", "job_id": docname}

    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Start Bulk OCR Job Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def get_bulk_job_status(docname):
    """
    Return current status/progress of a bulk job for live polling.
    """
    try:
        doc = frappe.get_doc("Bulk Statement Upload", docname)
        log_data = []
        if doc.extraction_log:
            try:
                log_data = json.loads(doc.extraction_log)
            except Exception:
                log_data = []

        # Enrich log with live QC confidence from actual statements
        statement_names = [r["statement"] for r in log_data if r.get("statement")]
        if statement_names:
            live_data = {
                row.name: row
                for row in frappe.get_all(
                    "Stockist Statement",
                    filters={"name": ["in", statement_names]},
                    fields=["name", "qc_confidence", "confidence_score"],
                )
            }
            for entry in log_data:
                st = entry.get("statement")
                if st and st in live_data:
                    entry["qc_confidence"] = live_data[st].qc_confidence or entry.get("qc_confidence", "")
                    entry["confidence_score"] = flt(live_data[st].confidence_score)

        return {
            "success": True,
            "status": doc.status,
            "progress": flt(doc.progress),
            "total_files": doc.total_files or 0,
            "success_count": doc.success_count or 0,
            "failed_count": doc.failed_count or 0,
            "skipped_count": doc.skipped_count or 0,
            "log": log_data,
        }
    except Exception as e:
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def upload_bulk_zip():
    """
    Custom ZIP upload endpoint for portal users who lack desk access.
    Frappe's built-in upload_file restricts non-desk users to a MIME allowlist
    that excludes application/zip, causing a 417.  This endpoint explicitly
    validates that the caller has write permission on Bulk Statement Upload and
    then saves the file with ignore_permissions=True so the ZIP is accepted.
    """
    if frappe.session.user == "Guest":
        frappe.throw("Authentication required", frappe.AuthenticationError)

    frappe.has_permission("Bulk Statement Upload", "write", throw=True)

    files = frappe.request.files
    if "file" not in files:
        frappe.throw("No file provided")

    uploaded = files["file"]
    filename = uploaded.filename or "bulk_upload.zip"
    if not filename.lower().endswith(".zip"):
        frappe.throw("Only ZIP files are accepted here")

    content = uploaded.stream.read()

    file_doc = frappe.get_doc({
        "doctype": "File",
        "file_name": filename,
        "is_private": 0,
        "content": content,
        "folder": "Home",
    })
    file_doc.save(ignore_permissions=True)
    frappe.db.commit()

    return {"file_url": file_doc.file_url}


@frappe.whitelist()
def get_bulk_jobs_list(division=None):
    """
    Return list of bulk OCR jobs for the portal list page.
    """
    try:
        user_division = division
        if not user_division:
            user_division = frappe.db.get_value("User", frappe.session.user, "division") or "Prima"

        filters = {"docstatus": ["in", [0, 1]]}
        if user_division:
            filters["division"] = user_division

        jobs = frappe.get_all(
            "Bulk Statement Upload",
            filters=filters,
            fields=["name", "statement_month", "status", "progress", "total_files",
                    "success_count", "failed_count", "skipped_count", "creation", "modified"],
            order_by="creation desc",
            limit_page_length=100,
        )

        return {"success": True, "jobs": jobs}

    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Get Bulk Jobs List Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def create_bulk_ocr_job(statement_month, zip_file_url, division=None, job_name=None):
    """
    Create a new Bulk Statement Upload document from portal.
    Returns the created doc name.
    """
    try:
        if not division:
            division = frappe.db.get_value("User", frappe.session.user, "division") or "Prima"

        doc = frappe.get_doc({
            "doctype": "Bulk Statement Upload",
            "statement_month": statement_month,
            "zip_file": zip_file_url,
            "division": division,
            "job_name": job_name or "",
            "status": "Pending",
        })
        doc.insert(ignore_permissions=True)
        frappe.db.commit()
        return {"success": True, "docname": doc.name}
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Create Bulk OCR Job Error")
        return {"success": False, "message": str(e)}


# ========== REST OF THE API FILE (UNCHANGED) ==========
@frappe.whitelist()
def bulk_extract_statements_async(docname):
    """
    Enqueue bulk extraction as background job
    """
    doc = frappe.get_doc("Bulk Statement Upload", docname)
    
    # Enqueue background job
    job = enqueue(
        method="scanify.api.process_bulk_extraction",
        queue="long",
        timeout=3600,  # 1 hour
        job_name=f"bulk_extract_{docname}",
        docname=docname,
        month=doc.statement_month,
        zip_file_url=doc.zip_file
    )
    
    # Update status
    doc.status = "Queued"
    doc.job_id = job.id
    doc.save(ignore_permissions=True)
    frappe.db.commit()
    
    return {
        "success": True,
        "message": "Bulk extraction job queued successfully",
        "job_id": job.id
    }

def process_bulk_extraction(docname, month, zip_file_url):
    """
    Background job to process bulk extraction
    """
    try:
        doc = frappe.get_doc("Bulk Statement Upload", docname)
        doc.status = "In Progress"
        doc.progress = 0
        doc.save(ignore_permissions=True)
        frappe.db.commit()
        
        # Get ZIP file
        from frappe.utils.file_manager import get_file_path
        file_path = get_file_path(zip_file_url)
        
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {zip_file_url}")
        
        # Extract ZIP
        with tempfile.TemporaryDirectory() as temp_dir:
            with zipfile.ZipFile(file_path, "r") as zip_ref:
                zip_ref.extractall(temp_dir)
            
            # Get all files
            all_files = []
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    if file.startswith(".") or file.startswith("__MACOSX") or file == "Thumbs.db":
                        continue
                    
                    file_full_path = os.path.join(root, file)
                    file_ext = os.path.splitext(file)[1].lower()
                    
                    supported_extensions = [".pdf", ".jpg", ".jpeg", ".png", ".csv", ".txt", ".xls", ".xlsx"]
                    if file_ext in supported_extensions:
                        all_files.append((file, file_full_path, file_ext))
            
            doc.total_files = len(all_files)
            doc.save(ignore_permissions=True)
            frappe.db.commit()
            
            # Process each file
            results = []
            success_count = 0
            failed_count = 0
            skipped_count = 0
            
            # Build product catalog once (reuse for all files)
            product_catalog, products_list = build_product_catalog_for_prompt()

            # Initialize Gemini client once for the entire batch
            bulk_api_key, model_name, _ = get_gemini_settings()
            bulk_genai_client = genai_sdk.Client(api_key=bulk_api_key)

            # --- STEP 1: Batch filename -> stockist mapping via Gemini (single call) ---
            all_filenames = [f for f, _, _ in all_files]
            gemini_mapping = {}
            try:
                filters = {"status": "Active"}
                if doc.division:
                    filters["division"] = doc.division
                    
                stockists_cat = frappe.get_all(
                    "Stockist Master",
                    filters=filters,
                    fields=["name", "stockist_name"],
                    order_by="stockist_name asc"
                )
                catalog_lines = [f"{s['name']}|{s['stockist_name']}" for s in stockists_cat]
                filenames_text = "\n".join([f"{i+1}. {f}" for i, f in enumerate(all_filenames)])
                map_prompt = (
                    "Match pharmaceutical statement filenames to stockist codes.\n\n"
                    "STOCKIST CATALOG (CODE|NAME):\n" + "\n".join(catalog_lines) + "\n\n"
                    "FILENAMES:\n" + filenames_text + "\n\n"
                    "Return JSON object: filename -> CODE (or null). Use exact filenames as keys.\n"
                    "Return ONLY valid JSON."
                )
                map_resp = bulk_genai_client.models.generate_content(
                    model=model_name,
                    contents=map_prompt,
                    config=genai_types.GenerateContentConfig(
                        thinking_config=genai_types.ThinkingConfig(thinking_budget=0)
                    )
                )
                resp_text = map_resp.text.strip()
                if resp_text.startswith("```"):
                    resp_text = resp_text.split("```", 1)[1]
                if resp_text.lower().startswith("json"):
                    resp_text = resp_text[4:]
                if resp_text.endswith("```"):
                    resp_text = resp_text.rsplit("```", 1)[0]
                raw_map = json.loads(resp_text.strip())
                valid_codes = {s["name"] for s in stockists_cat}
                for fname, code in raw_map.items():
                    gemini_mapping[fname] = code if (code and code in valid_codes) else None
                frappe.logger().info(f"Gemini batch mapping completed: {len(gemini_mapping)} entries")
            except Exception as map_err:
                frappe.logger().warning(f"Gemini batch mapping failed, using fuzzy fallback: {map_err}")

            for idx, (file, file_full_path, file_ext) in enumerate(all_files, 1):
                try:
                    # Identify stockist - Gemini mapping first, fuzzy fallback
                    stockist_code = gemini_mapping.get(file) if gemini_mapping else None
                    if not stockist_code:
                        stockist_code = identify_stockist_from_filename(file)
                    
                    if not stockist_code:
                        results.append({
                            "file": file,
                            "status": "Failed",
                            "message": "Could not identify stockist from filename"
                        })
                        failed_count += 1
                        # Update progress + counters so UI reflects this failure
                        doc.progress = (idx / len(all_files)) * 100
                        doc.success_count = success_count
                        doc.failed_count = failed_count
                        doc.skipped_count = skipped_count
                        doc.extraction_log = json.dumps(results)
                        doc.save(ignore_permissions=True)
                        frappe.db.commit()
                        continue
                    
                    # Check if already exists
                    existing = frappe.db.exists("Stockist Statement", {
                        "stockist_code": stockist_code,
                        "statement_month": month
                    })
                    
                    # Resolve stockist name for display
                    stockist_name = frappe.db.get_value("Stockist Master", stockist_code, "stockist_name") or stockist_code

                    if existing:
                        results.append({
                            "file": file,
                            "status": "Failed",
                            "message": f"Statement already exists for this stockist in this month: {existing}. Duplicate rejected.",
                            "stockist": stockist_name
                        })
                        failed_count += 1
                        # Update progress + counters so UI reflects this failure
                        doc.progress = (idx / len(all_files)) * 100
                        doc.success_count = success_count
                        doc.failed_count = failed_count
                        doc.skipped_count = skipped_count
                        doc.extraction_log = json.dumps(results)
                        doc.save(ignore_permissions=True)
                        frappe.db.commit()
                        continue
                    
                    # Create statement
                    statement_name = f"TEMP-{frappe.generate_hash(length=8)}"
                    
                    # Save file
                    from frappe.utils.file_manager import save_file_on_filesystem
                    file_doc = save_file_to_public(file, file_full_path, "Stockist Statement", statement_name)
                    
                    # Create statement doc
                    statement = frappe.get_doc({
                        "doctype": "Stockist Statement",
                        "stockist_code": stockist_code,
                        "statement_month": month,
                        "uploaded_file": file_doc.file_url,
                        "extracted_data_status": "Pending"
                    })
                    statement.insert(ignore_permissions=True)
                    
                    # Update file attachment
                    file_doc.attached_to_name = statement.name
                    file_doc.save(ignore_permissions=True)
                    
                    # Extract data using enhanced method (reuse already-configured client)
                    extracted_data = call_gemini_extraction_with_catalog(
                        file_full_path,
                        stockist_code,
                        product_catalog,
                        products_list,
                        model_name,
                        bulk_genai_client
                    )
                    
                    if extracted_data and len(extracted_data) > 0:
                        statement_rows, counts = _build_statement_rows(
                            extracted_data,
                            statement_division=doc.division,
                            products_list=products_list,
                        )
                        _replace_statement_items(statement, statement_rows)
                        statement.confidence_score = _calculate_confidence_score(statement_rows)

                        statement.extracted_data_status = "Completed"
                        statement.extraction_notes = _build_extraction_notes(
                            len(statement_rows),
                            statement.confidence_score,
                            unmapped_count=counts["unmapped_count"],
                            auto_mapped_count=counts["auto_mapped_count"],
                            special_row_count=counts["special_row_count"],
                            skipped_division_count=counts["skipped_division_count"],
                            statement_division=doc.division,
                        )
                    else:
                        statement.extracted_data_status = "Failed"
                        statement.extraction_notes = "No data extracted from file"
                    
                    statement.populate_previous_month_closing()
                    statement.calculate_closing_and_totals()
                    statement.calculate_qc_confidence()
                    statement.save(ignore_permissions=True)
                    
                    results.append({
                        "file": file,
                        "status": "Success",
                        "statement": statement.name,
                        "stockist": stockist_name,
                        "items_extracted": len(extracted_data) if extracted_data else 0,
                        "qc_confidence": statement.qc_confidence or "All Matched",
                    })
                    success_count += 1
                    
                except Exception as e:
                    error_msg = str(e)
                    frappe.log_error(
                        f"Error processing {file}: {error_msg}\n{frappe.get_traceback()}",
                        "Bulk Extract File Error"
                    )
                    # stockist_code / stockist_name may not be set if error happened before identification
                    _sc_display = "Unknown"
                    try:
                        _sc_display = stockist_name  # defined after stockist resolution
                    except NameError:
                        try:
                            _sc_display = stockist_code  # fall back to code
                        except NameError:
                            pass
                    results.append({
                        "file": file,
                        "status": "Failed",
                        "message": error_msg,
                        "stockist": _sc_display
                    })
                    failed_count += 1
                
                # Update progress + partial log so UI shows in-flight results
                doc.progress = (idx / len(all_files)) * 100
                doc.success_count = success_count
                doc.failed_count = failed_count
                doc.skipped_count = skipped_count
                doc.extraction_log = json.dumps(results)
                doc.save(ignore_permissions=True)
                frappe.db.commit()
        
        # Final update
        doc.status = "Completed" if failed_count == 0 else "Partially Completed"
        doc.progress = 100
        doc.success_count = success_count
        doc.failed_count = failed_count
        doc.skipped_count = skipped_count
        doc.extraction_log = json.dumps(results, indent=2)
        doc.save(ignore_permissions=True)
        frappe.db.commit()
        
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Bulk Extraction Background Job Error")
        
        try:
            doc = frappe.get_doc("Bulk Statement Upload", docname)
            doc.status = "Failed"
            doc.extraction_log = json.dumps([{"file": "—", "status": "Failed", "message": f"Job failed: {str(e)}"}])
            doc.save(ignore_permissions=True)
            frappe.db.commit()
        except Exception:
            pass

@frappe.whitelist()
def bulk_extract_statements(month, zip_file_url):
    """
    Bulk extract stock statements from ZIP file
    Creates draft statements for review
    """
    try:
        # Get ZIP file
        from frappe.utils.file_manager import get_file_path
        file_path = get_file_path(zip_file_url)

        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {zip_file_url}")

        # Create temporary directory
        with tempfile.TemporaryDirectory() as temp_dir:
            # Extract ZIP
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            # Process each file
            results = []
            
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    # Skip hidden/system files
                    if file.startswith('.') or file.startswith('__MACOSX') or file == 'Thumbs.db':
                        continue
                    
                    file_full_path = os.path.join(root, file)
                    file_ext = os.path.splitext(file)[1].lower()
                    
                    # Skip unsupported files
                    supported_extensions = ['.pdf', '.jpg', '.jpeg', '.png', '.csv', '.txt', '.xls', '.xlsx']
                    if file_ext not in supported_extensions:
                        results.append({
                            "file": file,
                            "status": "Skipped",
                            "message": f"Unsupported file type: {file_ext}"
                        })
                        continue
                    
                    # Try to identify stockist
                    stockist_code = identify_stockist_from_filename(file)
                    
                    if not stockist_code:
                        results.append({
                            "file": file,
                            "status": "Failed",
                            "message": "Could not identify stockist from filename"
                        })
                        continue
                    
                    try:
                        # Check if statement already exists
                        existing = frappe.db.exists("Stockist Statement", {
                            "stockist_code": stockist_code,
                            "statement_month": month
                        })
                        
                        if existing:
                            results.append({
                                "file": file,
                                "status": "Failed",
                                "message": f"Statement already exists for this stockist in this month: {existing}. Duplicate rejected.",
                                "stockist": stockist_code
                            })
                            continue
                        
                        # Create statement document (WITHOUT saving yet - to get name)
                        statement_name = f"TEMP-{frappe.generate_hash(length=8)}"
                        
                        # Copy file to public folder first
                        file_doc = save_file_to_public(file, file_full_path, "Stockist Statement", statement_name)
                        
                        # Now create the actual statement
                        statement = frappe.get_doc({
                            "doctype": "Stockist Statement",
                            "stockist_code": stockist_code,
                            "statement_month": month,
                            "uploaded_file": file_doc.file_url,
                            "extracted_data_status": "Pending"
                        })
                        
                        statement.insert(ignore_permissions=True)
                        
                        # Update file attachment to correct docname
                        file_doc.attached_to_name = statement.name
                        file_doc.save(ignore_permissions=True)
                        
                        # Extract data using the active Gemini extraction path
                        sync_api_key, sync_model_name, _ = get_gemini_settings()
                        sync_genai_client = genai_sdk.Client(api_key=sync_api_key)
                        product_catalog, products_list = build_product_catalog_for_prompt()
                        extracted_data = call_gemini_extraction_with_catalog(
                            file_full_path,
                            stockist_code,
                            product_catalog,
                            products_list,
                            sync_model_name,
                            sync_genai_client,
                        )
                        
                        if extracted_data and len(extracted_data) > 0:
                            statement_rows, counts = _build_statement_rows(
                                extracted_data,
                                statement_division=statement.division,
                                products_list=products_list,
                            )
                            _replace_statement_items(statement, statement_rows)
                            statement.confidence_score = _calculate_confidence_score(statement_rows)
                            statement.extracted_data_status = "Completed"
                            statement.extraction_notes = _build_extraction_notes(
                                len(statement_rows),
                                statement.confidence_score,
                                unmapped_count=counts["unmapped_count"],
                                auto_mapped_count=counts["auto_mapped_count"],
                                special_row_count=counts["special_row_count"],
                                skipped_division_count=counts["skipped_division_count"],
                                statement_division=statement.division,
                            )
                        else:
                            statement.extracted_data_status = "Failed"
                            statement.extraction_notes = "No data extracted from file"
                        
                        statement.populate_previous_month_closing()
                        statement.calculate_closing_and_totals()
                        statement.save(ignore_permissions=True)
                        
                        results.append({
                            "file": file,
                            "status": "Success",
                            "statement": statement.name,
                            "stockist": stockist_code,
                            "items_extracted": len(extracted_data) if extracted_data else 0
                        })
                        
                    except Exception as e:
                        error_msg = str(e)
                        frappe.log_error(
                            f"Error processing {file}: {error_msg}\n{frappe.get_traceback()}",
                            "Bulk Extract File Error"
                        )
                        results.append({
                            "file": file,
                            "status": "Failed",
                            "message": error_msg,
                            "stockist": stockist_code
                        })
            
            frappe.db.commit()
            
            success_count = len([r for r in results if r["status"] == "Success"])
            failed_count = len([r for r in results if r["status"] == "Failed"])
            
            return {
                "success": True,
                "total_files": len(results),
                "success_count": success_count,
                "failed_count": failed_count,
                "results": results
            }
            
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Bulk Extraction Error")
        return {
            "success": False,
            "message": str(e)
        }


@frappe.whitelist()
def get_unmatched_filenames_suggestion(zip_file_url):
    """
    Analyze ZIP file and suggest stockist matches for manual correction
    """
    try:
        file_path = frappe.get_site_path('public', zip_file_url.replace('/files/', ''))
        
        if not os.path.exists(file_path):
            return {"success": False, "message": "ZIP file not found"}
        
        suggestions = []
        
        with tempfile.TemporaryDirectory() as temp_dir:
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    if file.startswith('.') or file.startswith('__'):
                        continue
                    
                    stockist_code = identify_stockist_from_filename(file)
                    
                    # Get top 5 candidates
                    name_clean = os.path.splitext(file)[0].upper()
                    stockists = frappe.get_all("Stockist Master",
                        fields=["stockist_code", "stockist_name"],
                        filters={"status": "Active"})
                    
                    from difflib import SequenceMatcher
                    candidates = sorted(
                        [(s, SequenceMatcher(None, name_clean, s['stockist_name'].upper()).ratio()) 
                         for s in stockists],
                        key=lambda x: x[1],
                        reverse=True
                    )[:5]
                    
                    suggestions.append({
                        "filename": file,
                        "matched_stockist": stockist_code,
                        "top_candidates": [
                            {
                                "code": s['stockist_code'],
                                "name": s['stockist_name'],
                                "score": round(score * 100, 1)
                            }
                            for s, score in candidates
                        ]
                    })
        
        return {
            "success": True,
            "suggestions": suggestions
        }
    except Exception as e:
        return {"success": False, "message": str(e)}

def identify_stockist_from_filename(filename):
    """
    Identify stockist code from filename using robust fuzzy matching
    """
    import re
    from difflib import SequenceMatcher
    
    # Remove extension and normalize
    name = os.path.splitext(filename)[0]
    name_clean = name.upper().replace('-', ' ').replace('_', ' ').strip()
    
    # Remove common date patterns and keywords
    date_patterns = [
        r'\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b',
        r'\b\d{4}[-/]\d{1,2}[-/]\d{1,2}\b',
        r'\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\b',
        r'\b(JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\b',
        r'\b20\d{2}\b',  # Years 2000-2099
        r'\bSTATEMENT\b',
        r'\bSTOCK\b',
        r'\bSALES\b',
        r'\bREPORT\b'
    ]
    
    for pattern in date_patterns:
        name_clean = re.sub(pattern, '', name_clean, flags=re.IGNORECASE)
    
    # Clean up extra spaces
    name_clean = ' '.join(name_clean.split()).strip()
    
    if not name_clean or len(name_clean) < 3:
        frappe.log_error(f"Filename too short after cleaning: {filename}", "Stockist ID Failed")
        return None
    
    # Get all active stockists
    stockists = frappe.get_all("Stockist Master", 
        fields=["stockist_code", "stockist_name", "city"],
        filters={"status": "Active"})
    
    if not stockists:
        frappe.log_error("No active stockists found", "Stockist ID Failed")
        return None
    
    # Strategy 1: Exact stockist code match
    for s in stockists:
        if s['stockist_code'].upper() in name_clean:
            return s['stockist_code']
    
    # Strategy 2: Fuzzy match on stockist name
    best_match = None
    best_score = 0
    
    # Stop words to ignore
    stop_words = {
        'LLP', 'PVT', 'LTD', 'LIMITED', 'CO', 'COMPANY', 'AND', 'THE', 'A', 'AN', 
        'PHARMA', 'PHARMACEUTICAL', 'PHARMACEUTICALS', 'DIST', 'DISTRIBUTOR',
        'DISTRIBUTORS', 'TRADERS', 'ENTERPRISES', 'AGENCY', 'AGENCIES'
    }
    
    for s in stockists:
        stockist_name_clean = s['stockist_name'].upper().strip()
        
        # Calculate direct similarity
        similarity = SequenceMatcher(None, name_clean, stockist_name_clean).ratio()
        
        # Word-based matching
        stockist_words = set(stockist_name_clean.split())
        filename_words = set(name_clean.split())
        
        # Filter out stop words and short words
        stockist_words_filtered = {w for w in stockist_words 
                                   if len(w) > 2 and w not in stop_words}
        filename_words_filtered = {w for w in filename_words 
                                  if len(w) > 2 and w not in stop_words}
        
        if not stockist_words_filtered:
            # If no significant words, use all words
            stockist_words_filtered = {w for w in stockist_words if len(w) > 1}
        
        # Calculate word overlap
        common_words = stockist_words_filtered.intersection(filename_words_filtered)
        word_overlap_score = (len(common_words) / len(stockist_words_filtered) 
                             if stockist_words_filtered else 0)
        
        # Check for partial matches (important for names like "Dhanvantri" vs "Dhanvantari")
        partial_match_score = 0
        for s_word in stockist_words_filtered:
            for f_word in filename_words_filtered:
                if len(s_word) >= 4 and len(f_word) >= 4:
                    # Check if one is substring of other
                    if s_word in f_word or f_word in s_word:
                        partial_match_score += 0.3
                    # Check character-level similarity for typos
                    elif SequenceMatcher(None, s_word, f_word).ratio() > 0.8:
                        partial_match_score += 0.2
        
        partial_match_score = min(partial_match_score, 0.5)  # Cap at 0.5
        
        # Weighted combined score
        combined_score = (similarity * 0.4) + (word_overlap_score * 0.4) + (partial_match_score * 0.2)
        
        # Bonus for matching 2+ significant words
        if len(common_words) >= 2:
            combined_score += 0.15
        elif len(common_words) == 1 and len(stockist_words_filtered) == 1:
            # Single unique word match (e.g., "Jyoti")
            combined_score += 0.2
        
        # Bonus for exact word match
        if stockist_words_filtered == filename_words_filtered:
            combined_score += 0.2
        
        if combined_score > best_score:
            best_score = combined_score
            best_match = s
    
    # Strategy 3: City-based matching with additional context
    if best_score < 0.5:
        for s in stockists:
            if s.get('city') and s['city']:
                city_clean = s['city'].upper().strip()
                if len(city_clean) > 3 and city_clean in name_clean:
                    stockist_words = {w.upper() for w in s['stockist_name'].split() 
                                    if len(w) > 3 and w.upper() not in stop_words}
                    filename_words = set(name_clean.split())
                    
                    if stockist_words.intersection(filename_words):
                        city_match_score = 0.55
                        if city_match_score > best_score:
                            best_score = city_match_score
                            best_match = s
    
    # Accept match if confidence is above threshold
    CONFIDENCE_THRESHOLD = 0.40  # Lowered slightly for flexibility
    
    if best_match and best_score >= CONFIDENCE_THRESHOLD:
        frappe.logger().info(
            f"✓ Matched: {filename} -> {best_match['stockist_name']} "
            f"({best_match['stockist_code']}) [Score: {best_score:.2f}]"
        )
        return best_match['stockist_code']
    
    # Log failure with top 3 candidates for debugging
    if stockists:
        top_candidates = sorted(
            [(s, SequenceMatcher(None, name_clean, s['stockist_name'].upper()).ratio()) 
             for s in stockists],
            key=lambda x: x[1],
            reverse=True
        )[:3]
        
        candidates_info = "\n".join([
            f"  - {s['stockist_name']} ({s['stockist_code']}): {score:.2f}"
            for s, score in top_candidates
        ])
        
        frappe.log_error(
            f"Could not identify stockist from: {filename}\n"
            f"Clean name: '{name_clean}'\n"
            f"Best match: {best_match['stockist_name'] if best_match else 'None'}\n"
            f"Best score: {best_score:.2f}\n"
            f"Top candidates:\n{candidates_info}",
            "Stockist Identification Failed"
        )
    
    return None

def save_file_to_public(filename, file_path, doctype, docname):
    """Save file to public folder and create File document"""
    try:
        with open(file_path, 'rb') as f:
            content = f.read()
        
        # Ensure docname is valid
        if not docname or not isinstance(docname, (str, int)):
            # Create a temporary name if docname is invalid
            docname = frappe.generate_hash(length=10)
        
        file_doc = frappe.get_doc({
            "doctype": "File",
            "file_name": filename,
            "attached_to_doctype": doctype,
            "attached_to_name": str(docname),  # Ensure it's a string
            "content": content,
            "is_private": 0
        })
        file_doc.save(ignore_permissions=True)
        
        return file_doc
    except Exception as e:
        frappe.log_error(
            f"Error saving file {filename}: {str(e)}\n{frappe.get_traceback()}",
            "File Save Error"
        )
        raise

@frappe.whitelist()
def fetch_previous_month_closing(stockist_code, current_month):
    """Fetch previous month's closing balance to set as opening balance"""
    try:
        from dateutil.relativedelta import relativedelta

        if not stockist_code or not current_month:
            return []

        current_date = frappe.utils.getdate(current_month)
        previous_month = current_date - relativedelta(months=1)
        previous_month_first = get_first_day(previous_month)

        # Find previous month's statement
        prev_statement = frappe.db.get_value("Stockist Statement", {
            "stockist_code": stockist_code,
            "statement_month": previous_month_first,
            "docstatus": 1
        }, "name")

        if not prev_statement:
            frappe.msgprint("No previous month statement found", indicator='orange')
            return []

        # Get items from previous statement
        prev_items = frappe.get_all("Stockist Statement Item",
            filters={"parent": prev_statement},
            fields=["product_code", "product_name", "pack", "closing_qty", "pts", "closing_value"])

        return prev_items or []

    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Fetch Previous Month Error")
        return []

@frappe.whitelist()
def reroute_scheme_request(doc_name, comments):
    try:
        doc = frappe.get_doc("Scheme Request", doc_name)

        if doc.approval_status == "Rerouted":
            frappe.throw("Scheme request already rerouted")

        doc.approval_status = "Rerouted"
        doc.append("approval_log", {
            "approver": frappe.session.user,
            "approval_level": "Manager",
            "action": "Rerouted",
            "action_date": nowdate(),
            "comments": comments or "Rerouted for revision"
        })

        doc.save()
        frappe.db.commit()

        send_scheme_notification(doc, "Rerouted", comments)

        return True
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Reroute Scheme Error")
        frappe.throw(str(e))

@frappe.whitelist()
def approve_scheme_request(doc_name, comments):
    """Approve a scheme request"""
    try:
        doc = frappe.get_doc("Scheme Request", doc_name)

        if doc.approval_status == "Approved":
            frappe.throw("Scheme request already approved")

        doc.approval_status = "Approved"
        doc.append("approval_log", {
            "approver": frappe.session.user,
            "approval_level": "Manager",
            "action": "Approved",
            "action_date": nowdate(),
            "comments": comments or "Approved"
        })

        doc.save()
        # Submit the doc (docstatus=1) so it becomes available for scheme deduction
        doc.submit()
        frappe.db.commit()

        # Send notification
        send_scheme_notification(doc, "Approved", comments)

        return True
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Approve Scheme Error")
        frappe.throw(str(e))

@frappe.whitelist()
def reject_scheme_request(doc_name, comments):
    """Reject a scheme request"""
    try:
        doc = frappe.get_doc("Scheme Request", doc_name)

        if doc.approval_status == "Rejected":
            frappe.throw("Scheme request already rejected")

        doc.approval_status = "Rejected"
        doc.append("approval_log", {
            "approver": frappe.session.user,
            "approval_level": "Manager",
            "action": "Rejected",
            "action_date": nowdate(),
            "comments": comments or "Rejected"
        })

        doc.save()
        frappe.db.commit()

        # Send notification
        send_scheme_notification(doc, "Rejected", comments)

        return True
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Reject Scheme Error")
        frappe.throw(str(e))

def send_scheme_notification(doc, action, comments):
    """Send email notification for scheme request approval/rejection"""
    try:
        subject = f"Scheme Request {doc.name} - {action}"
        message = f"""
        <p>Dear {doc.requested_by},</p>
        <p>Your scheme request {doc.name} has been <strong>{action}</strong>.</p>
        <p><strong>Doctor:</strong> {doc.doctor_name or 'N/A'} ({doc.doctor_code or 'N/A'})</p>
        <p><strong>Stockist:</strong> {doc.stockist_name or 'N/A'}</p>
        <p><strong>Total Value:</strong> ₹{flt(doc.total_scheme_value or 0):,.2f}</p>
        <p><strong>Comments:</strong> {comments or 'None'}</p>
        <p>Please check the system for more details.</p>
        """

        frappe.sendmail(
            recipients=[doc.requested_by],
            subject=subject,
            message=message
        )
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Send Notification Error")

@frappe.whitelist()
def search_doctors(search_term):
    """Search doctors by name or place"""
    try:
        search_term = f"%{search_term}%"

        doctors = frappe.db.sql("""
            SELECT 
                name,
                doctor_code,
                doctor_name,
                place,
                specialization,
                hospital_address,
                hq,
                team,
                region
            FROM `tabDoctor Master`
            WHERE status = 'Active'
            AND (
                doctor_name LIKE %(search_term)s
                OR place LIKE %(search_term)s
                OR doctor_code LIKE %(search_term)s
            )
            ORDER BY doctor_name
            LIMIT 20
        """, {"search_term": search_term}, as_dict=True)

        return doctors or []
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Doctor Search Error")
        return []

@frappe.whitelist()
def get_stockists_by_team(team):
    """Get all active stockists for a team"""
    try:
        stockists = frappe.get_all("Stockist Master",
            filters={
                "team": team,
                "status": "Active"
            },
            fields=["name", "stockist_code", "stockist_name", "city", "hq"],
            order_by="stockist_name"
        )
        return stockists or []
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Get Stockists Error")
        return []

@frappe.whitelist()
def search_stockists(search_term, team=None):
    """Search stockists by name or city, optionally filter by team"""
    try:
        search_term = f"%{search_term}%"

        conditions = "status = 'Active'"

        if team:
            conditions += f" AND team = '{team}'"

        stockists = frappe.db.sql(f"""
            SELECT 
                name,
                stockist_code,
                stockist_name,
                city,
                hq,
                team,
                region
            FROM `tabStockist Master`
            WHERE {conditions}
            AND (
                stockist_name LIKE %(search_term)s
                OR city LIKE %(search_term)s
                OR stockist_code LIKE %(search_term)s
            )
            ORDER BY stockist_name
            LIMIT 20
        """, {"search_term": search_term}, as_dict=True)

        return stockists or []
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Stockist Search Error")
        return []

@frappe.whitelist()
def get_dashboard_data():
    """Get dashboard KPI data with null safety"""
    try:
        # Total stockists
        total_stockists = frappe.db.count("Stockist Master", {"status": "Active"}) or 0

        # Total schemes this month
        from frappe.utils import get_first_day, get_last_day, today
        first_day = get_first_day(today())
        last_day = get_last_day(today())

        total_schemes = frappe.db.count("Scheme Request", {
            "application_date": ["between", [first_day, last_day]]
        }) or 0

        pending_schemes = frappe.db.count("Scheme Request", {
            "approval_status": "Pending",
            "application_date": ["between", [first_day, last_day]]
        }) or 0

        approved_schemes = frappe.db.count("Scheme Request", {
            "approval_status": "Approved",
            "application_date": ["between", [first_day, last_day]]
        }) or 0

        # Total scheme value this month
        result = frappe.db.sql("""
            SELECT COALESCE(SUM(total_scheme_value), 0) as total
            FROM `tabScheme Request`
            WHERE application_date BETWEEN %s AND %s
            AND approval_status = 'Approved'
        """, (first_day, last_day), as_dict=True)

        total_scheme_value = flt(result[0].total) if result and result[0] else 0

        # Statements processed this month
        statements_processed = frappe.db.count("Stockist Statement", {
            "statement_month": ["between", [first_day, last_day]],
            "docstatus": 1
        }) or 0

        return {
            "total_stockists": total_stockists,
            "total_schemes": total_schemes,
            "pending_schemes": pending_schemes,
            "approved_schemes": approved_schemes,
            "total_scheme_value": total_scheme_value,
            "statements_processed": statements_processed
        }
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Get Dashboard Data Error")
        return {
            "total_stockists": 0,
            "total_schemes": 0,
            "pending_schemes": 0,
            "approved_schemes": 0,
            "total_scheme_value": 0,
            "statements_processed": 0
        }

@frappe.whitelist()
def upload_company_logo():
    """Upload company logo"""
    return {
        "message": "Upload your logo to /public/files/stedman_logo.png"
    }

@frappe.whitelist()
def get_workspace_settings():
    """Get workspace settings including logo"""
    logo_path = frappe.db.get_single_value("Scanify Settings", "company_logo")
    if not logo_path:
        logo_path = "/files/stedman_logo.png"  # default

    return {
        "logo": logo_path,
        "company_name": "Stedman Pharmaceuticals"
    }

@frappe.whitelist()
def get_product_history_for_scheme(product_code, doctor_code=None, hq=None):
    """Get historical data for a product in scheme context"""
    try:
        from frappe.utils import getdate, add_months
        import json
        
        # Get product details
        product = frappe.get_doc("Product Master", product_code)
        
        # Build filters
        filters = {
            "docstatus": ["<", 2],  # Not cancelled
            "approval_status": ["in", ["Approved", "Pending", "Rerouted"]]
        }
        
        if doctor_code:
            filters["doctor_code"] = doctor_code
        
        if hq:
            filters["hq"] = hq
        
        # Get past scheme requests with this product
        schemes = frappe.db.sql("""
            SELECT 
                sr.name,
                sr.application_date,
                sr.doctor_name,
                sr.doctor_code,
                sr.approval_status,
                sri.quantity,
                sri.product_value
            FROM 
                `tabScheme Request` sr
            INNER JOIN 
                `tabScheme Request Item` sri ON sr.name = sri.parent
            WHERE 
                sri.product_code = %(product_code)s
                AND sr.docstatus < 2
                AND sr.application_date >= %(six_months_ago)s
            ORDER BY 
                sr.application_date DESC
            LIMIT 10
        """, {
            "product_code": product_code,
            "six_months_ago": add_months(getdate(), -6)
        }, as_dict=True)
        
        # Calculate aggregates
        total_quantity = 0
        total_value = 0
        total_schemes = len(schemes)
        last_order_date = None
        
        for scheme in schemes:
            total_quantity += flt(scheme.quantity or 0)
            total_value += flt(scheme.product_value or 0)
            if not last_order_date and scheme.application_date:
                last_order_date = scheme.application_date
        
        # Get monthly trend data (last 6 months)
        chart_data = frappe.db.sql("""
            SELECT 
                DATE_FORMAT(sr.application_date, '%%Y-%%m') as month,
                SUM(sri.quantity) as quantity,
                SUM(sri.product_value) as value
            FROM 
                `tabScheme Request` sr
            INNER JOIN 
                `tabScheme Request Item` sri ON sr.name = sri.parent
            WHERE 
                sri.product_code = %(product_code)s
                AND sr.approval_status = 'Approved'
                AND sr.application_date >= %(six_months_ago)s
            GROUP BY 
                DATE_FORMAT(sr.application_date, '%%Y-%%m')
            ORDER BY 
                month DESC
        """, {
            "product_code": product_code,
            "six_months_ago": add_months(getdate(), -6)
        }, as_dict=True)
        
        return {
            "success": True,
            "product_code": product.product_code,
            "product_name": product.product_name,
            "pack": product.pack,
            "pts": product.pts,
            "total_schemes": total_schemes,
            "total_quantity": total_quantity,
            "total_value": total_value,
            "last_order_date": last_order_date.strftime("%Y-%m-%d") if last_order_date else None,
            "recent_schemes": schemes,
            "chart_data": chart_data
        }
        
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Get Product History Error")
        return {
            "success": False,
            "message": str(e)
        }


















@frappe.whitelist()
def get_incentive_calculation_data(filters):
    """
    Fetch incentive calculation data for Prima and Vektra
    Filter by: Month, HQ, Team, Region, Stockist
    """
    try:
        filters = json.loads(filters) if isinstance(filters, str) else filters
        
        # Get stockist statements
        filters_dict = {
            "docstatus": 1  # Submitted only
        }
        
        # Apply month filter
        if filters.get("month"):
            month = frappe.utils.getdate(f"{filters['month']}-01")
            from_date = get_first_day(month)
            to_date = frappe.utils.get_last_day_of_the_month(month)
            filters_dict["statement_month__gte"] = from_date
            filters_dict["statement_month__lte"] = to_date
        
        # Apply HQ filter if selected
        if filters.get("hq"):
            hq_stockists = frappe.get_all(
                "Stockist Master",
                filters={"hq": filters["hq"]},
                fields=["stockist_code"]
            )
            stockist_codes = [s["stockist_code"] for s in hq_stockists]
            filters_dict["stockist_code"] = ["in", stockist_codes]
        
        # Fetch all statements
        statements = frappe.get_all(
            "Stockist Statement",
            filters=filters_dict,
            fields=["name", "stockist_code", "statement_month"]
        )
        
        # Aggregate data
        incentive_data = {}
        
        for stmt in statements:
            stmt_doc = frappe.get_doc("Stockist Statement", stmt["name"])
            stockist = frappe.get_doc("Stockist Master", stmt["stockist_code"])
            
            # Get items grouped by product type (Prima/Vektra)
            for item in stmt_doc.items:
                product = frappe.get_doc("Product Master", item.product_code)
                product_type = product.product_type  # 'Prima' or 'Vektra'
                
                key = f"{stockist_code}_{product_type}_{item.product_code}"
                
                if key not in incentive_data:
                    incentive_data[key] = {
                        "stockist_code": stmt["stockist_code"],
                        "stockist_name": stockist.stockist_name,
                        "hq": stockist.hq,
                        "product_code": item.product_code,
                        "product_name": product.product_name,
                        "product_type": product_type,
                        "total_qty": 0,
                        "total_value": 0,
                        "incentive_amount": 0
                    }
                
                # Calculate values at PTR (Price to Retailer)
                qty = flt(item.sales_qty) + flt(item.free_qty)
                ptr_value = qty * flt(product.ptr or 0)
                
                incentive_data[key]["total_qty"] += qty
                incentive_data[key]["total_value"] += ptr_value
        
        # Calculate incentive based on target and achievement
        result = []
        for key, data in incentive_data.items():
            # Get target from Incentive Master
            incentive_config = frappe.db.get_value(
                "Incentive Master",
                {
                    "product_type": data["product_type"],
                    "status": "Active"
                },
                ["incentive_rate", "min_target"]
            )
            
            if incentive_config:
                incentive_rate, min_target = incentive_config
                incentive_rate = flt(incentive_rate) / 100  # Convert percentage
                
                if data["total_value"] >= flt(min_target):
                    data["incentive_amount"] = data["total_value"] * incentive_rate
            
            result.append(data)
        
        return {
            "success": True,
            "data": result,
            "total_records": len(result)
        }
    
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Incentive Calculation Error")
        return {
            "success": False,
            "message": str(e)
        }

@frappe.whitelist()
def export_incentive_report(filters, format_type="xlsx"):
    """Export incentive report to Excel or PDF"""
    try:
        filters = json.loads(filters) if isinstance(filters, str) else filters
        data = get_incentive_calculation_data(filters)
        
        if not data["success"]:
            frappe.throw(data["message"])
        
        records = data["data"]
        
        if format_type == "xlsx":
            return export_to_excel_incentive(records, filters)
        else:
            return export_to_pdf_incentive(records, filters)
    
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Export Incentive Error")
        frappe.throw(str(e))

def export_to_excel_incentive(records, filters):
    """Generate Excel file for incentive report"""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    import io
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Incentive Report"
    
    # Add header
    ws['A1'] = "Stedman Pharmaceuticals Pvt Ltd"
    ws['A1'].font = Font(bold=True, size=14)
    
    ws['A2'] = f"Incentive Calculation Report"
    ws['A2'].font = Font(bold=True, size=12)
    
    if filters.get("month"):
        ws['A3'] = f"Month: {filters['month']}"
    
    # Column headers
    headers = [
        "Stockist Code", "Stockist Name", "HQ", "Product Code", "Product Name",
        "Product Type", "Total Qty", "Total Value (₹)", "Incentive (₹)"
    ]
    
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=6, column=col)
        cell.value = header
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Add data rows
    row = 7
    total_incentive = 0
    
    for record in records:
        ws.cell(row, 1).value = record.get("stockist_code", "")
        ws.cell(row, 2).value = record.get("stockist_name", "")
        ws.cell(row, 3).value = record.get("hq", "")
        ws.cell(row, 4).value = record.get("product_code", "")
        ws.cell(row, 5).value = record.get("product_name", "")
        ws.cell(row, 6).value = record.get("product_type", "")
        ws.cell(row, 7).value = record.get("total_qty", 0)
        ws.cell(row, 8).value = record.get("total_value", 0)
        ws.cell(row, 9).value = record.get("incentive_amount", 0)
        
        # Format currency columns
        ws.cell(row, 8).number_format = '₹ #,##0.00'
        ws.cell(row, 9).number_format = '₹ #,##0.00'
        
        total_incentive += flt(record.get("incentive_amount", 0))
        row += 1
    
    # Add total row
    ws.cell(row, 8).value = "TOTAL INCENTIVE:"
    ws.cell(row, 8).font = Font(bold=True)
    ws.cell(row, 9).value = total_incentive
    ws.cell(row, 9).font = Font(bold=True)
    ws.cell(row, 9).number_format = '₹ #,##0.00'
    
    # Adjust column widths
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 20
    ws.column_dimensions['F'].width = 12
    ws.column_dimensions['G'].width = 12
    ws.column_dimensions['H'].width = 15
    ws.column_dimensions['I'].width = 15
    
    # Save to file
    filename = f"Incentive_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    filepath = frappe.get_site_path('public', 'exports', filename)
    
    import os
    os.makedirs(os.path.dirname(filepath), exist_ok=True)
    wb.save(filepath)
    
    return {
        "success": True,
        "message": "Report generated successfully",
        "file_url": f"/files/{filename}"
    }
@frappe.whitelist()
def get_doctor_history_for_scheme(doctor_code, hq=None):
    """
    Get historical data for a doctor in scheme context
    Shows all past scheme requests for this doctor
    """
    try:
        from frappe.utils import getdate, add_months
        import json
        
        # Get doctor details
        doctor = frappe.get_doc("Doctor Master", doctor_code)
        
        # Build filters
        filters = {
            "docstatus": ("!=", 2),  # Not cancelled
            "doctor_code": doctor_code
        }
        
        if hq:
            filters["hq"] = hq
        
        # Get past scheme requests for this doctor (last 12 months)
        twelve_months_ago = add_months(getdate(), -12)
        schemes = frappe.db.sql("""
            SELECT 
                sr.name,
                sr.application_date,
                sr.stockist_name,
                sr.hq,
                sr.approval_status,
                sr.total_scheme_value,
                COUNT(sri.name) as product_count
            FROM `tabScheme Request` sr
            LEFT JOIN `tabScheme Request Item` sri ON sr.name = sri.parent
            WHERE sr.doctor_code = %(doctor_code)s
                AND sr.docstatus != 2
                AND sr.application_date >= %(twelve_months_ago)s
            GROUP BY sr.name
            ORDER BY sr.application_date DESC
            LIMIT 20
        """, {
            "doctor_code": doctor_code,
            "twelve_months_ago": twelve_months_ago
        }, as_dict=True)
        
        # Calculate aggregates
        total_schemes = len(schemes)
        total_approved = len([s for s in schemes if s.approval_status == "Approved"])
        total_pending = len([s for s in schemes if s.approval_status == "Pending"])
        total_rejected = len([s for s in schemes if s.approval_status == "Rejected"])
        total_value = sum([flt(s.total_scheme_value or 0) for s in schemes])
        
        last_scheme_date = schemes[0].application_date if schemes else None
        
        # Get product-wise breakdown
        product_summary = frappe.db.sql("""
            SELECT 
                sri.product_code,
                sri.product_name,
                SUM(sri.quantity) as total_quantity,
                SUM(sri.free_quantity) as total_free_quantity,
                SUM(sri.product_value) as total_value,
                COUNT(DISTINCT sr.name) as scheme_count
            FROM `tabScheme Request` sr
            INNER JOIN `tabScheme Request Item` sri ON sr.name = sri.parent
            WHERE sr.doctor_code = %(doctor_code)s
                AND sr.approval_status = 'Approved'
                AND sr.application_date >= %(twelve_months_ago)s
            GROUP BY sri.product_code
            ORDER BY total_value DESC
            LIMIT 10
        """, {
            "doctor_code": doctor_code,
            "twelve_months_ago": twelve_months_ago
        }, as_dict=True)
        
        # Get monthly trend data (last 6 months)
        six_months_ago = add_months(getdate(), -6)
        chart_data = frappe.db.sql("""
            SELECT 
                DATE_FORMAT(sr.application_date, '%%Y-%%m') as month,
                COUNT(sr.name) as scheme_count,
                SUM(sr.total_scheme_value) as total_value
            FROM `tabScheme Request` sr
            WHERE sr.doctor_code = %(doctor_code)s
                AND sr.approval_status = 'Approved'
                AND sr.application_date >= %(six_months_ago)s
            GROUP BY DATE_FORMAT(sr.application_date, '%%Y-%%m')
            ORDER BY month DESC
        """, {
            "doctor_code": doctor_code,
            "six_months_ago": six_months_ago
        }, as_dict=True)
        
        return {
            "success": True,
            "doctor_code": doctor.doctor_code,
            "doctor_name": doctor.doctor_name,
            "place": doctor.place or "N/A",
            "specialization": doctor.specialization or "General",
            "hospital_address": doctor.hospital_address or "N/A",
            "hq": doctor.hq or "N/A",
            "total_schemes": total_schemes,
            "total_approved": total_approved,
            "total_pending": total_pending,
            "total_rejected": total_rejected,
            "total_value": total_value,
            "last_scheme_date": last_scheme_date.strftime("%Y-%m-%d") if last_scheme_date else None,
            "recent_schemes": schemes,
            "product_summary": product_summary,
            "chart_data": chart_data
        }
        
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Get Doctor History Error")
        return {"success": False, "message": str(e)}
@frappe.whitelist()
def set_user_division(division):
    """Set user's selected division in session and User document"""
    
    # Validate division
    if division not in ["Prima", "Vektra"]:
        frappe.throw("Invalid division")
    
    try:
        # Store in session
        frappe.session.user_division = division
        
        # Store in User document for persistence
        user_doc = frappe.get_doc("User", frappe.session.user)
        
        # Check if field exists
        if hasattr(user_doc, "division"):
            user_doc.division = division
            user_doc.save(ignore_permissions=True)
            frappe.db.commit()
        else:
            frappe.log_error("Division field not found in User doctype")
        
        return {
            "success": True, 
            "division": division,
            "message": f"Division switched to {division}"
        }
        
    except Exception as e:
        frappe.log_error(f"Could not save division to user: {str(e)}")
        return {
            "success": False,
            "message": f"Error: {str(e)}"
        }


@frappe.whitelist()
def get_user_division():
    """Get user's current division with proper fallback logic"""
    
    # Priority 1: Check session first (for current page load)
    if hasattr(frappe.session, "user_division") and frappe.session.user_division:
        return frappe.session.user_division
    
    # Priority 2: Check User document (persistent storage)
    try:
        user_doc = frappe.get_doc("User", frappe.session.user)
        if hasattr(user_doc, "division") and user_doc.division:
            # Sync to session
            frappe.session.user_division = user_doc.division
            return user_doc.division
    except:
        pass
    
    # Default: Prima
    default_division = "Prima"
    frappe.session.user_division = default_division
    return default_division


@frappe.whitelist()
def get_user_schemes(filters=None):
    """Get schemes for current user based on division and role"""
    user = frappe.session.user
    division = get_user_division(user)
    
    # Build filters
    scheme_filters = {"division": division}
    
    if filters:
        if isinstance(filters, str):
            filters = json.loads(filters)
        
        if filters.get("status"):
            scheme_filters["approval_status"] = filters["status"]
        if filters.get("from_date"):
            scheme_filters["application_date"] = [">=", filters["from_date"]]
        if filters.get("to_date"):
            if "application_date" in scheme_filters:
                scheme_filters["application_date"] = [
                    "between",
                    [filters["from_date"], filters["to_date"]]
                ]
            else:
                scheme_filters["application_date"] = ["<=", filters["to_date"]]
    
    # Get schemes
    schemes = frappe.get_all(
        "Scheme Request",
        filters=scheme_filters,
        fields=[
            "name", "application_date", "doctor_name", "stockist_name",
            "hq", "total_scheme_value", "approval_status"
        ],
        order_by="application_date desc",
        limit=100
    )
    
    return schemes

@frappe.whitelist()
def get_user_hqs(division=None):
    """Get HQs for current user based on division"""
    if not division:
        division = get_user_division()

    filters = {"status": "Active"}
    if division and division != "Both":
        filters["division"] = ["in", [division, "Both"]]
    
    hqs = frappe.get_all(
        "HQ Master",
        filters=filters,
        fields=["name", "hq_name", "team", "region"],
        order_by="hq_name asc",
    )

    # Fetch display names for team and region links
    team_names = {}
    region_names = {}
    for hq in hqs:
        if hq.team and hq.team not in team_names:
            team_names[hq.team] = frappe.db.get_value("Team Master", hq.team, "team_name") or hq.team
        if hq.region and hq.region not in region_names:
            region_names[hq.region] = frappe.db.get_value("Region Master", hq.region, "region_name") or hq.region
        hq["team_name"] = team_names.get(hq.team, hq.team or "")
        hq["region_name"] = region_names.get(hq.region, hq.region or "")
    
    return hqs

@frappe.whitelist()
def get_active_products(division=None):
    """Get all active products for scheme entry"""
    if not division:
        division = get_user_division()
    
    products = frappe.get_all(
        "Product Master",
        filters={"status": "Active"},
        fields=["name", "product_code", "product_name", "pack", "pts", "division"]
    )
    
    # Filter by division if applicable
    if division and division != "Both":
        products = [p for p in products if p.get("division") in (division, "Both")]
    
    return products

@frappe.whitelist()
def get_doctors_for_hq(hq=None, division=None):
    """Get all active doctors for a given HQ and division (for dropdown)"""
    if not division:
        division = get_user_division()

    conditions = ["status = 'Active'"]
    params = {}
    if hq:
        conditions.append("hq = %(hq)s")
        params["hq"] = hq
    if division and division != "Both":
        conditions.append("(division = %(division)s OR division = 'Both')")
        params["division"] = division

    where_sql = " AND ".join(conditions)
    doctors = frappe.db.sql("""
        SELECT name, doctor_code, doctor_name, place, specialization,
               hospital_address, hq, team, region
        FROM `tabDoctor Master`
        WHERE {where_sql}
        ORDER BY doctor_name
        LIMIT 500
    """.format(where_sql=where_sql), params, as_dict=True)

    return doctors

@frappe.whitelist()
def get_approved_doctors_for_hq(hq=None, division=None):
    """Get doctors who have at least one approved scheme in the given HQ/division"""
    if not division:
        division = get_user_division()

    conditions = ["dm.status = 'Active'"]
    params = {}
    if hq:
        conditions.append("sr.hq = %(hq)s")
        params["hq"] = hq
    if division and division != "Both":
        conditions.append("(sr.division = %(division)s OR sr.division = 'Both')")
        params["division"] = division

    where_sql = " AND ".join(conditions)
    doctors = frappe.db.sql("""
        SELECT DISTINCT dm.name, dm.doctor_code, dm.doctor_name, dm.place,
               dm.specialization, dm.hospital_address, dm.hq, dm.team, dm.region
        FROM `tabDoctor Master` dm
        INNER JOIN `tabScheme Request` sr ON sr.doctor_code = dm.name
        WHERE sr.approval_status = 'Approved'
          AND sr.docstatus != 2
          AND {where_sql}
        ORDER BY dm.doctor_name
        LIMIT 500
    """.format(where_sql=where_sql), params, as_dict=True)

    return doctors

@frappe.whitelist()
def get_approved_products_for_doctor(doctor_code=None, division=None):
    """Get products that appear in approved schemes for a specific doctor"""
    if not division:
        division = get_user_division()
    if not doctor_code:
        return []

    conditions = [
        "sr.approval_status = 'Approved'",
        "sr.docstatus != 2",
        "sr.doctor_code = %(doctor_code)s",
    ]
    params = {"doctor_code": doctor_code}
    if division and division != "Both":
        conditions.append("(sr.division = %(division)s OR sr.division = 'Both')")
        params["division"] = division

    where_sql = " AND ".join(conditions)
    products = frappe.db.sql("""
        SELECT DISTINCT pm.name, pm.product_code, pm.product_name, pm.pack, pm.pts, pm.division
        FROM `tabProduct Master` pm
        INNER JOIN `tabScheme Request Item` sri ON sri.product_code = pm.name
        INNER JOIN `tabScheme Request` sr ON sr.name = sri.parent
        WHERE pm.status = 'Active'
          AND {where_sql}
        ORDER BY pm.product_name
    """.format(where_sql=where_sql), params, as_dict=True)

    return products

@frappe.whitelist()
def get_stockists_by_hq(hq, division=None):
    """Get stockists for a specific HQ"""
    if not division:
        division = get_user_division()

    filters = {"hq": hq, "status": "Active"}
    if division and division != "Both":
        filters["division"] = ["in", [division, "Both"]]

    stockists = frappe.get_all(
        "Stockist Master",
        filters=filters,
        fields=["name", "stockist_code", "stockist_name"]
    )
    
    return stockists

@frappe.whitelist()
def search_doctors(searchterm=None, search_term=None, division=None, hq=None):
    """Search doctors by name or code, optionally filtered by HQ"""
    term = searchterm or search_term or ""
    if not division:
        division = get_user_division()
    
    division_clause = ""
    hq_clause = ""
    params = {"term": f"%{term}%"}
    if division and division != "Both":
        division_clause = "AND (division = %(division)s OR division = 'Both')"
        params["division"] = division
    if hq:
        hq_clause = "AND hq = %(hq)s"
        params["hq"] = hq
    
    doctors = frappe.db.sql("""
        SELECT name, doctor_code, doctor_name, place, specialization, hospital_address, hq, team, region
        FROM `tabDoctor Master`
        WHERE status = 'Active'
        {division_clause}
        {hq_clause}
        AND (doctor_name LIKE %(term)s OR doctor_code LIKE %(term)s OR place LIKE %(term)s)
        ORDER BY doctor_name
        LIMIT 20
    """.format(division_clause=division_clause, hq_clause=hq_clause), params, as_dict=True)
    
    return doctors

@frappe.whitelist()
def create_scheme_request(data):
    """Create a new scheme request from portal"""
    try:
        if isinstance(data, str):
            data = json.loads(data)
        
        # Create scheme document
        doc = frappe.get_doc({
            "doctype": "Scheme Request",
            "application_date": data.get("application_date"),
            "hq": data.get("hq"),
            "doctor_code": data.get("doctor_code"),
            "stockist_code": data.get("stockist_code"),
            "chemist": data.get("chemist"),
            "scheme_notes": data.get("scheme_notes"),
            "requestedby": frappe.session.user,
            "approval_status": "Pending"
        })
        
        # Add items
        for item in data.get("items", []):
            doc.append("items", {
                "product_code": item.get("product_code"),
                "quantity": item.get("quantity"),
                "free_quantity": item.get("free_quantity"),
                "special_rate": item.get("special_rate")
            })
        
        doc.insert(ignore_permissions=False)
        frappe.db.commit()
        
        return {
            "success": True,
            "name": doc.name,
            "message": "Scheme request created successfully"
        }
        
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Create Scheme Request Error")
        return {
            "success": False,
            "message": str(e)
        }
@frappe.whitelist()
def get_master_data(doctype, division=None, status=None):
    """Get records for a master doctype, all filtered by division where applicable."""
    try:
        filters = {}
        # All masters (Region, Team, Zone, State, HQ, Product, Doctor, Stockist) are division-scoped
        if doctype in ["Region Master", "Team Master", "Zone Master", "State Master", "HQ Master", "Product Master", "Doctor Master", "Stockist Master"]:
            if division:
                filters["division"] = ["in", [division, "Both"]]
        meta = frappe.get_meta(doctype)
        excluded_fieldtypes = [
            "Table", "HTML", "Button", "Column Break", "Section Break",
            "Tab Break", "Heading", "Image"
        ]

        fields = [
            f.fieldname for f in meta.fields
            if f.fieldtype not in excluded_fieldtypes
        ]
        fields = ["name"] + fields
        data = frappe.get_all(
            doctype,
            filters=filters,
            fields=fields,
            order_by="modified desc"
        )

        # Enrich link codes into human-readable labels for the frontend list view table
        for r in data:
            if "hq" in r and r["hq"]:
                r["hq_label"] = frappe.db.get_value("HQ Master", r["hq"], "hq_name") or r["hq"]
            if "team" in r and r["team"]:
                r["team_label"] = frappe.db.get_value("Team Master", r["team"], "team_name") or r["team"]
            if "region" in r and r["region"]:
                r["region_label"] = frappe.db.get_value("Region Master", r["region"], "region_name") or r["region"]
            if "zone" in r and r["zone"]:
                zone_val = frappe.db.get_value("Zone Master", r["zone"], "zone_name")
                r["zone_label"] = zone_val if zone_val else r["zone"]
            if "state" in r and r["state"]:
                state_val = frappe.db.get_value("State Master", r["state"], "state_name")
                r["state_label"] = state_val if state_val else r["state"]

        return {"success": True, "data": data}
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Get Master Data Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def save_master_record(doctype, name, data):
    """
    Save a master record.

    Region/Team/Zone/State/HQ Masters now use auto-generated codes (R0001, T0001, Z0001, ST0001, HQ0001).
    Division is injected from session when not provided.
    """
    try:
        data = frappe.parse_json(data) if isinstance(data, str) else data

        current_user_division = get_user_division()

        # Inject division for all division-scoped masters
        if doctype in ['HQ Master', 'Product Master', 'Doctor Master', 'Stockist Master', 'Region Master', 'Team Master', 'Zone Master', 'State Master']:
            if not data.get('division'):
                data['division'] = current_user_division

        # Strip auto-generated code fields — backend sets them via before_save hook
        for _dt, _cf in [
            ('HQ Master', 'hq_code'),
            ('Doctor Master', 'doctor_code'),
            ('Region Master', 'region_code'),
            ('Team Master', 'team_code'),
            ('Zone Master', 'zone_code'),
            ('State Master', 'state_code'),
        ]:
            if doctype == _dt:
                data.pop(_cf, None)

        # Stockist Code is system-generated for new records
        if doctype == 'Stockist Master' and not name:
            data.pop('stockist_code', None)

        # sanctioned_strength is auto-computed — never let frontend override it
        if doctype == 'Team Master':
            data.pop('sanctioned_strength', None)

        # Division uses autoname=field:division_name.
        # Accept legacy payloads that may send `name`.
        if doctype == 'Division':
            if data.get('name') and not data.get('division_name'):
                data['division_name'] = data.pop('name')
            if not name and not data.get('division_name'):
                return {'success': False, 'message': "Division Name is required"}

        # -----------------------------
        # CREATE OR LOAD DOC
        # -----------------------------
        if name:
            doc = frappe.get_doc(doctype, name)
        else:
            doc = frappe.new_doc(doctype)

        # -----------------------------
        # RESOLVE LINK FIELDS
        # If the frontend sends a display label instead of the actual doc name, resolve it.
        # For Region and Team, lookup needs to consider division as well.
        # -----------------------------

        meta = frappe.get_meta(doctype)
        for field, value in list(data.items()):
            df = meta.get_field(field)
            if not df or df.fieldtype != 'Link' or not value:
                continue

            linked_doctype = df.options
            
            # Skip if value is already a valid doc name
            if frappe.db.exists(linked_doctype, value):
                continue

            # Try to resolve by _name label field, considering division for composite keys
            label_field = next(
                (f.fieldname for f in frappe.get_meta(linked_doctype).fields
                 if f.fieldname.endswith('_name')),
                None
            )
            
            if label_field:
                lookup_filters = {label_field: value}

                # For division-scoped doctypes, also filter by division for accurate match
                if linked_doctype in ["Region Master", "Team Master", "Zone Master", "State Master", "HQ Master"]:
                    lookup_division = data.get('division') or current_user_division
                    if lookup_division:
                        lookup_filters["division"] = lookup_division

                match = frappe.db.get_value(linked_doctype, lookup_filters, 'name')
                if match:
                    data[field] = match

        # -----------------------------
        # APPLY DATA & SAVE
        # -----------------------------
        doc.update(data)
        doc.save(ignore_permissions=False)
        frappe.db.commit()

        # After saving an HQ, recalculate the linked team's sanctioned strength
        if doctype == 'HQ Master':
            team_name = doc.get('team') or data.get('team')
            if team_name:
                _recalculate_team_sanctioned_strength(team_name)

        return {
            'success': True,
            'message': 'Record saved successfully',
            'name': doc.name
        }

    except Exception as e:
        frappe.log_error(frappe.get_traceback(), 'Save Master Record Error')
        return {'success': False, 'message': str(e)}


@frappe.whitelist()
def delete_master_record(doctype, name):
    """Delete a master record"""
    try:
        # Before deleting an HQ, capture its team so we can recalculate
        team_name = None
        if doctype == 'HQ Master':
            team_name = frappe.db.get_value('HQ Master', name, 'team')

        frappe.delete_doc(doctype, name, ignore_permissions=False)
        frappe.db.commit()

        # Recalculate team strength after HQ deletion
        if team_name:
            _recalculate_team_sanctioned_strength(team_name)
        
        return {"success": True, "message": "Record deleted successfully"}
    
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Delete Master Record Error")
        return {"success": False, "message": str(e)}


def _recalculate_team_sanctioned_strength(team_name):
    """Sum per_capita across all HQ Masters linked to this team and update sanctioned_strength."""
    try:
        result = frappe.db.sql(
            "SELECT COALESCE(SUM(per_capita), 0) FROM `tabHQ Master` WHERE team = %s",
            team_name
        )
        total = int(result[0][0]) if result else 0
        frappe.db.set_value('Team Master', team_name, 'sanctioned_strength', total)
        frappe.db.commit()
    except Exception:
        frappe.log_error(frappe.get_traceback(), "Recalculate Team Sanctioned Strength Error")


@frappe.whitelist()
def recalculate_team_sanctioned_strength(team_name):
    """Public API to trigger sanctioned_strength recalculation for a team."""
    try:
        _recalculate_team_sanctioned_strength(team_name)
        strength = frappe.db.get_value('Team Master', team_name, 'sanctioned_strength') or 0
        return {"success": True, "sanctioned_strength": strength}
    except Exception as e:
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def get_hq_list(division=None, search=""):
    """Get HQ list filtered by division for dropdown selection in Stockist/Doctor forms."""
    try:
        filters = {"status": "Active"}
        if division:
            filters["division"] = ["in", [division, "Both"]]
        if search:
            filters["hq_name"] = ["like", f"%{search}%"]

        hqs = frappe.get_all(
            "HQ Master",
            filters=filters,
            fields=["name", "hq_name", "team", "region", "zone", "division"],
            order_by="hq_name asc",
            limit=50
        )

        # Resolve zone code (Z0001) to zone_name for display in text fields
        for hq in hqs:
            zone_code = hq.get("zone")
            if zone_code:
                # Check if zone_code looks like a Zone Master auto-code (starts with Z)
                zone_name = frappe.db.get_value("Zone Master", zone_code, "zone_name")
                if zone_name:
                    hq["zone"] = zone_name  # replace code with display name

        return {"success": True, "data": hqs}
    except Exception as e:
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def get_region_list(division=None, search=""):
    """Get Region list filtered by division for dropdown selection.
    Returns auto-code as value, region_name as label.
    Also resolves zone/state codes to display names for the list view.
    """
    try:
        filters = {"status": "Active"}
        if division:
            filters["division"] = ["in", [division, "Both"]]
        if search:
            filters["region_name"] = ["like", f"%{search}%"]

        regions = frappe.get_all(
            "Region Master",
            filters=filters,
            fields=["name", "region_name", "region_code", "zone", "state", "division"],
            order_by="region_name asc",
            limit=100
        )

        # Resolve zone/state codes to display names for the list table
        for region in regions:
            if region.get("zone"):
                region["zone_name"] = frappe.db.get_value("Zone Master", region["zone"], "zone_name") or region["zone"]
            else:
                region["zone_name"] = None
            if region.get("state"):
                region["state_name"] = frappe.db.get_value("State Master", region["state"], "state_name") or region["state"]
            else:
                region["state_name"] = None

        return {"success": True, "data": regions}
    except Exception as e:
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def get_team_list(division=None, search=""):
    """Get Team list filtered by division for dropdown selection in HQ form.
    Returns auto-code name as value, team_name as label.
    Also returns the linked region so the HQ form can auto-fill region and zone.
    """
    try:
        filters = {"status": "Active"}
        if division:
            filters["division"] = ["in", [division, "Both"]]
        if search:
            filters["team_name"] = ["like", f"%{search}%"]

        teams = frappe.get_all(
            "Team Master",
            filters=filters,
            fields=["name", "team_name", "region", "division"],
            order_by="team_name asc",
            limit=100
        )

        # Enrich each team with region_name and zone_name (from Region Master / Zone Master)
        for team in teams:
            if team.get("region"):
                region_data = frappe.db.get_value(
                    "Region Master", team["region"],
                    ["region_name", "zone"], as_dict=True
                )
                team["region_name"] = region_data.get("region_name") if region_data else None
                zone_code = region_data.get("zone") if region_data else None
                # Resolve zone code -> zone_name for display in text fields
                team["zone"] = zone_code  # Z0001 — used by zone_select dropdown
                team["zone_name"] = frappe.db.get_value("Zone Master", zone_code, "zone_name") if zone_code else None
            else:
                team["region_name"] = None
                team["zone"] = None
                team["zone_name"] = None

        return {"success": True, "data": teams}
    except Exception as e:
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def get_team_details(team_name):
    """Get Team details including linked region and zone for HQ form auto-population."""
    try:
        if not team_name:
            return {"success": False, "message": "Team name is required"}
        doc = frappe.get_doc("Team Master", team_name)
        region_name = None
        zone_code = None
        zone_name = None
        if doc.region:
            region_data = frappe.db.get_value(
                "Region Master", doc.region,
                ["region_name", "zone"], as_dict=True
            )
            if region_data:
                region_name = region_data.get("region_name")
                zone_code = region_data.get("zone")
                zone_name = frappe.db.get_value("Zone Master", zone_code, "zone_name") if zone_code else None
        return {
            "success": True,
            "data": {
                "team_name": doc.team_name,
                "team_id": doc.name,
                "region": doc.region,           # auto-code: "R0001"
                "region_name": region_name,     # display: "South Region"
                "zone": zone_code,              # auto-code: "Z0001" (for zone_select dropdown)
                "zone_name": zone_name,         # display text for Data zone fields
                "division": doc.division
            }
        }
    except Exception as e:
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def get_hq_details(hq_name):
    """Get HQ details including linked team and region for auto-population in forms."""
    try:
        if not hq_name:
            return {"success": False, "message": "HQ name is required"}
        doc = frappe.get_doc("HQ Master", hq_name)
        # Resolve display labels
        team_label = frappe.db.get_value("Team Master", doc.team, "team_name") if doc.team else None
        region_label = frappe.db.get_value("Region Master", doc.region, "region_name") if doc.region else None
        # Resolve zone code -> zone_name for display in Data fields
        zone_display = doc.zone
        if doc.zone:
            zone_name_val = frappe.db.get_value("Zone Master", doc.zone, "zone_name")
            if zone_name_val:
                zone_display = zone_name_val
        return {
            "success": True,
            "data": {
                "hq_name": doc.hq_name,
                "hq_id": doc.name,
                "team": doc.team,
                "team_label": team_label or doc.team,
                "region": doc.region,
                "region_label": region_label or doc.region,
                "zone": zone_display,  # human-readable zone_name for display
                "division": doc.division
            }
        }
    except Exception as e:
        return {"success": False, "message": str(e)}

@frappe.whitelist(allow_guest=True)
def portal_link_search(doctype=None, search=""):
    import json

    # 1️⃣ Try normal frappe args (form encoded)
    doctype = doctype or frappe.form_dict.get("doctype")
    search = search or frappe.form_dict.get("search", "")

    # 2️⃣ Try raw JSON body (portal fetch)
    if not doctype and frappe.request.data:
        try:
            raw = frappe.request.data.decode("utf-8")
            data = json.loads(raw)
            doctype = data.get("doctype")
            search = data.get("search", "")
        except:
            pass

    # 3️⃣ Still missing → error
    if not doctype:
        frappe.throw("doctype is required")

    meta = frappe.get_meta(doctype)

    # Find best display field (e.g. region_name, team_name, hq_name)
    display_field = next(
        (f.fieldname for f in meta.fields if f.fieldname.endswith("_name")),
        "name"
    )

    # For all division-scoped masters, filter by the current user's division.
    filters = {display_field: ["like", f"%{search}%"]}
    if doctype in ("Region Master", "Team Master", "Zone Master", "State Master", "HQ Master"):
        try:
            current_division = get_user_division()
            if current_division:
                filters["division"] = ["in", [current_division, "Both"]]
        except Exception:
            pass

    results = frappe.get_all(
        doctype,
        filters=filters,
        fields=["name", display_field],
        limit=15
    )

    # label = human-readable name (region_name, zone_name, etc.), value = auto-code (R0001, Z0001, etc.)
    return [
        {"label": r.get(display_field) or r.name, "value": r.name}
        for r in results
    ]

@frappe.whitelist()
def search_hq_targets(search="", division=None, limit=15):
    """HQ search helper for Sales Target portal entry."""
    search = (search or "").strip()

    division = division or get_user_division()
    limit = int(limit) if str(limit).isdigit() else 15
    limit = min(max(limit, 1), 50)

    filters = {"status": "Active"}
    if division:
        filters["division"] = ["in", [division, "Both"]]

    hq_rows = frappe.get_all(
        "HQ Master",
        filters=filters,
        or_filters=[
            ["name", "like", f"%{search}%"],
            ["hq_name", "like", f"%{search}%"],
        ],
        fields=["name", "hq_name", "team", "region", "zone"],
        order_by="hq_name asc",
        limit_page_length=limit,
    )

    team_ids = list({row.get("team") for row in hq_rows if row.get("team")})
    region_ids = list({row.get("region") for row in hq_rows if row.get("region")})

    team_map = {}
    if team_ids:
        team_map = {
            d.name: d.team_name
            for d in frappe.get_all("Team Master", filters={"name": ["in", team_ids]}, fields=["name", "team_name"])
        }

    region_map = {}
    if region_ids:
        region_map = {
            d.name: d.region_name
            for d in frappe.get_all("Region Master", filters={"name": ["in", region_ids]}, fields=["name", "region_name"])
        }

    for row in hq_rows:
        row["team_name"] = team_map.get(row.get("team"), row.get("team"))
        row["region_name"] = region_map.get(row.get("region"), row.get("region"))

    return hq_rows


@frappe.whitelist()
def get_hq_yearly_target_details(name):
    """Fetch HQ Yearly Target details for portal editing."""
    try:
        doc = frappe.get_doc("HQ Yearly Target", name)
        items = []
        for item in doc.hq_targets:
            hq_name = frappe.db.get_value("HQ Master", item.hq, "hq_name")
            items.append({
                "hq": item.hq,
                "hq_name": hq_name or item.hq,
                "team": item.team,
                "apr": item.apr, "may": item.may, "jun": item.jun,
                "jul": item.jul, "aug": item.aug, "sep": item.sep,
                "oct": item.oct, "nov": item.nov, "dec": item.dec,
                "jan": item.jan, "feb": item.feb, "mar": item.mar,
            })
        return {
            "success": True,
            "doc": {
                "name": doc.name,
                "financial_year": doc.financial_year,
                "start_date": str(doc.start_date),
                "end_date": str(doc.end_date),
                "status": doc.status,
                "docstatus": doc.docstatus,
                "items": items
            }
        }
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Get HQ Yearly Target Details Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def submit_hq_yearly_target_from_portal(name):
    """Submit / Approve an HQ yearly target from the portal."""
    try:
        if not frappe.has_permission("HQ Yearly Target", "submit"):
            return {"success": False, "message": "Not permitted to approve targets"}
        doc = frappe.get_doc("HQ Yearly Target", name)
        if doc.docstatus == 1:
            return {"success": True, "message": "Already approved"}
        doc.submit()
        frappe.db.commit()
        return {"success": True, "message": "Approved successfully"}
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Approve HQ Yearly Target Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def update_hq_yearly_target_from_portal(name, financial_year, start_date, end_date, hq_targets, status="Draft"):
    """Update existing draft HQ Yearly Target from portal."""
    try:
        if not frappe.has_permission("HQ Yearly Target", "write"):
            return {"success": False, "message": "Not permitted"}

        doc = frappe.get_doc("HQ Yearly Target", name)
        if doc.docstatus == 1:
            return {"success": False, "message": "Cannot edit an approved target"}

        if isinstance(hq_targets, str):
            hq_targets = frappe.parse_json(hq_targets)

        doc.financial_year = financial_year
        doc.start_date = start_date
        doc.end_date = end_date
        doc.status = status or "Draft"

        doc.hq_targets = []
        for raw in (hq_targets or []):
            hq = raw.get("hq")
            if not hq:
                continue
            meta = frappe.db.get_value("HQ Master", hq, ["team", "region"], as_dict=True)
            doc.append("hq_targets", {
                "hq": hq,
                "team": meta.get("team") if meta else None,
                "region": meta.get("region") if meta else None,
                "apr": flt(raw.get("apr")),
                "may": flt(raw.get("may")),
                "jun": flt(raw.get("jun")),
                "jul": flt(raw.get("jul")),
                "aug": flt(raw.get("aug")),
                "sep": flt(raw.get("sep")),
                "oct": flt(raw.get("oct")),
                "nov": flt(raw.get("nov")),
                "dec": flt(raw.get("dec")),
                "jan": flt(raw.get("jan")),
                "feb": flt(raw.get("feb")),
                "mar": flt(raw.get("mar")),
            })

        doc.save(ignore_permissions=False)
        frappe.db.commit()
        return {"success": True, "message": "Updated successfully", "name": doc.name}
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Update HQ Yearly Target Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def resolve_hq_target_rows_from_file(division=None):
    """Parse uploaded CSV/Excel bulk target file and resolve HQ names to IDs.
    Required columns: HQ Name, Apr, May, Jun, Jul, Aug, Sep, Oct, Nov, Dec, Jan, Feb, Mar
    Returns resolved rows ready to populate the portal form.
    """
    try:
        import pandas as pd
        import io as _io

        uploaded_file = frappe.request.files.get("file")
        if not uploaded_file:
            return {"success": False, "message": "No file uploaded"}

        if not division:
            division = get_user_division()

        filename = uploaded_file.filename or ""
        ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
        content = uploaded_file.read()

        if ext == "csv":
            df = pd.read_csv(_io.BytesIO(content), dtype=str)
        elif ext in ("xlsx", "xls"):
            df = pd.read_excel(_io.BytesIO(content), dtype=str)
        else:
            try:
                df = pd.read_csv(_io.BytesIO(content), dtype=str)
            except Exception:
                df = pd.read_excel(_io.BytesIO(content), dtype=str)

        df.columns = df.columns.str.strip()

        hq_filters = {"status": "Active"}
        if division:
            hq_filters["division"] = ["in", [division, "Both"]]
        all_hqs = frappe.get_all("HQ Master", filters=hq_filters, fields=["name", "hq_name"])
        hq_by_name = {h.hq_name.strip().lower(): h for h in all_hqs}
        hq_by_code = {h.name.strip().lower(): h for h in all_hqs}

        resolved = []
        errors = []
        months = ["apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec", "jan", "feb", "mar"]

        for idx, row in df.iterrows():
            hq_input = str(row.get("HQ Name", "") or "").strip()
            if not hq_input or hq_input.lower() in ("hq name", "hq"):
                continue  # skip header-like rows

            hq_rec = hq_by_name.get(hq_input.lower()) or hq_by_code.get(hq_input.lower())
            if not hq_rec:
                errors.append(f"Row {idx + 2}: HQ '{hq_input}' not found")
                continue

            result_row = {"hq": hq_rec.name, "hq_name": hq_rec.hq_name}
            for m in months:
                col_label = m.capitalize()
                val = row.get(col_label, 0) or row.get(m.upper(), 0) or row.get(m, 0) or 0
                try:
                    result_row[m] = float(val) if str(val).strip() not in ("nan", "None", "") else 0.0
                except (ValueError, TypeError):
                    result_row[m] = 0.0
            resolved.append(result_row)

        return {"success": True, "rows": resolved, "errors": errors}
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Resolve HQ Target Rows Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def create_hq_yearly_target_from_portal(financial_year, start_date, end_date, status="Draft", hq_targets=None):
    """Create HQ Yearly Target with HQ-wise monthly values from portal screen."""
    try:
        if not frappe.has_permission("HQ Yearly Target", "create"):
            frappe.throw(_("Not permitted to create HQ Yearly Target"))

        if isinstance(hq_targets, str):
            hq_targets = frappe.parse_json(hq_targets)

        if not hq_targets or not isinstance(hq_targets, list):
            frappe.throw(_("At least one HQ target row is required"))

        division = get_user_division()
        unique_hqs = list({(row or {}).get("hq") for row in hq_targets if (row or {}).get("hq")})
        if not unique_hqs:
            frappe.throw(_("At least one valid HQ is required"))

        hq_meta = frappe.get_all(
            "HQ Master",
            filters={"name": ["in", unique_hqs]},
            fields=["name", "team", "region", "division"],
            limit_page_length=0,
        )
        hq_map = {row.name: row for row in hq_meta}

        parsed_rows = []
        seen_hq = set()
        region_set = set()

        for row in hq_targets:
            row = row or {}
            hq = row.get("hq")
            if not hq:
                continue

            if hq in seen_hq:
                frappe.throw(_("Duplicate HQ found: {0}").format(hq))
            seen_hq.add(hq)

            meta = hq_map.get(hq)
            if not meta:
                frappe.throw(_("Invalid HQ selected: {0}").format(hq))

            if division and meta.get("division") not in [division, "Both"]:
                frappe.throw(_("HQ {0} does not belong to selected division {1}").format(hq, division))

            region = meta.get("region")
            if not region:
                frappe.throw(_("Region is missing in HQ Master for {0}").format(hq))
            region_set.add(region)

            parsed_rows.append({
                "hq": hq,
                "team": meta.get("team"),
                "apr": flt(row.get("apr")),
                "may": flt(row.get("may")),
                "jun": flt(row.get("jun")),
                "jul": flt(row.get("jul")),
                "aug": flt(row.get("aug")),
                "sep": flt(row.get("sep")),
                "oct": flt(row.get("oct")),
                "nov": flt(row.get("nov")),
                "dec": flt(row.get("dec")),
                "jan": flt(row.get("jan")),
                "feb": flt(row.get("feb")),
                "mar": flt(row.get("mar")),
            })

        if not parsed_rows:
            frappe.throw(_("At least one valid HQ target row is required"))
        if len(region_set) != 1:
            frappe.throw(_("All selected HQs must belong to one region"))

        region = list(region_set)[0]
        series = "HQT-.division.-.YYYY.-"
        doc = frappe.get_doc({
            "doctype": "HQ Yearly Target",
            "naming_series": series,
            "division": division,
            "region": region,
            "financial_year": financial_year,
            "start_date": start_date,
            "end_date": end_date,
            "status": status or "Draft",
            "hq_targets": parsed_rows,
        })
        doc.insert(ignore_permissions=False)
        frappe.db.commit()

        return {
            "success": True,
            "name": doc.name,
            "total_hqs": doc.total_hqs,
            "total_target_amount": doc.total_target_amount,
        }
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Create HQ Yearly Target From Portal Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def import_master_data(doctype, division):
    """Bulk import master data from Excel or CSV file.

    - Resolves link field values (names → codes) within the correct division.
    - Performs upsert: if a record with the same identifying fields + division
      already exists, it updates instead of creating a duplicate.
    - Code fields (hq_code, zone_code, etc.) are auto-generated — never required
      from the user.  Product Code is the only user-entered code.
    - Returns per-row error details so the portal can display them.
    """
    try:
        import pandas as pd

        file = frappe.request.files.get('file')
        if not file:
            return {"success": False, "message": "No file uploaded"}

        filename = file.filename or ""
        ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""

        if ext == "csv":
            df = pd.read_csv(file, dtype=str)
        elif ext in ("xlsx", "xls"):
            df = pd.read_excel(file, dtype=str)
        else:
            try:
                import io
                content = file.read()
                if hasattr(file, 'seek'):
                    file.seek(0)
                df = pd.read_csv(io.BytesIO(content), dtype=str)
            except Exception:
                if hasattr(file, 'seek'):
                    file.seek(0)
                df = pd.read_excel(file, dtype=str)

        df.columns = df.columns.str.strip()

        column_mapping = get_column_mapping(doctype)

        imported = 0
        updated = 0
        failed = 0
        errors = []

        for idx, row in df.iterrows():
            row_num = idx + 2  # Excel row (1-indexed header + 1)
            try:
                data = {}
                for excel_col, field_name in column_mapping.items():
                    if excel_col in row and pd.notna(row[excel_col]):
                        val = str(row[excel_col]).strip()
                        if val:
                            data[field_name] = val

                # ---- Inject division ----
                row_division = data.get("division") or division
                if doctype != "Division":
                    data["division"] = row_division

                # ---- Normalize Select field values (case-insensitive match) ----
                _normalize_select_fields(doctype, data)

                # ---- Resolve link-field labels → doc names ----
                _resolve_import_links(doctype, data, row_division)

                # ---- Strip auto-generated code fields (backend creates them) ----
                for _dt, _cf in [
                    ("HQ Master", "hq_code"),
                    ("Doctor Master", "doctor_code"),
                    ("Region Master", "region_code"),
                    ("Team Master", "team_code"),
                    ("Zone Master", "zone_code"),
                    ("State Master", "state_code"),
                    ("Stockist Master", "stockist_code"),
                ]:
                    if doctype == _dt:
                        data.pop(_cf, None)

                # sanctioned_strength is auto-computed
                if doctype == "Team Master":
                    data.pop("sanctioned_strength", None)

                # ---- Upsert: find existing record by name-field + division ----
                existing = _find_existing_master(doctype, data, row_division)

                if existing:
                    doc = frappe.get_doc(doctype, existing)
                    doc.update(data)
                    doc.save(ignore_permissions=True)
                    updated += 1
                else:
                    doc = frappe.get_doc({"doctype": doctype, **data})
                    doc.insert(ignore_permissions=True)
                    imported += 1

                if (imported + updated) % 50 == 0:
                    frappe.db.commit()

            except Exception as e:
                failed += 1
                err_msg = str(e)
                # Make validation errors more readable
                if "already exists in division" in err_msg:
                    errors.append(f"Row {row_num}: {err_msg}")
                else:
                    # Strip HTML tags from frappe error messages
                    import re as _re
                    clean = _re.sub(r"<[^>]+>", "", err_msg).strip()
                    errors.append(f"Row {row_num}: {clean}")

        frappe.db.commit()

        return {
            "success": True,
            "imported": imported,
            "updated": updated,
            "failed": failed,
            "errors": errors[:50]  # Return first 50 errors for visibility
        }

    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Import Master Data Error")
        return {"success": False, "message": str(e)}


def _normalize_select_fields(doctype, data):
    """Case-insensitively match imported Select values to their valid options."""
    meta = frappe.get_meta(doctype)
    for field in meta.fields:
        if field.fieldtype != "Select" or not field.options:
            continue
        val = data.get(field.fieldname)
        if not val:
            continue
        valid_options = [o for o in field.options.split("\n") if o]
        val_lower = val.lower()
        for opt in valid_options:
            if opt.lower() == val_lower:
                data[field.fieldname] = opt  # replace with correctly-cased version
                break


def _resolve_import_links(doctype, data, division):
    """Resolve display-name values in import data to actual doc names.

    E.g. Team='Team Alpha' → Team='T0001' (the actual name of that Team Master
    record in the given division).
    """
    # Mapping: data field → (linked doctype, label field)
    link_fields = {
        "team":   ("Team Master",   "team_name"),
        "region": ("Region Master", "region_name"),
        "zone":   ("Zone Master",   "zone_name"),
        "state":  ("State Master",  "state_name"),
        "hq":     ("HQ Master",     "hq_name"),
    }

    for field, (linked_dt, label_field) in link_fields.items():
        val = data.get(field)
        if not val:
            continue

        # Already a valid doc name?
        if frappe.db.exists(linked_dt, val):
            continue

        # Try to resolve by label within the same division
        filters = {label_field: val}
        linked_meta = frappe.get_meta(linked_dt)
        if linked_meta.has_field("division") and division:
            filters["division"] = division

        match = frappe.db.get_value(linked_dt, filters, "name")
        if match:
            data[field] = match


def _find_existing_master(doctype, data, division):
    """Find an existing master record for upsert based on the identifying
    name field + division combination."""
    upsert_keys = {
        "Zone Master":     "zone_name",
        "State Master":    "state_name",
        "Region Master":   "region_name",
        "Team Master":     "team_name",
        "HQ Master":       "hq_name",
        "Product Master":  "product_code",
        "Stockist Master": "stockist_name",
        "Doctor Master":   "doctor_name",
    }

    name_field = upsert_keys.get(doctype)
    if not name_field or not data.get(name_field):
        return None

    filters = {name_field: data[name_field]}

    # Division-scoped lookup
    meta = frappe.get_meta(doctype)
    if meta.has_field("division") and division:
        filters["division"] = division

    return frappe.db.get_value(doctype, filters, "name")


def get_column_mapping(doctype):
    """Get Excel column to field mapping for each doctype"""
    mappings = {
        "HQ Master": {
            "HQ Name": "hq_name",
            "Team": "team",
            "Region": "region",
            "Zone": "zone",
            "Per Capita": "per_capita",
            "Division": "division",
            "Status": "status"
        },
        "Stockist Master": {
            "Stockist Name": "stockist_name",
            "HQ": "hq",
            "Team": "team",
            "Region": "region",
            "Zone": "zone",
            "Address": "address",
            "Contact Person": "contact_person",
            "Phone": "phone",
            "Email": "email",
            "Status": "status"
        },
        "Product Master": {
            "Product Code": "product_code",
            "Product Name": "product_name",
            "Sequence": "sequence",
            "Product Group": "product_group",
            "Category": "category",
            "Pack": "pack",
            "Pack Conversion": "pack_conversion",
            "PTS": "pts",
            "PTR": "ptr",
            "MRP": "mrp",
            "GST Rate (%)": "gst_rate",
            "Division": "division",
            "Status": "status"
        },
        "Doctor Master": {
            # Doctor Code is auto-generated — skip from import mapping
            "Doctor Name": "doctor_name",
            "Qualification": "qualification",
            "Doctor Category": "doctor_category",
            "Specialization": "specialization",
            "Phone": "phone",
            "Place": "place",
            "Hospital Address": "hospital_address",
            "House Address": "house_address",
            "Division": "division",
            "HQ": "hq",
            "Team": "team",
            "Region": "region",
            "State": "state",
            "Zone": "zone",
            "Chemist Name": "chemist_name",
            "Status": "status"
        },
        "Team Master": {
            "Team Name": "team_name",
            "Region": "region",
            "Division": "division",
            "Status": "status"
            # sanctioned_strength is auto-computed - not in import
            # team_code is auto-generated - not in import
        },
        "Region Master": {
            "Region Name": "region_name",
            "Division": "division",
            "Zone": "zone",
            "State": "state",
            "Status": "status"
            # region_code is auto-generated - not in import
        },
        "Zone Master": {
            "Zone Name": "zone_name",
            "Division": "division",
            "Status": "status"
            # zone_code is auto-generated - not in import
        },
        "State Master": {
            "State Name": "state_name",
            "Division": "division",
            "Status": "status"
            # state_code is auto-generated - not in import
        }
    }
    return mappings.get(doctype, {})


@frappe.whitelist()
def get_zone_list(division=None, search=""):
    """Get all Zone Master records for dropdown population."""
    try:
        filters = {"status": "Active"}
        if division:
            filters["division"] = ["in", [division, "Both"]]
        if search:
            filters["zone_name"] = ["like", f"%{search}%"]
        zones = frappe.get_all(
            "Zone Master",
            filters=filters,
            fields=["name", "zone_name"],
            order_by="zone_name asc",
            limit=100
        )
        return {"success": True, "data": zones}
    except Exception as e:
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def get_state_list(division=None, search=""):
    """Get all State Master records for dropdown population."""
    try:
        filters = {"status": "Active"}
        if division:
            filters["division"] = ["in", [division, "Both"]]
        if search:
            filters["state_name"] = ["like", f"%{search}%"]
        states = frappe.get_all(
            "State Master",
            filters=filters,
            fields=["name", "state_name"],
            order_by="state_name asc",
            limit=100
        )
        return {"success": True, "data": states}
    except Exception as e:
        return {"success": False, "message": str(e)}


def get_code_field(doctype):
    """Legacy helper — kept for backward compatibility but no longer used by import."""
    code_fields = {
        "HQ Master": "hq_name",
        "Team Master": "team_name",
    }
    return code_fields.get(doctype)

@frappe.whitelist()
def searchstockists(searchterm=None, division=None, limit=20):
    searchterm = (searchterm or "").strip()

    limit = int(limit) if str(limit).isdigit() else 20
    limit = min(max(limit, 1), 50)

    # Basic filters
    filters = {"status": "Active"}
    if division:
        filters["division"] = ["in", [division, "Both"]]

    # Match by code or name
    # NOTE: adapt fieldnames if your doctype uses stockist_code/stockist_name exactly
    stockists = frappe.get_all(
        "Stockist Master",
        filters=filters,
        fields=["stockist_code", "stockist_name", "hq", "division"],
        or_filters=[
            ["stockist_code", "like", f"%{searchterm}%"],
            ["stockist_name", "like", f"%{searchterm}%"],
        ],
        limit_page_length=limit,
        order_by="stockist_name asc"
    )

    # Enrich with hierarchy from HQ Master (if available)
    for st in stockists:
        team = region = zone = None
        if st.get("hq"):
            team, region, zone = frappe.db.get_value(
                "HQ Master",
                st["hq"],
                ["team", "region", "zone"]
            ) or (None, None, None)
        st["team"] = team
        st["region"] = region
        st["zone"] = zone

    return stockists

@frappe.whitelist()
def get_scheme_detail(scheme_name):
    """Get full scheme request details for portal view"""
    try:
        doc = frappe.get_doc("Scheme Request", scheme_name)
        items = []
        for item in doc.items:
            items.append({
                "product_code": item.product_code,
                "product_name": item.product_name,
                "pack": item.pack,
                "quantity": flt(item.quantity),
                "free_quantity": flt(item.free_quantity),
                "product_rate": flt(item.product_rate),
                "special_rate": flt(item.special_rate),
                "scheme_percentage": flt(item.scheme_percentage),
                "product_value": flt(item.product_value),
            })
        logs = []
        for log in (doc.approval_log or []):
            logs.append({
                "approver": log.approver,
                "action": log.action,
                "action_date": str(log.action_date) if log.action_date else "",
                "comments": log.comments,
                "approval_level": log.approval_level,
            })
        return {
            "success": True,
            "name": doc.name,
            "application_date": str(doc.application_date) if doc.application_date else "",
            "doctor_code": doc.doctor_code,
            "doctor_name": doc.doctor_name,
            "doctor_place": doc.doctor_place,
            "specialization": doc.specialization,
            "hospital_address": doc.hospital_address,
            "hq": doc.hq,
            "region": doc.region,
            "team": doc.team,
            "stockist_code": doc.stockist_code,
            "stockist_name": doc.stockist_name,
            "approval_status": doc.approval_status,
            "total_scheme_value": flt(doc.total_scheme_value),
            "scheme_notes": doc.scheme_notes,
            "division": doc.division,
            "requested_by": doc.requested_by,
            "proof_attachment_1": doc.proof_attachment_1,
            "proof_attachment_2": doc.proof_attachment_2,
            "proof_attachment_3": doc.proof_attachment_3,
            "proof_attachment_4": doc.proof_attachment_4,
            "items": items,
            "approval_log": logs,
            "docstatus": doc.docstatus,
        }
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Get Scheme Detail Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def get_scheme_requests_for_deduction(division=None, search=""):
    """Get approved+submitted scheme requests for deduction selection"""
    try:
        if not division:
            division = get_user_division()
        search = (search or "").strip()
        rows = frappe.db.sql("""
            SELECT
                sr.name,
                sr.application_date,
                sr.doctor_name,
                sr.doctor_code,
                sr.stockist_code,
                sr.stockist_name,
                sr.hq,
                sr.total_scheme_value,
                sr.approval_status,
                sr.division
            FROM `tabScheme Request` sr
            WHERE sr.docstatus = 1
              AND sr.division = %(division)s
              AND (
                sr.name LIKE %(search)s
                OR sr.doctor_name LIKE %(search)s
                OR sr.stockist_name LIKE %(search)s
              )
            ORDER BY sr.application_date DESC
            LIMIT 50
        """, {"division": division, "search": f"%{search}%"}, as_dict=True)
        return {"success": True, "data": rows}
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Get Scheme Requests For Deduction Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def get_stockist_statements_for_deduction(stockist_code, division=None):
    """Get stockist statements for a stockist for deduction"""
    try:
        if not division:
            division = get_user_division()
        statements = frappe.db.sql("""
            SELECT
                ss.name,
                DATE_FORMAT(ss.statement_month, '%%b-%%Y') as month_label,
                ss.statement_month,
                sm.stockist_name
            FROM `tabStockist Statement` ss
            LEFT JOIN `tabStockist Master` sm ON ss.stockist_code = sm.name
            WHERE ss.stockist_code = %(stockist_code)s
              AND ss.docstatus != 2
              AND (
                ss.division IS NULL
                OR ss.division = %(division)s
                OR ss.division = 'Both'
              )
            ORDER BY ss.statement_month DESC
            LIMIT 24
        """, {"stockist_code": stockist_code, "division": division}, as_dict=True)
        return {"success": True, "data": statements}
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Get Statements For Deduction Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def fetch_deduction_items_portal(scheme_request, stockist_statement, division=None):
    """Fetch items for deduction from scheme + statement (portal version)"""
    try:
        if not division:
            division = get_user_division()
        scheme = frappe.get_doc("Scheme Request", scheme_request)
        statement = frappe.get_doc("Stockist Statement", stockist_statement)

        if getattr(scheme, "division", None) and division and scheme.division not in [division, "Both"]:
            return {"success": False, "message": f"Scheme {scheme_request} does not belong to division {division}"}
        if getattr(statement, "division", None) and division and statement.division not in [division, "Both"]:
            return {"success": False, "message": f"Statement {stockist_statement} does not belong to division {division}"}

        # Validate stockist match
        if scheme.stockist_code != statement.stockist_code:
            return {
                "success": False,
                "message": f"Stockist mismatch: Scheme ({scheme.stockist_code}) vs Statement ({statement.stockist_code})"
            }

        # Build statement product map — use sales_qty for deduction display
        stmt_map = {item.product_code: flt(item.sales_qty) for item in statement.items}

        items = []
        skipped = []
        for scheme_item in scheme.items:
            if scheme_item.product_code not in stmt_map:
                skipped.append(scheme_item.product_code)
                continue
            product = frappe.get_doc("Product Master", scheme_item.product_code)
            scheme_free_qty = flt(scheme_item.free_quantity)
            current_sales_qty = stmt_map[scheme_item.product_code]
            items.append({
                "product_code": scheme_item.product_code,
                "product_name": product.product_name,
                "pack": product.pack,
                "scheme_free_qty": scheme_free_qty,
                "current_sales_qty": current_sales_qty,
                "deduct_qty": scheme_free_qty,
                "pts": flt(product.pts),
                "deducted_value": scheme_free_qty * flt(product.pts),
            })

        return {
            "success": True,
            "items": items,
            "skipped": skipped,
            "scheme_doctor": scheme.doctor_name,
            "scheme_hq": scheme.hq,
        }
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Fetch Deduction Items Portal Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def create_scheme_deduction_portal(scheme_request, stockist_statement, items, deduction_date=None, division=None):
    """Create a Scheme Deduction document from portal"""
    try:
        if isinstance(items, str):
            items = json.loads(items)

        scheme = frappe.get_doc("Scheme Request", scheme_request)
        statement = frappe.get_doc("Stockist Statement", stockist_statement)

        # Check if deduction already exists
        existing = frappe.db.exists("Scheme Deduction", {
            "scheme_request": scheme_request,
            "stockist_statement": stockist_statement,
            "docstatus": ["!=", 2]
        })
        if existing:
            return {"success": False, "message": f"Deduction already exists: {existing}"}

        if not division:
            division = getattr(scheme, "division", None) or get_user_division()
        if getattr(scheme, "division", None) and scheme.division not in [division, "Both"]:
            return {"success": False, "message": f"Scheme {scheme_request} does not belong to division {division}"}
        if getattr(statement, "division", None) and statement.division not in [division, "Both"]:
            return {"success": False, "message": f"Statement {stockist_statement} does not belong to division {division}"}

        doc = frappe.new_doc("Scheme Deduction")
        doc.scheme_request = scheme_request
        doc.stockist_statement = stockist_statement
        doc.stockist_code = scheme.stockist_code
        doc.doctor_code = scheme.doctor_code
        doc.scheme_date = scheme.application_date
        doc.division = division
        if deduction_date:
            doc.deduction_date = deduction_date

        total_qty = 0
        total_value = 0
        for item in items:
            deduct_qty = flt(item.get("deduct_qty", 0))
            pts = flt(item.get("pts", 0))
            deducted_value = deduct_qty * pts
            doc.append("items", {
                "product_code": item.get("product_code"),
                "product_name": item.get("product_name"),
                "pack": item.get("pack"),
                "scheme_free_qty": flt(item.get("scheme_free_qty", 0)),
                "current_free_qty": flt(item.get("current_sales_qty") or item.get("current_free_qty") or 0),
                "deduct_qty": deduct_qty,
                "pts": pts,
                "deducted_value": deducted_value,
            })
            total_qty += deduct_qty
            total_value += deducted_value

        doc.total_deducted_qty = total_qty
        doc.total_deducted_value = total_value
        doc.insert(ignore_permissions=False)
        # Auto-submit the deduction (no draft state needed)
        doc.submit()
        frappe.db.commit()

        return {"success": True, "name": doc.name, "message": "Scheme deduction created successfully"}
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Create Scheme Deduction Portal Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def create_bulk_scheme_deductions_portal(deductions, deduction_date, division=None):
    """Create multiple Scheme Deductions at once from the auto-deduction portal page"""
    try:
        if isinstance(deductions, str):
            deductions = json.loads(deductions)

        if not division:
            division = get_user_division()

        results = []
        for entry in deductions:
            scheme_request = entry.get("scheme_request")
            stockist_statement = entry.get("stockist_statement")
            items = entry.get("items", [])

            if not scheme_request or not stockist_statement:
                results.append({"success": False, "message": "Missing scheme or statement"})
                continue

            result = create_scheme_deduction_portal(
                scheme_request=scheme_request,
                stockist_statement=stockist_statement,
                items=json.dumps(items) if not isinstance(items, str) else items,
                deduction_date=deduction_date,
                division=division,
            )
            results.append(result)

        created_count = sum(1 for r in results if r.get("success"))
        return {
            "success": True,
            "results": results,
            "created": created_count,
            "total": len(deductions),
            "message": f"{created_count} of {len(deductions)} deductions created",
        }
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Create Bulk Scheme Deductions Portal Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def get_scheme_deductions_portal(division=None, search=None, status=None, from_date=None, to_date=None):
    """List scheme deductions for the portal with filters"""
    try:
        if not division:
            division = get_user_division()

        filters = [["Scheme Deduction", "division", "=", division]]

        if search:
            filters = [
                ["Scheme Deduction", "division", "=", division],
                "|",
                ["Scheme Deduction", "name", "like", f"%{search}%"],
                ["Scheme Deduction", "scheme_request", "like", f"%{search}%"],
                ["Scheme Deduction", "stockist_statement", "like", f"%{search}%"],
            ]
            # Use simpler OR approach via SQL
            deductions = frappe.db.sql("""
                SELECT name, scheme_request, stockist_statement, deduction_date, 
                       total_deducted_qty, total_deducted_value, status, docstatus, creation
                FROM `tabScheme Deduction`
                WHERE division = %(division)s
                AND (name LIKE %(s)s OR scheme_request LIKE %(s)s OR stockist_statement LIKE %(s)s)
                ORDER BY creation DESC
                LIMIT 100
            """, {"division": division, "s": f"%{search}%"}, as_dict=True)
        else:
            q_filters = {"division": division}
            if status:
                q_filters["status"] = status
            if from_date:
                q_filters["creation"] = [">=", from_date]
            if to_date:
                if "creation" in q_filters:
                    q_filters["creation"] = ["between", [from_date, to_date]]
                else:
                    q_filters["creation"] = ["<=", to_date]

            deductions = frappe.get_all(
                "Scheme Deduction",
                filters=q_filters,
                fields=["name", "scheme_request", "stockist_statement", "deduction_date",
                        "total_deducted_qty", "total_deducted_value", "status", "docstatus", "creation"],
                order_by="creation desc",
                limit=200
            )

        return {"success": True, "data": deductions}
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Get Scheme Deductions Portal Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def portal_repeat_scheme_request(source_name):
    """Repeat an approved scheme request from portal"""
    try:
        source_doc = frappe.get_doc("Scheme Request", source_name)
        if source_doc.approval_status != "Approved":
            return {"success": False, "message": "Only approved scheme requests can be repeated"}

        new_doc = frappe.new_doc("Scheme Request")
        new_doc.application_date = nowdate()
        new_doc.requested_by = frappe.session.user
        new_doc.team = source_doc.team
        new_doc.region = source_doc.region
        new_doc.hq = source_doc.hq
        new_doc.stockist_code = source_doc.stockist_code
        new_doc.stockist_name = source_doc.stockist_name
        new_doc.doctor_code = source_doc.doctor_code
        new_doc.doctor_name = source_doc.doctor_name
        new_doc.doctor_place = source_doc.doctor_place
        new_doc.specialization = source_doc.specialization
        new_doc.hospital_address = source_doc.hospital_address
        new_doc.scheme_notes = f"Repeated from {source_doc.name}"
        new_doc.approval_status = "Pending"
        new_doc.repeated_request = 1

        for item in source_doc.items:
            new_doc.append("items", {
                "product_code": item.product_code,
                "product_name": item.product_name,
                "pack": item.pack,
                "quantity": item.quantity,
                "free_quantity": item.free_quantity,
                "product_rate": item.product_rate,
                "special_rate": item.special_rate,
                "product_value": item.product_value,
            })

        new_doc.insert()
        frappe.db.commit()
        return {"success": True, "name": new_doc.name, "message": "New scheme request created successfully"}
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Portal Repeat Scheme Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def get_doctor_monthly_limit_info(doctor_code, application_date=None):
    """Get per-product monthly request count for a doctor"""
    try:
        from frappe.utils import getdate, get_first_day, get_last_day
        if not application_date:
            application_date = nowdate()
        app_date = getdate(application_date)
        first_day = get_first_day(app_date)
        last_day = get_last_day(app_date)
        month_name = app_date.strftime("%B %Y")

        rows = frappe.db.sql("""
            SELECT sri.product_code, COUNT(DISTINCT sr.name) as request_count
            FROM `tabScheme Request` sr
            INNER JOIN `tabScheme Request Item` sri ON sr.name = sri.parent
            WHERE sr.doctor_code = %(doctor_code)s
              AND sr.application_date BETWEEN %(first_day)s AND %(last_day)s
              AND sr.docstatus != 2
            GROUP BY sri.product_code
        """, {"doctor_code": doctor_code, "first_day": first_day, "last_day": last_day}, as_dict=True)

        product_counts = {row.product_code: row.request_count for row in rows}
        total_requests = frappe.db.count("Scheme Request", {
            "doctor_code": doctor_code,
            "application_date": ["between", [first_day, last_day]],
            "docstatus": ["!=", 2]
        })

        return {
            "success": True,
            "month": month_name,
            "total_requests": total_requests,
            "product_counts": product_counts,
            "limit_per_product": 3,
        }
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Doctor Monthly Limit Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def create_scheme_request_v2(data):
    """Create a new scheme request from portal (robust version)"""
    try:
        if isinstance(data, str):
            data = json.loads(data)

        doc = frappe.new_doc("Scheme Request")
        doc.application_date = data.get("application_date") or nowdate()
        doc.requested_by = frappe.session.user
        doc.hq = data.get("hq")
        doc.doctor_code = data.get("doctor_code")
        doc.doctor_name = data.get("doctor_name")
        doc.doctor_place = data.get("doctor_place")
        doc.specialization = data.get("specialization")
        doc.hospital_address = data.get("hospital_address")
        doc.team = data.get("team")
        doc.region = data.get("region")
        doc.stockist_code = data.get("stockist_code")
        doc.stockist_name = data.get("stockist_name")
        doc.scheme_notes = data.get("scheme_notes")
        doc.approval_status = "Pending"
        doc.repeated_request = 1 if str(data.get("repeated_request", 0)).lower() in ("1", "true", "yes") else 0

        for item in data.get("items", []):
            if not item.get("product_code"):
                continue
            qty = flt(item.get("quantity", 0))
            free_qty = flt(item.get("free_quantity", 0))
            rate = flt(item.get("product_rate", 0))
            special_rate = flt(item.get("special_rate", 0))

            # Calculate scheme %
            scheme_pct = 0
            if special_rate > 0 and rate > 0:
                scheme_pct = ((rate - special_rate) / rate) * 100
            elif free_qty > 0 and qty > 0:
                scheme_pct = (free_qty / qty) * 100

            effective_rate = special_rate if special_rate > 0 else rate
            product_value = qty * effective_rate

            doc.append("items", {
                "product_code": item.get("product_code"),
                "product_name": item.get("product_name"),
                "pack": item.get("pack"),
                "quantity": qty,
                "free_quantity": free_qty,
                "product_rate": rate,
                "special_rate": special_rate,
                "scheme_percentage": round(scheme_pct, 2),
                "product_value": product_value,
            })

        doc.insert(ignore_permissions=False)
        frappe.db.commit()

        return {"success": True, "name": doc.name, "message": "Scheme request created successfully"}
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Create Scheme Request V2 Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def get_scheme_list_portal(division=None, filters=None):
    """Get schemes for portal list view with division filter"""
    try:
        if not division:
            division = get_user_division()
        if isinstance(filters, str):
            filters = json.loads(filters) if filters else {}
        if not filters:
            filters = {}

        where_clauses = ["sr.division = %(division)s"]
        params = {"division": division}

        if filters.get("status"):
            where_clauses.append("sr.approval_status = %(status)s")
            params["status"] = filters["status"]
        if filters.get("from_date") and filters.get("to_date"):
            where_clauses.append("sr.application_date BETWEEN %(from_date)s AND %(to_date)s")
            params["from_date"] = filters["from_date"]
            params["to_date"] = filters["to_date"]
        elif filters.get("from_date"):
            where_clauses.append("sr.application_date >= %(from_date)s")
            params["from_date"] = filters["from_date"]
        elif filters.get("to_date"):
            where_clauses.append("sr.application_date <= %(to_date)s")
            params["to_date"] = filters["to_date"]
        if filters.get("search"):
            where_clauses.append(
                "(sr.name LIKE %(search)s OR sr.doctor_name LIKE %(search)s "
                "OR sr.stockist_name LIKE %(search)s OR sr.hq LIKE %(search)s)"
            )
            params["search"] = f"%{filters['search']}%"
        if filters.get("request_type") == "Repeated":
            where_clauses.append("COALESCE(sr.repeated_request, 0) = 1")
        elif filters.get("request_type") == "Original":
            where_clauses.append("COALESCE(sr.repeated_request, 0) = 0")

        where_sql = " AND ".join(where_clauses)

        rows = frappe.db.sql(f"""
            SELECT
                sr.name,
                sr.application_date,
                sr.doctor_name,
                sr.doctor_code,
                sr.hq,
                hqm.hq_name as hq_name,
                sr.stockist_name,
                sr.stockist_code,
                sr.approval_status,
                COALESCE(sr.repeated_request, 0) as repeated_request,
                sr.total_scheme_value,
                sr.requested_by,
                sr.docstatus,
                COUNT(sri.name) as product_count
            FROM `tabScheme Request` sr
            LEFT JOIN `tabScheme Request Item` sri ON sr.name = sri.parent
            LEFT JOIN `tabHQ Master` hqm ON sr.hq = hqm.name
            WHERE {where_sql}
            GROUP BY sr.name
            ORDER BY sr.application_date DESC, sr.creation DESC
            LIMIT 200
        """, params, as_dict=True)

        return {"success": True, "data": rows, "division": division}
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Get Scheme List Portal Error")
        return {"success": False, "message": str(e)}



@frappe.whitelist()
def get_stockist_details(stockist_code):

    if not stockist_code:
        frappe.throw(_("Stockist code is required"))

    # Find Stockist Master by stockist_code
    name = frappe.db.get_value("Stockist Master", {"stockist_code": stockist_code}, "name")
    if not name:
        frappe.throw(_("Stockist not found"))

    st = frappe.get_doc("Stockist Master", name)

    team = region = zone = None
    if st.hq:
        team, region, zone = frappe.db.get_value("HQ Master", st.hq, ["team", "region", "zone"]) or (None, None, None)

    return {
        "stockist_code": st.stockist_code,
        "stockist_name": st.stockist_name,
        "division": getattr(st, "division", None),
        "hq": st.hq,
        "team": team,
        "region": region,
        "zone": zone,
    }


# ============================================================================
# INSIGHTS / ANALYTICS API
# ============================================================================

@frappe.whitelist()
def get_insights_filter_options(division=None):
    """Return filter dropdown options for Insights page."""
    if not division:
        division = get_user_division()

    regions = frappe.db.sql(
        "SELECT DISTINCT name FROM `tabRegion Master` WHERE status='Active' AND division IN (%s, 'Both') ORDER BY name",
        (division,), as_list=1,
    )
    teams = frappe.db.sql(
        "SELECT DISTINCT name FROM `tabTeam Master` WHERE status='Active' AND division IN (%s, 'Both') ORDER BY name",
        (division,), as_list=1,
    )
    hqs = frappe.db.sql(
        "SELECT DISTINCT name FROM `tabHQ Master` WHERE division=%s AND status='Active' ORDER BY name",
        (division,), as_list=1
    )
    financial_years = frappe.db.sql(
        "SELECT DISTINCT financial_year FROM `tabHQ Yearly Target` WHERE division=%s ORDER BY financial_year DESC",
        (division,), as_list=1
    )

    return {
        "regions": [r[0] for r in regions],
        "teams": [t[0] for t in teams],
        "hqs": [h[0] for h in hqs],
        "financial_years": [f[0] for f in financial_years],
    }


def _build_insights_conditions(division, from_date=None, to_date=None, region=None, team=None, hq=None, date_field="creation"):
    """Helper to build WHERE conditions for insights queries."""
    conditions = ["1=1"]
    values = []

    if division:
        conditions.append("division = %s")
        values.append(division)
    if from_date:
        conditions.append(f"{date_field} >= %s")
        values.append(from_date)
    if to_date:
        conditions.append(f"{date_field} <= %s")
        values.append(to_date)
    if region:
        conditions.append("region = %s")
        values.append(region)
    if team:
        conditions.append("team = %s")
        values.append(team)
    if hq:
        conditions.append("hq = %s")
        values.append(hq)

    return " AND ".join(conditions), values


@frappe.whitelist()
def get_insights_scheme_data(division=None, from_date=None, to_date=None, region=None, team=None, hq=None):
    """Scheme analytics data — trends, approval funnel, top doctors, top HQs, region breakdown."""
    if not division:
        division = get_user_division()

    cond, vals = _build_insights_conditions(division, from_date, to_date, region, team, hq, "application_date")

    # Monthly scheme trend
    monthly_trend = frappe.db.sql(f"""
        SELECT DATE_FORMAT(application_date, '%%Y-%%m') as month,
               COUNT(*) as count,
               COALESCE(SUM(total_scheme_value), 0) as total_value
        FROM `tabScheme Request`
        WHERE {cond}
        GROUP BY month ORDER BY month
    """, vals, as_dict=1)

    # Approval status breakdown
    approval_breakdown = frappe.db.sql(f"""
        SELECT approval_status as status, COUNT(*) as count
        FROM `tabScheme Request`
        WHERE {cond}
        GROUP BY approval_status
    """, vals, as_dict=1)

    # Top 10 doctors by scheme value
    top_doctors = frappe.db.sql(f"""
        SELECT doctor_name, COALESCE(SUM(total_scheme_value), 0) as total_value, COUNT(*) as count
        FROM `tabScheme Request`
        WHERE {cond} AND doctor_name IS NOT NULL AND doctor_name != ''
        GROUP BY doctor_code ORDER BY total_value DESC LIMIT 10
    """, vals, as_dict=1)

    # Top 10 HQs by scheme value
    top_hqs = frappe.db.sql(f"""
        SELECT hq, COALESCE(SUM(total_scheme_value), 0) as total_value, COUNT(*) as count
        FROM `tabScheme Request`
        WHERE {cond} AND hq IS NOT NULL AND hq != ''
        GROUP BY hq ORDER BY total_value DESC LIMIT 10
    """, vals, as_dict=1)

    # Scheme value by region
    region_breakdown = frappe.db.sql(f"""
        SELECT region, COALESCE(SUM(total_scheme_value), 0) as total_value, COUNT(*) as count
        FROM `tabScheme Request`
        WHERE {cond} AND region IS NOT NULL AND region != ''
        GROUP BY region ORDER BY total_value DESC
    """, vals, as_dict=1)

    # KPI: totals
    kpi = frappe.db.sql(f"""
        SELECT COUNT(*) as total_schemes,
               SUM(CASE WHEN approval_status='Approved' THEN 1 ELSE 0 END) as approved,
               COALESCE(SUM(total_scheme_value), 0) as total_value
        FROM `tabScheme Request`
        WHERE {cond}
    """, vals, as_dict=1)[0]

    return {
        "monthly_trend": monthly_trend,
        "approval_breakdown": approval_breakdown,
        "top_doctors": top_doctors,
        "top_hqs": top_hqs,
        "region_breakdown": region_breakdown,
        "kpi": kpi,
    }


@frappe.whitelist()
def get_insights_statement_data(division=None, from_date=None, to_date=None, region=None, team=None, hq=None):
    """Stock statement analytics — submissions, coverage, value trends."""
    if not division:
        division = get_user_division()

    cond, vals = _build_insights_conditions(division, from_date, to_date, region, team, hq, "statement_month")

    # Monthly statement count
    monthly_statements = frappe.db.sql(f"""
        SELECT DATE_FORMAT(statement_month, '%%Y-%%m') as month, COUNT(*) as count
        FROM `tabStockist Statement`
        WHERE {cond}
        GROUP BY month ORDER BY month
    """, vals, as_dict=1)

    # Closing value trend
    closing_value_trend = frappe.db.sql(f"""
        SELECT DATE_FORMAT(statement_month, '%%Y-%%m') as month,
               COALESCE(SUM(total_closing_value), 0) as total_closing
        FROM `tabStockist Statement`
        WHERE {cond}
        GROUP BY month ORDER BY month
    """, vals, as_dict=1)

    # Coverage by HQ — unique stockists who submitted per HQ
    hq_coverage = frappe.db.sql(f"""
        SELECT hq, COUNT(DISTINCT stockist_code) as stockist_count
        FROM `tabStockist Statement`
        WHERE {cond} AND hq IS NOT NULL AND hq != ''
        GROUP BY hq ORDER BY stockist_count DESC LIMIT 15
    """, vals, as_dict=1)

    # Extraction status
    extraction_status = frappe.db.sql(f"""
        SELECT COALESCE(extracted_data_status, 'Pending') as status, COUNT(*) as count
        FROM `tabStockist Statement`
        WHERE {cond}
        GROUP BY extracted_data_status
    """, vals, as_dict=1)

    # KPI
    kpi = frappe.db.sql(f"""
        SELECT COUNT(*) as total_statements,
               COUNT(DISTINCT stockist_code) as unique_stockists,
               COALESCE(SUM(total_closing_value), 0) as total_closing_value
        FROM `tabStockist Statement`
        WHERE {cond}
    """, vals, as_dict=1)[0]

    return {
        "monthly_statements": monthly_statements,
        "closing_value_trend": closing_value_trend,
        "hq_coverage": hq_coverage,
        "extraction_status": extraction_status,
        "kpi": kpi,
    }


@frappe.whitelist()
def get_insights_deduction_data(division=None, from_date=None, to_date=None, region=None, team=None, hq=None):
    """Scheme deduction analytics — monthly value, top stockists, status breakdown."""
    if not division:
        division = get_user_division()

    # Deduction table doesn't have region/team/hq directly — join via scheme_request
    base_cond = ["1=1"]
    base_vals = []
    if division:
        base_cond.append("sd.division = %s")
        base_vals.append(division)
    if from_date:
        base_cond.append("sd.deduction_date >= %s")
        base_vals.append(from_date)
    if to_date:
        base_cond.append("sd.deduction_date <= %s")
        base_vals.append(to_date)
    if region:
        base_cond.append("sr.region = %s")
        base_vals.append(region)
    if team:
        base_cond.append("sr.team = %s")
        base_vals.append(team)
    if hq:
        base_cond.append("sr.hq = %s")
        base_vals.append(hq)

    cond_str = " AND ".join(base_cond)

    # Monthly deduction value
    monthly_deductions = frappe.db.sql(f"""
        SELECT DATE_FORMAT(sd.deduction_date, '%%Y-%%m') as month,
               COUNT(*) as count,
               COALESCE(SUM(sd.total_deducted_value), 0) as total_value
        FROM `tabScheme Deduction` sd
        LEFT JOIN `tabScheme Request` sr ON sr.name = sd.scheme_request
        WHERE {cond_str}
        GROUP BY month ORDER BY month
    """, base_vals, as_dict=1)

    # Top stockists by deduction value
    top_stockists = frappe.db.sql(f"""
        SELECT sm.stockist_name, sd.stockist_code,
               COALESCE(SUM(sd.total_deducted_value), 0) as total_value,
               COUNT(*) as count
        FROM `tabScheme Deduction` sd
        LEFT JOIN `tabScheme Request` sr ON sr.name = sd.scheme_request
        LEFT JOIN `tabStockist Master` sm ON sm.name = sd.stockist_code
        WHERE {cond_str} AND sd.stockist_code IS NOT NULL
        GROUP BY sd.stockist_code ORDER BY total_value DESC LIMIT 10
    """, base_vals, as_dict=1)

    # Status breakdown
    status_breakdown = frappe.db.sql(f"""
        SELECT sd.status, COUNT(*) as count
        FROM `tabScheme Deduction` sd
        LEFT JOIN `tabScheme Request` sr ON sr.name = sd.scheme_request
        WHERE {cond_str}
        GROUP BY sd.status
    """, base_vals, as_dict=1)

    # KPI
    kpi = frappe.db.sql(f"""
        SELECT COUNT(*) as total_deductions,
               COALESCE(SUM(sd.total_deducted_value), 0) as total_value
        FROM `tabScheme Deduction` sd
        LEFT JOIN `tabScheme Request` sr ON sr.name = sd.scheme_request
        WHERE {cond_str}
    """, base_vals, as_dict=1)[0]

    return {
        "monthly_deductions": monthly_deductions,
        "top_stockists": top_stockists,
        "status_breakdown": status_breakdown,
        "kpi": kpi,
    }


@frappe.whitelist()
def get_insights_masters_data(division=None):
    """Masters overview — stockist count by HQ, HQ distribution by region."""
    if not division:
        division = get_user_division()

    # Stockist count by HQ (top 15)
    stockists_by_hq = frappe.db.sql("""
        SELECT hq, COUNT(*) as count
        FROM `tabStockist Master`
        WHERE division = %s AND status = 'Active'
        AND hq IS NOT NULL AND hq != ''
        GROUP BY hq ORDER BY count DESC LIMIT 15
    """, (division,), as_dict=1)

    # HQ distribution by region
    hqs_by_region = frappe.db.sql("""
        SELECT region, COUNT(*) as count
        FROM `tabHQ Master`
        WHERE division = %s AND status = 'Active'
        AND region IS NOT NULL AND region != ''
        GROUP BY region ORDER BY count DESC
    """, (division,), as_dict=1)

    # Doctors by HQ (top 15)
    doctors_by_hq = frappe.db.sql("""
        SELECT hq, COUNT(*) as count
        FROM `tabDoctor Master`
        WHERE division = %s AND status = 'Active'
        AND hq IS NOT NULL AND hq != ''
        GROUP BY hq ORDER BY count DESC LIMIT 15
    """, (division,), as_dict=1)

    # KPIs
    kpi = {
        "active_stockists": frappe.db.count("Stockist Master", {"division": division, "status": "Active"}),
        "active_hqs": frappe.db.count("HQ Master", {"division": division, "status": "Active"}),
        "active_doctors": frappe.db.count("Doctor Master", {"division": division, "status": "Active"}),
        "active_products": frappe.db.count("Product Master", {"division": division, "status": "Active"}),
    }

    return {
        "stockists_by_hq": stockists_by_hq,
        "hqs_by_region": hqs_by_region,
        "doctors_by_hq": doctors_by_hq,
        "kpi": kpi,
    }


@frappe.whitelist()
def get_insights_targets_data(division=None, financial_year=None):
    """Sales targets vs actuals — monthly comparison, region-wise breakdown."""
    if not division:
        division = get_user_division()

    if not financial_year:
        fy = frappe.db.sql(
            "SELECT financial_year FROM `tabHQ Yearly Target` WHERE division=%s ORDER BY financial_year DESC LIMIT 1",
            (division,), as_list=1
        )
        financial_year = fy[0][0] if fy else None

    if not financial_year:
        return {"monthly": [], "region_wise": [], "kpi": {"total_target": 0, "total_actual": 0}, "financial_year": None}

    # Monthly targets
    months_map = [
        ("apr", "Apr"), ("may", "May"), ("jun", "Jun"),
        ("jul", "Jul"), ("aug", "Aug"), ("sep", "Sep"),
        ("oct", "Oct"), ("nov", "Nov"), ("dec", "Dec"),
        ("jan", "Jan"), ("feb", "Feb"), ("mar", "Mar"),
    ]

    target_sums = frappe.db.sql("""
        SELECT
            COALESCE(SUM(ti.apr),0) as apr, COALESCE(SUM(ti.may),0) as may, COALESCE(SUM(ti.jun),0) as jun,
            COALESCE(SUM(ti.jul),0) as jul, COALESCE(SUM(ti.aug),0) as aug, COALESCE(SUM(ti.sep),0) as sep,
            COALESCE(SUM(ti.oct),0) as oct, COALESCE(SUM(ti.nov),0) as nov, COALESCE(SUM(ti.dec),0) as dec,
            COALESCE(SUM(ti.jan),0) as jan, COALESCE(SUM(ti.feb),0) as feb, COALESCE(SUM(ti.mar),0) as mar,
            COALESCE(SUM(ti.yearly_total),0) as yearly_total
        FROM `tabHQ Yearly Target` yt
        INNER JOIN `tabHQ Target Item` ti ON ti.parent = yt.name
        WHERE yt.docstatus = 1 AND yt.division = %s AND yt.financial_year = %s
    """, (division, financial_year), as_dict=1)[0]

    # Build monthly actuals from scheme requests (approved value by month)
    # Financial year format is like "2025-26" — parse start year
    try:
        start_year = int(financial_year.split("-")[0])
    except Exception:
        start_year = 2025

    month_to_date = {
        "apr": f"{start_year}-04", "may": f"{start_year}-05", "jun": f"{start_year}-06",
        "jul": f"{start_year}-07", "aug": f"{start_year}-08", "sep": f"{start_year}-09",
        "oct": f"{start_year}-10", "nov": f"{start_year}-11", "dec": f"{start_year}-12",
        "jan": f"{start_year+1}-01", "feb": f"{start_year+1}-02", "mar": f"{start_year+1}-03",
    }

    actual_data = frappe.db.sql("""
        SELECT DATE_FORMAT(application_date, '%%Y-%%m') as month,
               COALESCE(SUM(total_scheme_value), 0) as actual_value
        FROM `tabScheme Request`
        WHERE division = %s AND approval_status = 'Approved'
        AND application_date >= %s AND application_date <= %s
        GROUP BY month
    """, (division, f"{start_year}-04-01", f"{start_year+1}-03-31"), as_dict=1)

    actual_map = {a["month"]: a["actual_value"] for a in actual_data}

    monthly = []
    for key, label in months_map:
        monthly.append({
            "month": label,
            "target": flt(getattr(target_sums, key, 0) if hasattr(target_sums, key) else target_sums.get(key, 0)),
            "actual": flt(actual_map.get(month_to_date.get(key, ""), 0)),
        })

    # Region-wise target vs actual
    region_targets = frappe.db.sql("""
        SELECT yt.region,
               COALESCE(SUM(ti.yearly_total), 0) as target_value
        FROM `tabHQ Yearly Target` yt
        INNER JOIN `tabHQ Target Item` ti ON ti.parent = yt.name
        WHERE yt.docstatus = 1 AND yt.division = %s AND yt.financial_year = %s
        GROUP BY yt.region ORDER BY target_value DESC
    """, (division, financial_year), as_dict=1)

    region_actuals = frappe.db.sql("""
        SELECT region, COALESCE(SUM(total_scheme_value), 0) as actual_value
        FROM `tabScheme Request`
        WHERE division = %s AND approval_status = 'Approved'
        AND application_date >= %s AND application_date <= %s
        AND region IS NOT NULL AND region != ''
        GROUP BY region
    """, (division, f"{start_year}-04-01", f"{start_year+1}-03-31"), as_dict=1)

    actual_region_map = {a["region"]: a["actual_value"] for a in region_actuals}

    region_wise = []
    for rt in region_targets:
        region_wise.append({
            "region": rt["region"],
            "target": flt(rt["target_value"]),
            "actual": flt(actual_region_map.get(rt["region"], 0)),
        })

    total_target = flt(target_sums.get("yearly_total", 0))
    total_actual = sum(flt(actual_map.get(month_to_date.get(k, ""), 0)) for k, _ in months_map)

    return {
        "monthly": monthly,
        "region_wise": region_wise,
        "kpi": {"total_target": total_target, "total_actual": total_actual},
        "financial_year": financial_year,
    }


@frappe.whitelist()
def get_insights_products_data(division=None, from_date=None, to_date=None, region=None, team=None, hq=None):
    """Product movement insights — top products by closing value and by scheme value."""
    if not division:
        division = get_user_division()

    # Top products by closing value from statements
    stmt_cond = ["1=1"]
    stmt_vals = []
    if division:
        stmt_cond.append("ss.division = %s")
        stmt_vals.append(division)
    if from_date:
        stmt_cond.append("ss.statement_month >= %s")
        stmt_vals.append(from_date)
    if to_date:
        stmt_cond.append("ss.statement_month <= %s")
        stmt_vals.append(to_date)
    if region:
        stmt_cond.append("ss.region = %s")
        stmt_vals.append(region)
    if team:
        stmt_cond.append("ss.team = %s")
        stmt_vals.append(team)
    if hq:
        stmt_cond.append("ss.hq = %s")
        stmt_vals.append(hq)

    stmt_where = " AND ".join(stmt_cond)

    top_products_closing = frappe.db.sql(f"""
        SELECT si.product_name, si.product_code,
               COALESCE(SUM(si.closing_value), 0) as total_closing,
               COALESCE(SUM(si.sales_qty), 0) as total_sales_qty
        FROM `tabStockist Statement Item` si
        INNER JOIN `tabStockist Statement` ss ON ss.name = si.parent
        WHERE {stmt_where} AND si.product_name IS NOT NULL AND si.product_name != ''
        GROUP BY si.product_code ORDER BY total_closing DESC LIMIT 15
    """, stmt_vals, as_dict=1)

    # Top products by scheme value
    sch_cond = ["1=1"]
    sch_vals = []
    if division:
        sch_cond.append("sr.division = %s")
        sch_vals.append(division)
    if from_date:
        sch_cond.append("sr.application_date >= %s")
        sch_vals.append(from_date)
    if to_date:
        sch_cond.append("sr.application_date <= %s")
        sch_vals.append(to_date)
    if region:
        sch_cond.append("sr.region = %s")
        sch_vals.append(region)
    if team:
        sch_cond.append("sr.team = %s")
        sch_vals.append(team)
    if hq:
        sch_cond.append("sr.hq = %s")
        sch_vals.append(hq)

    sch_where = " AND ".join(sch_cond)

    top_products_scheme = frappe.db.sql(f"""
        SELECT sri.product_name, sri.product_code,
               COALESCE(SUM(sri.product_value), 0) as total_value,
               COALESCE(SUM(sri.quantity), 0) as total_qty
        FROM `tabScheme Request Item` sri
        INNER JOIN `tabScheme Request` sr ON sr.name = sri.parent
        WHERE {sch_where} AND sri.product_name IS NOT NULL AND sri.product_name != ''
        GROUP BY sri.product_code ORDER BY total_value DESC LIMIT 15
    """, sch_vals, as_dict=1)

    return {
        "top_products_closing": top_products_closing,
        "top_products_scheme": top_products_scheme,
    }


# ============================================================
#  PORTAL – USER MANAGEMENT & PROFILE APIs
# ============================================================

@frappe.whitelist()
def get_portal_users():
    """Return list of all non-Guest users for the Users management page (System Manager only)."""


    users = frappe.get_all(
        "User",
        filters={"name": ["not in", ["Guest", "Administrator"]], "user_type": "System User"},
        fields=["name", "email", "first_name", "middle_name", "last_name", "full_name",
                "mobile_no", "user_image", "enabled"],
        order_by="full_name asc",
        limit_page_length=500,
    )

    for u in users:
        roles = frappe.get_roles(u["name"])
        if "System Manager" in roles:
            u["role"] = "System Manager"
        elif "Sales Manager" in roles:
            u["role"] = "Sales Manager"
        else:
            u["role"] = "User"
        u["division"] = frappe.db.get_value("User", u["name"], "division") or ""

    return users


@frappe.whitelist()
def create_portal_user(email, first_name, last_name=None, middle_name=None,
                       role="User", division="Prima", mobile_no=None, password=None):
    """Create a new portal user. Restricted to System Manager."""
    if "System Manager" not in frappe.get_roles(frappe.session.user):
        frappe.throw("Not permitted", frappe.PermissionError)

    if not email or not first_name or not password:
        frappe.throw("Email, First Name and Password are required")

    if frappe.db.exists("User", email):
        frappe.throw(f"User {email} already exists")

    user = frappe.get_doc({
        "doctype": "User",
        "email": email,
        "first_name": first_name,
        "middle_name": middle_name or "",
        "last_name": last_name or "",
        "mobile_no": mobile_no or "",
        "division": division,
        "enabled": 1,
        "send_welcome_email": 0,
        "user_type": "System User",
    })
    user.insert(ignore_permissions=True)

    # Set password
    from frappe.utils.password import update_password as _update_password
    _update_password(email, password)

    # Assign role
    role_map = {
        "System Manager": "System Manager",
        "Sales Manager": "Sales Manager",
        "User": "Scanify User",
    }
    frappe_role = role_map.get(role, "Scanify User")
    if not frappe.db.exists("Role", frappe_role):
        frappe_role = role  # fall back to exact string

    user.add_roles(frappe_role)
    frappe.db.commit()

    return {"success": True, "user": email, "message": f"User {email} created successfully"}


@frappe.whitelist()
def get_my_profile():
    """Return the current user's profile data."""
    user = frappe.session.user
    if user == "Guest":
        frappe.throw("Please login to continue", frappe.PermissionError)

    doc = frappe.get_doc("User", user)
    roles = frappe.get_roles(user)
    if "System Manager" in roles:
        role = "System Manager"
    elif "Sales Manager" in roles:
        role = "Sales Manager"
    else:
        role = "User"

    return {
        "email": doc.email,
        "first_name": doc.first_name or "",
        "middle_name": doc.middle_name or "",
        "last_name": doc.last_name or "",
        "full_name": doc.full_name or "",
        "mobile_no": doc.mobile_no or "",
        "user_image": doc.user_image or "",
        "role": role,
        "division": doc.division or "",
    }


@frappe.whitelist()
def update_my_profile(first_name, middle_name=None, last_name=None, mobile_no=None):
    """Update the current user's own profile fields."""
    user = frappe.session.user
    if user == "Guest":
        frappe.throw("Please login to continue", frappe.PermissionError)

    doc = frappe.get_doc("User", user)
    doc.first_name = first_name
    doc.middle_name = middle_name or ""
    doc.last_name = last_name or ""
    doc.mobile_no = mobile_no or ""
    doc.save(ignore_permissions=True)
    frappe.db.commit()
    return {"success": True, "message": "Profile updated successfully"}


@frappe.whitelist()
def update_user_image(file_url):
    """Set the current user's profile image after upload."""
    user = frappe.session.user
    if user == "Guest":
        frappe.throw("Please login to continue", frappe.PermissionError)

    if not file_url:
        frappe.throw("No file URL provided")

    doc = frappe.get_doc("User", user)
    doc.user_image = file_url
    doc.save(ignore_permissions=True)
    frappe.db.commit()
    return {"success": True, "user_image": file_url, "message": "Profile image updated"}


# ───────────────────────────────────────────────────────────────
# Audit Trail / Change History — Portal API
# ───────────────────────────────────────────────────────────────

# Category → internal doctype(s) mapping
_AUDIT_CATEGORY_MAP = {
    "Masters": [
        "HQ Master", "Stockist Master", "Product Master",
        "Doctor Master", "Team Master", "Region Master",
        "Zone Master", "State Master",
    ],
    "Stock Statements": ["Stockist Statement"],
    "Scheme Requests": ["Scheme Request"],
    "Bulk Statement Upload": ["Bulk Statement Upload"],
    "Scheme Deductions": ["Scheme Deduction"],
    "Sales Targets": ["HQ Yearly Target"],
}

# Masters sub-type label → internal doctype
_MASTER_SUBTYPE_MAP = {
    "HQ": "HQ Master",
    "Stockist": "Stockist Master",
    "Product": "Product Master",
    "Doctor": "Doctor Master",
    "Team": "Team Master",
    "Region": "Region Master",
    "Zone": "Zone Master",
    "State": "State Master",
}

# Reverse: doctype → user-facing label
_DOCTYPE_LABEL_MAP = {
    "HQ Master": "HQ Master",
    "Stockist Master": "Stockist Master",
    "Product Master": "Product Master",
    "Doctor Master": "Doctor Master",
    "Team Master": "Team Master",
    "Region Master": "Region Master",
    "Zone Master": "Zone Master",
    "State Master": "State Master",
    "Stockist Statement": "Stock Statement",
    "Scheme Request": "Scheme Request",
    "Bulk Statement Upload": "Bulk Upload",
    "Scheme Deduction": "Scheme Deduction",
    "HQ Yearly Target": "Sales Target",
}

# Fields to hide from diffs (system / internal)
_SYSTEM_FIELDS = {
    "modified", "modified_by", "creation", "owner", "docstatus",
    "idx", "doctype", "name", "parent", "parenttype", "parentfield",
    "_liked_by", "_comments", "_assign", "_user_tags",
    "_seen", "amended_from",
}


def _get_user_display(email):
    """Return a display-friendly name for an email address."""
    if not email:
        return "System"
    full = frappe.db.get_value("User", email, "full_name")
    if full and full.strip():
        return full.strip()
    return email.split("@")[0].title()


def _parse_version_data(data_str):
    """Parse a Version record's data JSON, return (changed, added, removed)."""
    try:
        data = json.loads(data_str) if isinstance(data_str, str) else data_str
    except Exception:
        return [], [], []
    changed = data.get("changed", [])
    added = data.get("added", [])
    removed = data.get("removed", [])
    return changed, added, removed


def _field_label(doctype, fieldname):
    """Map a field name to its human-readable label using DocType meta."""
    try:
        meta = frappe.get_meta(doctype)
        df = meta.get_field(fieldname)
        if df and df.label:
            return df.label
    except Exception:
        pass
    return fieldname.replace("_", " ").title()


@frappe.whitelist()
def get_audit_trail_portal(
    category=None, sub_type=None, record_name=None,
    changed_by=None, from_date=None, to_date=None,
    page=1, page_size=25
):
    """Return paginated change-history entries for the portal."""
    try:
        roles = frappe.get_roles(frappe.session.user)
        page = max(int(page), 1)
        page_size = min(max(int(page_size), 5), 100)
        offset = (page - 1) * page_size

        # Determine which doctypes to query
        if category and category in _AUDIT_CATEGORY_MAP:
            doctypes = list(_AUDIT_CATEGORY_MAP[category])
            if category == "Masters" and sub_type and sub_type in _MASTER_SUBTYPE_MAP:
                doctypes = [_MASTER_SUBTYPE_MAP[sub_type]]
        else:
            # All tracked doctypes
            doctypes = []
            for dts in _AUDIT_CATEGORY_MAP.values():
                doctypes.extend(dts)

        if not doctypes:
            return {"success": True, "data": [], "total": 0, "page": page, "page_size": page_size}

        # Build SQL conditions
        conditions = []
        params = {}

        dt_keys = ["dt%d" % i for i in range(len(doctypes))]
        placeholders = ", ".join(["%%(%s)s" % k for k in dt_keys])
        for i, dt in enumerate(doctypes):
            params["dt%d" % i] = dt
        conditions.append("v.ref_doctype IN (%s)" % placeholders)

        if record_name:
            conditions.append("v.docname LIKE %(record_name)s")
            params["record_name"] = f"%{record_name}%"

        if changed_by:
            conditions.append("v.owner LIKE %(changed_by)s")
            params["changed_by"] = f"%{changed_by}%"

        if from_date:
            conditions.append("v.creation >= %(from_date)s")
            params["from_date"] = from_date

        if to_date:
            conditions.append("v.creation <= %(to_date)s")
            params["to_date"] = to_date + " 23:59:59"

        # Division filter for transactional docs
        division = get_user_division()
        div_doctypes_with_field = {
            "Stockist Statement", "Scheme Request",
            "Scheme Deduction", "HQ Yearly Target", "Bulk Statement Upload",
        }
        # Masters also have a division field on most records
        master_doctypes = set(_AUDIT_CATEGORY_MAP["Masters"])
        div_all = div_doctypes_with_field | master_doctypes
        # We'll filter via a LEFT JOIN approach — only if all requested doctypes support division
        # For simplicity, we add a sub-select condition when the category is specific
        if division and category and category != "Masters":
            # For transactional categories, filter by division on the parent record
            if len(doctypes) == 1 and doctypes[0] in div_doctypes_with_field:
                dt = doctypes[0]
                tab = f"tab{dt}"
                conditions.append(
                    f"EXISTS (SELECT 1 FROM `{tab}` dd WHERE dd.name = v.docname "
                    f"AND dd.division IN (%(div)s, 'Both'))"
                )
                params["div"] = division

        where = " AND ".join(conditions) if conditions else "1=1"

        # Count query
        count_sql = f"SELECT COUNT(*) as cnt FROM `tabVersion` v WHERE {where}"
        total = frappe.db.sql(count_sql, params, as_dict=True)[0].cnt

        # Data query
        data_sql = f"""
            SELECT v.name, v.ref_doctype, v.docname, v.owner, v.creation, v.data
            FROM `tabVersion` v
            WHERE {where}
            ORDER BY v.creation DESC
            LIMIT %(limit)s OFFSET %(offset)s
        """
        params["limit"] = page_size
        params["offset"] = offset

        rows = frappe.db.sql(data_sql, params, as_dict=True)

        result = []
        for row in rows:
            changed, added, removed = _parse_version_data(row.data)
            # Filter out system fields
            changed = [c for c in changed if c[0] not in _SYSTEM_FIELDS]
            change_count = len(changed) + len(added) + len(removed)

            result.append({
                "id": row.name,
                "record": row.docname,
                "category": _DOCTYPE_LABEL_MAP.get(row.ref_doctype, row.ref_doctype),
                "changed_by": _get_user_display(row.owner),
                "timestamp": str(row.creation),
                "change_count": change_count,
                "has_diff": change_count > 0,
            })

        return {
            "success": True,
            "data": result,
            "total": total,
            "page": page,
            "page_size": page_size,
        }
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Audit Trail Portal Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def get_audit_trail_detail(version_name):
    """Return field-level diff for a single Version record."""
    try:
        roles = frappe.get_roles(frappe.session.user)
        if "System Manager" not in roles and "Sales Manager" not in roles:
            return {"success": False, "message": "Insufficient permissions"}

        ver = frappe.db.get_value(
            "Version", version_name,
            ["ref_doctype", "docname", "data", "owner", "creation"],
            as_dict=True,
        )
        if not ver:
            return {"success": False, "message": "Record not found"}

        changed, added, removed = _parse_version_data(ver.data)
        ref_dt = ver.ref_doctype

        changes = []
        for c in changed:
            fname = c[0]
            if fname in _SYSTEM_FIELDS:
                continue
            changes.append({
                "label": _field_label(ref_dt, fname),
                "old_value": str(c[1]) if c[1] is not None else "",
                "new_value": str(c[2]) if c[2] is not None else "",
                "type": "changed",
            })

        for a in added:
            # added is a list of [row_doctype, {field: val, ...}]
            if isinstance(a, list) and len(a) >= 2:
                row_dt = a[0]
                row_data = a[1] if isinstance(a[1], dict) else {}
                summary = ", ".join(
                    f"{k}: {v}" for k, v in row_data.items()
                    if k not in _SYSTEM_FIELDS and v
                )
                changes.append({
                    "label": f"Row added",
                    "old_value": "",
                    "new_value": summary[:200] if summary else "(new row)",
                    "type": "added",
                })

        for r in removed:
            if isinstance(r, list) and len(r) >= 2:
                row_dt = r[0]
                row_data = r[1] if isinstance(r[1], dict) else {}
                summary = ", ".join(
                    f"{k}: {v}" for k, v in row_data.items()
                    if k not in _SYSTEM_FIELDS and v
                )
                changes.append({
                    "label": f"Row removed",
                    "old_value": summary[:200] if summary else "(removed row)",
                    "new_value": "",
                    "type": "removed",
                })

        return {
            "success": True,
            "record": ver.docname,
            "category": _DOCTYPE_LABEL_MAP.get(ref_dt, ref_dt),
            "changed_by": _get_user_display(ver.owner),
            "timestamp": str(ver.creation),
            "changes": changes,
        }
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Audit Trail Detail Error")
        return {"success": False, "message": str(e)}


# =============================================================================
# EXPORT MASTERS
# =============================================================================

# Master export configuration — maps frontend keys to doctypes and column definitions
_EXPORT_MASTER_CONFIGS = {
    "hq": {
        "title": "HQ Master",
        "doctype": "HQ Master",
        "columns": ["hq_code", "hq_name", "team", "region", "zone", "per_capita", "division", "status"],
        "headers": ["HQ Code", "HQ Name", "Team", "Region", "Zone", "Per Capita", "Division", "Status"],
        "resolve": {"team": "Team Master:team_name", "region": "Region Master:region_name", "zone": "Zone Master:zone_name"},
    },
    "stockist": {
        "title": "Stockist Master",
        "doctype": "Stockist Master",
        "columns": ["stockist_code", "stockist_name", "hq", "team", "region", "zone", "address", "contact_person", "phone", "email", "division", "status"],
        "headers": ["Stockist Code", "Stockist Name", "HQ", "Team", "Region", "Zone", "Address", "Contact Person", "Phone", "Email", "Division", "Status"],
        "resolve": {"hq": "HQ Master:hq_name", "team": "Team Master:team_name", "region": "Region Master:region_name", "zone": "Zone Master:zone_name"},
    },
    "product": {
        "title": "Product Master",
        "doctype": "Product Master",
        "columns": ["product_code", "product_name", "product_group", "category", "pack", "pack_conversion", "pts", "ptr", "mrp", "gst_rate", "division", "status"],
        "headers": ["Product Code", "Product Name", "Product Group", "Category", "Pack", "Pack Conversion", "PTS", "PTR", "MRP", "GST Rate", "Division", "Status"],
        "resolve": {},
    },
    "doctor": {
        "title": "Doctor Master",
        "doctype": "Doctor Master",
        "columns": ["doctor_code", "doctor_name", "qualification", "doctor_category", "specialization", "phone", "place", "hq", "team", "region", "state", "zone", "chemist_name", "division", "status"],
        "headers": ["Doctor Code", "Doctor Name", "Qualification", "Category", "Specialization", "Phone", "Place", "HQ", "Team", "Region", "State", "Zone", "Chemist Name", "Division", "Status"],
        "resolve": {"hq": "HQ Master:hq_name", "team": "Team Master:team_name", "region": "Region Master:region_name", "state": "State Master:state_name", "zone": "Zone Master:zone_name"},
    },
    "team": {
        "title": "Team Master",
        "doctype": "Team Master",
        "columns": ["team_code", "team_name", "region", "sanctioned_strength", "division", "status"],
        "headers": ["Team Code", "Team Name", "Region", "Sanctioned Strength", "Division", "Status"],
        "resolve": {"region": "Region Master:region_name"},
    },
    "region": {
        "title": "Region Master",
        "doctype": "Region Master",
        "columns": ["region_code", "region_name", "zone", "state", "division", "status"],
        "headers": ["Region Code", "Region Name", "Zone", "State", "Division", "Status"],
        "resolve": {"zone": "Zone Master:zone_name", "state": "State Master:state_name"},
    },
    "zone": {
        "title": "Zone Master",
        "doctype": "Zone Master",
        "columns": ["zone_code", "zone_name", "division", "status"],
        "headers": ["Zone Code", "Zone Name", "Division", "Status"],
        "resolve": {},
    },
    "state": {
        "title": "State Master",
        "doctype": "State Master",
        "columns": ["state_code", "state_name", "division", "status"],
        "headers": ["State Code", "State Name", "Division", "Status"],
        "resolve": {},
    },
}


def _fetch_export_data(config, division):
    """Fetch and resolve data for a master export."""
    doctype = config["doctype"]
    filters = {}
    if division:
        filters["division"] = ["in", [division, "Both"]]

    data = frappe.get_all(
        doctype,
        filters=filters,
        fields=["name"] + config["columns"],
        order_by="name asc",
        limit_page_length=0,
    )

    # Resolve Link fields to human-readable labels
    resolve_map = config.get("resolve", {})
    # Build caches for each linked doctype to avoid N+1 queries
    resolve_caches = {}
    for field, spec in resolve_map.items():
        linked_dt, label_field = spec.split(":")
        all_vals = frappe.get_all(linked_dt, fields=["name", label_field], limit_page_length=0)
        resolve_caches[field] = {v["name"]: v[label_field] for v in all_vals}

    for row in data:
        for field, spec in resolve_map.items():
            val = row.get(field)
            if val and field in resolve_caches:
                row[field] = resolve_caches[field].get(val, val)

    return data


def _generate_excel(config, data, division):
    """Generate a professional Excel file for a master."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = config["title"]

    # Company header
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(config["headers"]))
    ws["A1"] = "Stedman Pharmaceuticals Pvt Ltd"
    ws["A1"].font = Font(bold=True, size=14, color="1e293b")
    ws["A1"].alignment = Alignment(horizontal="center")

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(config["headers"]))
    ws["A2"] = config["title"]
    ws["A2"].font = Font(bold=True, size=12, color="4f46e5")
    ws["A2"].alignment = Alignment(horizontal="center")

    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(config["headers"]))
    ws["A3"] = f"Division: {division or 'All'}  |  Exported: {frappe.utils.now_datetime().strftime('%d %b %Y, %I:%M %p')}"
    ws["A3"].font = Font(size=10, color="64748b")
    ws["A3"].alignment = Alignment(horizontal="center")

    # Column headers at row 5
    header_fill = PatternFill(start_color="1e293b", end_color="1e293b", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=10)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin", color="d1d5db"),
        right=Side(style="thin", color="d1d5db"),
        top=Side(style="thin", color="d1d5db"),
        bottom=Side(style="thin", color="d1d5db"),
    )

    for col_idx, header in enumerate(config["headers"], 1):
        cell = ws.cell(row=5, column=col_idx)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    # Data rows starting at row 6
    alt_fill = PatternFill(start_color="f8fafc", end_color="f8fafc", fill_type="solid")
    data_align = Alignment(vertical="center", wrap_text=True)

    for row_idx, row in enumerate(data):
        excel_row = row_idx + 6
        for col_idx, col_key in enumerate(config["columns"], 1):
            cell = ws.cell(row=excel_row, column=col_idx)
            cell.value = row.get(col_key, "")
            cell.alignment = data_align
            cell.border = thin_border
            if row_idx % 2 == 1:
                cell.fill = alt_fill

    # Auto-fit column widths
    for col_idx, col_key in enumerate(config["columns"], 1):
        max_len = len(config["headers"][col_idx - 1])
        for row in data[:100]:  # Sample first 100 rows
            val = str(row.get(col_key, "") or "")
            if len(val) > max_len:
                max_len = len(val)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 40)

    # Footer
    footer_row = len(data) + 7
    ws.merge_cells(start_row=footer_row, start_column=1, end_row=footer_row, end_column=len(config["headers"]))
    ws.cell(row=footer_row, column=1).value = f"Total Records: {len(data)}"
    ws.cell(row=footer_row, column=1).font = Font(bold=True, size=10, color="64748b")

    return wb


def _generate_csv_content(config, data):
    """Generate CSV content for a master."""
    import csv
    import io

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(config["headers"])

    for row in data:
        writer.writerow([row.get(col, "") for col in config["columns"]])

    return output.getvalue()


def _generate_pdf_html(config, data, division):
    """Generate professional HTML for PDF conversion via wkhtmltopdf."""
    export_date = frappe.utils.now_datetime().strftime("%d %b %Y, %I:%M %p")
    num_cols = len(config["headers"])

    # Build table rows
    rows_html = ""
    for idx, row in enumerate(data):
        bg = "#f8fafc" if idx % 2 == 1 else "#ffffff"
        rows_html += f'<tr style="background:{bg};">'
        for col in config["columns"]:
            val = row.get(col, "") or ""
            rows_html += f'<td style="padding:6px 8px;border:1px solid #e2e8f0;font-size:9px;">{val}</td>'
        rows_html += "</tr>"

    # Build header cells
    header_cells = ""
    for h in config["headers"]:
        header_cells += f'<th style="padding:8px;border:1px solid #334155;background:#1e293b;color:#fff;font-size:9px;text-align:center;">{h}</th>'

    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            @page {{
                size: {"landscape" if num_cols > 6 else "portrait"};
                margin: 15mm;
            }}
            body {{
                font-family: 'Inter', 'Helvetica Neue', Arial, sans-serif;
                color: #1e293b;
                margin: 0;
            }}
            .header {{
                text-align: center;
                margin-bottom: 20px;
                border-bottom: 2px solid #4f46e5;
                padding-bottom: 12px;
            }}
            .header h1 {{
                font-size: 18px;
                margin: 0;
                color: #1e293b;
            }}
            .header h2 {{
                font-size: 14px;
                margin: 4px 0;
                color: #4f46e5;
                font-weight: 600;
            }}
            .header p {{
                font-size: 10px;
                color: #64748b;
                margin: 4px 0 0 0;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                margin-top: 10px;
            }}
            .footer {{
                text-align: center;
                margin-top: 15px;
                font-size: 9px;
                color: #94a3b8;
            }}
        </style>
    </head>
    <body>
        <div class="header">
            <h1>Stedman Pharmaceuticals Pvt Ltd</h1>
            <h2>{config["title"]}</h2>
            <p>Division: {division or "All"}  &bull;  Exported: {export_date}  &bull;  Total Records: {len(data)}</p>
        </div>
        <table>
            <thead><tr>{header_cells}</tr></thead>
            <tbody>{rows_html}</tbody>
        </table>
        <div class="footer">
            &copy; {frappe.utils.now_datetime().year} Stedman Pharmaceuticals Pvt Ltd &bull; Generated by Scanify
        </div>
    </body>
    </html>
    """
    return html


@frappe.whitelist()
def export_master_data(master_type, format_type="xlsx", division=None):
    """Export a single master to Excel, CSV, or PDF."""
    try:
        if master_type not in _EXPORT_MASTER_CONFIGS:
            return {"success": False, "message": f"Unknown master type: {master_type}"}

        config = _EXPORT_MASTER_CONFIGS[master_type]
        data = _fetch_export_data(config, division)

        timestamp = frappe.utils.now_datetime().strftime("%Y%m%d_%H%M%S")
        safe_title = config["title"].replace(" ", "_")

        files_dir = frappe.get_site_path("public", "files")
        os.makedirs(files_dir, exist_ok=True)

        if format_type == "xlsx":
            wb = _generate_excel(config, data, division)
            filename = f"{safe_title}_{timestamp}.xlsx"
            wb.save(os.path.join(files_dir, filename))

        elif format_type == "csv":
            csv_content = _generate_csv_content(config, data)
            filename = f"{safe_title}_{timestamp}.csv"
            with open(os.path.join(files_dir, filename), "w", encoding="utf-8-sig") as f:
                f.write(csv_content)

        elif format_type == "pdf":
            html = _generate_pdf_html(config, data, division)
            from frappe.utils.pdf import get_pdf
            pdf_content = get_pdf(html)
            filename = f"{safe_title}_{timestamp}.pdf"
            with open(os.path.join(files_dir, filename), "wb") as f:
                f.write(pdf_content)

        else:
            return {"success": False, "message": f"Unsupported format: {format_type}"}

        return {
            "success": True,
            "message": f"{config['title']} exported successfully",
            "file_url": f"/files/{filename}",
        }

    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Export Master Data Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def export_all_masters_zip(format_type="xlsx", division=None):
    """Export all masters into individual files and bundle them in a ZIP."""
    try:
        import zipfile
        import tempfile

        timestamp = frappe.utils.now_datetime().strftime("%Y%m%d_%H%M%S")
        zip_filename = f"All_Masters_{timestamp}.zip"
        files_dir = frappe.get_site_path("public", "files")
        os.makedirs(files_dir, exist_ok=True)
        zip_filepath = os.path.join(files_dir, zip_filename)

        with zipfile.ZipFile(zip_filepath, "w", zipfile.ZIP_DEFLATED) as zf:
            for master_key, config in _EXPORT_MASTER_CONFIGS.items():
                data = _fetch_export_data(config, division)
                safe_title = config["title"].replace(" ", "_")

                if format_type == "xlsx":
                    wb = _generate_excel(config, data, division)
                    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                        wb.save(tmp.name)
                        zf.write(tmp.name, f"{safe_title}.xlsx")
                        os.unlink(tmp.name)

                elif format_type == "csv":
                    csv_content = _generate_csv_content(config, data)
                    zf.writestr(f"{safe_title}.csv", csv_content)

                elif format_type == "pdf":
                    html = _generate_pdf_html(config, data, division)
                    from frappe.utils.pdf import get_pdf
                    pdf_content = get_pdf(html)
                    zf.writestr(f"{safe_title}.pdf", pdf_content)

        return {
            "success": True,
            "message": "All masters exported successfully",
            "file_url": f"/files/{zip_filename}",
        }

    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Export All Masters ZIP Error")
        return {"success": False, "message": str(e)}


# ===================== DELETE STATEMENT APIs =====================

@frappe.whitelist()
def search_stockist_statements(search_term="", division=None):
    """Search stockist statements by stockist name, code or statement name"""
    if not division:
        division = get_user_division()

    term = f"%{search_term}%"
    division_clause = ""
    params = {"term": term}

    if division and division != "Both":
        division_clause = "AND ss.division IN (%(division)s, 'Both')"
        params["division"] = division

    results = frappe.db.sql("""
        SELECT ss.name, ss.stockist_code, ss.stockist_name, ss.statement_month,
               ss.hq, ss.region, ss.docstatus, ss.division
        FROM `tabStockist Statement` ss
        WHERE (ss.stockist_name LIKE %(term)s
               OR ss.stockist_code LIKE %(term)s
               OR ss.name LIKE %(term)s)
        {division_clause}
        ORDER BY ss.modified DESC
        LIMIT 20
    """.format(division_clause=division_clause), params, as_dict=True)

    return results


@frappe.whitelist()
def get_statement_summary(doc_name):
    """Get full metadata summary for a stockist statement"""
    try:
        doc = frappe.get_doc("Stockist Statement", doc_name)
        return {
            "success": True,
            "name": doc.name,
            "stockist_code": doc.stockist_code,
            "stockist_name": doc.stockist_name,
            "statement_month": str(doc.statement_month) if doc.statement_month else None,
            "hq": doc.hq,
            "team": doc.team,
            "region": doc.region,
            "zone": doc.zone,
            "division": doc.division,
            "docstatus": doc.docstatus,
            "creation": str(doc.creation) if doc.creation else None,
            "uploaded_file": doc.uploaded_file,
            "total_items": len(doc.items) if doc.items else 0,
            "total_opening_value": flt(doc.total_opening_value),
            "total_purchase_value": flt(doc.total_purchase_value),
            "total_sales_value": flt(doc.total_sales_value_pts),
            "total_closing_value": flt(doc.total_closing_value),
        }
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Get Statement Summary Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def delete_stockist_statement(doc_name, reason):
    """Delete a stockist statement with a mandatory reason logged to audit trail"""
    try:
        reason = (reason or "").strip()
        if not reason or len(reason) < 5:
            return {"success": False, "message": "A reason of at least 5 characters is required."}

        if not frappe.db.exists("Stockist Statement", doc_name):
            return {"success": False, "message": f"Statement {doc_name} does not exist."}

        doc = frappe.get_doc("Stockist Statement", doc_name)

        # Store summary before deletion for audit trail
        summary = (
            f"Stockist: {doc.stockist_name} ({doc.stockist_code}), "
            f"Month: {doc.statement_month}, HQ: {doc.hq}, "
            f"Items: {len(doc.items) if doc.items else 0}"
        )

        # Add Comment for audit trail (persists after doc deletion)
        frappe.get_doc({
            "doctype": "Comment",
            "comment_type": "Info",
            "reference_doctype": "Stockist Statement",
            "reference_name": doc_name,
            "content": f"<b>Statement Deleted</b><br>Reason: {frappe.utils.escape_html(reason)}<br>{summary}",
            "comment_email": frappe.session.user,
        }).insert(ignore_permissions=True)

        # Cancel first if submitted
        if doc.docstatus == 1:
            doc.flags.ignore_permissions = True
            doc.cancel()

        # Delete the document
        frappe.delete_doc("Stockist Statement", doc_name, ignore_permissions=True, force=True)
        frappe.db.commit()

        return {"success": True, "message": f"Statement {doc_name} deleted. Reason logged to audit trail."}

    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Delete Stockist Statement Error")
        return {"success": False, "message": str(e)}


# ===================== SCHEME REQUEST: STOCKISTS BY REGION =====================

@frappe.whitelist()
def get_stockists_by_region(region, division=None):
    """Get all active stockists in a region (across all HQs in that region)"""
    if not division:
        division = get_user_division()

    filters = {"region": region, "status": "Active"}
    if division and division != "Both":
        filters["division"] = ["in", [division, "Both"]]

    stockists = frappe.get_all(
        "Stockist Master",
        filters=filters,
        fields=["name", "stockist_code", "stockist_name", "hq"],
        order_by="stockist_name asc"
    )

    # Fetch HQ names for display
    hq_names = {}
    for s in stockists:
        if s.hq and s.hq not in hq_names:
            hq_names[s.hq] = frappe.db.get_value("HQ Master", s.hq, "hq_name") or s.hq
        s["hq_name"] = hq_names.get(s.hq, s.hq or "")

    return stockists


# ===================== DUPLICATE STATEMENT CHECK =====================

@frappe.whitelist()
def check_statement_exists(stockist_code, statement_month):
    """Check if a stockist statement already exists for the given stockist + month"""
    if not stockist_code or not statement_month:
        return {"exists": False}

    # Normalise month to first-of-month
    if len(statement_month) == 7:
        statement_month = statement_month + "-01"

    existing = frappe.db.exists("Stockist Statement", {
        "stockist_code": stockist_code,
        "statement_month": statement_month
    })

    if existing:
        return {
            "exists": True,
            "statement_name": existing,
            "message": f"A statement already exists for this stockist in this month: {existing}"
        }

    return {"exists": False}


# ─────────────────────────────────────────────────────
# Primary Sales Upload, List & Export
# ─────────────────────────────────────────────────────

# Column mapping from Excel headers → doctype fields
_PRIMARY_SALES_COL_MAP = {
    "stockistcode": "stockist_code",
    "product_head": "product_head",
    "stockistname": "stockist_name",
    "citypool": "citypool",
    "team": "team",
    "region": "region",
    "act_region": "act_region",
    "zonee": "zonee",
    "invoiceno": "invoiceno",
    "invoicedate": "invoicedate",
    "pcode": "pcode",
    "product": "product",
    "pack": "pack",
    "batchno": "batchno",
    "quantity": "quantity",
    "freeqty": "freeqty",
    "expqty": "expqty",
    "pts": "pts",
    "ptsvalue": "ptsvalue",
    "ptr": "ptr",
    "ptrvalue": "ptrvalue",
    "mrp": "mrp",
    "mrpvalue": "mrpvalue",
    "nrv": "nrv",
    "nrvvalue": "nrvvalue",
    "dsort": "dsort",
    "direct_party": "direct_party",
    "iscancelled": "iscancelled",
}

_NUMERIC_FIELDS = {
    "quantity", "freeqty", "expqty",
    "pts", "ptsvalue", "ptr", "ptrvalue",
    "mrp", "mrpvalue", "nrv", "nrvvalue", "dsort",
}

_BOOL_FIELDS = {"direct_party", "iscancelled"}


def _parse_bool(val):
    """Convert various truthy representations to 1/0."""
    if val is None:
        return 0
    if isinstance(val, bool):
        return 1 if val else 0
    if isinstance(val, (int, float)):
        return 1 if val else 0
    s = str(val).strip().lower()
    return 1 if s in ("true", "1", "yes", "y") else 0


def _parse_num(val):
    """Safely parse a numeric value."""
    if val is None:
        return 0
    try:
        return flt(val)
    except Exception:
        return 0


@frappe.whitelist()
def process_primary_sales_upload(upload_month, file_url):
    """
    Process an uploaded Excel file of primary sales data.
    Validates columns, maps to masters, and creates Primary Sales Data records.
    """
    import openpyxl
    from datetime import datetime

    user_division = get_user_division() or "Prima"

    # Validate month format (YYYY-MM)
    if not upload_month or not re.match(r"^\d{4}-\d{2}$", str(upload_month)):
        return {"success": False, "error": "Invalid month format. Use YYYY-MM."}

    # Get the file path
    file_path = frappe.get_site_path(file_url.lstrip("/"))
    if not os.path.exists(file_path):
        # Try with 'private' prefix
        file_path = frappe.get_site_path("private", "files", os.path.basename(file_url))
    if not os.path.exists(file_path):
        return {"success": False, "error": "Uploaded file not found on server."}

    # Create upload record
    upload_doc = frappe.get_doc({
        "doctype": "Primary Sales Upload",
        "upload_month": upload_month,
        "division": user_division,
        "uploaded_by": frappe.session.user,
        "upload_date": frappe.utils.now(),
        "file": file_url,
        "status": "Processing",
    })
    upload_doc.insert(ignore_permissions=True)
    frappe.db.commit()

    errors = []
    success_count = 0
    total_rows = 0

    try:
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        ws = wb.active

        # Read headers from row 1
        raw_headers = [cell.value for cell in ws[1]]
        headers = []
        for h in raw_headers:
            if h is None:
                headers.append(None)
            else:
                headers.append(str(h).strip().lower())

        # Validate required columns exist
        required_cols = {"stockistcode", "iscancelled"}
        found_cols = set(h for h in headers if h)
        missing = required_cols - found_cols
        if missing:
            upload_doc.status = "Failed"
            upload_doc.error_log = f"Missing required columns: {', '.join(missing)}"
            upload_doc.save(ignore_permissions=True)
            frappe.db.commit()
            return {"success": False, "error": f"Missing required columns: {', '.join(missing)}"}

        # Build caches for validation
        stockist_cache = {}
        for s in frappe.get_all("Stockist Master",
                                filters={"division": ["in", [user_division, "Both"]]},
                                fields=["stockist_code", "name"],
                                limit_page_length=0):
            stockist_cache[s.stockist_code] = s.name

        product_cache = {}
        for p in frappe.get_all("Product Master",
                                filters={"division": ["in", [user_division, "Both", "ASPR", "Wellness"]]},
                                fields=["product_code", "name"],
                                limit_page_length=0):
            product_cache[p.product_code] = p.name

        # Process rows
        batch_size = 100
        batch = []

        for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            # Skip completely empty rows
            if all(v is None for v in row):
                continue

            total_rows += 1
            row_data = {}

            for col_idx, val in enumerate(row):
                if col_idx >= len(headers) or headers[col_idx] is None:
                    continue
                excel_col = headers[col_idx]
                if excel_col in _PRIMARY_SALES_COL_MAP:
                    field_name = _PRIMARY_SALES_COL_MAP[excel_col]
                    if field_name in _BOOL_FIELDS:
                        row_data[field_name] = _parse_bool(val)
                    elif field_name in _NUMERIC_FIELDS:
                        row_data[field_name] = _parse_num(val)
                    elif field_name == "invoicedate":
                        if isinstance(val, datetime):
                            row_data[field_name] = val.strftime("%Y-%m-%d")
                        elif val:
                            row_data[field_name] = str(val)
                        else:
                            row_data[field_name] = None
                    else:
                        row_data[field_name] = str(val).strip() if val else ""

            is_cancelled = row_data.get("iscancelled", 0)

            # Validate stockist_code (always required)
            stockist_code = row_data.get("stockist_code", "")
            if not stockist_code:
                errors.append(f"Row {row_idx}: Missing stockist code")
                continue

            # For non-cancelled rows, validate product code if present
            pcode = row_data.get("pcode", "")
            if not is_cancelled and pcode:
                # Warn if product not found but still store the data
                if pcode not in product_cache:
                    errors.append(f"Row {row_idx}: Product code '{pcode}' not found in Product Master (data saved anyway)")

            # Warn if stockist not found but still store the data
            if stockist_code not in stockist_cache:
                errors.append(f"Row {row_idx}: Stockist code '{stockist_code}' not found in Stockist Master (data saved anyway)")

            # Build the record
            row_data["upload_month"] = upload_month
            row_data["division"] = user_division
            row_data["upload_ref"] = upload_doc.name
            row_data["doctype"] = "Primary Sales Data"

            batch.append(row_data)
            success_count += 1

            # Insert in batches
            if len(batch) >= batch_size:
                for rec in batch:
                    doc = frappe.get_doc(rec)
                    doc.insert(ignore_permissions=True)
                batch = []
                frappe.db.commit()

        # Insert remaining
        if batch:
            for rec in batch:
                doc = frappe.get_doc(rec)
                doc.insert(ignore_permissions=True)
            frappe.db.commit()

        wb.close()

    except Exception as e:
        frappe.db.rollback()
        upload_doc.reload()
        upload_doc.status = "Failed"
        upload_doc.error_log = str(e)
        upload_doc.save(ignore_permissions=True)
        frappe.db.commit()
        frappe.log_error(f"Primary Sales Upload Error: {str(e)}", "Primary Sales Upload")
        return {"success": False, "error": str(e)}

    # Update upload record
    upload_doc.reload()
    upload_doc.total_rows = total_rows
    upload_doc.success_count = success_count
    upload_doc.error_count = len([e for e in errors if "not found" not in e])
    upload_doc.status = "Completed"
    if errors:
        upload_doc.error_log = "\n".join(errors[:200])
    upload_doc.save(ignore_permissions=True)
    frappe.db.commit()

    return {
        "success": True,
        "total_rows": total_rows,
        "success_count": success_count,
        "error_count": len(errors),
        "errors": errors[:20],
        "upload_name": upload_doc.name,
    }


@frappe.whitelist()
def get_primary_sales_data(division, page=1, page_size=50,
                           upload_month=None, zonee=None, region=None,
                           team=None, product_head=None, iscancelled=None,
                           stockist_search=None, pcode=None, product_search=None,
                           invoiceno=None):
    """Fetch paginated primary sales data with filters."""

    page = int(page)
    page_size = min(int(page_size), 200)
    offset = (page - 1) * page_size

    conditions = ["division = %(division)s"]
    params = {"division": division, "page_size": page_size, "offset": offset}

    if upload_month:
        conditions.append("upload_month = %(upload_month)s")
        params["upload_month"] = upload_month

    if zonee:
        conditions.append("zonee = %(zonee)s")
        params["zonee"] = zonee

    if region:
        conditions.append("region = %(region)s")
        params["region"] = region

    if team:
        conditions.append("team = %(team)s")
        params["team"] = team

    if product_head:
        conditions.append("product_head = %(product_head)s")
        params["product_head"] = product_head

    if iscancelled is not None and iscancelled != "":
        conditions.append("iscancelled = %(iscancelled)s")
        params["iscancelled"] = int(iscancelled)

    if stockist_search:
        conditions.append("(stockist_code LIKE %(stockist_search)s OR stockist_name LIKE %(stockist_search)s)")
        params["stockist_search"] = f"%{stockist_search}%"

    if pcode:
        conditions.append("pcode LIKE %(pcode)s")
        params["pcode"] = f"%{pcode}%"

    if product_search:
        conditions.append("product LIKE %(product_search)s")
        params["product_search"] = f"%{product_search}%"

    if invoiceno:
        conditions.append("invoiceno LIKE %(invoiceno)s")
        params["invoiceno"] = f"%{invoiceno}%"

    where_clause = " AND ".join(conditions)

    total = frappe.db.sql(
        f"SELECT COUNT(*) FROM `tabPrimary Sales Data` WHERE {where_clause}",
        params
    )[0][0]

    rows = frappe.db.sql(
        f"""SELECT
            name,
            stockist_code, stockist_name, product_head, pcode, product, pack,
            invoiceno, invoicedate, batchno,
            quantity, freeqty, expqty,
            pts, ptsvalue, ptr, ptrvalue,
            mrp, mrpvalue, nrv, nrvvalue,
            team, region, zonee, citypool,
            direct_party, iscancelled, upload_month, dsort
        FROM `tabPrimary Sales Data`
        WHERE {where_clause}
        ORDER BY stockist_code, invoicedate, pcode
        LIMIT %(page_size)s OFFSET %(offset)s""",
        params, as_dict=True
    )

    return {
        "rows": rows,
        "total": total,
        "page": page,
        "page_size": page_size,
    }


@frappe.whitelist()
def get_primary_sales_count(month, division):
    """Get count of primary sales records for a given month and division."""
    count = frappe.db.count("Primary Sales Data", filters={
        "upload_month": month,
        "division": division,
    })
    return count


@frappe.whitelist()
def export_primary_sales_data(month, division):
    """Export primary sales data as Excel for the given month and division."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    if not month or not division:
        frappe.throw("Month and division are required")

    rows = frappe.db.sql("""
        SELECT
            stockist_code, product_head, stockist_name, citypool,
            team, region, act_region, zonee,
            invoiceno, invoicedate, pcode, product, pack, batchno,
            quantity, freeqty, expqty,
            pts, ptsvalue, ptr, ptrvalue,
            mrp, mrpvalue, nrv, nrvvalue,
            dsort, direct_party, iscancelled
        FROM `tabPrimary Sales Data`
        WHERE upload_month = %(month)s AND division = %(division)s
        ORDER BY stockist_code, invoicedate, pcode
    """, {"month": month, "division": division}, as_dict=True)

    if not rows:
        frappe.throw(f"No data found for {month} in {division} division")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Primary Sales {month}"

    # Excel column headers (matching original upload format)
    excel_headers = [
        "stockistcode", "product_head", "stockistname", "citypool",
        "team", "region", "act_region", "zonee",
        "invoiceno", "invoicedate", "pcode", "product", "pack", "batchno",
        "quantity", "freeqty", "expqty",
        "pts", "ptsvalue", "ptr", "ptrvalue",
        "mrp", "mrpvalue", "nrv", "nrvvalue",
        "dsort", "direct_party", "iscancelled",
    ]

    # Field mapping for data rows
    field_keys = [
        "stockist_code", "product_head", "stockist_name", "citypool",
        "team", "region", "act_region", "zonee",
        "invoiceno", "invoicedate", "pcode", "product", "pack", "batchno",
        "quantity", "freeqty", "expqty",
        "pts", "ptsvalue", "ptr", "ptrvalue",
        "mrp", "mrpvalue", "nrv", "nrvvalue",
        "dsort", "direct_party", "iscancelled",
    ]

    # Header styling
    header_font = Font(bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    # Write headers
    for col_idx, header in enumerate(excel_headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    # Write data
    for row_idx, row in enumerate(rows, 2):
        for col_idx, key in enumerate(field_keys, 1):
            val = row.get(key, "")
            if key in ("direct_party", "iscancelled"):
                val = True if val else False
            elif key == "invoicedate" and val:
                val = str(val)
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.border = thin_border

    # Auto-width columns
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_length + 3, 30)

    # Freeze header row
    ws.freeze_panes = "A2"

    # Save to temp file and serve
    fname = f"Primary_Sales_{division}_{month}.xlsx"
    file_path = os.path.join(tempfile.gettempdir(), fname)
    wb.save(file_path)

    with open(file_path, "rb") as f:
        file_content = f.read()

    # Clean up
    try:
        os.remove(file_path)
    except OSError:
        pass

    frappe.local.response.filename = fname
    frappe.local.response.filecontent = file_content
    frappe.local.response.type = "download"


@frappe.whitelist()
def save_primary_sales_record(name=None, data=None):
    """Create or update a Primary Sales Data record."""
    try:
        data = frappe.parse_json(data) if isinstance(data, str) else data
        if not data:
            return {"success": False, "message": "No data provided"}

        current_division = get_user_division() or "Prima"
        data["division"] = current_division

        # Numeric fields
        num_fields = [
            "quantity", "freeqty", "expqty", "pts", "ptsvalue",
            "ptr", "ptrvalue", "mrp", "mrpvalue", "nrv", "nrvvalue", "dsort"
        ]
        for nf in num_fields:
            val = data.get(nf)
            if val is not None and val != "":
                try:
                    data[nf] = float(val)
                except (ValueError, TypeError):
                    data[nf] = 0
            else:
                data[nf] = 0

        # Boolean fields
        for bf in ["iscancelled", "direct_party"]:
            val = data.get(bf)
            data[bf] = 1 if val in (1, "1", True, "true", "True") else 0

        if name:
            doc = frappe.get_doc("Primary Sales Data", name)
            doc.update(data)
        else:
            doc = frappe.new_doc("Primary Sales Data")
            doc.update(data)

        doc.save(ignore_permissions=False)
        frappe.db.commit()

        return {"success": True, "message": "Record saved successfully", "name": doc.name}

    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Save Primary Sales Record Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def delete_primary_sales_record(name):
    """Delete a Primary Sales Data record."""
    try:
        frappe.delete_doc("Primary Sales Data", name, ignore_permissions=False)
        frappe.db.commit()
        return {"success": True, "message": "Record deleted successfully"}
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Delete Primary Sales Record Error")
        return {"success": False, "message": str(e)}


# ──────────────────────────────────────────────────────────────────────────────
# STOCKIST REPORTS
# ──────────────────────────────────────────────────────────────────────────────

@frappe.whitelist()
def get_stockist_report_filter_options(division=None):
    """Return dropdown options for the Stockist Reports page (active masters only)."""
    if not division:
        division = get_user_division()

    regions = frappe.db.sql(
        "SELECT DISTINCT name, region_name FROM `tabRegion Master` WHERE status='Active' AND division IN (%s, 'Both') ORDER BY name",
        (division,), as_dict=True,
    )
    teams = frappe.db.sql(
        "SELECT DISTINCT name, team_name FROM `tabTeam Master` WHERE status='Active' AND division IN (%s, 'Both') ORDER BY name",
        (division,), as_dict=True,
    )
    hqs = frappe.db.sql(
        "SELECT DISTINCT name, hq_name FROM `tabHQ Master` WHERE division=%s AND status='Active' ORDER BY name",
        (division,), as_dict=True,
    )
    zones = frappe.db.sql(
        "SELECT DISTINCT name, zone_name FROM `tabZone Master` WHERE status='Active' AND division IN (%s, 'Both') ORDER BY name",
        (division,), as_dict=True,
    )
    stockists = frappe.db.sql(
        "SELECT name, stockist_name FROM `tabStockist Master` WHERE division=%s AND status='Active' ORDER BY stockist_name",
        (division,), as_dict=True,
    )
    months = frappe.db.sql(
        "SELECT DISTINCT upload_month FROM `tabPrimary Sales Data` WHERE division=%s ORDER BY upload_month DESC",
        (division,), as_list=1,
    )
    statement_months = frappe.db.sql(
        "SELECT DISTINCT DATE_FORMAT(statement_month, '%%Y-%%m-%%d') as m FROM `tabStockist Statement` WHERE division=%s AND docstatus=1 ORDER BY statement_month DESC",
        (division,), as_list=1,
    )

    return {
        "regions": [{"code": r.name, "name": r.region_name} for r in regions],
        "teams": [{"code": t.name, "name": t.team_name} for t in teams],
        "hqs": [{"code": h.name, "name": h.hq_name} for h in hqs],
        "zones": [{"code": z.name, "name": z.zone_name} for z in zones],
        "stockists": [{
            "code": s.name, "name": s.stockist_name} for s in stockists],
        "months": [m[0] for m in months],
        "statement_months": [m[0] for m in statement_months],
    }


@frappe.whitelist()
def get_stockist_primary_sales_report(division=None, sales_type="primary", region=None,
                                       from_date=None, to_date=None):
    """Report 1 – Stockist Wise Primary Sales Report.
    sales_type: 'primary' (iscancelled=0) or 'creditnote' (iscancelled=1).
    """
    if not division:
        division = get_user_division()
    is_cancelled = 1 if sales_type == "creditnote" else 0

    conditions = ["division = %(division)s", "iscancelled = %(is_cancelled)s"]
    params = {"division": division, "is_cancelled": is_cancelled}

    if region:
        conditions.append("region = %(region)s")
        params["region"] = region
    if from_date:
        conditions.append("invoicedate >= %(from_date)s")
        params["from_date"] = from_date
    if to_date:
        conditions.append("invoicedate <= %(to_date)s")
        params["to_date"] = to_date

    where = " AND ".join(conditions)

    rows = frappe.db.sql(f"""
        SELECT stockist_code, stockist_name, pcode AS product_code,
               product AS product_name, pack,
               SUM(quantity) AS total_qty, SUM(ptsvalue) AS total_value
        FROM `tabPrimary Sales Data`
        WHERE {where}
        GROUP BY stockist_code, stockist_name, pcode, product, pack
        ORDER BY stockist_name, pcode
    """, params, as_dict=True)

    return {"success": True, "data": rows}


@frappe.whitelist()
def get_stockist_secondary_sales_report(division=None, region=None,
                                         from_date=None, to_date=None):
    """Report 2 – Stockist Wise Secondary Sales Report (from Draft or submitted Stockist Statements)."""
    if not division:
        division = get_user_division()

    conditions = ["ss.division = %(division)s", "ss.docstatus IN (0, 1)"]
    params = {"division": division}

    if region:
        conditions.append("ss.region = %(region)s")
        params["region"] = region
    if from_date:
        conditions.append("ss.statement_month >= %(from_date)s")
        params["from_date"] = from_date
    if to_date:
        conditions.append("ss.statement_month <= %(to_date)s")
        params["to_date"] = to_date

    where = " AND ".join(conditions)

    rows = frappe.db.sql(f"""
        SELECT ss.stockist_code, ss.stockist_name,
               si.product_code, si.product_name, si.pack,
               SUM(si.sales_qty) AS total_qty,
               SUM(si.sales_value_pts) AS total_value
        FROM `tabStockist Statement` ss
        INNER JOIN `tabStockist Statement Item` si
            ON si.parent = ss.name AND si.parenttype = 'Stockist Statement'
        WHERE {where}
        GROUP BY ss.stockist_code, ss.stockist_name,
                 si.product_code, si.product_name, si.pack
        ORDER BY ss.stockist_name, si.product_code
    """, params, as_dict=True)

    return {"success": True, "data": rows}


@frappe.whitelist()
def get_stockist_moving_trend_report(division=None, sales_type="secondary",
                                      stockist_code=None):
    """Report 3 – Moving Trend (monthly pivot Apr-Mar) for a single stockist.
    sales_type: 'primary' uses Primary Sales Data; 'secondary' uses Statement Items.
    """
    if not division:
        division = get_user_division()
    if not stockist_code:
        return {"success": False, "message": "Stockist code is required"}

    month_labels = ["apr", "may", "jun", "jul", "aug", "sep",
                    "oct", "nov", "dec", "jan", "feb", "mar"]

    if sales_type == "primary":
        rows = frappe.db.sql("""
            SELECT pcode AS product_code, product AS product_name, pack,
                   MONTH(invoicedate) AS m, YEAR(invoicedate) AS y,
                   SUM(quantity) AS qty
            FROM `tabPrimary Sales Data`
            WHERE division = %(division)s AND stockist_code = %(stockist)s
                  AND iscancelled = 0
            GROUP BY pcode, product, pack, MONTH(invoicedate), YEAR(invoicedate)
            ORDER BY pcode, y, m
        """, {"division": division, "stockist": stockist_code}, as_dict=True)
    else:
        rows = frappe.db.sql("""
            SELECT si.product_code, si.product_name, si.pack,
                   MONTH(ss.statement_month) AS m, YEAR(ss.statement_month) AS y,
                   SUM(si.sales_qty) AS qty
            FROM `tabStockist Statement` ss
            INNER JOIN `tabStockist Statement Item` si
                ON si.parent = ss.name AND si.parenttype = 'Stockist Statement'
            WHERE ss.division = %(division)s AND ss.stockist_code = %(stockist)s
                  AND ss.docstatus IN (0, 1)
            GROUP BY si.product_code, si.product_name, si.pack,
                     MONTH(ss.statement_month), YEAR(ss.statement_month)
            ORDER BY si.product_code, y, m
        """, {"division": division, "stockist": stockist_code}, as_dict=True)

    # Derive the financial year range from data
    if not rows:
        return {"success": True, "data": [], "fy_label": ""}

    all_dates = [(r.y, r.m) for r in rows]
    min_y = min(d[0] for d in all_dates)
    max_y = max(d[0] for d in all_dates)

    # Determine FY boundaries (Apr of min year to Mar of max year+1)
    fy_start_year = min_y if min(d[1] for d in all_dates if d[0] == min_y) >= 4 else min_y - 1
    fy_end_year = max_y + 1 if max(d[1] for d in all_dates if d[0] == max_y) >= 4 else max_y
    fy_label = f"Apr {str(fy_start_year)[2:]} to Mar {str(fy_end_year)[2:]}"

    # Map month number to FY column index: Apr(4)→0 … Mar(3)→11
    month_map = {4: 0, 5: 1, 6: 2, 7: 3, 8: 4, 9: 5,
                 10: 6, 11: 7, 12: 8, 1: 9, 2: 10, 3: 11}

    products = {}
    for r in rows:
        key = r.product_code
        if key not in products:
            products[key] = {
                "product_code": r.product_code,
                "product_name": r.product_name,
                "pack": r.pack,
                "months": [0] * 12,
                "total": 0,
            }
        idx = month_map.get(r.m)
        if idx is not None:
            products[key]["months"][idx] += flt(r.qty)
            products[key]["total"] += flt(r.qty)

    data = list(products.values())
    return {"success": True, "data": data, "fy_label": fy_label, "month_labels": month_labels}


@frappe.whitelist()
def get_stockist_closing_stock_report(division=None, region=None,
                                       from_date=None, to_date=None, group_by="stockist"):
    """Report 4 – Stockist Wise Closing Stock Report (from Draft or submitted Stockist Statements).
    group_by: 'stockist' (default) or 'hq'
    """
    if not division:
        division = get_user_division()

    conditions = ["ss.division = %(division)s", "ss.docstatus IN (0, 1)"]
    params = {"division": division}

    if region:
        conditions.append("ss.region = %(region)s")
        params["region"] = region
    if from_date:
        conditions.append("ss.statement_month >= %(from_date)s")
        params["from_date"] = from_date
    if to_date:
        conditions.append("ss.statement_month <= %(to_date)s")
        params["to_date"] = to_date

    where = " AND ".join(conditions)

    if group_by == "hq":
        rows = frappe.db.sql(f"""
            SELECT ss.hq AS hq_code,
                   COALESCE(hm.hq_name, ss.hq, '') AS hq_name,
                   si.product_code, si.product_name, si.pack,
                   SUM(si.opening_qty) AS opening_qty,
                   SUM(si.purchase_qty) AS purchase_qty,
                   SUM(si.sales_qty) AS sales_qty,
                   SUM(si.free_qty) AS free_qty,
                   SUM(si.free_qty_scheme) AS scheme_free_qty,
                   SUM(si.closing_qty) AS closing_qty
            FROM `tabStockist Statement` ss
            INNER JOIN `tabStockist Statement Item` si
                ON si.parent = ss.name AND si.parenttype = 'Stockist Statement'
            LEFT JOIN `tabHQ Master` hm ON hm.name = ss.hq AND hm.division = %(division)s
            WHERE {where}
            GROUP BY ss.hq, hm.hq_name,
                     si.product_code, si.product_name, si.pack
            ORDER BY hm.hq_name, si.product_code
        """, params, as_dict=True)
    else:
        rows = frappe.db.sql(f"""
            SELECT ss.stockist_code, ss.stockist_name,
                   si.product_code, si.product_name, si.pack,
                   SUM(si.opening_qty) AS opening_qty,
                   SUM(si.purchase_qty) AS purchase_qty,
                   SUM(si.sales_qty) AS sales_qty,
                   SUM(si.free_qty) AS free_qty,
                   SUM(si.free_qty_scheme) AS scheme_free_qty,
                   SUM(si.closing_qty) AS closing_qty
            FROM `tabStockist Statement` ss
            INNER JOIN `tabStockist Statement Item` si
                ON si.parent = ss.name AND si.parenttype = 'Stockist Statement'
            WHERE {where}
            GROUP BY ss.stockist_code, ss.stockist_name,
                     si.product_code, si.product_name, si.pack
            ORDER BY ss.stockist_name, si.product_code
        """, params, as_dict=True)

    return {"success": True, "data": rows, "group_by": group_by}


@frappe.whitelist()
def get_region_product_closing_stock(division=None, region=None, from_date=None, to_date=None, group_by="hq"):
    """Report 9 – Product Closing Stock for Region, pivoted by HQ or Stockist (grouped by Team).
    Returns products as rows, HQs/Stockists as columns grouped by team, with Team Totals and Region Total.
    group_by: 'hq' (default) or 'stockist'
    """
    if not division:
        division = get_user_division()
    if not region:
        frappe.throw("Region is required for this report.")

    conditions = ["ss.division = %(division)s", "ss.region = %(region)s", "ss.docstatus IN (0, 1)"]
    params = {"division": division, "region": region}

    if from_date:
        conditions.append("ss.statement_month >= %(from_date)s")
        params["from_date"] = from_date
    if to_date:
        conditions.append("ss.statement_month <= %(to_date)s")
        params["to_date"] = to_date

    where = " AND ".join(conditions)

    if group_by == "stockist":
        # Stockist-wise pivot: one column per stockist
        rows = frappe.db.sql(f"""
            SELECT
                ss.stockist_code AS col_code,
                ss.stockist_name AS col_name,
                COALESCE(hm.team, ss.team, '') AS team_code,
                COALESCE(tm.team_name, hm.team, ss.team, '') AS team_name,
                si.product_code,
                si.product_name,
                si.pack,
                SUM(si.closing_qty) AS closing_qty,
                SUM(COALESCE(si.closing_value, si.closing_qty * si.pts, 0)) AS closing_value
            FROM `tabStockist Statement` ss
            INNER JOIN `tabStockist Statement Item` si
                ON si.parent = ss.name AND si.parenttype = 'Stockist Statement'
            LEFT JOIN `tabHQ Master` hm ON hm.name = ss.hq AND hm.division = %(division)s
            LEFT JOIN `tabTeam Master` tm ON tm.name = COALESCE(hm.team, ss.team)
            WHERE {where}
            GROUP BY ss.stockist_code, ss.stockist_name, hm.team, tm.team_name,
                     si.product_code, si.product_name, si.pack
            ORDER BY hm.team, ss.stockist_name, si.product_code
        """, params, as_dict=True)

        # Build ordered column list from Stockist Master for this region
        col_list = frappe.db.sql("""
            SELECT sm.name AS col_code, sm.stockist_name AS col_name,
                   COALESCE(hm.team, sm.team, '') AS team_code,
                   COALESCE(tm.team_name, hm.team, sm.team, '') AS team_name
            FROM `tabStockist Master` sm
            LEFT JOIN `tabHQ Master` hm ON hm.name = sm.hq
            LEFT JOIN `tabTeam Master` tm ON tm.name = COALESCE(hm.team, sm.team)
            WHERE sm.division = %(division)s AND sm.region = %(region)s AND sm.status = 'Active'
            ORDER BY hm.team, sm.stockist_name
        """, params, as_dict=True)
    else:
        # HQ-wise pivot: one column per HQ
        rows = frappe.db.sql(f"""
            SELECT
                ss.hq AS col_code,
                COALESCE(hm.hq_name, ss.hq, '') AS col_name,
                COALESCE(hm.team, ss.team, '') AS team_code,
                COALESCE(tm.team_name, hm.team, ss.team, '') AS team_name,
                si.product_code,
                si.product_name,
                si.pack,
                SUM(si.closing_qty) AS closing_qty,
                SUM(COALESCE(si.closing_value, si.closing_qty * si.pts, 0)) AS closing_value
            FROM `tabStockist Statement` ss
            INNER JOIN `tabStockist Statement Item` si
                ON si.parent = ss.name AND si.parenttype = 'Stockist Statement'
            LEFT JOIN `tabHQ Master` hm ON hm.name = ss.hq AND hm.division = %(division)s
            LEFT JOIN `tabTeam Master` tm ON tm.name = COALESCE(hm.team, ss.team)
            WHERE {where}
            GROUP BY ss.hq, hm.hq_name, hm.team, tm.team_name,
                     si.product_code, si.product_name, si.pack
            ORDER BY hm.team, hm.hq_name, si.product_code
        """, params, as_dict=True)

        # Build ordered HQ column list from HQ Master for this region
        col_list = frappe.db.sql("""
            SELECT hm.name AS col_code, COALESCE(hm.hq_name, hm.name, '') AS col_name,
                   COALESCE(hm.team, '') AS team_code,
                   COALESCE(tm.team_name, hm.team, '') AS team_name
            FROM `tabHQ Master` hm
            LEFT JOIN `tabTeam Master` tm ON tm.name = hm.team
            WHERE hm.division = %(division)s AND hm.region = %(region)s AND hm.status = 'Active'
            ORDER BY hm.team, hm.hq_name
        """, params, as_dict=True)

    if not rows:
        return {"success": True, "products": [], "team_order": [], "hq_columns": [],
                "value_in_lakhs": {}, "region": region, "region_name": region, "group_by": group_by}

    # Fall back to data-derived column list if master has no entries
    if not col_list:
        seen = {}
        for r in rows:
            if r.col_code and r.col_code not in seen:
                seen[r.col_code] = {"col_code": r.col_code, "col_name": r.col_name,
                                    "team_code": r.team_code, "team_name": r.team_name}
        col_list = list(seen.values())

    # Build team order: only include teams/columns that actually have data
    data_col_codes = {r.col_code for r in rows if r.col_code}
    team_map = {}
    for c in col_list:
        tc = c.team_code or "Unknown"
        if tc not in team_map:
            team_map[tc] = {"team_code": tc, "team_name": c.team_name or tc, "hqs": []}
        if c.col_code in data_col_codes:
            team_map[tc]["hqs"].append({"col_code": c.col_code, "col_name": c.col_name})

    # Include any data columns not in master
    configured_cols = {c.col_code for c in col_list}
    for r in rows:
        if r.col_code and r.col_code not in configured_cols:
            tc = r.team_code or "Unknown"
            if tc not in team_map:
                team_map[tc] = {"team_code": tc, "team_name": r.team_name or tc, "hqs": []}
            if not any(h["col_code"] == r.col_code for h in team_map[tc]["hqs"]):
                team_map[tc]["hqs"].append({"col_code": r.col_code, "col_name": r.col_name})

    # Filter teams with at least one column
    team_order = [t for t in team_map.values() if t["hqs"]]
    flat_col_codes = [h["col_code"] for t in team_order for h in t["hqs"]]

    # Build product data matrix
    product_data = {}
    for r in rows:
        pc = r.product_code
        if pc not in product_data:
            product_data[pc] = {
                "product_code": pc,
                "product_name": r.product_name,
                "pack": r.pack,
                "col_qty": {},
                "col_value": {},
            }
        cc = r.col_code
        product_data[pc]["col_qty"][cc] = product_data[pc]["col_qty"].get(cc, 0) + (float(r.closing_qty) if r.closing_qty else 0)
        product_data[pc]["col_value"][cc] = product_data[pc]["col_value"].get(cc, 0) + (float(r.closing_value) if r.closing_value else 0)

    # Sort products by product_code (guard against None values)
    sorted_products = sorted(product_data.values(), key=lambda x: x["product_code"] or "")

    # Compute value_in_lakhs (total across all products per column/team/region)
    col_total_value = {}
    for pdata in sorted_products:
        for cc, val in pdata["col_value"].items():
            col_total_value[cc] = col_total_value.get(cc, 0) + val

    team_total_value = {}
    for t in team_order:
        tv = sum(col_total_value.get(h["col_code"], 0) for h in t["hqs"])
        team_total_value[t["team_code"]] = tv

    region_total_value = sum(col_total_value.values())

    # Build product rows
    product_rows = []
    for sno, pdata in enumerate(sorted_products, 1):
        team_qty = {}
        for t in team_order:
            team_qty[t["team_code"]] = sum(pdata["col_qty"].get(h["col_code"], 0) for h in t["hqs"])
        region_qty = sum(pdata["col_qty"].values())
        product_rows.append({
            "sno": sno,
            "product_code": pdata["product_code"],
            "product_name": pdata["product_name"],
            "pack": pdata["pack"],
            "col_qty": pdata["col_qty"],
            "team_qty": team_qty,
            "region_qty": region_qty,
        })

    region_name = frappe.db.get_value("Region Master", region, "region_name") or region

    return {
        "success": True,
        "region": region,
        "region_name": region_name,
        "from_date": from_date,
        "to_date": to_date,
        "group_by": group_by,
        "team_order": team_order,
        "hq_columns": flat_col_codes,
        "products": product_rows,
        "value_in_lakhs": {
            "col_values": {cc: round(v / 100000, 2) for cc, v in col_total_value.items()},
            "team_values": {tc: round(v / 100000, 2) for tc, v in team_total_value.items()},
            "region_total": round(region_total_value / 100000, 2),
        },
    }


@frappe.whitelist()
def get_hq_wise_stockist_report(division=None, region=None):
    """Report 5 – HQ Wise Stockist Report (active stockists grouped by HQ)."""
    if not division:
        division = get_user_division()

    conditions = ["sm.division = %(division)s", "sm.status = 'Active'"]
    params = {"division": division}

    if region:
        conditions.append("hm.region = %(region)s")
        params["region"] = region

    where = " AND ".join(conditions)

    rows = frappe.db.sql(f"""
        SELECT sm.name AS stockist_code, sm.stockist_name,
               sm.hq, COALESCE(hm.hq_name, sm.hq) AS hq_name,
               COALESCE(hm.team, '') AS team,
               COALESCE(hm.region, sm.region) AS region
        FROM `tabStockist Master` sm
        LEFT JOIN `tabHQ Master` hm ON hm.name = sm.hq
        WHERE {where}
        ORDER BY hm.hq_name, sm.stockist_name
    """, params, as_dict=True)

    # Group by HQ
    grouped = {}
    for r in rows:
        hq_key = r.hq or "Unassigned"
        if hq_key not in grouped:
            grouped[hq_key] = {
                "hq_code": r.hq,
                "hq_name": r.hq_name,
                "team": r.team,
                "region": r.region,
                "stockists": [],
            }
        grouped[hq_key]["stockists"].append({
            "stockist_code": r.stockist_code,
            "stockist_name": r.stockist_name,
        })

    return {"success": True, "data": list(grouped.values())}


@frappe.whitelist()
def get_stockist_address_report(division=None, region=None, criteria="ALL"):
    """Report 6 – Stockist Address Report.
    criteria: ALL, HQ WISE, TEAM WISE, CITY WISE, STOCKIST NAME
    """
    if not division:
        division = get_user_division()

    conditions = ["sm.division = %(division)s", "sm.status = 'Active'"]
    params = {"division": division}

    if region:
        conditions.append("sm.region = %(region)s")
        params["region"] = region

    where = " AND ".join(conditions)

    order_map = {
        "ALL": "sm.stockist_name",
        "HQ WISE": "sm.hq, sm.stockist_name",
        "TEAM WISE": "sm.team, sm.stockist_name",
        "CITY WISE": "sm.city, sm.stockist_name",
        "STOCKIST NAME": "sm.stockist_name",
    }
    order_by = order_map.get(criteria, "sm.stockist_name")

    rows = frappe.db.sql(f"""
        SELECT sm.name AS stockist_code, sm.stockist_name,
               sm.address, sm.city, sm.phone,
               sm.hq, sm.team, sm.region
        FROM `tabStockist Master` sm
        WHERE {where}
        ORDER BY {order_by}
    """, params, as_dict=True)

    # If grouping is needed, build groups
    group_field = None
    if criteria == "HQ WISE":
        group_field = "hq"
    elif criteria == "TEAM WISE":
        group_field = "team"
    elif criteria == "CITY WISE":
        group_field = "city"

    if group_field:
        grouped = {}
        for r in rows:
            key = r.get(group_field) or "Unassigned"
            grouped.setdefault(key, []).append(r)
        return {"success": True, "data": rows, "groups": grouped, "group_field": group_field}

    return {"success": True, "data": rows, "groups": None, "group_field": None}


@frappe.whitelist()
def export_stockist_report_excel(report_type, division=None, **kwargs):
    """Generate a styled Excel workbook for any of the 6 stockist report types."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    if not division:
        division = get_user_division()

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_font = Font(bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    group_font = Font(bold=True, size=11)
    group_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")

    wb = openpyxl.Workbook()
    ws = wb.active

    def write_header_row(ws, row_num, headers):
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=row_num, column=col, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border

    def write_data_row(ws, row_num, values):
        for col, v in enumerate(values, 1):
            cell = ws.cell(row=row_num, column=col, value=v)
            cell.border = thin_border

    def write_group_row(ws, row_num, label, num_cols):
        cell = ws.cell(row=row_num, column=1, value=label)
        cell.font = group_font
        cell.fill = group_fill
        for c in range(1, num_cols + 1):
            ws.cell(row=row_num, column=c).fill = group_fill
            ws.cell(row=row_num, column=c).border = thin_border

    def write_title_rows(ws, title, subtitle=""):
        ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=14)
        if subtitle:
            ws.cell(row=2, column=1, value=subtitle).font = Font(size=11, italic=True)
        return 4  # data starts at row 4

    region = kwargs.get("region", "")
    from_date = kwargs.get("from_date", "")
    to_date = kwargs.get("to_date", "")
    sales_type = kwargs.get("sales_type", "primary")
    stockist_code = kwargs.get("stockist_code", "")
    criteria = kwargs.get("criteria", "ALL")
    period_label = f"{from_date} to {to_date}" if from_date and to_date else ""
    region_label = region or "All Regions"

    if report_type == "primary_sales":
        ws.title = "Primary Sales"
        result = get_stockist_primary_sales_report(division, sales_type, region, from_date, to_date)
        data = result.get("data", [])
        type_label = "Credit Note" if sales_type == "creditnote" else "Primary"
        row = write_title_rows(ws, f"Stockist Wise {type_label} Sales Report – {division}",
                               f"Region: {region_label}  |  Period: {period_label}")
        headers = ["Stockist Code", "Stockist Name", "Product Code", "Product Name", "Pack", "Quantity", "Value (PTS)"]
        write_header_row(ws, row, headers)
        row += 1
        current_stockist = None
        for d in data:
            if d.stockist_code != current_stockist:
                current_stockist = d.stockist_code
                write_group_row(ws, row, f"{d.stockist_name} ({d.stockist_code})", len(headers))
                row += 1
            write_data_row(ws, row, [d.stockist_code, d.stockist_name, d.product_code,
                                     d.product_name, d.pack, flt(d.total_qty), flt(d.total_value)])
            row += 1

    elif report_type == "secondary_sales":
        ws.title = "Secondary Sales"
        result = get_stockist_secondary_sales_report(division, region, from_date, to_date)
        data = result.get("data", [])
        row = write_title_rows(ws, f"Stockist Wise Secondary Sales Report – {division}",
                               f"Region: {region_label}  |  Period: {period_label}")
        headers = ["Stockist Code", "Stockist Name", "Product Code", "Product Name", "Pack", "Quantity", "Value (PTS)"]
        write_header_row(ws, row, headers)
        row += 1
        current_stockist = None
        for d in data:
            if d.stockist_code != current_stockist:
                current_stockist = d.stockist_code
                write_group_row(ws, row, f"{d.stockist_name} ({d.stockist_code})", len(headers))
                row += 1
            write_data_row(ws, row, [d.stockist_code, d.stockist_name, d.product_code,
                                     d.product_name, d.pack, flt(d.total_qty), flt(d.total_value)])
            row += 1

    elif report_type == "moving_trend":
        ws.title = "Moving Trend"
        result = get_stockist_moving_trend_report(division, sales_type, stockist_code)
        data = result.get("data", [])
        fy_label = result.get("fy_label", "")
        ml = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]
        type_label = "Primary" if sales_type == "primary" else "Secondary"
        row = write_title_rows(ws, f"Moving Trend ({type_label} Sales) – {division}",
                               f"Stockist: {stockist_code}  |  {fy_label}")
        headers = ["Product Code", "Product Name", "Pack"] + ml + ["Total"]
        write_header_row(ws, row, headers)
        row += 1
        for d in data:
            vals = [d["product_code"], d["product_name"], d["pack"]] + d["months"] + [d["total"]]
            write_data_row(ws, row, vals)
            row += 1

    elif report_type == "closing_stock":
        ws.title = "Closing Stock"
        result = get_stockist_closing_stock_report(division, region, from_date, to_date)
        data = result.get("data", [])
        row = write_title_rows(ws, f"Stockist Wise Closing Stock Report – {division}",
                               f"Region: {region_label}  |  Period: {period_label}")
        headers = ["Stockist Code", "Stockist Name", "Product Code", "Product Name", "Pack",
                    "Opening Qty", "Purchase Qty", "Sales Qty", "Free Qty", "Scheme Free Qty", "Closing Qty"]
        write_header_row(ws, row, headers)
        row += 1
        current_stockist = None
        for d in data:
            if d.stockist_code != current_stockist:
                current_stockist = d.stockist_code
                write_group_row(ws, row, f"{d.stockist_name} ({d.stockist_code})", len(headers))
                row += 1
            write_data_row(ws, row, [d.stockist_code, d.stockist_name, d.product_code, d.product_name,
                                     d.pack, flt(d.opening_qty), flt(d.purchase_qty), flt(d.sales_qty),
                                     flt(d.free_qty), flt(d.scheme_free_qty), flt(d.closing_qty)])
            row += 1

    elif report_type == "hq_wise":
        ws.title = "HQ Wise Stockists"
        result = get_hq_wise_stockist_report(division, region)
        data = result.get("data", [])
        row = write_title_rows(ws, f"HQ Wise Stockist Report – {division}",
                               f"Region: {region_label}")
        headers = ["HQ", "HQ Name", "Team", "Region", "Stockist Code", "Stockist Name"]
        write_header_row(ws, row, headers)
        row += 1
        for hq_group in data:
            write_group_row(ws, row, f"{hq_group['hq_name']} ({hq_group['hq_code']})", len(headers))
            row += 1
            for s in hq_group["stockists"]:
                write_data_row(ws, row, [hq_group["hq_code"], hq_group["hq_name"],
                                         hq_group["team"], hq_group["region"],
                                         s["stockist_code"], s["stockist_name"]])
                row += 1

    elif report_type == "address":
        ws.title = "Stockist Address"
        result = get_stockist_address_report(division, region, criteria)
        data = result.get("data", [])
        groups = result.get("groups")
        row = write_title_rows(ws, f"Stockist Address Report – {division}",
                               f"Region: {region_label}  |  Criteria: {criteria}")
        headers = ["Stockist Code", "Stockist Name", "Address", "City", "Phone", "HQ", "Team", "Region"]
        write_header_row(ws, row, headers)
        row += 1
        if groups:
            for grp_name, grp_rows in groups.items():
                write_group_row(ws, row, str(grp_name), len(headers))
                row += 1
                for d in grp_rows:
                    write_data_row(ws, row, [d.stockist_code, d.stockist_name, d.address or "",
                                             d.city or "", d.phone or "", d.hq or "", d.team or "", d.region or ""])
                    row += 1
        else:
            for d in data:
                write_data_row(ws, row, [d.stockist_code, d.stockist_name, d.address or "",
                                         d.city or "", d.phone or "", d.hq or "", d.team or "", d.region or ""])
                row += 1
    else:
        frappe.throw("Invalid report type")

    # Auto-fit column widths
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_len + 4, 40)

    # Save to temp file and respond
    from io import BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    xlsx_data = output.getvalue()

    filename = f"Stockist_Report_{report_type}_{division}.xlsx"
    frappe.local.response.filename = filename
    frappe.local.response.filecontent = xlsx_data
    frappe.local.response.type = "download"
    frappe.local.response.content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ═══════════════════════════════════════════════════════════════
# SCHEME REPORTS  –  Portal API Methods
# ═══════════════════════════════════════════════════════════════

def get_scheme_report_filter_options(division=None):
    """Return dropdown options for the Scheme Reports portal page (active masters only)."""
    if not division:
        division = get_user_division()

    zones = frappe.db.sql(
        "SELECT name FROM `tabZone Master` WHERE status='Active' AND division IN (%s, 'Both') ORDER BY name",
        (division,), as_list=1,
    )
    regions = frappe.db.sql(
        "SELECT name, region_name, zone FROM `tabRegion Master` WHERE status='Active' AND division IN (%s, 'Both') ORDER BY name",
        (division,), as_dict=True,
    )
    teams = frappe.db.sql(
        "SELECT name, region FROM `tabTeam Master` WHERE status='Active' AND division IN (%s, 'Both') ORDER BY name",
        (division,), as_dict=True,
    )
    hqs = frappe.db.sql(
        "SELECT name, hq_name, team FROM `tabHQ Master` WHERE division=%s AND status='Active' ORDER BY name",
        (division,), as_dict=True,
    )
    products = frappe.db.sql(
        "SELECT product_code, product_name, category, product_group, pack "
        "FROM `tabProduct Master` WHERE division=%s AND status='Active' ORDER BY product_name",
        (division,), as_dict=True,
    )

    product_groups = sorted(set(p.product_group for p in products if p.product_group))

    return {
        "zones": [z[0] for z in zones],
        "regions": [{"name": r.name, "region_name": r.region_name or r.name, "zone": r.zone or ""} for r in regions],
        "teams": [{"name": t.name, "region": t.region or ""} for t in teams],
        "hqs": [{"name": h.name, "hq_name": h.hq_name or "", "team": h.team or ""} for h in hqs],
        "products": [{"code": p.product_code, "name": p.product_name,
                       "category": p.category or "", "group": p.product_group or "",
                       "pack": p.pack or ""} for p in products],
        "product_groups": product_groups,
    }


@frappe.whitelist()
def get_scheme_activity_trend_report(division=None, from_date=None, to_date=None,
                                      doctor_status="Active", product_type="All",
                                      product_codes=None, zone=None, region=None,
                                      criteria="Region", team_or_hq=None):
    """Report 1 – Activity Trend Report.
    Monthly pivot (Apr–Mar) showing product-qty pairs per doctor.
    """
    if not division:
        division = get_user_division()
    if not from_date or not to_date:
        return {"success": False, "message": "From Date and To Date are required"}

    conditions = [
        "sr.division = %(division)s",
        "sr.docstatus = 1",
        "sr.approval_status = 'Approved'",
        "sr.application_date BETWEEN %(from_date)s AND %(to_date)s",
    ]
    params = {"division": division, "from_date": from_date, "to_date": to_date}

    # Doctor status filter
    if doctor_status and doctor_status != "All":
        conditions.append("dm.status = %(doctor_status)s")
        params["doctor_status"] = doctor_status

    # Product type filter (category)
    if product_type and product_type != "All":
        cat_map = {"Hospital Products": "Hospital Product", "Other Products": "Main Product"}
        if product_type in cat_map:
            conditions.append("pm.category = %(product_category)s")
            params["product_category"] = cat_map[product_type]

    # Selected products filter
    if product_codes:
        if isinstance(product_codes, str):
            product_codes = json.loads(product_codes)
        if product_codes:
            placeholders = ", ".join(["%(pc_" + str(i) + ")s" for i in range(len(product_codes))])
            conditions.append("sri.product_code IN (" + placeholders + ")")
            for i, pc in enumerate(product_codes):
                params["pc_" + str(i)] = pc

    # Hierarchy filters
    if zone:
        conditions.append("dm.zone = %(zone)s")
        params["zone"] = zone
    if region:
        conditions.append("dm.region = %(region)s")
        params["region"] = region
    if team_or_hq and criteria in ("Team", "HQ"):
        if criteria == "Team":
            conditions.append("dm.team = %(team_or_hq)s")
        else:
            conditions.append("dm.hq = %(team_or_hq)s")
        params["team_or_hq"] = team_or_hq

    where = " AND ".join(conditions)

    rows = frappe.db.sql(f"""
        SELECT dm.region, hm.hq_name, dm.hq, dm.doctor_name, dm.name AS doctor_code,
               sri.product_code, sri.free_quantity,
               MONTH(sr.application_date) AS m, YEAR(sr.application_date) AS y
        FROM `tabScheme Request` sr
        INNER JOIN `tabScheme Request Item` sri ON sri.parent = sr.name
        INNER JOIN `tabDoctor Master` dm ON dm.name = sr.doctor_code
        INNER JOIN `tabHQ Master` hm ON hm.name = dm.hq
        LEFT JOIN `tabProduct Master` pm ON pm.product_code = sri.product_code AND pm.division = %(division)s
        WHERE {where}
        ORDER BY dm.region, hm.hq_name, dm.doctor_name, sr.application_date
    """, params, as_dict=True)

    # Build FY label from from_date
    from datetime import datetime
    fd = datetime.strptime(str(from_date), "%Y-%m-%d")
    td = datetime.strptime(str(to_date), "%Y-%m-%d")
    fy_label = f"From {fd.strftime('%B %Y')} To {td.strftime('%B %Y')}"

    # Month index: Apr=0..Mar=11
    month_idx = {4: 0, 5: 1, 6: 2, 7: 3, 8: 4, 9: 5, 10: 6, 11: 7, 12: 8, 1: 9, 2: 10, 3: 11}

    # Aggregate: key=(region, hq, doctor_name) → months[12] = list of "PROD-QTY" strings
    from collections import defaultdict
    doc_data = defaultdict(lambda: [[] for _ in range(12)])

    for r in rows:
        key = (r.region or "", r.hq or "", r.hq_name or "", r.doctor_name or "", r.doctor_code or "")
        idx = month_idx.get(r.m, 0)
        doc_data[key][idx].append(f"{r.product_code}-{int(r.free_quantity or 0)}")

    data = []
    for (reg, hq, hq_name, doctor_name, doctor_code), months in doc_data.items():
        cells = ["\n".join(sorted(set(m))) if m else "" for m in months]
        data.append({
            "region": reg, "hq": hq, "hq_name": hq_name,
            "doctor_name": doctor_name, "doctor_code": doctor_code,
            "months": cells,
        })

    # Sort by region, hq, doctor
    data.sort(key=lambda x: (x["region"], x["hq"], x["doctor_name"]))

    return {"success": True, "data": data, "fy_label": fy_label, "criteria": criteria}


@frappe.whitelist()
def get_scheme_activity_track_report(division=None, from_date=None, to_date=None,
                                      doctor_status="Active", product_type="All",
                                      product_codes=None, zone=None, region=None,
                                      criteria="Region", team_or_hq=None):
    """Report 2 – Activity Track Report.
    Transaction‐level rows: one per Scheme Request Item with full details.
    """
    if not division:
        division = get_user_division()
    if not from_date or not to_date:
        return {"success": False, "message": "From Date and To Date are required"}

    conditions = [
        "sr.division = %(division)s",
        "sr.docstatus = 1",
        "sr.approval_status = 'Approved'",
        "sr.application_date BETWEEN %(from_date)s AND %(to_date)s",
    ]
    params = {"division": division, "from_date": from_date, "to_date": to_date}

    if doctor_status and doctor_status != "All":
        conditions.append("dm.status = %(doctor_status)s")
        params["doctor_status"] = doctor_status
    if product_type and product_type != "All":
        cat_map = {"Hospital Products": "Hospital Product", "Other Products": "Main Product"}
        if product_type in cat_map:
            conditions.append("pm.category = %(product_category)s")
            params["product_category"] = cat_map[product_type]
    if product_codes:
        if isinstance(product_codes, str):
            product_codes = json.loads(product_codes)
        if product_codes:
            placeholders = ", ".join(["%(pc_" + str(i) + ")s" for i in range(len(product_codes))])
            conditions.append("sri.product_code IN (" + placeholders + ")")
            for i, pc in enumerate(product_codes):
                params["pc_" + str(i)] = pc
    if zone:
        conditions.append("dm.zone = %(zone)s")
        params["zone"] = zone
    if region:
        conditions.append("dm.region = %(region)s")
        params["region"] = region
    if team_or_hq and criteria in ("Team", "HQ"):
        if criteria == "Team":
            conditions.append("dm.team = %(team_or_hq)s")
        else:
            conditions.append("dm.hq = %(team_or_hq)s")
        params["team_or_hq"] = team_or_hq

    where = " AND ".join(conditions)

    rows = frappe.db.sql(f"""
        SELECT sr.application_date AS date, dm.region, dm.hq,
               dm.doctor_name, sri.product_code,
               COALESCE(pm.product_name, sri.product_name) AS product_name,
               sri.quantity, sri.free_quantity, sri.product_rate AS rate,
               sri.product_value AS value,
               sr.stockist_code, COALESCE(sm.stockist_name, '') AS stockist_name,
               dm.team
        FROM `tabScheme Request` sr
        INNER JOIN `tabScheme Request Item` sri ON sri.parent = sr.name
        INNER JOIN `tabDoctor Master` dm ON dm.name = sr.doctor_code
        LEFT JOIN `tabProduct Master` pm ON pm.product_code = sri.product_code AND pm.division = %(division)s
        LEFT JOIN `tabStockist Master` sm ON sm.name = sr.stockist_code
        WHERE {where}
        ORDER BY dm.team, dm.hq, sr.application_date, dm.doctor_name
    """, params, as_dict=True)

    # Add serial numbers and compute totals
    data = []
    total_qty = total_free = total_value = 0
    for i, r in enumerate(rows, 1):
        data.append({
            "sno": i,
            "date": str(r.date) if r.date else "",
            "region": r.region or "",
            "hq": r.hq or "",
            "doctor_name": r.doctor_name or "",
            "product_code": r.product_code or "",
            "product_name": r.product_name or "",
            "qty": flt(r.quantity),
            "free_qty": flt(r.free_quantity),
            "rate": flt(r.rate),
            "value": flt(r.value),
            "stockist_name": r.stockist_name or "",
            "team": r.team or "",
        })
        total_qty += flt(r.quantity)
        total_free += flt(r.free_quantity)
        total_value += flt(r.value)

    return {
        "success": True, "data": data,
        "totals": {"qty": total_qty, "free_qty": total_free, "value": total_value},
        "criteria": criteria,
    }


@frappe.whitelist()
def get_new_approval_doctors_report(division=None, from_date=None, to_date=None,
                                     product_codes=None, zone=None, region=None,
                                     criteria="Region", team_or_hq=None):
    """Report 3 – New Approval Doctors.
    Doctors whose first-ever approved scheme falls within the given date range.
    """
    if not division:
        division = get_user_division()
    if not from_date or not to_date:
        return {"success": False, "message": "From Date and To Date are required"}

    params = {"division": division, "from_date": from_date, "to_date": to_date}

    # Subquery: first ever approved date per doctor in this division
    # Then filter doctors whose first date falls in the range
    hierarchy_conds = ["dm.status = 'Active'"]
    if zone:
        hierarchy_conds.append("dm.zone = %(zone)s")
        params["zone"] = zone
    if region:
        hierarchy_conds.append("dm.region = %(region)s")
        params["region"] = region
    if team_or_hq and criteria in ("Team", "HQ"):
        if criteria == "Team":
            hierarchy_conds.append("dm.team = %(team_or_hq)s")
        else:
            hierarchy_conds.append("dm.hq = %(team_or_hq)s")
        params["team_or_hq"] = team_or_hq

    # Product filter on scheme items (optional)
    product_join = ""
    product_cond = ""
    if product_codes:
        if isinstance(product_codes, str):
            product_codes = json.loads(product_codes)
        if product_codes:
            placeholders = ", ".join(["%(pc_" + str(i) + ")s" for i in range(len(product_codes))])
            product_join = "INNER JOIN `tabScheme Request Item` sri ON sri.parent = sr.name"
            product_cond = "AND sri.product_code IN (" + placeholders + ")"
            for i, pc in enumerate(product_codes):
                params["pc_" + str(i)] = pc

    hierarchy_where = " AND ".join(hierarchy_conds)

    rows = frappe.db.sql(f"""
        SELECT dm.name AS doctor_code, dm.doctor_name, dm.place, dm.hospital_address,
               dm.hq, dm.team, dm.region, dm.zone,
               first_scheme.first_date, first_scheme.approved_by
        FROM `tabDoctor Master` dm
        INNER JOIN (
            SELECT sr.doctor_code,
                   MIN(sr.application_date) AS first_date,
                   SUBSTRING_INDEX(GROUP_CONCAT(sr.modified_by ORDER BY sr.application_date), ',', 1) AS approved_by
            FROM `tabScheme Request` sr
            {product_join}
            WHERE sr.division = %(division)s
              AND sr.docstatus = 1
              AND sr.approval_status = 'Approved'
              {product_cond}
            GROUP BY sr.doctor_code
            HAVING first_date BETWEEN %(from_date)s AND %(to_date)s
        ) first_scheme ON first_scheme.doctor_code = dm.name
        WHERE {hierarchy_where}
        ORDER BY dm.region, dm.hq, dm.doctor_name
    """, params, as_dict=True)

    data = []
    for i, r in enumerate(rows, 1):
        data.append({
            "sno": i,
            "approval_date": str(r.first_date) if r.first_date else "",
            "region": r.region or "",
            "hq": r.hq or "",
            "doctor_name": r.doctor_name or "",
            "hospital": r.hospital_address or "",
            "city": r.place or "",
            "approved_by": r.approved_by or "",
            "team": r.team or "",
            "zone": r.zone or "",
        })

    return {"success": True, "data": data}


@frappe.whitelist()
def get_scheme_periodic_report(division=None, from_date=None, to_date=None,
                                group_by="HQ", zone=None, region=None,
                                product_codes=None):
    """Report 4 – Periodic Report.
    Scheme summary aggregated by HQ / Team / Region / Doctor / Stockist / Month / Value.
    """
    if not division:
        division = get_user_division()
    if not from_date or not to_date:
        return {"success": False, "message": "From Date and To Date are required"}

    conditions = [
        "sr.division = %(division)s",
        "sr.docstatus = 1",
        "sr.approval_status = 'Approved'",
        "sr.application_date BETWEEN %(from_date)s AND %(to_date)s",
    ]
    params = {"division": division, "from_date": from_date, "to_date": to_date}

    if zone:
        conditions.append("dm.zone = %(zone)s")
        params["zone"] = zone
    if region:
        conditions.append("dm.region = %(region)s")
        params["region"] = region
    if product_codes:
        if isinstance(product_codes, str):
            product_codes = json.loads(product_codes)
        if product_codes:
            placeholders = ", ".join(["%(pc_" + str(i) + ")s" for i in range(len(product_codes))])
            conditions.append("sri.product_code IN (" + placeholders + ")")
            for i, pc in enumerate(product_codes):
                params["pc_" + str(i)] = pc

    where = " AND ".join(conditions)

    # Dynamic GROUP BY mapping
    group_map = {
        "HQ":       {"select": "dm.hq AS group_key, hm.hq_name AS group_label", "group": "dm.hq"},
        "Team":     {"select": "dm.team AS group_key, dm.team AS group_label", "group": "dm.team"},
        "Region":   {"select": "dm.region AS group_key, dm.region AS group_label", "group": "dm.region"},
        "Doctor":   {"select": "dm.name AS group_key, dm.doctor_name AS group_label, dm.hq, dm.region", "group": "dm.name"},
        "Stockist": {"select": "sr.stockist_code AS group_key, COALESCE(sm.stockist_name,'') AS group_label", "group": "sr.stockist_code"},
        "Month":    {"select": "DATE_FORMAT(sr.application_date, '%%Y-%%m') AS group_key, DATE_FORMAT(sr.application_date, '%%b %%Y') AS group_label", "group": "DATE_FORMAT(sr.application_date, '%%Y-%%m')"},
    }

    gb = group_map.get(group_by, group_map["HQ"])
    extra_select = gb["select"]
    group_clause = gb["group"]

    # Value mode: no grouping, just sorted by value desc
    if group_by == "Value":
        rows = frappe.db.sql(f"""
            SELECT sr.name AS group_key, CONCAT(sr.name, ' - ', dm.doctor_name) AS group_label,
                   dm.hq, dm.region,
                   SUM(sri.quantity) AS total_qty,
                   SUM(sri.free_quantity) AS free_qty,
                   SUM(sri.product_value) AS total_value
            FROM `tabScheme Request` sr
            INNER JOIN `tabScheme Request Item` sri ON sri.parent = sr.name
            INNER JOIN `tabDoctor Master` dm ON dm.name = sr.doctor_code
            LEFT JOIN `tabHQ Master` hm ON hm.name = dm.hq
            WHERE {where}
            GROUP BY sr.name, dm.doctor_name, dm.hq, dm.region
            ORDER BY total_value DESC
        """, params, as_dict=True)
    else:
        stockist_join = ""
        if group_by == "Stockist":
            stockist_join = "LEFT JOIN `tabStockist Master` sm ON sm.name = sr.stockist_code"

        rows = frappe.db.sql(f"""
            SELECT {extra_select},
                   SUM(sri.quantity) AS total_qty,
                   SUM(sri.free_quantity) AS free_qty,
                   SUM(sri.product_value) AS total_value
            FROM `tabScheme Request` sr
            INNER JOIN `tabScheme Request Item` sri ON sri.parent = sr.name
            INNER JOIN `tabDoctor Master` dm ON dm.name = sr.doctor_code
            LEFT JOIN `tabHQ Master` hm ON hm.name = dm.hq
            {stockist_join}
            WHERE {where}
            GROUP BY {group_clause}
            ORDER BY {group_clause}
        """, params, as_dict=True)

    data = []
    for r in rows:
        entry = {
            "group_key": r.group_key or "",
            "group_label": r.group_label or "",
            "total_qty": flt(r.total_qty),
            "free_qty": flt(r.free_qty),
            "total_value": flt(r.total_value),
        }
        if hasattr(r, "hq"):
            entry["hq"] = r.hq or ""
        if hasattr(r, "region"):
            entry["region"] = r.region or ""
        data.append(entry)

    return {"success": True, "data": data, "group_by": group_by}


@frappe.whitelist()
def export_scheme_report_excel(report_type, division=None, **kwargs):
    """Generate a styled Excel workbook for scheme reports."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    if not division:
        division = get_user_division()

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_font = Font(bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    group_font = Font(bold=True, size=11)
    group_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")

    wb = openpyxl.Workbook()
    ws = wb.active

    def write_header_row(ws, row_num, headers):
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=row_num, column=col, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border

    def write_data_row(ws, row_num, values):
        for col, v in enumerate(values, 1):
            cell = ws.cell(row=row_num, column=col, value=v)
            cell.border = thin_border

    def write_group_row(ws, row_num, label, num_cols):
        cell = ws.cell(row=row_num, column=1, value=label)
        cell.font = group_font
        cell.fill = group_fill
        for c in range(1, num_cols + 1):
            ws.cell(row=row_num, column=c).fill = group_fill
            ws.cell(row=row_num, column=c).border = thin_border

    def write_title_rows(ws, title, subtitle=""):
        ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=14)
        if subtitle:
            ws.cell(row=2, column=1, value=subtitle).font = Font(size=11, italic=True)
        return 4

    from_date = kwargs.get("from_date", "")
    to_date = kwargs.get("to_date", "")
    region_val = kwargs.get("region", "")
    zone_val = kwargs.get("zone", "")
    doctor_status = kwargs.get("doctor_status", "Active")
    product_type = kwargs.get("product_type", "All")
    product_codes = kwargs.get("product_codes", None)
    criteria = kwargs.get("criteria", "Region")
    team_or_hq = kwargs.get("team_or_hq", "")
    group_by = kwargs.get("group_by", "HQ")
    period_label = f"{from_date} to {to_date}" if from_date and to_date else ""
    region_label = region_val or "All Regions"

    ml = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]

    if report_type == "activity_trend":
        ws.title = "Activity Trend"
        result = get_scheme_activity_trend_report(
            division, from_date, to_date, doctor_status, product_type,
            product_codes, zone_val, region_val, criteria, team_or_hq
        )
        data = result.get("data", [])
        fy_label = result.get("fy_label", "")
        row = write_title_rows(ws, f"Activity Trend Report – {division}",
                               f"Region: {region_label}  |  {fy_label}")
        headers = ["Region", "HQ", "Doctor Name"] + ml
        write_header_row(ws, row, headers)
        row += 1
        current_region = None
        for d in data:
            if d["region"] != current_region:
                current_region = d["region"]
                write_group_row(ws, row, f"{current_region} Region", len(headers))
                row += 1
            vals = [d["region"], d.get("hq_name") or d["hq"], d["doctor_name"]] + d["months"]
            write_data_row(ws, row, vals)
            row += 1

    elif report_type == "activity_track":
        ws.title = "Activity Track"
        result = get_scheme_activity_track_report(
            division, from_date, to_date, doctor_status, product_type,
            product_codes, zone_val, region_val, criteria, team_or_hq
        )
        data = result.get("data", [])
        totals = result.get("totals", {})
        row = write_title_rows(ws, f"Activity Track Report – {division}",
                               f"Region: {region_label}  |  Period: {period_label}")
        headers = ["S.No", "Date", "Region", "HQ", "Doctor Name", "Product",
                    "Qty", "Free Qty", "Rate", "Value", "Stockist"]
        write_header_row(ws, row, headers)
        row += 1
        for d in data:
            write_data_row(ws, row, [d["sno"], d["date"], d["region"], d["hq"],
                                     d["doctor_name"], d["product_code"],
                                     d["qty"], d["free_qty"], d["rate"], d["value"],
                                     d["stockist_name"]])
            row += 1
        # Totals row
        write_group_row(ws, row, "Total", len(headers))
        ws.cell(row=row, column=7, value=totals.get("qty", 0)).font = group_font
        ws.cell(row=row, column=8, value=totals.get("free_qty", 0)).font = group_font
        ws.cell(row=row, column=10, value=totals.get("value", 0)).font = group_font

    elif report_type == "new_approval_doctors":
        ws.title = "New Approval Doctors"
        result = get_new_approval_doctors_report(
            division, from_date, to_date, product_codes, zone_val, region_val,
            criteria, team_or_hq
        )
        data = result.get("data", [])
        row = write_title_rows(ws, f"New Approval Doctors – {division}",
                               f"Period: {period_label}")
        headers = ["S.No", "Approval Date", "Region", "HQ", "Doctor Name",
                    "Hospital", "City", "Approved By"]
        write_header_row(ws, row, headers)
        row += 1
        for d in data:
            write_data_row(ws, row, [d["sno"], d["approval_date"], d["region"],
                                     d["hq"], d["doctor_name"], d["hospital"],
                                     d["city"], d["approved_by"]])
            row += 1

    elif report_type == "periodic":
        ws.title = "Periodic Report"
        result = get_scheme_periodic_report(
            division, from_date, to_date, group_by, zone_val, region_val,
            product_codes
        )
        data = result.get("data", [])
        gb = result.get("group_by", "HQ")
        row = write_title_rows(ws, f"Periodic Report ({gb} Wise) – {division}",
                               f"Period: {period_label}")
        headers = [gb, "Total Qty", "Free Qty", "Total Value"]
        write_header_row(ws, row, headers)
        row += 1
        for d in data:
            write_data_row(ws, row, [d["group_label"], d["total_qty"],
                                     d["free_qty"], d["total_value"]])
            row += 1
    else:
        frappe.throw("Invalid report type")

    # Auto-fit column widths
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_len + 4, 50)

    from io import BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    xlsx_data = output.getvalue()

    filename = f"Scheme_Report_{report_type}_{division}.xlsx"
    frappe.local.response.filename = filename
    frappe.local.response.filecontent = xlsx_data
    frappe.local.response.type = "download"
    frappe.local.response.content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ═══════════════════════════════════════════════════════════════
# RANKING REPORTS  –  Portal API Methods
# ═══════════════════════════════════════════════════════════════

_MONTH_LABELS = ["Apr", "May", "Jun", "Jul", "Aug", "Sep",
                 "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]
_MONTH_MAP = {4: 0, 5: 1, 6: 2, 7: 3, 8: 4, 9: 5,
              10: 6, 11: 7, 12: 8, 1: 9, 2: 10, 3: 11}


def _current_fy():
    """Return (fy_start_date, fy_end_date, fy_label) for the current financial year."""
    from datetime import date
    today = date.today()
    if today.month >= 4:
        fy_start = date(today.year, 4, 1)
        fy_end = date(today.year + 1, 3, 31)
    else:
        fy_start = date(today.year - 1, 4, 1)
        fy_end = date(today.year, 3, 31)
    fy_label = f"Apr {str(fy_start.year)[2:]} to Mar {str(fy_end.year)[2:]}"
    return str(fy_start), str(fy_end), fy_label


@frappe.whitelist()
def get_ranking_report_filter_options(division=None):
    """Return dropdown options for the Ranking Reports portal page (active masters only)."""
    if not division:
        division = get_user_division()

    zones = frappe.db.sql(
        "SELECT name FROM `tabZone Master` WHERE status='Active' AND division IN (%s, 'Both') ORDER BY name",
        (division,), as_list=1)
    regions = frappe.db.sql(
        "SELECT name, region_name, zone FROM `tabRegion Master` WHERE status='Active' AND division IN (%s, 'Both') ORDER BY name",
        (division,), as_dict=True)
    teams = frappe.db.sql(
        "SELECT name, region FROM `tabTeam Master` WHERE status='Active' AND division IN (%s, 'Both') ORDER BY name",
        (division,), as_dict=True)
    hqs = frappe.db.sql(
        "SELECT name, hq_name, team FROM `tabHQ Master` WHERE division=%s AND status='Active' ORDER BY name",
        (division,), as_dict=True)
    stockists = frappe.db.sql(
        "SELECT name, stockist_name, hq FROM `tabStockist Master` WHERE division=%s AND status='Active' ORDER BY stockist_name",
        (division,), as_dict=True)
    products = frappe.db.sql(
        "SELECT product_code, product_name, category, product_group, pack "
        "FROM `tabProduct Master` WHERE division=%s AND status='Active' ORDER BY product_name",
        (division,), as_dict=True)
    doctors = frappe.db.sql(
        "SELECT name, doctor_code, doctor_name, hq, region "
        "FROM `tabDoctor Master` WHERE division=%s AND status='Active' ORDER BY doctor_name",
        (division,), as_dict=True)

    return {
        "zones": [z[0] for z in zones],
        "regions": [{"name": r.name, "region_name": r.region_name or r.name, "zone": r.zone or ""} for r in regions],
        "teams": [{"name": t.name, "region": t.region or ""} for t in teams],
        "hqs": [{"name": h.name, "hq_name": h.hq_name or "", "team": h.team or ""} for h in hqs],
        "stockists": [{"code": s.name, "name": s.stockist_name, "hq": s.hq or ""} for s in stockists],
        "products": [{"code": p.product_code, "name": p.product_name,
                       "category": p.category or "", "group": p.product_group or "",
                       "pack": p.pack or ""} for p in products],
        "doctors": [{"code": d.name, "doctor_code": d.doctor_code or d.name,
                      "name": d.doctor_name, "hq": d.hq or "",
                      "region": d.region or ""} for d in doctors],
    }


# ─────────────────────────────────────────────────────────────
# Report 1: Moving Trend Report
# ─────────────────────────────────────────────────────────────
@frappe.whitelist()
def get_ranking_moving_trend_report(division=None, sales_type="secondary",
                                     criteria="Region", from_date=None, to_date=None,
                                     region=None, zone=None):
    """Monthly pivot (Apr–Mar) grouped by criteria (Region / HQ / Team / Doctor / Stockist / Product)."""
    if not division:
        division = get_user_division()

    fy_start, fy_end, fy_label = _current_fy()
    if not from_date:
        from_date = fy_start
    if not to_date:
        to_date = fy_end

    if sales_type == "primary":
        rows = _moving_trend_primary(division, criteria, from_date, to_date, region, zone)
    else:
        rows = _moving_trend_secondary(division, criteria, from_date, to_date, region, zone)

    # Build pivot
    pivoted = {}
    for r in rows:
        key = r.criteria_name
        if key not in pivoted:
            pivoted[key] = {"criteria_name": key, "months": [0] * 12, "total": 0}
        idx = _MONTH_MAP.get(r.m)
        if idx is not None:
            pivoted[key]["months"][idx] += flt(r.qty)
            pivoted[key]["total"] += flt(r.qty)

    data = list(pivoted.values())
    for d in data:
        d["average"] = round(d["total"] / 12, 2)
    data.sort(key=lambda x: x["total"], reverse=True)

    return {"success": True, "data": data, "fy_label": fy_label,
            "month_labels": _MONTH_LABELS, "criteria": criteria}


def _moving_trend_primary(division, criteria, from_date, to_date, region, zone):
    """Query Primary Sales Data for Moving Trend grouped by criteria."""
    criteria_col_map = {
        "Region": "ps.region",
        "HQ": "IFNULL(sm.hq, ps.team)",
        "Team": "ps.team",
        "Stockist": "CONCAT(ps.stockist_code, ' – ', ps.stockist_name)",
        "Product": "CONCAT(ps.pcode, ' – ', ps.product)",
        "Doctor": "ps.region",  # doctors not in primary sales; fallback to region
    }
    criteria_col = criteria_col_map.get(criteria, "ps.region")

    conditions = ["ps.division = %(division)s", "ps.iscancelled = 0",
                   "ps.invoicedate >= %(from_date)s", "ps.invoicedate <= %(to_date)s"]
    params = {"division": division, "from_date": from_date, "to_date": to_date}

    join_sm = ""
    if criteria == "HQ":
        join_sm = " LEFT JOIN `tabStockist Master` sm ON sm.name = ps.stockist_code AND sm.status = 'Active'"
    if region:
        conditions.append("ps.region = %(region)s")
        params["region"] = region
    if zone:
        conditions.append("ps.zonee = %(zone)s")
        params["zone"] = zone

    where = " AND ".join(conditions)
    return frappe.db.sql(f"""
        SELECT {criteria_col} AS criteria_name,
               MONTH(ps.invoicedate) AS m,
               SUM(ps.quantity) AS qty
        FROM `tabPrimary Sales Data` ps
        {join_sm}
        WHERE {where}
        GROUP BY criteria_name, MONTH(ps.invoicedate)
    """, params, as_dict=True)


def _moving_trend_secondary(division, criteria, from_date, to_date, region, zone):
    """Query Stockist Statement Items for Moving Trend grouped by criteria."""
    criteria_col_map = {
        "Region": "ss.region",
        "HQ": "IFNULL(sm.hq, ss.team)",
        "Team": "ss.team",
        "Stockist": "CONCAT(ss.stockist_code, ' – ', ss.stockist_name)",
        "Product": "CONCAT(si.product_code, ' – ', si.product_name)",
        "Doctor": "ss.region",  # doctors have no direct relation to statements; fallback
    }
    criteria_col = criteria_col_map.get(criteria, "ss.region")

    conditions = ["ss.division = %(division)s", "ss.docstatus IN (0, 1)",
                   "ss.statement_month >= %(from_date)s", "ss.statement_month <= %(to_date)s"]
    params = {"division": division, "from_date": from_date, "to_date": to_date}

    join_sm = ""
    if criteria == "HQ":
        join_sm = " LEFT JOIN `tabStockist Master` sm ON sm.name = ss.stockist_code AND sm.status = 'Active'"
    if region:
        conditions.append("ss.region = %(region)s")
        params["region"] = region
    if zone:
        conditions.append("ss.zone = %(zone)s")
        params["zone"] = zone

    where = " AND ".join(conditions)
    return frappe.db.sql(f"""
        SELECT {criteria_col} AS criteria_name,
               MONTH(ss.statement_month) AS m,
               SUM(si.sales_qty) AS qty
        FROM `tabStockist Statement` ss
        INNER JOIN `tabStockist Statement Item` si
            ON si.parent = ss.name AND si.parenttype = 'Stockist Statement'
        {join_sm}
        WHERE {where}
        GROUP BY criteria_name, MONTH(ss.statement_month)
    """, params, as_dict=True)


# ─────────────────────────────────────────────────────────────
# Report 2: Rupee Wise Report
# ─────────────────────────────────────────────────────────────
@frappe.whitelist()
def get_ranking_rupee_wise_report(division=None, sales_type="secondary",
                                   value_condition="gt", sale_value=0,
                                   from_date=None, to_date=None):
    """Transaction-level rows filtered by value > or < threshold."""
    if not division:
        division = get_user_division()
    sale_value = flt(sale_value)
    val_op = ">=" if value_condition == "gt" else "<="

    if sales_type == "primary":
        conditions = ["ps.division = %(division)s", "ps.iscancelled = 0",
                       f"ps.ptsvalue {val_op} %(sale_value)s"]
        params = {"division": division, "sale_value": sale_value}
        if from_date:
            conditions.append("ps.invoicedate >= %(from_date)s")
            params["from_date"] = from_date
        if to_date:
            conditions.append("ps.invoicedate <= %(to_date)s")
            params["to_date"] = to_date
        where = " AND ".join(conditions)

        rows = frappe.db.sql(f"""
            SELECT ps.invoicedate AS date, ps.region, ps.team AS hq,
                   ps.stockist_code, ps.stockist_name,
                   ps.pcode AS product_code, ps.product AS product_name,
                   ps.quantity AS qty, ps.pts AS rate, ps.ptsvalue AS value
            FROM `tabPrimary Sales Data` ps
            WHERE {where}
            ORDER BY ps.invoicedate, ps.stockist_code
            LIMIT 5000
        """, params, as_dict=True)
    else:
        conditions = ["ss.division = %(division)s", "ss.docstatus IN (0, 1)",
                       f"si.sales_value_pts {val_op} %(sale_value)s"]
        params = {"division": division, "sale_value": sale_value}
        if from_date:
            conditions.append("ss.statement_month >= %(from_date)s")
            params["from_date"] = from_date
        if to_date:
            conditions.append("ss.statement_month <= %(to_date)s")
            params["to_date"] = to_date
        where = " AND ".join(conditions)

        rows = frappe.db.sql(f"""
            SELECT ss.statement_month AS date, ss.region, ss.hq,
                   ss.stockist_code, ss.stockist_name,
                   si.product_code, si.product_name,
                   si.sales_qty AS qty, si.pts AS rate, si.sales_value_pts AS value
            FROM `tabStockist Statement` ss
            INNER JOIN `tabStockist Statement Item` si
                ON si.parent = ss.name AND si.parenttype = 'Stockist Statement'
            WHERE {where}
            ORDER BY ss.statement_month, ss.stockist_code
            LIMIT 5000
        """, params, as_dict=True)

    # Add serial number
    for i, r in enumerate(rows, 1):
        r["sno"] = i
        if r.get("date"):
            r["date"] = str(r["date"])

    return {"success": True, "data": rows}


# ─────────────────────────────────────────────────────────────
# Report 3: Productwise Ranking (Top N)
# ─────────────────────────────────────────────────────────────
@frappe.whitelist()
def get_ranking_productwise_topn(division=None, product_codes=None, top_n=5,
                                  from_date=None, to_date=None, sales_type="secondary"):
    """Rank selected products by total value, return top N with contribution %."""
    if not division:
        division = get_user_division()
    top_n = int(top_n or 5)

    codes = []
    if product_codes:
        codes = json.loads(product_codes) if isinstance(product_codes, str) else product_codes

    if sales_type == "primary":
        conditions = ["ps.division = %(division)s", "ps.iscancelled = 0"]
        params = {"division": division}
        if from_date:
            conditions.append("ps.invoicedate >= %(from_date)s")
            params["from_date"] = from_date
        if to_date:
            conditions.append("ps.invoicedate <= %(to_date)s")
            params["to_date"] = to_date
        if codes:
            conditions.append("ps.pcode IN %(codes)s")
            params["codes"] = codes
        where = " AND ".join(conditions)

        rows = frappe.db.sql(f"""
            SELECT ps.pcode AS product_code, ps.product AS product_name,
                   SUM(ps.quantity) AS total_qty, SUM(ps.ptsvalue) AS total_value
            FROM `tabPrimary Sales Data` ps
            WHERE {where}
            GROUP BY ps.pcode, ps.product
            ORDER BY total_value DESC
        """, params, as_dict=True)
    else:
        conditions = ["ss.division = %(division)s", "ss.docstatus IN (0, 1)"]
        params = {"division": division}
        if from_date:
            conditions.append("ss.statement_month >= %(from_date)s")
            params["from_date"] = from_date
        if to_date:
            conditions.append("ss.statement_month <= %(to_date)s")
            params["to_date"] = to_date
        if codes:
            conditions.append("si.product_code IN %(codes)s")
            params["codes"] = codes
        where = " AND ".join(conditions)

        rows = frappe.db.sql(f"""
            SELECT si.product_code, si.product_name,
                   SUM(si.sales_qty) AS total_qty,
                   SUM(si.sales_value_pts) AS total_value
            FROM `tabStockist Statement` ss
            INNER JOIN `tabStockist Statement Item` si
                ON si.parent = ss.name AND si.parenttype = 'Stockist Statement'
            WHERE {where}
            GROUP BY si.product_code, si.product_name
            ORDER BY total_value DESC
        """, params, as_dict=True)

    grand_total = sum(flt(r.total_value) for r in rows)
    data = []
    for rank, r in enumerate(rows[:top_n], 1):
        pct = round(flt(r.total_value) / grand_total * 100, 2) if grand_total else 0
        data.append({
            "rank": rank,
            "product_code": r.product_code,
            "product_name": r.product_name,
            "total_qty": flt(r.total_qty),
            "total_value": flt(r.total_value),
            "contribution_pct": pct,
        })

    return {"success": True, "data": data, "grand_total": grand_total}


# ─────────────────────────────────────────────────────────────
# Report 4: Productwise Ranking ALL (single product across HQs)
# ─────────────────────────────────────────────────────────────
@frappe.whitelist()
def get_ranking_productwise_all(division=None, product_code=None, region=None,
                                 sales_type="secondary", from_date=None, to_date=None):
    """Rank HQs for a single product by total value, with contribution %."""
    if not division:
        division = get_user_division()
    if not product_code:
        return {"success": False, "message": "Product code is required"}

    if sales_type == "primary":
        conditions = ["ps.division = %(division)s", "ps.iscancelled = 0",
                       "ps.pcode = %(product_code)s"]
        params = {"division": division, "product_code": product_code}
        if region:
            conditions.append("ps.region = %(region)s")
            params["region"] = region
        if from_date:
            conditions.append("ps.invoicedate >= %(from_date)s")
            params["from_date"] = from_date
        if to_date:
            conditions.append("ps.invoicedate <= %(to_date)s")
            params["to_date"] = to_date
        where = " AND ".join(conditions)

        rows = frappe.db.sql(f"""
            SELECT IFNULL(sm.hq, ps.team) AS hq,
                   ps.pcode AS product_code, ps.product AS product_name,
                   SUM(ps.quantity) AS total_qty, SUM(ps.ptsvalue) AS total_value
            FROM `tabPrimary Sales Data` ps
            LEFT JOIN `tabStockist Master` sm ON sm.name = ps.stockist_code AND sm.status = 'Active'
            WHERE {where}
            GROUP BY hq, ps.pcode, ps.product
            ORDER BY total_value DESC
        """, params, as_dict=True)
    else:
        conditions = ["ss.division = %(division)s", "ss.docstatus IN (0, 1)",
                       "si.product_code = %(product_code)s"]
        params = {"division": division, "product_code": product_code}
        if region:
            conditions.append("ss.region = %(region)s")
            params["region"] = region
        if from_date:
            conditions.append("ss.statement_month >= %(from_date)s")
            params["from_date"] = from_date
        if to_date:
            conditions.append("ss.statement_month <= %(to_date)s")
            params["to_date"] = to_date
        where = " AND ".join(conditions)

        rows = frappe.db.sql(f"""
            SELECT IFNULL(sm.hq, ss.hq) AS hq,
                   si.product_code, si.product_name,
                   SUM(si.sales_qty) AS total_qty,
                   SUM(si.sales_value_pts) AS total_value
            FROM `tabStockist Statement` ss
            INNER JOIN `tabStockist Statement Item` si
                ON si.parent = ss.name AND si.parenttype = 'Stockist Statement'
            LEFT JOIN `tabStockist Master` sm ON sm.name = ss.stockist_code AND sm.status = 'Active'
            WHERE {where}
            GROUP BY hq, si.product_code, si.product_name
            ORDER BY total_value DESC
        """, params, as_dict=True)

    grand_total = sum(flt(r.total_value) for r in rows)
    data = []
    for rank, r in enumerate(rows, 1):
        pct = round(flt(r.total_value) / grand_total * 100, 2) if grand_total else 0
        data.append({
            "rank": rank,
            "hq": r.hq,
            "product_code": r.product_code,
            "product_name": r.product_name,
            "total_qty": flt(r.total_qty),
            "total_value": flt(r.total_value),
            "contribution_pct": pct,
        })

    return {"success": True, "data": data, "grand_total": grand_total}


# ─────────────────────────────────────────────────────────────
# Report 5: Productwise Ranking Advanced
# ─────────────────────────────────────────────────────────────
@frappe.whitelist()
def get_ranking_productwise_advanced(division=None, sales_type="secondary",
                                      qty_filter=0, region=None, hq_wise=0,
                                      product_codes=None,
                                      from_date=None, to_date=None):
    """Advanced product ranking: Primary/Secondary/Closing, qty >= threshold, group by HQ or Stockist."""
    if not division:
        division = get_user_division()
    qty_filter = flt(qty_filter)
    hq_wise = int(hq_wise or 0)

    codes = []
    if product_codes:
        codes = json.loads(product_codes) if isinstance(product_codes, str) else product_codes

    if sales_type == "closing":
        # Closing stock from statement items
        conditions = ["ss.division = %(division)s", "ss.docstatus IN (0, 1)"]
        params = {"division": division}
        if region:
            conditions.append("ss.region = %(region)s")
            params["region"] = region
        if from_date:
            conditions.append("ss.statement_month >= %(from_date)s")
            params["from_date"] = from_date
        if to_date:
            conditions.append("ss.statement_month <= %(to_date)s")
            params["to_date"] = to_date
        if codes:
            conditions.append("si.product_code IN %(codes)s")
            params["codes"] = codes
        where = " AND ".join(conditions)

        group_col = "ss.hq" if hq_wise else "CONCAT(ss.stockist_code, ' – ', ss.stockist_name)"
        group_label = "HQ" if hq_wise else "Stockist"

        rows = frappe.db.sql(f"""
            SELECT ss.region, {group_col} AS group_key,
                   si.product_code, si.product_name,
                   SUM(si.closing_qty) AS qty,
                   SUM(si.closing_value) AS value
            FROM `tabStockist Statement` ss
            INNER JOIN `tabStockist Statement Item` si
                ON si.parent = ss.name AND si.parenttype = 'Stockist Statement'
            WHERE {where}
            GROUP BY ss.region, group_key, si.product_code, si.product_name
            HAVING qty >= %(qty_filter)s
            ORDER BY value DESC
            LIMIT 5000
        """, dict(**params, qty_filter=qty_filter), as_dict=True)

    elif sales_type == "primary":
        conditions = ["ps.division = %(division)s", "ps.iscancelled = 0"]
        params = {"division": division}
        if region:
            conditions.append("ps.region = %(region)s")
            params["region"] = region
        if from_date:
            conditions.append("ps.invoicedate >= %(from_date)s")
            params["from_date"] = from_date
        if to_date:
            conditions.append("ps.invoicedate <= %(to_date)s")
            params["to_date"] = to_date
        if codes:
            conditions.append("ps.pcode IN %(codes)s")
            params["codes"] = codes
        where = " AND ".join(conditions)

        if hq_wise:
            group_col = "IFNULL(sm.hq, ps.team)"
            join_sm = " LEFT JOIN `tabStockist Master` sm ON sm.name = ps.stockist_code AND sm.status = 'Active'"
        else:
            group_col = "CONCAT(ps.stockist_code, ' – ', ps.stockist_name)"
            join_sm = ""
        group_label = "HQ" if hq_wise else "Stockist"

        rows = frappe.db.sql(f"""
            SELECT ps.region, {group_col} AS group_key,
                   ps.pcode AS product_code, ps.product AS product_name,
                   SUM(ps.quantity) AS qty, SUM(ps.ptsvalue) AS value
            FROM `tabPrimary Sales Data` ps
            {join_sm}
            WHERE {where}
            GROUP BY ps.region, group_key, ps.pcode, ps.product
            HAVING qty >= %(qty_filter)s
            ORDER BY value DESC
            LIMIT 5000
        """, dict(**params, qty_filter=qty_filter), as_dict=True)

    else:  # secondary
        conditions = ["ss.division = %(division)s", "ss.docstatus IN (0, 1)"]
        params = {"division": division}
        if region:
            conditions.append("ss.region = %(region)s")
            params["region"] = region
        if from_date:
            conditions.append("ss.statement_month >= %(from_date)s")
            params["from_date"] = from_date
        if to_date:
            conditions.append("ss.statement_month <= %(to_date)s")
            params["to_date"] = to_date
        if codes:
            conditions.append("si.product_code IN %(codes)s")
            params["codes"] = codes
        where = " AND ".join(conditions)

        group_col = "ss.hq" if hq_wise else "CONCAT(ss.stockist_code, ' – ', ss.stockist_name)"
        group_label = "HQ" if hq_wise else "Stockist"

        rows = frappe.db.sql(f"""
            SELECT ss.region, {group_col} AS group_key,
                   si.product_code, si.product_name,
                   SUM(si.sales_qty) AS qty,
                   SUM(si.sales_value_pts) AS value
            FROM `tabStockist Statement` ss
            INNER JOIN `tabStockist Statement Item` si
                ON si.parent = ss.name AND si.parenttype = 'Stockist Statement'
            WHERE {where}
            GROUP BY ss.region, group_key, si.product_code, si.product_name
            HAVING qty >= %(qty_filter)s
            ORDER BY value DESC
            LIMIT 5000
        """, dict(**params, qty_filter=qty_filter), as_dict=True)

    data = []
    for rank, r in enumerate(rows, 1):
        data.append({
            "rank": rank,
            "region": r.region or "",
            "group_key": r.group_key or "",
            "product_code": r.product_code,
            "product_name": r.product_name,
            "qty": flt(r.qty),
            "value": flt(r.value),
        })

    return {"success": True, "data": data, "group_label": group_label}


# ─────────────────────────────────────────────────────────────
# Report 6: Moving Trend PCPM Tracker
# ─────────────────────────────────────────────────────────────
@frappe.whitelist()
def get_ranking_pcpm_tracker(division=None, sales_type="secondary",
                              region=None, product_codes=None):
    """Monthly pivot (Apr–Mar) per product with PCPM = total / sanctioned_strength / active_months."""
    if not division:
        division = get_user_division()

    fy_start, fy_end, fy_label = _current_fy()

    codes = []
    if product_codes:
        codes = json.loads(product_codes) if isinstance(product_codes, str) else product_codes

    # Get sanctioned strength = sum of per_capita from active HQ Masters in the region
    strength_conditions = ["hm.division = %(division)s", "hm.status = 'Active'"]
    strength_params = {"division": division}
    if region:
        strength_conditions.append("""hm.team IN (
            SELECT tm.name FROM `tabTeam Master` tm
            WHERE tm.region = %(region)s AND tm.status = 'Active')""")
        strength_params["region"] = region
    strength_where = " AND ".join(strength_conditions)

    strength_row = frappe.db.sql(f"""
        SELECT IFNULL(SUM(hm.per_capita), 0) AS total_strength
        FROM `tabHQ Master` hm
        WHERE {strength_where}
    """, strength_params, as_dict=True)
    sanctioned_strength = flt(strength_row[0].total_strength) if strength_row else 0

    # Get sales data
    if sales_type == "primary":
        conditions = ["ps.division = %(division)s", "ps.iscancelled = 0",
                       "ps.invoicedate >= %(from_date)s", "ps.invoicedate <= %(to_date)s"]
        params = {"division": division, "from_date": fy_start, "to_date": fy_end}
        if region:
            conditions.append("ps.region = %(region)s")
            params["region"] = region
        if codes:
            conditions.append("ps.pcode IN %(codes)s")
            params["codes"] = codes
        where = " AND ".join(conditions)

        rows = frappe.db.sql(f"""
            SELECT ps.pcode AS product_code, ps.product AS product_name,
                   MONTH(ps.invoicedate) AS m,
                   SUM(ps.quantity) AS qty
            FROM `tabPrimary Sales Data` ps
            WHERE {where}
            GROUP BY ps.pcode, ps.product, MONTH(ps.invoicedate)
        """, params, as_dict=True)
    else:
        conditions = ["ss.division = %(division)s", "ss.docstatus IN (0, 1)",
                       "ss.statement_month >= %(from_date)s", "ss.statement_month <= %(to_date)s"]
        params = {"division": division, "from_date": fy_start, "to_date": fy_end}
        if region:
            conditions.append("ss.region = %(region)s")
            params["region"] = region
        if codes:
            conditions.append("si.product_code IN %(codes)s")
            params["codes"] = codes
        where = " AND ".join(conditions)

        rows = frappe.db.sql(f"""
            SELECT si.product_code, si.product_name,
                   MONTH(ss.statement_month) AS m,
                   SUM(si.sales_qty) AS qty
            FROM `tabStockist Statement` ss
            INNER JOIN `tabStockist Statement Item` si
                ON si.parent = ss.name AND si.parenttype = 'Stockist Statement'
            WHERE {where}
            GROUP BY si.product_code, si.product_name, MONTH(ss.statement_month)
        """, params, as_dict=True)

    products = {}
    active_months = set()
    for r in rows:
        key = r.product_code
        if key not in products:
            products[key] = {
                "product_code": r.product_code,
                "product_name": r.product_name,
                "months": [0] * 12,
                "total": 0,
            }
        idx = _MONTH_MAP.get(r.m)
        if idx is not None:
            products[key]["months"][idx] += flt(r.qty)
            products[key]["total"] += flt(r.qty)
            active_months.add(idx)

    num_active_months = len(active_months) or 1
    data = []
    for p in products.values():
        p["average"] = round(p["total"] / 12, 2)
        p["pcpm"] = round(p["total"] / sanctioned_strength / num_active_months, 2) if sanctioned_strength else 0
        data.append(p)

    data.sort(key=lambda x: x["total"], reverse=True)

    return {"success": True, "data": data, "fy_label": fy_label,
            "month_labels": _MONTH_LABELS, "sanctioned_strength": sanctioned_strength,
            "active_months": num_active_months}


# ─────────────────────────────────────────────────────────────
# Ranking Reports – Excel Export
# ─────────────────────────────────────────────────────────────
@frappe.whitelist()
def export_ranking_report_excel(report_type, division=None, **kwargs):
    """Server-side Excel export for all 6 ranking report types."""
    if not division:
        division = get_user_division()

    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    wb = openpyxl.Workbook()
    ws = wb.active

    header_fill = PatternFill(start_color="1e293b", end_color="1e293b", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=10)
    group_fill = PatternFill(start_color="e8edf3", end_color="e8edf3", fill_type="solid")
    group_font = Font(bold=True, size=10)
    thin_border = Border(
        left=Side(style="thin", color="d1d5db"),
        right=Side(style="thin", color="d1d5db"),
        top=Side(style="thin", color="d1d5db"),
        bottom=Side(style="thin", color="d1d5db"),
    )

    def write_header_row(ws, row_num, headers):
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=row_num, column=col, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

    def write_data_row(ws, row_num, values):
        for col, v in enumerate(values, 1):
            cell = ws.cell(row=row_num, column=col, value=v)
            cell.border = thin_border

    def write_title_rows(ws, title, subtitle=""):
        ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=14)
        if subtitle:
            ws.cell(row=2, column=1, value=subtitle).font = Font(size=11, italic=True)
        return 4

    ml = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]

    if report_type == "moving_trend":
        ws.title = "Moving Trend"
        result = get_ranking_moving_trend_report(
            division, kwargs.get("sales_type", "secondary"),
            kwargs.get("criteria", "Region"),
            kwargs.get("from_date"), kwargs.get("to_date"),
            kwargs.get("region"), kwargs.get("zone"))
        data = result.get("data", [])
        criteria_label = result.get("criteria", "Region")
        row = write_title_rows(ws, f"Moving Trend Report – {division}",
                               result.get("fy_label", ""))
        headers = [criteria_label] + ml + ["Total", "Average"]
        write_header_row(ws, row, headers)
        row += 1
        for d in data:
            vals = [d["criteria_name"]] + d["months"] + [d["total"], d["average"]]
            write_data_row(ws, row, vals)
            row += 1

    elif report_type == "rupee_wise":
        ws.title = "Rupee Wise"
        result = get_ranking_rupee_wise_report(
            division, kwargs.get("sales_type", "secondary"),
            kwargs.get("value_condition", "gt"),
            kwargs.get("sale_value", 0),
            kwargs.get("from_date"), kwargs.get("to_date"))
        data = result.get("data", [])
        from_d = kwargs.get("from_date", "")
        to_d = kwargs.get("to_date", "")
        period_label = f"{from_d} to {to_d}" if from_d and to_d else ""
        row = write_title_rows(ws, f"Rupee Wise Report – {division}", period_label)
        headers = ["S.No", "Date", "Region", "HQ", "Stockist Code", "Stockist Name",
                    "Product Code", "Product", "Qty", "Rate", "Value"]
        write_header_row(ws, row, headers)
        row += 1
        for d in data:
            write_data_row(ws, row, [d["sno"], d.get("date", ""), d.get("region", ""),
                                      d.get("hq", ""), d.get("stockist_code", ""),
                                      d.get("stockist_name", ""), d.get("product_code", ""),
                                      d.get("product_name", ""), d.get("qty", 0),
                                      d.get("rate", 0), d.get("value", 0)])
            row += 1

    elif report_type == "productwise_topn":
        ws.title = "Product Ranking Top N"
        result = get_ranking_productwise_topn(
            division, kwargs.get("product_codes"),
            kwargs.get("top_n", 5),
            kwargs.get("from_date"), kwargs.get("to_date"),
            kwargs.get("sales_type", "secondary"))
        data = result.get("data", [])
        row = write_title_rows(ws, f"Productwise Ranking (Top N) – {division}", "")
        headers = ["Rank", "Product Code", "Product Name", "Total Qty",
                    "Total Value", "Contribution %"]
        write_header_row(ws, row, headers)
        row += 1
        for d in data:
            write_data_row(ws, row, [d["rank"], d["product_code"], d["product_name"],
                                      d["total_qty"], d["total_value"], d["contribution_pct"]])
            row += 1

    elif report_type == "productwise_all":
        ws.title = "Product Ranking All"
        result = get_ranking_productwise_all(
            division, kwargs.get("product_code"),
            kwargs.get("region"), kwargs.get("sales_type", "secondary"),
            kwargs.get("from_date"), kwargs.get("to_date"))
        data = result.get("data", [])
        row = write_title_rows(ws, f"Productwise Ranking ALL – {division}", "")
        headers = ["Rank", "HQ", "Product Code", "Product Name",
                    "Total Qty", "Total Value", "Contribution %"]
        write_header_row(ws, row, headers)
        row += 1
        for d in data:
            write_data_row(ws, row, [d["rank"], d["hq"], d["product_code"],
                                      d["product_name"], d["total_qty"],
                                      d["total_value"], d["contribution_pct"]])
            row += 1

    elif report_type == "productwise_advanced":
        ws.title = "Product Ranking Advanced"
        result = get_ranking_productwise_advanced(
            division, kwargs.get("sales_type", "secondary"),
            kwargs.get("qty_filter", 0), kwargs.get("region"),
            kwargs.get("hq_wise", 0), kwargs.get("product_codes"),
            kwargs.get("from_date"), kwargs.get("to_date"))
        data = result.get("data", [])
        gl = result.get("group_label", "HQ")
        row = write_title_rows(ws, f"Productwise Ranking Advanced – {division}", "")
        headers = ["Rank", "Region", gl, "Product Code", "Product Name", "Qty", "Value"]
        write_header_row(ws, row, headers)
        row += 1
        for d in data:
            write_data_row(ws, row, [d["rank"], d["region"], d["group_key"],
                                      d["product_code"], d["product_name"],
                                      d["qty"], d["value"]])
            row += 1

    elif report_type == "pcpm_tracker":
        ws.title = "PCPM Tracker"
        result = get_ranking_pcpm_tracker(
            division, kwargs.get("sales_type", "secondary"),
            kwargs.get("region"), kwargs.get("product_codes"))
        data = result.get("data", [])
        row = write_title_rows(ws, f"PCPM Tracker – {division}",
                               f"{result.get('fy_label', '')}  |  Sanctioned Strength: {result.get('sanctioned_strength', 0)}")
        headers = ["Product Code", "Product Name"] + ml + ["Total", "Average", "PCPM"]
        write_header_row(ws, row, headers)
        row += 1
        for d in data:
            vals = [d["product_code"], d["product_name"]] + d["months"] + [d["total"], d["average"], d["pcpm"]]
            write_data_row(ws, row, vals)
            row += 1
    else:
        frappe.throw("Invalid report type")

    # Auto-fit column widths
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_len + 4, 50)

    from io import BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    xlsx_data = output.getvalue()

    filename = f"Ranking_Report_{report_type}_{division}.xlsx"
    frappe.local.response.filename = filename
    frappe.local.response.filecontent = xlsx_data
    frappe.local.response.type = "download"
    frappe.local.response.content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ═══════════════════════════════════════════════════════════════
# SECONDARY SALES MOVING TREND REPORT  –  Portal API
# ═══════════════════════════════════════════════════════════════

@frappe.whitelist()
def get_secondary_sales_moving_trend(division=None, entity_type="Team",
                                     entity_name=None, financial_year=None):
    """Product-wise monthly secondary sales pivot grouped by product category.

    Shows MAIN PRODUCTS, HOSPITAL PRODUCTS, NEW PRODUCTS sections with
    Value in Lakhs, Target (from HQ Yearly Target), Sanctioned Strength,
    Per Capita and achievement %.
    """
    if not division:
        division = get_user_division()
    if not entity_name:
        return {"success": False, "message": f"{entity_type} is required"}

    from datetime import date
    today = date.today()

    # Determine financial year
    if financial_year:
        try:
            start_year = int(financial_year.split("-")[0])
        except Exception:
            start_year = today.year if today.month >= 4 else today.year - 1
    else:
        start_year = today.year if today.month >= 4 else today.year - 1
        financial_year = f"{start_year}-{str(start_year + 1)[2:]}"

    fy_start = f"{start_year}-04-01"
    fy_end = f"{start_year + 1}-03-31"
    fy_label = f"Apr {str(start_year)[2:]} to Mar {str(start_year + 1)[2:]}"

    # ── Resolve entity → list of HQs ──
    hq_list = []
    entity_display = entity_name
    sanctioned_strength = 0

    if entity_type == "HQ":
        hq_list = [entity_name]
        pc = frappe.db.get_value("HQ Master", entity_name, "per_capita") or 0
        sanctioned_strength = flt(pc)
        hq_name = frappe.db.get_value("HQ Master", entity_name, "hq_name") or entity_name
        entity_display = hq_name

    elif entity_type == "Team":
        hqs = frappe.get_all("HQ Master",
                             filters={"team": entity_name, "status": "Active", "division": division},
                             fields=["name", "hq_name"])
        hq_list = [h.name for h in hqs]
        hq_names = [h.hq_name for h in hqs]
        ss = frappe.db.get_value("Team Master", entity_name, "sanctioned_strength") or 0
        sanctioned_strength = flt(ss)
        team_name = frappe.db.get_value("Team Master", entity_name, "team_name") or entity_name
        entity_display = f"{team_name} ({', '.join(hq_names)})" if hq_names else team_name

    elif entity_type == "Region":
        teams = frappe.get_all("Team Master",
                               filters={"region": entity_name, "status": "Active",
                                         "division": ["in", [division, "Both"]]},
                               fields=["name"])
        for t in teams:
            sub_hqs = frappe.get_all("HQ Master",
                                 filters={"team": t.name, "status": "Active", "division": division},
                                 fields=["name"])
            hq_list.extend([h.name for h in sub_hqs])
        total_pc = frappe.db.sql(
            "SELECT COALESCE(SUM(per_capita), 0) FROM `tabHQ Master` "
            "WHERE team IN (SELECT name FROM `tabTeam Master` WHERE region=%s AND status='Active') "
            "AND status='Active' AND division=%s",
            (entity_name, division)
        )
        sanctioned_strength = flt(total_pc[0][0]) if total_pc else 0
        region_name = frappe.db.get_value("Region Master", entity_name, "region_name") or entity_name
        entity_display = region_name

    elif entity_type == "Zone":
        regions = frappe.get_all("Region Master",
                                  filters={"zone": entity_name, "status": "Active",
                                            "division": ["in", [division, "Both"]]},
                                  fields=["name"])
        for reg in regions:
            sub_teams = frappe.get_all("Team Master",
                                    filters={"region": reg.name, "status": "Active",
                                              "division": ["in", [division, "Both"]]},
                                    fields=["name"])
            for t in sub_teams:
                sub_hqs = frappe.get_all("HQ Master",
                                     filters={"team": t.name, "status": "Active", "division": division},
                                     fields=["name"])
                hq_list.extend([h.name for h in sub_hqs])
        total_pc = frappe.db.sql(
            "SELECT COALESCE(SUM(hm.per_capita), 0) FROM `tabHQ Master` hm "
            "INNER JOIN `tabTeam Master` tm ON hm.team = tm.name "
            "INNER JOIN `tabRegion Master` rm ON tm.region = rm.name "
            "WHERE rm.zone=%s AND hm.status='Active' AND hm.division=%s",
            (entity_name, division)
        )
        sanctioned_strength = flt(total_pc[0][0]) if total_pc else 0
        zone_name = frappe.db.get_value("Zone Master", entity_name, "zone_name") or entity_name
        entity_display = zone_name

    if not hq_list:
        return {"success": True, "sections": [],
                "entity_display": entity_display, "entity_type": entity_type,
                "sanctioned_strength": sanctioned_strength, "fy_label": fy_label}

    # ── Get all products for this division ──
    products = frappe.db.sql(
        "SELECT product_code, product_name, pack, category, pts "
        "FROM `tabProduct Master` WHERE division=%s AND status='Active' "
        "ORDER BY category, product_code",
        (division,), as_dict=True
    )

    # ── Get secondary sales (Stockist Statement Items) ──
    hq_placeholders = ", ".join(["%s"] * len(hq_list))
    sec_rows = frappe.db.sql(f"""
        SELECT si.product_code, si.pack,
               MONTH(ss.statement_month) AS m,
               SUM(si.sales_qty) AS qty,
               SUM(IFNULL(si.sales_value_pts, si.sales_qty * IFNULL(si.pts, 0))) AS value
        FROM `tabStockist Statement` ss
        INNER JOIN `tabStockist Statement Item` si
            ON si.parent = ss.name AND si.parenttype = 'Stockist Statement'
        WHERE ss.division = %s AND ss.docstatus IN (0, 1)
              AND ss.hq IN ({hq_placeholders})
              AND ss.statement_month BETWEEN %s AND %s
        GROUP BY si.product_code, si.pack, MONTH(ss.statement_month)
    """, [division] + hq_list + [fy_start, fy_end], as_dict=True)

    # ── Pivot data: product_code → 12-month array ──
    month_map = {4: 0, 5: 1, 6: 2, 7: 3, 8: 4, 9: 5,
                 10: 6, 11: 7, 12: 8, 1: 9, 2: 10, 3: 11}
    product_data = {}
    for r in sec_rows:
        pc = r.product_code
        if pc not in product_data:
            product_data[pc] = {"months_qty": [0] * 12, "months_val": [0.0] * 12}
        idx = month_map.get(r.m)
        if idx is not None:
            product_data[pc]["months_qty"][idx] += flt(r.qty)
            product_data[pc]["months_val"][idx] += flt(r.value)

    # ── Get HQ target value (in Lakhs) for the FY ──
    target_value = 0.0
    target_rows = frappe.db.sql("""
        SELECT COALESCE(SUM(ti.yearly_total), 0) AS total_target
        FROM `tabHQ Yearly Target` yt
        INNER JOIN `tabHQ Target Item` ti ON ti.parent = yt.name
        WHERE yt.docstatus = 1 AND yt.division = %s AND yt.financial_year = %s
              AND ti.hq IN ({hq_ph})
    """.format(hq_ph=hq_placeholders), [division, financial_year] + hq_list, as_dict=True)
    if target_rows:
        target_value = flt(target_rows[0].total_target)

    # ── Build sections by category ──
    category_order = [
        ("Main Products", "MAIN PRODUCTS"),
        ("Hospital Products", "HOSPITAL PRODUCTS"),
        ("New Products", "NEW PRODUCTS"),
    ]

    # Count months with actual data
    active_months = 0
    for i in range(12):
        has_data = any(pd["months_qty"][i] > 0 for pd in product_data.values())
        if has_data:
            active_months += 1
    active_months = max(active_months, 1)

    sections = []
    grand_total_val = 0.0

    # First pass: compute totals per section
    section_totals = {}
    for cat_key, cat_label in category_order:
        cat_products = [p for p in products if (p.category or "").lower() == cat_key.lower()]
        total_val = sum(sum(product_data.get(p.product_code, {"months_val": [0.0] * 12})["months_val"])
                        for p in cat_products)
        section_totals[cat_key] = total_val
        grand_total_val += total_val

    for cat_key, cat_label in category_order:
        cat_products = [p for p in products if (p.category or "").lower() == cat_key.lower()]
        if not cat_products:
            continue

        section_months_val = [0.0] * 12
        product_rows = []

        for p in cat_products:
            pc = p.product_code
            pd = product_data.get(pc, {"months_qty": [0] * 12, "months_val": [0.0] * 12})

            total_qty = sum(pd["months_qty"])
            avg_qty = round(total_qty / active_months) if total_qty else 0
            per_capita = round(avg_qty / sanctioned_strength) if sanctioned_strength else 0

            for i in range(12):
                section_months_val[i] += pd["months_val"][i]

            product_rows.append({
                "code": pc,
                "pack": p.pack or "",
                "target": 0,
                "months": pd["months_qty"],
                "total": total_qty,
                "average": avg_qty,
                "per_capita": per_capita,
            })

        # Section value in lakhs
        section_total_val = section_totals[cat_key]
        section_months_lakhs = [round(v / 100000, 2) for v in section_months_val]
        section_total_lakhs = round(section_total_val / 100000, 2)
        section_avg_lakhs = round(section_total_lakhs / active_months, 2)
        section_pc_lakhs = round(section_avg_lakhs / sanctioned_strength, 2) if sanctioned_strength else 0

        # Proportional target split
        cat_target = 0.0
        cat_pct = 0.0
        if target_value > 0 and grand_total_val > 0:
            share = section_total_val / grand_total_val if grand_total_val else 0
            cat_target = round(target_value * share, 2)
            cat_pct = round((section_total_lakhs / cat_target) * 100) if cat_target else 0

        sections.append({
            "label": cat_label,
            "category": cat_key,
            "products": product_rows,
            "value_in_lakhs": {
                "target": cat_target,
                "months": section_months_lakhs,
                "total": section_total_lakhs,
                "average": section_avg_lakhs,
                "per_capita": section_pc_lakhs,
                "pct": cat_pct,
            }
        })

    return {
        "success": True,
        "entity_type": entity_type,
        "entity_display": entity_display,
        "sanctioned_strength": sanctioned_strength,
        "fy_label": fy_label,
        "financial_year": financial_year,
        "active_months": active_months,
        "sections": sections,
        "month_labels": _MONTH_LABELS,
    }


# ═══════════════════════════════════════════════════════════════
# REGION-WISE STOCKIST SECONDARY SALES MOVING TREND
# Format: Rows = Stockists; Cols = Opening | Apr..Mar | Total Value | Closing
# Values in ₹ Lakhs; grouped by Team within Region
# Includes Draft (docstatus IN (0,1)) statements
# ═══════════════════════════════════════════════════════════════

@frappe.whitelist()
def get_region_wise_stockist_moving_trend(division=None, region=None, financial_year=None):
    """
    Region-wise stockist sales moving trend with opening and closing values.
    Matches the Orissa Secondary Sales Excel format:
      Columns: Stockist | HQ | Current Opening | Apr..Mar (12 months) | Total Sales Value | Current Closing
    Values in ₹ Lakhs.  Grouped by Team within the Region.
    Includes both Draft and Submitted statements (docstatus IN (0,1)).
    """
    if not division:
        division = get_user_division()
    if not region:
        return {"success": False, "message": "Region is required"}

    from datetime import date
    today = date.today()

    # Determine financial year
    if financial_year:
        try:
            start_year = int(financial_year.split("-")[0])
        except Exception:
            start_year = today.year if today.month >= 4 else today.year - 1
    else:
        start_year = today.year if today.month >= 4 else today.year - 1
        financial_year = f"{start_year}-{str(start_year + 1)[2:]}"

    fy_start = f"{start_year}-04-01"
    fy_end   = f"{start_year + 1}-03-31"
    fy_label = f"Apr {str(start_year)[2:]} to Mar {str(start_year + 1)[2:]}"

    # ── Month headers (FY order Apr=0 … Mar=11) ──
    month_map = {4: 0, 5: 1, 6: 2, 7: 3, 8: 4, 9: 5,
                 10: 6, 11: 7, 12: 8, 1: 9, 2: 10, 3: 11}
    month_labels = [
        f"Apr-{str(start_year)[2:]}", f"May-{str(start_year)[2:]}",
        f"Jun-{str(start_year)[2:]}", f"Jul-{str(start_year)[2:]}",
        f"Aug-{str(start_year)[2:]}", f"Sep-{str(start_year)[2:]}",
        f"Oct-{str(start_year)[2:]}", f"Nov-{str(start_year)[2:]}",
        f"Dec-{str(start_year)[2:]}", f"Jan-{str(start_year + 1)[2:]}",
        f"Feb-{str(start_year + 1)[2:]}", f"Mar-{str(start_year + 1)[2:]}",
    ]

    # ── Get all active stockists in this region ──
    stockists = frappe.db.sql("""
        SELECT sm.name AS code, sm.stockist_name, sm.hq,
               COALESCE(hm.hq_name, sm.hq, '') AS hq_name,
               COALESCE(hm.team, '') AS team,
               COALESCE(tm.team_name, hm.team, '') AS team_name
        FROM `tabStockist Master` sm
        LEFT JOIN `tabHQ Master` hm ON hm.name = sm.hq AND hm.division = %(division)s
        LEFT JOIN `tabTeam Master` tm ON tm.name = hm.team AND tm.division IN (%(division)s, 'Both')
        WHERE sm.division = %(division)s
          AND sm.region = %(region)s
          AND sm.status = 'Active'
        ORDER BY hm.team, sm.stockist_name
    """, {"division": division, "region": region}, as_dict=True)

    if not stockists:
        return {"success": True, "data": [], "teams": [], "fy_label": fy_label,
                "month_labels": month_labels, "region": region}

    stockist_codes = [s.code for s in stockists]
    placeholders = ", ".join(["%s"] * len(stockist_codes))

    # ── Monthly secondary sales value (₹) per stockist ──
    sales_rows = frappe.db.sql(f"""
        SELECT ss.stockist_code,
               MONTH(ss.statement_month) AS m,
               SUM(IFNULL(si.sales_value_pts, si.sales_qty * IFNULL(si.pts, 0))) AS sales_value
        FROM `tabStockist Statement` ss
        INNER JOIN `tabStockist Statement Item` si
            ON si.parent = ss.name AND si.parenttype = 'Stockist Statement'
        WHERE ss.division = %s
          AND ss.docstatus IN (0, 1)
          AND ss.stockist_code IN ({placeholders})
          AND ss.statement_month BETWEEN %s AND %s
        GROUP BY ss.stockist_code, MONTH(ss.statement_month)
    """, [division] + stockist_codes + [fy_start, fy_end], as_dict=True)

    # ── Opening stock of the FIRST statement in the FY (total value across all products) ──
    opening_rows = frappe.db.sql(f"""
        SELECT ss.stockist_code,
               SUM(IFNULL(si.opening_value, 0)) AS opening_value
        FROM `tabStockist Statement` ss
        INNER JOIN `tabStockist Statement Item` si
            ON si.parent = ss.name AND si.parenttype = 'Stockist Statement'
        WHERE ss.division = %s
          AND ss.docstatus IN (0, 1)
          AND ss.stockist_code IN ({placeholders})
          AND ss.statement_month = (
              SELECT MIN(ss2.statement_month)
              FROM `tabStockist Statement` ss2
              WHERE ss2.stockist_code = ss.stockist_code
                AND ss2.division = %s
                AND ss2.docstatus IN (0, 1)
                AND ss2.statement_month BETWEEN %s AND %s
          )
        GROUP BY ss.stockist_code
    """, [division] + stockist_codes + [division, fy_start, fy_end], as_dict=True)

    # ── Closing stock of the LATEST statement in the FY ──
    closing_rows = frappe.db.sql(f"""
        SELECT ss.stockist_code,
               SUM(IFNULL(si.closing_value, 0)) AS closing_value
        FROM `tabStockist Statement` ss
        INNER JOIN `tabStockist Statement Item` si
            ON si.parent = ss.name AND si.parenttype = 'Stockist Statement'
        WHERE ss.division = %s
          AND ss.docstatus IN (0, 1)
          AND ss.stockist_code IN ({placeholders})
          AND ss.statement_month = (
              SELECT MAX(ss2.statement_month)
              FROM `tabStockist Statement` ss2
              WHERE ss2.stockist_code = ss.stockist_code
                AND ss2.division = %s
                AND ss2.docstatus IN (0, 1)
                AND ss2.statement_month BETWEEN %s AND %s
          )
        GROUP BY ss.stockist_code
    """, [division] + stockist_codes + [division, fy_start, fy_end], as_dict=True)

    # ── Index into dicts ──
    sales_by_stockist = {}
    for r in sales_rows:
        sales_by_stockist.setdefault(r.stockist_code, {})[r.m] = flt(r.sales_value)

    opening_by_stockist = {r.stockist_code: flt(r.opening_value) for r in opening_rows}
    closing_by_stockist = {r.stockist_code: flt(r.closing_value) for r in closing_rows}

    # ── Build team-grouped output ──
    LAKH = 100000.0

    def to_lakhs(v):
        return round(v / LAKH, 2) if v else 0.0

    teams_dict = {}
    for s in stockists:
        team_key = s.team or "Unassigned"
        team_display = s.team_name or team_key
        if team_key not in teams_dict:
            teams_dict[team_key] = {"team": team_key, "team_name": team_display, "stockists": []}

        monthly_values = [to_lakhs(sales_by_stockist.get(s.code, {}).get(m_num, 0))
                          for m_num in [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]]
        total_sales = round(sum(monthly_values), 2)
        opening = to_lakhs(opening_by_stockist.get(s.code, 0))
        closing = to_lakhs(closing_by_stockist.get(s.code, 0))

        teams_dict[team_key]["stockists"].append({
            "stockist_code": s.code,
            "stockist_name": s.stockist_name,
            "hq": s.hq_name or s.hq or "",
            "opening": opening,
            "months": monthly_values,
            "total_sales": total_sales,
            "closing": closing,
        })

    # ── Compute team totals ──
    teams_list = []
    grand_opening = grand_months = None
    grand_total = 0.0
    grand_closing = 0.0

    for team_key, td in teams_dict.items():
        t_opening = round(sum(s["opening"] for s in td["stockists"]), 2)
        t_months = [round(sum(s["months"][i] for s in td["stockists"]), 2) for i in range(12)]
        t_total = round(sum(t_months), 2)
        t_closing = round(sum(s["closing"] for s in td["stockists"]), 2)
        td["total"] = {"opening": t_opening, "months": t_months, "total_sales": t_total, "closing": t_closing}
        teams_list.append(td)

        grand_total += t_total
        grand_closing += t_closing
        if grand_months is None:
            grand_opening = t_opening
            grand_months = list(t_months)
        else:
            grand_opening = round(grand_opening + t_opening, 2)
            grand_months = [round(grand_months[i] + t_months[i], 2) for i in range(12)]

    region_name = frappe.db.get_value("Region Master", region, "region_name") or region

    return {
        "success": True,
        "region": region,
        "region_name": region_name,
        "division": division,
        "fy_label": fy_label,
        "financial_year": financial_year,
        "month_labels": month_labels,
        "teams": teams_list,
        "grand_total": {
            "opening": round(grand_opening or 0, 2),
            "months": grand_months or [0] * 12,
            "total_sales": round(grand_total, 2),
            "closing": round(grand_closing, 2),
        },
    }


@frappe.whitelist()
def render_spreadsheet_preview(file_url):
    """Render an XLS/XLSX/CSV/TXT file as an HTML table for QC viewer preview."""
    import pandas as pd
    import html as html_mod

    if not file_url or not isinstance(file_url, str):
        return {"html": ""}

    # Resolve to absolute file path on disk
    if file_url.startswith("/files/"):
        file_path = frappe.get_site_path("public", file_url.lstrip("/"))
    elif file_url.startswith("/private/files/"):
        file_path = frappe.get_site_path(file_url.lstrip("/"))
    else:
        file_path = frappe.get_site_path("public", file_url.lstrip("/"))

    if not os.path.isfile(file_path):
        return {"html": "<p class='text-muted text-center'>File not found</p>"}

    ext = os.path.splitext(file_path)[1].lower()

    try:
        if ext == ".csv":
            df = pd.read_csv(file_path, dtype=str, na_filter=False)
        elif ext in (".xls", ".xlsx"):
            df = pd.read_excel(file_path, dtype=str, na_filter=False)
        elif ext == ".txt":
            df = pd.read_csv(file_path, dtype=str, sep="\t", na_filter=False)
        else:
            return {"html": "<p class='text-muted text-center'>Unsupported file type</p>"}

        # Limit to first 500 rows for performance
        truncated = len(df) > 500
        df = df.head(500)

        # Build HTML table with escaped content
        rows_html = []
        # Header
        header_cells = "".join(
            f"<th>{html_mod.escape(str(c))}</th>" for c in df.columns
        )
        rows_html.append(f"<thead><tr>{header_cells}</tr></thead>")

        # Body
        rows_html.append("<tbody>")
        for _, row in df.iterrows():
            cells = "".join(
                f"<td>{html_mod.escape(str(v))}</td>" for v in row
            )
            rows_html.append(f"<tr>{cells}</tr>")
        rows_html.append("</tbody>")

        table_html = f"<table>{''.join(rows_html)}</table>"
        if truncated:
            table_html += "<p class='text-muted text-center mt-2'><small>Showing first 500 rows</small></p>"

        return {"html": table_html}

    except Exception as e:
        frappe.log_error(f"Spreadsheet preview error: {e}", "render_spreadsheet_preview")
        return {"html": f"<p class='text-muted text-center'>Could not parse file</p>"}


@frappe.whitelist()
def get_products_for_division(division=None):
    """Return all active products for a division for manual statement entry."""
    if not division:
        division = get_user_division()

    products = frappe.get_all(
        "Product Master",
        filters={"status": "Active", "division": division},
        fields=["product_code", "product_name", "pack", "pts", "ptr", "pack_conversion"],
        order_by="product_name asc",
        limit_page_length=0,
    )
    return products


@frappe.whitelist()
def create_manual_statement(stockist_code, statement_month, items, uploaded_file=None, remarks=None):
    """
    Create a Stockist Statement from manual entry (no AI extraction).
    items: JSON array of product rows with quantities.
    """
    try:
        if isinstance(items, str):
            items = json.loads(items)

        if not stockist_code or not statement_month:
            return {"success": False, "message": "Stockist code and statement month are required."}

        # Normalise month
        if len(statement_month) == 7:
            statement_month = statement_month + "-01"

        # Check for duplicates
        existing = frappe.db.exists("Stockist Statement", {
            "stockist_code": stockist_code,
            "statement_month": statement_month,
        })
        if existing:
            return {"success": False, "message": f"A statement already exists: {existing}"}

        doc = frappe.new_doc("Stockist Statement")
        doc.stockist_code = stockist_code
        doc.statement_month = statement_month
        doc.extracted_data_status = "Draft"
        if uploaded_file:
            doc.uploaded_file = uploaded_file
        if remarks:
            doc.remarks = remarks

        for row in items:
            product_code = row.get("productcode") or row.get("product_code")
            if not product_code:
                continue
            doc.append("items", {
                "product_code": product_code,
                "opening_qty": flt(row.get("openingqty") or 0),
                "purchase_qty": flt(row.get("purchaseqty") or 0),
                "sales_qty": flt(row.get("salesqty") or 0),
                "free_qty": flt(row.get("freeqty") or 0),
                "free_qty_scheme": flt(row.get("freeqtyscheme") or 0),
                "return_qty": flt(row.get("returnqty") or 0),
                "misc_out_qty": flt(row.get("miscoutqty") or 0),
                "closing_qty": flt(row.get("closingqty") or 0),
                "closing_value": flt(row.get("closingvalue") or 0),
            })

        if not doc.items:
            return {"success": False, "message": "No valid product rows provided."}

        doc.calculate_closing_and_totals()
        doc.save(ignore_permissions=True)
        frappe.db.commit()

        return {
            "success": True,
            "message": f"Statement saved as Draft with {len(doc.items)} items.",
            "doc_name": doc.name,
        }

    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Create Manual Statement Error")
        return {"success": False, "message": str(e)}


# ═══════════════════════════════════════════════════════════════
# PRODUCT CORRECTION / MAPPING ENDPOINTS
# ═══════════════════════════════════════════════════════════════

@frappe.whitelist()
def save_product_correction(stockist_code, raw_product_name, mapped_product_code, statement_name=None):
    """
    Create or update a Stockist Product Correction record.
    Called from the portal when QC maps an unmapped item.
    """
    try:
        if not stockist_code or not raw_product_name or not mapped_product_code:
            return {"success": False, "message": "stockist_code, raw_product_name, and mapped_product_code are required"}

        raw_name = raw_product_name.strip().upper()

        # Check if correction already exists
        existing = frappe.db.get_value(
            "Stockist Product Correction",
            {"stockist_code": stockist_code, "raw_product_name": raw_name},
            "name",
        )

        if existing:
            doc = frappe.get_doc("Stockist Product Correction", existing)
            doc.mapped_product_code = mapped_product_code
            doc.status = "Active"
            doc.save(ignore_permissions=True)
        else:
            doc = frappe.get_doc({
                "doctype": "Stockist Product Correction",
                "stockist_code": stockist_code,
                "raw_product_name": raw_name,
                "mapped_product_code": mapped_product_code,
                "division": frappe.db.get_value("Stockist Master", stockist_code, "division"),
                "created_from_statement": statement_name or "",
                "status": "Active",
            })
            doc.insert(ignore_permissions=True)

        frappe.db.commit()
        return {"success": True, "correction_name": doc.name, "message": "Correction saved"}

    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Save Product Correction Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def get_product_search_for_mapping(query, division=None):
    """
    Search Product Master for mapping dropdown in QC screen.
    Returns matching products by code or name.
    """
    try:
        query = (query or "").strip()
        if len(query) < 2:
            return {"success": True, "results": []}

        filters = [
            ["Product Master", "status", "=", "Active"],
        ]
        if division:
            filters.append(["Product Master", "division", "=", division])

        # Search by product_code or product_name
        or_filters = [
            ["Product Master", "name", "like", f"%{query}%"],
            ["Product Master", "product_name", "like", f"%{query}%"],
        ]

        results = frappe.get_all(
            "Product Master",
            filters=filters,
            or_filters=or_filters,
            fields=["name as product_code", "product_name", "pack", "pts", "division"],
            order_by="name asc",
            limit_page_length=20,
        )

        return {"success": True, "results": results}

    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Product Search For Mapping Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def apply_mapping_and_save_correction(doc_name, row_idx, mapped_product_code):
    """
    Map a single unmapped item in a statement to a product,
    save the correction for future auto-mapping, and recalculate.
    row_idx: 0-based index in the items table.
    """
    try:
        row_idx = int(row_idx)
        doc = frappe.get_doc("Stockist Statement", doc_name)

        if doc.docstatus == 1:
            return {"success": False, "message": "Cannot modify a submitted statement"}

        if row_idx < 0 or row_idx >= len(doc.items):
            return {"success": False, "message": f"Invalid row index: {row_idx}"}

        item = doc.items[row_idx]
        raw_name = (item.raw_product_name or "").strip().upper()

        # Verify the product exists
        if not frappe.db.exists("Product Master", mapped_product_code):
            return {"success": False, "message": f"Product '{mapped_product_code}' not found"}

        # Update the item
        item.product_code = mapped_product_code
        item.mapping_status = "matched"

        # Keep extraction confidence intact; mapping review is reflected via QC status.
        conf_values = [flt(i.row_confidence) for i in doc.items]
        doc.confidence_score = round(sum(conf_values) / len(conf_values), 1) if conf_values else 0

        # Recalculate
        doc.calculate_closing_and_totals()
        doc.calculate_qc_confidence()
        doc.save(ignore_permissions=True)
        frappe.db.commit()

        return {
            "success": True,
            "message": f"Mapped '{raw_name}' → {mapped_product_code}",
            "qc_confidence": doc.qc_confidence,
            "confidence_score": flt(doc.confidence_score),
        }

    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Apply Mapping Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def override_qc_confidence(doc_name, new_confidence):
    """
    Manually override QC confidence for a statement.
    Used when a user confirms auto-mapped items are correct.
    Only allows setting to 'All Matched' from 'Verification Needed'.
    """
    try:
        if new_confidence != "All Matched":
            return {"success": False, "message": "Only 'All Matched' override is supported"}

        doc = frappe.get_doc("Stockist Statement", doc_name)

        if doc.docstatus != 0:
            return {"success": False, "message": "Can only override QC on draft statements"}

        if doc.qc_confidence != "Verification Needed":
            return {"success": False, "message": "Override only available for 'Verification Needed' statements"}

        # Mark all auto_mapped items as matched
        for item in doc.items:
            if item.mapping_status == "auto_mapped":
                item.mapping_status = "matched"

        doc.qc_confidence = "All Matched"
        doc.save(ignore_permissions=True)
        frappe.db.commit()

        return {"success": True, "qc_confidence": doc.qc_confidence}

    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Override QC Confidence Error")
        return {"success": False, "message": str(e)}


@frappe.whitelist()
def get_bulk_jobs_list_enhanced(division=None, page=1, page_size=20):
    """
    Enhanced bulk jobs list with job_name and QC confidence summary.
    """
    try:
        if not division:
            division = get_user_division()

        page = int(page or 1)
        page_size = int(page_size or 20)
        offset = (page - 1) * page_size

        filters = {}
        if division:
            filters["division"] = division

        jobs = frappe.get_all(
            "Bulk Statement Upload",
            filters=filters,
            fields=[
                "name", "job_name", "statement_month", "division", "status",
                "progress", "total_files", "success_count", "failed_count",
                "skipped_count", "creation", "zip_file",
            ],
            order_by="creation desc",
            start=offset,
            page_length=page_size,
        )

        total_count = frappe.db.count("Bulk Statement Upload", filters=filters)

        # For each job, get QC confidence summary from its statements
        for job in jobs:
            job["statement_month"] = str(job.get("statement_month") or "")
            job["creation"] = str(job.get("creation") or "")
            log_data = frappe.db.get_value("Bulk Statement Upload", job["name"], "extraction_log")
            if log_data:
                try:
                    log_entries = json.loads(log_data)
                    statement_names = [e.get("statement") for e in log_entries if e.get("statement")]
                    if statement_names:
                        qc_counts = frappe.db.sql("""
                            SELECT qc_confidence, COUNT(*) as cnt
                            FROM `tabStockist Statement`
                            WHERE name IN %(names)s
                            GROUP BY qc_confidence
                        """, {"names": statement_names}, as_dict=True)
                        job["qc_summary"] = {r.qc_confidence: r.cnt for r in qc_counts}
                    else:
                        job["qc_summary"] = {}
                except (json.JSONDecodeError, TypeError):
                    job["qc_summary"] = {}
            else:
                job["qc_summary"] = {}

        return {
            "success": True,
            "jobs": jobs,
            "total": total_count,
            "page": page,
            "page_size": page_size,
        }

    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Get Bulk Jobs List Enhanced Error")
        return {"success": False, "message": str(e)}


# ============================================================
# AI CHATBOT - Text2SQL Agentic Pipeline
# ============================================================

SCANIFY_SCHEMA_PROMPT = """
You are Scanify AI Assistant — an expert analytical chatbot for Stedman Pharmaceuticals' Stockist & Scheme Management System.
You convert natural language questions into SQL queries against a MariaDB database, then format results clearly.

## COMPANY CONTEXT
- Company: Stedman Pharmaceuticals
- Application: Scanify — manages stockist statements, scheme requests, primary sales, and sales targets
- Divisions: Prima, Vektra (division-controlled, user sees data for their active division)
- Current user division filter: {division}

## ORGANIZATIONAL HIERARCHY (top → bottom)
Division → Zone → Region → Team → HQ → Stockist
- Each level links to its parent. Zone is top geographic grouping, HQ is the lowest field unit.
- Stockists belong to an HQ. Doctors are also mapped to HQ/Team/Region.

## DATABASE SCHEMA

### Master Tables

**`tabDivision`** — Division master
- name (PK), division_name

**`tabZone Master`** — Zone master
- name (PK, format Z####), zone_code, zone_name, division (FK→tabDivision), status

**`tabState Master`** — State master
- name (PK, format ST####), state_code, state_name, division (FK→tabDivision), status

**`tabRegion Master`** — Region master
- name (PK, format R####), region_code, region_name, division (FK→tabDivision), zone (FK→tabZone Master), state (FK→tabState Master), status

**`tabTeam Master`** — Team master
- name (PK, format T####), team_code, team_name, division (FK→tabDivision), region (FK→tabRegion Master), sanctioned_strength, status

**`tabHQ Master`** — Headquarters master (lowest field unit)
- name (PK, format HQ####), hq_code, hq_name, division (FK→tabDivision), region (FK→tabRegion Master), team (FK→tabTeam Master), zone (Data), per_capita, status

**`tabProduct Master`** — Product catalog
- name (PK, format PR####), product_code, product_name, pack, pack_conversion, division (Select: Prima/Vektra/ASPR/Wellness), product_group (Select: Dentist/Derma/Contus/Xptum/Amino/Ortho/Drez/Gynae/Jusdee/Dygerm/Others), category (Main Product/Hospital Product/New Product), mrp, ptr, pts, gst_rate, status

**`tabStockist Master`** — Stockist (distributor) master
- name (PK, format S####), stockist_code, stockist_name, hq (FK→tabHQ Master), division (FK→tabDivision), team (FK→tabTeam Master), region (FK→tabRegion Master), zone, address, city, contact_person, phone, email, status

**`tabDoctor Master`** — Doctor master
- name (PK, format D####), doctor_code, doctor_name, division (FK→tabDivision), qualification, doctor_category, specialization, phone, place, hospital_address, house_address, hq (FK→tabHQ Master), team (FK→tabTeam Master), region (FK→tabRegion Master), state (FK→tabState Master), zone, chemist_name, status

### Transaction Tables

**`tabStockist Statement`** — Monthly stockist stock statements (Submittable: docstatus 0=Draft,1=Submitted,2=Cancelled)
- name (PK), stockist_code (FK→tabStockist Master), stockist_name, hq (FK→tabHQ Master), division (FK→tabDivision), team (FK→tabTeam Master), region (FK→tabRegion Master), zone, statement_month (Date, 1st of month), uploaded_file, extracted_data_status (Pending/In Progress/Completed/Failed), extraction_notes, qc_confidence, total_opening_value, total_purchase_value, total_free_value, total_sales_value_pts, total_sales_value_ptr, total_closing_value, docstatus, owner, creation, modified

**`tabStockist Statement Item`** — Line items of stockist statement (child table, parent=tabStockist Statement)
- name, parent (FK→tabStockist Statement), parenttype='Stockist Statement', idx, product_code (FK→tabProduct Master), product_name, raw_product_name, mapping_status, pack, opening_qty, prev_month_closing, purchase_qty, sales_qty, free_qty, free_qty_scheme, conversion_factor, return_qty, misc_out_qty, closing_qty, pts, sales_value_pts, sales_value_ptr, opening_value, purchase_value, closing_value, scheme_deducted_qty_calc

**`tabScheme Request`** — Scheme/discount requests (Submittable)
- name (PK, format SCH-YYYY-#####), application_date, requested_by (FK→tabUser), division (FK→tabDivision), region (FK→tabRegion Master), team (FK→tabTeam Master), doctor_code (FK→tabDoctor Master), doctor_name, doctor_place, specialization, hospital_address, hq (FK→tabHQ Master), stockist_code (FK→tabStockist Master), stockist_name, total_scheme_value, approval_status (Pending/Approved/Rejected/Rerouted), repeated_request, scheme_notes, docstatus, owner, creation, modified

**`tabScheme Request Item`** — Line items of scheme request (child table, parent=tabScheme Request)
- name, parent (FK→tabScheme Request), idx, product_code (FK→tabProduct Master), product_name, pack, quantity, free_quantity, product_rate, special_rate, scheme_percentage, product_value

**`tabScheme Approval Log`** — Approval workflow log (child table, parent=tabScheme Request)
- name, parent (FK→tabScheme Request), idx, approver (FK→tabUser), approval_level, action (Approved/Rejected/Rerouted), action_date, comments

**`tabScheme Deduction`** — Deductions applied against statements (Submittable)
- name (PK, format SD-####), scheme_request (FK→tabScheme Request), scheme_date, doctor_code (FK→tabDoctor Master), division (FK→tabDivision), stockist_statement (FK→tabStockist Statement), stockist_code (FK→tabStockist Master), deduction_date, status (Draft/Applied/Cancelled), total_deducted_qty, total_deducted_value, notes, docstatus, owner, creation, modified

**`tabScheme Deduction Item`** — Line items of deduction (child table, parent=tabScheme Deduction)
- name, parent (FK→tabScheme Deduction), idx, product_code (FK→tabProduct Master), product_name, pack, scheme_free_qty, current_free_qty, deduct_qty, pts, deducted_value

**`tabPrimary Sales Data`** — Invoice-level primary sales data
- name (PK, hash), upload_month (YYYY-MM), division (FK→tabDivision), upload_ref (FK→tabPrimary Sales Upload), iscancelled, direct_party, stockist_code (Data), stockist_name, citypool, team (Data), region (Data), act_region, zonee, invoiceno, invoicedate, pcode, product (product name), product_head (product group), pack, batchno, quantity, freeqty, expqty, pts, ptsvalue, ptr, ptrvalue, mrp, mrpvalue, nrv, nrvvalue, dsort

**`tabPrimary Sales Upload`** — Upload job tracking
- name (PK, format PRI-YYYY-#####), upload_month (YYYY-MM), division (FK→tabDivision), uploaded_by (FK→tabUser), upload_date, file, total_rows, success_count, error_count, status (Pending/Processing/Completed/Failed), error_log

**`tabBulk Statement Upload`** — Bulk OCR processing jobs (Submittable)
- name (PK, format BULK-month-####), job_name, statement_month, zip_file, division (FK→tabDivision), status (Pending/Queued/In Progress/Completed/Failed/Partially Completed), progress, total_files, success_count, failed_count, skipped_count, job_id, extraction_log, docstatus, owner, creation, modified

**`tabHQ Yearly Target`** — Annual sales targets per HQ
- name (PK), division (FK→tabDivision), region (FK→tabRegion Master), financial_year (e.g. 2025-26), start_date, end_date, status (Draft/Active/Completed/Cancelled), total_target_amount (Lakhs), total_hqs, upload_date, uploaded_by

**`tabHQ Target Item`** — Monthly target breakdowns (child table, parent=tabHQ Yearly Target)
- name, parent (FK→tabHQ Yearly Target), idx, hq (FK→tabHQ Master), hq_name, team (FK→tabTeam Master), apr, may, jun, jul, aug, sep, oct, nov, dec, jan, feb, mar (all Currency in Lakhs), q1_total, q2_total, q3_total, q4_total, yearly_total

**`tabStockist Product Correction`** — OCR product name corrections/mappings
- name (PK, format SPC-####), stockist_code (FK→tabStockist Master), raw_product_name, mapped_product_code (FK→tabProduct Master), division (FK→tabDivision), created_from_statement (FK→tabStockist Statement), status

## IMPORTANT SQL RULES
1. Always use backticks around table names with spaces: `tabStockist Statement`
2. For submitted documents, filter by docstatus=1 unless user asks for drafts
3. Date fields: statement_month stores 1st of month (e.g. '2025-06-01' for June 2025)
4. Primary Sales upload_month is YYYY-MM format string (e.g. '2025-06')
5. Division filter: ALWAYS add division='{division}' to WHERE clause unless query is explicitly cross-division
6. Currency values: pts/ptr/mrp are per-unit prices. Values (_value suffix) are computed totals.
7. Use LEFT JOINs for optional relationships, INNER JOINs for required ones
8. Limit results to 500 rows max unless user asks for all
9. For aggregations, always include meaningful GROUP BY
10. Financial year runs April to March (e.g. 2025-26 = Apr 2025 - Mar 2026)
11. Use COALESCE for nullable numeric fields in aggregations

## RESPONSE FORMAT
You MUST respond with a valid JSON object in one of these formats:

### For data queries:
{{
  "type": "data",
  "sql": "SELECT ... FROM ... WHERE ...",
  "title": "Brief title describing the results",
  "description": "One-line explanation of what the data shows",
  "columns": ["col1", "col2", ...],
  "format": "table"
}}

### For chart-worthy queries (trends, comparisons, distributions):
{{
  "type": "chart",
  "sql": "SELECT ... FROM ... WHERE ...",
  "title": "Chart title",
  "description": "What the chart shows",
  "chart_type": "bar|line|pie|doughnut|horizontalBar",
  "label_column": "column_name_for_labels",
  "value_columns": ["col1", "col2"],
  "columns": ["col1", "col2", ...],
  "format": "chart"
}}

### For count/single-value queries:
{{
  "type": "metric",
  "sql": "SELECT COUNT(*) as count FROM ...",
  "title": "Metric title",
  "description": "What this metric represents",
  "columns": ["count"],
  "format": "metric"
}}

### For conversational/non-data questions:
{{
  "type": "text",
  "message": "Your helpful response here"
}}

### For errors or unclear questions:
{{
  "type": "error",
  "message": "Explain what's wrong or ask for clarification"
}}

## CHART SELECTION GUIDELINES
- Time series / monthly trends → "line"
- Comparisons across categories (regions, teams, products) → "bar" or "horizontalBar"
- Distribution / share / percentage breakdown → "pie" or "doughnut"
- Top N rankings → "horizontalBar"
- If user says "chart" or "graph", pick the best chart_type automatically
- For multi-series charts, use multiple value_columns

## ANALYTICAL CAPABILITIES
You can answer questions like:
- Sales performance by region/team/HQ/stockist
- Product-wise sales trends and comparisons
- Scheme utilization and deduction analysis
- Target vs achievement analysis
- Stock movement patterns (opening, purchase, sales, closing)
- Top/bottom performers at any hierarchy level
- Month-over-month and year-over-year comparisons
- Stockist statement submission tracking
- Product group analysis
- Zone/Region wise breakdowns

Always think step by step about which tables to join and what aggregations are needed.
"""


@frappe.whitelist()
def chatbot_query(message, conversation_history=None):
    """
    Process a natural language query through the Text2SQL pipeline.
    Uses Gemini to convert question to SQL, executes it, and returns formatted results.
    """
    if not message or not message.strip():
        return {"success": False, "error": "Please enter a question."}

    try:
        api_key, model_name, is_enabled = get_gemini_settings()
    except Exception:
        return {"success": False, "error": "Gemini AI is not configured. Please enable it in Scanify Settings."}

    division = get_user_division()

    # Build the system prompt with division context
    system_prompt = SCANIFY_SCHEMA_PROMPT.format(division=division)

    # Build conversation messages
    messages = []
    if conversation_history:
        if isinstance(conversation_history, str):
            try:
                conversation_history = json.loads(conversation_history)
            except (json.JSONDecodeError, TypeError):
                conversation_history = []
        # Include last 10 exchanges for context
        for entry in conversation_history[-10:]:
            role = entry.get("role", "user")
            content = entry.get("content", "")
            if role == "user":
                messages.append({"role": "user", "content": content})
            elif role == "assistant":
                messages.append({"role": "assistant", "content": content})

    messages.append({"role": "user", "content": message})

    try:
        client = genai_sdk.Client(api_key=api_key)

        # Build contents for Gemini
        contents = []
        for msg in messages:
            role = "user" if msg["role"] == "user" else "model"
            contents.append(genai_types.Content(
                role=role,
                parts=[genai_types.Part(text=msg["content"])]
            ))

        response = client.models.generate_content(
            model=model_name,
            contents=contents,
            config=genai_types.GenerateContentConfig(
                system_instruction=system_prompt,
                temperature=0.1,
                max_output_tokens=8192,
            )
        )

        ai_text = response.text.strip()

        # Extract JSON from response (handle markdown code blocks)
        json_str = ai_text
        if "```json" in json_str:
            json_str = json_str.split("```json")[1].split("```")[0].strip()
        elif "```" in json_str:
            json_str = json_str.split("```")[1].split("```")[0].strip()

        try:
            ai_response = json.loads(json_str)
        except json.JSONDecodeError:
            # If AI returned plain text, treat as text response
            return {
                "success": True,
                "type": "text",
                "message": ai_text,
                "raw_response": ai_text
            }

        response_type = ai_response.get("type", "text")

        if response_type == "text":
            return {
                "success": True,
                "type": "text",
                "message": ai_response.get("message", ai_text),
                "raw_response": ai_text
            }

        if response_type == "error":
            return {
                "success": True,
                "type": "error",
                "message": ai_response.get("message", "I couldn't understand that query."),
                "raw_response": ai_text
            }

        sql = ai_response.get("sql", "")
        if not sql:
            return {
                "success": True,
                "type": "text",
                "message": ai_response.get("message", ai_text),
                "raw_response": ai_text
            }

        # Security: validate the SQL is read-only
        sql_upper = sql.strip().upper()
        # Use word-boundary regex to avoid false positives (e.g. LOAD in UPLOAD)
        forbidden_patterns = [
            r'\bINSERT\b', r'\bUPDATE\b', r'\bDELETE\b', r'\bDROP\b',
            r'\bALTER\b', r'\bCREATE\b', r'\bTRUNCATE\b', r'\bGRANT\b',
            r'\bREVOKE\b', r'\bEXEC\b', r'\bEXECUTE\b', r'\bCALL\b',
            r'\bLOAD\b', r'INTO\s+OUTFILE', r'INTO\s+DUMPFILE',
            r'\bINFORMATION_SCHEMA\b', r'\bSLEEP\s*\(', r'\bBENCHMARK\s*\('
        ]
        for pattern in forbidden_patterns:
            if re.search(pattern, sql_upper):
                return {
                    "success": False,
                    "error": "Only read-only SELECT queries are allowed.",
                    "raw_response": ai_text
                }

        if not sql_upper.lstrip().startswith("SELECT") and not sql_upper.lstrip().startswith("WITH"):
            return {
                "success": False,
                "error": "Only SELECT queries are permitted.",
                "raw_response": ai_text
            }

        # Execute the SQL query
        try:
            results = frappe.db.sql(sql, as_dict=True)
        except Exception as db_err:
            # If query fails, ask Gemini to fix it
            error_msg = str(db_err)
            retry_prompt = f"The SQL query failed with error: {error_msg}\n\nOriginal query: {sql}\n\nPlease fix the SQL and respond with the corrected JSON."
            contents.append(genai_types.Content(
                role="model",
                parts=[genai_types.Part(text=ai_text)]
            ))
            contents.append(genai_types.Content(
                role="user",
                parts=[genai_types.Part(text=retry_prompt)]
            ))

            retry_response = client.models.generate_content(
                model=model_name,
                contents=contents,
                config=genai_types.GenerateContentConfig(
                    system_instruction=system_prompt,
                    temperature=0.1,
                    max_output_tokens=8192,
                )
            )

            retry_text = retry_response.text.strip()
            retry_json = retry_text
            if "```json" in retry_json:
                retry_json = retry_json.split("```json")[1].split("```")[0].strip()
            elif "```" in retry_json:
                retry_json = retry_json.split("```")[1].split("```")[0].strip()

            try:
                retry_parsed = json.loads(retry_json)
                retry_sql = retry_parsed.get("sql", "")
                if retry_sql:
                    # Validate retry SQL too
                    retry_upper = retry_sql.strip().upper()
                    for pattern in forbidden_patterns:
                        if re.search(pattern, retry_upper):
                            return {"success": False, "error": "Only read-only queries are allowed."}
                    if not retry_upper.lstrip().startswith("SELECT") and not retry_upper.lstrip().startswith("WITH"):
                        return {"success": False, "error": "Only SELECT queries are permitted."}

                    results = frappe.db.sql(retry_sql, as_dict=True)
                    ai_response = retry_parsed
                    sql = retry_sql
                else:
                    return {
                        "success": False,
                        "error": f"Query failed: {error_msg}",
                        "raw_response": ai_text
                    }
            except Exception:
                return {
                    "success": False,
                    "error": f"Query failed: {error_msg}",
                    "raw_response": ai_text
                }

        # Format results
        rows = []
        for row in results[:500]:
            formatted_row = {}
            for key, val in row.items():
                if val is None:
                    formatted_row[key] = ""
                elif isinstance(val, (int, float)):
                    formatted_row[key] = val
                else:
                    formatted_row[key] = str(val)
            rows.append(formatted_row)

        columns = ai_response.get("columns", list(rows[0].keys()) if rows else [])

        response_data = {
            "success": True,
            "type": response_type,
            "title": ai_response.get("title", "Query Results"),
            "description": ai_response.get("description", ""),
            "columns": columns,
            "data": rows,
            "total_rows": len(rows),
            "sql": sql,
            "raw_response": ai_text
        }

        # Add chart config if chart type
        if response_type == "chart":
            response_data["chart_type"] = ai_response.get("chart_type", "bar")
            response_data["label_column"] = ai_response.get("label_column", columns[0] if columns else "")
            response_data["value_columns"] = ai_response.get("value_columns", columns[1:] if len(columns) > 1 else columns)

        return response_data

    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Chatbot Query Error")
        return {
            "success": False,
            "error": f"An error occurred: {str(e)}"
        }

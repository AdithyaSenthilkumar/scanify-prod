import frappe
import json
from frappe import _
from frappe.utils import flt, nowdate, add_months, get_first_day
import requests
import os
import base64
import mimetypes
from frappe import _
import google.generativeai as genai
from PIL import Image
import io
import zipfile
import tempfile
import time
from google.api_core.exceptions import ResourceExhausted
import re
from difflib import SequenceMatcher

GEMINI_API_KEY = "AIzaSyBT5OH6cAQ0oLNKYTmRENCoJDtzbivgeLE"
if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)

@frappe.whitelist()
def extract_stockist_statement(doc_name, file_url):
    """Extract stockist statement data using Gemini AI - TWO STAGE APPROACH"""
    doc = None
    try:
        doc = frappe.get_doc("Stockist Statement", doc_name)
        doc.extracted_data_status = "In Progress"
        doc.save()
        frappe.db.commit()

        if not file_url:
            raise ValueError("No file uploaded")

        # Get file path using Frappe's utility
        from frappe.utils.file_manager import get_file_path
        file_path = get_file_path(file_url)

        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_url}")

        # STAGE 1: Extract raw data WITHOUT product list (reduces tokens massively)
        extracted_data = call_gemini_extraction_two_stage(file_path, doc.stockist_code)

        if not extracted_data or len(extracted_data) == 0:
            doc.extracted_data_status = "Failed"
            doc.extraction_notes = "No data extracted - AI returned empty results"
            doc.save()
            frappe.db.commit()
            return {
                "success": False,
                "message": "No data extracted - AI returned empty results"
            }

        # Clear existing items
        doc.items = []

        # Add extracted items
        items_added = 0
        for item_data in extracted_data:
            if item_data.get("product_code"):  # Only add if product code exists
                doc.append("items", {
                    "product_code": item_data.get("product_code"),
                    "opening_qty": flt(item_data.get("opening_qty", 0)),
                    "purchase_qty": flt(item_data.get("purchase_qty", 0)),
                    "sales_qty": flt(item_data.get("sales_qty", 0)),
                    "free_qty": flt(item_data.get("free_qty", 0)),
                    "return_qty": flt(item_data.get("return_qty", 0)),
                    "misc_out_qty": flt(item_data.get("misc_out_qty", 0)),
                })
                items_added += 1

        if items_added == 0:
            doc.extracted_data_status = "Failed"
            doc.extraction_notes = "No valid products found in extracted data"
            doc.save()
            frappe.db.commit()
            return {
                "success": False,
                "message": "No valid products found in extracted data"
            }

        # Success - save with completed status
        doc.extracted_data_status = "Completed"
        doc.extraction_notes = f"Successfully extracted {items_added} items using AI (two-stage approach)"
        doc.calculate_closing_and_totals()
        doc.save()
        frappe.db.commit()

        return {
            "success": True,
            "message": f"Successfully extracted {items_added} items"
        }

    except Exception as e:
        error_msg = str(e)
        frappe.log_error(frappe.get_traceback(), "Gemini Extraction Error")

        # Update document with error if possible
        if doc:
            try:
                doc.extracted_data_status = "Failed"
                doc.extraction_notes = f"Extraction failed: {error_msg}"
                doc.save()
                frappe.db.commit()
            except:
                pass

        return {
            "success": False,
            "message": f"Extraction failed: {error_msg}"
        }


def call_gemini_extraction_two_stage(file_path, stockist_code):
    """
    TWO-STAGE EXTRACTION:
    Stage 1: Extract raw product data without sending product list (saves massive tokens)
    Stage 2: Match extracted products to product codes locally (no API call)
    """

    if not GEMINI_API_KEY:
        frappe.throw("Gemini API key not configured. Please set 'gemini_api_key' in site_config.json")

    try:
        # Get product master for Stage 2 matching (but DON'T send to Gemini)
        products = frappe.get_all("Product Master", 
            fields=["product_code", "product_name", "pack", "pts"],
            filters={"status": "Active"})

        # Determine file type and process accordingly
        mime_type, _ = mimetypes.guess_type(file_path)
        file_ext = os.path.splitext(file_path)[1].lower()

        # ========== STAGE 1: EXTRACT WITHOUT PRODUCT LIST ==========
        # Simplified prompt - NO PRODUCT LIST to reduce tokens
        prompt = """You are extracting product-wise stock data from a pharmaceutical stockist statement.

RULES FOR UNDERSTANDING TABLE ROWS:
1. A valid product row ALWAYS has at least one numeric quantity (opening, purchase, sales, free, return, closing).
2. If a line contains ONLY product name + unit but NO numeric columns, it is NOT a separate product.
   → It is a continuation/description of the next line. IGNORE such lines.
3. If two consecutive lines have the same product name/unit, merge them and extract quantities ONLY from the numeric row.
4. DO NOT create duplicate items for the same product in the same statement.

WHAT TO EXTRACT PER PRODUCT (ONE ENTRY PER PRODUCT ONLY):
- Product Name (exact text from document)
- Pack Size (e.g., 10TAB, 15CAP, 200ML)
- Opening Quantity (Op.Qty, OPSTK, Opening, Op Stock)
- Purchase Quantity (Purch.Qty, PURCH, Receipt, Pr.Qty, Purchase)
- Sales Quantity (Sales, Sale, Sl, Sold)
- Free Quantity (Free Qty, Free, Scheme Qty, Gift)
- Sales Return (Return, Ret, Sales Ret, Sl Ret)
- Misc Out/Transfer Out (Misc.Out, M.Out, Others, Trans Out, Transfer)

IMPORTANT:
- Many products may appear in multiple lines. Only keep the MOST DETAILED row — the one that has quantity values.
- Ignore lines that have NO quantities or NO numbers.
- Ignore unit-only lines (e.g. "AMINORICH CAP 10 CAP").
- Ignore header lines, totals, summaries, and continuation labels like "Continued..."

RETURN ONLY A JSON ARRAY LIKE:

[
  {
    "product_name": "AMINORICH CAP",
    "pack": "15 CAP",
    "opening_qty": 88,
    "purchase_qty": 29,
    "sales_qty": 59,
    "free_qty": 0,
    "return_qty": 0,
    "misc_out_qty": 0
  },
  ...
]

NO MARKDOWN, NO EXPLANATION — ONLY VALID JSON.

"""

        # Process based on file type with RETRY and EXPONENTIAL BACKOFF
        model = genai.GenerativeModel('gemini-2.5-flash')
        max_retries = 3
        base_delay = 2  # seconds

        response = None
        for attempt in range(max_retries):
            try:
                if file_ext in ['.pdf', '.jpg', '.jpeg', '.png']:
                    # Handle images and PDFs
                    with open(file_path, 'rb') as f:
                        file_data = f.read()

                    # For PDFs, convert first page to image with REDUCED QUALITY
                    if file_ext == '.pdf':
                        import fitz  # PyMuPDF
                        pdf = fitz.open(file_path)
                        page = pdf[0]
                        # Reduce resolution to 50% to save tokens
                        pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5))
                        img_data = pix.tobytes("png")
                        image_part = {
                            'mime_type': 'image/png',
                            'data': img_data
                        }
                    else:
                        image_part = {
                            'mime_type': mime_type or 'image/jpeg',
                            'data': file_data
                        }

                    response = model.generate_content([prompt, image_part])

                elif file_ext in ['.csv', '.txt']:
                    # Handle CSV/TXT files
                    with open(file_path, 'r', encoding='utf-8') as f:
                        file_content = f.read()

                    response = model.generate_content(f"{prompt}\n\nFile content:\n{file_content}")

                elif file_ext in ['.xls', '.xlsx']:
                    # Handle Excel files
                    import pandas as pd
                    df = pd.read_excel(file_path)
                    file_content = df.to_string()

                    response = model.generate_content(f"{prompt}\n\nExcel content:\n{file_content}")

                else:
                    frappe.throw(f"Unsupported file type: {file_ext}")

                # If we reach here, the API call succeeded
                break

            except ResourceExhausted as e:
                if attempt == max_retries - 1:
                    # Last attempt failed, re-raise
                    frappe.throw(f"Gemini API rate limit exceeded after {max_retries} retries. Please try again in a few minutes.")

                # Calculate exponential backoff delay
                delay = base_delay * (2 ** attempt)
                frappe.logger().warning(f"Rate limit hit (attempt {attempt + 1}/{max_retries}), retrying in {delay}s...")
                time.sleep(delay)

            except Exception as e:
                # Other errors, don't retry
                raise e

        # Log response
        try:
            frappe.logger().info("🤖 GEMINI RAW RESPONSE (TWO-STAGE) START >>>")
            frappe.logger().info(response.text)
            frappe.logger().info("🤖 GEMINI RAW RESPONSE END <<<")
        except Exception as log_err:
            frappe.logger().error(f"Failed to log Gemini response: {log_err}")

        # Parse response
        response_text = response.text.strip()

        # Remove markdown code blocks if present
        if response_text.startswith('```'):
            response_text = response_text.split('```')[1]
            if response_text.startswith('json'):
                response_text = response_text[4:]

        frappe.logger().info("🧹 Cleaned Response Text >>>")
        frappe.logger().info(response_text)

        extracted_items = json.loads(response_text or "[]")
        frappe.logger().info(f"📊 Parsed Items Count: {len(extracted_items)}")

        # ========== STAGE 2: MATCH PRODUCTS LOCALLY (NO API CALL) ==========
        result = match_products_locally(extracted_items, products)

        frappe.logger().info(f"✅ Final Matched Products: {len(result)}")
        return result

    except Exception as e:
        frappe.log_error(frappe.get_traceback(), "Gemini API Call Error (Two-Stage)")
        frappe.throw(f"Extraction failed: {str(e)}")


def match_products_locally(extracted_items, products):
    """
    STAGE 2: Match extracted product names to product codes LOCALLY (no API)

    - Robust normalisation (TAB/TABLET, CAP/CAPSULE, OINT/OINTMENT, GM/G, etc.)
    - Uses BOTH name similarity and pack similarity
    - Logs best candidate for unmatched items
    """
    result = []
    unmatched_logs = []

    # Pre-process products once
    enriched_products = []
    for p in products:
        enriched_products.append({
            "product_code": p["product_code"],
            "product_name": p["product_name"],
            "pack": p.get("pack") or "",
            "norm_tokens": _normalise_tokens(p["product_name"]),
        })

    for item in extracted_items:
        raw_name = (item.get("product_name") or "").strip()
        raw_pack = (item.get("pack") or "").strip()

        if not raw_name:
            continue

        best = None
        best_score = 0.0

        for p in enriched_products:
            # Name similarity
            name_score = _name_similarity(raw_name, p["product_name"])
            # Pack similarity
            pack_score = _pack_similarity(raw_pack, p["pack"])

            # Weight: name 80%, pack 20%
            total_score = 0.8 * name_score + 0.2 * pack_score

            if total_score > best_score:
                best_score = total_score
                best = p

        # Threshold – tweak if needed
        if best and best_score >= 0.55:
            result.append({
                "product_code": best["product_code"],
                "opening_qty": item.get("opening_qty", 0),
                "purchase_qty": item.get("purchase_qty", 0),
                "sales_qty": item.get("sales_qty", 0),
                "free_qty": item.get("free_qty", 0),
                "return_qty": item.get("return_qty", 0),
                "misc_out_qty": item.get("misc_out_qty", 0),
            })
        else:
            unmatched_logs.append(f"{raw_name} ({raw_pack}) [best_score={best_score:.2f}, best={best['product_name'] if best else 'None'}]")

    if unmatched_logs:
        frappe.logger().warning("⚠️ Unmatched / low-confidence products:")
        for line in unmatched_logs:
            frappe.logger().warning(f" - {line}")

    return result
def _clean_text(text: str) -> str:
    """Uppercase, remove extra punctuation/spaces."""
    if not text:
        return ""
    # Upper + replace separators with space
    text = text.upper()
    text = re.sub(r"[\-_/.,()]+", " ", text)
    # Collapse multiple spaces
    text = re.sub(r"\s+", " ", text).strip()
    return text

# Domain-specific normalisation
_SYNONYM_MAP = {
    "TAB": "TAB",
    "TABS": "TAB",
    "TABLET": "TAB",
    "TABLETS": "TAB",

    "CAP": "CAP",
    "CAPS": "CAP",
    "CAPSULE": "CAP",
    "CAPSULES": "CAP",

    "OINT": "OINT",
    "OINT.": "OINT",
    "ONT": "OINT",
    "OINTMENT": "OINT",

    "POWDER": "PWD",
    "POWDER.": "PWD",
    "PWD": "PWD",

    "CREAM": "CRM",
    "CRM": "CRM",

    "DROP": "DROP",
    "DROPS": "DROP",

    "LIQ": "LIQ",
    "LIQUID": "LIQ",

    "SOL": "SOL",
    "SOLUTION": "SOL",

    "GARGLE": "GARGLE",
}

def _normalise_tokens(text: str):
    """
    Turn a product name into a set of normalised tokens,
    stripping pack-like fragments (1X10S, 150 ML, 5GM).
    """
    text = _clean_text(text)
    if not text:
        return set()

    # Split and clean tokens
    tokens = []
    for tok in text.split():
        # Pack-like token? keep for pack matching but not for name similarity
        if re.match(r"^\d+(\.\d+)?(ML|GM|G|MG|KG)$", tok):
            continue
        if re.match(r"^\d+X\d+S?$", tok):
            continue

        mapped = _SYNONYM_MAP.get(tok, tok)
        tokens.append(mapped)

    return set(tokens)

def _normalise_pack(pack: str) -> str:
    """Normalise pack like '30 GM', '30G', '150 ML', '2x15s' -> comparable form."""
    if not pack:
        return ""

    pack = _clean_text(pack)

    # 2x15s -> 15CAP (we only care about the unit qty)
    m = re.search(r"(\d+)\s*X\s*(\d+)", pack)
    if m:
        inner = m.group(2)
        # treat as 'inner' units, type unknown -> just number
        return inner

    # Extract number + unit
    m = re.search(r"(\d+(\.\d+)?)\s*(ML|GM|G|MG|KG)?", pack)
    if not m:
        return pack

    qty = m.group(1)
    unit = m.group(3) or ""

    # GM and G are basically same for our comparison
    if unit in ("GM", "G"):
        unit = "G"
    return f"{qty}{unit}"

def _pack_similarity(p1: str, p2: str) -> float:
    """Simple pack similarity 0..1."""
    n1 = _normalise_pack(p1)
    n2 = _normalise_pack(p2)
    if not n1 or not n2:
        return 0.0
    if n1 == n2:
        return 1.0
    # If just the numeric part matches (e.g. 30G vs 30MG) give partial credit
    num1 = re.findall(r"\d+(\.\d+)?", n1)
    num2 = re.findall(r"\d+(\.\d+)?", n2)
    if num1 and num2 and num1[0] == num2[0]:
        return 0.7
    return 0.0

def _token_jaccard(a: set, b: set) -> float:
    """Jaccard similarity of token sets."""
    if not a or not b:
        return 0.0
    inter = len(a & b)
    union = len(a | b)
    return inter / union

def _name_similarity(extracted_name: str, master_name: str) -> float:
    """
    Combine token Jaccard + raw SequenceMatcher for robustness.
    """
    tokens_a = _normalise_tokens(extracted_name)
    tokens_b = _normalise_tokens(master_name)

    jacc = _token_jaccard(tokens_a, tokens_b)

    # Raw similarity on cleaned strings
    clean_a = _clean_text(extracted_name)
    clean_b = _clean_text(master_name)
    seq = SequenceMatcher(None, clean_a, clean_b).ratio()

    # Weight both
    return 0.6 * jacc + 0.4 * seq


# ========== REST OF THE API FILE (UNCHANGED) ==========

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
                                "status": "Skipped",
                                "message": f"Statement already exists: {existing}",
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
                        
                        # Extract data
                        extracted_data = call_gemini_extraction_two_stage(file_full_path, stockist_code)
                        
                        if extracted_data and len(extracted_data) > 0:
                            for item_data in extracted_data:
                                statement.append("items", {
                                    "product_code": item_data.get("product_code"),
                                    "opening_qty": flt(item_data.get("opening_qty", 0)),
                                    "purchase_qty": flt(item_data.get("purchase_qty", 0)),
                                    "sales_qty": flt(item_data.get("sales_qty", 0)),
                                    "free_qty": flt(item_data.get("free_qty", 0)),
                                    "return_qty": flt(item_data.get("return_qty", 0)),
                                    "misc_out_qty": flt(item_data.get("misc_out_qty", 0)),
                                })
                            
                            statement.extracted_data_status = "Completed"
                            statement.extraction_notes = f"Extracted {len(extracted_data)} products successfully"
                        else:
                            statement.extracted_data_status = "Failed"
                            statement.extraction_notes = "No data extracted from file"
                        
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
                hospital_clinic,
                city_pool,
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
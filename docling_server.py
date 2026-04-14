#!/usr/bin/env python3
"""
Docling-Powered Document Analysis Server
=========================================
Local REST API that wraps the Docling library to provide deep document parsing.
Returns structured JSON with precise bounding boxes for all text elements,
tables, key-value pairs, and document structure.

Usage:
    pip install docling flask flask-cors
    python docling_server.py
"""

import os
import sys
import json
import tempfile
import traceback
from pathlib import Path

from flask import Flask, request, jsonify
from flask_cors import CORS

# ═══════════════════════════════════════════════════════════════
#  Docling imports
# ═══════════════════════════════════════════════════════════════
try:
    from docling.document_converter import DocumentConverter
    from docling.datamodel.base_models import InputFormat
    from docling.datamodel.pipeline_options import PdfPipelineOptions
    from docling.pipeline.standard_pdf_pipeline import StandardPdfPipeline
    DOCLING_AVAILABLE = True
except ImportError:
    DOCLING_AVAILABLE = False
    print("⚠ Docling not installed. Run: pip install docling")
    print("  Server will start but return mock data for testing.")

app = Flask(__name__)
CORS(app)

# ═══════════════════════════════════════════════════════════════
#  Configuration
# ═══════════════════════════════════════════════════════════════
PORT = 5050
UPLOAD_FOLDER = tempfile.mkdtemp(prefix='docling_')

# ═══════════════════════════════════════════════════════════════
#  Helper: Extract structured data from DoclingDocument
# ═══════════════════════════════════════════════════════════════
def extract_structured_data(result):
    """
    Convert Docling's result into a structured JSON format
    optimized for the frontend generation pipeline.
    """
    doc = result.document
    output = {
        "status": "success",
        "documentType": "Document",
        "totalPages": 0,
        "pages": [],
        "textItems": [],
        "tables": [],
        "keyValuePairs": [],
        "fullText": "",
        "metadata": {}
    }

    # ── Export full text ──
    try:
        output["fullText"] = doc.export_to_markdown()
    except Exception:
        output["fullText"] = str(doc)

    # ── Detect document type from text ──
    text_upper = output["fullText"].upper()
    if "SALES CONTRACT" in text_upper or "SALE CONTRACT" in text_upper:
        output["documentType"] = "Sales Contract"
    elif "COMMERCIAL INVOICE" in text_upper:
        output["documentType"] = "Commercial Invoice"
    elif "PROFORMA INVOICE" in text_upper:
        output["documentType"] = "Proforma Invoice"
    elif "INVOICE" in text_upper:
        output["documentType"] = "Invoice"
    elif "CREDIT NOTE" in text_upper:
        output["documentType"] = "Credit Note"
    elif "DEBIT NOTE" in text_upper:
        output["documentType"] = "Debit Note"
    elif "PACKING LIST" in text_upper:
        output["documentType"] = "Packing List"
    elif "BILL OF LADING" in text_upper or "B/L" in text_upper:
        output["documentType"] = "Bill of Lading"
    elif "CERTIFICATE OF ORIGIN" in text_upper:
        output["documentType"] = "Certificate of Origin"
    elif "PURCHASE ORDER" in text_upper:
        output["documentType"] = "Purchase Order"
    elif "CONTRACT" in text_upper or "AGREEMENT" in text_upper:
        output["documentType"] = "Contract"

    # ── Extract text blocks with bounding boxes ──
    # Docling provides document items with provenance (bounding boxes)
    item_index = 0
    try:
        for item, _level in doc.iterate_items():
            text_item = {
                "index": item_index,
                "text": "",
                "type": "text",
                "page": 1,
                "bbox": None,
                "label": getattr(item, 'label', None),
            }

            # Get text content
            if hasattr(item, 'text'):
                text_item["text"] = item.text
            elif hasattr(item, 'export_to_markdown'):
                text_item["text"] = item.export_to_markdown()

            # Get bounding box from provenance
            if hasattr(item, 'prov') and item.prov:
                for prov in item.prov:
                    if hasattr(prov, 'page_no'):
                        text_item["page"] = prov.page_no
                    if hasattr(prov, 'bbox'):
                        bbox = prov.bbox
                        if hasattr(bbox, 'l'):
                            text_item["bbox"] = {
                                "x": float(bbox.l),
                                "y": float(bbox.t),
                                "width": float(bbox.r - bbox.l),
                                "height": float(bbox.b - bbox.t),
                                "x2": float(bbox.r),
                                "y2": float(bbox.b)
                            }
                        elif hasattr(bbox, 'x'):
                            text_item["bbox"] = {
                                "x": float(bbox.x),
                                "y": float(bbox.y),
                                "width": float(getattr(bbox, 'width', 100)),
                                "height": float(getattr(bbox, 'height', 14)),
                            }

            # Classify item type
            label = str(getattr(item, 'label', '')).lower()
            if 'table' in label:
                text_item["type"] = "table"
            elif 'title' in label or 'heading' in label:
                text_item["type"] = "heading"
            elif 'caption' in label:
                text_item["type"] = "caption"
            elif 'list' in label:
                text_item["type"] = "list_item"
            elif 'paragraph' in label or 'text' in label:
                text_item["type"] = "paragraph"

            if text_item["text"].strip():
                output["textItems"].append(text_item)
                item_index += 1
    except Exception as e:
        print(f"  ⚠ Error iterating items: {e}")

    # ── Extract tables ──
    try:
        if hasattr(doc, 'tables') and doc.tables:
            for table in doc.tables:
                table_data = {
                    "page": 1,
                    "rows": [],
                    "bbox": None,
                    "numRows": 0,
                    "numCols": 0,
                }

                # Get table bounding box
                if hasattr(table, 'prov') and table.prov:
                    for prov in table.prov:
                        if hasattr(prov, 'page_no'):
                            table_data["page"] = prov.page_no
                        if hasattr(prov, 'bbox'):
                            bbox = prov.bbox
                            if hasattr(bbox, 'l'):
                                table_data["bbox"] = {
                                    "x": float(bbox.l),
                                    "y": float(bbox.t),
                                    "width": float(bbox.r - bbox.l),
                                    "height": float(bbox.b - bbox.t),
                                }

                # Extract table content
                if hasattr(table, 'export_to_dataframe'):
                    try:
                        df = table.export_to_dataframe()
                        table_data["numRows"] = len(df)
                        table_data["numCols"] = len(df.columns)
                        table_data["headers"] = list(df.columns)
                        table_data["rows"] = df.values.tolist()
                    except Exception:
                        pass

                output["tables"].append(table_data)
    except Exception as e:
        print(f"  ⚠ Error extracting tables: {e}")

    # ── Extract key-value pairs ──
    # Docling's structured output often includes key-value patterns
    try:
        _extract_key_value_pairs(output)
    except Exception as e:
        print(f"  ⚠ Error extracting key-value pairs: {e}")

    # ── Count pages ──
    if output["textItems"]:
        output["totalPages"] = max(item.get("page", 1) for item in output["textItems"])
    else:
        output["totalPages"] = 1

    # ── Build per-page structure ──
    for p in range(1, output["totalPages"] + 1):
        page_items = [item for item in output["textItems"] if item.get("page") == p]
        output["pages"].append({
            "pageNumber": p,
            "itemCount": len(page_items),
            "items": page_items,
        })

    return output


def _extract_key_value_pairs(output):
    """
    Analyze text items to detect label:value patterns.
    Uses spatial analysis — items on the same Y-level with
    a label-like item on the left and a value on the right.
    """
    import re

    # Common label patterns in trade documents
    label_patterns = [
        # English
        (r'(?:^|\b)(NO|No\.?|Number|Ref)[\s.:]*$', 'SO_CHUNG_TU', 'Document Number'),
        (r'(?:^|\b)(DATE|Date|Dated)[\s.:]*$', 'NGAY', 'Date'),
        (r'(?:^|\b)(BUYER|Buyer)[\s.:]*$', 'TEN_NGUOI_MUA', 'Buyer'),
        (r'(?:^|\b)(SELLER|Seller)[\s.:]*$', 'TEN_NGUOI_BAN', 'Seller'),
        (r'(?:^|\b)(ADDRESS|Address)[\s.:]*$', 'DIA_CHI', 'Address'),
        (r'(?:^|\b)(COMMODITY|Goods|Description)[\s.:]*$', 'MO_TA_HANG', 'Goods Description'),
        (r'(?:^|\b)(QUANTITY|Qty|QTY)[\s.:]*$', 'SO_LUONG', 'Quantity'),
        (r'(?:^|\b)(UNIT\s*PRICE|Price)[\s.:]*$', 'DON_GIA', 'Unit Price'),
        (r'(?:^|\b)(TOTAL|Amount|AMOUNT)[\s.:]*$', 'TONG_TIEN', 'Total Amount'),
        (r'(?:^|\b)(PAYMENT|Payment\s*Terms?)[\s.:]*$', 'THANH_TOAN', 'Payment Terms'),
        (r'(?:^|\b)(SHIPMENT|Delivery)[\s.:]*$', 'GIAO_HANG', 'Shipment'),
        (r'(?:^|\b)(PORT\s*OF\s*LOADING)[\s.:]*$', 'CANG_XUAT', 'Loading Port'),
        (r'(?:^|\b)(PORT\s*OF\s*(?:DISCHARGE|DESTINATION))[\s.:]*$', 'CANG_NHAP', 'Discharge Port'),
        (r'(?:^|\b)(VESSEL|Ship)[\s.:]*$', 'TAU', 'Vessel'),
        (r'(?:^|\b)(BANK|Bank\s*Name)[\s.:]*$', 'NGAN_HANG', 'Bank'),
        (r'(?:^|\b)(SWIFT)[\s.:]*$', 'SWIFT', 'SWIFT Code'),
        (r'(?:^|\b)(ACCOUNT)[\s.:]*$', 'TAI_KHOAN', 'Account'),
        (r'(?:^|\b)(TEL|Phone|Telephone)[\s.:]*$', 'DIEN_THOAI', 'Phone'),
        (r'(?:^|\b)(FAX)[\s.:]*$', 'FAX', 'Fax'),
        (r'(?:^|\b)(CURRENCY)[\s.:]*$', 'LOAI_TIEN', 'Currency'),
        (r'(?:^|\b)(CONTAINER)[\s.:]*$', 'SO_CONTAINER', 'Container No.'),
        (r'(?:^|\b)(NOTIFY\s*PARTY)[\s.:]*$', 'BEN_THONG_BAO', 'Notify Party'),
        (r'(?:^|\b)(CONSIGNEE)[\s.:]*$', 'NGUOI_NHAN', 'Consignee'),
        (r'(?:^|\b)(SHIPPER)[\s.:]*$', 'NGUOI_GUI', 'Shipper'),
        (r'(?:^|\b)(APPLICANT)[\s.:]*$', 'NGUOI_YEU_CAU', 'Applicant'),
        (r'(?:^|\b)(BENEFICIARY)[\s.:]*$', 'NGUOI_THU_HUONG', 'Beneficiary'),
        (r'(?:^|\b)(ISSUING\s*BANK)[\s.:]*$', 'NH_PHAT_HANH', 'Issuing Bank'),
        # Vietnamese
        (r'(?:^|\b)(Số|SO)[\s.:]*$', 'SO_CHUNG_TU', 'Số chứng từ'),
        (r'(?:^|\b)(Ngày|NGAY)[\s.:]*$', 'NGAY', 'Ngày'),
        (r'(?:^|\b)(Người\s*mua|NGUOI\s*MUA)[\s.:]*$', 'TEN_NGUOI_MUA', 'Người mua'),
        (r'(?:^|\b)(Người\s*bán|NGUOI\s*BAN)[\s.:]*$', 'TEN_NGUOI_BAN', 'Người bán'),
        (r'(?:^|\b)(Địa\s*chỉ|DIA\s*CHI)[\s.:]*$', 'DIA_CHI', 'Địa chỉ'),
    ]

    # Strategy 1: Look for inline "LABEL: VALUE" patterns in text items
    for item in output["textItems"]:
        text = item.get("text", "").strip()
        if not text or len(text) < 3:
            continue

        # Check for "LABEL: VALUE" or "LABEL : VALUE" patterns
        inline_match = re.match(r'^(.+?)\s*[:：]\s*(.+)$', text)
        if inline_match:
            label_part = inline_match.group(1).strip()
            value_part = inline_match.group(2).strip()

            # Match the label part against known patterns
            for pattern, code, label_name in label_patterns:
                if re.search(pattern, label_part, re.IGNORECASE):
                    kv = {
                        "fieldCode": code,
                        "label": label_name,
                        "labelText": label_part,
                        "value": value_part,
                        "page": item.get("page", 1),
                        "confidence": 0.9,
                        "source": "inline",
                    }
                    # Use the item's bbox to estimate value position
                    if item.get("bbox"):
                        bbox = item["bbox"]
                        label_ratio = len(label_part) / max(len(text), 1)
                        kv["labelBbox"] = {
                            "x": bbox["x"],
                            "y": bbox["y"],
                            "width": bbox["width"] * label_ratio,
                            "height": bbox["height"]
                        }
                        kv["valueBbox"] = {
                            "x": bbox["x"] + bbox["width"] * label_ratio,
                            "y": bbox["y"],
                            "width": bbox["width"] * (1 - label_ratio),
                            "height": bbox["height"]
                        }
                    output["keyValuePairs"].append(kv)
                    break

    # Strategy 2: Look for adjacent items (label on left, value on right or next line)
    items_with_bbox = [item for item in output["textItems"] if item.get("bbox")]
    items_with_bbox.sort(key=lambda x: (x.get("page", 1), x["bbox"]["y"], x["bbox"]["x"]))

    for i, item in enumerate(items_with_bbox):
        text = item.get("text", "").strip()
        if not text:
            continue

        for pattern, code, label_name in label_patterns:
            if re.search(pattern, text, re.IGNORECASE):
                # Already found as inline? Skip
                if any(kv["fieldCode"] == code and kv["page"] == item.get("page", 1)
                       for kv in output["keyValuePairs"]):
                    continue

                # Look for value in the next item(s) on the same page
                for j in range(i + 1, min(i + 3, len(items_with_bbox))):
                    next_item = items_with_bbox[j]
                    if next_item.get("page") != item.get("page"):
                        break

                    next_text = next_item.get("text", "").strip()
                    if not next_text:
                        continue

                    # Check if next item is NOT a label (it should be a value)
                    is_label = any(re.search(p, next_text, re.IGNORECASE)
                                   for p, _, _ in label_patterns)
                    if not is_label and len(next_text) > 0:
                        kv = {
                            "fieldCode": code,
                            "label": label_name,
                            "labelText": text,
                            "value": next_text,
                            "page": item.get("page", 1),
                            "confidence": 0.85,
                            "source": "adjacent",
                            "labelBbox": item.get("bbox"),
                            "valueBbox": next_item.get("bbox"),
                        }
                        output["keyValuePairs"].append(kv)
                        break
                break


# ═══════════════════════════════════════════════════════════════
#  API Routes
# ═══════════════════════════════════════════════════════════════

@app.route('/api/docling/health', methods=['GET'])
def health():
    """Health check endpoint."""
    return jsonify({
        "status": "online",
        "engine": "docling",
        "version": "2.88.0" if DOCLING_AVAILABLE else "mock",
        "docling_available": DOCLING_AVAILABLE,
    })


@app.route('/api/docling/parse', methods=['POST'])
def parse_document():
    """
    Parse a PDF document using Docling.
    
    Accepts: multipart/form-data with a 'file' field
    Returns: Structured JSON with text items, tables, key-value pairs
    """
    if 'file' not in request.files:
        return jsonify({"status": "error", "message": "No file uploaded"}), 400

    file = request.files['file']
    if not file.filename:
        return jsonify({"status": "error", "message": "Empty filename"}), 400

    # Save file temporarily
    ext = file.filename.rsplit('.', 1)[-1].lower() if '.' in file.filename else 'pdf'
    temp_path = os.path.join(UPLOAD_FOLDER, f"upload_{os.getpid()}.{ext}")

    try:
        file.save(temp_path)
        print(f"\n═══ Docling Parse: {file.filename} ({os.path.getsize(temp_path)} bytes) ═══")

        if not DOCLING_AVAILABLE:
            # Return mock data for testing when Docling is not installed
            return jsonify(_generate_mock_data(file.filename))

        # ── Run Docling conversion ──
        try:
            from docling.document_converter import PdfFormatOption
            pipeline_options = PdfPipelineOptions()
            pipeline_options.do_ocr = True 
            
            converter = DocumentConverter(
                allowed_formats=[InputFormat.PDF, InputFormat.IMAGE],
                format_options={
                    InputFormat.PDF: PdfFormatOption(pipeline_options=pipeline_options)
                }
            )
        except Exception as e:
            print(f"  ⚠ Could not initialize advanced Docling options (OCR/Image): {e}")
            converter = DocumentConverter()
            
        result = converter.convert(temp_path)

        print(f"  ✓ Docling conversion complete")

        # Extract structured data
        structured = extract_structured_data(result)
        structured["filename"] = file.filename

        print(f"  ✓ Extracted: {len(structured['textItems'])} text items, "
              f"{len(structured['tables'])} tables, "
              f"{len(structured['keyValuePairs'])} key-value pairs")
        print(f"  ✓ Document type: {structured['documentType']}")
        print(f"  ✓ Pages: {structured['totalPages']}")

        return jsonify(structured)

    except Exception as e:
        print(f"  ✗ Error: {e}")
        traceback.print_exc()
        return jsonify({
            "status": "error",
            "message": str(e),
            "traceback": traceback.format_exc()
        }), 500

    finally:
        # Cleanup
        if os.path.exists(temp_path):
            os.remove(temp_path)


def _generate_mock_data(filename):
    """Generate mock structured data when Docling is not available."""
    return {
        "status": "success",
        "filename": filename,
        "documentType": "Sales Contract",
        "totalPages": 1,
        "mock": True,
        "textItems": [
            {"index": 0, "text": "SALES CONTRACT", "type": "heading", "page": 1,
             "bbox": {"x": 150, "y": 30, "width": 300, "height": 20}},
            {"index": 1, "text": "NO: SC-MY250922", "type": "text", "page": 1,
             "bbox": {"x": 350, "y": 55, "width": 180, "height": 14}},
            {"index": 2, "text": "DATE: 22/09/2025", "type": "text", "page": 1,
             "bbox": {"x": 350, "y": 75, "width": 180, "height": 14}},
            {"index": 3, "text": "BUYER:", "type": "text", "page": 1,
             "bbox": {"x": 50, "y": 100, "width": 60, "height": 14}},
            {"index": 4, "text": "THIEN LOC VIET NAM IMPORT EXPORT CO.,LTD", "type": "text", "page": 1,
             "bbox": {"x": 120, "y": 100, "width": 350, "height": 14}},
            {"index": 5, "text": "SELLER:", "type": "text", "page": 1,
             "bbox": {"x": 50, "y": 145, "width": 60, "height": 14}},
            {"index": 6, "text": "NINGBO MING YUE GLOBAL TRADING CO., LTD", "type": "text", "page": 1,
             "bbox": {"x": 120, "y": 145, "width": 350, "height": 14}},
        ],
        "tables": [],
        "keyValuePairs": [
            {
                "fieldCode": "SO_CHUNG_TU", "label": "Document Number",
                "labelText": "NO", "value": "SC-MY250922",
                "page": 1, "confidence": 0.95, "source": "inline",
                "labelBbox": {"x": 350, "y": 55, "width": 40, "height": 14},
                "valueBbox": {"x": 390, "y": 55, "width": 140, "height": 14}
            },
            {
                "fieldCode": "NGAY", "label": "Date",
                "labelText": "DATE", "value": "22/09/2025",
                "page": 1, "confidence": 0.95, "source": "inline",
                "labelBbox": {"x": 350, "y": 75, "width": 50, "height": 14},
                "valueBbox": {"x": 405, "y": 75, "width": 125, "height": 14}
            },
            {
                "fieldCode": "TEN_NGUOI_MUA", "label": "Buyer",
                "labelText": "BUYER", "value": "THIEN LOC VIET NAM IMPORT EXPORT CO.,LTD",
                "page": 1, "confidence": 0.9, "source": "adjacent",
                "labelBbox": {"x": 50, "y": 100, "width": 60, "height": 14},
                "valueBbox": {"x": 120, "y": 100, "width": 350, "height": 14}
            },
            {
                "fieldCode": "TEN_NGUOI_BAN", "label": "Seller",
                "labelText": "SELLER", "value": "NINGBO MING YUE GLOBAL TRADING CO., LTD",
                "page": 1, "confidence": 0.9, "source": "adjacent",
                "labelBbox": {"x": 50, "y": 145, "width": 60, "height": 14},
                "valueBbox": {"x": 120, "y": 145, "width": 350, "height": 14}
            },
        ],
        "pages": [{
            "pageNumber": 1,
            "itemCount": 7,
        }],
        "fullText": "SALES CONTRACT\nNO: SC-MY250922\nDATE: 22/09/2025\nBUYER: THIEN LOC VIET NAM...\nSELLER: NINGBO MING YUE..."
    }


# ═══════════════════════════════════════════════════════════════
#  Main
# ═══════════════════════════════════════════════════════════════

if __name__ == '__main__':
    print("═══════════════════════════════════════════════════════")
    print("  Docling Document Analysis Server")
    print(f"  Docling available: {'✓' if DOCLING_AVAILABLE else '✗ (mock mode)'}")
    print(f"  Listening on: http://localhost:{PORT}")
    print(f"  Upload folder: {UPLOAD_FOLDER}")
    print("═══════════════════════════════════════════════════════")
    app.run(host='0.0.0.0', port=PORT, debug=True)

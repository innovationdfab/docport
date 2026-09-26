from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from pypdf import PdfReader
from io import BytesIO

import re

from excel_generator import generate_excel


# =========================================================
# APP
# =========================================================

app = FastAPI(
    title="ExportFlow API",
    version="3.0.0"
)


# =========================================================
# CORS
# =========================================================

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",
        "http://127.0.0.1:5173",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# =========================================================
# HEALTH CHECK
# =========================================================

@app.get("/")
def health_check():
    return {
        "status": "ok",
        "service": "ExportFlow API",
        "parser": "whole-document",
        "excel_generator": True
    }


# =========================================================
# HELPERS
# =========================================================

def clean(value):
    if value is None:
        return None

    value = str(value).strip()
    value = re.sub(r"[ \t]+", " ", value)

    return value if value else None


def find_first(patterns, text):
    for pattern in patterns:
        match = re.search(
            pattern,
            text,
            re.IGNORECASE | re.MULTILINE
        )

        if match:
            return clean(match.group(1))

    return None


def extract_section(
    text,
    start_pattern,
    end_patterns
):
    start = re.search(
        start_pattern,
        text,
        re.IGNORECASE
    )

    if not start:
        return None

    section_start = start.end()
    section_end = len(text)

    for pattern in end_patterns:
        end = re.search(
            pattern,
            text[section_start:],
            re.IGNORECASE
        )

        if end:
            possible_end = (
                section_start
                + end.start()
            )

            if possible_end < section_end:
                section_end = possible_end

    return clean(
        text[
            section_start:
            section_end
        ]
    )


def first_meaningful_line(section):
    if not section:
        return None

    for line in section.splitlines():
        line = clean(line)

        if line:
            return line

    return None


# =========================================================
# PDF EXTRACTION
# =========================================================

@app.post("/api/invoices/extract")
async def extract_invoice(
    file: UploadFile = File(...)
):

    # -----------------------------------------------------
    # VALIDATION
    # -----------------------------------------------------

    if not file.filename:
        raise HTTPException(
            status_code=400,
            detail="No PDF selected"
        )

    if not file.filename.lower().endswith(
        ".pdf"
    ):
        raise HTTPException(
            status_code=400,
            detail="Only PDF files are supported"
        )

    contents = await file.read()

    if not contents:
        raise HTTPException(
            status_code=400,
            detail="Uploaded PDF is empty"
        )

    # -----------------------------------------------------
    # READ PDF
    # -----------------------------------------------------

    try:
        reader = PdfReader(
            BytesIO(contents)
        )

        pages = []

        for page_number, page in enumerate(
            reader.pages
        ):
            try:
                page_text = (
                    page.extract_text()
                    or ""
                )

            except Exception:
                page_text = ""

            pages.append({
                "page": page_number + 1,
                "text": page_text
            })

        # -------------------------------------------------
        # WHOLE DOCUMENT TEXT
        # -------------------------------------------------

        full_text = "\n".join(
            page["text"]
            for page in pages
        )

        full_text = full_text.replace(
            "\xa0",
            " "
        )

        if not full_text.strip():
            return {
                "success": False,
                "requires_ocr": True,
                "message": (
                    "No readable text found. "
                    "PDF requires OCR."
                ),
                "filename": file.filename,
                "page_count": len(pages),
                "extracted": {},
                "missing_fields": [],
                "raw_text": ""
            }

        # =================================================
        # 1. INVOICE DETAILS
        # =================================================

        invoice_number = find_first(
            [
                (
                    r"^\s*Invoice\s+Number"
                    r"\s*:\s*([^\r\n]+)"
                ),
                (
                    r"^\s*Invoice\s+No\.?"
                    r"\s*:\s*([^\r\n]+)"
                ),
                (
                    r"^\s*Invoice\s*#"
                    r"\s*:\s*([^\r\n]+)"
                )
            ],
            full_text
        )

        invoice_date = find_first(
            [
                (
                    r"^\s*Invoice\s+Issue"
                    r"\s+Date\s*:\s*"
                    r"([^\r\n]+)"
                ),
                (
                    r"^\s*Invoice\s+Date"
                    r"\s*:\s*([^\r\n]+)"
                )
            ],
            full_text
        )

        po_number = find_first(
            [
                (
                    r"^\s*Purchase\s+Order"
                    r"\s*:\s*([^\r\n]+)"
                ),
                (
                    r"^\s*PO\s*"
                    r"(?:No\.?|Number)?"
                    r"\s*:\s*([^\r\n]+)"
                ),
                (
                    r"^\s*Order\s+No\.?"
                    r"\s*:\s*([^\r\n]+)"
                )
            ],
            full_text
        )

        carrier = find_first(
            [
                (
                    r"^\s*Carrier\s*:"
                    r"\s*([^\r\n]+)"
                )
            ],
            full_text
        )

        # =================================================
        # 2. EXPORTER / IMPORTER /
        #    CONSIGNEE / SOLD TO
        # =================================================

        exporter_section = extract_section(
            full_text,
            (
                r"Shipper\s*/\s*Exporter"
                r"\s+of\s+record\s*:"
            ),
            [
                r"Importer\s+of\s+record\s*:",
                (
                    r"Ship\s+to\s*/"
                    r"\s*Consignee\s*:"
                )
            ]
        )

        importer_section = extract_section(
            full_text,
            r"Importer\s+of\s+record\s*:",
            [
                (
                    r"Ship\s+to\s*/"
                    r"\s*Consignee\s*:"
                ),
                r"Sold\s+to\s*:"
            ]
        )

        consignee_section = extract_section(
            full_text,
            (
                r"Ship\s+to\s*/"
                r"\s*Consignee\s*:"
            ),
            [
                r"Sold\s+to\s*:",
                r"Invoice\s+details\s*:"
            ]
        )

        sold_to_section = extract_section(
            full_text,
            r"Sold\s+to\s*:",
            [
                r"Invoice\s+details\s*:"
            ]
        )

        exporter_name = (
            first_meaningful_line(
                exporter_section
            )
        )

        importer_name = (
            first_meaningful_line(
                importer_section
            )
        )

        consignee_name = (
            first_meaningful_line(
                consignee_section
            )
        )

        sold_to_name = (
            first_meaningful_line(
                sold_to_section
            )
        )

        # =================================================
        # 3. PRODUCT / MATERIAL
        # =================================================

        part_numbers = re.findall(
            (
                r"Part\s*#\s*"
                r"([A-Za-z0-9._/\-]+)"
            ),
            full_text,
            re.IGNORECASE
        )

        materials = re.findall(
            (
                r"Material\s*:\s*"
                r"([^\r\n]+"
                r"(?:\n(?!\s*\d)"
                r"[^\r\n]+)?)"
            ),
            full_text,
            re.IGNORECASE
        )

        part_number = (
            clean(part_numbers[0])
            if part_numbers
            else None
        )

        material = (
            clean(materials[0])
            if materials
            else None
        )

        # =================================================
        # 4. PRODUCT TABLE ROWS
        #
        # Example:
        # 1.71 9817850100 India
        # 3 $59.93 $179.79
        # =================================================

        line_items = []

        row_pattern = re.compile(
            r"(?m)^\s*"
            r"(\d+(?:\.\d+)?)\s+"
            r"(\d{6,12})\s+"
            r"([A-Za-z][A-Za-z\s]*?)\s+"
            r"(\d+)\s+"
            r"([$€£₹]?)\s*"
            r"([\d,]+(?:\.\d+)?)\s+"
            r"([$€£₹]?)\s*"
            r"([\d,]+(?:\.\d+)?)\s*$"
        )

        for index, match in enumerate(
            row_pattern.finditer(
                full_text
            )
        ):

            symbol = (
                match.group(5)
                or match.group(7)
            )

            item = {
                "net_weight_kg":
                    clean(match.group(1)),

                "hs_code":
                    clean(match.group(2)),

                "country_of_origin":
                    clean(match.group(3)),

                "quantity":
                    clean(match.group(4)),

                "currency_symbol":
                    clean(symbol),

                "unit_price":
                    clean(match.group(6)),

                "position_price":
                    clean(match.group(8))
            }

            if index < len(part_numbers):
                item["part_number"] = clean(
                    part_numbers[index]
                )
            else:
                item["part_number"] = None

            if index < len(materials):
                item["material"] = clean(
                    materials[index]
                )
            else:
                item["material"] = None

            line_items.append(item)

        # =================================================
        # 5. PACKAGING
        #
        # Example:
        # Packaging:
        # 35x15x10 cm, 1.95 kg
        # =================================================

        packaging_match = re.search(
            (
                r"Packaging\s*:\s*"
                r"([0-9.]+\s*[xX×]\s*"
                r"[0-9.]+\s*[xX×]\s*"
                r"[0-9.]+\s*cm)"
                r"\s*,\s*"
                r"([0-9.]+)\s*kg"
            ),
            full_text,
            re.IGNORECASE
        )

        package_dimensions = None
        package_weight_kg = None

        if packaging_match:
            package_dimensions = clean(
                packaging_match.group(1)
            )

            package_weight_kg = clean(
                packaging_match.group(2)
            )

        # =================================================
        # 6. TOTAL QUANTITY /
        #    TOTAL AMOUNT
        # =================================================

        totals_match = re.search(
            (
                r"Q-ty\s*:\s*(\d+)\s+"
                r"([$€£₹]?)\s*"
                r"([\d,]+(?:\.\d+)?)"
            ),
            full_text,
            re.IGNORECASE
        )

        total_quantity = None
        total_amount = None
        total_symbol = None

        if totals_match:
            total_quantity = clean(
                totals_match.group(1)
            )

            total_symbol = clean(
                totals_match.group(2)
            )

            total_amount = clean(
                totals_match.group(3)
            )

        # ---------------------------------------------
        # FALLBACK QUANTITY FROM TABLE
        # ---------------------------------------------

        if (
            not total_quantity
            and line_items
        ):
            try:
                total_quantity = str(
                    sum(
                        int(
                            item["quantity"]
                        )
                        for item in line_items
                        if item[
                            "quantity"
                        ]
                    )
                )

            except Exception:
                pass

        # ---------------------------------------------
        # FALLBACK TOTAL FROM TABLE
        # ---------------------------------------------

        if (
            not total_amount
            and line_items
        ):
            try:
                total_amount = str(
                    round(
                        sum(
                            float(
                                item[
                                    "position_price"
                                ].replace(
                                    ",",
                                    ""
                                )
                            )
                            for item
                            in line_items
                            if item[
                                "position_price"
                            ]
                        ),
                        2
                    )
                )

            except Exception:
                pass

        # =================================================
        # 7. CURRENCY
        # =================================================

        currency = None

        if re.search(
            r"\bUSD\b",
            full_text,
            re.IGNORECASE
        ):
            currency = "USD"

        elif re.search(
            r"\bINR\b",
            full_text,
            re.IGNORECASE
        ):
            currency = "INR"

        elif re.search(
            r"\bEUR\b",
            full_text,
            re.IGNORECASE
        ):
            currency = "EUR"

        elif "$" in full_text:
            currency = "USD"

        elif "₹" in full_text:
            currency = "INR"

        elif "€" in full_text:
            currency = "EUR"

        # =================================================
        # 8. NET WEIGHT
        # =================================================

        net_weight_kg = None

        if line_items:
            try:
                calculated_weight = sum(
                    float(
                        item[
                            "net_weight_kg"
                        ]
                    )
                    for item
                    in line_items
                    if item[
                        "net_weight_kg"
                    ]
                )

                net_weight_kg = (
                    f"{calculated_weight:.3f}"
                    .rstrip("0")
                    .rstrip(".")
                )

            except Exception:
                net_weight_kg = None

        # =================================================
        # 9. COMPANY VALUES FOUND
        #    DIRECTLY IN PDF
        # =================================================

        gstin = find_first(
            [
                (
                    r"\bGSTIN\s*"
                    r"(?:ID)?\s*"
                    r"[:\-]\s*"
                    r"([0-9A-Z]{15})"
                )
            ],
            full_text
        )

        iec = find_first(
            [
                (
                    r"\bIEC\s*"
                    r"(?:No\.?|Number)?"
                    r"\s*[:\-]\s*"
                    r"([0-9A-Z]{10})"
                )
            ],
            full_text
        )

        # =================================================
        # 10. BANK DETAILS
        # =================================================

        bank_name = find_first(
            [
                (
                    r"^\s*Bank\s*:"
                    r"\s*([^\r\n]+)"
                )
            ],
            full_text
        )

        bank_code = find_first(
            [
                (
                    r"^\s*Bank\s+code\s*:"
                    r"\s*([^\r\n]+)"
                )
            ],
            full_text
        )

        bank_account = find_first(
            [
                (
                    r"^\s*Account\s*#\s*:"
                    r"\s*([^\r\n]+)"
                ),
                (
                    r"^\s*Account\s+Number"
                    r"\s*:\s*([^\r\n]+)"
                )
            ],
            full_text
        )

        # =================================================
        # 11. SHIPPING /
        #     CUSTOMS VALUES
        # =================================================

        awb_number = find_first(
            [
                (
                    r"\bAWB\s*"
                    r"(?:No\.?|Number)?"
                    r"\s*[:\-]\s*"
                    r"([A-Z0-9\-]+)"
                ),
                (
                    r"\bAir\s*Waybill\s*"
                    r"(?:No\.?|Number)?"
                    r"\s*[:\-]\s*"
                    r"([A-Z0-9\-]+)"
                )
            ],
            full_text
        )

        shipping_bill_number = find_first(
            [
                (
                    r"Shipping\s+Bill\s*"
                    r"(?:No\.?|Number)?"
                    r"\s*[:\-]\s*"
                    r"([A-Z0-9\-]+)"
                )
            ],
            full_text
        )

        shipping_bill_date = find_first(
            [
                (
                    r"Shipping\s+Bill.*?"
                    r"Date\s*[:\-]\s*"
                    r"([0-9./\-]+)"
                )
            ],
            full_text
        )

        exchange_rate = find_first(
            [
                (
                    r"Exchange\s+Rate"
                    r"\s*[:\-]\s*"
                    r"([0-9.]+)"
                )
            ],
            full_text
        )

        # =================================================
        # 12. BUILD NORMALIZED DATA
        # =================================================

        extracted = {
            "invoice": {
                "number":
                    invoice_number,

                "date":
                    invoice_date,

                "po_number":
                    po_number,

                "carrier":
                    carrier
            },

            "exporter": {
                "name":
                    exporter_name,

                "raw_section":
                    exporter_section
            },

            "importer": {
                "name":
                    importer_name,

                "raw_section":
                    importer_section
            },

            "consignee": {
                "name":
                    consignee_name,

                "raw_section":
                    consignee_section
            },

            "sold_to": {
                "name":
                    sold_to_name,

                "raw_section":
                    sold_to_section
            },

            "product": {
                "part_number":
                    part_number,

                "material":
                    material,

                "line_items":
                    line_items
            },

            "shipment": {
                "net_weight_kg":
                    net_weight_kg,

                "package_weight_kg":
                    package_weight_kg,

                "package_dimensions":
                    package_dimensions,

                "total_quantity":
                    total_quantity,

                "currency":
                    currency,

                "total_amount":
                    total_amount,

                "awb_number":
                    awb_number,

                "shipping_bill_number":
                    shipping_bill_number,

                "shipping_bill_date":
                    shipping_bill_date,

                "exchange_rate":
                    exchange_rate
            },

            "company": {
                "gstin":
                    gstin,

                "iec":
                    iec
            },

            "bank": {
                "name":
                    bank_name,

                "code":
                    bank_code,

                "account_number":
                    bank_account
            }
        }

        # =================================================
        # 13. MISSING FIELD REPORT
        # =================================================

        missing_fields = []

        required_checks = {
            "invoice.number":
                invoice_number,

            "invoice.date":
                invoice_date,

            "invoice.po_number":
                po_number,

            "invoice.carrier":
                carrier,

            "shipment.net_weight_kg":
                net_weight_kg,

            "shipment.package_weight_kg":
                package_weight_kg,

            "shipment.package_dimensions":
                package_dimensions,

            "shipment.total_quantity":
                total_quantity,

            "shipment.currency":
                currency,

            "shipment.total_amount":
                total_amount,

            "shipping.awb_number":
                awb_number,

            "shipping.shipping_bill_number":
                shipping_bill_number,

            "shipping.shipping_bill_date":
                shipping_bill_date,

            "shipping.exchange_rate":
                exchange_rate
        }

        for (
            field_name,
            value
        ) in required_checks.items():

            if value is None:
                missing_fields.append(
                    field_name
                )

        # =================================================
        # FINAL EXTRACTION RESPONSE
        # =================================================

        return {
            "success": True,
            "requires_ocr": False,
            "message": (
                "Whole PDF extracted "
                "successfully"
            ),
            "filename": file.filename,
            "page_count": len(pages),
            "extracted": extracted,
            "missing_fields":
                missing_fields,
            "raw_text": full_text
        }

    except Exception as exc:

        raise HTTPException(
            status_code=500,
            detail=(
                "Unable to process PDF: "
                f"{str(exc)}"
            )
        )


# =========================================================
# EXCEL GENERATION
# =========================================================

@app.post("/api/documents/generate")
async def generate_documents(
    payload: dict
):

    try:
        # Generate a COPY of the master workbook.
        #
        # excel_generator.py changes only the
        # approved dynamic fields.
        #
        # All standard workbook content remains
        # untouched.

        output_path = generate_excel(
            payload
        )

        return FileResponse(
            path=str(output_path),

            media_type=(
                "application/"
                "vnd.openxmlformats-"
                "officedocument."
                "spreadsheetml.sheet"
            ),

            filename=output_path.name
        )

    except FileNotFoundError as exc:

        raise HTTPException(
            status_code=500,
            detail=str(exc)
        )

    except ValueError as exc:

        raise HTTPException(
            status_code=400,
            detail=str(exc)
        )

    except Exception as exc:

        raise HTTPException(
            status_code=500,
            detail=(
                "Unable to generate Excel: "
                f"{str(exc)}"
            )
        )


   # =========================================================
# DIRECT PDF -> EXCEL GENERATION
# =========================================================

@app.post("/api/documents/generate-from-pdf")
async def generate_documents_from_pdf(
    file: UploadFile = File(...)
):
    try:
        extracted_payload = await extract_invoice(file)

        if not extracted_payload.get("success"):
            raise HTTPException(
                status_code=400,
                detail=(
                    extracted_payload.get("message")
                    or "PDF extraction failed"
                )
            )

        output_path = generate_excel(
            extracted_payload
        )

        return FileResponse(
            path=str(output_path),
            media_type=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            filename=output_path.name,
        )

    except HTTPException:
        raise

    except FileNotFoundError as exc:
        raise HTTPException(
            status_code=500,
            detail=str(exc),
        )

    except ValueError as exc:
        raise HTTPException(
            status_code=400,
            detail=str(exc),
        )

    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=(
                "Unable to generate Excel from PDF: "
                f"{str(exc)}"
            ),
        )
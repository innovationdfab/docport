import asyncio
import os
import sys
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from google import genai
from google.genai import types
from pydantic import BaseModel, Field
from starlette.background import BackgroundTask


# ============================================================
# PROJECT PATHS
# ============================================================

ROOT_DIR = Path(__file__).resolve().parents[1]
BACKEND_DIR = ROOT_DIR / "backend"

# Load local environment variables.
load_dotenv(ROOT_DIR / ".env.local")

# Allow this API to reuse the existing DocPort backend.
if str(BACKEND_DIR) not in sys.path:
    sys.path.insert(0, str(BACKEND_DIR))

from excel_generator import generate_excel, MAX_PRODUCTS


# ============================================================
# FASTAPI
# ============================================================

app = FastAPI(
    title="DocPort API",
    version="1.0.0",
)


# ============================================================
# CORS
#
# Allows the Vite development website running on port 5173
# to communicate with this API on port 8001.
# ============================================================

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


# ============================================================
# STRUCTURED GEMINI RESPONSE MODELS
# ============================================================


class InvoiceData(BaseModel):
    number: Optional[str] = None
    date: Optional[str] = None
    order_date: Optional[str] = None
    po_number: Optional[str] = None


class PartyData(BaseModel):
    name: Optional[str] = None

    contact: Optional[str] = Field(
        default=None,
        description=(
            "Actual named contact person only. "
            "Do not copy a company name into this field "
            "unless explicitly identified as the contact."
        ),
    )

    address: Optional[str] = None

    raw_section: Optional[str] = Field(
        default=None,
        description=(
            "Complete useful visible address/contact block "
            "for this party."
        ),
    )

    email: Optional[str] = None
    phone: Optional[str] = None
    country: Optional[str] = None


class LineItem(BaseModel):
    part_number: Optional[str] = None

    description: Optional[str] = Field(
        default=None,
        description=(
            "Useful product description exactly once. "
            "Do not duplicate the same filename."
        ),
    )

    material: Optional[str] = Field(
        default=None,
        description=(
            "Material grade exactly as printed. "
            "Preserve numbers, punctuation and spacing."
        ),
    )

    net_weight_kg: Optional[float] = None

    hs_code: Optional[str] = None

    india_hsn: Optional[str] = None

    india_hsn_code: Optional[str] = None

    country_of_origin: Optional[str] = None

    quantity: Optional[int] = None

    currency_symbol: Optional[str] = None

    unit_price: Optional[float] = None

    position_price: Optional[float] = None


class ProductData(BaseModel):
    line_items: list[LineItem] = Field(
        default_factory=list
    )


class ShipmentData(BaseModel):
    package_dimensions: Optional[str] = None

    package_weight_kg: Optional[float] = None

    net_weight_kg: Optional[float] = None

    total_quantity: Optional[int] = None

    total_amount: Optional[float] = None

    currency: Optional[str] = None

    exchange_rate: Optional[float] = None

    awb_number: Optional[str] = None

    shipping_bill_number: Optional[str] = None

    shipping_bill_date: Optional[str] = None

    rodtep_yes_no: Optional[str] = None

    rodtep_total_line_items: Optional[int] = None

    rodtep_claimed_line_items: Optional[int] = None

    freight: Optional[float] = None

    insurance: Optional[float] = None

    commission: Optional[float] = None

    discount: Optional[float] = None

    packing_charges: Optional[float] = None

    package_count: Optional[int] = None

    state_origin: Optional[str] = None

    district_origin: Optional[str] = None


class CompanyData(BaseModel):
    iec: Optional[str] = None
    gstin: Optional[str] = None


class BankData(BaseModel):
    code: Optional[str] = Field(
        default=None,
        description=(
            "Bank AD Code only. "
            "Never put a SWIFT or BIC code here."
        ),
    )

    account_number: Optional[str] = None

    ifsc: Optional[str] = Field(
        default=None,
        description="Indian IFSC code only.",
    )

    swift_code: Optional[str] = Field(
        default=None,
        description="SWIFT/BIC code when explicitly present.",
    )


class ExtractedData(BaseModel):
    invoice: InvoiceData = Field(
        default_factory=InvoiceData
    )

    exporter: PartyData = Field(
        default_factory=PartyData
    )

    importer: PartyData = Field(
        default_factory=PartyData
    )

    consignee: PartyData = Field(
        default_factory=PartyData
    )

    sold_to: PartyData = Field(
        default_factory=PartyData
    )

    product: ProductData = Field(
        default_factory=ProductData
    )

    shipment: ShipmentData = Field(
        default_factory=ShipmentData
    )

    company: CompanyData = Field(
        default_factory=CompanyData
    )

    bank: BankData = Field(
        default_factory=BankData
    )


# ============================================================
# GEMINI EXTRACTION PROMPT
# ============================================================

EXTRACTION_PROMPT = """
You are the PDF extraction engine for DocPort.

The attached PDF is a commercial/export invoice.

Read the ENTIRE PDF carefully before returning data.

Extract only information that is actually visible in the document.

Never invent, infer, assume, calculate or manufacture missing values.

If a field is absent, unclear or uncertain, return null.

Accuracy is more important than filling every field.


============================================================
INVOICE
============================================================

Extract:

- commercial invoice number
- invoice date
- order date only when explicitly shown
- purchase order / PO number

Preserve identifiers exactly as printed.

When a date is clearly identifiable, normalize it to:

DD/MM/YYYY


============================================================
PARTIES
============================================================

Keep all parties separate:

- exporter / shipper
- importer
- consignee / ship-to
- sold-to

Do not mix their names, addresses, contacts, email addresses or
phone numbers.

For raw_section, preserve the complete useful visible block for
that party.

For contact, return an actual contact person only.

Do not copy the company name into the contact field unless the
document explicitly identifies it as the contact.


============================================================
PRODUCTS
============================================================

Extract every visible product line in the same order as the PDF.

Do not merge different products.

For every product extract separately:

- part_number
- description
- material
- net_weight_kg
- hs_code
- india_hsn
- country_of_origin
- quantity
- currency_symbol
- unit_price
- position_price


DESCRIPTION:

Do not repeat the same product name or filename twice.

Example:

Wrong:
Ultrasonic_toolchangerplate A.1.stp, prototype,
Ultrasonic_toolchangerplate A.1.stp

Correct:
Ultrasonic_toolchangerplate A.1.stp, prototype


MATERIAL:

Preserve material grade numbers exactly.

Example:

3.3211

must not become:

3.3 211


WEIGHT:

product.line_items[].net_weight_kg is the net weight of the
individual product line.

shipment.package_weight_kg is the gross/package shipment weight.

shipment.net_weight_kg is the total shipment net weight only when
explicitly printed.

Do not confuse these weights.


PRICES:

unit_price is price per unit.

position_price is the printed total amount for that product line.

Do not calculate a missing position_price.


============================================================
SHIPMENT
============================================================

Extract when explicitly available:

- package dimensions
- package/gross weight
- overall net shipment weight
- total quantity
- total amount
- currency
- exchange rate
- AWB number
- Shipping Bill number
- Shipping Bill date
- freight
- insurance
- commission
- discount
- packing charges
- package count
- state of origin
- district of origin
- RoDTEP information

Never fabricate these fields.

If exchange rate is not printed, return null.

If AWB is not printed, return null.

If Shipping Bill information is not printed, return null.


============================================================
BANK INFORMATION
============================================================

Be very careful with bank codes.

bank.code means BANK AD CODE only.

A SWIFT/BIC code must NEVER be placed in bank.code.

Place SWIFT/BIC only in:

bank.swift_code

Place IFSC only in:

bank.ifsc

If AD Code is not explicitly identified in the PDF:

bank.code = null


============================================================
COMPANY IDENTIFIERS
============================================================

Extract IEC and GSTIN only when explicitly printed.

Do not infer them.


============================================================
EXCEL
============================================================

Do NOT choose Excel cells.

Do NOT generate an Excel workbook.

Do NOT change document formatting.

Do NOT decide where data belongs in Excel.

DocPort's existing deterministic field_mapping.py and
excel_generator.py control all Excel placement and formatting.
"""


# ============================================================
# GEMINI CLIENT
# ============================================================


def get_gemini_client():

    api_key = os.getenv("GEMINI_API_KEY")

    if not api_key:
        raise HTTPException(
            status_code=500,
            detail="GEMINI_API_KEY is not configured.",
        )

    return genai.Client(
        api_key=api_key
    )


# ============================================================
# MISSING FIELD HELPERS
# ============================================================


def is_missing(value):

    if value is None:
        return True

    if isinstance(value, str):
        return value.strip() == ""

    return False


def build_missing_fields(
    data: ExtractedData,
):

    missing = []

    important_fields = [
        (
            "invoice.number",
            data.invoice.number,
        ),
        (
            "invoice.date",
            data.invoice.date,
        ),
        (
            "invoice.po_number",
            data.invoice.po_number,
        ),
        (
            "consignee.name",
            data.consignee.name,
        ),
        (
            "consignee.raw_section",
            (
                data.consignee.raw_section
                or data.consignee.address
            ),
        ),
        (
            "shipment.package_dimensions",
            data.shipment.package_dimensions,
        ),
        (
            "shipment.package_weight_kg",
            data.shipment.package_weight_kg,
        ),
        (
            "shipment.exchange_rate",
            data.shipment.exchange_rate,
        ),
        (
            "shipment.awb_number",
            data.shipment.awb_number,
        ),
        (
            "shipment.shipping_bill_number",
            data.shipment.shipping_bill_number,
        ),
    ]

    for field_name, value in important_fields:

        if is_missing(value):
            missing.append(field_name)

    for index, item in enumerate(
        data.product.line_items
    ):

        product_fields = [
            (
                "part_number",
                item.part_number,
            ),
            (
                "description",
                item.description,
            ),
            (
                "material",
                item.material,
            ),
            (
                "net_weight_kg",
                item.net_weight_kg,
            ),
            (
                "hs_code",
                item.hs_code,
            ),
            (
                "country_of_origin",
                item.country_of_origin,
            ),
            (
                "quantity",
                item.quantity,
            ),
            (
                "unit_price",
                item.unit_price,
            ),
            (
                "position_price",
                item.position_price,
            ),
        ]

        for field_name, value in product_fields:

            if is_missing(value):

                missing.append(
                    (
                        f"product.line_items"
                        f"[{index}]."
                        f"{field_name}"
                    )
                )

    return missing


# ============================================================
# TEMPORARY GEMINI ERROR DETECTION
# ============================================================


def is_temporary_gemini_error(
    error: Exception,
):

    error_text = str(error).upper()

    temporary_markers = [
        "429",
        "500",
        "502",
        "503",
        "504",
        "RESOURCE_EXHAUSTED",
        "UNAVAILABLE",
        "HIGH DEMAND",
        "OVERLOADED",
        "TIMEOUT",
        "DEADLINE_EXCEEDED",
    ]

    return any(
        marker in error_text
        for marker in temporary_markers
    )


# ============================================================
# GEMINI EXTRACTION WITH RETRIES
# ============================================================


async def extract_with_gemini(
    pdf_bytes: bytes,
):

    client = get_gemini_client()

    # Uses the model configured in .env.local.
    #
    # Your current working .env.local should contain:
    #
    # GEMINI_MODEL=gemini-3.8-flash
    #
    model = os.getenv(
        "GEMINI_MODEL",
        "gemini-3.8-flash",
    )

    pdf_part = types.Part.from_bytes(
        data=pdf_bytes,
        mime_type="application/pdf",
    )

    # First attempt immediately.
    #
    # If Gemini temporarily returns 429/503/etc:
    #
    # attempt 2 -> wait 2 seconds
    # attempt 3 -> wait 4 seconds
    # attempt 4 -> wait 8 seconds
    #
    retry_delays = [
        0,
        2,
        4,
        8,
    ]

    last_error = None

    for attempt_number, delay in enumerate(
        retry_delays,
        start=1,
    ):

        if delay > 0:

            print(
                (
                    "Gemini temporarily unavailable. "
                    f"Retrying in {delay} seconds..."
                )
            )

            await asyncio.sleep(delay)

        try:

            print(
                (
                    "DocPort Gemini extraction "
                    f"attempt {attempt_number}/"
                    f"{len(retry_delays)} "
                    f"using {model}"
                )
            )

            response = await asyncio.to_thread(
                client.models.generate_content,

                model=model,

                contents=[
                    pdf_part,
                    EXTRACTION_PROMPT,
                ],

                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=ExtractedData,
                    temperature=0,
                ),
            )

            if response is None:
                raise ValueError(
                    "Gemini returned no response."
                )

            parsed = getattr(
                response,
                "parsed",
                None,
            )

            if isinstance(
                parsed,
                ExtractedData,
            ):
                return parsed

            if parsed is not None:

                return ExtractedData.model_validate(
                    parsed
                )

            response_text = getattr(
                response,
                "text",
                None,
            )

            if response_text:

                return ExtractedData.model_validate_json(
                    response_text
                )

            raise ValueError(
                "Gemini returned an empty structured response."
            )

        except Exception as exc:

            last_error = exc

            if not is_temporary_gemini_error(
                exc
            ):
                raise

            print(
                (
                    f"Gemini attempt "
                    f"{attempt_number} failed "
                    f"temporarily: {exc}"
                )
            )

    print(
        f"Gemini retries exhausted: {last_error}"
    )

    raise HTTPException(
        status_code=503,
        detail=(
            "Gemini is temporarily busy after "
            "multiple automatic attempts. "
            "Please retry the invoice shortly."
        ),
    )


# ============================================================
# API STATUS
# ============================================================


@app.get("/api")
def api_home():

    return {
        "service": "DocPort",
        "status": "online",
        "gemini_configured": bool(
            os.getenv("GEMINI_API_KEY")
        ),
        "gemini_model": os.getenv(
            "GEMINI_MODEL",
            "gemini-3.8-flash",
        ),
    }


@app.get("/api/health")
def api_health():

    return {
        "ok": True,
        "service": "DocPort",
        "gemini_configured": bool(
            os.getenv("GEMINI_API_KEY")
        ),
        "gemini_model": os.getenv(
            "GEMINI_MODEL",
            "gemini-3.8-flash",
        ),
    }


# ============================================================
# PDF EXTRACTION
# ============================================================


@app.post("/api/invoices/extract")
async def extract_invoice(
    file: UploadFile = File(...),
):

    filename = (
        file.filename
        or "invoice.pdf"
    )

    is_pdf = (
        file.content_type == "application/pdf"
        or filename.lower().endswith(".pdf")
    )

    if not is_pdf:

        raise HTTPException(
            status_code=400,
            detail=(
                "Only PDF commercial invoices "
                "are supported."
            ),
        )

    pdf_bytes = await file.read()

    if not pdf_bytes:

        raise HTTPException(
            status_code=400,
            detail="The uploaded PDF is empty.",
        )

    # 50 MB upload limit.
    if len(pdf_bytes) > (
        50 * 1024 * 1024
    ):

        raise HTTPException(
            status_code=413,
            detail=(
                "The PDF must be "
                "50 MB or smaller."
            ),
        )

    try:

        extracted = await extract_with_gemini(
            pdf_bytes
        )

        product_count = len(
            extracted.product.line_items
        )

        if product_count == 0:

            raise HTTPException(
                status_code=422,
                detail=(
                    "No product line items "
                    "were detected in the PDF."
                ),
            )

        # Your current approved master workbook has
        # a fixed number of product rows.
        if product_count > MAX_PRODUCTS:

            raise HTTPException(
                status_code=422,
                detail=(
                    f"The invoice contains "
                    f"{product_count} products, "
                    f"but the current approved "
                    f"Excel template supports "
                    f"{MAX_PRODUCTS} products."
                ),
            )

        return {
            "filename": filename,

            "extracted":
                extracted.model_dump(),

            "missing_fields":
                build_missing_fields(
                    extracted
                ),

            # Kept for compatibility with the existing
            # DocPort frontend/backend data structure.
            "raw_text": "",
        }

    except HTTPException:
        raise

    except Exception as exc:

        print(
            (
                "DocPort Gemini extraction "
                f"failed: {exc}"
            )
        )

        raise HTTPException(
            status_code=500,
            detail=(
                "Gemini PDF extraction failed: "
                f"{exc}"
            ),
        )


# ============================================================
# EXISTING DOCPORT EXCEL GENERATOR
#
# Gemini DOES NOT design or generate the workbook.
#
# Existing:
#
# backend/field_mapping.py
# backend/excel_generator.py
# backend/templates/ExportFlow_Master_Template.xlsx
#
# continue to control the exact Excel layout.
# ============================================================


@app.post("/api/documents/generate")
async def generate_document(
    payload: dict,
):

    try:

        output_path = Path(
            generate_excel(
                payload
            )
        )

        if not output_path.exists():

            raise FileNotFoundError(
                "The Excel workbook was not generated."
            )

        return FileResponse(
            path=str(output_path),

            media_type=(
                "application/"
                "vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),

            filename=output_path.name,

            background=BackgroundTask(
                lambda:
                output_path.unlink(
                    missing_ok=True
                )
            ),
        )

    except Exception as exc:

        print(
            (
                "DocPort Excel generation "
                f"failed: {exc}"
            )
        )

        raise HTTPException(
            status_code=500,
            detail=(
                "Excel generation failed: "
                f"{exc}"
            ),
        )
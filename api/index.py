import asyncio
import os
import sys
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from google import genai
from google.genai import types
from pydantic import BaseModel, Field
from starlette.background import BackgroundTask


# ============================================================
# DOCPORT PROJECT PATHS
# ============================================================

ROOT_DIR = Path(__file__).resolve().parents[1]
BACKEND_DIR = ROOT_DIR / "backend"

# Local development environment.
load_dotenv(ROOT_DIR / ".env.local")

# Allow this API to reuse the existing approved Excel code.
if str(BACKEND_DIR) not in sys.path:
    sys.path.insert(0, str(BACKEND_DIR))

from excel_generator import generate_excel, MAX_PRODUCTS


# ============================================================
# FASTAPI APP
# ============================================================

app = FastAPI(
    title="DocPort API",
    version="1.0.0",
)


# ============================================================
# GEMINI STRUCTURED OUTPUT
#
# IMPORTANT:
# These field names are kept compatible with the existing
# DocPort frontend and existing Excel generator.
# ============================================================


class InvoiceData(BaseModel):
    number: Optional[str] = Field(
        default=None,
        description="Commercial invoice number exactly as printed.",
    )

    date: Optional[str] = Field(
        default=None,
        description="Commercial invoice date in DD/MM/YYYY when clear.",
    )

    order_date: Optional[str] = Field(
        default=None,
        description="Order date only when explicitly shown.",
    )

    po_number: Optional[str] = Field(
        default=None,
        description="Purchase order number exactly as printed.",
    )


class PartyData(BaseModel):
    name: Optional[str] = None

    contact: Optional[str] = Field(
        default=None,
        description=(
            "Explicit contact person only. "
            "Do not copy the company name into contact."
        ),
    )

    address: Optional[str] = None

    raw_section: Optional[str] = Field(
        default=None,
        description=(
            "Complete visible address/contact block for this party."
        ),
    )

    email: Optional[str] = None
    phone: Optional[str] = None
    country: Optional[str] = None


class LineItem(BaseModel):
    part_number: Optional[str] = Field(
        default=None,
        description="Product or part number exactly as printed.",
    )

    description: Optional[str] = Field(
        default=None,
        description=(
            "Product description exactly once. "
            "Do not duplicate the same product/file name."
        ),
    )

    material: Optional[str] = Field(
        default=None,
        description=(
            "Material designation exactly as printed. "
            "Preserve grade punctuation and spacing accurately."
        ),
    )

    net_weight_kg: Optional[float] = Field(
        default=None,
        description="Net weight for this individual product line.",
    )

    hs_code: Optional[str] = Field(
        default=None,
        description=(
            "HS/HTS code printed for this individual product line."
        ),
    )

    india_hsn: Optional[str] = Field(
        default=None,
        description=(
            "Indian HSN only if explicitly printed in the PDF."
        ),
    )

    india_hsn_code: Optional[str] = None

    country_of_origin: Optional[str] = None

    quantity: Optional[int] = None

    currency_symbol: Optional[str] = None

    unit_price: Optional[float] = None

    position_price: Optional[float] = Field(
        default=None,
        description=(
            "Complete printed line total for this product."
        ),
    )


class ProductData(BaseModel):
    line_items: list[LineItem] = Field(
        default_factory=list
    )


class ShipmentData(BaseModel):
    package_dimensions: Optional[str] = Field(
        default=None,
        description=(
            "Package dimensions as visible in the PDF, "
            "for example 47x38x7 cm."
        ),
    )

    package_weight_kg: Optional[float] = Field(
        default=None,
        description=(
            "Gross/package shipment weight, not product net weight."
        ),
    )

    net_weight_kg: Optional[float] = Field(
        default=None,
        description=(
            "Overall shipment net weight only when explicitly printed."
        ),
    )

    total_quantity: Optional[int] = None
    total_amount: Optional[float] = None

    currency: Optional[str] = Field(
        default=None,
        description="ISO currency such as USD, EUR, GBP or INR.",
    )

    exchange_rate: Optional[float] = Field(
        default=None,
        description=(
            "Exchange rate only if explicitly shown. Never calculate it."
        ),
    )

    awb_number: Optional[str] = Field(
        default=None,
        description=(
            "Air Waybill/AWB number only if explicitly shown."
        ),
    )

    shipping_bill_number: Optional[str] = Field(
        default=None,
        description=(
            "Shipping Bill number only if explicitly shown."
        ),
    )

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
    iec: Optional[str] = Field(
        default=None,
        description="IEC only if explicitly printed.",
    )

    gstin: Optional[str] = Field(
        default=None,
        description="GSTIN only if explicitly printed.",
    )


class BankData(BaseModel):
    code: Optional[str] = Field(
        default=None,
        description=(
            "Bank AD Code only. "
            "DO NOT put SWIFT/BIC code in this field."
        ),
    )

    account_number: Optional[str] = None

    ifsc: Optional[str] = Field(
        default=None,
        description="Indian IFSC code only.",
    )

    swift_code: Optional[str] = Field(
        default=None,
        description="SWIFT/BIC code when present.",
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
# GEMINI EXTRACTION INSTRUCTIONS
# ============================================================

EXTRACTION_PROMPT = """
You are the document extraction engine for DocPort.

The attached document is a commercial/export invoice PDF.

Read the ENTIRE PDF carefully, including:

- header fields
- exporter/shipper
- importer
- consignee
- sold-to party
- product table
- packaging information
- totals
- bank information
- footer information

Return ONLY information matching the supplied structured schema.


============================================================
GENERAL ACCURACY RULES
============================================================

1. Extract only information actually visible in the PDF.

2. Never invent, infer, assume or manufacture a missing value.

3. If a field is absent or unclear, return null.

4. Accuracy is more important than filling every field.

5. Read the complete document before producing the result.

6. Preserve identifiers exactly as printed.

Examples:

- invoice numbers
- PO numbers
- part numbers
- HS/HTS codes
- IEC
- GSTIN
- bank account numbers
- AWB numbers
- shipping bill numbers


============================================================
DATES
============================================================

If a date can be confidently identified, normalize it to:

DD/MM/YYYY

Never assign one date to another field unless the PDF explicitly
identifies it for that field.


============================================================
PARTIES
============================================================

Keep these parties separate:

- exporter
- importer
- consignee / ship-to
- sold-to

Do not mix addresses or contacts between them.

For raw_section, preserve the complete useful visible block.

For contact:
return an actual named contact person only.

Do NOT copy a company name into the contact field unless the PDF
explicitly identifies that value as the contact.


============================================================
PRODUCT LINE ITEMS
============================================================

Extract EVERY product line in the same order shown in the PDF.

Never merge multiple products into one product.

Never omit a product.

For every product, keep these fields separate:

- part_number
- description
- material
- net_weight_kg
- hs_code
- india_hsn
- country_of_origin
- quantity
- unit_price
- position_price
- currency_symbol


DESCRIPTION RULE:

Do not repeat the same filename/product name twice.

For example, if the source effectively says:

Ultrasonic_toolchangerplate A.1.stp
prototype

the description should NOT become:

Ultrasonic_toolchangerplate A.1.stp, prototype,
Ultrasonic_toolchangerplate A.1.stp

It should contain the useful description only once.


MATERIAL RULE:

Read material grades carefully.

Preserve punctuation and digits exactly.

For example:

3.3211

must not become:

3.3 211


WEIGHT RULE:

product.line_items[].net_weight_kg
means the individual product line's net weight.

shipment.package_weight_kg
means package/gross shipment weight.

Do not confuse them.


PRICE RULE:

unit_price
means price per unit.

position_price
means the complete line-item amount printed in the invoice.

If quantity and unit price are printed but a position total is not
printed, return null for position_price.

Do not calculate missing financial fields.


============================================================
SHIPMENT
============================================================

package_dimensions should preserve the meaningful visible dimensions.

Example:

47x38x7 cm

package_weight_kg is package/gross weight.

shipment.net_weight_kg is only the overall shipment net weight when
the PDF explicitly provides one.

Do NOT manufacture:

- exchange rate
- AWB number
- Shipping Bill number
- Shipping Bill date
- freight
- insurance
- commission
- discount
- packing charges
- RoDTEP information


============================================================
CURRENCY
============================================================

currency should use ISO-style values:

USD
EUR
GBP
INR

currency_symbol examples:

$
€
£
₹


============================================================
BANK INFORMATION
============================================================

Be very careful with banking codes.

bank.code means BANK AD CODE only.

A SWIFT/BIC code must NEVER be placed in bank.code.

Put a SWIFT/BIC code only in:

bank.swift_code

Put IFSC only in:

bank.ifsc

If an AD Code is not explicitly present, bank.code must be null.


============================================================
EXCEL
============================================================

Do NOT choose Excel cells.

Do NOT generate Excel formatting.

Do NOT redesign the workbook.

Do NOT modify document wording for presentation purposes.

DocPort's existing deterministic Excel mapping system will handle
all workbook placement after extraction.
"""


# ============================================================
# GEMINI CLIENT
# ============================================================

def get_gemini_client():

    api_key = os.getenv(
        "GEMINI_API_KEY"
    )

    if not api_key:

        raise HTTPException(
            status_code=500,
            detail=(
                "GEMINI_API_KEY is not configured."
            ),
        )

    return genai.Client(
        api_key=api_key
    )


# ============================================================
# MISSING VALUE HELPERS
# ============================================================

def is_missing(value):

    if value is None:
        return True

    if isinstance(value, str):
        return value.strip() == ""

    return False


def build_missing_fields(
    data: ExtractedData
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

            missing.append(
                field_name
            )

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
# TEMPORARY GEMINI ERROR CHECK
# ============================================================

def is_temporary_gemini_error(
    error: Exception
):

    text = str(error).upper()

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
        "INTERNAL",
    ]

    return any(
        marker in text
        for marker in temporary_markers
    )


# ============================================================
# GEMINI EXTRACTION WITH AUTOMATIC RETRIES
# ============================================================

async def extract_with_gemini(
    pdf_bytes: bytes
):

    client = get_gemini_client()

    model = os.getenv(
        "GEMINI_MODEL",
        "gemini-3.5-flash",
    )

    pdf_part = types.Part.from_bytes(
        data=pdf_bytes,
        mime_type="application/pdf",
    )

    # Four total attempts.
    #
    # Attempt 1: immediately
    # Attempt 2: after 2 sec
    # Attempt 3: after 4 sec
    # Attempt 4: after 8 sec
    #
    # This specifically helps with temporary Gemini
    # capacity/high-demand/503 problems.
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

        if delay:

            print(
                (
                    "Gemini temporarily unavailable. "
                    f"Retrying in {delay} seconds..."
                )
            )

            await asyncio.sleep(
                delay
            )

        try:

            print(
                (
                    "DocPort Gemini extraction "
                    f"attempt {attempt_number}/"
                    f"{len(retry_delays)} "
                    f"using {model}"
                )
            )

            # Run the synchronous Gemini SDK call away
            # from FastAPI's async event loop.
            response = await asyncio.to_thread(
                client.models.generate_content,

                model=model,

                contents=[
                    pdf_part,
                    EXTRACTION_PROMPT,
                ],

                config=types.GenerateContentConfig(
                    response_mime_type=(
                        "application/json"
                    ),

                    response_schema=(
                        ExtractedData
                    ),
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

                return (
                    ExtractedData
                    .model_validate_json(
                        response_text
                    )
                )

            raise ValueError(
                (
                    "Gemini returned an empty "
                    "structured response."
                )
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

    raise HTTPException(
        status_code=503,
        detail=(
            "Gemini is temporarily busy after "
            "multiple automatic attempts. "
            "Please retry the invoice shortly."
        ),
    )


# ============================================================
# API HEALTH
# ============================================================

@app.get("/api")
def api_home():

    return {
        "service": "DocPort",
        "status": "online",
        "gemini_configured": bool(
            os.getenv(
                "GEMINI_API_KEY"
            )
        ),
        "gemini_model": os.getenv(
            "GEMINI_MODEL",
            "gemini-3.5-flash",
        ),
    }


@app.get("/api/health")
def api_health():

    return {
        "ok": True,
        "service": "DocPort",
        "gemini_configured": bool(
            os.getenv(
                "GEMINI_API_KEY"
            )
        ),
    }


# ============================================================
# GEMINI PDF EXTRACTION ENDPOINT
# ============================================================

@app.post("/api/invoices/extract")
async def extract_invoice(
    file: UploadFile = File(...)
):

    filename = (
        file.filename
        or "invoice.pdf"
    )

    is_pdf = (
        file.content_type
        == "application/pdf"
        or filename.lower().endswith(
            ".pdf"
        )
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
            detail=(
                "The uploaded PDF is empty."
            ),
        )

    # Safety limit.
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
            extracted
            .product
            .line_items
        )

        if product_count == 0:

            raise HTTPException(
                status_code=422,
                detail=(
                    "No product line items "
                    "were detected in the PDF."
                ),
            )

        # Your approved workbook currently has
        # fixed product-row slots.
        if product_count > MAX_PRODUCTS:

            raise HTTPException(
                status_code=422,
                detail=(
                    f"The invoice contains "
                    f"{product_count} products, "
                    f"but the approved Excel "
                    f"template supports "
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

            # Kept so existing DocPort data
            # structure remains compatible.
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
# EXISTING APPROVED EXCEL GENERATION
#
# IMPORTANT:
#
# Gemini does NOT generate the workbook.
#
# Gemini only supplies structured invoice values.
#
# Your existing:
#
#   field_mapping.py
#   excel_generator.py
#   ExportFlow_Master_Template.xlsx
#
# continue to control the Excel output.
# ============================================================

@app.post("/api/documents/generate")
async def generate_document(
    payload: dict
):

    try:

        output_path = Path(
            generate_excel(
                payload
            )
        )

        if not output_path.exists():

            raise FileNotFoundError(
                (
                    "The Excel workbook "
                    "was not generated."
                )
            )

        return FileResponse(
            path=str(
                output_path
            ),

            media_type=(
                "application/"
                "vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),

            filename=(
                output_path.name
            ),

            background=(
                BackgroundTask(
                    lambda:
                    output_path.unlink(
                        missing_ok=True
                    )
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
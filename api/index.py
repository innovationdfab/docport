import asyncio
import io
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
from pypdf import PdfReader
from starlette.background import BackgroundTask


# ============================================================
# PROJECT PATHS
# ============================================================

ROOT_DIR = Path(__file__).resolve().parents[1]
BACKEND_DIR = ROOT_DIR / "backend"

load_dotenv(ROOT_DIR / ".env.local")

if str(BACKEND_DIR) not in sys.path:
    sys.path.insert(0, str(BACKEND_DIR))


# Existing approved DocPort Excel system.
from excel_generator import generate_excel, MAX_PRODUCTS

# New deterministic local invoice parser.
from local_invoice_parser import (
    parse_invoice_locally,
    local_result_is_complete,
)


# ============================================================
# FASTAPI
# ============================================================

app = FastAPI(
    title="DocPort API",
    version="3.0.0",
)


# ============================================================
# CORS
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
# STRUCTURED DATA MODELS
# ============================================================


class InvoiceData(BaseModel):
    number: Optional[str] = None
    date: Optional[str] = None
    order_date: Optional[str] = None
    po_number: Optional[str] = None


class PartyData(BaseModel):
    name: Optional[str] = None

    contact: Optional[str] = None

    address: Optional[str] = None

    raw_section: Optional[str] = None

    email: Optional[str] = None

    phone: Optional[str] = None

    country: Optional[str] = None


class LineItem(BaseModel):
    part_number: Optional[str] = None

    description: Optional[str] = None

    material: Optional[str] = None

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
            "Do not put SWIFT/BIC here."
        ),
    )

    account_number: Optional[str] = None

    ifsc: Optional[str] = None

    swift_code: Optional[str] = None


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
# GEMINI FALLBACK PROMPT
#
# Gemini is NOT used for normal supported invoices.
# It is used only when local parsing fails.
# ============================================================

GEMINI_FALLBACK_PROMPT = """
You are the fallback invoice extraction engine for DocPort.

Read the attached commercial/export invoice PDF carefully.

Extract only information explicitly visible in the document.

Never invent or assume missing values.

Return null for absent or uncertain values.

Keep these parties separate:

- exporter / shipper
- importer
- consignee / ship-to
- sold-to

Extract every visible product line in its original order.

For every product keep these fields separate:

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

Do not duplicate product descriptions.

Preserve material grades correctly.

Keep product net weight separate from package/gross shipment weight.

Do not calculate missing prices, exchange rates, AWB numbers,
Shipping Bill numbers or other missing fields.

bank.code means Bank AD Code only.

SWIFT/BIC must go only into bank.swift_code.

IFSC must go only into bank.ifsc.

Do not choose Excel cells.

Do not create or redesign Excel.

DocPort's existing Excel mapping system handles workbook generation.
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
# LOCAL PDF TEXT EXTRACTION
#
# ZERO GEMINI TOKENS
# ============================================================


def extract_pdf_text_locally(
    pdf_bytes: bytes,
) -> str:

    try:

        reader = PdfReader(
            io.BytesIO(pdf_bytes)
        )

        pages = []

        for page_number, page in enumerate(
            reader.pages,
            start=1,
        ):

            try:

                page_text = (
                    page.extract_text()
                    or ""
                )

            except Exception:

                page_text = ""

            page_text = page_text.strip()

            if page_text:

                pages.append(
                    (
                        f"\n--- PAGE {page_number} ---\n"
                        f"{page_text}"
                    )
                )

        return "\n".join(
            pages
        ).strip()

    except Exception as exc:

        print(
            f"Local PDF reading failed: {exc}"
        )

        return ""


# ============================================================
# CHECK IF PDF CONTAINS USABLE TEXT
# ============================================================


def has_usable_invoice_text(
    text: str,
) -> bool:

    if not text:

        return False

    if len(text) < 200:

        return False

    lowered = text.lower()

    invoice_signals = [
        "invoice",
        "shipper",
        "exporter",
        "consignee",
        "ship to",
        "purchase order",
        "goods description",
        "unit price",
        "position price",
        "packaging",
        "total",
    ]

    matches = sum(
        1
        for signal in invoice_signals
        if signal in lowered
    )

    return matches >= 2


# ============================================================
# TEMPORARY GEMINI ERRORS
# ============================================================


def is_temporary_gemini_error(
    error: Exception,
) -> bool:

    text = str(error).upper()

    markers = [
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
        marker in text
        for marker in markers
    )


# ============================================================
# PARSE GEMINI RESPONSE
# ============================================================


def parse_gemini_response(
    response,
) -> ExtractedData:

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
        "Gemini returned an empty response."
    )


# ============================================================
# GEMINI PDF FALLBACK
#
# Only called if local parsing fails.
# ============================================================


async def extract_with_gemini_fallback(
    pdf_bytes: bytes,
) -> ExtractedData:

    client = get_gemini_client()

    model = os.getenv(
        "GEMINI_MODEL",
        "gemini-3.8-flash",
    )

    pdf_part = types.Part.from_bytes(
        data=pdf_bytes,
        mime_type="application/pdf",
    )

    last_error = None

    # Maximum 2 attempts.
    #
    # No long 2s/4s/8s retry chain.
    for attempt in range(2):

        try:

            print(
                (
                    "Gemini fallback attempt "
                    f"{attempt + 1}/2 "
                    f"using {model}"
                )
            )

            response = await asyncio.to_thread(
                client.models.generate_content,

                model=model,

                contents=[
                    pdf_part,
                    GEMINI_FALLBACK_PROMPT,
                ],

                config=types.GenerateContentConfig(
                    response_mime_type=(
                        "application/json"
                    ),

                    response_schema=(
                        ExtractedData
                    ),

                    temperature=0,
                ),
            )

            return parse_gemini_response(
                response
            )

        except Exception as exc:

            last_error = exc

            if not is_temporary_gemini_error(
                exc
            ):

                raise

            if attempt == 0:

                print(
                    (
                        "Gemini temporarily busy. "
                        "One quick retry..."
                    )
                )

                await asyncio.sleep(1)

    print(
        (
            "Gemini fallback failed after "
            f"quick retry: {last_error}"
        )
    )

    raise HTTPException(
        status_code=503,
        detail=(
            "Gemini fallback is temporarily busy. "
            "Please retry this invoice shortly."
        ),
    )


# ============================================================
# MAIN SMART EXTRACTION
#
# THIS IS THE IMPORTANT PART.
#
# 1. pypdf locally
# 2. local deterministic parser
# 3. Gemini ONLY when necessary
# ============================================================


async def smart_extract_invoice(
    pdf_bytes: bytes,
):

    # --------------------------------------------------------
    # STEP 1
    # Read PDF locally.
    # --------------------------------------------------------

    invoice_text = await asyncio.to_thread(
        extract_pdf_text_locally,
        pdf_bytes,
    )

    print(
        (
            "Local PDF text characters: "
            f"{len(invoice_text)}"
        )
    )

    # --------------------------------------------------------
    # STEP 2
    # Try local deterministic parser first.
    #
    # NO GEMINI
    # NO AI TOKENS
    # --------------------------------------------------------

    if has_usable_invoice_text(
        invoice_text
    ):

        print(
            "Trying DocPort local invoice parser..."
        )

        try:

            local_data = (
                parse_invoice_locally(
                    invoice_text
                )
            )

            if local_result_is_complete(
                local_data
            ):

                extracted = (
                    ExtractedData.model_validate(
                        local_data
                    )
                )

                print(
                    "========================================"
                )

                print(
                    "LOCAL PARSER SUCCESS"
                )

                print(
                    "GEMINI NOT USED"
                )

                print(
                    "========================================"
                )

                return (
                    extracted,
                    invoice_text,
                    "local_parser",
                )

            print(
                (
                    "Local parser did not find "
                    "enough required fields."
                )
            )

        except Exception as exc:

            print(
                (
                    "Local parser error: "
                    f"{exc}"
                )
            )

    else:

        print(
            (
                "PDF text is missing or "
                "not readable enough locally."
            )
        )

    # --------------------------------------------------------
    # STEP 3
    # Only now use Gemini.
    # --------------------------------------------------------

    print(
        "Using Gemini PDF fallback..."
    )

    extracted = (
        await extract_with_gemini_fallback(
            pdf_bytes
        )
    )

    return (
        extracted,
        invoice_text,
        "gemini_pdf_fallback",
    )


# ============================================================
# MISSING FIELD HELPERS
# ============================================================


def is_missing(
    value,
):

    if value is None:

        return True

    if isinstance(
        value,
        str,
    ):

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

        if is_missing(
            value
        ):

            missing.append(
                field_name
            )

    for index, item in enumerate(
        data.product.line_items
    ):

        fields = [
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

        for field_name, value in fields:

            if is_missing(
                value
            ):

                missing.append(
                    (
                        f"product.line_items"
                        f"[{index}]."
                        f"{field_name}"
                    )
                )

    return missing


# ============================================================
# API STATUS
# ============================================================


@app.get("/api")
def api_home():

    return {
        "service": "DocPort",
        "status": "online",
        "version": "3.0.0",

        "gemini_configured": bool(
            os.getenv(
                "GEMINI_API_KEY"
            )
        ),

        "gemini_model": os.getenv(
            "GEMINI_MODEL",
            "gemini-3.8-flash",
        ),

        "primary_extraction": (
            "local_parser"
        ),

        "fallback_extraction": (
            "gemini_pdf"
        ),
    }


@app.get("/api/health")
def api_health():

    return {
        "ok": True,
        "service": "DocPort",
        "version": "3.0.0",
    }


# ============================================================
# PDF EXTRACTION ENDPOINT
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

        (
            extracted,
            local_text,
            extraction_mode,
        ) = await smart_extract_invoice(
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
                    "were detected."
                ),
            )

        if product_count > MAX_PRODUCTS:

            raise HTTPException(
                status_code=422,
                detail=(
                    f"The invoice contains "
                    f"{product_count} products, "
                    f"but the approved Excel "
                    f"template currently supports "
                    f"{MAX_PRODUCTS} products."
                ),
            )

        print(
            (
                "Extraction completed: "
                f"{filename}"
            )
        )

        print(
            (
                "Extraction mode: "
                f"{extraction_mode}"
            )
        )

        return {
            "filename":
                filename,

            "extracted":
                extracted.model_dump(),

            "missing_fields":
                build_missing_fields(
                    extracted
                ),

            # Keep existing DocPort compatibility.
            "raw_text":
                local_text,

            # Extra debugging/info field.
            "extraction_mode":
                extraction_mode,
        }

    except HTTPException:

        raise

    except Exception as exc:

        print(
            (
                "DocPort extraction failed: "
                f"{exc}"
            )
        )

        raise HTTPException(
            status_code=500,
            detail=(
                "Invoice extraction failed: "
                f"{exc}"
            ),
        )


# ============================================================
# EXISTING EXCEL GENERATION
#
# IMPORTANT:
#
# Nothing here changes your Excel layout.
#
# Existing:
#
# backend/field_mapping.py
# backend/excel_generator.py
# backend/templates/ExportFlow_Master_Template.xlsx
#
# remain responsible for your six-sheet workbook.
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
                (
                    "Excel workbook "
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
                "DocPort Excel generation failed: "
                f"{exc}"
            )
        )

        raise HTTPException(
            status_code=500,
            detail=(
                "Excel generation failed: "
                f"{exc}"
            ),
        )
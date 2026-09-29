from pathlib import Path
from datetime import datetime
import os
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

from field_mapping import (
    SHEETS,
    CELLS,
    INVOICE_PRODUCT_ROWS,
    DHL_EXTRA_CELLS,
)


# ============================================================
# PATHS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

TEMPLATE_PATH = (
    BASE_DIR
    / "templates"
    / "ExportFlow_Master_Template.xlsx"
)

GENERATED_DIR = (
    Path("/tmp")
    / "docport-generated"
    if os.environ.get("VERCEL")
    else BASE_DIR
    / "generated"
)


# ============================================================
# XLSX NAMESPACES
# ============================================================

MAIN_NS = (
    "http://schemas.openxmlformats.org/"
    "spreadsheetml/2006/main"
)

REL_NS = (
    "http://schemas.openxmlformats.org/"
    "officeDocument/2006/relationships"
)


# ============================================================
# SETTINGS
# ============================================================

MAX_PRODUCTS = len(
    INVOICE_PRODUCT_ROWS
)

DEFAULT_INDIA_HSN = "76169990"


# ============================================================
# BASIC HELPERS
# ============================================================

def clean(value):
    if value is None:
        return None

    value = str(
        value
    ).strip()

    return (
        value
        if value
        else None
    )


def as_float(value):
    if (
        value is None
        or value == ""
    ):
        return None

    try:
        return float(
            str(value)
            .replace(",", "")
            .replace("$", "")
            .replace("₹", "")
            .replace("€", "")
            .replace("£", "")
            .strip()
        )

    except Exception:
        return None


def as_int(value):
    number = as_float(
        value
    )

    return (
        int(number)
        if number is not None
        else None
    )


# ============================================================
# DATE HELPERS
# ============================================================

def parse_date(value):
    if not value:
        return None

    text = str(
        value
    ).strip()

    for fmt in (
        "%d.%m.%Y",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d",
    ):
        try:
            return datetime.strptime(
                text,
                fmt,
            )

        except ValueError:
            pass

    return None


def date_slash(value):
    parsed = parse_date(
        value
    )

    return (
        parsed.strftime(
            "%d/%m/%Y"
        )
        if parsed
        else (
            clean(value)
            or ""
        )
    )


def date_dash(value):
    parsed = parse_date(
        value
    )

    return (
        parsed.strftime(
            "%d-%m-%Y"
        )
        if parsed
        else (
            clean(value)
            or ""
        )
    )


# ============================================================
# PRODUCT HELPERS
# ============================================================

def get_line_items(data):
    items = (
        data
        .get(
            "product",
            {},
        )
        .get(
            "line_items",
            [],
        )
    )

    return (
        items
        if isinstance(
            items,
            list,
        )
        else []
    )


def build_product_description(item):
    """
    Build the complete reviewed Description of Goods.

    Preserve Part #, description and Material together in
    downloaded Excel, matching the website preview.

    Avoid duplicates when description already contains
    Part # or Material.
    """

    if not item:
        return None

    description = clean(
        item.get(
            "description"
        )
    )

    part_number = clean(
        item.get(
            "part_number"
        )
    )

    material = clean(
        item.get(
            "material"
        )
    )

    description_text = (
        description
        or ""
    )

    parts = []

    if part_number:
        part_exists = re.search(
            r"Part\s*#\s*"
            + re.escape(
                part_number
            ),
            description_text,
            re.IGNORECASE,
        )

        if not part_exists:
            parts.append(
                f"Part #{part_number}"
            )

    if description:
        parts.append(
            description
        )

    if material:
        material_exists = re.search(
            r"Material\s*:",
            description_text,
            re.IGNORECASE,
        )

        if not material_exists:
            parts.append(
                f"Material: {material}"
            )

    return (
        "\n".join(
            parts
        ).strip()
        or None
    )

def calculate_multi_product_totals(
    line_items
):
    total_quantity = 0
    total_net_weight = 0.0
    total_value = 0.0

    has_quantity = False
    has_weight = False
    has_value = False

    for item in line_items:
        qty = as_int(
            item.get(
                "quantity"
            )
        )

        if qty is not None:
            total_quantity += qty
            has_quantity = True

        weight = as_float(
            item.get(
                "net_weight_kg"
            )
        )

        if weight is not None:
            total_net_weight += weight
            has_weight = True

        value = as_float(
            item.get(
                "position_price"
            )
        )

        if value is None:
            unit_price = as_float(
                item.get(
                    "unit_price"
                )
            )

            if (
                qty is not None
                and unit_price is not None
            ):
                value = (
                    qty
                    * unit_price
                )

        if value is not None:
            total_value += value
            has_value = True

    return {
        "quantity":
            (
                total_quantity
                if has_quantity
                else None
            ),

        "net_weight":
            (
                round(
                    total_net_weight,
                    3,
                )
                if has_weight
                else None
            ),

        "total_value":
            (
                round(
                    total_value,
                    2,
                )
                if has_value
                else None
            ),
    }


def currency_symbol_for(
    item,
    shipment,
):
    symbol = clean(
        item.get(
            "currency_symbol"
        )
    )

    if symbol:
        return symbol

    currency = (
        clean(
            shipment.get(
                "currency"
            )
        )
        or "USD"
    ).upper()

    return {
        "USD": "$",
        "EUR": "€",
        "GBP": "£",
        "INR": "₹",
    }.get(
        currency,
        "",
    )


# ============================================================
# BUILD EXCEL UPDATES
# ============================================================

def build_excel_updates(payload):
    data = payload.get(
        "extracted",
        payload,
    )

    invoice = (
        data.get(
            "invoice",
            {},
        )
        or {}
    )

    consignee = (
        data.get(
            "consignee",
            {},
        )
        or {}
    )

    shipment = (
        data.get(
            "shipment",
            {},
        )
        or {}
    )

    company = (
        data.get(
            "company",
            {},
        )
        or {}
    )

    bank = (
        data.get(
            "bank",
            {},
        )
        or {}
    )

    line_items = get_line_items(
        data
    )

    if not line_items:
        raise ValueError(
            "No product line items "
            "were extracted from the PDF."
        )

    if (
        len(line_items)
        > MAX_PRODUCTS
    ):
        raise ValueError(
            f"The PDF contains "
            f"{len(line_items)} products, "
            f"but the final Excel template "
            f"supports {MAX_PRODUCTS}. "
            "No product data was removed."
        )

    # ========================================================
    # INVOICE NUMBER
    # ========================================================

    invoice_number = clean(
        invoice.get(
            "number"
        )
    )

    # ========================================================
    # ONE DYNAMIC DATE FOR ALL 6 SHEETS
    # ========================================================
    #
    # No hardcoded date.
    #
    # Every new PDF uses the date extracted from that PDF.
    #
    # Priority:
    # 1. invoice.date
    # 2. shipment.shipping_bill_date only if invoice.date missing
    #
    # This SAME actual date is used in:
    #
    # Invoice
    # DHL
    # FORM SDF
    # EVD
    # AlumSteel
    # SCOMET
    #
    # ========================================================

    extracted_invoice_date = clean(
        invoice.get(
            "date"
        )
    )

    extracted_shipping_date = clean(
        shipment.get(
            "shipping_bill_date"
        )
    )

    pdf_document_date = (
        extracted_invoice_date
        or extracted_shipping_date
    )

    invoice_date_slash = date_slash(
        pdf_document_date
    )

    invoice_date_dash = date_dash(
        pdf_document_date
    )

    shipping_bill_date = (
        pdf_document_date
    )

    shipping_date_slash = date_slash(
        pdf_document_date
    )

    shipping_date_dash = date_dash(
        pdf_document_date
    )

    # ========================================================
    # OTHER VALUES
    # ========================================================

    po_number = clean(
        invoice.get(
            "po_number"
        )
    )

    consignee_name = (
        clean(
            consignee.get(
                "name"
            )
        )
        or clean(
            consignee.get(
                "contact"
            )
        )
    )

    consignee_raw = (
        clean(
            consignee.get(
                "raw_section"
            )
        )
        or clean(
            consignee.get(
                "address"
            )
        )
    )

    dimensions = clean(
        shipment.get(
            "package_dimensions"
        )
    )

    package_weight = as_float(
        shipment.get(
            "package_weight_kg"
        )
    )

    exchange_rate = as_float(
        shipment.get(
            "exchange_rate"
        )
    )

    awb = clean(
        shipment.get(
            "awb_number"
        )
    )

    shipping_bill = clean(
        shipment.get(
            "shipping_bill_number"
        )
    )

    currency = (
        clean(
            shipment.get(
                "currency"
            )
        )
        or "USD"
    ).upper()

    # ========================================================
    # TOTALS
    # ========================================================

    totals = (
        calculate_multi_product_totals(
            line_items
        )
    )

    total_quantity = (
        totals[
            "quantity"
        ]
    )

    if total_quantity is None:
        total_quantity = as_int(
            shipment.get(
                "total_quantity"
            )
        )

    total_net_weight = as_float(
        shipment.get(
            "net_weight_kg"
        )
    )

    if total_net_weight is None:
        total_net_weight = (
            totals[
                "net_weight"
            ]
        )

    total_value = as_float(
        shipment.get(
            "total_amount"
        )
    )

    if total_value is None:
        total_value = (
            totals[
                "total_value"
            ]
        )

    amount_inr = None

    if (
        total_value is not None
        and exchange_rate is not None
    ):
        amount_inr = round(
            total_value
            * exchange_rate,
            2,
        )

    # ========================================================
    # COUNTRY OF ORIGIN
    # ========================================================

    countries = []

    for item in line_items:
        country = clean(
            item.get(
                "country_of_origin"
            )
        )

        if (
            country
            and country not in countries
        ):
            countries.append(
                country
            )

    country_origin = (
        ", ".join(
            countries
        )
        if countries
        else "INDIA"
    )

    # ========================================================
    # PRODUCT DESCRIPTIONS
    # ========================================================

    all_product_descriptions = []

    for (
        index,
        item,
    ) in enumerate(
        line_items
    ):
        description = (
            build_product_description(
                item
            )
        )

        if not description:
            continue

        if (
            len(line_items)
            > 1
        ):
            all_product_descriptions.append(
                f"Product "
                f"{index + 1}: "
                f"{description}"
            )

        else:
            all_product_descriptions.append(
                description
            )

    all_products_text = (
        "\n".join(
            all_product_descriptions
        )
        or None
    )

    first_item = (
        line_items[0]
        if line_items
        else {}
    )

    first_part = clean(
        first_item.get(
            "part_number"
        )
    )

    first_title = (
        clean(
            first_item.get(
                "description"
            )
        )
        or first_part
        or ""
    )

    updates = {}

    # ========================================================
    # 1. INVOICE
    # ========================================================

    invoice_updates = {
        CELLS[
            "invoice"
        ][
            "invoice_date"
        ]:
            (
                f"Date-"
                f"{invoice_date_slash}"
                if invoice_date_slash
                else ""
            ),

        CELLS[
            "invoice"
        ][
            "invoice_number"
        ]:
            invoice_number,

        CELLS[
            "invoice"
        ][
            "order_date"
        ]:
            (
                f"Date-"
                f"{invoice_date_slash}"
                if invoice_date_slash
                else ""
            ),

        CELLS[
            "invoice"
        ][
            "po_number"
        ]:
            po_number,

        CELLS[
            "invoice"
        ][
            "consignee_label"
        ]:
            (
                "Ship to / Consignee: "
                f"{consignee_name or ''}"
            ),

        CELLS[
            "invoice"
        ][
            "consignee_address"
        ]:
            consignee_raw,

        CELLS[
            "invoice"
        ][
            "consignee_contact"
        ]:
            (
                "Ctc person- "
                f"{consignee_name or ''}"
            ),

        CELLS[
            "invoice"
        ][
            "country_origin_1"
        ]:
            country_origin,

        CELLS[
            "invoice"
        ][
            "country_origin_2"
        ]:
            country_origin,

        CELLS[
            "invoice"
        ][
            "exchange_rate"
        ]:
            {
                "value":
                    exchange_rate,

                "kind":
                    "number",
            },

        CELLS[
            "invoice"
        ][
            "amount_inr"
        ]:
            {
                "value":
                    amount_inr,

                "kind":
                    "number",
            },

        CELLS[
            "invoice"
        ][
            "amount_usd"
        ]:
            {
                "value":
                    total_value,

                "kind":
                    "number",
            },

        CELLS[
            "invoice"
        ][
            "total_before_tax"
        ]:
            {
                "value":
                    total_value,

                "kind":
                    "number",
            },

        CELLS[
            "invoice"
        ][
            "total_after_tax"
        ]:
            {
                "value":
                    total_value,

                "kind":
                    "number",
            },

        CELLS[
            "invoice"
        ][
            "package_count"
        ]:
            {
                "value":
                    (
                        as_int(
                            shipment.get(
                                "package_count"
                            )
                        )
                        or 1
                    ),

                "kind":
                    "number",
            },

        CELLS[
            "invoice"
        ][
            "dimensions"
        ]:
            dimensions,

        CELLS[
            "invoice"
        ][
            "gross_weight"
        ]:
            (
                f"{package_weight:g} kg"
                if package_weight is not None
                else None
            ),

        CELLS[
            "invoice"
        ][
            "net_weight"
        ]:
            (
                f"{total_net_weight:g} kg"
                if total_net_weight is not None
                else None
            ),
    }

    # ========================================================
    # OPTIONAL COMPANY DATA
    # ========================================================

    if (
        "iec"
        in CELLS[
            "invoice"
        ]
    ):
        invoice_updates[
            CELLS[
                "invoice"
            ][
                "iec"
            ]
        ] = clean(
            company.get(
                "iec"
            )
        )

    if (
        "gstin"
        in CELLS[
            "invoice"
        ]
    ):
        invoice_updates[
            CELLS[
                "invoice"
            ][
                "gstin"
            ]
        ] = clean(
            company.get(
                "gstin"
            )
        )

    if (
        "bank_ad_code"
        in CELLS[
            "invoice"
        ]
    ):
        invoice_updates[
            CELLS[
                "invoice"
            ][
                "bank_ad_code"
            ]
        ] = clean(
            bank.get(
                "code"
            )
        )

    if (
        "bank_account"
        in CELLS[
            "invoice"
        ]
    ):
        invoice_updates[
            CELLS[
                "invoice"
            ][
                "bank_account"
            ]
        ] = clean(
            bank.get(
                "account_number"
            )
        )

    if (
        "bank_ifsc"
        in CELLS[
            "invoice"
        ]
    ):
        invoice_updates[
            CELLS[
                "invoice"
            ][
                "bank_ifsc"
            ]
        ] = clean(
            bank.get(
                "ifsc"
            )
        )

    # ========================================================
    # CLEAR PRODUCT ROWS
    # ========================================================

    for slot in (
        INVOICE_PRODUCT_ROWS
    ):
        invoice_updates[
            slot[
                "sr_no"
            ]
        ] = None

        invoice_updates[
            slot[
                "india_hsn"
            ]
        ] = None

        invoice_updates[
            slot[
                "us_hts"
            ]
        ] = None

        invoice_updates[
            slot[
                "description"
            ]
        ] = None

        invoice_updates[
            slot[
                "quantity"
            ]
        ] = None

        invoice_updates[
            slot[
                "unit_price"
            ]
        ] = None

        invoice_updates[
            slot[
                "taxable_value"
            ]
        ] = None

        invoice_updates[
            slot[
                "igst_percent"
            ]
        ] = None

        invoice_updates[
            slot[
                "igst_amount"
            ]
        ] = None

    # ========================================================
    # FILL PRODUCT ROWS
    # ========================================================

    for (
        index,
        item,
    ) in enumerate(
        line_items
    ):
        slot = (
            INVOICE_PRODUCT_ROWS[
                index
            ]
        )

        description = (
            build_product_description(
                item
            )
        )

        qty = as_int(
            item.get(
                "quantity"
            )
        )

        unit_price = as_float(
            item.get(
                "unit_price"
            )
        )

        position_price = as_float(
            item.get(
                "position_price"
            )
        )

        if (
            position_price is None
            and qty is not None
            and unit_price is not None
        ):
            position_price = round(
                qty
                * unit_price,
                2,
            )

        india_hsn = (
            clean(
                item.get(
                    "india_hsn"
                )
            )
            or clean(
                item.get(
                    "india_hsn_code"
                )
            )
            or DEFAULT_INDIA_HSN
        )

        us_hts = clean(
            item.get(
                "hs_code"
            )
        )

        symbol = (
            currency_symbol_for(
                item,
                shipment,
            )
        )

        invoice_updates[
            slot[
                "sr_no"
            ]
        ] = {
            "value":
                index + 1,

            "kind":
                "number",
        }

        invoice_updates[
            slot[
                "india_hsn"
            ]
        ] = india_hsn

        invoice_updates[
            slot[
                "us_hts"
            ]
        ] = us_hts

        invoice_updates[
            slot[
                "description"
            ]
        ] = description

        invoice_updates[
            slot[
                "quantity"
            ]
        ] = {
            "value":
                qty,

            "kind":
                "number",
        }

        invoice_updates[
            slot[
                "unit_price"
            ]
        ] = (
            f"{symbol}"
            f"{unit_price:.2f}"
            if unit_price is not None
            else None
        )

        invoice_updates[
            slot[
                "taxable_value"
            ]
        ] = {
            "value":
                position_price,

            "kind":
                "number",
        }

        invoice_updates[
            slot[
                "igst_percent"
            ]
        ] = "0%"

        invoice_updates[
            slot[
                "igst_amount"
            ]
        ] = "0"

    updates[
        SHEETS[
            "invoice"
        ]
    ] = invoice_updates

    # ========================================================
    # 2. DHL EXPRESS SLI
    # ========================================================

    taxable_text = (
        "TAXABLE AMOUNT - "
        + (
            f"{amount_inr:.2f}"
            if amount_inr is not None
            else ""
        )
        + "\n"
        + "IGST RATE - 0"
        + "\n"
        + "IGST AMOUNT - 0"
        + "\n"
        + "GST Compensation Cess - 0"
    )

    taxable_text = (
        clean(
            shipment.get(
                "dhl_tax_block_override"
            )
        )
        or taxable_text
    )

    dhl_updates = {
        CELLS[
            "dhl"
        ][
            "invoice_number"
        ]:
            invoice_number,

        CELLS[
            "dhl"
        ][
            "consignee"
        ]:
            consignee_name,

        # SAME PDF DATE
        CELLS[
            "dhl"
        ][
            "invoice_date"
        ]:
            invoice_date_dash,

        CELLS[
            "dhl"
        ][
            "awb"
        ]:
            awb,

        CELLS[
            "dhl"
        ][
            "tax_block"
        ]:
            taxable_text,

        CELLS[
            "dhl"
        ][
            "fob_value"
        ]:
            (
                f"{total_value:.2f} "
                f"{currency}"
                if total_value is not None
                else None
            ),

        CELLS[
            "dhl"
        ][
            "net_weight"
        ]:
            (
                f"{total_net_weight:g} kg"
                if total_net_weight is not None
                else None
            ),

        CELLS[
            "dhl"
        ][
            "gross_weight"
        ]:
            (
                f"{package_weight:g} kg"
                if package_weight is not None
                else None
            ),
    }

    extra_sources = {
        "rodtep_yes_no":
            shipment.get(
                "rodtep_yes_no"
            ),

        "rodtep_total_line_items":
            shipment.get(
                "rodtep_total_line_items"
            ),

        "rodtep_claimed_line_items":
            shipment.get(
                "rodtep_claimed_line_items"
            ),

        "freight":
            shipment.get(
                "freight"
            ),

        "insurance":
            shipment.get(
                "insurance"
            ),

        "commission":
            shipment.get(
                "commission"
            ),

        "discount":
            shipment.get(
                "discount"
            ),

        "packing_charges":
            shipment.get(
                "packing_charges"
            ),

        "package_count":
            (
                shipment.get(
                    "package_count"
                )
                or 1
            ),

        "state_origin":
            shipment.get(
                "state_origin"
            ),

        "district_origin":
            shipment.get(
                "district_origin"
            ),
    }

    for (
        key,
        cell,
    ) in (
        DHL_EXTRA_CELLS.items()
    ):
        value = (
            extra_sources.get(
                key
            )
        )

        if value is not None:
            dhl_updates[
                cell
            ] = value

    updates[
        SHEETS[
            "dhl"
        ]
    ] = dhl_updates

    # ========================================================
    # 3. FORM SDF
    # ========================================================

    updates[
        SHEETS[
            "sdf"
        ]
    ] = {
        CELLS[
            "sdf"
        ][
            "shipping_bill_header"
        ]:
            shipping_bill,

        # SAME PDF DATE
        CELLS[
            "sdf"
        ][
            "date"
        ]:
            shipping_date_slash,

        CELLS[
            "sdf"
        ][
            "shipping_bill_declaration"
        ]:
            (
                "We here by declare that "
                "we are the SELLER / CONSIGNOR "
                "of the goods in Respect of which "
                "this declaration made and that "
                "particulars given to shipping Bill No: "
                f"{shipping_bill or ''} "
                "Date "
                f"{shipping_date_slash or ''} "
                "are true and that :"
            ),
    }

    # ========================================================
    # 4. EVD
    # ========================================================

    updates[
        SHEETS[
            "evd"
        ]
    ] = {
        CELLS[
            "evd"
        ][
            "shipping_bill"
        ]:
            (
                "1. Shipping Bill No. "
                f"{shipping_bill or ''}"
                " & Date:- "
                f"{shipping_date_slash or ''}"
            ),

        CELLS[
            "evd"
        ][
            "invoice_details"
        ]:
            (
                "2. Invoice No. & Date   "
                f"{invoice_number or ''}"
                " & "
                f"{invoice_date_slash or ''}"
            ),

        # RIGHT-SIDE EVD DATE
        # Uses SAME CURRENT PDF date.
        CELLS[
            "evd"
        ][
            "invoice_date"
        ]:
            (
                "Date: "
                f"{invoice_date_slash or ''}"
            ),

        CELLS[
            "evd"
        ][
            "previous_shipping_bill"
        ]:
            (
                "Shipping Bill No: "
                f"{shipping_bill or ''}"
                " and date: "
                f"{shipping_date_slash or ''}"
            ),

        CELLS[
            "evd"
        ][
            "date"
        ]:
            (
                "Date: "
                f"{shipping_date_slash or ''}"
            ),
    }

    # ========================================================
    # 5. ALUMINIUM / STEEL
    # ========================================================

    weight_text = (
        (
            "3) Full weight of the product "
            f"{total_net_weight:g} kg"
            "\n"
            "Aluminum content weight "
            f"{total_net_weight:g} kg"
        )
        if total_net_weight is not None
        else (
            "3) Full weight of the product"
            "\n"
            "Aluminum content weight"
        )
    )

    value_text = (
        (
            "Total Value of the product "
            f"{total_value:.2f} "
            f"{currency.lower()}"
            "\n"
            "Value of the Aluminum content "
            f"{total_value:.2f} "
            f"{currency.lower()}"
        )
        if total_value is not None
        else (
            "Total Value of the product"
            "\n"
            "Value of the Aluminum content"
        )
    )

    updates[
        SHEETS[
            "alum"
        ]
    ] = {
        CELLS[
            "alum"
        ][
            "product"
        ]:
            (
                "Product:- "
                f"{all_products_text}"
                if all_products_text
                else None
            ),

        CELLS[
            "alum"
        ][
            "weight_block"
        ]:
            weight_text,

        CELLS[
            "alum"
        ][
            "value_block"
        ]:
            value_text,

        # SAME PDF DATE
        CELLS[
            "alum"
        ][
            "completed_date"
        ]:
            (
                "Completed by : "
                f"{invoice_date_dash or ''}"
            ),

        CELLS[
            "alum"
        ][
            "title"
        ]:
            (
                "Title       : "
                f"{first_title}"
                if first_title
                else "Title       :"
            ),

        # SAME PDF DATE
        CELLS[
            "alum"
        ][
            "date"
        ]:
            (
                "Date        : "
                f"{invoice_date_dash or ''}"
            ),
    }

    # ========================================================
    # 6. SCOMET
    # ========================================================

    hsn_text = ", ".join(
        dict.fromkeys(
            clean(
                item.get(
                    "hs_code"
                )
            )

            for item
            in line_items

            if clean(
                item.get(
                    "hs_code"
                )
            )
        )
    )

    updates[
        SHEETS[
            "scomet"
        ]
    ] = {
        CELLS[
            "scomet"
        ][
            "invoice_number"
        ]:
            (
                "Invoice No: - "
                f"{invoice_number or ''}"
            ),

        # SAME PDF DATE
        CELLS[
            "scomet"
        ][
            "invoice_date"
        ]:
            (
                "Date: "
                f"{invoice_date_dash or ''}"
            ),

        CELLS[
            "scomet"
        ][
            "product"
        ]:
            (
                "Product description:- "
                f"{all_products_text or ''}"
            ),

        CELLS[
            "scomet"
        ][
            "hsn_code"
        ]:
            (
                "HSN code: - "
                f"{hsn_text}"
            ),
    }

    return updates


# ============================================================
# WORKBOOK / SHEET HELPERS
# ============================================================

def get_sheet_paths(
    zip_file
):
    workbook_root = (
        ET.fromstring(
            zip_file.read(
                "xl/workbook.xml"
            )
        )
    )

    relationships_root = (
        ET.fromstring(
            zip_file.read(
                "xl/_rels/"
                "workbook.xml.rels"
            )
        )
    )

    relationships = {
        relationship.attrib.get(
            "Id"
        ):
            relationship.attrib.get(
                "Target"
            )

        for relationship
        in relationships_root
    }

    sheets_node = (
        workbook_root.find(
            f"{{{MAIN_NS}}}sheets"
        )
    )

    if sheets_node is None:
        raise ValueError(
            "Unable to read worksheets "
            "from master workbook."
        )

    sheet_paths = {}

    for sheet in sheets_node:
        sheet_name = (
            sheet.attrib[
                "name"
            ]
        )

        relationship_id = (
            sheet.attrib[
                f"{{{REL_NS}}}id"
            ]
        )

        target = (
            relationships.get(
                relationship_id
            )
        )

        if not target:
            raise ValueError(
                "Unable to resolve "
                f"worksheet '{sheet_name}'."
            )

        if target.startswith(
            "/"
        ):
            sheet_path = (
                target.lstrip(
                    "/"
                )
            )

        else:
            sheet_path = (
                posixpath.normpath(
                    posixpath.join(
                        "xl",
                        target,
                    )
                )
            )

        sheet_paths[
            sheet_name
        ] = sheet_path

    return sheet_paths


def make_cell_pattern(
    reference
):
    return re.compile(
        (
            r'(<(?:'
            r'[A-Za-z_]'
            r'[\w.\-]*:'
            r')?c\b'
            r'(?=[^>]*\br="'
            + re.escape(
                reference
            )
            + r'")'
            r'[^>]*?)'
            r'(?:/>|>'
            r'(.*?)'
            r'</(?:'
            r'[A-Za-z_]'
            r'[\w.\-]*:'
            r')?c>)'
        ),
        re.DOTALL,
    )


def remove_type_attribute(
    start_tag
):
    return re.sub(
        r'\s+t="[^"]*"',
        "",
        start_tag,
    )


def escape_excel_text(
    value
):
    if value is None:
        return ""

    return escape(
        str(value),
        {
            '"':
                "&quot;",

            "'":
                "&apos;",
        },
    )


def xml_prefix_from_start_tag(
    start_tag
):
    match = re.match(
        r"<(?P<prefix>"
        r"[A-Za-z_]"
        r"[\w.\-]*:"
        r")?"
        r"c\b",
        start_tag,
    )

    return (
        (
            match.group(
                "prefix"
            )
            or ""
        )
        if match
        else ""
    )


# ============================================================
# PATCH XLSX CELL
# ============================================================

def patch_cell(
    xml_text,
    reference,
    change,
):
    pattern = (
        make_cell_pattern(
            reference
        )
    )

    match = (
        pattern.search(
            xml_text
        )
    )

    if not match:
        raise ValueError(
            f"Cell {reference} "
            "was not found in "
            "the master workbook."
        )

    start_tag = (
        remove_type_attribute(
            match.group(1)
        )
    )

    original_inner = (
        match.group(2)
        or ""
    )

    prefix = (
        xml_prefix_from_start_tag(
            start_tag
        )
    )

    if isinstance(
        change,
        dict,
    ):
        value = change.get(
            "value"
        )

        kind = change.get(
            "kind",
            "string",
        )

    else:
        value = change
        kind = "string"

    formula_match = re.search(
        (
            r'(<(?:'
            r'[A-Za-z_]'
            r'[\w.\-]*:'
            r')?f\b'
            r'[^>]*>'
            r'.*?'
            r'</(?:'
            r'[A-Za-z_]'
            r'[\w.\-]*:'
            r')?f>'
            r'|'
            r'<(?:'
            r'[A-Za-z_]'
            r'[\w.\-]*:'
            r')?f\b'
            r'[^>]*/>)'
        ),
        original_inner,
        flags=re.DOTALL,
    )

    existing_formula = (
        formula_match.group(1)
        if formula_match
        else None
    )

    c_close = (
        f"</{prefix}c>"
    )

    v_open = (
        f"<{prefix}v>"
    )

    v_close = (
        f"</{prefix}v>"
    )

    is_open = (
        f"<{prefix}is>"
    )

    is_close = (
        f"</{prefix}is>"
    )

    t_open = (
        f'<{prefix}t '
        f'xml:space="preserve">'
    )

    t_close = (
        f"</{prefix}t>"
    )

    # ========================================================
    # NUMBER
    # ========================================================

    if kind == "number":
        if (
            value is None
            or value == ""
        ):
            if existing_formula:
                replacement = (
                    f"{start_tag}>"
                    f"{existing_formula}"
                    f"{c_close}"
                )

            else:
                replacement = (
                    f"{start_tag}/>"
                )

        else:
            if existing_formula:
                replacement = (
                    f"{start_tag}>"
                    f"{existing_formula}"
                    f"{v_open}"
                    f"{value}"
                    f"{v_close}"
                    f"{c_close}"
                )

            else:
                replacement = (
                    f"{start_tag}>"
                    f"{v_open}"
                    f"{value}"
                    f"{v_close}"
                    f"{c_close}"
                )

    # ========================================================
    # STRING
    # ========================================================

    else:
        escaped_value = (
            escape_excel_text(
                value
            )
        )

        replacement = (
            f'{start_tag} '
            f't="inlineStr">'
            f"{is_open}"
            f"{t_open}"
            f"{escaped_value}"
            f"{t_close}"
            f"{is_close}"
            f"{c_close}"
        )

    return (
        xml_text[
            :
            match.start()
        ]
        +
        replacement
        +
        xml_text[
            match.end()
            :
        ]
    )


def patch_sheet(
    xml_bytes,
    changes,
):
    xml_text = (
        xml_bytes.decode(
            "utf-8"
        )
    )

    for (
        reference,
        change,
    ) in (
        changes.items()
    ):
        xml_text = (
            patch_cell(
                xml_text,
                reference,
                change,
            )
        )

    return (
        xml_text.encode(
            "utf-8"
        )
    )


# ============================================================
# VALIDATE GENERATED XLSX
# ============================================================

def validate_generated_xlsx(
    path
):
    try:
        with zipfile.ZipFile(
            path,
            "r",
        ) as workbook:
            broken_file = (
                workbook.testzip()
            )

            if broken_file:
                raise ValueError(
                    "Generated workbook "
                    "contains damaged file: "
                    f"{broken_file}"
                )

            for filename in (
                workbook.namelist()
            ):
                if not (
                    filename.endswith(
                        ".xml"
                    )
                    or filename.endswith(
                        ".rels"
                    )
                ):
                    continue

                try:
                    ET.fromstring(
                        workbook.read(
                            filename
                        )
                    )

                except Exception as exc:
                    raise ValueError(
                        "Generated workbook "
                        "contains invalid XML "
                        f"in {filename}: "
                        f"{exc}"
                    )

    except zipfile.BadZipFile as exc:
        raise ValueError(
            "Generated file is not "
            "a valid XLSX workbook: "
            f"{exc}"
        )


# ============================================================
# GENERATE EXCEL
# ============================================================

def generate_excel(
    payload
):
    if not (
        TEMPLATE_PATH.exists()
    ):
        raise FileNotFoundError(
            "Master Excel template "
            "was not found: "
            f"{TEMPLATE_PATH}"
        )

    GENERATED_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    extracted_data = (
        payload.get(
            "extracted"
        )
    )

    if not isinstance(
        extracted_data,
        dict,
    ):
        extracted_data = (
            payload
        )

    invoice_number = clean(
        extracted_data
        .get(
            "invoice",
            {},
        )
        .get(
            "number"
        )
    )

    safe_invoice = re.sub(
        r"[^A-Za-z0-9_-]+",
        "_",
        invoice_number
        or "export",
    )

    output_path = (
        GENERATED_DIR
        /
        (
            "Export_Documents_"
            f"{safe_invoice}.xlsx"
        )
    )

    updates = (
        build_excel_updates(
            payload
        )
    )

    with zipfile.ZipFile(
        TEMPLATE_PATH,
        "r",
    ) as source_zip:
        sheet_paths = (
            get_sheet_paths(
                source_zip
            )
        )

        patched_files = {}

        for (
            sheet_name,
            changes,
        ) in (
            updates.items()
        ):
            if (
                sheet_name
                not in sheet_paths
            ):
                raise ValueError(
                    "Worksheet "
                    f"'{sheet_name}' "
                    "was not found inside "
                    "the master workbook."
                )

            worksheet_path = (
                sheet_paths[
                    sheet_name
                ]
            )

            original_xml = (
                source_zip.read(
                    worksheet_path
                )
            )

            # IMPORTANT:
            #
            # No row insertion or deletion happens here.
            # The master template already contains the correct
            # US row / heading / product-row structure.
            #
            # We only replace values in existing cells.
            patched_xml = (
                patch_sheet(
                    original_xml,
                    changes,
                )
            )

            patched_files[
                worksheet_path
            ] = patched_xml

        with zipfile.ZipFile(
            output_path,
            "w",
            compression=(
                zipfile.ZIP_DEFLATED
            ),
        ) as output_zip:
            for zip_info in (
                source_zip.infolist()
            ):
                file_data = (
                    source_zip.read(
                        zip_info.filename
                    )
                )

                if (
                    zip_info.filename
                    in patched_files
                ):
                    file_data = (
                        patched_files[
                            zip_info.filename
                        ]
                    )

                output_zip.writestr(
                    zip_info,
                    file_data,
                )

    validate_generated_xlsx(
        output_path
    )

    return output_path
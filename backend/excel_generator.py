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

# The master workbook starts with three formatted product rows
# (27-29). When an invoice contains more products, additional
# rows are inserted dynamically before row 30 using row 29 as
# the formatting template. There is intentionally no fixed
# application-level product-count limit.
INVOICE_PRODUCT_START_ROW = 27
INVOICE_STATIC_START_ROW = 30
INVOICE_TEMPLATE_PRODUCT_ROW_COUNT = len(
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
# DYNAMIC INVOICE PRODUCT ROW HELPERS
# ============================================================

def invoice_product_slot(index):
    """Return the Invoice-sheet cell map for any product index."""
    row = (
        INVOICE_PRODUCT_START_ROW
        + index
    )

    return {
        "sr_no": f"A{row}",
        "india_hsn": f"B{row}",
        "us_hts": f"C{row}",
        "description": f"D{row}",
        "quantity": f"I{row}",
        "unit_price": f"J{row}",
        "taxable_value": f"K{row}",
        "igst_percent": f"M{row}",
        "igst_amount": f"N{row}",
    }


def shift_cell_reference(
    reference,
    start_row,
    row_delta,
):
    """Shift a simple A1 reference when rows are inserted above it."""
    if (
        not row_delta
        or not reference
    ):
        return reference

    match = re.fullmatch(
        r"([A-Z]+)(\d+)",
        str(reference),
    )

    if not match:
        return reference

    column = match.group(1)
    row = int(
        match.group(2)
    )

    if row >= start_row:
        row += row_delta

    return f"{column}{row}"


def shift_a1_range_text(
    value,
    start_row,
    row_delta,
):
    """Shift row numbers inside an A1 reference/range string."""
    if (
        not row_delta
        or not value
    ):
        return value

    pattern = re.compile(
        r"(\$?[A-Z]{1,3})(\$?)(\d+)"
    )

    def replace(match):
        column = match.group(1)
        absolute_row = match.group(2)
        row = int(
            match.group(3)
        )

        if row >= start_row:
            row += row_delta

        return (
            f"{column}"
            f"{absolute_row}"
            f"{row}"
        )

    return pattern.sub(
        replace,
        str(value),
    )


def _move_invoice_row_xml(
    row_xml,
    source_row,
    target_row,
):
    """Move one raw worksheet row while preserving every style/cell."""
    def replace(match):
        prefix = match.group(1)
        column = match.group(2)
        row = int(
            match.group(3)
        )
        suffix = match.group(4)

        if row == source_row:
            row = target_row

        return (
            f"{prefix}"
            f"{column}"
            f"{row}"
            f"{suffix}"
        )

    return re.sub(
        r'(\br=")([A-Z]*)(\d+)(")',
        replace,
        row_xml,
    )


def _shift_invoice_row_xml(
    row_xml,
    start_row,
    row_delta,
):
    """Shift an existing worksheet row and all of its cell references."""
    row_match = re.search(
        r'<(?:[A-Za-z_][\w.\-]*:)?row\b[^>]*\br="(\d+)"',
        row_xml,
    )

    if not row_match:
        return row_xml

    source_row = int(
        row_match.group(1)
    )

    if source_row < start_row:
        return row_xml

    return _move_invoice_row_xml(
        row_xml,
        source_row,
        source_row + row_delta,
    )


def _shift_reference_attributes(
    xml_fragment,
    start_row,
    row_delta,
):
    """Shift ref/sqref attributes outside sheetData after row insertion."""
    if not row_delta:
        return xml_fragment

    def replace(match):
        return (
            match.group(1)
            + shift_a1_range_text(
                match.group(2),
                start_row,
                row_delta,
            )
            + match.group(3)
        )

    return re.sub(
        r'(\b(?:ref|sqref)=")([^"]+)(")',
        replace,
        xml_fragment,
    )


def expand_invoice_product_rows(
    xml_bytes,
    extra_rows,
):
    """
    Insert extra formatted product rows into the Invoice worksheet.

    The existing template has product rows 27-29. For every product
    beyond the third, clone row 29, insert it immediately after the
    existing product block, then shift all original rows/references
    from row 30 downward. This preserves all unrelated workbook XML.
    """
    if extra_rows <= 0:
        return xml_bytes

    xml_text = xml_bytes.decode(
        "utf-8"
    )

    sheet_data_match = re.search(
        r'(<(?:[A-Za-z_][\w.\-]*:)?sheetData>)(.*?)(</(?:[A-Za-z_][\w.\-]*:)?sheetData>)',
        xml_text,
        re.DOTALL,
    )

    if not sheet_data_match:
        raise ValueError(
            "Invoice worksheet sheetData was not found."
        )

    before_sheet_data = xml_text[
        :sheet_data_match.start()
    ]
    sheet_data_open = sheet_data_match.group(1)
    sheet_data_body = sheet_data_match.group(2)
    sheet_data_close = sheet_data_match.group(3)
    after_sheet_data = xml_text[
        sheet_data_match.end():
    ]

    template_row_match = re.search(
        r'<(?:[A-Za-z_][\w.\-]*:)?row\b[^>]*\br="29"[^>]*>.*?</(?:[A-Za-z_][\w.\-]*:)?row>',
        sheet_data_body,
        re.DOTALL,
    )

    if not template_row_match:
        raise ValueError(
            "Invoice product template row 29 was not found."
        )

    template_row_xml = (
        template_row_match.group(0)
    )

    row_pattern = re.compile(
        r'<(?:[A-Za-z_][\w.\-]*:)?row\b[^>]*>.*?</(?:[A-Za-z_][\w.\-]*:)?row>',
        re.DOTALL,
    )

    shifted_body = row_pattern.sub(
        lambda match: _shift_invoice_row_xml(
            match.group(0),
            INVOICE_STATIC_START_ROW,
            extra_rows,
        ),
        sheet_data_body,
    )

    extra_row_xml = "".join(
        _move_invoice_row_xml(
            template_row_xml,
            29,
            INVOICE_STATIC_START_ROW
            + offset,
        )
        for offset in range(
            extra_rows
        )
    )

    row_29_match = re.search(
        r'<(?:[A-Za-z_][\w.\-]*:)?row\b[^>]*\br="29"[^>]*>.*?</(?:[A-Za-z_][\w.\-]*:)?row>',
        shifted_body,
        re.DOTALL,
    )

    if not row_29_match:
        raise ValueError(
            "Invoice product row 29 disappeared during expansion."
        )

    shifted_body = (
        shifted_body[:row_29_match.end()]
        + extra_row_xml
        + shifted_body[row_29_match.end():]
    )

    before_sheet_data = (
        _shift_reference_attributes(
            before_sheet_data,
            INVOICE_STATIC_START_ROW,
            extra_rows,
        )
    )

    after_sheet_data = (
        _shift_reference_attributes(
            after_sheet_data,
            INVOICE_STATIC_START_ROW,
            extra_rows,
        )
    )

    # Every product row has the same three merged regions as row 29.
    product_merge_columns = (
        ("D", "H"),
        ("K", "L"),
        ("N", "O"),
    )

    new_merge_tags = "".join(
        f'<mergeCell ref="{start}{row}:{end}{row}"/>'
        for row in range(
            INVOICE_STATIC_START_ROW,
            INVOICE_STATIC_START_ROW
            + extra_rows,
        )
        for start, end in product_merge_columns
    )

    merge_cells_match = re.search(
        r'(<(?:[A-Za-z_][\w.\-]*:)?mergeCells\b[^>]*\bcount=")(\d+)("[^>]*>)(.*?)(</(?:[A-Za-z_][\w.\-]*:)?mergeCells>)',
        after_sheet_data,
        re.DOTALL,
    )

    if not merge_cells_match:
        raise ValueError(
            "Invoice worksheet mergeCells block was not found."
        )

    merge_count = (
        int(merge_cells_match.group(2))
        + (
            len(product_merge_columns)
            * extra_rows
        )
    )

    replacement_merge_cells = (
        merge_cells_match.group(1)
        + str(merge_count)
        + merge_cells_match.group(3)
        + merge_cells_match.group(4)
        + new_merge_tags
        + merge_cells_match.group(5)
    )

    after_sheet_data = (
        after_sheet_data[:merge_cells_match.start()]
        + replacement_merge_cells
        + after_sheet_data[merge_cells_match.end():]
    )

    expanded_xml = (
        before_sheet_data
        + sheet_data_open
        + shifted_body
        + sheet_data_close
        + after_sheet_data
    )

    return expanded_xml.encode(
        "utf-8"
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

    extra_product_rows = max(
        0,
        len(line_items)
        - INVOICE_TEMPLATE_PRODUCT_ROW_COUNT,
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
        ] = (
            clean(
                company.get(
                    "iec"
                )
            )
            or "AAGCD8689R"
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
        ] = (
            clean(
                company.get(
                    "gstin"
                )
            )
            or "29AAGCD8689R1ZB"
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
        ] = (
            clean(
                bank.get(
                    "code"
                )
            )
            or "6361504-8400009"
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
        ] = (
            clean(
                bank.get(
                    "account_number"
                )
            )
            or "926030028134123"
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
        ] = (
            clean(
                bank.get(
                    "ifsc"
                )
            )
            or "UTIB0001541"
        )

    # If more than three products are present, the original
    # totals/bank/packing section is moved down by the same
    # number of newly inserted product rows. Shift only those
    # static Invoice-sheet update references here.
    if extra_product_rows:
        invoice_updates = {
            shift_cell_reference(
                reference,
                INVOICE_STATIC_START_ROW,
                extra_product_rows,
            ): change
            for reference, change
            in invoice_updates.items()
        }

    # ========================================================
    # CLEAR PRODUCT ROWS
    # ========================================================

    product_row_count = max(
        INVOICE_TEMPLATE_PRODUCT_ROW_COUNT,
        len(line_items),
    )

    for product_index in range(
        product_row_count
    ):
        slot = invoice_product_slot(
            product_index
        )
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
        slot = invoice_product_slot(
            index
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
                "We here by declare that we are the SELLER / CONSIGNOR of the goods\\n"
                "in Respect of which this declaration made and that particulars given to shipping Bill\\n"
                f"No: {shipping_bill or ''} Date {shipping_date_slash or ''} are true and that :"
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

    # AlumSteel short product summary for downloaded Excel.
    # Keep this consistent with the website preview:
    # Product :-
    # Part # <part number>
    # Aluminium / Steel
    alum_product_entries = []

    for index, item in enumerate(line_items):

        part = clean(
            item.get("part_number")
        )

        material_source = " ".join(
            value
            for value in [
                clean(item.get("material")),
                clean(item.get("description")),
            ]
            if value
        )

        material_lower = (
            material_source.lower()
        )

        material_type = ""

        if any(
            token in material_lower
            for token in [
                "aluminium",
                "aluminum",
                "en aw-",
                "en aw ",
                "aw-6061",
                "aw 6061",
                "aw-6082",
                "aw 6082",
                "aw-7075",
                "aw 7075",
                "aw-5083",
                "aw 5083",
                "al-mg",
                "almg",
            ]
        ):
            material_type = "Aluminium"

        elif any(
            token in material_lower
            for token in [
                "stainless steel",
                "steel",
                "aisi",
                "sus ",
                "ss304",
                "ss 304",
                "ss316",
                "ss 316",
                "304l",
                "316l",
                "17-4ph",
                "17-4 ph",
                "42crmo4",
                "s235",
                "s275",
                "s355",
            ]
        ):
            material_type = "Steel"

        lines = []

        if part:
            lines.append(
                f"Part # {part}"
            )

        if material_type:
            lines.append(
                material_type
            )

        if lines:
            alum_product_entries.append(
                "\n".join(lines)
            )

    alum_product_text = "Product :-"

    if alum_product_entries:
        alum_product_text += (
            "\n"
            + "\n\n".join(
                alum_product_entries
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
            alum_product_text,

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
                f"{' '.join(first_title.split())}"
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
                "Product description:-"
                + chr(10)
                + chr(10).join(
                    " ".join(
                        value.replace("\\n", " ").split()
                    )
                    for value in all_product_descriptions
                    if value
                )
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


# DOC_PORT_SMART_ROW_HEIGHT_V2
# ============================================================
# CONTENT-AWARE GENERATED EXCEL ROW HEIGHTS
#
# Keeps workbook compact for A4 while preventing text clipping.
# It does not alter:
# - wording
# - borders
# - columns
# - formulas
# - merged cells
# - website preview
# - print-area settings
# ============================================================


def _smart_value(change):
    if isinstance(change, dict):
        value = change.get("value")
    else:
        value = change

    if value is None:
        return ""

    return (
        str(value)
        .replace("\\n", "\n")
        .replace("\r\n", "\n")
        .replace("\r", "\n")
        .strip()
    )


def _smart_estimate_lines(
    value,
    chars_per_line,
):
    value = _smart_value(value)

    if not value:
        return 1

    chars_per_line = max(
        15,
        int(chars_per_line),
    )

    total = 0

    for part in value.split("\n"):
        part = part.strip()

        if not part:
            total += 1
            continue

        total += max(
            1,
            (
                len(part)
                + chars_per_line
                - 1
            )
            // chars_per_line,
        )

    return max(
        1,
        total,
    )


def _smart_height(
    value,
    chars_per_line,
    minimum,
    maximum,
    line_height=15.0,
    padding=9.0,
):
    lines = _smart_estimate_lines(
        value,
        chars_per_line,
    )

    height = (
        lines * float(line_height)
        + float(padding)
    )

    return max(
        float(minimum),
        min(
            float(maximum),
            float(height),
        ),
    )


def _smart_set_row_height(
    xml_bytes,
    row_number,
    height,
    hidden=False,
):
    """
    Modify ONLY row height/hidden state in worksheet XML.
    """

    xml_text = xml_bytes.decode(
        "utf-8"
    )

    worksheet_match = re.search(
        r"<(?P<prefix>[A-Za-z_][\w.\-]*:)?worksheet\b",
        xml_text,
    )

    prefix = (
        worksheet_match.group("prefix")
        if worksheet_match
        else ""
    ) or ""

    row_tag = re.escape(
        prefix + "row"
    )

    pattern = re.compile(
        rf'(<{row_tag}\b'
        rf'(?=[^>]*\br="{int(row_number)}")'
        rf'[^>]*>)'
    )

    match = pattern.search(
        xml_text
    )

    if not match:
        return xml_bytes

    tag = match.group(1)

    # Remove ONLY height/hidden attributes.
    tag = re.sub(
        r'\s+(?:ht|customHeight|hidden)="[^"]*"',
        "",
        tag,
    )

    tag = tag[:-1]

    tag += (
        f' ht="{float(height):.2f}"'
        ' customHeight="1"'
    )

    if hidden:
        tag += ' hidden="1"'

    tag += ">"

    xml_text = (
        xml_text[:match.start()]
        + tag
        + xml_text[match.end():]
    )

    return xml_text.encode(
        "utf-8"
    )


def optimize_generated_row_heights(
    xml_bytes,
    sheet_name,
    changes,
    payload,
):
    """
    Apply compact content-aware row heights only where needed.
    """

    # ========================================================
    # FORM SDF + EVD
    #
    # Already manually finalized. Do not touch.
    # ========================================================

    if sheet_name in (
        SHEETS.get("sdf"),
        SHEETS.get("evd"),
    ):
        return xml_bytes


    # ========================================================
    # INVOICE
    # ========================================================

    if sheet_name == SHEETS["invoice"]:

        extracted = payload.get(
            "extracted"
        )

        if not isinstance(
            extracted,
            dict,
        ):
            extracted = payload

        items = get_line_items(
            extracted
        )

        product_count = len(
            items
        )

        # ----------------------------------------------------
        # ACTIVE PRODUCT ROWS
        #
        # 1 product = 1 compact row
        # 4 products = 4 individually-sized rows
        # etc.
        # ----------------------------------------------------

        for index in range(
            product_count
        ):

            row = (
                INVOICE_PRODUCT_START_ROW
                + index
            )

            description_ref = (
                f"D{row}"
            )

            description = (
                changes.get(
                    description_ref
                )
            )

            height = _smart_height(
                description,
                chars_per_line=46,
                minimum=50,
                maximum=160,
                line_height=17,
                padding=18,
            )

            xml_bytes = (
                _smart_set_row_height(
                    xml_bytes,
                    row,
                    height,
                    hidden=False,
                )
            )

        # ----------------------------------------------------
        # UNUSED ORIGINAL TEMPLATE PRODUCT ROWS
        #
        # Example:
        # 1 product => rows 28 and 29 consume NO print space.
        # ----------------------------------------------------

        if (
            product_count
            < INVOICE_TEMPLATE_PRODUCT_ROW_COUNT
        ):

            for row in range(
                (
                    INVOICE_PRODUCT_START_ROW
                    + product_count
                ),
                (
                    INVOICE_PRODUCT_START_ROW
                    + INVOICE_TEMPLATE_PRODUCT_ROW_COUNT
                ),
            ):

                xml_bytes = (
                    _smart_set_row_height(
                        xml_bytes,
                        row,
                        2.0,
                        hidden=True,
                    )
                )

        return xml_bytes


    # ========================================================
    # DHL EXPRESS SLI
    #
    # Resize only dynamically generated long values.
    # Existing fixed form structure stays intact.
    # ========================================================

    if sheet_name == SHEETS["dhl"]:

        row_requirements = {}

        for reference, change in (
            changes.items()
        ):

            value = _smart_value(
                change
            )

            if not value:
                continue

            row_match = re.search(
                r"(\d+)$",
                str(reference),
            )

            if not row_match:
                continue

            row = int(
                row_match.group(1)
            )

            # Short values such as invoice no/date/weight
            # already fit correctly.
            if (
                len(value) < 45
                and "\n" not in value
            ):
                continue

            height = _smart_height(
                value,
                chars_per_line=58,
                minimum=24,
                maximum=74,
                line_height=15,
                padding=8,
            )

            row_requirements[row] = max(
                row_requirements.get(
                    row,
                    0,
                ),
                height,
            )

        for row, height in (
            row_requirements.items()
        ):
            xml_bytes = (
                _smart_set_row_height(
                    xml_bytes,
                    row,
                    height,
                    hidden=False,
                )
            )

        # DOC_PORT_DHL_EXACT_HEIGHT_OVERRIDE_V2
        # Final compact DHL heights.
        # These run AFTER normal content-aware sizing so the
        # known oversized areas cannot become large again.
        exact_dhl_heights = {
            13: 15.0,
            14: 15.0,
            15: 15.0,
            16: 15.0,
            18: 48.0,
            44: 34.0,
            45: 28.0,
            46: 36.0,
        }

        for exact_row, exact_height in exact_dhl_heights.items():
            xml_bytes = _smart_set_row_height(
                xml_bytes,
                exact_row,
                exact_height,
                hidden=False,
            )

        return xml_bytes


    # ========================================================
    # ALUMINIUM / STEEL DERIVATIVES
    # ========================================================

    if sheet_name == SHEETS["alum"]:

        row_requirements = {}

        for reference, change in (
            changes.items()
        ):

            value = _smart_value(
                change
            )

            if not value:
                continue

            row_match = re.search(
                r"(\d+)$",
                str(reference),
            )

            if not row_match:
                continue

            row = int(
                row_match.group(1)
            )

            if (
                len(value) < 45
                and "\n" not in value
            ):
                continue

            height = _smart_height(
                value,
                chars_per_line=82,
                minimum=24,
                maximum=78,
                line_height=15,
                padding=9,
            )

            row_requirements[row] = max(
                row_requirements.get(
                    row,
                    0,
                ),
                height,
            )

        for row, height in (
            row_requirements.items()
        ):
            xml_bytes = (
                _smart_set_row_height(
                    xml_bytes,
                    row,
                    height,
                    hidden=False,
                )
            )

        # DOC_PORT_ALUM_PRODUCT_DYNAMIC_HEIGHT_V1
        # ----------------------------------------------------
        # ALUMINUM / STEEL DERIVATIVES PRODUCT BOX
        #
        # Layout:
        #
        # Product :-
        # Part # XXXXX
        # Aluminium / Steel
        #
        # Part # XXXXX
        # Aluminium / Steel
        #
        # The product row grows according to actual product
        # count so no product is cut in Excel or A4/PDF output.
        # ----------------------------------------------------

        extracted = payload.get("extracted")

        if not isinstance(extracted, dict):
            extracted = payload

        alum_items = get_line_items(
            extracted
        )

        alum_product_count = max(
            1,
            len(alum_items),
        )

        # Required visual lines:
        #
        # 1 heading line
        # + 2 lines per product
        # + 1 blank separator between products
        #
        # This simplifies to approximately 3 lines per product.
        required_alum_lines = (
            3 * alum_product_count
        )

        alum_product_height = (
            required_alum_lines * 15.0
            + 10.0
        )

        # Minimum keeps a 1-product document neat.
        alum_product_height = max(
            55.0,
            alum_product_height,
        )

        xml_bytes = _smart_set_row_height(
            xml_bytes,
            2,
            alum_product_height,
            hidden=False,
        )

        return xml_bytes


    # ========================================================
    # SCOMET DECLARATION
    # ========================================================

    if sheet_name == SHEETS["scomet"]:

        # Product description row C5:H5.
        product_value = (
            changes.get(
                CELLS["scomet"]["product"]
            )
        )

        if product_value:

            product_height = (
                _smart_height(
                    product_value,
                    chars_per_line=82,
                    minimum=42,
                    maximum=165,
                    line_height=15,
                    padding=10,
                )
            )

            xml_bytes = (
                _smart_set_row_height(
                    xml_bytes,
                    5,
                    product_height,
                    hidden=False,
                )
            )

        # Static declaration wording:
        # compact enough for A4 but enough to avoid clipping.
        compact_rows = {
            8: 58.0,
            9: 82.0,
            10: 68.0,
            11: 54.0,
            13: 34.0,
            14: 34.0,
        }

        for row, height in (
            compact_rows.items()
        ):

            xml_bytes = (
                _smart_set_row_height(
                    xml_bytes,
                    row,
                    height,
                    hidden=False,
                )
            )

        return xml_bytes


    return xml_bytes





# DOC_PORT_DYNAMIC_INVOICE_PRINT_AREA
def update_dynamic_invoice_print_area(
    workbook_path,
    payload,
):
    """
    Keep Invoice Print_Area synchronized with dynamically inserted
    product rows.

    <= 3 products : A1:O41
    4 products    : A1:O42
    5 products    : A1:O43
    etc.
    """

    extracted = payload.get("extracted")

    if not isinstance(extracted, dict):
        extracted = payload

    product_count = len(
        get_line_items(extracted)
    )

    extra_rows = max(
        0,
        product_count
        - INVOICE_TEMPLATE_PRODUCT_ROW_COUNT,
    )

    last_row = 41 + extra_rows

    workbook_path = Path(
        workbook_path
    )

    temp_path = workbook_path.with_name(
        workbook_path.stem
        + ".printarea.tmp.xlsx"
    )

    found_invoice_area = False

    with zipfile.ZipFile(
        workbook_path,
        "r",
    ) as source_zip:

        with zipfile.ZipFile(
            temp_path,
            "w",
            compression=zipfile.ZIP_DEFLATED,
        ) as output_zip:

            for info in source_zip.infolist():

                data = source_zip.read(
                    info.filename
                )

                if (
                    info.filename
                    == "xl/workbook.xml"
                ):

                    xml_text = data.decode(
                        "utf-8"
                    )

                    pattern = re.compile(
                        r'('
                        r'<(?:[A-Za-z_][\w.\-]*:)?definedName\b'
                        r'[^>]*'
                        r'name="_xlnm\.Print_Area"'
                        r'[^>]*>'
                        r')'
                        r'([^<]*)'
                        r'('
                        r'</(?:[A-Za-z_][\w.\-]*:)?definedName>'
                        r')'
                    )

                    def replace_area(match):
                        nonlocal found_invoice_area

                        formula = match.group(2)

                        if not (
                            "Invoice!"
                            in formula
                            or "'Invoice'!"
                            in formula
                        ):
                            return match.group(0)

                        new_formula, count = re.subn(
                            r'\$A\$1:\$O\$\d+',
                            (
                                "$A$1:$O$"
                                + str(last_row)
                            ),
                            formula,
                            count=1,
                        )

                        if count != 1:
                            return match.group(0)

                        found_invoice_area = True

                        return (
                            match.group(1)
                            + new_formula
                            + match.group(3)
                        )

                    xml_text = pattern.sub(
                        replace_area,
                        xml_text,
                    )

                    data = xml_text.encode(
                        "utf-8"
                    )

                output_zip.writestr(
                    info,
                    data,
                )

    if not found_invoice_area:

        try:
            temp_path.unlink()
        except FileNotFoundError:
            pass

        raise ValueError(
            "Invoice Print_Area definition was not found."
        )

    temp_path.replace(
        workbook_path
    )





# DOC_PORT_INVOICE_PRINTAREA_FINAL_V2
def ensure_dynamic_invoice_print_area(
    workbook_path,
    payload,
):
    """
    Keep the Invoice printable area synchronized with dynamically
    inserted product rows.

    Base:
        1-3 products -> A1:O41

    Additional products:
        4 products   -> A1:O42
        5 products   -> A1:O43
        etc.

    Only workbook Print_Area metadata is changed.
    """

    extracted = payload.get(
        "extracted"
    )

    if not isinstance(
        extracted,
        dict,
    ):
        extracted = payload

    product_count = len(
        get_line_items(
            extracted
        )
    )

    extra_product_rows = max(
        0,
        product_count
        - INVOICE_TEMPLATE_PRODUCT_ROW_COUNT,
    )

    last_print_row = (
        41
        + extra_product_rows
    )

    workbook_path = Path(
        workbook_path
    )

    temp_path = workbook_path.with_name(
        workbook_path.stem
        + ".printarea-final.tmp.xlsx"
    )

    invoice_area_found = False

    with zipfile.ZipFile(
        workbook_path,
        "r",
    ) as source_zip:

        with zipfile.ZipFile(
            temp_path,
            "w",
            compression=zipfile.ZIP_DEFLATED,
        ) as target_zip:

            for info in source_zip.infolist():

                data = source_zip.read(
                    info.filename
                )

                if (
                    info.filename
                    == "xl/workbook.xml"
                ):

                    xml_text = data.decode(
                        "utf-8"
                    )

                    pattern = re.compile(
                        r'('
                        r'<(?:[A-Za-z_][\w.\-]*:)?definedName\b'
                        r'[^>]*'
                        r'\bname="_xlnm\.Print_Area"'
                        r'[^>]*>'
                        r')'
                        r'(.*?)'
                        r'('
                        r'</(?:[A-Za-z_][\w.\-]*:)?definedName>'
                        r')',
                        re.DOTALL,
                    )

                    def replace_print_area(match):
                        nonlocal invoice_area_found

                        formula = match.group(2)

                        if not re.search(
                            r"(?:'Invoice'|Invoice)!",
                            formula,
                        ):
                            return match.group(0)

                        new_formula, count = re.subn(
                            r"\$A\$1:\$O\$\d+",
                            (
                                "$A$1:$O$"
                                + str(last_print_row)
                            ),
                            formula,
                            count=1,
                        )

                        if count != 1:
                            return match.group(0)

                        invoice_area_found = True

                        return (
                            match.group(1)
                            + new_formula
                            + match.group(3)
                        )

                    xml_text = pattern.sub(
                        replace_print_area,
                        xml_text,
                    )

                    data = xml_text.encode(
                        "utf-8"
                    )

                target_zip.writestr(
                    info,
                    data,
                )

    if not invoice_area_found:

        try:
            temp_path.unlink()
        except FileNotFoundError:
            pass

        raise ValueError(
            "Invoice Print_Area could not be located "
            "inside workbook.xml."
        )

    # ========================================================
    # VERIFY BEFORE REPLACING ORIGINAL
    # ========================================================

    expected = (
        "$A$1:$O$"
        + str(last_print_row)
    )

    with zipfile.ZipFile(
        temp_path,
        "r",
    ) as check_zip:

        workbook_xml = check_zip.read(
            "xl/workbook.xml"
        ).decode(
            "utf-8"
        )

        if expected not in workbook_xml:
            raise ValueError(
                "Dynamic Invoice PrintArea verification failed. "
                f"Expected {expected}."
            )

    temp_path.replace(
        workbook_path
    )




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

    line_items = get_line_items(
        extracted_data
    )

    extra_product_rows = max(
        0,
        len(line_items)
        - INVOICE_TEMPLATE_PRODUCT_ROW_COUNT,
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

            # The Invoice master starts with three formatted product
            # rows. Add more rows only when the uploaded invoice has
            # more than three products. All other sheets are untouched.
            if (
                sheet_name
                == SHEETS["invoice"]
                and extra_product_rows > 0
            ):
                original_xml = (
                    expand_invoice_product_rows(
                        original_xml,
                        extra_product_rows,
                    )
                )

            patched_xml = (
                patch_sheet(
                    original_xml,
                    changes,
                )
            )

            patched_xml = (
                optimize_generated_row_heights(
                    patched_xml,
                    sheet_name,
                    changes,
                    payload,
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

    update_dynamic_invoice_print_area(
        output_path,
        payload,
    )

    ensure_dynamic_invoice_print_area(
        output_path,
        payload,
    )

    validate_generated_xlsx(
        output_path
    )

    return output_path

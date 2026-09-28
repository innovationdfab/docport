from pathlib import Path
from datetime import datetime
import tempfile
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
    Path(tempfile.gettempdir())
    / "docport-generated"
)

GENERATED_DIR.mkdir(
    parents=True,
    exist_ok=True,
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

    value = str(value).strip()

    return value if value else None


def as_float(value):
    if value is None or value == "":
        return None

    try:
        return float(
            str(value)
            .replace(",", "")
            .replace("$", "")
            .replace("₹", "")
            .replace("€", "")
            .replace("£", "")
            .replace("kg", "")
            .strip()
        )

    except Exception:
        return None


def as_int(value):
    number = as_float(value)

    if number is None:
        return None

    return int(number)


# ============================================================
# DATE HELPERS
# ============================================================

def parse_date(value):
    if not value:
        return None

    text = str(value).strip()

    formats = (
        "%d.%m.%Y",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d",
    )

    for fmt in formats:
        try:
            return datetime.strptime(
                text,
                fmt,
            )

        except ValueError:
            continue

    return None


def date_slash(value):
    parsed = parse_date(value)

    if parsed:
        return parsed.strftime(
            "%d/%m/%Y"
        )

    return clean(value) or ""


def date_dash(value):
    parsed = parse_date(value)

    if parsed:
        return parsed.strftime(
            "%d-%m-%Y"
        )

    return clean(value) or ""


# ============================================================
# PRODUCT HELPERS
# ============================================================

def get_line_items(data):
    product = (
        data.get("product", {})
        or {}
    )

    items = product.get(
        "line_items",
        [],
    )

    if not isinstance(
        items,
        list,
    ):
        return []

    return items


def extract_product_blocks_from_pdf(payload):
    raw_text = (
        payload.get(
            "raw_text",
            "",
        )
        or ""
    )

    if not raw_text:
        return []

    pattern = re.compile(
        r"Part\s*#\s*"
        r"([A-Za-z0-9._/\-]+)"
        r"\s*,?\s*(.*?)"
        r"(?=\n\s*\d+(?:\.\d+)?\s+"
        r"\d{6,12}\s+"
        r"[A-Za-z][A-Za-z\s]*?\s+"
        r"\d+\s+"
        r"[$€£₹]?\s*[\d,.]+\s+"
        r"[$€£₹]?\s*[\d,.]+)",
        re.IGNORECASE
        | re.DOTALL,
    )

    blocks = []

    for match in pattern.finditer(
        raw_text
    ):
        part_number = clean(
            match.group(1)
        )

        body = (
            match.group(2)
            .strip()
        )

        material = None
        description = body

        material_match = re.search(
            r"Material\s*:\s*(.*)$",
            body,
            re.IGNORECASE
            | re.DOTALL,
        )

        if material_match:
            material = re.sub(
                r"\s*\n\s*",
                " ",
                material_match
                .group(1)
                .strip(),
            )

            description = (
                body[
                    : material_match.start()
                ]
                .strip()
            )

        description = re.sub(
            r"\n\s*",
            "\n",
            description,
        ).strip(
            " ,\n"
        )

        blocks.append(
            {
                "part_number":
                    part_number,

                "description":
                    description
                    or None,

                "material":
                    clean(material),
            }
        )

    return blocks


def build_product_description(
    item,
    product_block=None,
):
    product_block = (
        product_block
        or {}
    )

    part_number = (
        clean(
            item.get(
                "part_number"
            )
        )
        or
        clean(
            product_block.get(
                "part_number"
            )
        )
    )

    description = (
        clean(
            item.get(
                "description"
            )
        )
        or
        clean(
            product_block.get(
                "description"
            )
        )
    )

    material = (
        clean(
            item.get(
                "material"
            )
        )
        or
        clean(
            product_block.get(
                "material"
            )
        )
    )

    lines = []

    if part_number:
        lines.append(
            f"Part #{part_number}"
        )

    if description:
        lower_description = (
            description.lower()
        )

        part_text = (
            f"part #{part_number}"
            .lower()
            if part_number
            else ""
        )

        if (
            not part_number
            or
            part_text
            not in lower_description
        ):
            lines.append(
                description
            )

    if material:
        if (
            not description
            or
            "material:"
            not in description.lower()
        ):
            lines.append(
                f"Material: {material}"
            )

    result = "\n".join(
        lines
    ).strip()

    return result or None


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

        quantity = as_int(
            item.get(
                "quantity"
            )
        )

        if quantity is not None:
            total_quantity += (
                quantity
            )
            has_quantity = True

        weight = as_float(
            item.get(
                "net_weight_kg"
            )
        )

        if weight is not None:
            total_net_weight += (
                weight
            )
            has_weight = True

        position_price = as_float(
            item.get(
                "position_price"
            )
        )

        if position_price is None:

            unit_price = as_float(
                item.get(
                    "unit_price"
                )
            )

            if (
                quantity is not None
                and
                unit_price is not None
            ):
                position_price = (
                    quantity
                    * unit_price
                )

        if position_price is not None:
            total_value += (
                position_price
            )
            has_value = True

    return {
        "quantity":
            total_quantity
            if has_quantity
            else None,

        "net_weight":
            round(
                total_net_weight,
                3,
            )
            if has_weight
            else None,

        "total_value":
            round(
                total_value,
                2,
            )
            if has_value
            else None,
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

    symbols = {
        "USD": "$",
        "EUR": "€",
        "GBP": "£",
        "INR": "₹",
    }

    return symbols.get(
        currency,
        "",
    )


# ============================================================
# SAFE OPTIONAL MAPPING
# ============================================================

def add_optional_mapping(
    target,
    mapping,
    key,
    value,
):
    """
    Adds a cell only when the key
    exists inside field_mapping.py.

    This prevents KeyErrors for
    template versions that do not
    contain the optional cell.
    """

    cell = mapping.get(
        key
    )

    if not cell:
        return

    target[
        cell
    ] = value


# ============================================================
# BUILD ALL WORKBOOK UPDATES
# ============================================================

def build_excel_updates(
    payload
):

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


    # ========================================================
    # PRODUCTS
    # ========================================================

    line_items = get_line_items(
        data
    )

    if not line_items:
        raise ValueError(
            "No product line items were extracted from the PDF."
        )

    if (
        len(line_items)
        > MAX_PRODUCTS
    ):
        raise ValueError(
            f"The PDF contains {len(line_items)} products, "
            f"but this Excel template supports only "
            f"{MAX_PRODUCTS} product rows."
        )


    product_blocks = (
        extract_product_blocks_from_pdf(
            payload
        )
    )

    blocks_by_part = {}

    for block in (
        product_blocks
    ):

        part = clean(
            block.get(
                "part_number"
            )
        )

        if part:
            blocks_by_part[
                part
            ] = block


    # ========================================================
    # INVOICE DETAILS
    # ========================================================

    invoice_number = clean(
        invoice.get(
            "number"
        )
    )

    invoice_date = clean(
        invoice.get(
            "date"
        )
    )

    invoice_date_slash = (
        date_slash(
            invoice_date
        )
    )

    invoice_date_dash = (
        date_dash(
            invoice_date
        )
    )

    po_number = clean(
        invoice.get(
            "po_number"
        )
    )


    # ========================================================
    # CONSIGNEE
    # ========================================================

    consignee_name = (
        clean(
            consignee.get(
                "name"
            )
        )
        or
        clean(
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
        or
        clean(
            consignee.get(
                "address"
            )
        )
    )


    # ========================================================
    # SHIPMENT
    # ========================================================

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

    shipping_bill_date = (
        clean(
            shipment.get(
                "shipping_bill_date"
            )
        )
        or
        invoice_date
    )

    shipping_date_slash = (
        date_slash(
            shipping_bill_date
        )
    )


    # ========================================================
    # CURRENCY
    # ========================================================

    currency = (
        clean(
            shipment.get(
                "currency"
            )
        )
        or
        "USD"
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
        totals.get(
            "quantity"
        )
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
            totals.get(
                "net_weight"
            )
        )


    total_value = as_float(
        shipment.get(
            "total_amount"
        )
    )

    if total_value is None:
        total_value = (
            totals.get(
                "total_value"
            )
        )


    # ========================================================
    # INR VALUE
    # ========================================================

    amount_inr = None

    if (
        total_value is not None
        and
        exchange_rate is not None
    ):
        amount_inr = round(
            total_value
            * exchange_rate,
            2,
        )


    # Website editable override
    amount_inr_override = (
        as_float(
            shipment.get(
                "amount_inr_override"
            )
        )
    )

    if (
        amount_inr_override
        is not None
    ):
        amount_inr = (
            amount_inr_override
        )


    # ========================================================
    # COUNTRY OF ORIGIN
    # ========================================================

    countries = []

    for item in (
        line_items
    ):

        country = clean(
            item.get(
                "country_of_origin"
            )
        )

        if (
            country
            and
            country
            not in countries
        ):
            countries.append(
                country
            )

    country_origin = (
        ", ".join(
            countries
        )
        if countries
        else
        "INDIA"
    )


    # ========================================================
    # PRODUCT DESCRIPTION BLOCKS
    # ========================================================

    all_product_descriptions = []

    for (
        index,
        item,
    ) in enumerate(
        line_items
    ):

        part = clean(
            item.get(
                "part_number"
            )
        )

        block = (
            blocks_by_part.get(
                part
            )
            if part
            else None
        )

        if (
            block is None
            and
            index
            < len(
                product_blocks
            )
        ):
            block = (
                product_blocks[
                    index
                ]
            )

        description = (
            build_product_description(
                item,
                block,
            )
        )

        if not description:
            continue

        if (
            len(line_items)
            > 1
        ):
            all_product_descriptions.append(
                f"Product {index + 1}: {description}"
            )

        else:
            all_product_descriptions.append(
                description
            )


    all_products_text = (
        "\n\n".join(
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
        or
        first_part
        or
        ""
    )


    updates = {}


    # ========================================================
    # 1. INVOICE
    # ========================================================

    invoice_map = (
        CELLS[
            "invoice"
        ]
    )


    invoice_updates = {

        invoice_map[
            "invoice_date"
        ]:
            (
                f"Date-{invoice_date_slash}"
                if invoice_date_slash
                else ""
            ),

        invoice_map[
            "invoice_number"
        ]:
            invoice_number,

        invoice_map[
            "order_date"
        ]:
            (
                f"Date-{invoice_date_slash}"
                if invoice_date_slash
                else ""
            ),

        invoice_map[
            "po_number"
        ]:
            po_number,

        invoice_map[
            "consignee_label"
        ]:
            (
                "Ship to / Consignee: "
                f"{consignee_name or ''}"
            ),

        invoice_map[
            "consignee_address"
        ]:
            consignee_raw,

        invoice_map[
            "consignee_contact"
        ]:
            (
                "Ctc person- "
                f"{consignee_name or ''}"
            ),

        invoice_map[
            "country_origin_1"
        ]:
            country_origin,

        invoice_map[
            "country_origin_2"
        ]:
            country_origin,

        invoice_map[
            "exchange_rate"
        ]:
            {
                "value":
                    exchange_rate,

                "kind":
                    "number",
            },

        invoice_map[
            "amount_inr"
        ]:
            {
                "value":
                    amount_inr,

                "kind":
                    "number",
            },

        invoice_map[
            "amount_usd"
        ]:
            {
                "value":
                    total_value,

                "kind":
                    "number",
            },

        invoice_map[
            "total_before_tax"
        ]:
            {
                "value":
                    total_value,

                "kind":
                    "number",
            },

        invoice_map[
            "total_after_tax"
        ]:
            {
                "value":
                    total_value,

                "kind":
                    "number",
            },

        invoice_map[
            "dimensions"
        ]:
            dimensions,

        invoice_map[
            "gross_weight"
        ]:
            (
                f"{package_weight:g} kg"

                if package_weight
                is not None

                else None
            ),

        invoice_map[
            "net_weight"
        ]:
            (
                f"{total_net_weight:g} kg"

                if total_net_weight
                is not None

                else None
            ),
    }


    # --------------------------------------------------------
    # OPTIONAL INVOICE FIELDS
    # --------------------------------------------------------

    add_optional_mapping(
        invoice_updates,
        invoice_map,
        "iec",
        clean(
            company.get(
                "iec"
            )
        ),
    )

    add_optional_mapping(
        invoice_updates,
        invoice_map,
        "gstin",
        clean(
            company.get(
                "gstin"
            )
        ),
    )

    add_optional_mapping(
        invoice_updates,
        invoice_map,
        "bank_ad_code",
        clean(
            bank.get(
                "code"
            )
        ),
    )

    add_optional_mapping(
        invoice_updates,
        invoice_map,
        "bank_account",
        clean(
            bank.get(
                "account_number"
            )
        ),
    )

    add_optional_mapping(
        invoice_updates,
        invoice_map,
        "bank_ifsc",
        clean(
            bank.get(
                "ifsc"
            )
        ),
    )


    # --------------------------------------------------------
    # OPTIONAL PACKAGE COUNT
    # --------------------------------------------------------

    package_count_cell = (
        invoice_map.get(
            "package_count"
        )
    )

    if package_count_cell:

        invoice_updates[
            package_count_cell
        ] = {
            "value":
                (
                    as_int(
                        shipment.get(
                            "package_count"
                        )
                    )
                    or
                    1
                ),

            "kind":
                "number",
        }


    # --------------------------------------------------------
    # CLEAR PRODUCT ROWS
    # --------------------------------------------------------

    for slot in (
        INVOICE_PRODUCT_ROWS
    ):

        for key in (
            "sr_no",
            "india_hsn",
            "us_hts",
            "description",
            "quantity",
            "unit_price",
            "taxable_value",
            "igst_percent",
            "igst_amount",
        ):

            cell = (
                slot.get(
                    key
                )
            )

            if cell:
                invoice_updates[
                    cell
                ] = None


    # --------------------------------------------------------
    # WRITE PRODUCT ROWS
    # --------------------------------------------------------

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

        part = clean(
            item.get(
                "part_number"
            )
        )

        block = (
            blocks_by_part.get(
                part
            )
            if part
            else None
        )

        if (
            block is None
            and
            index
            < len(
                product_blocks
            )
        ):
            block = (
                product_blocks[
                    index
                ]
            )

        description = (
            build_product_description(
                item,
                block,
            )
        )

        quantity = as_int(
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
            and
            quantity is not None
            and
            unit_price is not None
        ):
            position_price = round(
                quantity
                * unit_price,
                2,
            )


        india_hsn = (
            clean(
                item.get(
                    "india_hsn"
                )
            )
            or
            clean(
                item.get(
                    "india_hsn_code"
                )
            )
            or
            DEFAULT_INDIA_HSN
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


        serial_cell = (
            slot.get(
                "sr_no"
            )
        )

        if serial_cell:
            invoice_updates[
                serial_cell
            ] = {
                "value":
                    index + 1,

                "kind":
                    "number",
            }


        india_cell = (
            slot.get(
                "india_hsn"
            )
        )

        if india_cell:
            invoice_updates[
                india_cell
            ] = india_hsn


        hts_cell = (
            slot.get(
                "us_hts"
            )
        )

        if hts_cell:
            invoice_updates[
                hts_cell
            ] = us_hts


        description_cell = (
            slot.get(
                "description"
            )
        )

        if description_cell:
            invoice_updates[
                description_cell
            ] = description


        quantity_cell = (
            slot.get(
                "quantity"
            )
        )

        if quantity_cell:
            invoice_updates[
                quantity_cell
            ] = {
                "value":
                    quantity,

                "kind":
                    "number",
            }


        unit_cell = (
            slot.get(
                "unit_price"
            )
        )

        if unit_cell:

            invoice_updates[
                unit_cell
            ] = (
                f"{symbol}{unit_price:.2f}"

                if unit_price
                is not None

                else None
            )


        taxable_cell = (
            slot.get(
                "taxable_value"
            )
        )

        if taxable_cell:

            invoice_updates[
                taxable_cell
            ] = {
                "value":
                    position_price,

                "kind":
                    "number",
            }


        igst_percent_cell = (
            slot.get(
                "igst_percent"
            )
        )

        if igst_percent_cell:
            invoice_updates[
                igst_percent_cell
            ] = "0%"


        igst_amount_cell = (
            slot.get(
                "igst_amount"
            )
        )

        if igst_amount_cell:
            invoice_updates[
                igst_amount_cell
            ] = "0"


    updates[
        SHEETS[
            "invoice"
        ]
    ] = invoice_updates


    # ========================================================
    # 2. DHL EXPRESS SLI
    # ========================================================

    dhl_map = (
        CELLS[
            "dhl"
        ]
    )


    taxable_text = (
        "TAXABLE AMOUNT - "
        +
        (
            f"{amount_inr:.2f}"

            if amount_inr
            is not None

            else ""
        )
        +
        "\nIGST RATE - 0"
        +
        "\nIGST AMOUNT - 0"
        +
        "\nGST Compensation Cess - 0"
    )


    dhl_tax_override = clean(
        shipment.get(
            "dhl_tax_block_override"
        )
    )

    if dhl_tax_override:
        taxable_text = (
            dhl_tax_override
        )


    dhl_updates = {

        dhl_map[
            "invoice_number"
        ]:
            invoice_number,

        dhl_map[
            "consignee"
        ]:
            consignee_name,

        dhl_map[
            "invoice_date"
        ]:
            invoice_date_dash,

        dhl_map[
            "awb"
        ]:
            awb,

        dhl_map[
            "tax_block"
        ]:
            taxable_text,

        dhl_map[
            "fob_value"
        ]:
            (
                f"{total_value:.2f} {currency}"

                if total_value
                is not None

                else None
            ),

        dhl_map[
            "net_weight"
        ]:
            (
                f"{total_net_weight:g} kg"

                if total_net_weight
                is not None

                else None
            ),

        dhl_map[
            "gross_weight"
        ]:
            (
                f"{package_weight:g} kg"

                if package_weight
                is not None

                else None
            ),
    }


    # --------------------------------------------------------
    # OPTIONAL DHL FIELDS
    # --------------------------------------------------------

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
                or
                1
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

    sdf_map = (
        CELLS[
            "sdf"
        ]
    )


    sdf_updates = {

        sdf_map[
            "shipping_bill_header"
        ]:
            shipping_bill,

        sdf_map[
            "shipping_bill_declaration"
        ]:
            (
                "We here by declare that we are the "
                "SELLER / CONSIGNOR of the goods "
                "in Respect of which this declaration made "
                "and that particulars given to shipping Bill No: "
                f"{shipping_bill or ''} "
                f"Date {shipping_date_slash or ''} "
                "are true and that :"
            ),
    }


    # Current field_mapping.py does not
    # contain sdf.date, so use it only
    # if a future mapping adds it.

    sdf_date_cell = (
        sdf_map.get(
            "date"
        )
    )

    if sdf_date_cell:
        sdf_updates[
            sdf_date_cell
        ] = shipping_date_slash


    updates[
        SHEETS[
            "sdf"
        ]
    ] = sdf_updates


    # ========================================================
    # 4. EVD
    # ========================================================

    evd_map = (
        CELLS[
            "evd"
        ]
    )


    evd_updates = {

        evd_map[
            "shipping_bill"
        ]:
            (
                "1. Shipping Bill No. "
                f"{shipping_bill or ''}"
                " & Date:- "
                f"{shipping_date_slash or ''}"
            ),

        evd_map[
            "invoice_details"
        ]:
            (
                "2. Invoice No. & Date   "
                f"{invoice_number or ''}"
                " & "
                f"{invoice_date_slash or ''}"
            ),

        evd_map[
            "previous_shipping_bill"
        ]:
            (
                "Shipping Bill No: "
                f"{shipping_bill or ''}"
                " and date: "
                f"{shipping_date_slash or ''}"
            ),

        evd_map[
            "date"
        ]:
            (
                "Date: "
                f"{shipping_date_slash or invoice_date_slash}"
            ),
    }


    # Current mapping does not have
    # evd.invoice_date.
    # Use it only if available later.

    evd_invoice_date_cell = (
        evd_map.get(
            "invoice_date"
        )
    )

    if evd_invoice_date_cell:

        evd_updates[
            evd_invoice_date_cell
        ] = (
            "Date: "
            f"{invoice_date_slash or ''}"
        )


    updates[
        SHEETS[
            "evd"
        ]
    ] = evd_updates


    # ========================================================
    # 5. ALUMINIUM / STEEL
    # ========================================================

    alum_map = (
        CELLS[
            "alum"
        ]
    )


    full_product_weight = as_float(
        shipment.get(
            "full_product_weight_kg"
        )
    )

    if full_product_weight is None:
        full_product_weight = (
            total_net_weight
        )


    aluminium_weight = as_float(
        shipment.get(
            "aluminium_content_weight_kg"
        )
    )

    if aluminium_weight is None:
        aluminium_weight = (
            total_net_weight
        )


    aluminium_value = as_float(
        shipment.get(
            "aluminium_content_value"
        )
    )

    if aluminium_value is None:
        aluminium_value = (
            total_value
        )


    weight_text = (
        "3) Full weight of the product "
        +
        (
            f"{full_product_weight:g} kg"

            if full_product_weight
            is not None

            else ""
        )
        +
        "\nAluminum content weight "
        +
        (
            f"{aluminium_weight:g} kg"

            if aluminium_weight
            is not None

            else ""
        )
    )


    value_text = (
        "Total Value of the product "
        +
        (
            f"{total_value:.2f} {currency.lower()}"

            if total_value
            is not None

            else ""
        )
        +
        "\nValue of the Aluminum content "
        +
        (
            f"{aluminium_value:.2f} {currency.lower()}"

            if aluminium_value
            is not None

            else ""
        )
    )


    alum_product_override = clean(
        shipment.get(
            "alum_product_text_override"
        )
    )


    if alum_product_override:
        alum_product_text = (
            "Product:- "
            f"{alum_product_override}"
        )

    elif all_products_text:
        alum_product_text = (
            "Product:- "
            f"{all_products_text}"
        )

    else:
        alum_product_text = None


    updates[
        SHEETS[
            "alum"
        ]
    ] = {

        alum_map[
            "product"
        ]:
            alum_product_text,

        alum_map[
            "weight_block"
        ]:
            weight_text,

        alum_map[
            "value_block"
        ]:
            value_text,

        alum_map[
            "completed_date"
        ]:
            (
                "Completed by : "
                f"{invoice_date_dash or ''}"
            ),

        alum_map[
            "title"
        ]:
            (
                f"Title       : {first_title}"

                if first_title

                else "Title       :"
            ),

        alum_map[
            "date"
        ]:
            (
                "Date        : "
                f"{invoice_date_dash or ''}"
            ),
    }


    # ========================================================
    # 6. SCOMET DECLARATION
    # ========================================================

    scomet_map = (
        CELLS[
            "scomet"
        ]
    )


    unique_hsn = []

    for item in line_items:

        hsn = clean(
            item.get(
                "hs_code"
            )
        )

        if (
            hsn
            and
            hsn not in unique_hsn
        ):
            unique_hsn.append(
                hsn
            )


    hsn_text = ", ".join(
        unique_hsn
    )


    scomet_override = clean(
        shipment.get(
            "scomet_product_text_override"
        )
    )


    if scomet_override:
        scomet_product = (
            scomet_override
        )

    else:
        scomet_product = (
            "Product description:- "
            f"{all_products_text or ''}"
        )


    updates[
        SHEETS[
            "scomet"
        ]
    ] = {

        scomet_map[
            "invoice_number"
        ]:
            (
                "Invoice No: - "
                f"{invoice_number or ''}"
            ),

        scomet_map[
            "invoice_date"
        ]:
            (
                "Date: "
                f"{invoice_date_dash or ''}"
            ),

        scomet_map[
            "product"
        ]:
            scomet_product,

        scomet_map[
            "hsn_code"
        ]:
            (
                "HSN code: - "
                f"{hsn_text}"
            ),
    }


    return updates


# ============================================================
# XLSX SHEET PATHS
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
                "xl/_rels/workbook.xml.rels"
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
            "Unable to read worksheets from master workbook."
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
                f"Unable to resolve worksheet '{sheet_name}'."
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


# ============================================================
# XLSX CELL PATCHING
# ============================================================

def make_cell_pattern(
    reference
):

    return re.compile(
        (
            r'(<(?:[A-Za-z_][\w.\-]*:)?c\b'
            r'(?=[^>]*\br="'
            +
            re.escape(
                reference
            )
            +
            r'")'
            r'[^>]*?)'
            r'(?:/>|>(.*?)</(?:[A-Za-z_][\w.\-]*:)?c>)'
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
        r"<(?P<prefix>[A-Za-z_][\w.\-]*:)?c\b",
        start_tag,
    )

    if not match:
        return ""

    return (
        match.group(
            "prefix"
        )
        or ""
    )


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
            f"Cell {reference} was not found in the master workbook."
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
            r'(<(?:[A-Za-z_][\w.\-]*:)?f\b[^>]*>'
            r'.*?'
            r'</(?:[A-Za-z_][\w.\-]*:)?f>'
            r'|'
            r'<(?:[A-Za-z_][\w.\-]*:)?f\b[^>]*/>)'
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
        f'<{prefix}t xml:space="preserve">'
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
            or
            value == ""
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
    # TEXT
    # ========================================================

    else:

        escaped_value = (
            escape_excel_text(
                value
            )
        )


        replacement = (
            f'{start_tag} t="inlineStr">'
            f"{is_open}"
            f"{t_open}"
            f"{escaped_value}"
            f"{t_close}"
            f"{is_close}"
            f"{c_close}"
        )


    return (
        xml_text[
            : match.start()
        ]
        +
        replacement
        +
        xml_text[
            match.end() :
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
# VALIDATION
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
                    "Generated workbook contains damaged file: "
                    f"{broken_file}"
                )


            for filename in (
                workbook.namelist()
            ):

                if not (
                    filename.endswith(
                        ".xml"
                    )
                    or
                    filename.endswith(
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
                        "Generated workbook contains invalid XML "
                        f"in {filename}: {exc}"
                    )


    except zipfile.BadZipFile as exc:

        raise ValueError(
            "Generated file is not a valid XLSX workbook: "
            f"{exc}"
        )


# ============================================================
# FINAL GENERATION
# ============================================================

def generate_excel(
    payload
):

    if not TEMPLATE_PATH.exists():

        raise FileNotFoundError(
            "Master Excel template was not found: "
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
        or
        "export",
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
                    f"Worksheet '{sheet_name}' "
                    "was not found inside the master workbook."
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


            patched_files[
                worksheet_path
            ] = (
                patch_sheet(
                    original_xml,
                    changes,
                )
            )


        with zipfile.ZipFile(
            output_path,
            "w",
            compression=
                zipfile.ZIP_DEFLATED,
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
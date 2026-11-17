import re
from typing import Optional


# ============================================================
# BASIC HELPERS
# ============================================================

def clean(value: Optional[str]) -> Optional[str]:
    if value is None:
        return None

    value = re.sub(r"\s+", " ", value).strip()

    return value or None


def normalize_date(value: Optional[str]) -> Optional[str]:
    if not value:
        return None

    value = value.strip()

    match = re.search(
        r"(\d{1,2})[./-](\d{1,2})[./-](\d{4})",
        value,
    )

    if not match:
        return value

    day, month, year = match.groups()

    return f"{int(day):02d}/{int(month):02d}/{year}"


def get_section(
    text: str,
    start_marker: str,
    end_marker: str,
) -> str:
    pattern = (
        re.escape(start_marker)
        + r"(.*?)"
        + re.escape(end_marker)
    )

    match = re.search(
        pattern,
        text,
        flags=re.I | re.S,
    )

    if not match:
        return ""

    return match.group(1).strip()


def first_line(block: str) -> Optional[str]:
    for line in block.splitlines():
        line = line.strip()

        if line:
            return line

    return None


# ============================================================
# PARTY HELPERS
# ============================================================

def extract_email(block: str) -> Optional[str]:
    match = re.search(
        r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}",
        block,
        flags=re.I,
    )

    return match.group(0) if match else None


def extract_phone(block: str) -> Optional[str]:
    match = re.search(
        r"\+\d[\d\s-]{6,}\d",
        block,
    )

    if not match:
        return None

    return re.sub(
        r"\s+",
        " ",
        match.group(0),
    ).strip()


def extract_contact(block: str) -> Optional[str]:
    lines = block.splitlines()

    for index, line in enumerate(lines):
        if not re.search(
            r"Contact\s*person",
            line,
            flags=re.I,
        ):
            continue

        current = re.sub(
            r"^.*?Contact\s*person\s*:?\s*",
            "",
            line,
            flags=re.I,
        ).strip()

        parts = []

        if current:
            parts.append(current)

        next_index = index + 1

        while next_index < len(lines):
            next_line = lines[next_index].strip()

            if not next_line:
                break

            if "@" in next_line:
                break

            if re.match(
                r"^(tel\s*:|\+?\d[\d\s-]{5,})",
                next_line,
                flags=re.I,
            ):
                break

            if re.match(
                r"^(Pickup\s+address|Address)",
                next_line,
                flags=re.I,
            ):
                break

            parts.append(next_line)

            next_index += 1

        if parts:
            return clean(" ".join(parts))

    return None


def extract_address(block: str) -> Optional[str]:
    lines = [
        line.strip()
        for line in block.splitlines()
        if line.strip()
    ]

    address_lines = []
    collecting = False

    for line in lines:
        if re.match(
            r"^Address\s*:?",
            line,
            flags=re.I,
        ):
            collecting = True

            first_address_line = re.sub(
                r"^Address\s*:?\s*",
                "",
                line,
                flags=re.I,
            ).strip()

            if first_address_line:
                address_lines.append(
                    first_address_line
                )

            continue

        if collecting:
            if re.match(
                (
                    r"^(Contact\s*person|"
                    r"Pickup\s+address|"
                    r"tel\s*:)"
                ),
                line,
                flags=re.I,
            ):
                break

            if "@" in line:
                break

            if re.match(
                r"^\+\d",
                line,
            ):
                break

            address_lines.append(line)

    return clean(
        ", ".join(address_lines)
    )


def extract_country(
    address: Optional[str],
) -> Optional[str]:
    if not address:
        return None

    known_countries = [
        "United States",
        "Singapore",
        "India",
        "Germany",
        "United Kingdom",
        "France",
        "Italy",
        "Canada",
        "Mexico",
        "Australia",
        "China",
        "Japan",
    ]

    lowered = address.lower()

    for country in known_countries:
        if country.lower() in lowered:
            return country

    return None


def parse_party(
    block: str,
) -> dict:
    name = first_line(block)

    address = extract_address(block)

    return {
        "name": name,
        "contact": extract_contact(block),
        "address": address,
        "raw_section": block.strip() or None,
        "email": extract_email(block),
        "phone": extract_phone(block),
        "country": extract_country(address),
    }


# ============================================================
# PRODUCT HELPERS
# ============================================================

def deduplicate_description(
    description: str,
) -> Optional[str]:
    if not description:
        return None

    parts = [
        part.strip()
        for part in description.split(",")
        if part.strip()
    ]

    result = []
    seen = set()

    for part in parts:
        key = part.lower()

        if key in seen:
            continue

        seen.add(key)
        result.append(part)

    return ", ".join(result) or None


def normalize_material(
    material: Optional[str],
) -> Optional[str]:
    if not material:
        return None

    material = re.sub(
        r"\s+",
        " ",
        material,
    ).strip()

    # Repair PDF wrapping such as:
    #
    # 3.3
    # 211
    #
    # -> 3.3211
    material = re.sub(
        r"(\d\.\d)\s+(\d{3})\b",
        r"\1\2",
        material,
    )

    return material or None


def _build_product(
    *,
    part_number: Optional[str],
    description: Optional[str],
    material: Optional[str],
    net_weight: str,
    hs_code: str,
    country: str,
    quantity: str,
    unit_price: str,
    position_price: str,
) -> dict:
    return {
        "part_number": clean(part_number),

        "description":
            deduplicate_description(
                description or ""
            ),

        "material":
            normalize_material(material),

        "net_weight_kg":
            float(net_weight),

        "hs_code":
            clean(hs_code),

        "india_hsn":
            None,

        "india_hsn_code":
            None,

        "country_of_origin":
            clean(country),

        "quantity":
            int(quantity),

        "currency_symbol":
            "$",

        "unit_price":
            float(
                unit_price.replace(",", "")
            ),

        "position_price":
            float(
                position_price.replace(",", "")
            ),
    }


# ============================================================
# PRODUCT PARSER
#
# Supports BOTH:
#
# 1. Existing DocPort invoices containing:
#       Part # XXXXX, ...
#
# 2. Xometry position-based invoices:
#       10
#       Description...
#       Material: ...
#       1.26 9817850100 India 1 US$222.27 US$222.27
#
# Existing behaviour is preserved.
# ============================================================

def parse_products(
    text: str,
) -> list[dict]:
    products: list[dict] = []

    # --------------------------------------------------------
    # Locate product table.
    # --------------------------------------------------------

    table_match = re.search(
        r"Pos\.?\s+Goods\s+description(.*?)Totals:",
        text,
        flags=re.I | re.S,
    )

    if not table_match:
        return products

    product_text = (
        table_match
        .group(1)
        .replace("\r", "")
    )

    # --------------------------------------------------------
    # FORMAT 1:
    # Existing Part # invoice format.
    # --------------------------------------------------------

    existing_text = re.sub(
        r"\s+",
        " ",
        product_text,
    )

    existing_pattern = re.compile(
        r"Part\s*#\s*"
        r"(?P<part_number>[^,\s]+)"
        r"\s*,\s*"
        r"(?P<description>.*?)"
        r"\s+Material:\s*"
        r"(?P<material>.*?)"
        r"\s+"
        r"(?P<net_weight>\d+(?:\.\d+)?)"
        r"\s+"
        r"(?P<hs_code>\d{6,12})"
        r"\s+"
        r"(?P<country>[A-Za-z][A-Za-z ]*?)"
        r"\s+"
        r"(?P<quantity>\d+)"
        r"\s+"
        r"(?:US|USD)?\s*\$\s*"
        r"(?P<unit_price>\d[\d,]*(?:\.\d+)?)"
        r"\s+"
        r"(?:US|USD)?\s*\$\s*"
        r"(?P<position_price>\d[\d,]*(?:\.\d+)?)",
        flags=re.I,
    )

    for match in existing_pattern.finditer(
        existing_text
    ):
        products.append(
            _build_product(
                part_number=(
                    match.group(
                        "part_number"
                    )
                ),
                description=(
                    match.group(
                        "description"
                    ).strip()
                ),
                material=(
                    match.group(
                        "material"
                    ).strip()
                ),
                net_weight=(
                    match.group(
                        "net_weight"
                    )
                ),
                hs_code=(
                    match.group(
                        "hs_code"
                    )
                ),
                country=(
                    match.group(
                        "country"
                    )
                ),
                quantity=(
                    match.group(
                        "quantity"
                    )
                ),
                unit_price=(
                    match.group(
                        "unit_price"
                    )
                ),
                position_price=(
                    match.group(
                        "position_price"
                    )
                ),
            )
        )

    # Existing supported invoice worked.
    # Return immediately and preserve its behaviour.
    if products:
        return products

    # --------------------------------------------------------
    # FORMAT 2:
    # Xometry position-based invoice.
    # --------------------------------------------------------

    normalized = re.sub(
        r"[ \t]+",
        " ",
        product_text,
    )

    # Repair wrapped grade number if PDF split it.
    normalized = re.sub(
        r"(\d\.\d)\s*\n?\s*(\d{3})\b",
        r"\1\2",
        normalized,
    )

    # A Xometry row looks similar to:
    #
    # 10
    # This is a test fixture...
    # Material: Aluminium AW-6061-T6
    # x / 3.3211 / Al-Mg1SiCu
    # Prototype...
    # 1.26 9817850100 India 1 US$222.27 US$222.27
    #
    xometry_pattern = re.compile(
        r"(?:^|\n)\s*"
        r"(?P<position>\d+)\s+"

        r"(?P<body>.*?)"

        r"(?P<net_weight>\d+(?:\.\d+)?)"
        r"\s+"

        r"(?P<hs_code>\d{6,12})"
        r"\s+"

        r"(?P<country>[A-Za-z][A-Za-z ]*?)"
        r"\s+"

        r"(?P<quantity>\d+)"
        r"\s+"

        r"(?:US|USD)?\s*\$\s*"
        r"(?P<unit_price>\d[\d,]*(?:\.\d+)?)"
        r"\s+"

        r"(?:US|USD)?\s*\$\s*"
        r"(?P<position_price>\d[\d,]*(?:\.\d+)?)",

        flags=re.I | re.S,
    )

    for match in xometry_pattern.finditer(
        normalized
    ):
        body = (
            match.group("body")
            .strip()
        )

        description = body
        material = None

        # ----------------------------------------------------
        # Separate description from Material section.
        # ----------------------------------------------------

        material_match = re.search(
            r"Material:\s*(.+)",
            body,
            flags=re.I | re.S,
        )

        if material_match:
            description = (
                body[
                    :material_match.start()
                ]
                .strip()
            )

            material_block = (
                material_match
                .group(1)
                .strip()
            )

            # Remove explanatory notes that follow material.
            #
            # Example:
            # Prototype for testing & evaluation...
            #
            material_block = re.split(
                (
                    r"\bPrototype\s+for\b"
                    r"|"
                    r"\bPrimary\s+HTS\s*:"
                ),
                material_block,
                maxsplit=1,
                flags=re.I,
            )[0].strip()

            material = normalize_material(
                material_block
            )

        # Convert multiline description into clean text.
        description = re.sub(
            r"\s+",
            " ",
            description,
        ).strip(" ,")

        products.append(
            _build_product(
                # Position "10" is not a manufacturer
                # part number, therefore keep part_number
                # empty rather than inventing one.
                part_number=None,

                description=description,

                material=material,

                net_weight=(
                    match.group(
                        "net_weight"
                    )
                ),

                hs_code=(
                    match.group(
                        "hs_code"
                    )
                ),

                country=(
                    match.group(
                        "country"
                    )
                ),

                quantity=(
                    match.group(
                        "quantity"
                    )
                ),

                unit_price=(
                    match.group(
                        "unit_price"
                    )
                ),

                position_price=(
                    match.group(
                        "position_price"
                    )
                ),
            )
        )

    return products


# ============================================================
# SHIPMENT PARSER
# ============================================================

def parse_shipment(
    text: str,
) -> dict:
    dimensions = None
    package_weight = None
    total_quantity = None
    total_amount = None

    # Example:
    #
    # Totals: 1 positions
    # Packaging: 27x21x8 cm, 1.5 kg
    # Q-ty: 1
    # US$222.27
    #
    totals_match = re.search(
        r"Totals:\s*"
        r"\d+\s+positions?"
        r"\s+Packaging:\s*"
        r"([^,]+)"
        r",\s*"
        r"(\d+(?:\.\d+)?)\s*kg"
        r"\s+Q-ty:\s*"
        r"(\d+)"
        r"\s+"
        r"(?:US|USD)?\s*\$\s*"
        r"(\d[\d,]*(?:\.\d+)?)",
        text,
        flags=re.I | re.S,
    )

    if totals_match:
        dimensions = clean(
            totals_match.group(1)
        )

        package_weight = float(
            totals_match.group(2)
        )

        total_quantity = int(
            totals_match.group(3)
        )

        total_amount = float(
            totals_match
            .group(4)
            .replace(",", "")
        )

    return {
        "package_dimensions":
            dimensions,

        "package_weight_kg":
            package_weight,

        "net_weight_kg":
            None,

        "total_quantity":
            total_quantity,

        "total_amount":
            total_amount,

        # Keep existing DocPort behaviour.
        "currency":
            None,

        "exchange_rate":
            None,

        "awb_number":
            None,

        "shipping_bill_number":
            None,

        "shipping_bill_date":
            None,

        "rodtep_yes_no":
            None,

        "rodtep_total_line_items":
            None,

        "rodtep_claimed_line_items":
            None,

        "freight":
            None,

        "insurance":
            None,

        "commission":
            None,

        "discount":
            None,

        "packing_charges":
            None,

        "package_count":
            None,

        "state_origin":
            None,

        "district_origin":
            None,
    }


# ============================================================
# MAIN LOCAL INVOICE PARSER
# ============================================================

def parse_invoice_locally(
    text: str,
) -> dict:
    exporter_block = get_section(
        text,
        "Shipper / Exporter of record:",
        "Importer of record:",
    )

    importer_block = get_section(
        text,
        "Importer of record:",
        "Ship to / Consignee:",
    )

    consignee_block = get_section(
        text,
        "Ship to / Consignee:",
        "Sold to:",
    )

    sold_to_block = get_section(
        text,
        "Sold to:",
        "Invoice details:",
    )

    invoice_number = None
    invoice_date = None
    po_number = None

    # --------------------------------------------------------
    # Invoice number
    # --------------------------------------------------------

    match = re.search(
        (
            r"Invoice\s+Number\s*:\s*"
            r"(.+?)"
            r"(?="
            r"\s+Purchase\s+Order\s*:"
            r"|\r?\n"
            r"|$)"
        ),
        text,
        flags=re.I,
    )

    if match:
        invoice_number = clean(
            match.group(1)
        )

    # --------------------------------------------------------
    # Invoice date
    # --------------------------------------------------------

    match = re.search(
        (
            r"Invoice\s+Issue\s+Date\s*:\s*"
            r"(.+?)"
            r"(?="
            r"\s+Carrier\s*:"
            r"|\r?\n"
            r"|$)"
        ),
        text,
        flags=re.I,
    )

    if match:
        invoice_date = normalize_date(
            match.group(1)
        )

    # --------------------------------------------------------
    # Purchase order
    # --------------------------------------------------------

    match = re.search(
        (
            r"Purchase\s+Order\s*:\s*"
            r"(.+?)"
            r"(?="
            r"\s+Carrier\s*:"
            r"|\r?\n"
            r"|$)"
        ),
        text,
        flags=re.I,
    )

    if match:
        po_number = clean(
            match.group(1)
        )

    # --------------------------------------------------------
    # Products and shipment.
    # --------------------------------------------------------

    products = parse_products(text)

    shipment = parse_shipment(text)

    return {
        "invoice": {
            "number":
                invoice_number,

            "date":
                invoice_date,

            "order_date":
                None,

            "po_number":
                po_number,
        },

        "exporter":
            parse_party(
                exporter_block
            ),

        "importer":
            parse_party(
                importer_block
            ),

        "consignee":
            parse_party(
                consignee_block
            ),

        "sold_to":
            parse_party(
                sold_to_block
            ),

        "product": {
            "line_items":
                products
        },

        "shipment":
            shipment,

        "company": {
            "iec": None,
            "gstin": None,
        },

        "bank": {
            "code": None,
            "account_number": None,
            "ifsc": None,
            "swift_code": None,
        },
    }


# ============================================================
# LOCAL RESULT VALIDATION
#
# This determines whether api/index.py can immediately use the
# local result without sending the PDF to Gemini.
# ============================================================

def local_result_is_complete(
    data: dict,
) -> bool:
    invoice = data.get(
        "invoice",
        {},
    )

    product = data.get(
        "product",
        {},
    )

    shipment = data.get(
        "shipment",
        {},
    )

    line_items = product.get(
        "line_items",
        [],
    )

    if not invoice.get("number"):
        return False

    if not invoice.get("date"):
        return False

    if not invoice.get("po_number"):
        return False

    if not line_items:
        return False

    if not shipment.get(
        "package_dimensions"
    ):
        return False

    if shipment.get(
        "total_amount"
    ) is None:
        return False

    return True
SHEETS = {
    "invoice": "Invoice",
    "dhl": "Dhl Express SLI",
    "sdf": "FORM SDF",
    "evd": "EVD",
    "alum": "AlumSteel Derivatives declarati",
    "scomet": "Scomet Declaration",
}


CELLS = {
    # ========================================================
    # 1. EXPORT INVOICE
    # ========================================================
    "invoice": {
        "invoice_date": "N2",
        "invoice_number": "K3",
        "order_date": "N4",
        "po_number": "K5",

        "consignee_label": "A9",
        "consignee_address": "A10",
        "consignee_contact": "A15",

        "country_origin_1": "K7",
        "country_origin_2": "K9",

        "exchange_rate": "J17",

        # Master workbook has the added Final Destination
        # value row, therefore these cells are shifted by 1.
        "amount_inr": "D30",
        "amount_usd": "D31",

        "iec": "D32",
        "gstin": "D33",
        "bank_ad_code": "D34",
        "bank_account": "D35",
        "bank_ifsc": "D36",

        "package_count": "I31",
        "dimensions": "I32",
        "gross_weight": "I33",
        "net_weight": "I34",

        "total_before_tax": "O30",
        "total_after_tax": "O32",
    },


    # ========================================================
    # 2. DHL EXPRESS SLI
    # ========================================================
    "dhl": {
        "invoice_number": "I2",
        "consignee": "C3",

        # Actual workbook date cell.
        "invoice_date": "I3",

        "awb": "E5",
        "tax_block": "E13",

        "fob_value": "E22",

        "net_weight": "E29",
        "gross_weight": "E30",
    },


    # ========================================================
    # 3. FORM SDF
    # ========================================================
    "sdf": {
        "shipping_bill_header": "C2",

        # Actual visible SDF date.
        "date": "G2",

        # Sentence also contains the same dynamic date.
        "shipping_bill_declaration": "B4",
    },


    # ========================================================
    # 4. EVD
    # ========================================================
    "evd": {
        # Row:
        # 1. Shipping Bill No. ... & Date
        "shipping_bill": "A4",

        # Row:
        # 2. Invoice No. & Date ...
        "invoice_details": "A5",

        # IMPORTANT FIX:
        #
        # The visible right-side EVD date is E5.
        #
        # Previous mapping incorrectly used D5,
        # which is why Date: 19/09/2026 never changed.
        "invoice_date": "E5",

        # Previous shipping bill/date sentence.
        "previous_shipping_bill": "A14",

        # Bottom EVD date.
        "date": "D19",
    },


    # ========================================================
    # 5. ALUMINIUM / STEEL DERIVATIVES
    # ========================================================
    "alum": {
        "product": "A2",

        "weight_block": "A6",
        "value_block": "A7",

        # Uses current uploaded PDF date.
        "completed_date": "A12",

        "title": "A13",

        # Uses current uploaded PDF date.
        "date": "A15",
    },


    # ========================================================
    # 6. SCOMET DECLARATION
    # ========================================================
    "scomet": {
        "invoice_number": "A4",

        # Actual SCOMET date cell.
        "invoice_date": "D4",

        "product": "A5",
        "hsn_code": "A6",
    },
}


# ============================================================
# INVOICE PRODUCT ROWS
# ============================================================
#
# Current master structure:
#
# Row 24 = Port of Discharge / Final Destination
# Row 25 = US
# Row 26 = Product headings
# Row 27 = Product 1
# Row 28 = Product 2
# Row 29 = Product 3
#
# ============================================================

INVOICE_PRODUCT_ROWS = [
    {
        "sr_no": "A27",
        "india_hsn": "B27",
        "us_hts": "C27",
        "description": "D27",
        "quantity": "I27",
        "unit_price": "J27",
        "taxable_value": "K27",
        "igst_percent": "M27",
        "igst_amount": "N27",
    },

    {
        "sr_no": "A28",
        "india_hsn": "B28",
        "us_hts": "C28",
        "description": "D28",
        "quantity": "I28",
        "unit_price": "J28",
        "taxable_value": "K28",
        "igst_percent": "M28",
        "igst_amount": "N28",
    },

    {
        "sr_no": "A29",
        "india_hsn": "B29",
        "us_hts": "C29",
        "description": "D29",
        "quantity": "I29",
        "unit_price": "J29",
        "taxable_value": "K29",
        "igst_percent": "M29",
        "igst_amount": "N29",
    },
]


# ============================================================
# DHL ADDITIONAL CELLS
# ============================================================

DHL_EXTRA_CELLS = {
    "rodtep_yes_no": "G18",
    "rodtep_total_line_items": "J18",
    "rodtep_claimed_line_items": "L18",

    "freight": "E23",
    "insurance": "E24",
    "commission": "E25",
    "discount": "E26",
    "packing_charges": "E27",

    "package_count": "E28",

    "state_origin": "E31",
    "district_origin": "E32",
}
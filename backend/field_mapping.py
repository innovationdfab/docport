# field_mapping.py
#
# FINAL MAPPING FOR:
# ExportFlow_Website_Exact_Excel_Workbook.xlsx
#
# IMPORTANT:
# These locations match the NEW clean workbook design.
# Do not use the old N5 / L6 / K23 mappings anymore.

SHEETS = {
    "invoice": "Invoice",
    "dhl": "Dhl Express SLI",
    "sdf": "FORM SDF",
    "evd": "EVD",
    "alum": "AlumSteel Derivatives declarati",
    "scomet": "Scomet Declaration",
}


CELLS = {
    # =========================================================
    # INVOICE
    # =========================================================
    "invoice": {
        # Header
        "invoice_date": "N2",
        "invoice_number": "K3",

        "order_date": "N4",
        "po_number": "K5",

        # Consignee
        "consignee_label": "A9",
        "consignee_address": "A10",
        "consignee_contact": "A15",

        # Country
        "country_origin_1": "K7",
        "country_origin_2": "K9",

        # Exchange rate
        "exchange_rate": "J17",

        # First product row
        "product_description": "D26",
        "quantity": "I26",
        "unit_price": "J26",
        "taxable_value": "K26",

        # Totals
        "amount_inr": "D29",
        "amount_usd": "D30",

        "total_before_tax": "O29",
        "total_after_tax": "O31",

        # Packing
        "dimensions": "I31",
        "gross_weight": "I32",
        "net_weight": "I33",
    },

    # =========================================================
    # DHL EXPRESS SLI
    # =========================================================
    "dhl": {
        "invoice_number": "I2",
        "consignee": "C3",
        "invoice_date": "I3",

        "awb": "E5",

        # Composite TAXABLE / IGST block
        "tax_block": "E13",

        "fob_value": "E22",

        "net_weight": "E29",
        "gross_weight": "E30",
    },

    # =========================================================
    # FORM SDF
    # =========================================================
    "sdf": {
        # Blue shipping bill entry cell
        "shipping_bill_header": "C2",

        # Complete declaration sentence containing
        # shipping bill number + date
        "shipping_bill_declaration": "B4",
    },

    # =========================================================
    # EVD
    # =========================================================
    "evd": {
        "shipping_bill": "A4",
        "invoice_details": "A5",
        "previous_shipping_bill": "A14",
        "date": "D19",
    },

    # =========================================================
    # ALUMINIUM / STEEL DECLARATION
    #
    # The final clean workbook combines some values into
    # single visual blocks. excel_generator.py must therefore
    # write complete block text to these cells.
    # =========================================================
    "alum": {
        "product": "A2",

        # One combined red block:
        # Full weight + Aluminium content weight
        "weight_block": "A6",

        # One combined blue block:
        # Total value + Aluminium content value
        "value_block": "A7",

        "completed_date": "A12",
        "title": "A13",
        "date": "A15",
    },

    # =========================================================
    # SCOMET DECLARATION
    # =========================================================
    "scomet": {
        # Invoice number and date are separate cells now
        "invoice_number": "A4",
        "invoice_date": "D4",

        # Combined product description/material block
        "product": "A5",

        "hsn_code": "A6",
    },
}


# =============================================================
# PRODUCT ROWS
#
# The final Invoice sheet has three clean product rows.
# excel_generator.py should use these instead of the old
# row 32 / 37 / 42 structure.
# =============================================================

INVOICE_PRODUCT_ROWS = [
    {
        "sr_no": "A26",
        "india_hsn": "B26",
        "us_hts": "C26",
        "description": "D26",
        "quantity": "I26",
        "unit_price": "J26",
        "taxable_value": "K26",
        "igst_percent": "M26",
        "igst_amount": "N26",
    },
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
]


# =============================================================
# EXTRA DHL MANUAL / EDITABLE CELLS
# =============================================================

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
from datetime import datetime
from pathlib import Path

import xlsxwriter
from xlsxwriter.utility import xl_cell_to_rowcol


OUTPUT_FILE = Path(__file__).with_name("stock_management.xlsx")

MAX_ARTICLE_ROWS = 200
MAX_SUPPLIER_ROWS = 100
MAX_PROCUREMENT_ROWS = 2000
MAX_MOVEMENT_ROWS = 2000
MAX_CATALOG_ROWS = 500
ARCHIVE_AGE_DAYS = 90

categories = [
    "Essential Oils",
    "Carrier Oils",
    "Stones",
    "Herbs",
    "Boxes",
    "Bottles",
    "Bottle Labels",
    "Box Labels",
    "Finished Goods",
]

units = ["ml", "L", "g", "kg", "unit"]
movement_reasons = ["Production", "Waste", "Return", "Correction", "Other"]
yes_no = ["Yes", "No"]

articles = [
    ["ART-001", "Argan Oil", "Carrier Oils", "L", 2, "Yes", "Cold-pressed carrier oil"],
    ["ART-002", "Jojoba Oil", "Carrier Oils", "L", 2, "Yes", "Carrier oil"],
    ["ART-003", "Sweet Almond Oil", "Carrier Oils", "L", 3, "Yes", "Carrier oil"],
    ["ART-004", "Coconut Oil", "Carrier Oils", "kg", 5, "Yes", "Carrier oil"],
    ["ART-005", "Olive Oil", "Carrier Oils", "L", 3, "Yes", "Carrier oil"],
    ["ART-006", "Castor Oil", "Carrier Oils", "L", 2, "Yes", "Carrier oil"],
    ["ART-007", "Lavender EO", "Essential Oils", "ml", 200, "Yes", "Essential oil"],
    ["ART-008", "Tea Tree EO", "Essential Oils", "ml", 150, "Yes", "Essential oil"],
    ["ART-009", "Peppermint EO", "Essential Oils", "ml", 150, "Yes", "Essential oil"],
    ["ART-010", "Rosemary EO", "Essential Oils", "ml", 150, "Yes", "Essential oil"],
    ["ART-011", "Eucalyptus EO", "Essential Oils", "ml", 150, "Yes", "Essential oil"],
    ["ART-012", "Lemon EO", "Essential Oils", "ml", 150, "Yes", "Essential oil"],
    ["ART-013", "Dried Rosemary", "Herbs", "kg", 1, "Yes", "Dried herb"],
    ["ART-014", "Dried Lavender", "Herbs", "kg", 1, "Yes", "Dried herb"],
    ["ART-015", "Rose Quartz", "Stones", "unit", 25, "Yes", "Decorative stone"],
    ["ART-016", "Amethyst", "Stones", "unit", 25, "Yes", "Decorative stone"],
    ["ART-017", "30ml Amber Bottle", "Bottles", "unit", 50, "Yes", "Packaging"],
    ["ART-018", "50ml Clear Bottle", "Bottles", "unit", 50, "Yes", "Packaging"],
    ["ART-019", "Kraft Box Small", "Boxes", "unit", 50, "Yes", "Packaging"],
    ["ART-020", "Kraft Box Large", "Boxes", "unit", 50, "Yes", "Packaging"],
    ["ART-021", "Bottle Label 30ml", "Bottle Labels", "unit", 100, "Yes", "Packaging label"],
    ["ART-022", "Bottle Label 50ml", "Bottle Labels", "unit", 100, "Yes", "Packaging label"],
    ["ART-023", "Box Label Small", "Box Labels", "unit", 100, "Yes", "Packaging label"],
    ["ART-024", "Box Label Large", "Box Labels", "unit", 100, "Yes", "Packaging label"],
    ["ART-025", "Lavender Roller 10ml", "Finished Goods", "unit", 20, "Yes", "Finished product"],
    ["ART-026", "Sleep Blend Oil 5ml", "Finished Goods", "unit", 15, "Yes", "Finished product"],
    ["ART-027", "Argan Face Serum 30ml", "Finished Goods", "unit", 10, "Yes", "Finished product"],
]

suppliers = [
    ["SUP-001", "Nature's Best Oils", "Maya Green", "+41 22 100 0101", "orders@naturesbest.example", "30 days", "Carrier oils supplier"],
    ["SUP-002", "Aromatherapy Supplies Co", "Leo Martin", "+41 22 100 0102", "sales@aromasupplies.example", "Prepaid", "Essential oils supplier"],
    ["SUP-003", "Crystal World", "Sara Stone", "+41 22 100 0103", "hello@crystalworld.example", "15 days", "Stones supplier"],
    ["SUP-004", "Herb Farm Direct", "Nora Field", "+41 22 100 0104", "orders@herbfarm.example", "30 days", "Herbs supplier"],
    ["SUP-005", "PackRight Solutions", "Adam Box", "+41 22 100 0105", "sales@packright.example", "30 days", "Packaging supplier"],
    ["SUP-006", "Alive Alchemy", "-", "-", "-", "-", "In-house production"],
]

supplier_articles = [
    ["SUP-001", "ART-001", "ARG-1L", 28.50, "2024-01-10", "Primary argan oil supplier"],
    ["SUP-001", "ART-002", "JOJ-1L", 22.00, "2024-01-10", "Primary jojoba oil supplier"],
    ["SUP-001", "ART-003", "ALM-1L", 14.25, "2024-01-10", "Sweet almond oil"],
    ["SUP-002", "ART-007", "LAV-500", 0.12, "2024-01-12", "Lavender EO per ml"],
    ["SUP-002", "ART-008", "TTO-500", 0.15, "2024-01-12", "Tea Tree EO per ml"],
    ["SUP-002", "ART-009", "PEP-500", 0.10, "2024-01-12", "Peppermint EO per ml"],
    ["SUP-003", "ART-015", "RQ-UNIT", 1.50, "2024-01-20", "Rose quartz pieces"],
    ["SUP-003", "ART-016", "AME-UNIT", 1.75, "2024-01-20", "Amethyst pieces"],
    ["SUP-004", "ART-013", "ROSEMARY-KG", 18.00, "2024-01-18", "Dried rosemary"],
    ["SUP-004", "ART-014", "LAVENDER-KG", 24.00, "2024-01-18", "Dried lavender"],
    ["SUP-005", "ART-017", "BOT-30-AM", 0.85, "2024-01-22", "30ml amber bottle"],
    ["SUP-005", "ART-018", "BOT-50-CL", 0.95, "2024-01-22", "50ml clear bottle"],
    ["SUP-005", "ART-019", "BOX-KR-S", 0.55, "2024-01-22", "Small kraft box"],
    ["SUP-005", "ART-021", "LBL-BOT-30", 0.08, "2024-01-22", "30ml bottle label"],
]

procurement = [
    ["2024-01-15", "Nature's Best Oils", "Argan Oil", 5, 28.50, "INV-001", "LOT-ARG-2401", "Yes", "Initial stock"],
    ["2024-01-15", "Nature's Best Oils", "Jojoba Oil", 3, 22.00, "INV-001", "LOT-JOJ-2401", "Yes", "Initial stock"],
    ["2024-01-20", "Aromatherapy Supplies Co", "Lavender EO", 500, 0.12, "INV-002", "LOT-LAV-2401", "Yes", "Initial stock"],
    ["2024-01-20", "Aromatherapy Supplies Co", "Tea Tree EO", 300, 0.15, "INV-002", "LOT-TTO-2401", "Yes", "Initial stock"],
    ["2024-02-01", "Crystal World", "Rose Quartz", 100, 1.50, "INV-003", "LOT-RQ-2402", "Yes", "Initial stock"],
    ["2024-02-05", "PackRight Solutions", "30ml Amber Bottle", 200, 0.85, "INV-004", "LOT-BOT30-2402", "Yes", "Initial stock"],
    ["2024-02-05", "PackRight Solutions", "Bottle Label 30ml", 250, 0.08, "INV-004", "LOT-LBL30-2402", "Yes", "Initial stock"],
    ["2024-02-07", "Herb Farm Direct", "Dried Rosemary", 2, 18.00, "INV-005", "LOT-HRB-RO-2402", "Yes", "Initial stock"],
    ["2024-02-10", "Nature's Best Oils", "Sweet Almond Oil", 6, 14.25, "INV-006", "LOT-ALM-2402", "Yes", "Initial stock"],
    ["2024-02-12", "PackRight Solutions", "Kraft Box Small", 150, 0.55, "INV-007", "LOT-BOXS-2402", "Yes", "Initial stock"],
    ["2024-03-01", "Alive Alchemy", "Lavender Roller 10ml", 30, 3.50, "PROD-B-001", "LOT-FG-LAV-2403", "Yes", "Production batch"],
    ["2024-03-05", "Alive Alchemy", "Sleep Blend Oil 5ml", 20, 4.20, "PROD-B-002", "LOT-FG-SLP-2403", "Yes", "Production batch"],
]

stock_movements = [
    # date, article, batch, quantity, reason, notes
    ["2024-02-10", "Argan Oil", "LOT-ARG-2401", 1.5, "Production", "Batch production draw"],
    ["2024-02-10", "Lavender EO", "LOT-LAV-2401", 150, "Production", "Batch production draw"],
    ["2024-02-15", "30ml Amber Bottle", "LOT-BOT30-2402", 50, "Production", "Packaging draw"],
    ["2024-02-20", "Jojoba Oil", "LOT-JOJ-2401", 0.5, "Waste", "Spillage"],
    ["2024-02-25", "Rose Quartz", "LOT-RQ-2402", 10, "Production", "Batch production draw"],
]

supplier_dict = {row[0]: row[1] for row in suppliers}
article_dict = {row[0]: row[1] for row in articles}


def write_headers(ws, headers, header_format, row=0):
    for col, header in enumerate(headers):
        ws.write(row, col, header, header_format)


def apply_common_sheet_format(ws, last_col, freeze=True):
    ws.autofilter(0, 0, 0, last_col)
    if freeze:
        ws.freeze_panes(1, 0)


def main():
    workbook = xlsxwriter.Workbook(OUTPUT_FILE)
    workbook.set_calc_mode("auto")

    # ── Formats ───────────────────────────────────────────────────────────────
    title_fmt = workbook.add_format({"bold": True, "font_size": 16, "font_color": "#2D3748"})
    subtitle_fmt = workbook.add_format({"bold": True, "font_size": 12, "font_color": "#4A5568"})

    input_header_fmt = workbook.add_format({
        "bold": True, "bg_color": "#4E7B5C", "font_color": "white",
        "border": 1, "align": "center", "valign": "vcenter",
    })
    auto_header_fmt = workbook.add_format({
        "bold": True, "bg_color": "#546E7A", "font_color": "white",
        "border": 1, "align": "center", "valign": "vcenter",
    })
    dashboard_header_fmt = workbook.add_format({
        "bold": True, "bg_color": "#455A64", "font_color": "white",
        "border": 1, "align": "center", "valign": "vcenter",
    })
    filter_label_fmt = workbook.add_format({
        "bold": True, "font_color": "#4A5568", "align": "right", "valign": "vcenter",
    })

    date_fmt = workbook.add_format({"num_format": "yyyy-mm-dd"})
    money_fmt = workbook.add_format({"num_format": "#,##0.00"})
    qty_fmt = workbook.add_format({"num_format": "#,##0.00"})
    note_fmt = workbook.add_format({"italic": True, "font_color": "#718096"})

    metric_label_fmt = workbook.add_format({"bold": True, "bg_color": "#EDF2F7", "border": 1})
    metric_value_fmt = workbook.add_format({"bold": True, "bg_color": "#F7FAFC", "border": 1, "num_format": "#,##0.00"})
    metric_int_fmt = workbook.add_format({"bold": True, "bg_color": "#F7FAFC", "border": 1, "num_format": "#,##0"})

    low_fmt = workbook.add_format({"bg_color": "#FED7D7", "font_color": "#9B2335"})
    warning_fmt = workbook.add_format({"bg_color": "#FEFCBF", "font_color": "#744210"})
    ok_fmt = workbook.add_format({"bg_color": "#C6F6D5", "font_color": "#276749"})
    depleted_fmt = workbook.add_format({"bg_color": "#E2E8F0", "font_color": "#4A5568"})
    open_fmt = workbook.add_format({"bg_color": "#BEE3F8", "font_color": "#2C5282"})
    missing_fmt = workbook.add_format({"bg_color": "#FED7D7", "font_color": "#9B2335", "border": 1})
    archived_fmt = workbook.add_format({"bg_color": "#EDF2F7", "font_color": "#718096", "italic": True})

    # Unlocked variants for input cells (sheet protection — formula cells stay locked)
    input_date_fmt = workbook.add_format({"num_format": "yyyy-mm-dd", "locked": False})
    input_qty_fmt = workbook.add_format({"num_format": "#,##0.00", "locked": False})
    input_money_fmt = workbook.add_format({"num_format": "#,##0.00", "locked": False})
    input_text_fmt = workbook.add_format({"locked": False})

    # ── Sheets ────────────────────────────────────────────────────────────────
    ws_dashboard = workbook.add_worksheet("Dashboard")
    ws_procurement = workbook.add_worksheet("Procurement")
    ws_movements = workbook.add_worksheet("Stock_Movements")
    ws_register = workbook.add_worksheet("Stock_Register")
    # Backing / admin (hidden)
    ws_detail = workbook.add_worksheet("Stock_Detail")
    ws_summary = workbook.add_worksheet("Stock_Summary")
    ws_movements_archive = workbook.add_worksheet("Stock_Movements_Archive")
    ws_articles = workbook.add_worksheet("Articles")
    ws_suppliers = workbook.add_worksheet("Suppliers")
    ws_catalog = workbook.add_worksheet("Supplier_Catalog")
    ws_lists = workbook.add_worksheet("Lists")

    ws_detail.hide()
    ws_summary.hide()
    ws_movements_archive.hide()
    ws_articles.hide()
    ws_suppliers.hide()
    ws_catalog.hide()
    ws_lists.hide()

    # ── Lists helper ──────────────────────────────────────────────────────────
    for row, value in enumerate(categories):
        ws_lists.write(row, 0, value)
    for row, value in enumerate(units):
        ws_lists.write(row, 1, value)
    for row, value in enumerate(movement_reasons):
        ws_lists.write(row, 2, value)
    for row, value in enumerate(yes_no):
        ws_lists.write(row, 3, value)
    # Columns F (5) and G (6) are left blank for VBA scratch use (validation bypass)
    ws_lists.set_column("A:G", 32, input_text_fmt)
    ws_lists.protect()

    # ── Articles ──────────────────────────────────────────────────────────────
    article_headers = ["Article_ID", "Article_Name", "Category", "Unit", "Reorder_Level", "Active", "Notes"]
    write_headers(ws_articles, article_headers, input_header_fmt)
    for row_idx, row in enumerate(articles, start=1):
        ws_articles.write_row(row_idx, 0, row)

    ws_articles.data_validation(1, 2, MAX_ARTICLE_ROWS, 2, {"validate": "list", "source": "=Lists!$A$1:$A$9"})
    ws_articles.data_validation(1, 3, MAX_ARTICLE_ROWS, 3, {"validate": "list", "source": "=Lists!$B$1:$B$5"})
    ws_articles.data_validation(1, 4, MAX_ARTICLE_ROWS, 4, {
        "validate": "decimal", "criteria": ">=", "value": 0,
        "input_title": "Reorder Level", "input_message": "Enter reorder level >= 0",
        "error_title": "Invalid", "error_message": "Must be zero or positive.",
    })
    ws_articles.data_validation(1, 5, MAX_ARTICLE_ROWS, 5, {"validate": "list", "source": "=Lists!$D$1:$D$2"})
    ws_articles.set_column("A:A", 14, input_text_fmt)
    ws_articles.set_column("B:B", 24, input_text_fmt)
    ws_articles.set_column("C:C", 18, input_text_fmt)
    ws_articles.set_column("D:D", 10, input_text_fmt)
    ws_articles.set_column("E:E", 15, input_text_fmt)
    ws_articles.set_column("F:F", 10, input_text_fmt)
    ws_articles.set_column("G:G", 30, input_text_fmt)
    # Highlight duplicate Article_Name entries (causes silent VLOOKUP/SUMIFS corruption)
    ws_articles.conditional_format(1, 1, MAX_ARTICLE_ROWS, 1, {
        "type": "formula",
        "criteria": '=AND($B2<>"",COUNTIF($B$2:$B$201,$B2)>1)',
        "format": missing_fmt,
    })
    ws_articles.protect()
    # Note: no autofilter/freeze — sheet is hidden (Excel rejects those on hidden sheets)

    # ── Suppliers ─────────────────────────────────────────────────────────────
    supplier_headers = ["Supplier_ID", "Supplier_Name", "Contact_Person", "Phone", "Email", "Payment_Terms", "Notes"]
    write_headers(ws_suppliers, supplier_headers, input_header_fmt)
    for row_idx, row in enumerate(suppliers, start=1):
        ws_suppliers.write_row(row_idx, 0, row)
    ws_suppliers.set_column("A:A", 12, input_text_fmt)
    ws_suppliers.set_column("B:B", 28, input_text_fmt)
    ws_suppliers.set_column("C:C", 18, input_text_fmt)
    ws_suppliers.set_column("D:D", 18, input_text_fmt)
    ws_suppliers.set_column("E:E", 32, input_text_fmt)
    ws_suppliers.set_column("F:F", 16, input_text_fmt)
    ws_suppliers.set_column("G:G", 30, input_text_fmt)
    # Highlight duplicate Supplier_Name entries
    ws_suppliers.conditional_format(1, 1, MAX_SUPPLIER_ROWS, 1, {
        "type": "formula",
        "criteria": '=AND($B2<>"",COUNTIF($B$2:$B$101,$B2)>1)',
        "format": missing_fmt,
    })
    ws_suppliers.protect()
    # Note: no autofilter/freeze — sheet is hidden

    # ── Supplier_Catalog ──────────────────────────────────────────────────────
    sc_headers = ["Supplier", "Article", "Supplier_SKU", "Price_Per_Unit", "Last_Updated", "Notes"]
    write_headers(ws_catalog, sc_headers, input_header_fmt)
    for row_idx in range(1, MAX_CATALOG_ROWS + 1):
        excel_row = row_idx + 1
        # Price auto-fills from latest procurement for this article
        ws_catalog.write_formula(
            row_idx, 3,
            f'=IF(B{excel_row}="","",IFERROR(LOOKUP(2,1/(Procurement!$C$2:$C$2001=B{excel_row}),Procurement!$F$2:$F$2001),""))',
            money_fmt,
        )
        # Last_Updated auto-fills from latest procurement date for this article
        ws_catalog.write_formula(
            row_idx, 4,
            f'=IF(B{excel_row}="","",IFERROR(LOOKUP(2,1/(Procurement!$C$2:$C$2001=B{excel_row}),Procurement!$A$2:$A$2001),""))',
            date_fmt,
        )
    for row_idx, row in enumerate(supplier_articles, start=1):
        supplier_id, article_id, sku, _price, _last_updated, notes = row
        supplier_name = supplier_dict.get(supplier_id, supplier_id)
        article_name = article_dict.get(article_id, article_id)
        ws_catalog.write(row_idx, 0, supplier_name)
        ws_catalog.write(row_idx, 1, article_name)
        ws_catalog.write(row_idx, 2, sku)
        # D (Price) and E (Last_Updated) are formulas — auto-calculated above
        ws_catalog.write(row_idx, 5, notes)
    ws_catalog.data_validation(1, 0, MAX_CATALOG_ROWS, 0, {"validate": "list", "source": "=Suppliers!$B$2:$B$101"})
    ws_catalog.data_validation(1, 1, MAX_CATALOG_ROWS, 1, {"validate": "list", "source": "=Articles!$B$2:$B$201"})
    ws_catalog.set_column("A:A", 28, input_text_fmt)
    ws_catalog.set_column("B:B", 24, input_text_fmt)
    ws_catalog.set_column("C:C", 16, input_text_fmt)
    ws_catalog.set_column("D:D", 14, input_money_fmt)
    ws_catalog.set_column("E:E", 14, input_date_fmt)
    ws_catalog.set_column("F:F", 32, input_text_fmt)
    ws_catalog.protect()
    # Note: no autofilter/freeze — sheet is hidden

    # ── Procurement ───────────────────────────────────────────────────────────
    procurement_input_headers = [
        "Date", "Supplier", "Article", "Unit", "Quantity",
        "Unit_Cost", "Total_Cost", "Invoice_Number", "Batch_Lot_Number", "Active", "Notes",
    ]

    for col, header in enumerate(procurement_input_headers):
        ws_procurement.write(0, col, header, input_header_fmt)

    for row_idx in range(1, MAX_PROCUREMENT_ROWS + 1):
        excel_row = row_idx + 1
        ws_procurement.write_formula(
            row_idx, 3,
            f'=IF(C{excel_row}="","",IFERROR(VLOOKUP(C{excel_row},Articles!$B$2:$D${MAX_ARTICLE_ROWS + 1},3,FALSE),""))',
        )
        ws_procurement.write_formula(
            row_idx, 6,
            f'=IF(OR(E{excel_row}="",F{excel_row}=""),"",E{excel_row}*F{excel_row})',
        )

    for row_idx, row in enumerate(procurement, start=1):
        date, supplier, article, quantity, unit_cost, invoice, batch, active, notes = row
        ws_procurement.write_datetime(row_idx, 0, datetime.strptime(date, "%Y-%m-%d"), input_date_fmt)
        ws_procurement.write(row_idx, 1, supplier)
        ws_procurement.write(row_idx, 2, article)
        ws_procurement.write_number(row_idx, 4, quantity, input_qty_fmt)
        ws_procurement.write_number(row_idx, 5, unit_cost, input_money_fmt)
        ws_procurement.write(row_idx, 7, invoice)
        ws_procurement.write(row_idx, 8, batch)
        ws_procurement.write(row_idx, 9, active)
        ws_procurement.write(row_idx, 10, notes)

    ws_procurement.data_validation(1, 1, MAX_PROCUREMENT_ROWS, 1, {"validate": "list", "source": "=Suppliers!$B$2:$B$101"})
    ws_procurement.data_validation(1, 9, MAX_PROCUREMENT_ROWS, 9, {"validate": "list", "source": "=Lists!$D$1:$D$2"})
    # C: Article (Dependent list)
    ws_procurement.data_validation(1, 2, MAX_PROCUREMENT_ROWS, 2, {
        "validate": "list", "source": ["(Select Supplier First)"], # Placeholder, handled by VBA
        "ignore_blank": True,
    })
    ws_procurement.data_validation(1, 0, MAX_PROCUREMENT_ROWS, 0, {
        "validate": "custom", "value": '=AND(ISNUMBER(A2),A2<=TODAY())',
        "input_title": "Date", "input_message": "Enter a valid date (not in the future)",
        "error_title": "Invalid date", "error_message": "Enter a valid date that is not in the future.",
    })
    ws_procurement.data_validation(1, 4, MAX_PROCUREMENT_ROWS, 4, {
        "validate": "decimal", "criteria": ">", "value": 0,
        "input_title": "Quantity", "input_message": "Enter quantity > 0",
        "error_title": "Invalid quantity", "error_message": "Quantity must be a positive number.",
    })
    ws_procurement.data_validation(1, 5, MAX_PROCUREMENT_ROWS, 5, {
        "validate": "decimal", "criteria": ">=", "value": 0,
        "input_title": "Unit Cost", "input_message": "Enter unit cost >= 0",
        "error_title": "Invalid cost", "error_message": "Unit cost must be zero or positive.",
    })

    # Highlight Date in red when Article is filled but Date is missing
    ws_procurement.conditional_format(1, 0, MAX_PROCUREMENT_ROWS, 0, {
        "type": "formula",
        "criteria": '=AND($C2<>"",$A2="")',
        "format": missing_fmt,
    })
    # Highlight Batch_Lot_Number in red when Article is filled but Batch is missing
    ws_procurement.conditional_format(1, 8, MAX_PROCUREMENT_ROWS, 8, {
        "type": "formula",
        "criteria": '=AND($C2<>"",$I2="")',
        "format": missing_fmt,
    })

    ws_procurement.set_column("A:A", 12, input_date_fmt)
    ws_procurement.set_column("B:B", 28, input_text_fmt)
    ws_procurement.set_column("C:C", 24, input_text_fmt)
    ws_procurement.set_column("D:D", 10)
    ws_procurement.set_column("E:E", 14, input_qty_fmt)
    ws_procurement.set_column("F:F", 14, input_money_fmt)
    ws_procurement.set_column("G:G", 14, money_fmt)
    ws_procurement.set_column("H:I", 20, input_text_fmt)
    ws_procurement.set_column("J:J", 10, input_text_fmt)
    ws_procurement.set_column("K:K", 32, input_text_fmt)
    apply_common_sheet_format(ws_procurement, len(procurement_input_headers) - 1)
    ws_procurement.protect()

    # ── Stock_Movements ───────────────────────────────────────────────────────
    movement_headers = [
        "Date", "Article", "Batch_Lot_Number", "Unit",
        "Available_Qty", "Qty_Drawn", "Reason", "Notes", "Status", "Archived_On",
    ]
    write_headers(ws_movements, movement_headers, input_header_fmt)

    for row_idx in range(1, MAX_MOVEMENT_ROWS + 1):
        excel_row = row_idx + 1
        ws_movements.write_formula(
            row_idx, 3,
            f'=IF(B{excel_row}="","",IFERROR(VLOOKUP(B{excel_row},Articles!$B$2:$D${MAX_ARTICLE_ROWS + 1},3,FALSE),""))',
        )
        ws_movements.write_formula(
            row_idx, 4,
            f'=IF(OR(B{excel_row}="",C{excel_row}=""),"",IFERROR(SUMIFS(Stock_Detail!$I$2:$I$2001,Stock_Detail!$A$2:$A$2001,B{excel_row},Stock_Detail!$E$2:$E$2001,C{excel_row}),""))',
        )
        ws_movements.write_formula(row_idx, 8, f'=IF(B{excel_row}="","","Active")')
        ws_movements.write_formula(row_idx, 9, f'=IF(B{excel_row}="","","")')

    for row_idx, row in enumerate(stock_movements, start=1):
        date, article, batch, quantity, reason, notes = row
        ws_movements.write_datetime(row_idx, 0, datetime.strptime(date, "%Y-%m-%d"), input_date_fmt)
        ws_movements.write(row_idx, 1, article)
        ws_movements.write(row_idx, 2, batch)
        ws_movements.write_number(row_idx, 5, quantity, input_qty_fmt)
        ws_movements.write(row_idx, 6, reason)
        ws_movements.write(row_idx, 7, notes)

    ws_movements.data_validation(1, 1, MAX_MOVEMENT_ROWS, 1, {"validate": "list", "source": "=Articles!$B$2:$B$201"})
    # Batch validation: warn if batch doesn't exist in Procurement (VBA overrides with filtered list)
    ws_movements.data_validation(1, 2, MAX_MOVEMENT_ROWS, 2, {
        "validate": "custom", "value": '=OR(C2="",COUNTIF(Procurement!$I$2:$I$2001,C2)>0)',
        "input_title": "Batch", "input_message": "Select or type a batch from Procurement",
        "error_title": "Unknown batch", "error_message": "This batch does not exist in Procurement. Check spelling or add it first.",
    })
    ws_movements.data_validation(1, 6, MAX_MOVEMENT_ROWS, 6, {"validate": "list", "source": "=Lists!$C$1:$C$5"})
    ws_movements.data_validation(1, 0, MAX_MOVEMENT_ROWS, 0, {
        "validate": "custom", "value": '=AND(ISNUMBER(A2),A2<=TODAY())',
        "input_title": "Date", "input_message": "Enter a valid date (not in the future)",
        "error_title": "Invalid date", "error_message": "Enter a valid date that is not in the future.",
    })
    # Qty_Drawn: must be positive, but over-draw is ALLOWED (real-world variations)
    ws_movements.data_validation(1, 5, MAX_MOVEMENT_ROWS, 5, {
        "validate": "decimal", "criteria": ">", "value": 0,
        "input_title": "Qty Drawn", "input_message": "Enter quantity drawn (> 0)",
        "error_title": "Invalid", "error_message": "Quantity must be a positive number.",
    })

    # Highlight Date in red when Article is filled but Date is missing
    ws_movements.conditional_format(1, 0, MAX_MOVEMENT_ROWS, 0, {
        "type": "formula",
        "criteria": '=AND($B2<>"",$A2="")',
        "format": missing_fmt,
    })

    # Highlight Reason in amber when Article is set but Reason is empty
    ws_movements.conditional_format(1, 6, MAX_MOVEMENT_ROWS, 6, {
        "type": "formula",
        "criteria": '=AND($B2<>"",$G2="")',
        "format": warning_fmt,
    })

    ws_movements.set_column("A:A", 12, input_date_fmt)
    ws_movements.set_column("B:B", 24, input_text_fmt)
    ws_movements.set_column("C:C", 22, input_text_fmt)
    ws_movements.set_column("D:D", 10)
    ws_movements.set_column("E:E", 14)
    ws_movements.set_column("F:F", 14, input_qty_fmt)
    ws_movements.set_column("G:G", 16, input_text_fmt)
    ws_movements.set_column("H:H", 32, input_text_fmt)
    ws_movements.set_column("I:I", 12)
    ws_movements.set_column("J:J", 14)
    apply_common_sheet_format(ws_movements, len(movement_headers) - 1)
    ws_movements.protect()

    # ── Stock_Movements_Archive ───────────────────────────────────────────────
    write_headers(ws_movements_archive, movement_headers, auto_header_fmt)
    ws_movements_archive.set_column("A:A", 12, date_fmt)
    ws_movements_archive.set_column("B:B", 24)
    ws_movements_archive.set_column("C:C", 22)
    ws_movements_archive.set_column("D:J", 14, input_text_fmt)
    ws_movements_archive.protect()
    # Note: no autofilter/freeze — sheet is hidden

    # ── Stock_Detail ──────────────────────────────────────────────────────────
    detail_headers = [
        "Article", "Category", "Supplier", "Invoice_Number", "Batch_Lot_Number",
        "Date", "Qty_Purchased", "Qty_Drawn", "Qty_Remaining", "Status",
        "Unit_Cost", "Stock_Value", "Last_Movement_Date", "Days_Since_Last_Movement",
    ]
    write_headers(ws_detail, detail_headers, auto_header_fmt)
    for row_idx in range(1, MAX_PROCUREMENT_ROWS + 1):
        excel_row = row_idx + 1
        proc_row = excel_row
        vlookup_cat = f'VLOOKUP(A{excel_row},Articles!$B$2:$C${MAX_ARTICLE_ROWS + 1},2,FALSE)'
        qty_drawn_formula = (
            f'IF(A{excel_row}="","",'
            f'IFERROR(SUMIFS(Stock_Movements!$F$2:$F$2001,Stock_Movements!$B$2:$B$2001,A{excel_row},Stock_Movements!$C$2:$C$2001,E{excel_row},Stock_Movements!$I$2:$I$2001,"Active"),0)'
            f'+IFERROR(SUMIFS(Stock_Movements_Archive!$F$2:$F$2001,Stock_Movements_Archive!$B$2:$B$2001,A{excel_row},Stock_Movements_Archive!$C$2:$C$2001,E{excel_row}),0))'
        )
        last_movement_formula = (
            f'IF(A{excel_row}="","",'
            f'IFERROR(MAXIFS(Stock_Movements!$A$2:$A$2001,Stock_Movements!$B$2:$B$2001,A{excel_row},Stock_Movements!$C$2:$C$2001,E{excel_row},Stock_Movements!$I$2:$I$2001,"Active"),'
            f'IFERROR(MAXIFS(Stock_Movements_Archive!$A$2:$A$2001,Stock_Movements_Archive!$B$2:$B$2001,A{excel_row},Stock_Movements_Archive!$C$2:$C$2001,E{excel_row}),\"\")))'
        )
        formulas = [
            f'=IF(Procurement!C{proc_row}="","",Procurement!C{proc_row})',
            f'=IF(A{excel_row}="","",IFERROR({vlookup_cat},""))',
            f'=IF(Procurement!B{proc_row}="","",Procurement!B{proc_row})',
            f'=IF(Procurement!H{proc_row}="","",Procurement!H{proc_row})',
            f'=IF(Procurement!I{proc_row}="","",Procurement!I{proc_row})',
            f'=IF(Procurement!A{proc_row}="","",Procurement!A{proc_row})',
            f'=IF(Procurement!E{proc_row}="","",Procurement!E{proc_row})',
            f'={qty_drawn_formula}',
            f'=IF(G{excel_row}="","",G{excel_row}-H{excel_row})',
            f'=IF(A{excel_row}="","",IF(ROUND(I{excel_row},2)>0,"Open","Depleted"))',
            f'=IF(Procurement!F{proc_row}="","",Procurement!F{proc_row})',
            f'=IF(I{excel_row}="","",I{excel_row}*K{excel_row})',
            f'={last_movement_formula}',
            f'=IF(M{excel_row}="","",IF(ISNUMBER(M{excel_row}),TODAY()-M{excel_row},""))',
        ]
        for col, formula in enumerate(formulas):
            ws_detail.write_formula(row_idx, col, formula)

    ws_detail.set_column("A:A", 24)
    ws_detail.set_column("B:B", 18)
    ws_detail.set_column("C:C", 28)
    ws_detail.set_column("D:E", 20)
    ws_detail.set_column("F:F", 12, date_fmt)
    ws_detail.set_column("G:I", 14, qty_fmt)
    ws_detail.set_column("J:J", 12)
    ws_detail.set_column("K:L", 14, money_fmt)
    ws_detail.set_column("M:M", 16, date_fmt)
    ws_detail.set_column("N:N", 18)
    ws_detail.protect()
    # Note: no conditional formatting, autofilter, or freeze — sheet is hidden

    # ── Stock_Summary ─────────────────────────────────────────────────────────
    summary_headers = ["Article", "Category", "Unit", "Total_Stock", "Reorder_Level", "Active", "Status", "Stock_Value"]
    write_headers(ws_summary, summary_headers, auto_header_fmt)
    for row_idx in range(1, MAX_ARTICLE_ROWS + 1):
        excel_row = row_idx + 1
        formulas = [
            f'=IF(Articles!B{excel_row}="","",Articles!B{excel_row})',
            f'=IF(A{excel_row}="","",Articles!C{excel_row})',
            f'=IF(A{excel_row}="","",Articles!D{excel_row})',
            f'=IF(A{excel_row}="","",SUMIF(Stock_Detail!$A$2:$A$2001,A{excel_row},Stock_Detail!$I$2:$I$2001))',
            f'=IF(A{excel_row}="","",Articles!E{excel_row})',
            f'=IF(A{excel_row}="","",Articles!F{excel_row})',
            f'=IF(A{excel_row}="","",IF(D{excel_row}<E{excel_row},"Low",IF(D{excel_row}<E{excel_row}*1.5,"Warning","OK")))',
            f'=IF(A{excel_row}="","",SUMIF(Stock_Detail!$A$2:$A$2001,A{excel_row},Stock_Detail!$L$2:$L$2001))',
        ]
        for col, formula in enumerate(formulas):
            ws_summary.write_formula(row_idx, col, formula)
    ws_summary.set_column("A:A", 24)
    ws_summary.set_column("B:B", 18)
    ws_summary.set_column("C:C", 10)
    ws_summary.set_column("D:E", 14, qty_fmt)
    ws_summary.set_column("F:G", 12)
    ws_summary.set_column("H:H", 14, money_fmt)
    ws_summary.protect()
    # Note: no conditional formatting, autofilter, or freeze — sheet is hidden

    # ── Stock_Register ────────────────────────────────────────────────────────
    ws_register.merge_range(0, 0, 0, 2, "Stock Register", title_fmt)
    ws_register.write(0, 3, "Show:", filter_label_fmt)
    ws_register.write(0, 4, "Open Batches", input_text_fmt)
    ws_register.data_validation(0, 4, 0, 4, {
        "validate": "list",
        "source": ["Open Batches", "Depleted Batches", "All Batches"],
        "input_title": "Filter view",
        "input_message": "Select which batches to display",
    })
    ws_register.set_row(0, 26)

    write_headers(ws_register, detail_headers, auto_header_fmt, row=1)

    register_formula = (
        f'=IFERROR('
        f'IF(E1="Depleted Batches",'
        f'SORT(FILTER(Stock_Detail!A2:N{MAX_PROCUREMENT_ROWS + 1},Stock_Detail!$J$2:$J$2001="Depleted"),1,TRUE),'
        f'IF(E1="All Batches",'
        f'SORT(FILTER(Stock_Detail!A2:N{MAX_PROCUREMENT_ROWS + 1},Stock_Detail!$A$2:$A$2001<>""),1,TRUE),'
        f'SORT(FILTER(Stock_Detail!A2:N{MAX_PROCUREMENT_ROWS + 1},Stock_Detail!$J$2:$J$2001="Open"),1,1)'
        f')),'
        f'"No batches to show")'
    )
    ws_register.write_dynamic_array_formula(2, 0, 2, 0, register_formula)
    ws_register.freeze_panes(2, 0)
    ws_register.protect()

    ws_register.set_column("A:A", 24)
    ws_register.set_column("B:B", 18)
    ws_register.set_column("C:C", 28)
    ws_register.set_column("D:E", 20)
    ws_register.set_column("F:F", 12, date_fmt)
    ws_register.set_column("G:I", 14, qty_fmt)
    ws_register.set_column("J:J", 12)
    ws_register.set_column("K:L", 14, money_fmt)
    ws_register.set_column("M:M", 16, date_fmt)
    ws_register.set_column("N:N", 18)

    # ── Dashboard ─────────────────────────────────────────────────────────────
    ws_dashboard.merge_range("A1:H1", "Alive Alchemy — Stock Dashboard", title_fmt)
    ws_dashboard.write("A3", "Metric", dashboard_header_fmt)
    ws_dashboard.write("B3", "Value", dashboard_header_fmt)
    metrics = [
        ("A4", "B4", "Total Articles", f'=COUNTIF(Articles!$B$2:$B$201,"<>")', metric_int_fmt),
        ("A5", "B5", "Active Articles", f'=COUNTIF(Articles!$F$2:$F${MAX_ARTICLE_ROWS + 1},"Yes")', metric_int_fmt),
        ("A6", "B6", "Low Stock Items", f'=COUNTIF(Stock_Summary!$G$2:$G${MAX_ARTICLE_ROWS + 1},"Low")', metric_int_fmt),
        ("A7", "B7", "Warning Stock Items", f'=COUNTIF(Stock_Summary!$G$2:$G${MAX_ARTICLE_ROWS + 1},"Warning")', metric_int_fmt),
        ("A8", "B8", "Total Stock Value (CHF)", f'=SUM(Stock_Summary!$H$2:$H${MAX_ARTICLE_ROWS + 1})', metric_value_fmt),
    ]
    for label_cell, value_cell, label, formula, fmt in metrics:
        ws_dashboard.write(label_cell, label, metric_label_fmt)
        ws_dashboard.write_formula(value_cell, formula, fmt)

    # Stock Tools buttons — reliable on both Mac and Windows Excel
    ws_dashboard.write("J3", "Stock Tools", subtitle_fmt)
    ws_dashboard.set_column("J:J", 22)
    ws_dashboard.set_column("K:K", 26)
    dashboard_buttons = [
        ("J4", "GoToFirstEmptyProcurementRow", "Go to empty Procurement row"),
        ("J5", "GoToFirstEmptyMovementRow", "Go to empty Movement row"),
        ("J6", "ArchiveOldMovements", "Archive old movements"),
        ("J7", "RestoreArchivedMovements", "Restore archived movements"),
    ]
    for cell, macro, caption in dashboard_buttons:
        row, col = xl_cell_to_rowcol(cell)
        ws_dashboard.insert_button(
            row, col,
            {
                "macro": macro,
                "caption": caption,
                "width": 220,
                "height": 22,
            },
        )

    ws_dashboard.write("A11", "Low / Warning Stock Alerts", subtitle_fmt)
    low_headers = ["Article", "Category", "Unit", "Total_Stock", "Reorder_Level", "Active", "Status", "Stock_Value"]
    for col, header in enumerate(low_headers):
        ws_dashboard.write(11, col, header, dashboard_header_fmt)
    ws_dashboard.write_dynamic_array_formula(
        12, 0, 12, 0,
        f'=IFERROR(FILTER(Stock_Summary!A2:H{MAX_ARTICLE_ROWS + 1},'
        f'(Stock_Summary!G2:G{MAX_ARTICLE_ROWS + 1}="Low")+'
        f'(Stock_Summary!G2:G{MAX_ARTICLE_ROWS + 1}="Warning")),"No low or warning stock")',
    )

    ws_dashboard.write("I11", "Recent Procurement", subtitle_fmt)
    recent_headers = ["Date", "Supplier", "Article", "Unit", "Quantity", "Unit_Cost", "Total_Cost", "Invoice_Number", "Batch_Lot_Number", "Active", "Notes"]
    for col, header in enumerate(recent_headers, start=8):
        ws_dashboard.write(11, col, header, dashboard_header_fmt)
    ws_dashboard.write_dynamic_array_formula(
        12, 8, 12, 8,
        f'=IFERROR(TAKE(SORT(FILTER(Procurement!A2:K{MAX_PROCUREMENT_ROWS + 1},'
        f'Procurement!A2:A{MAX_PROCUREMENT_ROWS + 1}<>""),1,FALSE),10),"No procurement yet")',
    )

    ws_dashboard.write("A32", "How It Works", subtitle_fmt)
    lines = [
        "DAILY WORKFLOW — 4 visible tabs: Dashboard · Procurement · Stock_Movements · Stock_Register",
        "Admin/backing sheets are hidden. To reveal them: right-click any tab > Unhide (Excel).",
        "",
        "VBA SETUP (one-time, Excel 365):",
        "  1. File > Save As > stock_management.xlsm (Macro-Enabled Workbook — required to run VBA)",
        "  2. Open VBA editor: Alt+F11 (Windows) or Option+F11 (Mac)",
        "  3. Import StockDropdowns_Module.bas: File > Import File > select it (creates 'StockDropdowns' module)",
        "  4. Double-click 'ThisWorkbook', delete any code, paste ThisWorkbook.bas",
        "  5. Delete any old code in the 'Procurement' and 'Stock_Movements' sheet modules from previous setups.",
        "  6. Save (Ctrl+S) and close the VBA editor. Reopen the workbook with macros enabled.",
        "",
        "STOCK TOOLS: use the buttons on the Dashboard (top-right). They work on Mac and Windows.",
        "",
        "DATE ENTRY TIP: double-click the Date cell to stamp today's date instantly. Or press Ctrl+;.",
        "",
        "DEPENDENT DROPDOWNS (requires VBA): Article in Procurement filters by Supplier.",
        "Batch_Lot_Number in Stock_Movements filters to open batches for the selected article.",
        "",
        "PROCUREMENT (IN): log every purchase. Select Supplier → Article dropdown auto-filters.",
        "Type your own Batch_Lot_Number (column I), e.g. LOT-ARG-2401.",
        "Missing Batch_Lot_Number cells are highlighted in red when Article is already filled.",
        "",
        "STOCK MOVEMENTS (OUT): select Article → Batch dropdown auto-filters to open batches.",
        "Qty_Drawn can exceed Available_Qty (real-world variations allowed — adjust later if needed).",
        "Missing Reason cells are highlighted in amber when Article is filled.",
        "",
        "STOCK REGISTER: one tab replaces three. Use the 'Show:' dropdown (top-right, cell E1):",
        "  • Open Batches (default) — daily working view",
        "  • Depleted Batches — audit trail",
        "  • All Batches — full register",
        "Note: FILTER/SORT/TAKE require Excel 365 (not compatible with older Excel versions).",
        "ARCHIVING: Stock Tools > Archive old movements (90+ days). Rows are cleared from Stock_Movements and stored in the hidden archive sheet.",
        "ARCHIVING: Stock Tools > Archive old movements (90+ days). Rows are cleared from Stock_Movements and stored in the hidden archive sheet.",
        "To bring them back: Stock Tools > Restore archived movements.",
        "",
        "MONITOR: Dashboard shows low-stock alerts and recent procurement at a glance.",
    ]
    for i, line in enumerate(lines):
        ws_dashboard.write(33 + i, 0, line, note_fmt)

    ws_dashboard.set_column("A:A", 24)
    ws_dashboard.set_column("B:H", 16)
    ws_dashboard.set_column("I:I", 12)
    ws_dashboard.set_column("J:J", 28)
    ws_dashboard.set_column("K:K", 24)
    ws_dashboard.set_column("L:S", 14)
    ws_dashboard.freeze_panes(11, 0)
    ws_dashboard.protect()

    # ── Tab colours ───────────────────────────────────────────────────────────
    ws_dashboard.set_tab_color("#455A64")
    ws_procurement.set_tab_color("#D97706")
    ws_movements.set_tab_color("#D97706")
    ws_register.set_tab_color("#546E7A")
    ws_detail.set_tab_color("#B0BEC5")
    ws_summary.set_tab_color("#B0BEC5")
    ws_movements_archive.set_tab_color("#B0BEC5")
    ws_articles.set_tab_color("#4E7B5C")
    ws_suppliers.set_tab_color("#4E7B5C")
    ws_catalog.set_tab_color("#4E7B5C")

    workbook.close()
    print(f"Created {OUTPUT_FILE}")


if __name__ == "__main__":
    main()

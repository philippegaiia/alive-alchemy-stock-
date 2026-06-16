from datetime import datetime
from pathlib import Path

import xlsxwriter


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
    ["2024-03-01", "", "Lavender Roller 10ml", 30, 3.50, "PROD-B-001", "LOT-FG-LAV-2403", "Yes", "Production batch"],
    ["2024-03-05", "", "Sleep Blend Oil 5ml", 20, 4.20, "PROD-B-002", "LOT-FG-SLP-2403", "Yes", "Production batch"],
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
    # Typography
    title_fmt = workbook.add_format({"bold": True, "font_size": 16, "font_color": "#2D3748"})
    subtitle_fmt = workbook.add_format({"bold": True, "font_size": 12, "font_color": "#4A5568"})

    # Headers: sage-green for editable input sheets, slate-blue for computed views
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

    # Inline label for the Stock_Register filter control
    filter_label_fmt = workbook.add_format({
        "bold": True, "font_color": "#4A5568", "align": "right", "valign": "vcenter",
    })

    # Data formats
    date_fmt = workbook.add_format({"num_format": "yyyy-mm-dd"})
    money_fmt = workbook.add_format({"num_format": "#,##0.00"})
    qty_fmt = workbook.add_format({"num_format": "#,##0.00"})
    note_fmt = workbook.add_format({"italic": True, "font_color": "#718096"})

    # Dashboard metrics
    metric_label_fmt = workbook.add_format({"bold": True, "bg_color": "#EDF2F7", "border": 1})
    metric_value_fmt = workbook.add_format({"bold": True, "bg_color": "#F7FAFC", "border": 1, "num_format": "#,##0.00"})
    metric_int_fmt = workbook.add_format({"bold": True, "bg_color": "#F7FAFC", "border": 1, "num_format": "#,##0"})

    # Status / alert colours — muted palette
    low_fmt = workbook.add_format({"bg_color": "#FED7D7", "font_color": "#9B2335"})
    warning_fmt = workbook.add_format({"bg_color": "#FEFCBF", "font_color": "#744210"})
    ok_fmt = workbook.add_format({"bg_color": "#C6F6D5", "font_color": "#276749"})
    depleted_fmt = workbook.add_format({"bg_color": "#E2E8F0", "font_color": "#4A5568"})
    open_fmt = workbook.add_format({"bg_color": "#BEE3F8", "font_color": "#2C5282"})
    # Missing-data highlight (red border, light red bg)
    missing_fmt = workbook.add_format({"bg_color": "#FED7D7", "font_color": "#9B2335", "border": 1})
    archived_fmt = workbook.add_format({"bg_color": "#EDF2F7", "font_color": "#718096", "italic": True})

    # ── Sheets — tab order determines display order ────────────────────────────
    # Visible working tabs
    ws_dashboard = workbook.add_worksheet("Dashboard")
    ws_procurement = workbook.add_worksheet("Procurement")
    ws_movements = workbook.add_worksheet("Stock_Movements")
    ws_register = workbook.add_worksheet("Stock_Register")
    ws_movement_history = workbook.add_worksheet("Movement_History")
    # Backing / admin sheets (hidden — access via Format > Sheet > Show)
    ws_detail = workbook.add_worksheet("Stock_Detail")
    ws_summary = workbook.add_worksheet("Stock_Summary")
    ws_movements_archive = workbook.add_worksheet("Stock_Movements_Archive")
    ws_articles = workbook.add_worksheet("Articles")
    ws_suppliers = workbook.add_worksheet("Suppliers")
    ws_catalog = workbook.add_worksheet("Supplier_Catalog")
    ws_lists = workbook.add_worksheet("Lists")

    # Hide backing / admin sheets
    ws_detail.hide()
    ws_summary.hide()
    ws_movements_archive.hide()
    ws_articles.hide()
    ws_suppliers.hide()
    ws_catalog.hide()
    ws_lists.hide()

    # ── Lists helper sheet ────────────────────────────────────────────────────
    for row, value in enumerate(categories):
        ws_lists.write(row, 0, value)
    for row, value in enumerate(units):
        ws_lists.write(row, 1, value)
    for row, value in enumerate(movement_reasons):
        ws_lists.write(row, 2, value)
    for row, value in enumerate(yes_no):
        ws_lists.write(row, 3, value)
    ws_lists.set_column("A:D", 32)

    # ── Named ranges ─────────────────────────────────────────────────────────
    workbook.define_name("Categories", "=Lists!$A$1:$A$9")
    workbook.define_name("Units", "=Lists!$B$1:$B$5")
    workbook.define_name("MovementReasons", "=Lists!$C$1:$C$5")
    workbook.define_name("YesNo", "=Lists!$D$1:$D$2")
    workbook.define_name("ArticleIDs", f"=Articles!$A$2:$A${MAX_ARTICLE_ROWS + 1}")
    workbook.define_name("ArticleNames", f"=Articles!$B$2:$B${MAX_ARTICLE_ROWS + 1}")
    workbook.define_name("ArticleCategories", f"=Articles!$C$2:$C${MAX_ARTICLE_ROWS + 1}")
    workbook.define_name("ArticleUnits", f"=Articles!$D$2:$D${MAX_ARTICLE_ROWS + 1}")
    workbook.define_name("ReorderLevels", f"=Articles!$E$2:$E${MAX_ARTICLE_ROWS + 1}")
    workbook.define_name("SupplierIDs", f"=Suppliers!$A$2:$A${MAX_SUPPLIER_ROWS + 1}")
    workbook.define_name("SupplierNames", f"=Suppliers!$B$2:$B${MAX_SUPPLIER_ROWS + 1}")
    # Procurement columns
    workbook.define_name("ProcDates", f"=Procurement!$A$2:$A${MAX_PROCUREMENT_ROWS + 1}")
    workbook.define_name("ProcSuppliers", f"=Procurement!$B$2:$B${MAX_PROCUREMENT_ROWS + 1}")
    workbook.define_name("ProcArticles", f"=Procurement!$C$2:$C${MAX_PROCUREMENT_ROWS + 1}")
    workbook.define_name("ProcQuantities", f"=Procurement!$E$2:$E${MAX_PROCUREMENT_ROWS + 1}")
    workbook.define_name("ProcUnitCosts", f"=Procurement!$F$2:$F${MAX_PROCUREMENT_ROWS + 1}")
    workbook.define_name("ProcInvoices", f"=Procurement!$H$2:$H${MAX_PROCUREMENT_ROWS + 1}")
    workbook.define_name("ProcBatches", f"=Procurement!$I$2:$I${MAX_PROCUREMENT_ROWS + 1}")
    workbook.define_name("ProcActive", f"=Procurement!$J$2:$J${MAX_PROCUREMENT_ROWS + 1}")
    # Stock_Movements columns
    workbook.define_name("MoveDates", f"=Stock_Movements!$A$2:$A${MAX_MOVEMENT_ROWS + 1}")
    workbook.define_name("MoveArticles", f"=Stock_Movements!$B$2:$B${MAX_MOVEMENT_ROWS + 1}")
    workbook.define_name("MoveBatches", f"=Stock_Movements!$C$2:$C${MAX_MOVEMENT_ROWS + 1}")
    workbook.define_name("MoveQtyDrawn", f"=Stock_Movements!$F$2:$F${MAX_MOVEMENT_ROWS + 1}")
    workbook.define_name("MoveReasons", f"=Stock_Movements!$G$2:$G${MAX_MOVEMENT_ROWS + 1}")
    workbook.define_name("MoveStatus", f"=Stock_Movements!$I$2:$I${MAX_MOVEMENT_ROWS + 1}")
    # Stock_Movements_Archive columns
    workbook.define_name("ArchiveDates", f"=Stock_Movements_Archive!$A$2:$A${MAX_MOVEMENT_ROWS + 1}")
    workbook.define_name("ArchiveArticles", f"=Stock_Movements_Archive!$B$2:$B${MAX_MOVEMENT_ROWS + 1}")
    workbook.define_name("ArchiveBatches", f"=Stock_Movements_Archive!$C$2:$C${MAX_MOVEMENT_ROWS + 1}")
    workbook.define_name("ArchiveQtyDrawn", f"=Stock_Movements_Archive!$F$2:$F${MAX_MOVEMENT_ROWS + 1}")
    # Stock_Detail columns
    workbook.define_name("DetailArticles", f"=Stock_Detail!$A$2:$A${MAX_PROCUREMENT_ROWS + 1}")
    workbook.define_name("DetailBatches", f"=Stock_Detail!$E$2:$E${MAX_PROCUREMENT_ROWS + 1}")
    workbook.define_name("DetailQtyRemaining", f"=Stock_Detail!$I$2:$I${MAX_PROCUREMENT_ROWS + 1}")
    workbook.define_name("DetailStatus", f"=Stock_Detail!$J$2:$J${MAX_PROCUREMENT_ROWS + 1}")
    workbook.define_name("DetailStockValue", f"=Stock_Detail!$L$2:$L${MAX_PROCUREMENT_ROWS + 1}")

    # ── Articles (hidden admin sheet) ─────────────────────────────────────────
    articles_headers = ["Article_ID", "Article_Name", "Category", "Unit", "Reorder_Level", "Active", "Notes"]
    write_headers(ws_articles, articles_headers, input_header_fmt)
    for row_idx, row in enumerate(articles, start=1):
        ws_articles.write_row(row_idx, 0, row)
    ws_articles.data_validation(1, 2, MAX_ARTICLE_ROWS, 2, {"validate": "list", "source": "=Categories"})
    ws_articles.data_validation(1, 3, MAX_ARTICLE_ROWS, 3, {"validate": "list", "source": "=Units"})
    ws_articles.data_validation(1, 4, MAX_ARTICLE_ROWS, 4, {
        "validate": "decimal", "criteria": ">=", "value": 0,
        "input_title": "Reorder Level", "input_message": "Enter reorder level >= 0",
        "error_title": "Invalid reorder level", "error_message": "Reorder level must be zero or positive.",
    })
    ws_articles.data_validation(1, 5, MAX_ARTICLE_ROWS, 5, {"validate": "list", "source": "=YesNo"})
    ws_articles.set_column("A:A", 14)
    ws_articles.set_column("B:B", 24)
    ws_articles.set_column("C:C", 18)
    ws_articles.set_column("D:D", 10)
    ws_articles.set_column("E:E", 15)
    ws_articles.set_column("F:F", 10)
    ws_articles.set_column("G:G", 30)
    # Note: no autofilter/freeze — sheet is hidden (Excel rejects those on hidden sheets)

    # ── Suppliers (hidden admin sheet) ────────────────────────────────────────
    supplier_headers = ["Supplier_ID", "Supplier_Name", "Contact_Person", "Phone", "Email", "Payment_Terms", "Notes"]
    write_headers(ws_suppliers, supplier_headers, input_header_fmt)
    for row_idx, row in enumerate(suppliers, start=1):
        ws_suppliers.write_row(row_idx, 0, row)
    ws_suppliers.set_column("A:A", 12)
    ws_suppliers.set_column("B:B", 28)
    ws_suppliers.set_column("C:C", 18)
    ws_suppliers.set_column("D:D", 18)
    ws_suppliers.set_column("E:E", 32)
    ws_suppliers.set_column("F:F", 16)
    ws_suppliers.set_column("G:G", 30)
    # Note: no autofilter/freeze — sheet is hidden

    # ── Supplier_Catalog (hidden admin sheet) ─────────────────────────────────
    sc_headers = ["Supplier", "Supplier_Name", "Article", "Article_Name", "Supplier_SKU", "Price_Per_Unit", "Last_Updated", "Notes"]
    write_headers(ws_catalog, sc_headers, input_header_fmt)
    for row_idx in range(1, MAX_CATALOG_ROWS + 1):
        excel_row = row_idx + 1
        ws_catalog.write_formula(row_idx, 1, f'=IF(A{excel_row}="","",A{excel_row})')
        ws_catalog.write_formula(row_idx, 3, f'=IF(C{excel_row}="","",C{excel_row})')
    for row_idx, row in enumerate(supplier_articles, start=1):
        supplier_id, article_id, sku, price, last_updated, notes = row
        supplier_name = supplier_dict.get(supplier_id, supplier_id)
        article_name = article_dict.get(article_id, article_id)
        ws_catalog.write(row_idx, 0, supplier_name)
        ws_catalog.write(row_idx, 2, article_name)
        ws_catalog.write(row_idx, 4, sku)
        ws_catalog.write_number(row_idx, 5, price, money_fmt)
        ws_catalog.write(row_idx, 6, last_updated)
        ws_catalog.write(row_idx, 7, notes)
    ws_catalog.data_validation(1, 0, MAX_CATALOG_ROWS, 0, {"validate": "list", "source": "=SupplierNames"})
    ws_catalog.data_validation(1, 2, MAX_CATALOG_ROWS, 2, {"validate": "list", "source": "=ArticleNames"})
    ws_catalog.set_column("A:B", 28)
    ws_catalog.set_column("C:D", 24)
    ws_catalog.set_column("E:E", 16)
    ws_catalog.set_column("F:G", 14)
    ws_catalog.set_column("H:H", 32)
    # Note: no autofilter/freeze — sheet is hidden

    # ── Procurement ───────────────────────────────────────────────────────────
    # Columns A–K are user-input (green header); L is auto-computed (slate header).
    # L = Suggested_Batch_ID: a generated code the user can copy into column I.
    procurement_input_headers = [
        "Date", "Supplier", "Article", "Unit", "Quantity",
        "Unit_Cost", "Total_Cost", "Invoice_Number", "Batch_Lot_Number", "Active", "Notes",
    ]
    procurement_all_headers = procurement_input_headers + ["Suggested_Batch_ID"]

    for col, header in enumerate(procurement_input_headers):
        ws_procurement.write(0, col, header, input_header_fmt)
    ws_procurement.write(0, 11, "Suggested_Batch_ID", auto_header_fmt)

    for row_idx in range(1, MAX_PROCUREMENT_ROWS + 1):
        excel_row = row_idx + 1
        # Unit auto-lookup
        ws_procurement.write_formula(
            row_idx, 3,
            f'=IF(C{excel_row}="","",IFERROR(VLOOKUP(C{excel_row},Articles!$B$2:$D${MAX_ARTICLE_ROWS + 1},3,FALSE),""))',
        )
        # Total_Cost
        ws_procurement.write_formula(
            row_idx, 6,
            f'=IF(OR(E{excel_row}="",F{excel_row}=""),"",E{excel_row}*F{excel_row})',
        )
        # Suggested_Batch_ID — e.g. "ARGO-240115" from "Argan Oil" on 2024-01-15
        ws_procurement.write_formula(
            row_idx, 11,
            f'=IF(OR(A{excel_row}="",C{excel_row}=""),"",UPPER(SUBSTITUTE(LEFT(C{excel_row},4)," ",""))&"-"&TEXT(A{excel_row},"YYMMDD"))',
        )

    for row_idx, row in enumerate(procurement, start=1):
        date, supplier, article, quantity, unit_cost, invoice, batch, active, notes = row
        ws_procurement.write_datetime(row_idx, 0, datetime.strptime(date, "%Y-%m-%d"), date_fmt)
        ws_procurement.write(row_idx, 1, supplier)
        ws_procurement.write(row_idx, 2, article)
        ws_procurement.write_number(row_idx, 4, quantity, qty_fmt)
        ws_procurement.write_number(row_idx, 5, unit_cost, money_fmt)
        ws_procurement.write(row_idx, 7, invoice)
        ws_procurement.write(row_idx, 8, batch)
        ws_procurement.write(row_idx, 9, active)
        ws_procurement.write(row_idx, 10, notes)

    ws_procurement.data_validation(1, 1, MAX_PROCUREMENT_ROWS, 1, {"validate": "list", "source": "=SupplierNames"})
    ws_procurement.data_validation(1, 9, MAX_PROCUREMENT_ROWS, 9, {"validate": "list", "source": "=YesNo"})
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

    # Highlight Batch_Lot_Number cell in red when Article is filled but Batch is missing
    ws_procurement.conditional_format(1, 8, MAX_PROCUREMENT_ROWS, 8, {
        "type": "formula",
        "criteria": '=AND($C2<>"",$I2="")',
        "format": missing_fmt,
    })

    ws_procurement.set_column("A:A", 12, date_fmt)
    ws_procurement.set_column("B:B", 28)
    ws_procurement.set_column("C:C", 24)
    ws_procurement.set_column("D:D", 10)
    ws_procurement.set_column("E:E", 14, qty_fmt)
    ws_procurement.set_column("F:G", 14, money_fmt)
    ws_procurement.set_column("H:I", 20)
    ws_procurement.set_column("J:J", 10)
    ws_procurement.set_column("K:K", 32)
    ws_procurement.set_column("L:L", 22)
    apply_common_sheet_format(ws_procurement, len(procurement_all_headers) - 1)

    # ── Stock_Movements ───────────────────────────────────────────────────────
    movement_headers = [
        "Date", "Article", "Batch_Lot_Number", "Unit",
        "Available_Qty", "Qty_Drawn", "Reason", "Notes", "Status", "Archived_On",
    ]
    write_headers(ws_movements, movement_headers, input_header_fmt)

    for row_idx in range(1, MAX_MOVEMENT_ROWS + 1):
        excel_row = row_idx + 1
        # Unit auto-lookup
        ws_movements.write_formula(
            row_idx, 3,
            f'=IF(B{excel_row}="","",IFERROR(VLOOKUP(B{excel_row},ArticleNames:ArticleUnits,3,FALSE),""))',
        )
        # Available_Qty: remaining stock for this article+batch from Stock_Detail
        ws_movements.write_formula(
            row_idx, 4,
            f'=IF(OR(B{excel_row}="",C{excel_row}=""),"",IFNA(SUMIFS(DetailQtyRemaining,DetailArticles,B{excel_row},DetailBatches,C{excel_row}),""))',
        )
        # Status defaults to Active when row has data; archive macro overwrites with "Archived"
        ws_movements.write_formula(row_idx, 8, f'=IF(A{excel_row}="","","Active")')
        ws_movements.write_formula(row_idx, 9, f'=IF(A{excel_row}="","","")')

    for row_idx, row in enumerate(stock_movements, start=1):
        date, article, batch, quantity, reason, notes = row
        ws_movements.write_datetime(row_idx, 0, datetime.strptime(date, "%Y-%m-%d"), date_fmt)
        ws_movements.write(row_idx, 1, article)
        ws_movements.write(row_idx, 2, batch)
        ws_movements.write_number(row_idx, 5, quantity, qty_fmt)
        ws_movements.write(row_idx, 6, reason)
        ws_movements.write(row_idx, 7, notes)
        ws_movements.write(row_idx, 8, "Active")

    ws_movements.data_validation(1, 1, MAX_MOVEMENT_ROWS, 1, {"validate": "list", "source": "=ArticleNames"})
    ws_movements.data_validation(1, 6, MAX_MOVEMENT_ROWS, 6, {"validate": "list", "source": "=MovementReasons"})
    ws_movements.data_validation(1, 8, MAX_MOVEMENT_ROWS, 8, {"validate": "list", "source": ["Active", "Archived"]})
    ws_movements.data_validation(1, 0, MAX_MOVEMENT_ROWS, 0, {
        "validate": "custom", "value": '=AND(ISNUMBER(A2),A2<=TODAY())',
        "input_title": "Date", "input_message": "Enter a valid date (not in the future)",
        "error_title": "Invalid date", "error_message": "Enter a valid date that is not in the future.",
    })
    # Over-draft prevention: Qty_Drawn must be > 0 AND <= Available_Qty (when set)
    ws_movements.data_validation(1, 5, MAX_MOVEMENT_ROWS, 5, {
        "validate": "custom",
        "value": '=AND(F2>0,OR(E2="",F2<=E2))',
        "input_title": "Qty Drawn",
        "input_message": "Enter a positive number that does not exceed the Available Qty shown",
        "error_title": "Over-draft blocked",
        "error_message": "Qty drawn must be > 0 and cannot exceed the available quantity for this batch.",
    })

    # Highlight Reason cell in amber when Article is set but Reason is missing
    ws_movements.conditional_format(1, 6, MAX_MOVEMENT_ROWS, 6, {
        "type": "formula",
        "criteria": '=AND($B2<>"",$G2="")',
        "format": warning_fmt,
    })
    # Grey-out archived rows
    ws_movements.conditional_format(1, 0, MAX_MOVEMENT_ROWS, 9, {
        "type": "formula", "criteria": '=$I2="Archived"', "format": archived_fmt,
    })

    ws_movements.set_column("A:A", 12, date_fmt)
    ws_movements.set_column("B:B", 24)
    ws_movements.set_column("C:C", 22)
    ws_movements.set_column("D:D", 10)
    ws_movements.set_column("E:E", 14)
    ws_movements.set_column("F:F", 14, qty_fmt)
    ws_movements.set_column("G:G", 16)
    ws_movements.set_column("H:H", 32)
    ws_movements.set_column("I:I", 12)
    ws_movements.set_column("J:J", 14)
    apply_common_sheet_format(ws_movements, len(movement_headers) - 1)

    # ── Stock_Movements_Archive (hidden backing) ───────────────────────────────
    write_headers(ws_movements_archive, movement_headers, auto_header_fmt)
    ws_movements_archive.set_column("A:A", 12, date_fmt)
    ws_movements_archive.set_column("B:B", 24)
    ws_movements_archive.set_column("C:C", 22)
    ws_movements_archive.set_column("D:D", 10)
    ws_movements_archive.set_column("E:E", 14)
    ws_movements_archive.set_column("F:F", 14, qty_fmt)
    ws_movements_archive.set_column("G:G", 16)
    ws_movements_archive.set_column("H:H", 32)
    ws_movements_archive.set_column("I:I", 12)
    ws_movements_archive.set_column("J:J", 14)
    # Note: no autofilter/freeze — sheet is hidden

    # ── Movement_History (live view — visible, audit log) ─────────────────────
    write_headers(ws_movement_history, movement_headers, auto_header_fmt)
    ws_movement_history.set_column("A:A", 12, date_fmt)
    ws_movement_history.set_column("B:B", 24)
    ws_movement_history.set_column("C:C", 22)
    ws_movement_history.set_column("D:D", 10)
    ws_movement_history.set_column("E:E", 14)
    ws_movement_history.set_column("F:F", 14, qty_fmt)
    ws_movement_history.set_column("G:G", 16)
    ws_movement_history.set_column("H:H", 32)
    ws_movement_history.set_column("I:I", 12)
    ws_movement_history.set_column("J:J", 14)
    ws_movement_history.write_formula(
        1, 0,
        f'=IFERROR(SORT(FILTER(Stock_Movements!A2:J{MAX_MOVEMENT_ROWS + 1},MoveDates<>""),1,FALSE),"No movements yet")',
    )
    ws_movement_history.freeze_panes(1, 0)

    # ── Stock_Detail (hidden backing — one row per procurement batch) ─────────
    detail_headers = [
        "Article", "Category", "Supplier", "Invoice_Number", "Batch_Lot_Number",
        "Date", "Qty_Purchased", "Qty_Drawn", "Qty_Remaining", "Status",
        "Unit_Cost", "Stock_Value", "Last_Movement_Date", "Days_Since_Last_Movement",
        "Archive_Status",
    ]
    write_headers(ws_detail, detail_headers, auto_header_fmt)
    for row_idx in range(1, MAX_PROCUREMENT_ROWS + 1):
        excel_row = row_idx + 1
        proc_row = excel_row
        vlookup_cat = f'VLOOKUP(A{excel_row},ArticleNames:ArticleCategories,2,FALSE)'
        qty_drawn_formula = (
            f'IF(A{excel_row}="","",'
            f'IFERROR(SUMIFS(MoveQtyDrawn,MoveArticles,A{excel_row},MoveBatches,E{excel_row},MoveStatus,"Active"),0)'
            f'+IFERROR(SUMIFS(ArchiveQtyDrawn,ArchiveArticles,A{excel_row},ArchiveBatches,E{excel_row}),0))'
        )
        last_movement_formula = (
            f'IF(A{excel_row}="","",'
            f'IFERROR(MAXIFS(MoveDates,MoveArticles,A{excel_row},MoveBatches,E{excel_row},MoveStatus,"Active"),'
            f'IFERROR(MAXIFS(ArchiveDates,ArchiveArticles,A{excel_row},ArchiveBatches,E{excel_row}),\"\")))'
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
            f'=IF(A{excel_row}="","",IF(I{excel_row}>0,"Open","Depleted"))',
            f'=IF(Procurement!F{proc_row}="","",Procurement!F{proc_row})',
            f'=IF(I{excel_row}="","",I{excel_row}*K{excel_row})',
            f'={last_movement_formula}',
            f'=IF(M{excel_row}="","",IF(ISNUMBER(M{excel_row}),TODAY()-M{excel_row},""))',
            f'=IF(A{excel_row}="","",IF(AND(J{excel_row}="Depleted",ISNUMBER(N{excel_row}),N{excel_row}>{ARCHIVE_AGE_DAYS}),"Archive","Active"))',
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
    ws_detail.set_column("O:O", 14)
    # Note: no conditional formatting, autofilter, or freeze — sheet is hidden

    # ── Stock_Summary (hidden backing — one row per article) ──────────────────
    summary_headers = ["Article", "Category", "Unit", "Total_Stock", "Reorder_Level", "Active", "Status", "Stock_Value"]
    write_headers(ws_summary, summary_headers, auto_header_fmt)
    for row_idx in range(1, MAX_ARTICLE_ROWS + 1):
        excel_row = row_idx + 1
        formulas = [
            f'=IF(Articles!B{excel_row}="","",Articles!B{excel_row})',
            f'=IF(A{excel_row}="","",Articles!C{excel_row})',
            f'=IF(A{excel_row}="","",Articles!D{excel_row})',
            f'=IF(A{excel_row}="","",SUMIF(DetailArticles,A{excel_row},DetailQtyRemaining))',
            f'=IF(A{excel_row}="","",Articles!E{excel_row})',
            f'=IF(A{excel_row}="","",Articles!F{excel_row})',
            f'=IF(A{excel_row}="","",IF(D{excel_row}<E{excel_row},"Low",IF(D{excel_row}<E{excel_row}*1.5,"Warning","OK")))',
            f'=IF(A{excel_row}="","",SUMIF(DetailArticles,A{excel_row},DetailStockValue))',
        ]
        for col, formula in enumerate(formulas):
            ws_summary.write_formula(row_idx, col, formula)
    ws_summary.set_column("A:A", 24)
    ws_summary.set_column("B:B", 18)
    ws_summary.set_column("C:C", 10)
    ws_summary.set_column("D:E", 14, qty_fmt)
    ws_summary.set_column("F:F", 10)
    ws_summary.set_column("G:G", 12)
    ws_summary.set_column("H:H", 14, money_fmt)
    # Note: no conditional formatting, autofilter, or freeze — sheet is hidden

    # ── Stock_Register (consolidated view — replaces 3 separate detail tabs) ──
    # Row 0: title + filter dropdown
    # Row 1: column headers
    # Row 2+: dynamic FILTER/SORT formula driven by the dropdown in E1
    ws_register.merge_range(0, 0, 0, 2, "Stock Register", title_fmt)
    ws_register.write(0, 3, "Show:", filter_label_fmt)
    ws_register.write(0, 4, "Open Batches")   # default selection (pre-filled)
    ws_register.data_validation(0, 4, 0, 4, {
        "validate": "list",
        "source": ["Open Batches", "Depleted Batches", "All Batches"],
        "input_title": "Filter view",
        "input_message": "Select which batches to display",
    })
    ws_register.set_row(0, 26)  # slightly taller title/filter row

    write_headers(ws_register, detail_headers, auto_header_fmt, row=1)

    register_formula = (
        f'=IFERROR('
        f'IF(E1="Depleted Batches",'
        f'SORT(FILTER(Stock_Detail!A2:O{MAX_PROCUREMENT_ROWS + 1},DetailStatus="Depleted"),6,FALSE),'
        f'IF(E1="All Batches",'
        f'SORT(FILTER(Stock_Detail!A2:O{MAX_PROCUREMENT_ROWS + 1},DetailArticles<>""),1,TRUE),'
        f'SORT(FILTER(Stock_Detail!A2:O{MAX_PROCUREMENT_ROWS + 1},DetailStatus="Open"),1,1)'
        f')),'
        f'"No batches to show")'
    )
    ws_register.write_formula(2, 0, register_formula)
    ws_register.freeze_panes(2, 0)  # freeze title row + header row

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
    ws_register.set_column("O:O", 14)

    # ── Dashboard ─────────────────────────────────────────────────────────────
    ws_dashboard.merge_range("A1:H1", "Alive Alchemy — Stock Dashboard", title_fmt)
    ws_dashboard.write("A3", "Metric", dashboard_header_fmt)
    ws_dashboard.write("B3", "Value", dashboard_header_fmt)
    metrics = [
        ("A4", "B4", "Total Articles", f'=COUNTIF(ArticleNames,"<>")', metric_int_fmt),
        ("A5", "B5", "Active Articles", f'=COUNTIF(Articles!$F$2:$F${MAX_ARTICLE_ROWS + 1},"Yes")', metric_int_fmt),
        ("A6", "B6", "Low Stock Items", f'=COUNTIF(Stock_Summary!$G$2:$G${MAX_ARTICLE_ROWS + 1},"Low")', metric_int_fmt),
        ("A7", "B7", "Warning Stock Items", f'=COUNTIF(Stock_Summary!$G$2:$G${MAX_ARTICLE_ROWS + 1},"Warning")', metric_int_fmt),
        ("A8", "B8", "Total Stock Value (CHF)", f'=SUM(Stock_Summary!$H$2:$H${MAX_ARTICLE_ROWS + 1})', metric_value_fmt),
    ]
    for label_cell, value_cell, label, formula, fmt in metrics:
        ws_dashboard.write(label_cell, label, metric_label_fmt)
        ws_dashboard.write_formula(value_cell, formula, fmt)

    ws_dashboard.write("A11", "Low / Warning Stock Alerts", subtitle_fmt)
    low_headers = ["Article", "Category", "Unit", "Total_Stock", "Reorder_Level", "Active", "Status", "Stock_Value"]
    for col, header in enumerate(low_headers):
        ws_dashboard.write(11, col, header, dashboard_header_fmt)
    ws_dashboard.write_formula(
        "A13",
        f'=IFERROR(FILTER(Stock_Summary!A2:H{MAX_ARTICLE_ROWS + 1},'
        f'(Stock_Summary!G2:G{MAX_ARTICLE_ROWS + 1}="Low")+'
        f'(Stock_Summary!G2:G{MAX_ARTICLE_ROWS + 1}="Warning")),"No low or warning stock")',
    )

    ws_dashboard.write("I11", "Recent Procurement", subtitle_fmt)
    recent_headers = ["Date", "Supplier", "Article", "Unit", "Quantity", "Unit_Cost", "Total_Cost", "Invoice_Number", "Batch_Lot_Number", "Active", "Notes"]
    for col, header in enumerate(recent_headers, start=8):
        ws_dashboard.write(11, col, header, dashboard_header_fmt)
    ws_dashboard.write_formula(
        "I13",
        f'=IFERROR(TAKE(SORT(FILTER(Procurement!A2:K{MAX_PROCUREMENT_ROWS + 1},'
        f'Procurement!A2:A{MAX_PROCUREMENT_ROWS + 1}<>""),1,FALSE),10),"No procurement yet")',
    )

    ws_dashboard.write("A32", "How It Works", subtitle_fmt)
    lines = [
        "DAILY WORKFLOW — 5 visible tabs: Dashboard · Procurement · Stock_Movements · Stock_Register · Movement_History",
        "Admin/backing sheets are hidden. To reveal them: Format > Sheet > Show (Google Sheets).",
        "",
        "APPS SCRIPT (one-time): import into Google Sheets, then paste dependent_dropdown.gs via Extensions > Apps Script.",
        "Adds a 'Stock Tools' menu: Go to first empty row (Procurement / Stock_Movements) · Archive old movements.",
        "Dependent dropdowns: Article in Procurement filters by selected Supplier; Batch in Stock_Movements filters to open batches.",
        "",
        "DATE ENTRY TIP: press Cmd+; (Mac) or Ctrl+; (Windows/ChromeOS) to stamp today's date instantly — no script needed.",
        "",
        "PROCUREMENT (IN): log every purchase. Select Supplier → Article dropdown auto-filters to that supplier's catalogue.",
        "Copy the auto-generated Suggested_Batch_ID (column L) into the Batch_Lot_Number column (I) to avoid typos.",
        "Missing Batch_Lot_Number cells are highlighted in red when Article is already filled.",
        "",
        "STOCK MOVEMENTS (OUT): select Article → Batch dropdown auto-filters to open batches for that article.",
        "Qty_Drawn is blocked if it exceeds Available_Qty — over-drafts are prevented by validation.",
        "Missing Reason cells are highlighted in amber when Article is filled.",
        "",
        "STOCK REGISTER: one tab replaces three. Use the 'Show:' dropdown (top-right, cell E1) to switch view:",
        "  • Open Batches (default) — your daily working view",
        "  • Depleted Batches — for audit trail",
        "  • All Batches — full register",
        "",
        "MOVEMENT HISTORY: all movements sorted by date descending — read-only audit log.",
        "",
        "ARCHIVING: Stock Tools > Archive old movements (90+ days) moves old rows to Stock_Movements_Archive (hidden).",
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
    ws_dashboard.set_column("L:L", 10)
    ws_dashboard.set_column("M:O", 14)
    ws_dashboard.set_column("P:Q", 20)
    ws_dashboard.set_column("R:R", 10)
    ws_dashboard.set_column("S:S", 32)
    ws_dashboard.freeze_panes(11, 0)

    # ── Tab colours ───────────────────────────────────────────────────────────
    ws_dashboard.set_tab_color("#455A64")          # dark slate — overview
    ws_procurement.set_tab_color("#D97706")         # amber — data entry IN
    ws_movements.set_tab_color("#D97706")           # amber — data entry OUT
    ws_register.set_tab_color("#546E7A")            # slate-blue — stock view
    ws_movement_history.set_tab_color("#546E7A")    # slate-blue — audit view
    ws_detail.set_tab_color("#B0BEC5")              # light grey — hidden backing
    ws_summary.set_tab_color("#B0BEC5")
    ws_movements_archive.set_tab_color("#B0BEC5")
    ws_articles.set_tab_color("#4E7B5C")            # sage green — admin/setup
    ws_suppliers.set_tab_color("#4E7B5C")
    ws_catalog.set_tab_color("#4E7B5C")

    workbook.close()
    print(f"Created {OUTPUT_FILE}")


if __name__ == "__main__":
    main()

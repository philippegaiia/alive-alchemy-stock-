from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

WORKBOOK = Path(__file__).with_name("stock_management.xlsx")

REQUIRED_SHEETS = {
    "Dashboard",
    "Procurement",
    "Stock_Movements",
    "Stock_Register",
    "Movement_History",
    "Articles",
    "Suppliers",
    "Supplier_Catalog",
    "Stock_Detail",
    "Stock_Summary",
    "Stock_Movements_Archive",
    "Lists",
}

WB_NS = {"": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
RELS_NS = {"": "http://schemas.openxmlformats.org/package/2006/relationships"}
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def _sheet_names() -> set:
    with ZipFile(WORKBOOK) as zf:
        root = ET.fromstring(zf.read("xl/workbook.xml").decode("utf-8"))
    return {s.get("name") for s in root.findall(".//sheet", WB_NS)}


def _sheet_xml(zf, sheet_name) -> str:
    """Read a sheet's XML by name from an open ZipFile."""
    wb_root = ET.fromstring(zf.read("xl/workbook.xml").decode("utf-8"))
    sheets_meta = wb_root.findall(".//sheet", WB_NS)

    target_rid = None
    for s in sheets_meta:
        if s.get("name") == sheet_name:
            target_rid = s.get(RID)
            break
    assert target_rid is not None, f"{sheet_name} sheet not found"

    wb_rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels").decode("utf-8"))
    for rel in wb_rels.findall("Relationship", RELS_NS):
        if rel.get("Id") == target_rid:
            sheet_path = "xl/" + rel.get("Target").lstrip("/")
            return zf.read(sheet_path).decode("utf-8")
    raise RuntimeError(f"Could not find worksheet XML for {sheet_name}")


def main():
    assert WORKBOOK.exists(), f"{WORKBOOK} not found"

    sheets = _sheet_names()
    missing = REQUIRED_SHEETS - sheets
    assert not missing, f"Missing sheets: {missing}"

    with ZipFile(WORKBOOK) as zf:
        movements_xml = _sheet_xml(zf, "Stock_Movements")

    # No named ranges in formulas (would cause #NAME?)
    forbidden_names = ["DetailQtyRemaining", "DetailArticles", "DetailBatches", "MoveArticles"]
    found = [n for n in forbidden_names if n in movements_xml]
    assert not found, f"Stock_Movements references named ranges: {found}"

    # IFNA replaced with IFERROR for broad compatibility
    assert "IFNA" not in movements_xml, "Stock_Movements still uses IFNA"

    # Articles sheet must have data — without it ALL downstream formulas fail
    # xlsxwriter stores strings in sharedStrings.xml, so check there + verify
    # the Articles sheet has row elements beyond the header.
    with ZipFile(WORKBOOK) as zf:
        shared = zf.read("xl/sharedStrings.xml").decode("utf-8")
        articles_xml = _sheet_xml(zf, "Articles")
    assert "Argan Oil" in shared, (
        "Articles data missing from shared strings — all VLOOKUPs/SUMIFS/COUNTIFs will return empty/0"
    )
    assert articles_xml.count("<row ") >= 28, (
        f"Articles sheet has only {articles_xml.count('<row ')} rows — expected at least 28 (header + 27 articles)"
    )

    print("PASS")


if __name__ == "__main__":
    main()

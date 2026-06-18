# AGENTS.md

> **Read this file completely before doing anything.** This is the source of truth for this project. Do not guess, do not assume — if something isn't documented here, read the relevant source file and update this document.

## Project Overview

**Alive Alchemy** — batch-tracked raw materials stock management system for a small cosmetic-oils manufacturer (~80 articles, 15 suppliers, 30 production batches/month).

Two parallel versions, both generated from Python (`xlsxwriter`):
- **`google_sheets/`** — primary version, uses Apps Script for dependent dropdowns
- **`excel/`** — Excel 365 version (current active focus), uses VBA for dependent dropdowns

## Current State

### Working
- Excel workbook generator produces correct `.xlsx` with all 11 sheets
- All formulas use **direct sheet references** (no named ranges) — eliminates `#NAME?`
- All formulas use `IFERROR` (not `IFNA`) for broad compatibility
- Articles sheet populated with 27 articles
- VBA dropdowns (2 files) work on Mac Excel with `Collection` instead of `Scripting.Dictionary`
- Sheet protection active — formula cells locked, input cells unlocked
- Double-click Date cell (col A) in Procurement/Stock_Movements to stamp today's date
- Dependent dropdowns: unprotect sheet before writing validation, re-protect after
- Catalog Price/Last_Updated auto-calculate from latest Procurement (LOOKUP formula)
- Archive clears rows (not hide); Restore button brings them back
- Tests pass: `python3 test_excel_workbook.py`

### Known limitations
- `Stock_Register` and Dashboard alert tables use `FILTER`/`SORT`/`TAKE` — requires Excel 365/2021+
- `LOOKUP` array pattern in Supplier_Catalog requires Excel 365
- No LibreOffice version in this repo

## Architecture

### Data flow
```
Procurement (IN)  →  Stock_Detail (hidden engine)  ←  Stock_Movements (OUT)
                         ↓
                    Stock_Summary (hidden, per-article aggregates)
                         ↓
                      Dashboard (visible, metrics + alerts)
```

### Sheet inventory (11 sheets)

| Sheet | Visible | Tab colour | Purpose |
|---|---|---|---|
| Dashboard | Yes | Slate | Metrics, low-stock alerts, recent procurement, Stock Tools buttons |
| Procurement | Yes | Amber | Log every purchase/production — one row per batch |
| Stock_Movements | Yes | Amber | Log every draw — one row per usage event |
| Stock_Register | Yes | Slate-blue | Live view of batches (Open/Depleted/All filter in E1) |
| Stock_Detail | No | Grey | Core calculation engine — one row per batch (14 cols) |
| Stock_Summary | No | Grey | One row per article — feeds Dashboard |
| Stock_Movements_Archive | No | Grey | Movements 90+ days old (moved by archive macro) |
| Articles | No | Green | Master list (ID, Name, Category, Unit, Reorder_Level, Active) |
| Suppliers | No | Green | Supplier contacts (6 suppliers incl. "Alive Alchemy") |
| Supplier_Catalog | No | Green | Supplier→Article mapping (6 cols, Price/Date auto-calc) |
| Lists | No | — | Static dropdown sources + VBA scratch columns F/G |

### Column layouts

**Articles** (hidden): `Article_ID | Article_Name | Category | Unit | Reorder_Level | Active | Notes`

**Suppliers** (hidden): `Supplier_ID | Supplier_Name | Contact_Person | Phone | Email | Payment_Terms | Notes`

**Supplier_Catalog** (hidden, 6 cols): `Supplier | Article | Supplier_SKU | Price_Per_Unit(auto) | Last_Updated(auto) | Notes`

**Procurement** (visible): `Date | Supplier | Article | Unit(auto) | Quantity | Unit_Cost | Total_Cost(auto) | Invoice_Number | Batch_Lot_Number | Active | Notes`

**Stock_Movements** (visible): `Date | Article | Batch_Lot_Number | Unit(auto) | Available_Qty(auto) | Qty_Drawn | Reason | Notes | Status(auto) | Archived_On(auto)`

**Stock_Detail** (hidden, 14 cols): `Article | Category(auto) | Supplier | Invoice_Number | Batch_Lot_Number | Date | Qty_Purchased | Qty_Drawn(auto) | Qty_Remaining(auto) | Status(auto) | Unit_Cost | Stock_Value(auto) | Last_Movement_Date(auto) | Days_Since(auto)`

### Key formulas
- **Available_Qty** (Stock_Movements col E): `IFERROR(SUMIFS(Stock_Detail!I, article, batch))`
- **Status** (Stock_Detail col J): `IF(ROUND(Qty_Remaining,2)>0,"Open","Depleted")` — ROUND prevents floating-point issues
- **Qty_Drawn** (Stock_Detail col H): `SUMIFS(Stock_Movements) + SUMIFS(Stock_Movements_Archive)`
- **Qty_Remaining** (Stock_Detail col I): `Qty_Purchased - Qty_Drawn`
- **Catalog Price** (Supplier_Catalog col D): `LOOKUP(2, 1/(Procurement article match), Procurement unit cost)` — latest price
- **Status** (Stock_Movements col I): `IF(B="","","Active")` — formula only, no dropdown

## How to work on this project

### Generate the workbook
```bash
cd excel && python3 generate_excel_workbook.py
```
Requires: `pip install xlsxwriter`

### Run tests
```bash
cd excel && python3 test_excel_workbook.py
```

### Commit conventions
- Write clear commit messages explaining what changed and why
- Always run the test before committing
- Never commit generated `.xlsx`/`.xlsm` files (they're gitignored)
- Push after committing: `git push`

### When fixing a bug
1. Read the relevant section of the generator script
2. Fix it
3. Regenerate: `python3 generate_excel_workbook.py`
4. Test: `python3 test_excel_workbook.py`
5. Update the "Current State" section in this file
6. Commit and push

## VBA setup (Excel version only)

Only **2 files** are needed:
1. `StockDropdowns_Module.bas` — **File → Import File** in VBA editor (has `Attribute VB_Name`, required for import)
2. `ThisWorkbook.bas` — **paste** into the ThisWorkbook module (do NOT import; no `Attribute VB_Name`)

User must Save As `.xlsm` before VBA works.

`ThisWorkbook.bas` contains three event handlers:
- `Workbook_Open` — re-protects all sheets with `UserInterfaceOnly:=True` (allows VBA to modify cells)
- `Workbook_SheetChange` — dependent dropdowns + Stock_Register recalculation; unprotects sheet before writing validation, re-protects after
- `Workbook_SheetBeforeDoubleClick` — double-click Date cell (col A) in Procurement/Stock_Movements stamps today's date

## Constraints and gotchas

### DO NOT do these things
- **Do NOT use named ranges in formulas.** They cause `#NAME?` on the user's Mac. Use direct references like `Stock_Detail!$I$2:$I$2001`.
- **Do NOT use `IFNA`.** Not recognized on some Mac Excel versions. Use `IFERROR`.
- **Do NOT use `Scripting.Dictionary` in VBA.** Causes "ActiveX component can't create object" error 429 on Mac. Use VBA `Collection` instead.
- **Do NOT use `TEXT(date, "YYMMDD")` in formulas.** Format codes are locale-specific and break in Swiss/French locale.
- **Do NOT add autofilter or freeze_panes to hidden sheets.** Excel rejects these on hidden sheets.
- **Do NOT put `.Select` before `.Activate`.** On Mac, a sheet must be activated before selecting a range on it.
- **Do NOT create 4 separate VBA files.** Use 2: the module (import) + ThisWorkbook (paste).
- **Do NOT put event handlers in individual sheet modules.** Handle everything in `ThisWorkbook`'s `Workbook_SheetChange`.
- **Do NOT write `Validation.Add` on a protected sheet without unprotecting first.** Even `UserInterfaceOnly:=True` does not allow it.
- **Do NOT overwrite formula cells with static values in the generator.** (e.g., writing "Active" to a Status formula cell.)

### User environment
- **OS**: macOS (MacBook Pro)
- **Excel**: Excel 365 on Mac (free trial; file will be handed off to another user)
- **Locale**: Swiss/French — semicolons are list separators, date format codes are localized
- **Python**: system `python3` (no virtualenv)
- **Git**: repo lives inside Google Drive (`~/My Drive/alive_alchemy_stock/both_versions/`)

### Constants (in generator)
```python
MAX_ARTICLE_ROWS = 200
MAX_SUPPLIER_ROWS = 100
MAX_PROCUREMENT_ROWS = 2000
MAX_MOVEMENT_ROWS = 2000
MAX_CATALOG_ROWS = 500
ARCHIVE_AGE_DAYS = 90
```

## File map

```
both_versions/                    ← repo root
├── AGENTS.md                     ← THIS FILE — read first
├── README.md                     ← project overview
├── .gitignore                    ← excludes *.xlsx, *.xlsm, .DS_Store, ~$*
├── excel/
│   ├── generate_excel_workbook.py  ← main generator (Python + xlsxwriter)
│   ├── test_excel_workbook.py      ← regression test
│   ├── StockDropdowns_Module.bas   ← VBA standard module (import via File > Import)
│   ├── ThisWorkbook.bas            ← VBA workbook events (paste into ThisWorkbook)
│   ├── WORKBOOK_AUDIT.md           ← sheet-by-sheet analysis
│   └── README.md                   ← detailed Excel setup guide
└── google_sheets/
    ├── generate_stock_workbook.py  ← Google Sheets generator (writes .xlsx)
    └── dependent_dropdown.gs       ← Google Apps Script for dependent dropdowns
```

## Decision log

| Date | Decision | Reason |
|---|---|---|
| Jun 2026 | Removed Movement_History sheet | Redundant with Stock_Movements; broken (spill blocked by protection) |
| Jun 2026 | Added sheet protection (formula cells locked) | Prevent end user from accidentally overwriting formulas |
| Jun 2026 | Added double-click date stamping | UX: faster data entry without Ctrl+; |
| Jun 2026 | Catalog Price auto-calculates from Procurement | Prices stay current without manual updates |
| Jun 2026 | Removed Supplier_Catalog redundant cols B&D | Exact copies of A&C, never read by VBA |
| Jun 2026 | Removed Archive_Status col from Stock_Detail | Dead flag, no macro reads it |
| Jun 2026 | Archive now clears rows (not hide); added Restore | Old approach left hidden rows with stale Status values |
| Jun 2026 | Removed over-draw validation on Qty_Drawn | Real-world business has variations; user adjusts later |
| Jun 2026 | Status formula uses ROUND to prevent float issues | 100-100 can = 1.4e-15 in Excel, keeping batch "Open" |
| Jun 2026 | Removed all named ranges from formulas | `#NAME?` errors on user's Mac Excel |
| Jun 2026 | Replaced `IFNA` with `IFERROR` | `IFNA` not recognized on some Mac Excel configs |
| Jun 2026 | Added Articles data to generator | Sheet was completely empty — root cause of zero Dashboard metrics |
| Jun 2026 | `Scripting.Dictionary` → `Collection` | Mac VBA doesn't support ActiveX Dictionary (error 429) |
| Jun 2026 | Consolidated 4 VBA files → 2 | Reduce setup errors; fewer paste steps for user |
| Jun 2026 | Moved sheet events → `ThisWorkbook` | Eliminates paste-into-sheet-module errors |
| Jun 2026 | Added "Alive Alchemy" as supplier | Production batches had empty Supplier field |

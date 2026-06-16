# Alive Alchemy — Stock Management System

Batch-tracked raw materials stock register for a small cosmetic-oils manufacturer. Tracks 50–80 articles across 9 categories, 10–15 suppliers, and ~30 production batches/month.

Built in two parallel versions — **Google Sheets** (primary) and **Excel 365** — generated from Python scripts so the structure is always reproducible.

## Repo structure

```
excel/            Excel 365 version (current focus)
google_sheets/    Google Sheets version (primary)
```

Each version has its own generator script, test file, and setup instructions.

## Key design decisions

- Stock tracked **per batch** (one Procurement row = one batch)
- `Stock_Detail` (hidden) is the calculation engine — one row per batch
- `Stock_Movements` uses positive `Qty_Drawn`; `Stock_Detail` subtracts it
- Separate `Invoice_Number` and `Batch_Lot_Number` columns
- Hidden backing sheets keep the workspace clean (5 visible tabs)
- Dependent dropdowns: Article filters by Supplier (Procurement), Batch filters by Article (Movements)

## Generate a workbook

```bash
# Excel version
cd excel && python3 generate_excel_workbook.py

# Google Sheets version
cd google_sheets && python3 generate_stock_workbook.py
```

Requires `pip install xlsxwriter`.

## Test

```bash
cd excel && python3 test_excel_workbook.py
```

## Changelog

- **v3** (current): Fixed Articles sheet missing data (root cause of all formulas returning 0). Replaced `IFNA` with `IFERROR` for broad compatibility. All formulas use direct sheet references (no named ranges).
- **v2**: Replaced named ranges with direct references. Consolidated VBA from 4 files to 2. Moved sheet events to `ThisWorkbook`.
- **v1**: Initial generator with 13 sheets, 27 named ranges, data validations, and dependent dropdowns.

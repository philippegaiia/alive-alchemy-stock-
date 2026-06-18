# Alive Alchemy — Raw Materials Stock Register (Excel 365)

A batch-tracked raw materials register built entirely in Excel 365. No database, no cloud service — just a well-structured `.xlsm` workbook with a thin VBA layer for dependent dropdowns and archiving.

---

## What it does

### The register in one sentence
Every purchase creates a **batch** (one row in the hidden `Stock_Detail` sheet). Every time you draw from that batch, the remaining quantity updates automatically. The `Stock_Register` tab shows you, at a glance, what you still have and what is depleted.

### The 4 working tabs

| Tab | Colour | Purpose |
|---|---|---|
| **Dashboard** | Dark slate | Stock alerts (Low / Warning), recent procurement, metrics, Stock Tools buttons |
| **Procurement** | Amber | Log every purchase or production — one row per batch received |
| **Stock_Movements** | Amber | Log every draw — one row per usage event |
| **Stock_Register** | Slate-blue | Live view of all batches; filter by Open / Depleted / All |

Everything else (Articles, Suppliers, Supplier_Catalog, Stock_Detail, Stock_Summary, Stock_Movements_Archive, Lists) is **hidden** to keep the workspace clean. Right-click any tab → **Unhide** to access them.

---

## Daily workflow

### Logging a purchase or production batch (IN)
1. Go to **Procurement** — **double-click the Date cell** (column A) to stamp today's date.
2. Select **Supplier** from the dropdown (column B). The **Article** dropdown (column C) auto-filters to that supplier's catalogue.
3. Fill in **Quantity** and **Unit Cost**. `Total_Cost` calculates automatically.
4. Fill in the **Invoice Number** (column H).
5. Type your **Batch_Lot_Number** (column I) — e.g. `LOT-ARG-2401`.  
   *The cell turns red if you leave Batch_Lot_Number empty after filling Article.*
6. Set **Active** to `Yes`.

### Logging a draw (OUT)
1. Go to **Stock_Movements** — **double-click the Date cell** to stamp today's date.
2. Select **Article** (column B). The **Batch_Lot_Number** dropdown (column C) auto-filters to open batches with remaining stock.
3. The **Available_Qty** (column E) shows the current remaining quantity for that batch — read-only.
4. Enter **Qty_Drawn** (column F). Can exceed Available_Qty (real-world variations — adjust later).  
   *The Reason cell turns amber if you leave it empty.*
5. Select a **Reason** (column G): Production / Waste / Return / Correction / Other.

### Checking stock
- Open **Stock_Register**. Cell **E1** has a dropdown — choose:
  - `Open Batches` (default) — your daily view
  - `Depleted Batches` — audit trail of finished batches
  - `All Batches` — full register
- The **Dashboard** highlights any article that is **Low** (below reorder level) or **Warning** (below 1.5× reorder level).

---

## Setup

### Step 1 — Generate the workbook

```bash
cd excel
python3 generate_excel_workbook.py
```

Requires Python 3 and `xlsxwriter`:
```bash
pip install xlsxwriter
```

### Step 2 — Save as macro-enabled

Open `stock_management.xlsx` in Excel 365, then:

**File → Save As → stock_management.xlsm** *(Macro-Enabled Workbook)*

> [!IMPORTANT]
> VBA will not run from a `.xlsx` file. You must save as `.xlsm` before the next step.

### Step 3 — Import the VBA code

Open the VBA editor: **Alt+F11** (Windows) or **Option+F11** (Mac).

You will see a **Project Explorer** pane on the left. If it's hidden, press **Ctrl+R**.

#### 3a — Import the standard module

**File → Import File** → select `StockDropdowns_Module.bas`

This adds all dropdown logic, navigation, archiving, and restore procedures.

#### 3b — Workbook events

Double-click **ThisWorkbook** in Project Explorer → delete any existing code → paste the entire contents of `ThisWorkbook.bas`.

> [!IMPORTANT]
> If you previously pasted code into the **Procurement** or **Stock_Movements** sheet modules, delete it now. The new setup handles sheet events from `ThisWorkbook`.

#### 3c — Save and reopen

Press **Ctrl+S** in the VBA editor, close it, then **close and reopen** `stock_management.xlsm`.

On reopen, the workbook automatically re-applies sheet protection with `UserInterfaceOnly:=True`, which allows VBA to manage dropdowns while protecting formulas from accidental edits.

The **Stock Tools buttons** live on the **Dashboard** sheet (top-right). They work on both Mac and Windows.

> [!NOTE]
> If Excel asks whether to enable macros when opening, always click **Enable**.

---

## VBA files reference

| File | Where to paste | What it does |
|---|---|---|
| `StockDropdowns_Module.bas` | Import as module | Dropdown logic, navigation, archiving, restore |
| `ThisWorkbook.bas` | ThisWorkbook module | Sheet protection on open + sheet-change events for dependent dropdowns + double-click date stamping |

---

## Sheet protection

All sheets are protected. This prevents accidental edits to formula cells.

- **Input cells** (white background) are editable
- **Formula cells** (auto-calculated) are locked — cannot be overwritten
- VBA manages dropdowns and archiving normally (unprotects/re-protects internally)
- To manually unprotect: **Review → Unprotect Sheet** (no password)

---

## Hidden sheets reference

| Sheet | Purpose |
|---|---|
| **Articles** | Master list of raw materials and reorder levels |
| **Suppliers** | Supplier contact details (includes "Alive Alchemy" for production) |
| **Supplier_Catalog** | Which articles each supplier carries. Price and Last_Updated auto-fill from latest Procurement |
| **Stock_Detail** | One row per procurement batch — the core calculation engine (14 columns) |
| **Stock_Summary** | One row per article — aggregates Stock_Detail for the Dashboard |
| **Stock_Movements_Archive** | Movements older than 90 days, moved here by the archive macro |
| **Lists** | Dropdown source lists (Categories, Units, Reasons). Columns F–G used as VBA scratch |

To reveal any hidden sheet: **right-click any tab → Unhide**.

---

## Stock Tools (Dashboard buttons)

| Button | What it does |
|---|---|
| Go to empty Procurement row | Navigates to the first empty row in Procurement |
| Go to empty Movement row | Navigates to the first empty row in Stock_Movements |
| Archive old movements | Moves movements older than 90 days to the hidden archive sheet. Clears the rows in Stock_Movements (reusable for new data). |
| Restore archived movements | Moves all archived movements back to Stock_Movements and clears the archive |

---

## Built-in guardrails

| Guard | Where | What it prevents |
|---|---|---|
| **Missing Date highlight** | Procurement & Stock_Movements | Red cell when Article is filled but Date is empty |
| **Missing Batch highlight** | Procurement, Batch_Lot_Number | Red cell when Article is filled but Batch is empty |
| **Missing Reason highlight** | Stock_Movements, Reason | Amber cell when Article is filled but Reason is empty |
| **Batch validation** | Stock_Movements, Batch_Lot_Number | Warns if batch doesn't exist in Procurement |
| **Duplicate detection** | Articles & Suppliers | Red highlight on duplicate names |
| **Sheet protection** | All sheets | Prevents overwriting formula cells |
| **Date validation** | Procurement & Stock_Movements | Future dates are rejected |

---

## Tips

- **Date shortcut**: double-click the Date cell (column A) to stamp today's date. Or press `Ctrl+;`.
- **Add a new article**: right-click any tab → Unhide → **Articles**. Add a row. The article will appear in all dropdowns immediately.
- **Add a new supplier**: same process with the **Suppliers** sheet. Then add its articles to **Supplier_Catalog**.
- **Catalog prices auto-update**: when you add a procurement entry, the Supplier_Catalog automatically shows the latest unit cost and date.

---

## Compatibility

| Feature | Requirement |
|---|---|
| `FILTER`, `SORT`, `TAKE` (Stock_Register, Dashboard) | **Excel 365** or Excel 2021+ |
| `LOOKUP` array pattern (Supplier_Catalog) | Excel 365 |
| `MAXIFS`, `SUMIFS` (Stock_Detail calculations) | Excel 2019+ |
| VBA dependent dropdowns | Any Excel version that supports `.xlsm` |
| Sheet protection + double-click date | Any modern Excel |

> [!WARNING]
> The `Stock_Register` and Dashboard alert tables use dynamic array functions (`FILTER`, `SORT`, `TAKE`) that are **not available in Excel 2016 or earlier**.

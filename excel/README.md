# Alive Alchemy — Raw Materials Stock Register (Excel 365)

A batch-tracked raw materials register built entirely in Excel 365. No database, no cloud service — just a well-structured `.xlsm` workbook with a thin VBA layer for dependent dropdowns and archiving.

---

## What it does

### The register in one sentence
Every purchase creates a **batch** (one row in the hidden `Stock_Detail` sheet). Every time you draw from that batch, the remaining quantity updates automatically. The `Stock_Register` tab shows you, at a glance, what you still have and what is depleted.

### The 5 working tabs

| Tab | Colour | Purpose |
|---|---|---|
| **Dashboard** | 🔵 Dark slate | Stock alerts (Low / Warning), recent procurement, metrics |
| **Procurement** | 🟠 Amber | Log every purchase — one row per batch received |
| **Stock_Movements** | 🟠 Amber | Log every draw — one row per usage event |
| **Stock_Register** | 🔵 Slate-blue | Live view of all batches; filter by Open / Depleted / All |
| **Movement_History** | 🔵 Slate-blue | All movements sorted newest-first (read-only audit log) |

Everything else (Articles, Suppliers, Supplier_Catalog, Stock_Detail, Stock_Summary, Stock_Movements_Archive) is **hidden** to keep the workspace clean. Right-click any tab → **Unhide** to access them.

---

## Daily workflow

### Logging a purchase (IN)
1. Go to **Procurement** — press `Ctrl+;` to stamp today's date in column A.
2. Select **Supplier** from the dropdown (column B). If VBA is set up, the **Article** dropdown (column C) auto-filters to that supplier's catalogue.
3. Fill in **Quantity** and **Unit Cost**. `Total_Cost` calculates automatically.
4. Fill in the **Invoice Number** (column H).
5. Type your **Batch_Lot_Number** (column I) — e.g. `LOT-ARG-2401`.  
   *The cell turns red if you leave Batch_Lot_Number empty after filling Article.*
6. Set **Active** to `Yes`.

### Logging a draw (OUT)
1. Go to **Stock_Movements** — press `Ctrl+;` for the date.
2. Select **Article** (column B). If VBA is set up, the **Batch_Lot_Number** dropdown (column C) auto-filters to open batches with remaining stock.
3. The **Available_Qty** (column E) shows the current remaining quantity for that batch — read-only.
4. Enter **Qty_Drawn** (column F). Excel blocks entry if the quantity exceeds the available stock.  
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
cd both_versions/excel
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

This adds all menu procedures, dropdown logic, archiving, and named-range repair.

#### 3b — Workbook events

Double-click **ThisWorkbook** in Project Explorer → delete any existing code → paste the entire contents of `ThisWorkbook.bas`.

> [!IMPORTANT]
> If you previously pasted code into the **Procurement** or **Stock_Movements** sheet modules, delete it now. The new setup handles sheet events from `ThisWorkbook`, so old code in those modules would make dropdowns refresh twice.

#### 3c — Save and reopen

Press **Ctrl+S** in the VBA editor, close it, then **close and reopen** `stock_management.xlsm`.

On reopen, the workbook will automatically repair missing named ranges, so `Available_Qty` will calculate correctly.

The **Stock Tools buttons** live on the **Dashboard** sheet (top-right). They work on both Mac and Windows.

> [!NOTE]
> If Excel asks whether to enable macros when opening, always click **Enable**.

---

## VBA files reference

| File | Where to paste | What it does |
|---|---|---|
| `StockDropdowns_Module.bas` | Import as module | Dropdown logic, navigation, archiving, named-range repair |
| `ThisWorkbook.bas` | ThisWorkbook module | Auto-repair on open + sheet-change events for dependent dropdowns |

(The old `Procurement_Sheet.bas` and `StockMovements_Sheet.bas` files are no longer needed.)

---

## Hidden sheets reference

| Sheet | Purpose |
|---|---|
| **Articles** | Master list of raw materials and reorder levels |
| **Suppliers** | Supplier contact details |
| **Supplier_Catalog** | Which articles each supplier carries (drives Article dropdown) |
| **Stock_Detail** | One row per procurement batch — the core calculation engine |
| **Stock_Summary** | One row per article — aggregates Stock_Detail for the Dashboard |
| **Stock_Movements_Archive** | Movements older than 90 days, moved here by the archive macro |
| **Lists** | Dropdown source lists (Categories, Units, Reasons). Columns F–G used as VBA scratch to bypass Excel's 256-character validation limit |

To reveal any hidden sheet: **right-click any tab → Unhide**.

---

## Built-in guardrails

| Guard | Where | What it prevents |
|---|---|---|
| **Over-draft block** | Stock_Movements, Qty_Drawn | Entering a draw quantity that exceeds Available_Qty |
| **Missing Batch highlight** | Procurement, Batch_Lot_Number | Red cell when Article is filled but Batch is empty |
| **Missing Reason highlight** | Stock_Movements, Reason | Amber cell when Article is filled but Reason is empty |
| **Date validation** | Procurement & Stock_Movements | Future dates are rejected |

---

## Tips

- **Date shortcut**: `Ctrl+;` inserts today's date instantly — no macro needed.
- **Navigate to next empty row**: click the **Dashboard** Stock Tools buttons, or use Stock Tools → *Go to first empty row (Procurement / Stock Movements)*.
- **Archive old movements**: click the **Dashboard** button, or use Stock Tools → *Archive old Stock Movements (90+ days)*. Rows older than 90 days are moved to `Stock_Movements_Archive` (hidden) and hidden in `Stock_Movements`. Use *Unhide all archived rows* to bring them back.
- **Add a new article**: right-click any tab → Unhide → **Articles**. Add a row. The article will appear in all dropdowns immediately.
- **Add a new supplier**: same process with the **Suppliers** sheet. Then add its articles to **Supplier_Catalog**.

---

## Compatibility

| Feature | Requirement |
|---|---|
| `FILTER`, `SORT`, `TAKE` (Stock_Register, Movement_History, Dashboard) | **Excel 365** or Excel 2021+ |
| `MAXIFS`, `SUMIFS` (Stock_Detail calculations) | Excel 2019+ |
| VBA dependent dropdowns | Any Excel version that supports `.xlsm` |
| Core procurement/movements data entry | Any modern Excel |

> [!WARNING]
> The `Stock_Register`, `Movement_History`, and Dashboard alert tables use dynamic array functions (`FILTER`, `SORT`, `TAKE`) that are **not available in Excel 2016 or earlier**. Those cells will show `#NAME?` on older versions.

# Workbook Audit — Sheet-by-Sheet Analysis

## Dashboard (visible)

**Purpose:** Overview screen with KPIs, stock alerts, recent procurement, and Stock Tools buttons.

### Sections
| Section | Usefulness | Notes |
|---|---|---|
| Metrics (A4:B8) | High | Total Articles, Active, Low, Warning, Stock Value. Straightforward COUNTIF/SUM formulas. |
| Stock Tools buttons (J4:J7) | High | Navigation + archiving. Only UI for macros. |
| Low/Warning Alerts (A13) | High | Dynamic array from Stock_Summary. Daily-use feature. |
| Recent Procurement (I13) | Medium | Last 10 procurements. Useful for quick glance. |
| How It Works text (A34+) | Low | Setup instructions. Useful first time, clutter after. |

### Issues
- Title says "v3 - articles data + IFERROR" — technical jargon meaningless to end user.
- `COUNTIF(range,"<>")` for Total Articles counts ALL articles including inactive. Label should say "Total Articles" (it does) but metric might mislead if many are inactive.
- No metric for total batches or open batches count.
- How It Works text is baked into the sheet — can't be edited without regenerating.

---

## Procurement (visible)

**Purpose:** Log every stock IN event (purchases and production batches).

### Columns
| Col | Header | Type | Usefulness | Notes |
|---|---|---|---|---|
| A | Date | Input | Essential | Validated: must be date, not future. Red highlight if missing. |
| B | Supplier | Input (dropdown) | Essential | From Suppliers sheet. Empty for production batches. |
| C | Article | Input (VBA dropdown) | Essential | Dependent on Supplier. Placeholder "(Select Supplier First)" without VBA. |
| D | Unit | Auto (VLOOKUP) | Verification only | From Articles. Read-only. |
| E | Quantity | Input | Essential | Must be > 0. |
| F | Unit_Cost | Input | Essential | Must be >= 0. Drives stock valuation. |
| G | Total_Cost | Auto (E×F) | Reference | Calculated. |
| H | Invoice_Number | Input | Accounting | Free text. |
| I | Batch_Lot_Number | Input | Essential | Key field for batch tracking. Red highlight if missing. |
| J | Active | Input (dropdown) | Essential | Yes/No. Controls whether batch appears in stock. |
| K | Notes | Input | Optional | Free text. |

### Issues
- **Production batches have empty Supplier** (rows 11-12 in sample data). These are self-produced goods, not purchases. The dependent dropdown can't filter articles by supplier when supplier is blank. Design flaw: production and procurement are mixed in one sheet.
- **No validation on column C without VBA.** The placeholder list `["(Select Supplier First)"]` is useless if macros aren't enabled. User can still type freely but gets no dropdown.
- **No highlight when Quantity or Unit_Cost is missing.** Only Date and Batch are conditionally formatted.
- **Active column semantics unclear.** "No" makes a batch vanish from stock calculations silently. No explanation on sheet.
- **Date validation rejects future dates.** Prevents logging expected/pre-paid deliveries dated in the future.

---

## Stock_Movements (visible)

**Purpose:** Log every stock OUT event (production draws, waste, returns, corrections).

### Columns
| Col | Header | Type | Usefulness | Notes |
|---|---|---|---|---|
| A | Date | Input | Essential | Validated: not future. Red highlight if missing. |
| B | Article | Input (dropdown) | Essential | From Articles. Triggers batch dropdown via VBA. |
| C | Batch_Lot_Number | Input (VBA dropdown) | Essential | Filtered to open batches for selected article. **No validation without VBA.** |
| D | Unit | Auto (VLOOKUP) | Verification | From Articles. |
| E | Available_Qty | Auto (SUMIFS) | High | Shows remaining stock for article+batch. Display-only. |
| F | Qty_Drawn | Input | Essential | Must be > 0. Over-draw allowed. |
| G | Reason | Input (dropdown) | Essential | Production/Waste/Return/Correction/Other. Amber if missing. |
| H | Notes | Input | Optional | Free text. |
| I | Status | Auto formula + dropdown | **Problematic** | See issues below. |
| J | Archived_On | Auto formula + VBA | Low | Only populated by archive macro. |

### Issues
- **Status column conflict:** Column I has BOTH a formula (`=IF(B="","","Active")`) AND a data validation dropdown (Active/Archived). If user manually selects "Archived" from dropdown, it overwrites the formula. Then if the row is edited, the formula is gone. The dropdown should be removed — Status is managed by VBA only.
- **No batch validation without VBA.** Column C has no data_validation at all in the generator. Without VBA, user can type any batch number, including nonexistent ones.
- **Available_Qty can show stale values.** No `Application.Calculate` triggered when user selects article in column B. The SUMIFS reads Stock_Detail which might not have recalculated.
- **Archived_On formula** `=IF(B="","","")` always returns "". It exists only as a placeholder. The archive macro replaces it with a timestamp. Redundant formula.

---

## Stock_Register (visible)

**Purpose:** Live view of all batches with a filter dropdown (Open/Depleted/All).

### Issues
- **Inconsistent sort keys:** Depleted Batches sorts by Date descending (col 6). Open Batches and All Batches sort by Article ascending (col 1). User switching between views sees data jump around.
- **Shows 15 columns** including Archive_Status and Days_Since_Last_Movement which are operational flags, not useful for daily reading.
- Filter dropdown in E1 has no VBA recalculate on all platforms — we added `Application.Calculate` in ThisWorkbook but only fires if VBA is set up.

### Utility
High when working. Replaces what would otherwise be three separate sheets. The filter is the right design.

---

## Movement_History (visible)

**Purpose:** Read-only audit log of all stock movements, sorted newest-first.

### Issues
- **Shows archived movements too.** Formula filters on `Stock_Movements!A2:A2001<>""` (Date not empty). Archived rows are hidden in Stock_Movements but still have dates, so they appear here. This is probably intentional (complete audit log) but could confuse users who expect Movement_History to mirror Stock_Movements.
- No filter or date-range selector. For 2000+ movements, this becomes a long scroll.

### Utility
Medium. Useful for audit, but Stock_Movements with autofilter provides similar capability.

---

## Stock_Detail (hidden) — calculation engine

**Purpose:** One row per procurement batch. Calculates drawn quantities, remaining stock, status.

### Columns
| Col | Header | Source | Usefulness | Notes |
|---|---|---|---|---|
| A | Article | Formula → Procurement C | Essential | 1:1 mapping with Procurement row. |
| B | Category | Formula → VLOOKUP Articles | Medium | For Register display. |
| C | Supplier | Formula → Procurement B | Medium | For Register display. |
| D | Invoice_Number | Formula → Procurement H | Low | For Register display. |
| E | Batch_Lot_Number | Formula → Procurement I | Essential | Key for SUMIFS matching. |
| F | Date | Formula → Procurement A | Medium | For Register display. |
| G | Qty_Purchased | Formula → Procurement E | Essential | From Procurement. |
| H | Qty_Drawn | Formula → SUMIFS | Essential | Sums active + archived movements. |
| I | Qty_Remaining | Formula → G−H | Essential | Core metric. |
| J | Status | Formula → ROUND(I,2)>0 | Essential | Open/Depleted. |
| K | Unit_Cost | Formula → Procurement F | Medium | For stock value calc. |
| L | Stock_Value | Formula → I×K | Medium | For Dashboard total. |
| M | Last_Movement_Date | Formula → MAXIFS | Medium | For staleness tracking. |
| N | Days_Since_Last_Movement | Formula → TODAY()−M | Low | Informational. |
| O | Archive_Status | Formula → flag | Low | Flags depleted+90-day batches. **No macro reads this automatically.** |

### Issues
- **1:1 row mapping with Procurement is fragile.** Stock_Detail row N reads from Procurement row N. If user inserts or deletes a row in Procurement (rather than adding at the end), the mapping breaks silently. No protection against this.
- **2000 formula rows pre-written.** Makes file larger and initial calculation slower. Empty rows return "" which is harmless but wasteful.
- **Archive_Status (col O) is a dead flag.** It computes "Archive" for depleted batches older than 90 days, but no macro automatically acts on it. The ArchiveOldMovements macro archives movements, not batches. This flag serves no automated purpose.
- **Columns B, C, D, F, K** exist only to feed the Stock_Register view. They add calculation overhead for data that's already visible in Procurement.

---

## Stock_Summary (hidden)

**Purpose:** One row per article. Aggregates Stock_Detail for Dashboard metrics.

### Columns
| Col | Header | Usefulness | Notes |
|---|---|---|---|
| A | Article | Essential | From Articles. |
| B | Category | Low | Duplicate of Articles data. |
| C | Unit | Low | Duplicate of Articles data. |
| D | Total_Stock | Essential | SUMIF of Qty_Remaining across batches. |
| E | Reorder_Level | Essential | From Articles. Drives Status. |
| F | Active | Low | Duplicate of Articles data. |
| G | Status | Essential | Low/Warning/OK based on thresholds. |
| H | Stock_Value | Medium | SUMIF of Stock_Detail L. |

### Issues
- **Status thresholds hardcoded:** Low = Total < Reorder_Level. Warning = Total < 1.5×Reorder_Level. The 1.5× multiplier is arbitrary and not configurable by the user.
- Columns B, C, F duplicate Articles data. They exist to feed the Dashboard alert table without needing a second VLOOKUP. Justified for performance but adds maintenance overhead.

---

## Stock_Movements_Archive (hidden)

**Purpose:** Stores movements older than 90 days, moved here by archive macro.

### Issues
- No formulas, no validation. Pure data dump. This is correct for an archive.
- **No way to un-archive.** Once movements are archived, the UnhideAllArchivedRows macro only unhides rows in Stock_Movements — it doesn't move data back from the archive sheet. Archived movements are permanently split across two sheets.

### Utility
Essential for keeping Stock_Movements manageable over time.

---

## Articles (hidden)

**Purpose:** Master list of all raw materials, packaging, and finished goods.

### Columns
| Col | Header | Usefulness | Notes |
|---|---|---|---|
| A | Article_ID | Low | Not used in any formula or dropdown. Human reference only. |
| B | Article_Name | Essential | Used in all dropdowns and VLOOKUPs. |
| C | Category | Medium | For grouping. Dropdown-validated. |
| D | Unit | Essential | Drives Unit column in Procurement/Movements. |
| E | Reorder_Level | Essential | Drives low-stock alerts. |
| F | Active | Essential | Yes/No. Inactive articles excluded from Summary. |
| G | Notes | Low | Free text. |

### Issues
- **No duplicate prevention on Article_Name.** Two articles with the same name cause VLOOKUPs to return the first match and SUMIFS to aggregate both — unpredictable behavior.
- **Article_ID is dead weight.** Not referenced anywhere. Could be removed or could be used as the key for dropdowns (more stable than names).

---

## Suppliers (hidden)

**Purpose:** Supplier contact directory.

### Columns
| Col | Header | Usefulness | Notes |
|---|---|---|---|
| A | Supplier_ID | Low | Used only in seed data setup, not in formulas or VBA. |
| B | Supplier_Name | Essential | Used in all dropdowns. |
| C | Contact_Person | Medium | Reference. |
| D | Phone | Medium | Reference. |
| E | Email | Medium | Reference. |
| F | Payment_Terms | Low | Reference. Not used in any calculation. |
| G | Notes | Low | Free text. |

### Issues
- **No duplicate prevention on Supplier_Name.**
- **Supplier_ID is dead weight.** VBA matches by name, not ID.

---

## Supplier_Catalog (hidden)

**Purpose:** Maps which articles each supplier carries. Drives the dependent Article dropdown in Procurement.

### Columns
| Col | Header | Usefulness | Notes |
|---|---|---|---|
| A | Supplier | Essential | Dropdown from Suppliers. Read by VBA. |
| B | Supplier_Name | **Redundant** | Formula copies A. Never read by VBA or formulas. |
| C | Article | Essential | Dropdown from Articles. Read by VBA. |
| D | Article_Name | **Redundant** | Formula copies C. Never read by VBA or formulas. |
| E | Supplier_SKU | Low | Supplier's own product code. Reference only. |
| F | Price_Per_Unit | Low | Reference only. Not used in any calculation. Procurement has its own Unit_Cost. |
| G | Last_Updated | Low | Manual date. Not used in any formula. |
| H | Notes | Low | Free text. |

### Issues
- **Columns B and D are completely redundant.** They are exact copies of A and C. Remove them.
- **Price_Per_Unit is disconnected from Procurement.** When the user logs a procurement, they re-enter the unit cost manually. The catalog price is never suggested or auto-filled. This means catalog prices drift from actual purchase prices with no feedback loop.
- **No enforcement that catalog articles match Articles sheet.** User can select any article from the dropdown, but the Supplier→Article mapping is manual. No validation that the mapping makes sense.

---

## Lists (hidden)

**Purpose:** Static dropdown sources + VBA scratch columns.

### Columns
| Col | Header | Usefulness | Notes |
|---|---|---|---|
| A | Categories (9) | Essential | Feeds Articles Category dropdown. |
| B | Units (5) | Essential | Feeds Articles Unit dropdown. |
| C | Movement Reasons (5) | Essential | Feeds Stock_Movements Reason dropdown. |
| D | Yes/No | Essential | Feeds Active dropdowns. |
| E | (unused) | None | Empty. |
| F | Article scratch | Essential | VBA writes filtered article lists here. |
| G | Batch scratch | Essential | VBA writes filtered batch lists here. |

### Issues
- None. Simple and functional.

---

## Cross-cutting issues

1. **Production batches mixed with procurement.** Self-produced goods (rows 11-12 in sample) have empty Supplier. They break the dependent dropdown and conflate two different workflows (buying vs making).

2. **No duplicate prevention** on Article_Name or Supplier_Name. A duplicate name causes silent data corruption in VLOOKUPs and SUMIFS.

3. **VBA dependency for basic functionality.** Without VBA: no dependent dropdowns, no batch validation in Stock_Movements, no archiving, no navigation buttons. The workbook is significantly less functional.

4. **Status column conflict in Stock_Movements.** Formula + dropdown on same column. The dropdown should be removed.

5. **Archive is one-way.** Movements moved to Stock_Movements_Archive cannot be moved back. The UnhideAllArchivedRows macro only unhides rows still in Stock_Movements.

6. **Archive_Status flag in Stock_Detail (col O) is unused.** No automation reads it. Either remove it or wire it to a macro.

7. **Supplier_Catalog price is disconnected** from Procurement. No auto-suggest or validation that procurement prices match catalog prices.

8. **Stock_Detail 1:1 row mapping is fragile.** Row insertions/deletions in Procurement break the mapping silently.

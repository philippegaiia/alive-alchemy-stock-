/**
 * Dependent dropdowns + navigation tools for Google Sheets.
 *
 * After importing stock_management.xlsx into Google Sheets:
 * 1. Extensions -> Apps Script
 * 2. Delete any default code, paste this entire file
 * 3. Save (Ctrl+S / Cmd+S)
 * 4. Close the Apps Script tab
 *
 * Automatic on edit (handles single cells AND multi-row pastes/autofill):
 * - Procurement: editing Supplier (col B) filters the Article dropdown.
 * - Stock_Movements: editing Article (col B) filters the Batch_Lot_Number dropdown.
 *
 * Date entry tip: use Cmd+; (Mac) or Ctrl+; (Windows) — no script needed.
 *
 * Stock Tools menu (added on file open):
 * - Go to first empty row in Procurement
 * - Go to first empty row in Stock_Movements
 * - Archive old Stock_Movements (90+ days) to Stock_Movements_Archive
 * - Unhide all archived rows
 *
 * The menu only NAVIGATES and APPENDS. It never clears or overwrites existing data.
 */

// ── Event handlers ────────────────────────────────────────────────────────────

function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const col = e.range.getColumn();
  const startRow = e.range.getRow();
  const numRows = e.range.getNumRows();

  if (startRow < 2) return;

  // Read values once — handles both single-cell edits and multi-row paste/autofill
  const values = numRows === 1
    ? [[e.value !== undefined ? e.value : e.range.getValue()]]
    : e.range.getValues();

  if (sheet.getName() === 'Procurement' && col === 2) {
    const ss = e.source;
    for (let i = 0; i < numRows; i++) {
      const row = startRow + i;
      if (row < 2) continue;
      refreshProcurementArticleDropdown(ss, sheet, row, String(values[i][0] || ''));
    }
    return;
  }

  if (sheet.getName() === 'Stock_Movements' && col === 2) {
    const ss = e.source;
    for (let i = 0; i < numRows; i++) {
      const row = startRow + i;
      if (row < 2) continue;
      refreshStockMovementBatchDropdown(ss, sheet, row, String(values[i][0] || ''));
    }
  }
}

function onOpen() {
  // Wrapped in try-catch: silently skips when running outside a browser context
  // (e.g. API triggers, scheduled runs from a mobile device).
  try {
    SpreadsheetApp.getUi()
      .createMenu('Stock Tools')
      .addItem('Go to first empty row (Procurement)', 'goToFirstEmptyProcurementRow')
      .addItem('Go to first empty row (Stock Movements)', 'goToFirstEmptyMovementRow')
      .addSeparator()
      .addItem('Archive old Stock Movements (90+ days)', 'archiveOldMovements')
      .addItem('Unhide all archived rows', 'unhideAllArchivedRows')
      .addToUi();
  } catch (_) {
    // No UI context — skip silently
  }
}

// ── Menu actions ──────────────────────────────────────────────────────────────

function goToFirstEmptyProcurementRow() {
  goToFirstEmptyRow('Procurement');
}

function goToFirstEmptyMovementRow() {
  goToFirstEmptyRow('Stock_Movements');
}

function goToFirstEmptyRow(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    showAlert('Sheet "' + sheetName + '" not found.');
    return;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    sheet.setActiveRange(sheet.getRange(2, 1));
    return;
  }
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (!data[i][0]) {
      sheet.setActiveRange(sheet.getRange(i + 2, 1));
      return;
    }
  }
  sheet.setActiveRange(sheet.getRange(lastRow + 1, 1));
}

function archiveOldMovements() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Stock_Movements');
  const archive = ss.getSheetByName('Stock_Movements_Archive');

  if (!sheet || !archive) {
    showAlert('Stock_Movements or Stock_Movements_Archive sheet is missing.');
    return;
  }

  const cutoffDays = 90;
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - cutoffDays);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    showAlert('No movements to archive.');
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  const archiveRows = [];
  const archiveSourceIndices = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (row[0] && row[0] instanceof Date && row[0] < cutoff && row[8] !== 'Archived') {
      const archivedRow = row.slice();
      archivedRow[9] = new Date();  // stamp Archived_On
      archiveRows.push(archivedRow);
      archiveSourceIndices.push(i);
    }
  }

  if (archiveRows.length === 0) {
    showAlert('No movements older than ' + cutoffDays + ' days to archive.');
    return;
  }

  // Append to archive sheet
  const archiveStart = archive.getLastRow() + 1;
  archive.getRange(archiveStart, 1, archiveRows.length, archiveRows[0].length).setValues(archiveRows);

  // Mark originals as Archived and hide — process in reverse to keep indices stable
  for (let i = archiveSourceIndices.length - 1; i >= 0; i--) {
    const realRow = archiveSourceIndices[i] + 2;  // +1 for header, +1 for 1-indexed
    sheet.getRange(realRow, 9).setValue('Archived');
    sheet.getRange(realRow, 10).setValue(new Date());
    sheet.hideRow(sheet.getRange(realRow, 1));
  }

  showAlert('Archived ' + archiveRows.length + ' movement(s) to Stock_Movements_Archive. Original rows are now hidden.');
}

function unhideAllArchivedRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Stock_Movements');
  if (!sheet) {
    showAlert('Stock_Movements sheet not found.');
    return;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    showAlert('No rows to unhide.');
    return;
  }
  const data = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
  let count = 0;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === 'Archived') {
      sheet.showRow(sheet.getRange(i + 2, 1));
      count++;
    }
  }
  showAlert('Unhid ' + count + ' archived row(s).');
}

// ── Dropdown refresh helpers ──────────────────────────────────────────────────

function refreshProcurementArticleDropdown(ss, sheet, row, supplier) {
  supplier = normalize(supplier);

  const articleCell = sheet.getRange(row, 3);  // column C = Article
  articleCell.clearContent().clearNote();

  if (!supplier) {
    articleCell.clearDataValidations();
    return;
  }

  const catalogSheet = ss.getSheetByName('Supplier_Catalog');
  if (!catalogSheet) return;

  const lastRow = catalogSheet.getLastRow();
  if (lastRow < 2) return;

  const catalogData = catalogSheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const articles = uniqueSorted(
    catalogData
      .filter(function(r) { return normalize(r[0]) === supplier && r[2]; })
      .map(function(r) { return r[2]; })
  );

  setDropdownOrNote(articleCell, articles, 'No articles in Supplier_Catalog for this supplier', { autoSelectSingle: false });
}

function refreshStockMovementBatchDropdown(ss, sheet, row, article) {
  article = normalize(article);

  const batchCell = sheet.getRange(row, 3);  // column C = Batch_Lot_Number
  batchCell.clearContent().clearNote();

  if (!article) {
    batchCell.clearDataValidations();
    return;
  }

  const batches = getOpenBatchesForArticle(ss, article);
  setDropdownOrNote(batchCell, batches, 'No active batches with stock remaining for this article', { autoSelectSingle: true });
}

/**
 * Finds open batches for the given article by reading Stock_Detail directly.
 * Stock_Detail columns (0-indexed): A=Article(0), E=Batch(4), I=QtyRemaining(8), J=Status(9).
 * Cross-references Procurement to confirm batch is Active=Yes.
 */
function getOpenBatchesForArticle(ss, article) {
  article = normalize(article);

  const detailSheet = ss.getSheetByName('Stock_Detail');
  const procSheet = ss.getSheetByName('Procurement');
  if (!detailSheet || !procSheet) return [];

  const detailLastRow = detailSheet.getLastRow();
  if (detailLastRow < 2) return [];

  // Build set of article+batch keys that are Active=Yes in Procurement
  // Procurement columns (0-indexed): C=Article(2), I=Batch(8), J=Active(9)
  const activeBatchKeys = new Set();
  const procLastRow = procSheet.getLastRow();
  if (procLastRow >= 2) {
    const procData = procSheet.getRange(2, 1, procLastRow - 1, 10).getValues();
    for (const r of procData) {
      const procArticle = normalize(r[2]);
      const batch = normalize(String(r[8]));
      const active = r[9];
      if (procArticle === article && batch && active === 'Yes') {
        activeBatchKeys.add(procArticle + '||' + batch);
      }
    }
  }

  // Filter Stock_Detail for open, active batches with remaining stock
  const detailData = detailSheet.getRange(2, 1, detailLastRow - 1, 10).getValues();
  return uniqueSorted(
    detailData
      .filter(function(r) {
        const detailArticle = normalize(r[0]);
        const batch = normalize(String(r[4]));
        const remaining = Number(r[8]) || 0;
        const status = r[9];
        return (
          detailArticle === article &&
          batch &&
          remaining > 0 &&
          status === 'Open' &&
          activeBatchKeys.has(detailArticle + '||' + batch)
        );
      })
      .map(function(r) { return String(r[4]); })
  );
}

function setDropdownOrNote(cell, values, note, options) {
  options = options || {};
  cell.clearDataValidations().clearNote();

  if (values.length === 0) {
    cell.setNote(note);
    return;
  }

  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(true)
    .build();
  cell.setDataValidation(validation);

  if (options.autoSelectSingle && values.length === 1) {
    cell.setValue(values[0]);
  }
}

// ── Utility functions ─────────────────────────────────────────────────────────

/**
 * Wraps getUi().alert() safely.
 * Silently logs if no UI context (mobile, API trigger, headless run).
 */
function showAlert(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (_) {
    Logger.log('[showAlert] ' + message);
  }
}

function uniqueSorted(values) {
  return [...new Set(values.filter(Boolean))].sort();
}

function normalize(value) {
  return String(value || '').trim();
}

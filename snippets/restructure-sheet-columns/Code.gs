/**
 * Restructure Sheet Columns
 * ------------------------------------------------------------
 * A safe, re-runnable migration for workbooks with many tabs that
 * share the same header row (e.g. one tab per subject):
 *
 *   1. Removes the columns listed in COLS_TO_DROP from every tab.
 *   2. Renames a header on every tab (RENAME_FROM -> RENAME_TO).
 *   3. Creates a new tab from a template tab, adds an extra column
 *      after a given header, and attaches dropdowns.
 *
 * Safe to re-run: missing columns are skipped and the new tab is
 * only created once. Edit the CONFIG block, then run
 * restructureWorkbook() from the Apps Script editor.
 */

/* ---------- CONFIG ---------- */
var CONFIG = {
  COLS_TO_DROP: ["Division", "Subject"],
  RENAME_FROM: "Stage",
  RENAME_TO: "Stage/Division",

  NEW_TAB: "Social Sciences (G11-G12)",
  TEMPLATE_TAB: "English (G1-G12)",
  PLACE_AFTER: "History (G1-G12)",
  INSERT_AFTER_HEADER: "Grade Level",
  NEW_COLUMN_HEADER: "Subject",
  NEW_TAB_LEVELS: ["G11", "G12"],
  NEW_COLUMN_OPTIONS: ["Economics", "Political Science", "Business", "Sociology", "Psychology"]
};

function restructureWorkbook() {
  var ss = SpreadsheetApp.getActive();

  // 1) Drop columns on every tab except the new one
  ss.getSheets()
    .filter(function (sh) { return sh.getName() !== CONFIG.NEW_TAB; })
    .forEach(function (sh) { removeColumnsByHeader_(sh, CONFIG.COLS_TO_DROP); });

  // 2) Rename a header on every tab
  ss.getSheets().forEach(function (sh) { renameHeader_(sh, CONFIG.RENAME_FROM, CONFIG.RENAME_TO); });

  // 3) Build the new tab (once)
  if (ss.getSheetByName(CONFIG.NEW_TAB)) {
    ss.toast('Columns updated. "' + CONFIG.NEW_TAB + '" already exists — skipped.', "Done", 6);
    return;
  }

  var template = ss.getSheetByName(CONFIG.TEMPLATE_TAB);
  if (!template) { throw new Error("Template tab not found: " + CONFIG.TEMPLATE_TAB); }

  var sh = template.copyTo(ss).setName(CONFIG.NEW_TAB);

  // Clear content below the header, keep formatting
  var dataRows = sh.getMaxRows() - 1;
  if (dataRows > 0) { sh.getRange(2, 1, dataRows, sh.getMaxColumns()).clearContent(); }

  // Insert the extra column after the anchor header
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var anchorCol = headers.indexOf(CONFIG.INSERT_AFTER_HEADER) + 1;
  if (anchorCol < 1) { throw new Error('"' + CONFIG.INSERT_AFTER_HEADER + '" header not found on template.'); }

  sh.insertColumnAfter(anchorCol);
  var newCol = anchorCol + 1;
  sh.getRange(1, anchorCol).copyTo(sh.getRange(1, newCol), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  sh.getRange(1, newCol).setValue(CONFIG.NEW_COLUMN_HEADER);
  sh.setColumnWidth(newCol, sh.getColumnWidth(anchorCol));

  if (dataRows > 0) {
    sh.getRange(2, anchorCol, dataRows, 1).setDataValidation(listRule_(CONFIG.NEW_TAB_LEVELS));
    sh.getRange(2, newCol, dataRows, 1).setDataValidation(listRule_(CONFIG.NEW_COLUMN_OPTIONS));
  }

  // Position the new tab
  var after = ss.getSheetByName(CONFIG.PLACE_AFTER);
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(after ? after.getIndex() + 1 : ss.getNumSheets());

  ss.toast('Columns updated and "' + CONFIG.NEW_TAB + '" created.', "Done", 6);
}

/** Deletes every column whose header matches one of `names` (right-to-left so indexes don't shift). */
function removeColumnsByHeader_(sh, names) {
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) { return; }
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  headers
    .map(function (h, i) { return names.indexOf(h) > -1 ? i + 1 : 0; })
    .filter(function (c) { return c > 0; })
    .sort(function (a, b) { return b - a; })
    .forEach(function (c) { sh.deleteColumn(c); });
}

/** Renames the header cell that matches `from` (case-insensitive). */
function renameHeader_(sh, from, to) {
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) { return; }
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function (h) { return String(h).trim().toLowerCase(); });
  var idx = headers.indexOf(String(from).toLowerCase());
  if (idx > -1) { sh.getRange(1, idx + 1).setValue(to); }
}

function listRule_(values) {
  return SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build();
}

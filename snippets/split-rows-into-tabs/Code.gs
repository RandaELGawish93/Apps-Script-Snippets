/**
 * Split Rows Into Tabs
 * ------------------------------------------------------------
 * Reads a list from a "Source" sheet and writes one tab per
 * category, e.g. a training catalog grouped by scope, a book list
 * grouped by genre, or a task list grouped by owner.
 *
 * Source sheet layout (row 1 = headers):
 *   | Title | Category | Link |
 *
 * Re-running is safe: existing category tabs are cleared and rebuilt.
 */

var SOURCE_SHEET = "Source";
var TITLE_COL = 1;     // column A
var CATEGORY_COL = 2;  // column B
var LINK_COL = 3;      // column C

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Split Tabs")
    .addItem("Create sample Source sheet", "createSampleSource")
    .addItem("Split rows into tabs", "splitRowsIntoTabs")
    .addToUi();
}

function splitRowsIntoTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var source = ss.getSheetByName(SOURCE_SHEET);
  if (!source) {
    throw new Error('No "' + SOURCE_SHEET + '" sheet found. Run "Create sample Source sheet" first.');
  }

  var values = source.getDataRange().getValues().slice(1); // drop header row

  // Group rows by category, preserving first-seen order
  var order = [];
  var groups = {};
  values.forEach(function (row) {
    var title = String(row[TITLE_COL - 1] || "").trim();
    var category = String(row[CATEGORY_COL - 1] || "").trim();
    if (!title || !category) { return; }
    if (!groups[category]) {
      groups[category] = [];
      order.push(category);
    }
    groups[category].push([title, row[LINK_COL - 1] || ""]);
  });

  var summary = [];
  order.forEach(function (category) {
    var tabName = sanitizeTabName_(category);
    var sheet = ss.getSheetByName(tabName) || ss.insertSheet(tabName);
    sheet.clearContents();

    sheet.getRange(1, 1, 1, 2).setValues([["Title", "Link"]]).setFontWeight("bold");
    var rows = groups[category];
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 2);

    summary.push(tabName + " (" + rows.length + ")");
  });

  SpreadsheetApp.getUi().alert(
    "Done! Created/updated " + order.length + " tabs:\n\n" + summary.join("\n")
  );
}

function createSampleSource() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SOURCE_SHEET) || ss.insertSheet(SOURCE_SHEET);
  sheet.clearContents();
  sheet.getRange(1, 1, 7, 3).setValues([
    ["Title", "Category", "Link"],
    ["Getting Started with Formative Assessment", "Assessment", "https://example.com/course-1"],
    ["Designing Rubrics", "Assessment", "https://example.com/course-2"],
    ["Routines and Procedures", "Classroom Management", "https://example.com/course-3"],
    ["Positive Behavior Support", "Classroom Management", "https://example.com/course-4"],
    ["Using Google Forms for Quizzes", "Educational Technology", "https://example.com/course-5"],
    ["Cooperative Learning Structures", "Active Learning", "https://example.com/course-6"]
  ]);
  sheet.getRange(1, 1, 1, 3).setFontWeight("bold");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 3);
}

// Sheet tab names can't exceed 100 characters or contain: / \ ? * [ ]
function sanitizeTabName_(name) {
  return name.replace(/[\/\\\?\*\[\]]/g, "-").substring(0, 100);
}

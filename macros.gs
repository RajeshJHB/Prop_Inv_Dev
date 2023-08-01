//----------Macros------------------RS

function SortSTByDate() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C7:H7').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).sort({column: 3, ascending: true});
  spreadsheet.getRange('C7').activate();
};

function clearStatement() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C:H').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
};

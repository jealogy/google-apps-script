function isRedText(richTextValue, columnIndex) {
  if (columnIndex === 1) {
    return richTextValue.getTextStyle().getForegroundColor() === '#ff0000';
  }
  return false;
}

function LuzonSummary() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('NORTH LUZON - ALAND'), true);
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(spreadsheet.getCurrentCell().getRow() + 2, 1, 3441, sheet.getMaxColumns()).activate();
  spreadsheet.getActiveRange().createFilter();
  spreadsheet.getCurrentCell().offset(0, 14).activate();
  var criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['Bounced', 'Deposited-Manual', 'Deposited-RCD', 'For Warehousing', 'Rejected - Treasury', 'Rejected - UBP', 'Warehoused-RCD'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(spreadsheet.getActiveRange().getColumn(), criteria);
  spreadsheet.getCurrentCell().offset(0, -14, 1, 16).activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Summary'), true);
  spreadsheet.getCurrentCell().offset(4, 1).activate();
  spreadsheet.getRange('\'NORTH LUZON - ALAND\'!A3:P3443').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  var summarySheet = spreadsheet.getSheetByName('Summary');
  var lastRow = summarySheet.getLastRow();
  summarySheet.getRange(lastRow + 1, 2).activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('SOUTH LUZON - ALAND'), true);
  sheet = spreadsheet.getActiveSheet();
  sheet.getRange(spreadsheet.getCurrentCell().getRow() + 2, 1, 793, sheet.getMaxColumns()).activate();
  spreadsheet.getActiveRange().createFilter();
  spreadsheet.getCurrentCell().offset(0, 14).activate();
  criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['', 'Bounced', 'Deposited-Manual', 'Deposited-RCD', 'For Warehousing', 'On Hold', 'Rejected - UBP', 'Warehoused-RCD'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(spreadsheet.getActiveRange().getColumn(), criteria);
  spreadsheet.getCurrentCell().offset(59, -14, 1, 16).activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Summary'), true);
  spreadsheet.getRange('\'SOUTH LUZON - ALAND\'!A62:P502').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  var summarySheet = spreadsheet.getSheetByName('Summary');
  var lastRow = summarySheet.getLastRow();
  summarySheet.getRange(lastRow + 1, 2).activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('SOUTH LUZON - LIMA'), true);
  sheet = spreadsheet.getActiveSheet();
  sheet.getRange(spreadsheet.getCurrentCell().getRow() + 2, 1, 1222, sheet.getMaxColumns()).activate();
  spreadsheet.getActiveRange().createFilter();
  spreadsheet.getCurrentCell().offset(0, 14).activate();
  criteria = SpreadsheetApp.newFilterCriteria()
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(spreadsheet.getActiveRange().getColumn(), criteria);
  criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['', 'Bounced', 'Deposited-Manual', 'Deposited-RCD', 'For RCD Deposit', 'For Warehousing', 'Rejected - Treasury', 'Rejected - UBP', 'Warehoused-RCD'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(spreadsheet.getActiveRange().getColumn(), criteria);
  spreadsheet.getCurrentCell().offset(153, -14, 1, 16).activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Summary'), true);
  spreadsheet.getRange('\'SOUTH LUZON - LIMA\'!A156:P1037').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  var summarySheet = spreadsheet.getSheetByName('Summary');
  var lastRow = summarySheet.getLastRow();
  summarySheet.getRange(lastRow + 1, 2).activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('LUZON - ALAND'), true);
  sheet = spreadsheet.getActiveSheet();
  sheet.getRange(spreadsheet.getCurrentCell().getRow() + 2, 1, 3876, sheet.getMaxColumns()).activate();
  spreadsheet.getActiveRange().createFilter();
  spreadsheet.getCurrentCell().offset(0, 14).activate();
  criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['Bounced', 'Deposited-Manual', 'Deposited-RCD', 'Deposited-Warehouse', 'For Manual Deposit', 'For Warehousing', 'HOLD', 'On Hold', 'Rejected - Treasury', 'Rejected - UBP', 'Warehoused-RCD', 'Warehoused-UBP'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(spreadsheet.getActiveRange().getColumn(), criteria);
  spreadsheet.getCurrentCell().offset(27, -14, 1, 16).activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Summary'), true);
  spreadsheet.getRange('\'LUZON - ALAND\'!A30:P3181').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  var summarySheet = spreadsheet.getSheetByName('Summary');
  var lastRow = summarySheet.getLastRow();
  summarySheet.getRange(lastRow + 1, 2).activate();
  spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveSheet().getMaxRows(), 1000);
  var lastRow = summarySheet.getLastRow();
  summarySheet.getRange(lastRow + 1, 2).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
  spreadsheet.getCurrentCell().offset(1, 0).activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('LUZON - LIMA/CIPDI'), true);
  sheet = spreadsheet.getActiveSheet();
  sheet.getRange(spreadsheet.getCurrentCell().getRow() + 1, 1, 1087, sheet.getMaxColumns()).activate();
  spreadsheet.getActiveRange().createFilter();
  spreadsheet.getCurrentCell().offset(0, 14).activate();
  criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['Bounced', 'Deposited-Manual', 'Deposited-RCD', 'Deposited-Warehouse', 'On Hold', 'Rejected - UBP', 'Warehoused-RCD', 'Warehoused-UBP'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(spreadsheet.getActiveRange().getColumn(), criteria);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('SOUTH LUZON - LIMA'), true);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('CEBU - ALAND'), true);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('LUZON - LIMA/CIPDI'), true);
  spreadsheet.getCurrentCell().offset(5, -14, 1, 16).activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Summary'), true);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Summary'), true);
  spreadsheet.getRange('\'LUZON - LIMA/CIPDI\'!A7:P1084').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  var summarySheet = spreadsheet.getSheetByName('Summary');
  var lastRow = summarySheet.getLastRow();
  summarySheet.getRange(lastRow + 1, 2).activate();
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('NORTH LUZON - ALAND'), true);
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(spreadsheet.getCurrentCell().getRow(), 1, 1, sheet.getMaxColumns()).activate();
  spreadsheet.getActiveSheet().getFilter().remove();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('SOUTH LUZON - ALAND'), true);
  spreadsheet.getActiveSheet().getFilter().remove();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('SOUTH LUZON - LIMA'), true);
  spreadsheet.getActiveSheet().getFilter().remove();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('LUZON - ALAND'), true);
  spreadsheet.getActiveSheet().getFilter().remove();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('LUZON - LIMA/CIPDI'), true);
  spreadsheet.getActiveSheet().getFilter().remove();
}

function filterAndCopyToBounced(sourceSheetName, sourceSpreadsheetId) {
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
  var targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = targetSpreadsheet.getSheetByName('Bounced');

  var rangeToFilter = sourceSheet.getDataRange();
  var richTextValues = rangeToFilter.getRichTextValues();
  var data = rangeToFilter.getValues();
  var filteredValues = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rowHasRedText = false;

    for (var j = 0; j < row.length; j++) {
      var richTextValue = richTextValues[i][j];
      if (isRedText(richTextValue, j + 1)) { 
        rowHasRedText = true;
        break; 
      }
    }

    if (rowHasRedText) {
      filteredValues.push(row);
    }
  }

  if (filteredValues.length > 0) {
    var targetRow = targetSheet.getLastRow() + 1; 
    var targetRange = targetSheet.getRange(targetRow, 1, filteredValues.length, filteredValues[0].length);
    targetRange.setValues(filteredValues);
  } else {
    Logger.log("No matching data found in " + sourceSheetName);
  }
}

function processBothTabs() {
  filterAndCopyToBounced('UBP_7088 LUZON', '1yOw4RBovjJWPGQI7Q5rv1efi5EV6imoh4yIDAYvJmmA');
  filterAndCopyToBounced('UBP_11087 LUZON', '1yOw4RBovjJWPGQI7Q5rv1efi5EV6imoh4yIDAYvJmmA');
  filterAndCopyToBounced('CIPDI_UBP_Disb/Depo', '1xtCwcAziii_7AFhWl487q9zvbhdbmN7Z3zq2TOezR1c');
}

processBothTabs();




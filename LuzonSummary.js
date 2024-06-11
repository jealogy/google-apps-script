function isRedText(richTextValue, columnIndex) {
  if (columnIndex === 1) {
    return richTextValue.getTextStyle().getForegroundColor() === '#ff0000';
  }
  return false;
}

function LuzonSummary() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet, criteria, currentCell;
  
  try {
    // Function to create filter and apply criteria
    function applyFilterAndCopy(sourceSheetName, filterRangeStartRow, filterHiddenValues, copyRange) {
      spreadsheet.setActiveSheet(spreadsheet.getSheetByName(sourceSheetName), true);
      sheet = spreadsheet.getActiveSheet();
      var range = sheet.getRange(filterRangeStartRow, 1, sheet.getLastRow() - filterRangeStartRow + 1, sheet.getMaxColumns());
      range.activate();
      spreadsheet.getActiveRange().createFilter();
      spreadsheet.getCurrentCell().offset(0, 14).activate();
      criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(filterHiddenValues).build();
      var column = spreadsheet.getActiveRange().getColumn();
      if (column >= 1) {
        spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(column, criteria);
      }
      spreadsheet.getCurrentCell().offset(0, -14, 1, 16).activate();
      currentCell = spreadsheet.getCurrentCell();
      spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
      currentCell.activateAsCurrentCell();
      spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Pulled Out - Luzon'), true);
      spreadsheet.getRange(copyRange).copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      var summarySheet = spreadsheet.getSheetByName('Pulled Out - Luzon');
      var lastRow = summarySheet.getLastRow();
      summarySheet.getRange(lastRow + 1, 2).activate();
    }

    // NORTH LUZON - ALAND
    applyFilterAndCopy('NORTH LUZON - ALAND', 3, ['Bounced', 'Deposited-Manual', 'Deposited-RCD', 'For Warehousing', 'Rejected - Treasury', 'Rejected - UBP', 'Warehoused-RCD'], '\'NORTH LUZON - ALAND\'!A3:P3443');

    // SOUTH LUZON - ALAND
    applyFilterAndCopy('SOUTH LUZON - ALAND', 3, ['Bounced', 'Deposited-Manual', 'Deposited-RCD', 'For Warehousing', 'On Hold', 'Rejected - UBP', 'Warehoused-RCD'], '\'SOUTH LUZON - ALAND\'!A62:P502');

    // SOUTH LUZON - LIMA
    applyFilterAndCopy('SOUTH LUZON - LIMA', 3, ['Bounced', 'Deposited-Manual', 'Deposited-RCD', 'For RCD Deposit', 'For Warehousing', 'Rejected - Treasury', 'Rejected - UBP', 'Warehoused-RCD'], '\'SOUTH LUZON - LIMA\'!A156:P1037');

    // LUZON - ALAND
    applyFilterAndCopy('LUZON - ALAND', 3, ['Bounced', 'Deposited-Manual', 'Deposited-RCD', 'Deposited-Warehouse', 'For Manual Deposit', 'For Warehousing', 'HOLD', 'On Hold', 'Rejected - Treasury', 'Rejected - UBP', 'Warehoused-RCD', 'Warehoused-UBP'], '\'LUZON - ALAND\'!A30:P3181');

    // LUZON - LIMA/CIPDI
    applyFilterAndCopy('LUZON - LIMA/CIPDI', 3, ['Bounced', 'Deposited-Manual', 'Deposited-RCD', 'Deposited-Warehouse', 'On Hold', 'Rejected - UBP', 'Warehoused-RCD', 'Warehoused-UBP'], '\'LUZON - LIMA/CIPDI\'!A7:P1084');

    // Remove filters
    function removeFilters(sheetName) {
      spreadsheet.setActiveSheet(spreadsheet.getSheetByName(sheetName), true);
      sheet = spreadsheet.getActiveSheet();
      var filter = sheet.getFilter();
      if (filter) {
        filter.remove();
      }
    }

    removeFilters('NORTH LUZON - ALAND');
    removeFilters('SOUTH LUZON - ALAND');
    removeFilters('SOUTH LUZON - LIMA');
    removeFilters('LUZON - ALAND');
    removeFilters('LUZON - LIMA/CIPDI');
  } catch (e) {
    Logger.log('Error: ' + e.message);
  }
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

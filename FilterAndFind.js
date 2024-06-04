function FilterAndFind() {
  var sourceSpreadsheetId = "1yOw4RBovjJWPGQI7Q5rv1efi5EV6imoh4yIDAYvJmmA";
  var targetSpreadsheetId = "1SXs6PMrFO1Epr3HVuEpzdpw57ufuzFw9KNL_aWOsqok";
  
  var sourceSheetName = "UBP_7088 LUZON";
  var targetSheetNames = ["NORTH LUZON - ALAND v2", "SOUTH LUZON - ALAND", "SOUTH LUZON - LIMA", "LUZON - ALAND", "LUZON - LIMA/CIPDI"];
  var logSheetName = "logSheet"; 

  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
  
  var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
  var logSheet = targetSpreadsheet.getSheetByName(logSheetName);

  logSheet.clear();
  
  var lastProcessedRowRange = logSheet.getRange("A1");
  var lastProcessedRow = lastProcessedRowRange.getValue();
  if (!lastProcessedRow) {
    lastProcessedRow = 0;
  }

  var sourceRange = sourceSheet.getDataRange();
  var sourceValues = sourceRange.getValues();
  var richTextValues = sourceRange.getRichTextValues();

  var newLogEntries = [];
  var copiedValuesToFind = [];

  for (var i = lastProcessedRow; i < richTextValues.length; i++) {
    if (isRedText(richTextValues[i][0], 1)) {
      var copiedValue = sourceValues[i][31];
      if (copiedValue !== "" && copiedValue !== "0000000000" && !copiedValuesToFind.includes(copiedValue)) {
        copiedValuesToFind.push(copiedValue);
        newLogEntries.push([copiedValue]);
      }
    }
  }

  targetSheetNames.forEach(function(sheetName) {
    var targetSheet = targetSpreadsheet.getSheetByName(sheetName);
    if (!targetSheet) {
      Logger.log("Target sheet not found: " + sheetName);
      return;
    }

    var targetColumnRange = targetSheet.getRange("AY:AY");
    var targetColumnValues = targetColumnRange.getValues();
    var targetStatuses = targetSheet.getRange("O:O").getValues();
    var processedMarkers = targetSheet.getRange("AZ:AZ").getValues();

    var updates = [];

    for (var rowIndex = 0; rowIndex < targetColumnValues.length; rowIndex++) {
      var cellValue = targetColumnValues[rowIndex][0];
      var processedMarker = processedMarkers[rowIndex][0];
      if (copiedValuesToFind.includes(cellValue) && processedMarker !== "Processed") {
        updates.push(rowIndex + 1);
      }
    }

    if (updates.length > 0) {
      var bgRanges = targetSheet.getRangeList(updates.map(row => "AY" + row));
      var statusRanges = targetSheet.getRangeList(updates.map(row => "O" + row));
      var markerRanges = targetSheet.getRangeList(updates.map(row => "AZ" + row));

      bgRanges.setBackground('#FFFF00');
      statusRanges.setValue("Bounced");
      markerRanges.setValue("Processed");
    }
  });

  if (newLogEntries.length > 0) {
    var logSheetLastRow = logSheet.getLastRow();
    logSheet.getRange(logSheetLastRow + 1, 1, newLogEntries.length, 1).setValues(newLogEntries);

    lastProcessedRowRange.setValue(richTextValues.length);
  }

  Logger.log("Processed rows count: " + newLogEntries.length);
}

function isRedText(richTextValue, columnIndex) {
  if (columnIndex === 1) {
    return richTextValue.getTextStyle().getForegroundColor() === '#ff0000';
  }
  return false;
}

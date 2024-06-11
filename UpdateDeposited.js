function updateDeposited() {
   var sourceSpreadsheetId = "1yOw4RBovjJWPGQI7Q5rv1efi5EV6imoh4yIDAYvJmmA";
  var targetSpreadsheetId = "1SXs6PMrFO1Epr3HVuEpzdpw57ufuzFw9KNL_aWOsqok";
  
  var sourceSheetName = "UBP_7088 LUZON";
  var targetSheetNames = ["NORTH LUZON - ALAND v2", "SOUTH LUZON - ALAND", "SOUTH LUZON - LIMA", "LUZON - ALAND"];

  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
  
  var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);

  var sourceRange = sourceSheet.getDataRange();
  var sourceValues = sourceRange.getValues();

  var copiedValuesToFind = [];

  // Collect values to process
  for (var i = 0; i < sourceValues.length; i++) {
    if (sourceValues[i][8] > 0) { // Column I is index 8
      var copiedValue = sourceValues[i][31]; 
      if (copiedValue !== "" && !copiedValuesToFind.includes(copiedValue)) {
        copiedValuesToFind.push(copiedValue);
      }
    }
  }

  // Process each target sheet
  targetSheetNames.forEach(function(sheetName) {
    var targetSheet = targetSpreadsheet.getSheetByName(sheetName);
    if (!targetSheet) {
      Logger.log("Target sheet not found: " + sheetName);
      return;
    }

    var targetColumnRange = targetSheet.getRange("AY:AY");
    var targetColumnValues = targetColumnRange.getValues();
    var targetStatuses = targetSheet.getRange("O:O").getValues();
    var processedMarkers = targetSheet.getRange("BA:BA").getValues();

    var updates = [];

    for (var rowIndex = 0; rowIndex < targetColumnValues.length; rowIndex++) {
      var cellValue = targetColumnValues[rowIndex][0];
      if (copiedValuesToFind.includes(cellValue) && targetStatuses[rowIndex][0] !== "Deposited-RCD") {
        updates.push(rowIndex + 1);
      }
    }

    // Perform batch updates
    if (updates.length > 0) {
      var statusRanges = targetSheet.getRangeList(updates.map(row => "O" + row));
      var markerRanges = targetSheet.getRangeList(updates.map(row => "BA" + row));

      statusRanges.setValue("Deposited-RCD");
      markerRanges.setValue("Processed-Deposited");
    }
  });

  Logger.log("Processed rows count: " + copiedValuesToFind.length);
}

/* Management Magic for Google Analytics
*    Auxiliary functions for dimension management
* Based in project (https://github.com/narcan/api-tools/tree/master/GA%20Management%20Magic) - Pedro Avila (pdro@google.com)
* Copyright ©2018 Gustavo Decarlo (gustavodecarlo@gmail.com)
***************************************************************************/


/**************************************************************************
* Adds a formatted sheet to the spreadsheet to faciliate data management.
*/
function formatDimensionSheet(createNew) {
  // Get common values
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var date = new Date();
  var sheetName = "Dimensions@"+ date.getTime();
  var NUMBER_OF_COLUMNS = 7;

  // Normalize/format the values of the parameters
  createNew = (createNew === undefined) ? false : createNew;

  // Insert a new sheet or warn the user that formatting will erase data on the current sheet
  try {
    if (createNew) {
      sheet = ss.insertSheet(sheetName, 0);
    } else if (!createNew) {
      // Show warning to user and ask to proceed
      var response = Browser.msgBox('WARNING: This will erase all data on the current sheet, Would you like to proceed?', Browser.Buttons.YES_NO);
      if (response == 'yes') { // delete/remove all data/formatting/validations on sheet
        SpreadsheetApp.getActiveRange().clearDataValidations();
        sheet.clear();
        sheet.setName(sheetName);
      } else if (response == 'no') {
        Browser.msgBox('Format cancelled.');
        return sheet;
      } else {
        console.log(response);
        return sheet;
      }
    }
  } catch (error) {
    Browser.msgBox(error.message);
    return sheet;
  }

  // set local vars
  var numRows = sheet.getMaxRows();
  var numCols = sheet.getMaxColumns();
  var deltaCols = numCols - NUMBER_OF_COLUMNS;

  // set the number of columns
  try {
    if (deltaCols > 0) {
      sheet.deleteColumns(NUMBER_OF_COLUMNS, deltaCols);
    } else if (deltaCols < 0) {
      sheet.insertColumnsAfter(numCols, -deltaCols);
    }
  } catch (e) {
    return "failed to set the number of columns\n"+ e.message;
  }

  // Set ranges.
  var includeCol = sheet.getRange("A2:A");
  //var propertyCol = sheet.getRange("B2:B");
  var indexCol = sheet.getRange("E2:E");
  var scopeCol = sheet.getRange("F2:F");
  var activeCol = sheet.getRange("G2:G");

  // Set header range, values and formatting.
  try {
    var headerRange = sheet.getRange(1,1,1,sheet.getMaxColumns());
    sheet.getRange("A1").setValue("Include");
    sheet.getRange("B1").setValue("Account");
    sheet.getRange("C1").setValue("Property");
    sheet.getRange("D1").setValue("Name");
    sheet.getRange("E1").setValue("Index");
    sheet.getRange("F1").setValue("Scope");
    sheet.getRange("G1").setValue("Active");
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#4285F4");
    headerRange.setFontColor("#FFFFFF");

    // Index Column: protect & set background & font color
    indexCol.protect().setDescription("prevent others from modifying the dimension indices");
    indexCol.setBackground("#BABABA");
    indexCol.setFontColor("#FFFFFF");

    // Include Column: modify data validation values
    var includeValues = ['✓'];
    var includeRule = SpreadsheetApp.newDataValidation().requireValueInList(includeValues, true).build();
    includeCol.setDataValidation(includeRule);

    // Scope Column: modify data validation values
    var scopeValues = ['USER','SESSION','HIT','PRODUCT'];
    var scopeRule = SpreadsheetApp.newDataValidation().requireValueInList(scopeValues, true).build();
    scopeCol.setDataValidation(scopeRule);

    // Active Column: modify data validation values
    var activeValues = ['TRUE','FALSE'];
    var activeRule = SpreadsheetApp.newDataValidation().requireValueInList(activeValues, true).build();
    activeCol.setDataValidation(activeRule);
  } catch (e) {
    return "failed to set the header values and format ranges\n"+ e.message;
  }

  // send Measurement Protocol hit to Google Analytics
  /*var label = '';
  var value = '';
  var httpResponse = mpHit(SpreadsheetApp.getActiveSpreadsheet().getUrl(),'format dimension sheet',label,value);
  Logger.log(httpResponse);*/

  return sheet;
}

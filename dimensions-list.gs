/* Management Magic for Google Analytics
*    Lists dimensions from a GA property
* Based in project (https://github.com/narcan/api-tools/tree/master/GA%20Management%20Magic) - Pedro Avila (pdro@google.com)
* Copyright ©2018 Gustavo Decarlo (gustavodecarlo@gmail.com)
***************************************************************************/


/**************************************************************************
* Obtains input from user necessary for listing dimensions.
*/

function requestDimensionData() {
  var html = HtmlService.createHtmlOutputFromFile('dimension');
  SpreadsheetApp.getUi().showModalDialog(html, 'Dimension Get Data');
}

function requestDimensionList(formObject) {
  // Get account and property id from formatMetricSheet
  var accountId = formObject.gaAccount,
      propertyId = formObject.gaProp;

  // Process the user's response.
  if (accountId && propertyId) {
    // Construct the array of one or more properties from the user's input.
    var propertyList = propertyId;
    var propertyListArray = propertyList.split(/\s*,\s*/);

    // List dimensions from all properties entered by the user.
    var listResponse = listDimensions(accountId,propertyListArray);

    // Output errors and log successes.
    if (listResponse != "success") {
      Browser.msgBox(listResponse);
    } else {
      console.log("List dimensions response: "+ listResponse)
    }
  }
  // Log method by which the user chose not to proceed.
  else {
    console.log('The user did not provide a Account and property ID.');
  }
}

/**************************************************************************
* Lists dimension settings from the property into a new sheet
* @param {string} property The tracking ID of the GA property
* @return {string} Operation output ('success' or error message)
*/
function listDimensions(accountId,propertyList) {
  // Set common values
  var include = "✓";
  var allDimensions = [];
  var dataColumns = 7;

  // Iterate through the array of properties from which to list dimensions
  for (p = 0; p < propertyList.length; p++) {
    var property = propertyList[p];

    // Process a property id if it matches a valid format.
    if (property.match(/UA-\d+-\d+/)) {


      // Attempt to get property information from the Management API
      try {
        var dimensionList = Analytics.Management.CustomDimensions.list(accountId, property);
      } catch (e) {
        return e.message;
      }

      // Attempt to store the information received from the Management API in an array
      try {
        var dimensions = [];

        // Parse each result of the API request and push it to an array
        for (var i = 0; i < dimensionList.totalResults; i++) {
          var dimension = dimensionList.items[i];
          dimensions[i] = [include,accountId,dimension.webPropertyId,dimension.name,dimension.index,dimension.scope,dimension.active];
          allDimensions.push(dimensions[i]);
        }
      } catch (e) {
        return e.message;
      }
    }
    // Return an error message if the property id does not match the correct format.
    else return property +" is an invalid property format";
  }

  // Insert the values processed from the API into a formatted sheet
  try {
    // Set the values in the sheet
    var sheet = formatDimensionSheet(true);
    sheet.getRange(2,1,allDimensions.length,dataColumns).setValues(allDimensions);
  } catch (e) {
    return e.message;
  }

  // send Measurement Protocol event hit to Google Analytics
  /*var label = propertyList;
  var value = propertyList.length;
  var httpResponse = mpHit(SpreadsheetApp.getActiveSpreadsheet().getUrl(),'list dimensions',label,value);
  console.log(httpResponse);*/

  return "success";
}

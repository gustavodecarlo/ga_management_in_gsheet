/* Management Magic for Google Analytics
*    Lists metrics from a GA property
* Based in project (https://github.com/narcan/api-tools/tree/master/GA%20Management%20Magic) - Pedro Avila (pdro@google.com)
* Copyright ©2018 Gustavo Decarlo (gustavodecarlo@gmail.com)
***************************************************************************/


/**************************************************************************
* Obtains input from user necessary for listing metrics.
*/

function requestMetricData() {
  var html = HtmlService.createHtmlOutputFromFile('metric');
  SpreadsheetApp.getUi().showModalDialog(html, 'Metric Get Data');
}

function requestMetricList(formObject) {
  // Get account and property id from formatMetricSheet
  var accountId = formObject.gaAccount,
      propertyId = formObject.gaProp;

  // Process the user's response.
  if (accountId && propertyId) {
    // Construct the array of one or more properties from the user's input.
    var propertyList = propertyId;
    var propertyListArray = propertyList.split(/\s*,\s*/);

    // List metrics from all properties entered by the user.
    var listResponse = listMetrics(accountId,propertyListArray);

    // Output errors and log successes.
    if (listResponse != "success") {
      Browser.msgBox(listResponse);
    } else {
      console.log("List metrics response: "+ listResponse)
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
function listMetrics(accountId,propertyList) {
  // Set common values
  var include = "✓";
  var allMetrics = [];
  var dataColumns = 10;

  // Iterate through the array of properties from which to list metrics
  for (p = 0; p < propertyList.length; p++) {
    var property = propertyList[p];

    // Process a property id if it matches a valid format.
    if (property.match(/UA-\d+-\d+/)) {

      // Attempt to get property information from the Management API
      try {
        var metricList = Analytics.Management.CustomMetrics.list(accountId, property);
      } catch (e) {
        return e.message;
      }

      // Attempt to store the information received from the Management API in an array
      try {
        var metrics = [];

        // Parse each result of the API request and push it to an array
        for (var i = 0; i < metricList.totalResults; i++) {
          var metric = metricList.items[i];
          metrics[i] = [include,accountId,metric.webPropertyId,metric.name,metric.index,metric.scope,metric.type,metric.min_value,metric.max_value,metric.active];
          allMetrics.push(metrics[i]);
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
    var sheet = formatMetricSheet(true);
    sheet.getRange(2,1,allMetrics.length,dataColumns).setValues(allMetrics);
  } catch (e) {
    return e.message;
  }

  // send Measurement Protocol event hit to Google Analytics
  /*var label = propertyList;
  var value = propertyList.length;
  var httpResponse = mpHit(SpreadsheetApp.getActiveSpreadsheet().getUrl(),'list metrics',label,value);
  Logger.log(httpResponse);*/

  return "success";
}

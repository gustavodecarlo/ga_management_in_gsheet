/* Management Magic for Google Analytics
*   Adds a menu item to manage Google Analytics Properties
* Based in project (https://github.com/narcan/api-tools/tree/master/GA%20Management%20Magic) - Pedro Avila (pdro@google.com)
* Copyright ©2018 Gustavo Decarlo (gustavodecarlo@gmail.com)
**************************************************************************/


/**************************************************************************
* Main function runs on application open, setting the menu of commands
*/
function onOpen(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // create the addon menu
  try {
    var menu = ui.createAddonMenu();
    if (e && e.authMode == ScriptApp.AuthMode.NONE) {
      // Add a normal menu item (works in all authorization modes).
      menu.addItem('List filters', 'requestFilterData')
      .addItem('Update filters', 'requestFilterUpdate')
      .addSeparator()
      .addItem('List custom dimensions', 'requestDimensionData')
      .addItem('Update custom dimensions', 'requestDimensionUpdate')
      .addSeparator()
      .addItem('List custom metrics', 'requestMetricList')
      .addItem('Update custom metrics', 'requestMetricUpdate')
      .addSeparator()
      .addItem('Format filter sheet', 'formatFilterSheet')
      .addItem('Format dimension sheet', 'formatDimensionSheet')
      .addItem('Format metric sheet', 'formatMetricSheet')
      .addSeparator()
      .addItem('About this Add-on','about');
    } else {
      menu.addItem('List filters', 'requestFilterData')
      .addItem('Update filters', 'requestFilterUpdate')
      .addSeparator()
      .addItem('List custom dimensions', 'requestDimensionData')
      .addItem('Update custom dimensions', 'requestDimensionUpdate')
      .addSeparator()
      .addItem('List custom metrics', 'requestMetricList')
      .addItem('Update custom metrics', 'requestMetricUpdate')
      .addSeparator()
      .addItem('Format filter sheet', 'formatFilterSheet')
      .addItem('Format dimension sheet', 'formatDimensionSheet')
      .addItem('Format metric sheet', 'formatMetricSheet')
      .addSeparator()
      .addItem('About this Add-on','about');
    }
    menu.addToUi();

    // send Measurement Protocol hitType to Google Analytics
    //mpHit(ss.getUrl(),'open');
  } catch (e) {
    Browser.msgBox(e.message);
  }
}

/**************************************************************************
* Edit function runs when the application is edited
*/
function onEdit(e) {
  // send Measurement Protocol hitType to Google Analytics
  //mpHit(ss.getUrl(),'edit');
}

/**************************************************************************
* Install function runs when the application is installed
*/
function onInstall(e) {
  onOpen(e);
  // send Measurement Protocol hitType to Google Analytics
  //mpHit(ss.getUrl(),'install');
}

/**
* Shows the side bar populated with the content from the instructions page
*/
function about() {
  var html = HtmlService.createHtmlOutputFromFile('about')
  .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  .setTitle('About')
  .setWidth(300);

  SpreadsheetApp.getUi().showSidebar(html);
}

/**************************************************************************
* Example function for Google Analytics Measurement Protocol.
* @param {string} tid Tracking ID / Web Property ID
* @param {string} url Document location URL
* @return {string} HTTP response
*/
function mpHit(url, intent, label, value){
  if (intent == 'open'|| intent == 'edit' || intent == 'install' || intent == '') {
    var hitType = 'pageview';
    var category, action, label, value = '';
  } else {
    var hitType = 'event';
    var category = 'interaction';
    var action = intent;
  }

  var v = '1';
  var tid = 'UA-XXXXXX-YY'; // Put your GA account if you want to track
  var cid = createUUID();
  var z = Math.floor(Math.random()*10E7);
  var ds = 'web';
  var dr = encodeURIComponent('GA Management Magic Addon');
  var dl = encodeURIComponent(url);
  var el = encodeURIComponent(label);

  var hit = "https://www.google-analytics.com/collect?"
  +"v="+ v
  +"&tid="+ tid
  +"&cid="+ cid
  +"&z="+ z
  +"&t="+ hitType
  +"&ds="+ ds
  +"&dr="+ dl
  +"&dl="+ dl;
  if (hitType == 'event') hit = hit +"&ec="+ category +"&ea="+ action +"&el="+ el +"&ev="+ value;

  return UrlFetchApp.fetch(hit).getHeaders();
}

function createUUID() {
  // http://www.ietf.org/rfc/rfc4122.txt
  var s = [];
  var hexDigits = "0123456789abcdef";
  for (var i = 0; i < 36; i++) {
    s[i] = hexDigits.substr(Math.floor(Math.random() * 0x10), 1);
  }
  s[14] = "4";  // bits 12-15 of the time_hi_and_version field to 0010
  s[19] = hexDigits.substr((s[19] & 0x3) | 0x8, 1);  // bits 6-7 of the clock_seq_hi_and_reserved to 01
  s[8] = s[13] = s[18] = s[23] = "-";

  var uuid = s.join("");
  return uuid;
}

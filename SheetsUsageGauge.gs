/**
 * @OnlyCurrentDoc
 *
 * Only request permissions for the currently running sheet
 * {@link https://developers.google.com/apps-script/guides/services/authorization#permissions_and_types_of_scripts}
 */

/**
 * Creates a menu entry in the Google Sheets UI when the spreadsheet is opened.
 * {@link https://developers.google.com/apps-script/guides/triggers}
 */
function onOpen() {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Show gauge', 'showGaugeSidebar')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 * {@link https://developers.google.com/apps-script/guides/triggers}
 */
function onInstall() {
  onOpen();
}

/**
 * Gets the usage stats.
 *
 * @return {Array} size stats for each sheet included in this spreadsheet
 */
function getUsageStats() {
  
  return SpreadsheetApp.getActive().getSheets().map(function(sheet) {
    return {
      name: sheet.getName(),
      maxRows: sheet.getMaxRows(),
      maxColumns: sheet.getMaxColumns(),
    }
  });
}


/**
 * Display the sidebar containing the gauge (and the JavaScript to run it)
 */
function showGaugeSidebar() {
   
  var htmlOutput = HtmlService.createHtmlOutputFromFile('GaugeSidebar.html')

  htmlOutput.setTitle('Cells Used');  

  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}  


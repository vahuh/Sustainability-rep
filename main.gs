/* 
  Sustainability reporting plugin 
  
  
 */


function tag(param) {
  let document = DocumentApp.getActiveDocument()
  let selection = document.getSelection()
  if (selection) {
    Drive.Comments.insert({ anchor: selection, content: param }, document.getId())
    console.log("range", selection)
  }

  console.log("func tag called with", param)
}



/* Function to create Sheets in current folder */
function createSheetOnCurrentFolder() {
  /* Get current folder */
  let document = DocumentApp.getActiveDocument()
  let file = DriveApp.getFileById(document.getId())
  let folder = file.getParents().next()
  /* Create SpreadSheet */
  let sheet = SpreadsheetApp.create("SuSaf output")
  let sheetfile = DriveApp.getFileById(sheet.getId())
  /* Move sheet to current folder */
  if (folder) sheetfile.moveTo(folder)
}

function toCsv() {
  /* Get current document folder */
  let document = DocumentApp.getActiveDocument()
  let file = DriveApp.getFileById(document.getId())
  let folder = file.getParents().next()
  /* Get first sheet */
  let sheets = folder.getFilesByType(MimeType.GOOGLE_SHEETS)
  let sheet = sheets.next()
  if (sheet) {
    /* Open spreadsheet */
    let ss = SpreadsheetApp.openById(sheet.getId())
    /* Get all data */
    let range = ss.getDataRange()
    /* Returns a list of lists */
    let values = range.getValues()
    console.log("values", values)
  }
}

/* Function that creates a new Spreadsheet */
function generateSheet() {
  var newSheet = SpreadsheetApp.create("Tag List");
  writeOnSheet(newSheet);
  //lets the user know that a spreadesheet was created
  DocumentApp.getUi().alert('Tag List spreadsheet was created')
}

// Function that allows writing on a created Spreadsheet
function writeOnSheet(sSheet) {
  console.log("write to sheet", sSheet.getId())
  let values = [
    [
      1, 2, 3// Cell values ...
    ],
    ["Tekstiä", "testi", "123"]
    // Additional rows ...
  ];

  range = sSheet.getRange("A1:C2")
  range.setValues(values)

  console.log("result = ", result)
}
/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
    .addItem('Start', 'showSidebar')
    .addToUi();
}

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}


/**
 * Opens a sidebar in the document containing the add-on's user interface.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 */
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Sustainability Reporting');
  DocumentApp.getUi().showSidebar(ui);

}

//Function to show the Pop
function showPopup(){
  var html = HtmlService.createHtmlOutputFromFile('popup')
  .setWidth(500)
  .setHeight(600)
  DocumentApp.getUi().showModalDialog(html, 'Feature tagging')
}

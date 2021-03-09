/* 
  Sustainability reporting plugin 
  
  
 */

function tag(param) {
  let document = DocumentApp.getActiveDocument()
  let selection = document.getSelection()
  if (selection) {
    let selectedText = getTextSelection(selection)
    showPopup(selectedText, param)
    console.log("range", selection)
  }
  console.log("func tag called with", param)
}


//function used to get the text selected by the user  
function getTextSelection(selection) {
  var textAsString = ""
  var selectedText = ""
  var selectedElements = selection.getSelectedElements()
  for (var i = 0; i < selectedElements.length; i++) {
    var currentElement = selectedElements[i]
    //beginning of the user selection
    selectedText = currentElement.getElement().asText().getText()
    if (currentElement.isPartial()) {
      var startIndex = currentElement.getStartOffset()
      //end of the user selection
      var endIndex = currentElement.getEndOffsetInclusive() + 1
      //getting the selected text as a string
      selectedText = selectedText.substring(startIndex, endIndex)
    }
    textAsString += " " + selectedText.trim()
  }
  return textAsString
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

//Function to show the Pop Up 
function showPopup(feature, dimension) {
  var htmlPopup = HtmlService.createHtmlOutputFromFile('popup')
    .setWidth(600)
    .setHeight(700)
  //variable to append the selected feature as hidden division in the html file
  var strFeature = "<div id='selectedFeature' style='display:none;'>" + Utilities.base64Encode(JSON.stringify(feature)) + "</div>"
  //appending the Feature division to html popup file
  htmlPopup = htmlPopup.append(strFeature)
  //variables to append the sustainability dimensions as hidden division in the html file
  var strDimension = "<div id='susDimension' style='display:none;'>" + Utilities.base64Encode(JSON.stringify(dimension)) + "</div>"
  //appending the sustainability dimension to html file  
  htmlPopup = htmlPopup.append(strDimension)
  //getting the form visible to the user
  DocumentApp.getUi().showModalDialog(htmlPopup, 'Feature tagging')

}


/** 
 * This doesn't work properly, needs to be fixed, the spreadsheet url needs also to be changed once we have a working code* */
function processFeatures(formObject) {
  /* var spreadsheetURL="https://docs.google.com/spreadsheets/d/152NZwO02mdxcgO1nKpyCYGHWyD07vGEI9JZestKjwsI/edit#gid=0"
  var usedSpreadSheet = SpreadsheetApp.openByUrl(spreadsheetURL)
  var currentSheet = usedSpreadSheet.getSheetByName("Sheet1")
 */
  let document = DocumentApp.getActiveDocument()
  let file = DriveApp.getFileById(document.getId())
  let folder = file.getParents().next()
  /* Get first sheet */
  let sheets = folder.getFilesByType(MimeType.GOOGLE_SHEETS)
  if (!sheets.hasNext()) {
    createSheetOnCurrentFolder()
    sheets = folder.getFilesByType(MimeType.GOOGLE_SHEETS)
  }
  let sheet = sheets.next()
  if (sheet){
    /* Open spreadsheet */
    let currentSheet = SpreadsheetApp.openById(sheet.getId())
    console.log("formObject", formObject.inputCategory,"subcat", formObject.inputSubCategory, formObject.topicSelection, formObject)
    currentSheet.appendRow([formObject.inputCategory, formObject.inputSubCategory, formObject.topicSelection])
  }
}

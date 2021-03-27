/* 
  Sustainability reporting plugin 
  
  
 */

function tag(dimension) {
  let document = DocumentApp.getActiveDocument()
  let selection = document.getSelection()
  if (selection) {
    let selectedText = getTextSelection(selection)
    highlightSelectedText(selection, dimension)
    showPopup(selectedText, dimension)
    console.log("range", selection)
  }
  else {
    DocumentApp.getUi().alert("You need to select text to tag!")
  }
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

//Function used to highlight the selected text in the color of the associated dimension
function highlightSelectedText(selection, dimension) {
  //an element is a paragraph 
  let color;
  switch (dimension) {
    case "economic":
      color = "#9370db"
      break
    case "environmental":
      color = "#2e8b57"
      break
    case "social":
      color = "#DDA0DD"
      break
    case "technical":
      color = "#708090"
      break
    case "individual":
      color = "#bc8f8f"
      break
    default:
      color = '#FFFF00'
  }

  var selectedElements = selection.getRangeElements()
  for (var i = 0; i < selectedElements.length; i++) {
    var currentElement = selectedElements[i]
    //Checking that the selection is text and not images
    if (currentElement.getElement().editAsText) {
      var text = currentElement.getElement().editAsText()
      //checking if the current element is complete or not
      if (currentElement.isPartial()) {
        var startIndex = currentElement.getStartOffset()
        var endIndex = currentElement.getEndOffsetInclusive()
        //Highlighting the text by changing its background color based on the associated dimension, sets the color only to the selected text
        text.setBackgroundColor(startIndex, endIndex, color)
      } else {
        text.setBackgroundColor(color)
      }
    }
  }
}

/* Function to create Sheets in current folder */
function createSheetOnCurrentFolder() {
  /* Get current folder */
  let document = DocumentApp.getActiveDocument()
  let file = DriveApp.getFileById(document.getId())
  let folder = file.getParents().next()
  /* Create SpreadSheet */
  let sheet = SpreadsheetApp.create("SuSaf output")
  sheet.appendRow(['Feature', 'Dimension', 'Category', 'SubCategory', 'Topic', 'Impact', 'Order of effect', 'Memo'])

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
    .addItem('Topic list', 'askTopics')
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

/**
*  Display a dialog box with a title, message, input field, and "Yes" and "No" buttons.
*  The user can also close the dialog by clicking the close button in its title bar.
*/
function askTopics() {
  // Get previous properties
  let documentProperties = PropertiesService.getDocumentProperties();
  let topics = documentProperties.getProperty('TOPICS');
  let ui = DocumentApp.getUi();
  let response = ui.prompt('List of topics', topics + '\nSeparate topics with comma', ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response.getSelectedButton() == ui.Button.YES) {
    // Got response, save as new topic list
    let text = response.getResponseText();
    documentProperties.setProperty('TOPICS', text);
  } else if (response.getSelectedButton() == ui.Button.NO) {
    // the user clicked no
  } else {
    // the user closed popup
  }
}

// Function to show the Pop Up 
function showPopup(feature, dimension) {
  // Get document topics
  let documentProperties = PropertiesService.getDocumentProperties();
  let topics = documentProperties.getProperty('TOPICS')
  let htmlTemplate = HtmlService.createTemplateFromFile('popup')

  if (topics) {
    // set topics on template
    htmlTemplate.data = topics.split(",");
  } else {
    // send empty array
    htmlTemplate.data = []
  }
  // evaluate template
  let htmlPopup = htmlTemplate.evaluate();
  // set width and height
  htmlPopup
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



async function processFeatures(formObject) {
  try {
    let document = DocumentApp.getActiveDocument()
    let file = DriveApp.getFileById(document.getId())
    let folder = file.getParents().next()
    /* Get first sheet */
    let sheets = folder.getFilesByType(MimeType.GOOGLE_SHEETS)
    if (!sheets.hasNext()) {
      await createSheetOnCurrentFolder()
      sheets = folder.getFilesByType(MimeType.GOOGLE_SHEETS)
    }
    let sheet = sheets.next()
    if (sheet) {
      /* Open spreadsheet */
      let currentSheet = SpreadsheetApp.openById(sheet.getId())
      console.log("formObject", formObject.inputCategory, "subcat", formObject.inputSubCategory, formObject.topicSelection, formObject)

      currentSheet.appendRow([formObject.selectedFeature, formObject.susDimension, formObject.inputCategory, formObject.inputSubCategory, formObject.topicSelection, formObject.impactPosNeg, formObject.orderEffect, formObject.memoArea])
    }
  } catch (e) {
    DocumentApp.getUi().alert("You don't have permission to write to parent folder. Please contact project owner.");
  }
}

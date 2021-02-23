 /* 
  Sustainability reporting plugin 
 */

//Function that adds a comment to the active google doc file 
function tag(param, color) {
  let document = DocumentApp.getActiveDocument()
  let selection = document.getSelection()
//Comment is added only if text is selected
  if (selection) {
    //content of the comment will be the tag and the selected feature
    var stringContent = "Tag: " + param + " Feature: " + getTextSelection(selection)
    Drive.Comments.insert({ content: stringContent }, document.getId())
    highlightSelectedText(selection,color)
    console.log("range", selection)
  }
  else{
    DocumentApp.getUi().alert("No text selected")
  }
  console.log("func tag called with", param)
}

//Function used to highlight the selected text in the color of the associated dimension
function highlightSelectedText(selection,color){
  //an element is a paragraph 
  var selectedElements = selection.getRangeElements()
  for (var i = 0; i < selectedElements.length; i++){
    var currentElement = selectedElements[i]
    //Checking that the selection is text and not images
    if (currentElement.getElement().editAsText){  
      var text = currentElement.getElement().editAsText()
      //checking if the current element is complete or not
      if (currentElement.isPartial()){
        var startIndex = currentElement.getStartOffset()
        var endIndex = currentElement.getEndOffsetInclusive()
        //Highlighting the text by changing its background color based on the assoviated dimension, sets the color only to the selected text
        text.setBackgroundColor(startIndex, endIndex, color)
      }else{
        text.setBackgroundColor(color)
      }
    }
  }
}

//function used to get the text selected by the user  
function getTextSelection(selection){
  var textAsString = ""
  var selectedText= ""
  var selectedElements = selection.getSelectedElements()
  for (var i = 0; i < selectedElements.length;i++){
    var currentElement = selectedElements[i]
    //beginning of the user selection
    selectedText=currentElement.getElement().asText().getText()
    if (currentElement.isPartial()){
      var startIndex = currentElement.getStartOffset()
      //end of the user selection
      var endIndex = currentElement.getEndOffsetInclusive()+1
      //getting the selected text as a string
      selectedText = selectedText.substring(startIndex, endIndex)
    }
    textAsString += " " + selectedText.trim()
  }
  return textAsString

}


// Function that creates a new Spreadsheet 
function generateSheet() {
  var newSheet = SpreadsheetApp.create("Tag List")
  writeOnSheet(newSheet)
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

//function for getting comments from a file 
function getComments(){
  var docID = DocumentApp.getActiveDocument().getId()
  var comments = Drive.Comments.list(docID)

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

/*function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
    .addItem('Start', 'showSidebar')
    .addToUi();
}*/

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

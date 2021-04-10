/* 
  Sustainability reporting plugin 
 */


/**Function that tags a selected text with the selected dimension 
 * Takes the dimension from the button clicked by the user
 * If no text is selected, the user gets an alert to select text 
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

function categorize(categoryType){
  showCatPopup(categoryType)
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
  sheet.appendRow(['ID','Effect', 'Dimension','Category', 'SubCategory','Topic', 'Impact', 'Order of effect', 'Memo','Leads to'])

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
  let response = ui.prompt('List of topics', topics + '\n\nSeparate topics with comma.\n\nDo you want to overwrite existing topics?', ui.ButtonSet.YES_NO_CANCEL);

  // Process the user's response.
  if (response.getSelectedButton() == ui.Button.YES) {
    // Got response, save as new topic list
    let text = response.getResponseText();
    documentProperties.setProperty('TOPICS', text);
  } else if (response.getSelectedButton() == ui.Button.NO) {
    // the user clicked no
    let text = response.getResponseText();
    documentProperties.setProperty('TOPICS', topics + "," + text);
  } else {
    // the user closed popup
  }
}

/**
 * Function to set Category or SubCategory properties for document
 * It appends new categories or subcategories to document properties 
 */
function setCatProperties(catType, inputText){
  //Get Previous properties 
  let documentProperties = PropertiesService.getDocumentProperties();
  let categories = documentProperties.getProperty('CATEGORIES')
  let subCategories = documentProperties.getProperty('SUBCATEGORIES')
  if (catType == 'Category'){
    documentProperties.setProperty('CATEGORIES', categories + "," + inputText);
  }
  else if(catType == 'SubCategory'){
    documentProperties.setProperty('SUBCATEGORIES', subCategories + "," + inputText);
  }
}


// Function to show the Pop Up 
function showPopup(effect, dimension) {
  // Get document topics
  let documentProperties = PropertiesService.getDocumentProperties();
  let topics = documentProperties.getProperty('TOPICS')
  let htmlTemplate = HtmlService.createTemplateFromFile('popup')
  let ddOptionDict = populateDropdown()
  var ddValues = []
  
  if (!isEmpty(ddOptionDict)){
    for (var id in ddOptionDict){
    var contentArray = ddOptionDict[id]
    var currentEffect = contentArray[0]
    var currentDimension = contentArray[1]
    var ddOption = currentEffect + " Dimension: "+ currentDimension + " ("+id+") "
    ddValues.push(ddOption)
    }
  }
  //set feature dropdown options to template
  htmlTemplate.dropdownOptions = ddValues
  console.log("test",ddValues)

  if (topics) {
    // set topics on template
    htmlTemplate.data = topics.split(",")
  } else {
    // send empty array
    htmlTemplate.data = []
  }
  // evaluate template
  let htmlPopup = htmlTemplate.evaluate()
  // set width and height
  htmlPopup
    .setWidth(600)
    .setHeight(750)
  //variable to append the selected effect as hidden division in the html file
  var strEffect = "<div id='selectedEffect' style='display:none;'>" + Utilities.base64Encode(JSON.stringify(effect)) + "</div>"
  //appending the Effect division to html popup file
  htmlPopup = htmlPopup.append(strEffect)
  //variables to append the sustainability dimensions as hidden division in the html file
  var strDimension = "<div id='susDimension' style='display:none;'>" + Utilities.base64Encode(JSON.stringify(dimension)) + "</div>"
  //appending the sustainability dimension to html file  
  htmlPopup = htmlPopup.append(strDimension)
  //getting the form visible to the user
  DocumentApp.getUi().showModalDialog(htmlPopup, 'Effect tagging')

}

/**Function that shows the popup for categorizing an effect
 * Gets all the existing effect values and pushes them to a list that is visible to the user on the effect selection dropdown 
 * No values are added if there are no existing values 
*/
function showCatPopup(catType){
  let htmlTemplate = HtmlService.createTemplateFromFile('categories')

  let ddOptionDict = populateDropdown()
  var ddValues = []
  
  if (!isEmpty(ddOptionDict)){
    for (var id in ddOptionDict){
    var contentArray = ddOptionDict[id]
    var currentEffect = contentArray[0]
    var currentDimension = contentArray[1]
    var ddOption = currentEffect + " Dimension: "+ currentDimension + " ("+id+") "
    ddValues.push(ddOption)
    }
  }
  //set effect dropdown options to template
  htmlTemplate.dropdownOptions = ddValues
  console.log("test",ddValues)
  
  let htmlCategories = htmlTemplate.evaluate()
  htmlCategories
  .setWidth(600)
  .setHeight(500)

  var strCatType = "<div id='catType' style='display:none;'>" + Utilities.base64Encode(JSON.stringify(catType)) + "</div>"

  htmlCategories = htmlCategories.append(strCatType)
  DocumentApp.getUi().showModalDialog(htmlCategories,'Categorizing')
}

function processCategories(formObject, catType){
  var document = DocumentApp.getActiveDocument()
  var file = DriveApp.getFileById(document.getId())
  var folder = file.getParents().next()
  var spreadsheets = folder.getFilesByType(MimeType.GOOGLE_SHEETS)

  if (spreadsheets.hasNext()){
    var spreadsheet = spreadsheets.next()
    var currentSheet = SpreadsheetApp.openById(spreadsheet.getId())
    var lastrow = currentSheet.getLastRow()
    var idColumn = currentSheet.getRange("A2:A"+lastrow)
    var catColumn = currentSheet.getRange("D2:D"+lastrow)
    var subCatColumn = currentSheet.getRange("E2:E"+lastrow)

    try{
      categorizedEffect = formObject.effectDdl
      selectedID = findEffectID(categorizedEffect)
      console.log('selected ID:',selectedID)
    }catch (e){
      DocumentApp.getUi.alert("An error occured, ID couldn't be found.")
    }

    var idData = idColumn.getValues()
    var effectRow = null
    for(var i = 0; i<idData.length;i++){
      //if value is equal to ID, we return the value
      if (idData[i]==selectedID){
        effectRow = i
        break
      }
    }

//still needs to add categories to specific property in doc, in order to get the ddl working
    if (effectRow == null){
      DocumentApp.getUi.alert("No match found for ID")
    }else {
      rowNumber = effectRow + 2
      rowNumberStr = rowNumber.toString()
      if (catType == "Category"){
        var catCell = currentSheet.getRange("D"+rowNumberStr+":D"+rowNumberStr).getCell(1,1)
        catCell.setValue('this is test') 
      }
      else if(catType == "SubCategory"){
        var subCatCell = currentSheet.getRange("E"+rowNumberStr+":E"+rowNumberStr).getCell(1,1)
        subCatCell.setValue('this is test') 
      }
    
    }
    
  }
}


/**
 * Function to get ID associated with an effect from dropdown list 
 * Works only if only one set of '()' is present in the input string */ 
function findEffectID(textString){
  effectID = textString.split('(').pop().split(')')[0]
  return effectID

}

//function to check if a dictionary is empty
function isEmpty(dictionary){
  for (var key in dictionary){
    if(dictionary.hasOwnProperty(key))
    return false
  }
  return true
}

/**function to process the form from tagging an effect
 * If the user does not have approriate authorization to access parent folder of the file, the spreadsheet cannot be created or updated 
 * In such case, the user is alerted to ask project owner to give editor rights 
*/
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
      var lastRowInt = currentSheet.getLastRow()
      var elementID = "ID"+lastRowInt.toString()
      console.log("id", elementID, "formObject", "subcat", formObject.topicSelection, formObject)

      currentSheet.appendRow([elementID, formObject.selectedEffect, formObject.susDimension,"","", formObject.topicSelection, formObject.impactPosNeg, formObject.orderEffect, formObject.memoArea, formObject.linkDdl])
      DocumentApp.getUi().alert("Tag was added succesfully to spreadsheet")
    }
  } catch (e) {
    DocumentApp.getUi().alert("You don't have permission to write to parent folder. Please contact project owner.");
  }
}


/** Function used to populate effect dropdown 
 * This function gets the tagged effect with its ID and associated dimension from SpreadSheet 
 * It returns a dictionnary where each ID is associated with an effect and a dimension 
 */
function populateDropdown(){
  
  var document = DocumentApp.getActiveDocument()
  var file = DriveApp.getFileById(document.getId())
  var folder = file.getParents().next()
  var spreadsheets = folder.getFilesByType(MimeType.GOOGLE_SHEETS)
  var dropdownValues = {}

  if (spreadsheets.hasNext()){
    var spreadsheet = spreadsheets.next()
    var currentSheet = SpreadsheetApp.openById(spreadsheet.getId())
  
    //we want values from the first three columns of the spreadsheet
    var lastrow = currentSheet.getLastRow()
    var idColumn = currentSheet.getRange("A2:A"+lastrow)
    var effectColumn = currentSheet.getRange("B2:B" +lastrow)
    var dimensionColumn = currentSheet.getRange("C2:C"+lastrow)

    var idData = idColumn.getValues();
    var effectData = effectColumn.getValues();
    var dimensionData = dimensionColumn.getValues();

    for(var i = 0; i<idData.length;i++){
      //if row is empty, we go to the following row
      if (idData[i]==""){
        continue
      }else{
        dropdownValues[idData[i]] = [effectData[i],dimensionData[i]]
      }
    }
  }else{
    dropdownValues = {}
  }
  return dropdownValues
}
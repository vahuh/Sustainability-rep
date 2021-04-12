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


/** Function that gets the selected text from the user 
 * Returns the user selection as a string so that it can be used */ 
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

/** Function to create Sheets in current folder */  
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
 * If document categories are empty, the new value is directly added as first value 
 */
function setCatProperties(catType, inputText){
  //Get Previous properties 
  let documentProperties = PropertiesService.getDocumentProperties();
  let categories = documentProperties.getProperty('CATEGORIES')
  let subCategories = documentProperties.getProperty('SUBCATEGORIES')
  if (catType == 'Category'){
    if (!addCatProperty(categories,inputText)){
      if (categories == null){
        documentProperties.setProperties('CATEGORIES',inputText)
      }
      else{
        documentProperties.setProperty('CATEGORIES', categories + "," + inputText);
      }
    }
  }
  else if(catType == 'SubCategory'){
    if (!addCatProperty(subCategories,inputText)){
      if (subcategories == null){
        documentProperties.setProperties('SUBCATEGORIES',inputText)
      }else{
        documentProperties.setProperty('SUBCATEGORIES', subCategories + "," + inputText);
      }
    }
  }
}


/** Function to check if a category or subcategory value is already existing in the document properties 
 * If not, the function returns false
 * If yes, the function returns true
 */
function addCatProperty(propertyValues,newProperty){
  propertyValueList = propertyValues.split(',')
  for (var i in propertyValueList){
    if (propertyValueList[i].toString() == newProperty.toString()){
      return true
    }
  }
}


/** Function that shows the popup for tagging selected text 
 * It adds the defined topics to a dropdown in the window 
 * It adds the already tagged effects from spreadsheet in the "leads to" dropdown list
 */
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

//function used to check if a value is already in a list, returns true if it is the case
function checkValuePresence(list, text){
  for (var i in list){
    if (list[i]==text){
      return true
    }
  }
}

/**Function that shows the popup for categorizing an effect
 * Gets all the existing effect values and pushes them to a list that is visible to the user on the effect selection dropdown 
 * No values are added if there are no existing values 
*/
function showCatPopup(catType){
  let documentProperties = PropertiesService.getDocumentProperties()
  let categories = documentProperties.getProperty('CATEGORIES')
  let subCategories = documentProperties.getProperty('SUBCATEGORIES')
  let htmlTemplate = HtmlService.createTemplateFromFile('categories')
  let ddOptionDict = populateDropdown()
  var ddValues = []
  var catValues = []
  var subCatValues = []
  
  if (catType == 'Category'){
    if (categories){
      var existingCategories = categories.split(",")
      for (var i in existingCategories){
        //Only distinct strings are sent to html, if the element is null, it is also left out
        if (!checkValuePresence(catValues, existingCategories[i]) && existingCategories[i] != 'null'){
          catValues.push(existingCategories[i])
        }
        else if (existingCategories[i]=='null'){
          continue
        }
      }
      htmlTemplate.data = catValues
    }else{
      htmlTemplate.data = []
    }
  }else if (catType == 'SubCategory'){
    if (subCategories){
      var existingSubCategories = subCategories.split(",")
      for (var i in existingSubCategories){
        //Only distinct strings are sent to html, if the element is null, it is also left out
        if (!checkValuePresence(catValues, existingSubCategories[i]) && existingSubCategories[i] != 'null'){
          subCatValues.push(existingSubCategories[i])
        }
      }
      htmlTemplate.data = subCatValues
    }else{
      htmlTemplate.data = []
    }
  }

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


/** Submit function from button in categories.html 
 * It takes as an input all the elements that have an id in the html form and the category type 
 * Category type is defined by the initial button click when opening the categorization page
 * The function checks if the selected effect is already associated with a category or subcategory 
 * If it is the case, it is checked from the user if the value should be overwritten
 * When there is no value associated with the effect, the category or subcategory value is set in the corresponding cell in spreadsheet
 */
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
    var idData = idColumn.getValues()
    var effectRow = null
    var catText = checkInputType(formObject.inputCategory,formObject.catDdl)

    try{
      categorizedEffect = formObject.effectDdl
      selectedID = findEffectID(categorizedEffect)
      console.log('selected ID:',selectedID)
    }catch (e){
      DocumentApp.getUi.alert("An error occured, ID couldn't be found.")
    }
    
    for(var i = 0; i<idData.length;i++){
      //if value is equal to ID, we return the value
      if (idData[i]==selectedID){
        effectRow = i
        break
      }
    }

    if (effectRow == null){
      DocumentApp.getUi.alert("No match found for ID")
    }else {
      rowNumber = effectRow + 2
      rowNumberStr = rowNumber.toString()
      setCatProperties(catType,catText)
      if (catType == "Category"){
        var catCell = currentSheet.getRange("D"+rowNumberStr+":D"+rowNumberStr).getCell(1,1)

        if (catCell.getValue() != ''){
          checkChoice(catCell,catType,catText)
          } 
          else{
          catCell.setValue(catText);
          }
          }
      else if(catType == "SubCategory"){
        var subCatCell = currentSheet.getRange("E"+rowNumberStr+":E"+rowNumberStr).getCell(1,1)
        if (subCatCell.getValue() !=''){
          checkChoice(subCatCell,catType,catText)
        }
        else{
          subCatCell.setValue(catText)
        } 
      }
    }
    
  }
}


/** Function to check if the value for category or subcategory comes from input field or ddl 
 * If there is an existing value in input field, it is taken as the (sub)category value 
 * If input field is left empty, the category will be set as the ddl value 
 * If both values are empty, an empty string is returned 
 */
function checkInputType(inputValue, ddlValue){
  let finalCatValue = ''
  if (inputValue != ''){
    finalCatValue = inputValue
  }else if(ddlValue != '' && inputValue == ''){
    finalCatValue = ddlValue
  }else{
    finalCatValue
  }
  return finalCatValue
}

/** Function that asks the user if they are sure that they want to change the value 
 * Function is called when a value is already set in the cell 
 */

function checkChoice(cell, catType, newText){
  let ui = DocumentApp.getUi()
  let answer = ui.alert('Current '+catType+' value: '+cell.getValue(),'Are you sure you want to overwrite current value?', ui.ButtonSet.YES_NO)
  if (answer == ui.Button.YES) {
    cell.setValue(newText);
    }
    else{
      ui.alert("Value change cancelled")
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
      let lastRowInt = currentSheet.getLastRow()
      let elementID = "ID" + lastRowInt.toString()
      console.log("id", elementID, "formObject", "subcat", formObject.topicSelection, formObject)
      let data = currentSheet.getDataRange().getValues();
      let ind = null;
      for (var i = 0; i < data.length; i++) {
        if (data[i][1] == formObject.selectedEffect) {
          ind = i
          break;
        }
      }
      if (ind != null) {
        let range = currentSheet.getRange(`A${ind + 1}:J${ind + 1}`)
        range.setValues([[elementID, formObject.selectedEffect, formObject.susDimension, "", "", formObject.topicSelection, formObject.impactPosNeg, formObject.orderEffect, formObject.memoArea, formObject.linkDdl]])
      }
      else {
        currentSheet.appendRow([elementID, formObject.selectedEffect, formObject.susDimension, "", "", formObject.topicSelection, formObject.impactPosNeg, formObject.orderEffect, formObject.memoArea, formObject.linkDdl])
      }
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

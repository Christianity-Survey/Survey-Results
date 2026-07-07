// Open a form by ID.
var form = FormApp.openById('FILL IN FORM ID HERE');
// Open a sheet by ID.
var sheet = SpreadsheetApp.openById('FILL IN SPREADSHEET ID HERE').getSheets()[0];

// variables for putting the questions and answers in the right position
var question_position = 0;
var answers_position = 0;

// main function to run
function getFormValues() {
  form.getItems().forEach(callback);
}

// Iterate over all questions 
function callback(el){
  //console.log(el)
  
  if(el.getType() === FormApp.ItemType.CHECKBOX) {
    sheet.getRange(question_position +1, 1).setValue(el.getTitle())
    sheet.getRange(question_position +1, 2).setValue(el.getType().toString().toLowerCase())
    sheet.getRange(question_position +1, 3).setValue(el.asCheckboxItem().getChoices().map(c=>c.getValue()).join(','))
    // console.log(`${el.getId()}\t${el.getTitle()}\t${el.getType()}\t${el.asCheckboxItem().getChoices().map(c=>c.getValue())}`)
  } else if(el.getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
    sheet.getRange(question_position +1, 1).setValue(el.getTitle())
    sheet.getRange(question_position +1, 2).setValue(el.getType().toString().toLowerCase())
    sheet.getRange(question_position +1, 3).setValue(el.asMultipleChoiceItem().getChoices().map(c=>c.getValue()).join(','))
    // console.log(`${el.getId()}\t${el.getTitle()}\t${el.getType()}\t${el.asMultipleChoiceItem().getChoices().map(c=>c.getValue())}`)
  } else if(el.getType() === FormApp.ItemType.SCALE) {
    sheet.getRange(question_position +1, 1).setValue(el.getTitle())
    sheet.getRange(question_position +1, 2).setValue(el.getType().toString().toLowerCase())
    sheet.getRange(question_position +1, 3).setValue(`${el.asScaleItem().getLeftLabel()},${el.asScaleItem().getRightLabel()}`)
    // console.log(`${el.getId()}\t${el.getTitle()}\t${el.getType()}\t${el.asScaleItem().getLeftLabel()},${el.asScaleItem().getRightLabel()}`)
  } else if(el.getType() === FormApp.ItemType.CHECKBOX_GRID) {
    sheet.getRange(question_position +1, 1).setValue(el.getTitle())
    sheet.getRange(question_position +1, 2).setValue(el.getType().toString().toLowerCase())
    sheet.getRange(question_position +1, 3).setValue(el.asCheckboxGridItem().getRows().join(','))
    sheet.getRange(question_position +1, 3).setValue(el.asCheckboxGridItem().getColumns().join(','))
    // last scale.columns
    // scale.multiple true
    // console.log(`${el.getId()}\t${el.getTitle()}\t${el.getType()}\t${el.asCheckboxGridItem().getRows().join(',')}\t${el.asCheckboxGridItem().getColumns()}`)
  } else if(el.getType() === FormApp.ItemType.GRID) {
    sheet.getRange(question_position +1, 1).setValue(el.getTitle())
    sheet.getRange(question_position +1, 2).setValue(el.getType().toString().toLowerCase())
    sheet.getRange(question_position +1, 3).setValue(el.asGridItem().getRows().join(','))
    sheet.getRange(question_position +1, 3).setValue(el.asGridItem().getColumns().join(','))
    // last scale.columns
    // scale.multiple false
    // console.log(`${el.getId()}\t${el.getTitle()}\t${el.getType()}\t${el.asGridItem().getRows().join(',')}\t${el.asGridItem().getColumns()}`)
  } else if(el.getType() === FormApp.ItemType.LIST) {
    sheet.getRange(question_position +1, 1).setValue(el.getTitle())
    sheet.getRange(question_position +1, 2).setValue(el.getType().toString().toLowerCase())
    sheet.getRange(question_position +1, 3).setValue(el.asListItem().getChoices().map(c=>c.getValue()).join(','))
    // console.log(`${el.getId()}\t${el.getTitle()}\t${el.getType()}\t${el.asListItem().getChoices().map(c=>c.getValue())}`)
  } else if(el.getType() !== FormApp.ItemType.PAGE_BREAK)  {
    sheet.getRange(question_position +1, 1).setValue(el.getTitle())
    sheet.getRange(question_position +1, 2).setValue(el.getType().toString().toLowerCase())
    // console.log(`${el.getId()}\t${el.getTitle()}\t${el.getType()}`)
  }
  question_position++
  
}
/// ---- Main ------ RS
/*
1P75hcrQt2o0D6iGslz5ixL-obXL81eeIHAL8bTuLuL6QzBia80cATdp6

*/

var ss
var invoiceSheet
var glSheet
var glData
var propSheet
var propSheetData
var entrySheet
var entrySheetData
var chargedUtil
var printInv = false

function onOpen() {
  menuOne()
}

function  menuOne(){
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Invoicing')
    .addItem('Save & Clear', 'justSaveNClear')
    .addItem('Make Invoice', 'writeToInvoice')
    .addItem('Clear Entry', 'menuClearEntry') // calls funtion clearDataEntry(false)
    .addSeparator()
    .addSubMenu(ui.createMenu('Utils')
      .addItem('Make Entry Sheet', 'menuItem2')
      .addItem('Make Invoice Template', 'makeInvoiceSheet')
      .addItem('Edit Menu', 'menuTwo'))
    .addToUi();
}

function menuTwo() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Invoicing')
    .addItem('Edit this/old Invoice', 'editEntry')
    .addItem('Print Invoice', 'printInvoice')
    .addSubMenu(ui.createMenu('Utils')
      .addItem('Full Menu', 'menuOne'))
    .addToUi();
}

function menuThree() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Statements')
    .addItem('Generate Invoice', 'makeStatement')
  .addToUi();
}



function stillToDo(){

}

function menuClearEntry() {
  clearDataEntry(false)
}

function anotherItem() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('You clicked the second menu item!');
}

function menuItem2() {
  makeEntryDataSheet()
  //SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  //   .alert('You clicked the second menu item!');
}

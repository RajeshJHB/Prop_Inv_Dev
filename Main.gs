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
var menuNum = 0

function onOpen() {
  menuZero()
}

function  menuZero(){
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Properties')
    .addItem('Make Invoice Template', 'makeInvoiceSheet')
    .addItem('Make Entry Sheet', 'makeES')
    .addItem('Make Statement Sheet', 'makeSTSheet')
 .addSeparator()
  .addToUi();
  menuNum = 0

}

function  menuOne(){
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Properties')
    .addItem('Save & Clear', 'justSaveNClear')
    .addItem('Make Invoice', 'writeToInvoice')
    .addItem('Clear Entry', 'menuClearEntry') // calls funtion clearDataEntry(false)
    .addSeparator()
    .addSubMenu(ui.createMenu('Utils')
      .addItem('Make Entry Sheet', 'makeES')
      .addItem('Make Invoice Template', 'makeInvoiceSheet')
      .addItem('Edit Menu', 'menuTwo'))
    .addToUi();
  menuNum = 1
}

function menuTwo() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Properties')
    .addItem('Edit this/old Invoice', 'editEntry')
    .addItem('Print Invoice', 'printInvoice')
    .addSubMenu(ui.createMenu('Utils')
      .addItem('Full Menu', 'menuOne'))
    .addToUi();
  menuNum = 2
}

function menuThree() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Properties')
    .addItem('Generate Statement', 'makeStatement')
    .addItem('Sort By Date', 'sortSTByDate')
    .addSeparator()
    .addSubMenu(ui.createMenu('Utils')
      .addItem('Make Statement Sheet', 'makeSTSheet'))
  .addToUi();
  menuNum = 3
}



function stillToDo(){

}

function menuClearEntry() {
  clearDataEntry(false)
}

function makeES() {
  makeEntryDataSheet()
  //SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  //   .alert('You clicked the second menu item!');
}

function anotherItem() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('You clicked the second menu item!');
}

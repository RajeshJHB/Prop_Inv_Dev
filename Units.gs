/// ---- Units ------ RS

function getEntrySheet(){
  ss = SpreadsheetApp.getActiveSpreadsheet();
  entrySheet = ss.getSheetByName("DataEntry")
  entrySheetData = entrySheet.getDataRange().getValues();
  myLog("Found Entry Sheet")
}

function getPropSheet(){
  getEntrySheet()
  theProp = entrySheetData[1][1]    //B2
  if (theProp == ""){
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('First Pick a Property');
    return
  }
  myLog("theProp  "+ theProp)
  propSheet = ss.getSheetByName(theProp)
  if (propSheet == null){
    ss.insertSheet(theProp)
    propSheet = ss.getSheetByName("DataEntry")
  }
  propSheetData = propSheet.getDataRange().getValues()
  return(theProp)
}


function getGlobalSheet()
{
  getEntrySheet()
  glSheet = ss.getSheetByName("Globals")
  glData = glSheet.getDataRange().getValues();

}



function getEntrySheetOnce(){
 // myLog(entrySheet)
  if (entrySheet == null){
    getEntrySheet()
  }
}

function getPropSheetLastDataCol(){
  var maxCol = propSheet.getMaxColumns()
  myLog("maxCols  "+ maxCol)
  var prColIndex = findLastColInSheet(propSheetData)   // in row 1
  if (maxCol == prColIndex) [
    propSheet.insertColumnAfter(maxCol)   // ("add a col" )
  ]
  return (prColIndex)
}

function findLastColInSheet(sheetData) {  //find first column with blank in first row.
  for(var i = 0; i<sheetData.length+5;i++)
  {
    if ((sheetData[0][i] ==   "")  || (sheetData[0][i] ==  null))  //(theItem || null))    // 0 - row 0
    {
      myLog("Empty Col @ " +(i))
      return (i)
    }
  }
}

//-------------- that is why 2 functions -----------------------


function myLog(msg) { Logger.log(msg) }
function myArLog(msg) {Logger.log(JSON.stringify(msg))};
function onetoTwod(inArr,div){
  var newArr = [];
  while(inArr.length) newArr.push(inArr.splice(0,div));
  return newArr
}

function findRowInSheet(sheetData,theItem,index)
{
  var sheetRowData = []
    for(var i = 0; i<sheetData.length+1;i++)
    {
      if(sheetData[i][index] == theItem)    // 0 - row 1
      {
        myLog("Found Row (Index 0) " + theItem +" at "+(i))
        sheetRowData = sheetData[i]
        sheetRowData[0] = i;
        return sheetRowData
      }
    }
}

function findColInSheet(sheetData,theItem,index) {  // 0 - col 1
  for(var i = 0; i<sheetData.length+1;i++)
  { var colVal = sheetData[index][i]
    if (colVal == "") {return (i)}
    if( colVal == theItem)    // 0 - row 0
    {
      myLog("Found Column (Index 0) " + theItem +" at "+(i))
      return (i+1)  //Return base 1
    }
  }
}




function myYNPrompt(question){
    var result = SpreadsheetApp.getUi().alert(question, SpreadsheetApp.getUi().ButtonSet.YES_NO);
    SpreadsheetApp.getActive().toast(result);
    if (result == "CLOSE") {return("NO")}
    return(result)
}

function myAlert(theNote){
    SpreadsheetApp.getUi()
     .alert(theNote);
}

function writeSaved(wrtSaved){
  getEntrySheet()
  var myRange = entrySheet.getRange("C1")
  if (!wrtSaved){
    entrySheet.getRange("C1").setBackground("#f70713")
    var theStyle = SpreadsheetApp.newTextStyle()
      .setBold(true)
      .setForegroundColor("#ffff00")
      .build();
    var theText = SpreadsheetApp.newRichTextValue()
        .setText("Not Saved")
        .setTextStyle(theStyle)
        .build();
    myRange.setRichTextValue(theText)
  } else
  {
    entrySheet.getRange("C1").setBackground("#07f74f")
    var theStyle = SpreadsheetApp.newTextStyle()
      .setBold(true)
      .setForegroundColor("#000000")
      .build();
    var theText = SpreadsheetApp.newRichTextValue()
        .setText("Saved")
        .setTextStyle(theStyle)
        .build();
    myRange.setRichTextValue(theText)
  }
}

function myFormat(x, n){
    x = parseFloat(x);
    n = n || 2;
    return parseFloat(x.toFixed(n))
}

function getHex(input) {
  return SpreadsheetApp.getActiveSpreadsheet().getRange(input).getBackgrounds();
}

// ____________________________________________________________________________________________________________________

function testFormat(){
  var num = 11/7
  num = myFormat(num,2)
  myLog(num)
}

function checkEntrySheet(){
  getEntrySheet()
  myArLog(entrySheetData)
  uVal = entrySheet.getRange(4,9,1,1).getValue()
  myLog(uVal)
}


function testOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Custom Menu')
      .addItem('First item', 'menuItem1')
      .addSeparator()
      .addSubMenu(ui.createMenu('Sub-menu')
          .addItem('Second item', 'menuItem2'))
      .addToUi();
}

function testmenu(){
var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Invoicing')
    .addItem('Save & Clear', 'justSaveNClear')
    .addItem('Invoice & Clear', 'writeToInvoice')
    .addItem('Clear Entry', 'menuClearEntry') // calls funtion clearDataEntry(false)
    .addSeparator()
    .addSubMenu(ui.createMenu('Utils')
      .addItem('Make Entry Sheet', 'menuItem2'))
    .addToUi();

}

function testCol() {
  getPropSheet()
  theIndi = getPropSheetLastDataCol()
  myLog("theIndi " + theIndi)

}

function testNote(){
  ss = SpreadsheetApp.getActiveSpreadsheet();
  entrySheet = ss.getSheetByName("DataEntry")
  entrySheet.getRange("E7:E14").setNote("Transaction Value" )
}

function testIn(){
  var theCell = [0,0]
  theCell[0] = 15
  theCell[1] = 3
  drCrRange = [7,8,9,10,11,12,13,14]
  drCrIn = drCrRange.indexOf(theCell[0])
  myLog(drCrIn)

  if (drCrIn != -1) {
    myLog("In range")
  }
  else {
    myLog("Out of Range")
  }

}


function testMisc(){
    prop =   getPropSheet() // from data entry screen
    myLog(prop)
    lastCol = getPropSheetLastDataCol() //next new column
    lastV = propSheet.getRange(23,lastCol,1,1).getValue()
    myLog(lastV)
}

function testMisc2(){
    drList = ["Rent","Fibre", "Rates","Service Fee", "Interest Charge","Penalty","Deposit Charge","None" ]
    getEntrySheet()
    subListVal = entrySheet.getRange(7,4,1,1).getValue()
    checkWord = drList.includes(subListVal);
    myLog("sublistValue " + subListVal + "Checkword " +checkWord)
}
function testMisc3(){
    var result = SpreadsheetApp.getUi().alert("Overwrite Current Data", SpreadsheetApp.getUi().ButtonSet.YES_NO);
      SpreadsheetApp.getActive().toast(result);
    myLog(result)
}

function testMisc4(){
  getEntrySheet()
  d2Val =  entrySheet.getRange('D2').getValue()
  d2fVal = entrySheet.getRange('D2').getFormulaR1C1()
  if (d2fVal == ""){
   myLog(d2Val)
  }
  myLog(d2Val)
  myLog(d2fVal)
}

function testMisc5(){
  prRet = myYNPrompt("Test Prompt")
  myLog(prRet)
  if (prRet != "YES"){
    myLog("YES -> " +prRet)
  }
}



function testCreate(){
  ss = SpreadsheetApp.getActiveSpreadsheet();
  invSheet = ss.getSheetByName("Invoice")
  if (invSheet == null){
    ss.insertSheet("Invoice")
    invSheet = ss.getSheetByName("Invoice")
  }
  ss.setActiveSheet(ss.getSheetByName("Invoice"))
}

var saveToRootFolder = true

function _exportBlob(blob, fileName, spreadsheet) {
  blob = blob.setName(fileName)
  var folder = saveToRootFolder ? DriveApp : DriveApp.getFileById(spreadsheet.getId()).getParents().next()
  var pdfFile = folder.createFile(blob)
  
  // Display a modal dialog box with custom HtmlService content.
  const htmlOutput = HtmlService
    .createHtmlOutput('<p>Click to open <a href="' + pdfFile.getUrl() + '" target="_blank">' + fileName + '</a></p>')
    .setWidth(300)
    .setHeight(80)
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export Successful')
}


function _getAsBlob(url, sheet, range) {
  var rangeParam = ''
  var sheetParam = ''
  if (range) {
    rangeParam =
      '&r1=' + (range.getRow() - 1)
      + '&r2=' + range.getLastRow()
      + '&c1=' + (range.getColumn() - 1)
      + '&c2=' + range.getLastColumn()
  }
  if (sheet) {
    sheetParam = '&gid=' + sheet.getSheetId()
  }
  // A credit to https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
  // these parameters are reverse-engineered (not officially documented by Google)
  // they may break overtime.
  var exportUrl = url.replace(/\/edit.*$/, '')
      + '/export?exportFormat=pdf&format=pdf'
      + '&size=LETTER'
      + '&portrait=true'
      + '&fitw=true'
      + '&top_margin=0.75'
      + '&bottom_margin=0.75'
      + '&left_margin=0.7'
      + '&right_margin=0.7'
      + '&sheetnames=false&printtitle=false'
      + '&pagenum=UNDEFINED' // change it to CENTER to print page numbers
      + '&gridlines=true'
      + '&fzr=FALSE'
      + sheetParam
      + rangeParam
      
  Logger.log('exportUrl=' + exportUrl)
   var response
  var i = 0
  for (; i < 5; i += 1) {
    response = UrlFetchApp.fetch(exportUrl, {
      muteHttpExceptions: true,
      headers: {
        Authorization: 'Bearer ' +  ScriptApp.getOAuthToken(),
      },
    })
    if (response.getResponseCode() === 429) {
      // printing too fast, retrying
      Utilities.sleep(3000)
    } else {
      break
    }
  }
  
  if (i === 5) {
    throw new Error('Printing failed. Too many sheets to print.')
  }
  
  return response.getBlob()
}

function exportCurrentSheetAsPDF() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var currentSheet = SpreadsheetApp.getActiveSheet()
  
  var blob = _getAsBlob(spreadsheet.getUrl(), currentSheet)
  _exportBlob(blob, currentSheet.getName(), spreadsheet)
}

function exportPartAsPDF(predefinedRanges) {
  var ui = SpreadsheetApp.getUi()
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  
  var selectedRanges
  var fileSuffix
  if (predefinedRanges) {
    selectedRanges = predefinedRanges
    fileSuffix = '-predefined'
  } else {
    var activeRangeList = spreadsheet.getActiveRangeList()
    if (!activeRangeList) {
      ui.alert('Please select at least one range to export')
      return
    }
    selectedRanges = activeRangeList.getRanges()
    fileSuffix = '-selected'
  }
  
  if (selectedRanges.length === 1) {
    // special export with formatting
    var currentSheet = selectedRanges[0].getSheet()
    var blob = _getAsBlob(spreadsheet.getUrl(), currentSheet, selectedRanges[0])
    
    var fileName = spreadsheet.getName() + fileSuffix
    _exportBlob(blob, fileName, spreadsheet)
    return
  }
  
  var tempSpreadsheet = SpreadsheetApp.create(spreadsheet.getName() + fileSuffix)
  if (!saveToRootFolder) {
    DriveApp.getFileById(tempSpreadsheet.getId()).moveTo(DriveApp.getFileById(spreadsheet.getId()).getParents().next())
  }
  var tempSheets = tempSpreadsheet.getSheets()
  var sheet1 = tempSheets.length > 0 ? tempSheets[0] : undefined
  SpreadsheetApp.setActiveSpreadsheet(tempSpreadsheet)
  tempSpreadsheet.setSpreadsheetTimeZone(spreadsheet.getSpreadsheetTimeZone())
  tempSpreadsheet.setSpreadsheetLocale(spreadsheet.getSpreadsheetLocale())
  for (var i = 0; i < selectedRanges.length; i++) {
    var selectedRange = selectedRanges[i]
    var originalSheet = selectedRange.getSheet()
    var originalSheetName = originalSheet.getName()
    
    var destSheet = tempSpreadsheet.getSheetByName(originalSheetName)
    if (!destSheet) {
      destSheet = tempSpreadsheet.insertSheet(originalSheetName)
    }
    
    Logger.log('a1notation=' + selectedRange.getA1Notation())
    var destRange = destSheet.getRange(selectedRange.getA1Notation())
    destRange.setValues(selectedRange.getValues())
    destRange.setTextStyles(selectedRange.getTextStyles())
    destRange.setBackgrounds(selectedRange.getBackgrounds())
    destRange.setFontColors(selectedRange.getFontColors())
    destRange.setFontFamilies(selectedRange.getFontFamilies())
    destRange.setFontLines(selectedRange.getFontLines())
    destRange.setFontStyles(selectedRange.getFontStyles())
    destRange.setFontWeights(selectedRange.getFontWeights())
    destRange.setHorizontalAlignments(selectedRange.getHorizontalAlignments())
    destRange.setNumberFormats(selectedRange.getNumberFormats())
    destRange.setTextDirections(selectedRange.getTextDirections())
    destRange.setTextRotations(selectedRange.getTextRotations())
    destRange.setVerticalAlignments(selectedRange.getVerticalAlignments())
    destRange.setWrapStrategies(selectedRange.getWrapStrategies())
  }
  
  // remove empty Sheet1
  if (sheet1) {
    Logger.log('lastcol = ' + sheet1.getLastColumn() + ',lastrow=' + sheet1.getLastRow())
    if (sheet1 && sheet1.getLastColumn() === 0 && sheet1.getLastRow() === 0) {
      tempSpreadsheet.deleteSheet(sheet1)
    }
  }
  exportAsPDF()
  SpreadsheetApp.setActiveSpreadsheet(spreadsheet)
  DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true)
}





function exportNamedRangesAsPDF() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var allNamedRanges = spreadsheet.getNamedRanges()
  var toPrintNamedRanges = []
  for (var i = 0; i < allNamedRanges.length; i++) {
    var namedRange = allNamedRanges[i]
    if (/^print_area_.*$/.test(namedRange.getName())) {
      Logger.log('found named range ' + namedRange.getName())
      toPrintNamedRanges.push(namedRange.getRange())
    }
  }
  if (toPrintNamedRanges.length === 0) {
    SpreadsheetApp.getUi().alert('No print areas found. Please add at least one \'print_area_1\' named range in the menu Data > Named ranges.')
    return
  } else {
    toPrintNamedRanges.sort(function (a, b) {
      return a.getSheet().getIndex() - b.getSheet().getIndex()
    })
    exportPartAsPDF(toPrintNamedRanges)
  }
}


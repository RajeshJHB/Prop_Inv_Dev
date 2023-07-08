/// ---- StatementEntry ------ RS

function getSpecificPropSheet(theProp){
  ss = SpreadsheetApp.getActiveSpreadsheet();
  if (theProp == ""){
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('First Pick a Property');
    return
  }
//  myLog("theProp  "+ theProp)
  propSheet = ss.getSheetByName(theProp)
  if (propSheet == null){
    ss.insertSheet(theProp)
    propSheet = ss.getSheetByName(theProp)
  }
  propSheetData = propSheet.getDataRange().getValues()
  return(theProp)
}


function makeSTSheet(){
  ss = SpreadsheetApp.getActiveSpreadsheet();
  var statementSheet = ss.getSheetByName("Statement")
  if (statementSheet == null){
    ss.insertSheet("Statement")
    statementSheet = ss.getSheetByName("Statement")
  }
  maxCol = statementSheet.getMaxColumns()
  if (maxCol <10){
    statementSheet.insertColumnsAfter(1,10-maxCol)
  }
//  prRange = glSheet.getRange('B2:B32')
  // Wite static info
  statementSheet.getRange('A1:A9').setValues([
    ["Pick A Property"],
    [""],
    [""],
    [""],
    [""],
    ["Invoice Start"],
    [""],
    [""],
    ["Invoice End"],
  ])

 
  statementSheet.getRange(['A1:A11']).setBackground([["#a2c4c9"]])  // Set grey colour
  
  glSheet = ss.getSheetByName("Globals")
  var propRange = glSheet.getRange('B2:B50')
  var arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Choose from Menu Only").requireValueInRange(propRange).build() // from items in the sheet
  statementSheet.getRange('A2').setDataValidation(arule) // Choose property

  statementSheet.getRange('A3').setValue("=index((Globals!C2:C40),match(A2,Globals!B2:B40,0),)") // Display Current tenat code
  statementSheet.getRange('A4').setValue("=INDEX(Globals!C53:C100,MATCH(A3,Globals!B53:B100))") // Display Current Tenant name

  // Set all notes in cell
  statementSheet.getRange("A2").setNote("Statement For Which Roperty")
  statementSheet.getRange("A7").setNote("Statement Start Invoice")
  statementSheet.getRange("A10").setNote("Statement End Value")
 }

 function populateInvNumber(aProp){
  ss = SpreadsheetApp.getActiveSpreadsheet();
  var statementSheet = ss.getSheetByName("Statement")
  var invList = []
//  aProp = "CHU10"
  getSpecificPropSheet(aProp)
  invList = getINumbers()
  if (invList.length > 0){
    arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Choose from List Only").requireValueInList(invList).build()
    statementSheet.getRange(7,1,1,1).setDataValidation(arule)
    statementSheet.getRange(10,1,1,1).setDataValidation(arule)
  }
  statementSheet.getRange(7,1,1,1).setValue(invList[0])
  statementSheet.getRange(10,1,1,1).setValue(invList[0])
}

function sortSTByDate() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C7:H7').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).sort({column: 3, ascending: true});
  spreadsheet.getRange('C7').activate();
};

function makeStatement(){
  ss = SpreadsheetApp.getActiveSpreadsheet();
  statementSheet = ss.getSheetByName("Statement")
  statementData = statementSheet.getDataRange().getValues();
  var statementText = []
  var sProp = statementSheet.getRange(2,1,1,1).getValue()
  var stInvNum = statementSheet.getRange(7,1,1,1).getValue()
  var endInvNum = statementSheet.getRange(10,1,1,1).getValue()
  if (stInvNum > endInvNum) {
    var aTemp = stInvNum
    stInvNum = endInvNum
    endInvNum = aTemp
  }
  var numInvs = endInvNum - stInvNum
  getSpecificPropSheet(sProp)
  var maxCol = propSheet.getMaxColumns()
  myLog(maxCol)
  var invAt = 22
  firstInv = true
  prnStRow = 2
  prnStCol = 3
  printRowNum = 1
  drTotal = 0
  crTotal = 0

  for(var i = 0; i<maxCol+1;i++)
  {
    var invNum = propSheetData[invAt][i]
    if((invNum != null) &&((invNum >= stInvNum) && (invNum <= endInvNum)))     // 0 - row 0
    {
      myLog("Inv : " + invNum)
      if (firstInv){
        getInvoiceDrCr(i)
        firstInv = false
      }
      else {
        getInvoiceDrCr(i)
      }
    }
    sortSTByDate()
  }

  /*
  statementText = [["", "", "","","",""]]
  statementSheet.getRange(prnStRow + printRowNum,prnStCol,1,5).setValues(statementText)
  printRowNum++
 
  statementText = [["", "Totals", drTotal,crTotal,""]]
  statementSheet.getRange(prnStRow + printRowNum,prnStCol,1,5).setValues(statementText)
  printRowNum++
  statementText = [["", "BALANCE", drTotal-crTotal,"",""]]
  statementSheet.getRange(prnStRow + printRowNum,prnStCol,1,5).setValues(statementText)
  */
}

function getInvoiceDrCr(invIndex)
{
//  getPropSheet()
  var dEntryIndex = 0
  var wrtData = []
  var inArr = []
  var items =   propSheetData.length
  var minus1 = -1
  var statementText = []
  
  if (firstInv){
    var now  = new Date()

    statementText = [["", "STATEMENT", "",now,"",""]]
    statementSheet.getRange(prnStRow + printRowNum,prnStCol,1,6).setValues(statementText)
    printRowNum++

    statementText = [["Property Code", propSheetData[minus1+11][invIndex], "","","",""]]
    statementSheet.getRange(prnStRow + printRowNum,prnStCol,1,6).setValues(statementText)
    printRowNum++

    statementText = [["Occupant", propSheetData[minus1+14][invIndex], "","","",""]]
    statementSheet.getRange(prnStRow + printRowNum,prnStCol,1,6).setValues(statementText)
    printRowNum++
    statementText = [["", "", "","","",""]]
    statementSheet.getRange(prnStRow + printRowNum,prnStCol,1,6).setValues(statementText)
    printRowNum++
    statementText = [["Date", "Description", "Debit","Credit","Balance","Notes"]]
    statementSheet.getRange(prnStRow + printRowNum,prnStCol,1,6).setValues(statementText)
    printRowNum++
    statementText = [[propSheetData[minus1+2][invIndex],
                  "Previous Balance Inv: " +propSheetData[minus1+5][invIndex],
                  propSheetData[minus1+7][invIndex],
                  "",
                  propSheetData[minus1+7][invIndex],
                  ""]]
    drTotal = drTotal + propSheetData[minus1+7][invIndex]
    statementSheet.getRange(prnStRow + printRowNum,prnStCol,1,6).setValues(statementText)
    printRowNum++
  }
  statementText = [[propSheetData[minus1+20][invIndex],
                  "Invoice : " +propSheetData[minus1+23][invIndex],
                  "",
                  "",
                  "=G"+(printRowNum+1)+("+E"+(printRowNum+2))+("-F"+(printRowNum+2)),
                  ""]]
  statementSheet.getRange(prnStRow + printRowNum,prnStCol,1,6).setValues(statementText)
  printRowNum++

  // Do Utils
  var offset = 0
  for(var j = 0; j<3;j++){
    offset = j*9
    if (propSheetData[minus1+30+offset][invIndex] != "")
    {
      statementText = [[propSheetData[minus1+29+offset][invIndex],
                      propSheetData[minus1+30+offset][invIndex] + " " + propSheetData[minus1+31][invIndex],
                      propSheetData[minus1+36+offset][invIndex],
                      "",
                      "=G"+(printRowNum+1)+("+E"+(printRowNum+2))+("-F"+(printRowNum+2)),
                      ""]]
      statementSheet.getRange(prnStRow + printRowNum,prnStCol,1,6).setValues(statementText)
      printRowNum++
      drTotal = drTotal + propSheetData[minus1+36+offset][invIndex]
    }
  }
  for(var j = 0; j<8;j++){
    offset = j*9
    if (propSheetData[minus1+57+offset][invIndex] == "Debit (+)")
    {
      statementText = [[propSheetData[minus1+56+offset][invIndex],
                      propSheetData[minus1+58+offset][invIndex],
                      propSheetData[minus1+59+offset][invIndex],
                      "",
                      "=G"+(printRowNum+1)+("+E"+(printRowNum+2))+("-F"+(printRowNum+2)),
                      propSheetData[minus1+60+offset][invIndex]]]
      statementSheet.getRange(prnStRow + printRowNum,prnStCol,1,6).setValues(statementText)
      printRowNum++
      drTotal = drTotal + propSheetData[minus1+59+offset][invIndex]
    }
    else if (propSheetData[minus1+57+offset][invIndex] == "Credit (-)")
    {
      statementText = [[propSheetData[minus1+56+offset][invIndex],
                      propSheetData[minus1+58+offset][invIndex],
                      "",
                      propSheetData[minus1+59+offset][invIndex],
                      "=G"+(printRowNum+1)+("+E"+(printRowNum+2))+("-F"+(printRowNum+2)),
                      propSheetData[minus1+60+offset][invIndex]]]
      statementSheet.getRange(prnStRow + printRowNum,prnStCol,1,6).setValues(statementText)
      printRowNum++
      crTotal = crTotal + propSheetData[minus1+59+offset][invIndex]
    }
    else if (propSheetData[minus1+57+offset][invIndex] == "Note")
    {
      statementText = [[propSheetData[minus1+56+offset][invIndex],
                      propSheetData[minus1+58+offset][invIndex],
                      "",
                      "",
                      "=G"+(printRowNum+1)+("+E"+(printRowNum+2))+("-F"+(printRowNum+2)),
                      propSheetData[minus1+60+offset][invIndex]]]
      statementSheet.getRange(prnStRow + printRowNum,prnStCol,1,6).setValues(statementText)
      printRowNum++
    }
    else{

    }
  }
}




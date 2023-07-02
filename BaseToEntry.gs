/// ---- BaseToEntry ------ RS

function loadUp(){
  clearDataEntry(false)  // ClearEntry - don't clear pick property
  addTenFormula()  //ClearEntry - fill in missing formulae
  getLastSaved()  // if last 'temp' load last else setup for new invoice
}

function addTenFormula(){
  getEntrySheet()
  var d2fVal = entrySheet.getRange('D2').getFormulaR1C1()
  if (d2fVal == ""){
    entrySheet.getRange('D2').setValue("=index((Globals!C2:C40),match(B2,Globals!B2:B40,0),)") // Display Current tenat code
  }
  var e2fVal = entrySheet.getRange('E2').getFormulaR1C1()
  if (e2fVal == ""){
    entrySheet.getRange('E2').setValue("=INDEX(Globals!C53:C100,MATCH(D2,Globals!B53:B100))") // Display Current Tenant name
  }
}

function testBE(){
  fromBaseToEntryData(1)
}

function fromBaseToEntryData(colIndex){
//  var proCode = getPropSheet()
  var dEntryIndex = 0
  var wrtData = []
  var items =   propSheetData.length
  myLog("Lenght -> " + items +" colIndex -> " + colIndex)
  myArLog(entrySheetData)
  for (var i =0 ; i <=13; i++ ){    // i = 13
    var inArr = []
    for (var j= 0 ; j<=8; j++) {    // j = 8
          myLog("i-> " +i + "j->" +j + "theIndex -> " + dEntryIndex + "the data -> " +propSheetData[dEntryIndex][colIndex] )
          inArr.push(propSheetData[dEntryIndex][colIndex])
          dEntryIndex ++
    }
    wrtData.push(inArr)
  }
  entrySheet.getRange(1,1,14,9).setValues(wrtData)
  chargedUtil = entrySheet.getRange(4,9,3,1).getDisplayValues() // Charged Util
}

function loadInvAtBottom(colIndex){
//  getPropSheet()
  var dEntryIndex = 0
  var wrtData = []
  var items =   propSheetData.length
  myLog("Lenght -> " + items +" colIndex -> " + colIndex)
  myArLog(entrySheetData)
  for (var i =0 ; i <=13; i++ ){    // i = 13
    var inArr = []
    for (var j= 0 ; j<=8; j++) {    // j = 8
          myLog("i-> " +i + "j->" +j + "theIndex -> " + dEntryIndex + "the data -> " +propSheetData[dEntryIndex][colIndex] )
          inArr.push(propSheetData[dEntryIndex][colIndex])
          dEntryIndex ++
    }
    wrtData.push(inArr)
  }
  entrySheet.getRange(21,1,14,9).setValues(wrtData)
}


function makeAllPickList()
{
  for (var i = 4 ; i<= 14; i++) {
    makePickList(i)
  } // put in the pick lists
}

function getLastSaved()
{
    var prop =   getPropSheet() // from data entry screen
    var lastCol = getPropSheetLastDataCol() //next new column
    myLog("LastCol " + lastCol + " Property : " + prop)
    if (lastCol < 2) {
      entrySheet.getRange(1,2,1,1).setValue(0)
      entrySheet.getRange(1,5,1,1).setValue(0)
      entrySheet.getRange(1,7,1,1).setValue(0)
      entrySheet.getRange(1,9,1,1).setValue(0)
      entrySheet.getRange(2,9,1,1).setValue(0)
      getDeposits()
      getInvNumbers()
      buildMainPickList()
      writeSaved(true)
    return
    }
    lastInv = propSheet.getRange(23,lastCol,1,1).getValue()
    if (lastInv == "Temp"){
      fromBaseToEntryData(lastCol-1)
      buildMainPickList()
      makeAllPickList()
      fixUtilValues() // Calls getInvNumbers()
      writeSaved(true)
    }else  {
      getLastUtil()
      getLastInvNum()
      buildMainPickList()
     // makeAllPickList()
      getDeposits()
      getInvNumbers()
    //  fixUtilValues() // calls getInvNumbers()
      writeSaved(false)
    }

}


function fixUtilValues(){
  getEntrySheet()
/*
  var v1 = entrySheet.getRange(4,5,3,2).getDisplayValues()  // Calculated Util
  myArLog(v1)
  entrySheet.getRange(4,5,3,2).setValues([[0,1],[0,1],[0,1]])
  getInvNumbers()
  entrySheet.getRange(4,5,3,2).setValues(v1)
  */
  entrySheet.getRange(4,9,3,1).setValues(chargedUtil)
  getInvNumbers()
}


function restoreUtilVals(){
  getEntrySheet()
  var v1 = entrySheet.getRange(4,7,3,3).getValues()
  myLog(v1)
  entrySheet.getRange(4,7,3,3).setValues(v1)
}

function getLastUtil() {
  getPropSheet()
  colIndex1 = getPropSheetLastDataCol()
  if (colIndex1 == 1) {
    entrySheet.getRange(2,2,1,1).setValue("None")
     entrySheet.getRange(2,5,1,1).setValue("None")
     entrySheet.getRange(2,7,1,1).setValue("None")
    return
  }
 
  utilA_Name =  propSheetData[29][colIndex1-1]
  entrySheet.getRange(4,3,1,1).setValue(utilA_Name)
  utilA_Type =  propSheetData[30][colIndex1-1]
  entrySheet.getRange(4,4,1,1).setValue(utilA_Type)
  utilA_Last =  propSheetData[32][colIndex1-1]
  entrySheet.getRange(4,5,1,1).setValue(utilA_Last)
  if (utilA_Type != ""){entrySheet.getRange(4,6,1,1).setValue(utilA_Last+1)}
  makePickList(4)

  utilB_Name =  propSheetData[38][colIndex1-1]
  entrySheet.getRange(5,3,1,1).setValue(utilB_Name)
  utilB_Type =  propSheetData[39][colIndex1-1]
  entrySheet.getRange(5,4,1,1).setValue(utilB_Type)
  utilB_Last =  propSheetData[41][colIndex1-1]
  entrySheet.getRange(5,5,1,1).setValue(utilB_Last)
  if (utilB_Type != ""){entrySheet.getRange(5,6,1,1).setValue(utilB_Last+1)}
  makePickList(5)

  utilC_Name =  propSheetData[47][colIndex1-1]
  entrySheet.getRange(6,3,1,1).setValue(utilC_Name)
  utilC_Type =  propSheetData[48][colIndex1-1]
  entrySheet.getRange(6,4,1,1).setValue(utilC_Type)
  utilC_Last =  propSheetData[50][colIndex1-1]
  entrySheet.getRange(6,5,1,1).setValue(utilC_Last)
  if (utilC_Type != ""){entrySheet.getRange(6,6,1,1).setValue(utilC_Last+1)}
//  chargedUtil = entrySheet.getRange(4,9,3,1).getDisplayValues()
  makePickList(6)
}

function getLastInvNum() {
  getPropSheet()
  colIndex1 = getPropSheetLastDataCol()
  if (colIndex1 == 1) {
    entrySheet.getRange(2,2,1,1).setValue("None")
     entrySheet.getRange(2,5,1,1).setValue("None")
     entrySheet.getRange(2,7,1,1).setValue("None")
    return
  }
  lastInvDate = propSheetData[19][colIndex1-1]
  entrySheet.getRange(1,2,1,1).setValue(lastInvDate)
  lastInvNumber = propSheetData[22][colIndex1-1]
  entrySheet.getRange(1,5,1,1).setValue(lastInvNumber)
  lastInvTotal = propSheetData[24][colIndex1-1]
  entrySheet.getRange(1,7,1,1).setValue(lastInvTotal)
  loadInvAtBottom(colIndex1-1)
}

function getDeposits(){
  getGlobalSheet()
  var tenantCode = entrySheetData[1][3] //next new column
  var tenantData = findRowInSheet(glData,tenantCode,1)  // get invoice end digit setup
  var iRow = tenantData[0] + 1
  var depReqSheet = glSheet.getRange(iRow,12,1,1).getValue()
  var depPaid =  glSheet.getRange(iRow,13,1,1).getValue()
  myLog("Deposit Req " + depReqSheet + "Deposit Paid" +  depPaid)
  entrySheet.getRange(1,9,1,1).setValue(depReqSheet)
  entrySheet.getRange(2,9,1,1).setValue(depPaid)
  
}

function editEntry(){
  getEntrySheet()
  chargedUtil = entrySheet.getRange(4,9,3,1).getDisplayValues()
  buildMainPickList() //MakeEntrySheet
  makeAllPickList() //BaseToEntry
  fixUtilValues()
  menuOne()
 
}

//-------------------------------------------------------------------

function getPickInvoice(){
  getPropSheet()
  invPicked = entrySheet.getRange(14,9,1,1).getValue()
  if (invPicked == ""){
    myAlert("No invoice item picked")
    return
  }
  theIndex = getInvIdexOf(invPicked)
  myLog(invPicked)
  theInvIs = getInvIdexOf(invPicked)
  clearDataEntry(false)
  fromBaseToEntryData(theInvIs)
//  fixUtilValues()
  getInvNumbers()
  writeSaved(true)
  menuTwo()
}


function getInvIdexOf(invPicked){
  getPropSheet()
  maxCol = propSheet.getMaxColumns()
  invAt = 22
  for(var i = 0; i<maxCol+1;i++)
  {
    invNum = propSheetData[invAt][i]
    if(invNum == invPicked)    // 0 - row 0
    {
      return(i)
    }
  }
  myAlert("Invoice number " + invNum +  " not found")
  return(0)
}

function getINumbers()
{
  var maxCol = propSheet.getMaxColumns()
  myLog(maxCol)
  var inv1stItem = ""
  var invAt = 22
  var i=0
var options = {
   formatOnly: false,
   contentsOnly: true,
   validationsOnly : true
  };

  try{
    myLog(propSheetData[invAt][i])
    }
    catch (e) {
      entrySheet.getRange(14,9,1,1).clear(options)
      entrySheet.getRange(14,9,1,1).setValue("No Old Data")
      //myAlert("Missing data or corrupt data")
      return
    }
  myLog(propSheetData[invAt][i])
  var invNumbers = []
  for(var i = 0; i<maxCol+1;i++)
  {
    var invNum = propSheetData[invAt][i]
    if(invNum != null)    // 0 - row 0
    {
      if (i == 0){
        inv1stItem = invNum
      }
      else
      {
        invNumbers.push(invNum)
      }
    }
  }
  invNumbers.push(inv1stItem)
  if (invNumbers.length > 0){
    invNumbers.reverse()
  }
  return invNumbers
}

function getInvNumbers(){
  var invList = []
  getPropSheet()
  invList = getINumbers()
  entrySheet.getRange(14,9,1,1).clearContent()
  if (invList.length > 0){
    arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Choose from List Only").requireValueInList(invList).build()
    entrySheet.getRange(14,9,1,1).setDataValidation(arule)
  }
  entrySheet.getRange(14,9,1,1).setValue(invList[0])
}




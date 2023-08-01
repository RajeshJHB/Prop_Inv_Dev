/// ---- Make Invoice ------ RS
// Write to invoice - genInvoice
// Write to base - writeToTBase
// Read Last - getLastforDataEntry
//-------- Java Script does not return 2 items ---------

function writeToInvoice(){
  printInv = true
  clearInvoice()
  var newCol = makeInvNumber()
  setPropTransDataInv() // Had to move before FixData so that The Deposits display
  getSetPropFixData()
  packPutFullEntryData(newCol)
  clearDataEntry(true)
  switchToInvoice()
  printInv = false
}

function printInvoice(){
  clearInvoice()
  setPropTransDataInv() // Had to move before FixData so that The Deposits display
  getSetPropFixData()
  clearDataEntry(true)
  switchToInvoice()
}


function switchToInvoice(){
  ss = SpreadsheetApp.getActiveSpreadsheet();
  var invSheet = ss.setActiveSheet(ss.getSheetByName("Invoice"))
  var range = invSheet.getRange(2,2,39,3);  //row,col,num of rows, num of columns
  invSheet.setActiveRange(range);
}


function makeInvNumber () // according to data entry property
{
  var prop = getPropSheet()
  var currInv = entrySheet.getRange(3,5,1,1).getValue()
  if (currInv != "Temp"){
    myLog("No new Invoice number")
    var colNum = findInvIndex()
    return(colNum)
  }
  getGlobalSheet()
  var newInvNum = 0
  var lastInvNumber = 0
  var lastColIndex = getPropSheetLastDataCol()
  myLog("lastColIndex " + lastColIndex)
  var propData = findRowInSheet(glData,prop,1)  // get invoice end digit setup
  var endDigits = propData[12]  // from the global setup
  myLog('endDigits ' + endDigits)
  if (lastColIndex <= 1) {    // Empty data base
    newInvNum = 100 + endDigits
  }
  else if (lastColIndex == 2) { // have only one data record
    lastInvNumber = propSheetData[22][lastColIndex-1]
    if (lastInvNumber == "Temp") {
      newInvNum = 100 + endDigits
    }
    else {
      myLog("last inv number "+ lastInvNumber)
      newInvNum = parseInt(lastInvNumber/100) + 1
      newInvNum = (newInvNum*100) + endDigits
      lastColIndex = lastColIndex + 1
    }
  }
  else  // have more than 1 record
  {
    lastInvNumber = propSheetData[22][lastColIndex-1]
    if (lastInvNumber == "Temp"){
      lastInvNumber = propSheetData[22][lastColIndex-2]
    }
    else {
      lastColIndex = lastColIndex + 1
    }
    myLog("last inv number "+ lastInvNumber)
    newInvNum = parseInt(lastInvNumber/100) + 1
    newInvNum = (newInvNum*100) + endDigits
  }
//  myRes = myYNPrompt("Create new 'YES" +  newInvNum +  "Update " + lastInvNumber)
  myLog("new inv number "+ newInvNum)
//  invoiceSheet.getRange(4,5,1,1).setValue(newInv)  // Se the new invoice number
  entrySheet.getRange(3,5,1,1).setValue(newInvNum)
  return(lastColIndex)
}

function genInvoice(){
  theProp = getPropSheet()
  glData = ss.getSheetByName("Globals").getDataRange().getValues();
  invoiceSheet = ss.getSheetByName("Invoice")
  getInvInfo(theProp)
}


function getInvInfo(thisProp){
      ClearInvoice()
      colIndex = setPropTransDataInv()
      myLog("Overwrite Index @ 3 " + colIndex)
      getSetPropFixData(thisProp) // landloard name & bank Name & Property details to invoice
      myLog("Overwrite Index @ 4 " + colIndex)
      genEntryToPropBase(colIndex)
    // Set all transactions to invoice
    
}

function makeAddDeposit(depValue){
  getGlobalSheet()
  tenantCode = entrySheetData[1][3] //next new column
  myLog("Tenant Code  " + tenantCode)
  tenantData = findRowInSheet(glData,tenantCode,1)  // get invoice end digit setup
  myArLog(tenantData)
  iRow = tenantData[0] + 1
  myLog("iRow " + iRow)
  depInSheet = glSheet.getRange(iRow,12,1,1).getValue()
  myLog("Deposit in Sheet  " + depInSheet)
  if (depInSheet > 0){
    var resp = myYNPrompt("Adding " + depValue + " Current Deposit Required : " + depInSheet)
    if (resp == "YES"){
      depValue = depValue + depInSheet    // Update deposit amount
    }
    else
    {
      depValue = depInSheet
    }
  }
  entrySheet.getRange(1,9,1,1).setValue(depValue)
  glSheet.getRange(iRow,12,1,1).setValue(depValue)
}

function makeDepositPayment(depValue){
  getGlobalSheet()
  tenantCode = entrySheetData[1][3] //next new column
  myLog("Tenant Code  " + tenantCode)
  tenantData = findRowInSheet(glData,tenantCode,1)  // get invoice end digit setup
  myArLog(tenantData)
  iRow = tenantData[0] + 1
  myLog("iRow " + iRow)
  depInSheet = glSheet.getRange(iRow,13,1,1).getValue()
  myLog("Deposit in Sheet  " + depInSheet)
  if (depInSheet > 0){
    var resp = myYNPrompt("Adding " + depValue + " to Current Deposit Held : " + depInSheet)
      if (resp == "YES"){
      depValue = depValue + depInSheet    // Update deposit amount
    }
    else
    {
      depValue = depInSheet
    }
  }
  entrySheet.getRange(2,9,1,1).setValue(depValue)
  glSheet.getRange(iRow,13,1,1).setValue(depValue)
}


function getSetPropFixData(){
  var prop = getPropSheet()
  invoiceSheet = ss.getSheetByName("Invoice")
  glData = ss.getSheetByName("Globals").getDataRange().getValues();
  invoiceSheet.getRange(37,3,1,1).setValue(prop)

  var propData = findRowInSheet(glData,prop,1)
  getSetPropAddr(propData)

  var landLord = propData[10]
  getSetLandLordDetails(landLord)
  entrySheet.getRange(13,8,1,1).setValue(landLord)

  depBank = propData[11]
  getSetBankDetails (depBank)
  entrySheet.getRange(14,8,1,1).setValue(depBank)

  theTenant = propData[2]
  getSetTenant(theTenant)

  invoiceSheet.getRange(38,3,1,1).setValue(prop)  // Bank referance

   // getDataEntryLastInfo() - for now - either on RHS Pick List or event driveen for old data

  myLog("Property Data " + propData)
  myLog("The tenant is  "+  theTenant)
  myLog("The landlord is "+  landLord)
  myLog("The Bank is "+  depBank)
//    newInv = makeInvNumber(prop)
//    invoiceSheet.getRange(4,5,1,1).setValue(newInv)  // Se the new invoice number
//    entrySheet.getRange(3,5,1,1).setValue(newInv)
}

function getSetLandLordDetails(landLord){
  var landLordData = findRowInSheet(glData,landLord,15)
  var landLordData = landLordData.slice(16)
//  myArLog("Landloard Data "+landLordData)
  var landLordData2D = onetoTwod(landLordData,1)
//  myArLog(landLordData2D)
  invoiceSheet.getRange(2,2,6,1).setValues(landLordData2D)
}

function getSetBankDetails (depBank) {
  var depBankData = findRowInSheet(glData,depBank,15)
  myLog ("DepBank Data 1 "+ depBankData)
  var depBankDataS = depBankData.slice(16,20)
  myLog("Bank Data" + depBankDataS)
  var depBankData2D = onetoTwod(depBankDataS,1)
 // myArLog(depBankData2D)
  invoiceSheet.getRange(34,3,4,1).setValues(depBankData2D)
}

function getSetPropAddr(propData) {
  var propAddr = propData.slice(4,9)
  myLog ("Prop Address  " + propAddr)
  var propAddr2D = onetoTwod(propAddr,1)
 // myArLog(propAddr2D)
  invoiceSheet.getRange(9,4,5,1).setValues(propAddr2D)
//  invoiceSheet.getRange(38,3,1,1).setValue(prop)
}

function getSetTenant(tData){
  myLog ("Find this Tenant "+tData)
  var tenant = findRowInSheet(glData,tData,1)
  var rentalDeposit = tenant[11]
  var rentalDepositPaid = tenant[12]
  myLog("Tenant Deposit "+ rentalDeposit)
  myLog("Tenant Data "+tenant)

  var tenantS = tenant.slice(2,9)
  myLog("Tenant Data Sliced "+tenantS)

  var tenant2D = onetoTwod(tenantS,1)
//  myArLog("Tenant Data 2D "+tenant2D)

  invoiceSheet.getRange(8,3,7,1).setValues(tenant2D)

  invoiceSheet.getRange(15,5,1,1).setValue(rentalDepositPaid)   // todo ...credit when payment made
  invoiceSheet.getRange(16,5,1,1).setValue(rentalDeposit)
}


function setPropDataToInv(){
  theProp = getPropSheet()
  invoiceSheet = ss.getSheetByName("Invoice")
  myArLog(propSheetData)

}

function setPropTransDataInv(){
  var theProp = getPropSheet()
  invoiceSheet = ss.getSheetByName("Invoice")
  var invoiceTotal = 0
  var invDate = entrySheetData[2][1]
  if (invDate == "")
  {
      invDate = new Date()
      invDate.setHours(0,0,0,0)
  }
  entrySheet.getRange(3,2,1,1).setValue(invDate)
  invoiceSheet.getRange(5,5,1,1).setValue(invDate)  // set date of invoice
  var invMonthDate = entrySheetData[2][2]
  if (invMonthDate == "")
  {
    invMonthDate = invDate
  }
  entrySheet.getRange(3,3,1,1).setValue(invMonthDate)
  invoiceSheet.getRange(6,5,1,1).setValue(invMonthDate)  // set invoice date - for which Month and Year
  invoiceSheet.getRange(18,3,1,1).setValue("Balance Brought Forward")
  invDate = entrySheetData[0][6] // Balance B/F
  invoiceSheet.getRange(18,5,1,1).setValue(invDate)  // set invoice date - for which Month and Year

  var invNum = entrySheetData[2][4] // Invoice Number
  invoiceSheet.getRange(4,5,1,1).setValue(invNum)  // Se the new invoice number

  var nextReadRow = 3 // 0 Based
  var nextWriteRow = 19 // 1 Based
  for (var i = 0; i <= 2; i++ )  // 3 rows
  {
    utilType = entrySheetData[nextReadRow+i][2]
    if ((utilType == "None") || (utilType == "")) {
      
    }
    else
    {
      var utilChargeType = entrySheetData[nextReadRow+i][3]
      myLog (utilType + " ->  "  + utilChargeType)
      var utilDate = entrySheetData[nextReadRow+i][1]
      var utilStartReading = entrySheetData[nextReadRow+i][4]
      var utilEndReading = entrySheetData[nextReadRow+i][5]
      var utilConsumed = entrySheetData[nextReadRow+i][6]
      utilConsumed = myFormat(utilConsumed,2)
     // var utilCalcCharge = entrySheetData[nextReadRow+i][7]
      var utilActCharged = entrySheetData[nextReadRow+i][8]
      myLog("Actual Charge :" + utilActCharged)
      invoiceSheet.getRange(nextWriteRow,2,1,1).setValue(utilDate)   // write Date
      if (utilChargeType == "PostPaid"){
        invoiceSheet.getRange(nextWriteRow,3,1,1).setValue(utilType +" from : " + utilStartReading + " to " + utilEndReading  + " Total Units "+ utilConsumed)
      }
      else if (utilChargeType == "PrePaid"){
        invoiceSheet.getRange(nextWriteRow,3,1,1).setValue(utilType + "->"+ utilChargeType)
      }
      else if (utilChargeType == "Fix"){
        invoiceSheet.getRange(nextWriteRow,3,1,1).setValue(utilType + "->"+ utilChargeType)
      }
      else if (utilChargeType == "New"){
        invoiceSheet.getRange(nextWriteRow,3,1,1).setValue(utilType + "->"+ utilChargeType + " Start Reading : " + utilEndReading ) // previous end reading
      }
      invoiceSheet.getRange(nextWriteRow,5,1,1).setValue(utilActCharged)
      nextWriteRow +=1
      invoiceTotal += utilActCharged
    }
  }
  myLog("Next write Row " + nextWriteRow)

  nextReadRow = 6   // 0 based

  for (var i = 0; i <= 7; i++ )  // 8 rows
  {
    var transType = entrySheetData[nextReadRow+i][2] // Debit , Credi , note or None
   // myLog("Next write Row " + nextWriteRow + "Trans Type " + transType)
    if ((transType == "None") || (transType == "")) {
      
    }
    else
    {
      var transDate = entrySheetData[nextReadRow+i][1] // Read Date
      invoiceSheet.getRange(nextWriteRow,2,1,1).setValue(transDate) // Write Date
      var transNote = entrySheetData[nextReadRow+i][5]  // Note text
      var transSubType = entrySheetData[nextReadRow+i][3] // Tye of transaction, rent, payment etc
      var transValue = entrySheetData[nextReadRow+i][4] // Value of above transaction
      if (transType == "Debit (+)") {
        invoiceSheet.getRange(nextWriteRow,3,1,1).setValue(transSubType + " " + transNote )
        invoiceSheet.getRange(nextWriteRow,5,1,1).setValue(transValue)
        if (transSubType == "Deposit Reqd"){
          makeAddDeposit(transValue)
        }
      }
      else if (transType == "Credit (-)") {
        invoiceSheet.getRange(nextWriteRow,3,1,1).setValue(transSubType + " " + transNote)
        invoiceSheet.getRange(nextWriteRow,5,1,1).setValue("-"+transValue)    // Show a minus sign on the invoice
        if (transSubType == "Deposit Paid"){
          makeDepositPayment(transValue)
        }
      }
      else if (transType== "Note") {
        invoiceSheet.getRange(nextWriteRow,3,1,1).setValue(transNote)
        invoiceSheet.getRange(nextWriteRow,5,1,1).setValue(transValue)
      }
      nextWriteRow +=1
    }
  }
  invoiceTotal = entrySheetData[2][6] // Value of above transaction
  invoiceSheet.getRange(30,5,1,1).setValue(invoiceTotal)
  invoiceSheet.getRange(39,3,1,1).setValue(invoiceTotal)
}

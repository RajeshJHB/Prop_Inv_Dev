/// ---- EntryToBase ------ RS

function justSaveNClear(){
  var theCol = findInvIndex()
  packPutFullEntryData(theCol)
  clearDataEntry(true)
}


function packPutFullEntryData(wrtIndex){
  getPropSheet()
  var colIndex = getPropSheetLastDataCol()
  myLog("colIndex : " + colIndex)
  myLog("wrtIndex : " + wrtIndex)
  if (colIndex < 2){
    initializeTable() // that writes to column 1 with headings
    wrtIndex = 2  //data now in colum 2
  }
  /*
  if (colIndex > 1) {
    currInvNumber = entrySheetData[2][4]
    storedInvNum = propSheetData[22][colIndex-1]
    if (currInvNumber == storedInvNum) {
      result = myYNPrompt("Overwrite Current Data - CurrInv " + currInvNumber + "Stored Inv " + storedInvNum)
      colIndex -= 1; // Overwrite index
      if (result == "NO") {return}
    } else if(storedInvNum == "Temp"){
      colIndex -= 1;
    }
  }
*/
  if ((colIndex != 0) && (colIndex == wrtIndex) && (!printInv)) {
      var currInvNumber = entrySheetData[2][4]
      result = myYNPrompt("Overwrite Current Data " + currInvNumber)
      if (result == "NO") {return}
  }

  var dEntry =[]
  var dEntryIndex = 0
  for (var i =0 ; i <=13; i++ ){    // Row 6 to 8
    for (var j= 0 ; j<=8; j++) {
          dEntry[dEntryIndex] = entrySheetData[i][j]
          dEntryIndex ++
    }
  }
  myArLog(dEntry)
  var dEntry2D = onetoTwod(dEntry,1)
  propSheet.getRange(1,wrtIndex,dEntryIndex,1).setValues(dEntry2D)
}


function findInvIndex(){
  getPropSheet()
  var colIndex = getPropSheetLastDataCol()
  if (colIndex == 0){ // no data in sheet to get last invoice
    return(1)
  }
  var dispInvNumber = entrySheetData[2][4]
  var theCol = findColInSheet(propSheetData,dispInvNumber,22) // row 22 (base 0) is the invoice numbers
  if (theCol == null){
    myLog("Temp on screen only")
    theCol = colIndex + 1
  }
  myLog("colIndex : " + colIndex)
  myLog(theCol)
  return(theCol)

  /*
  if (dispInvNumber == "Temp"){ // saving over temp
    theCol = colIndex
  }
  else
  {
    theCol = findColInSheet(propSheetData,dispInvNumber,22) // row 22 (base 0) is the invoice numbers
  }
  */
}

function initializeTable(){
  var dEntry = []
  dEntry = ["Last Inv Date","last inv date","saved status","Last Inv Number","last inv number","Last Inv Balance","last inv balance","Deposit Required","deposit required",
            "Pick Property","property code","Current Tenant","tenant code","tenant name","tenant number","Blank","Deposit Paid","deposit paid",
            "Invoice Date","invoice date","Blank","Invoice Number","invoice number","Invoice Total","invoice total","Blank","Blank",
            "Utility A","date of reading ","utility A name","utility A cost type","utility A start reading","utility A end reading", "utility A units used", "utility A calculated cost", "utility A cost charged",
            "Utility B","date of reading ","utility B name","utility B cost type","utility B start reading","utility B end reading",  "utility B units used", "utility B calculated cost", "utility B cost charged",
            "Utility C","date of reading ","utility C name","utility C cost type","utility C start reading","utility C end reading",  "utility C units used","utility C calculated cost", "utility C cost charged",
            "Trans 1","date of transaction","trans1 type","trans1 sub type","trans1 cost","trans1 note","Blank","Blank","Blank",
            "Trans 2","date of transaction","trans2 type","trans2 sub type","trans2 cost","trans2 note","Blank","Blank","Blank",
            "Trans 3","date of transaction","trans3 type","trans3 sub type","trans3 cost","trans3 note","Blank","Blank","Blank",
            "Trans 4","date of transaction","trans4 type","trans4 sub type","trans4 cost","trans4 note","Blank","Blank","Blank",
            "Trans 5","date of transaction","trans5 type","trans5 sub type","trans5 cost","trans5 note","Blank","Blank","Blank",
            "Trans 6","date of transaction","trans6 type","trans6 sub type","trans6 cost","trans6 note","Blank","Blank","Blank",
            "Trans 7","date of transaction","trans7 type","trans7 sub type","trans7 cost","trans7 note","Blank","Blank","Blank",
            "Trans 8","date of transaction","trans8 type","trans8 sub type","trans8 cost","trans8 note","Blank","Blank","Blank",
            ]
  dEntry2D = onetoTwod(dEntry,1)
  proCode = getPropSheet()
  propSheet.getRange(1,1,126,1).setValues(dEntry2D)

}

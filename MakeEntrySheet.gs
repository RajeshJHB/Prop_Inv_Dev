/// ---- MakeEntrySheet ------ RS

function makeEntryDataSheet(){
  ss = SpreadsheetApp.getActiveSpreadsheet();
  newEntrySheet = ss.getSheetByName("DataEntry")
  if (newEntrySheet == null){
    ss.insertSheet("DataEntry")
    newEntrySheet = ss.getSheetByName("DataEntry")
  }
  else
  {
     SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('Data Entry Sheets Exists already, no action performed');
    return
  }
  maxCol = newEntrySheet.getMaxColumns()
  if (maxCol <10){
    newEntrySheet.insertColumnsAfter(1,10-maxCol)
  }
  glSheet = ss.getSheetByName("Globals")
  prRange = glSheet.getRange('B2:B32')
  // Wite static info
  newEntrySheet.getRange('A1:I14').setValues([
  ["Last Inv Date","","","Last Inv Number","", "Last Inv Balance","","Deposit Reqd",""],
  ["Pick Property","","Current Tenant","","","","","Deposit Paid",""],
  ["Invoice Date","","","Invoice Number","","Invoice Total","","Calculated Util","Charged Util"],
  ["Date","","","","","","","",""],
  ["Date","","","","","","","",""],
  ["Date","","","","","","","",""],
  ["Date","","","","","","","",""],
  ["Date","","","","","","","",""],
  ["Date","","","","","","","",""],
  ["Date","","","","","","","",""],
  ["Date","","","","","","","",""],
  ["Date","","","","","","","",""],
  ["Date","","","","","","","","Pick Old Invoice"],
  ["Date","","","","","","","",""]
  ])

  // Write utils colours
  newEntrySheet.getRange('C18:C19').setValues([["Utility - Electricity/Water/Gas"],["Debits/Credits"]])
 
 // Legend cell colours
  newEntrySheet.getRange(['C18:D18']).setBackground([["#ffe599"]])  // Set beige colour
  newEntrySheet.getRange(['C19:D19']).setBackground([["#a2c4c9"]])  // Set grey colour

  // Write last inv cell colours
  newEntrySheet.getRangeList(['B1','E1','G1','I1','I2']).setBackground([["#ffff00"]])  // Set yellow colour
  
  newEntrySheet.getRange('B3:C3').setBackground([["#00ff00"]])    // Set Green colour enter invoice date
  newEntrySheet.getRangeList(['E3','G3']).setBackground([["#f19292"]])  // Set Pink Red for invoice number and Total

  
  newEntrySheet.getRange(['A4:G6']).setBackground([["#ffe599"]])  // Set beige colour - Util Entry
  newEntrySheet.getRange(['A7:F14']).setBackground([["#a2c4c9"]])  // Set grey colour - Debit/Credit/ Note entry
  newEntrySheet.getRange(['H3:H6']).setBackground([["#d7e38a"]])  // Calc Util
  newEntrySheet.getRange(['I3:I6']).setBackground([["#befa99"]]) //Charged Util



  arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Pick a Date").requireDate().build()  // Set validation for all dates
  newEntrySheet.getRange('B1').setDataValidation(arule)
  newEntrySheet.getRange('B3:B14').setDataValidation(arule)
  newEntrySheet.getRangeList(['B3:B14','B1']).setNumberFormat("dd-mm-yyyy")
  newEntrySheet.getRange('C3').setDataValidation(arule)
  newEntrySheet.getRange("C3").setNumberFormat('MMM-yyyy');

  newEntrySheet.getRangeList(["G1","G3","E4:E14","H4:I6"]).setNumberFormat("0.00");
 
  
  arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Choose from Menu 1 Only").requireValueInRange(prRange).build() // from items in the sheet
  newEntrySheet.getRange('B2').setDataValidation(arule) // Choose property

  newEntrySheet.getRange('D2').setValue("=index((Globals!C2:C40),match(B2,Globals!B2:B40,0),)") // Display Current tenat code
  newEntrySheet.getRange('E2').setValue("=INDEX(Globals!C53:C100,MATCH(D2,Globals!B53:B100))") // Display Current Tenant name

//  arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireNumberLessThan(10000000).build();
//  newEntrySheet.getRange("E4:E14").setDataValidation(arule)
//  newEntrySheet.getRange("F4:I6").setDataValidation(arule)

  buildMainPickList()
/*
  utilList = ["Electricity","Water","Gas","None"]   // Make Pick List for utils
  arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Choose from Menu 2 Only").requireValueInList(utilList).build()
  newEntrySheet.getRange('C4:C6').setDataValidation(arule)

  accList =["Debit (+)","Credit (-)", "Note","None"]  // Make pick list for Debits /Credits
  arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Choose from Menu 3 Only").requireValueInList(accList).build()
  newEntrySheet.getRange('C7:C14').setDataValidation(arule)
*/

  // Set all notes in cell
  newEntrySheet.getRange("C3").setNote("Invoice For Which Month")
  newEntrySheet.getRange("E4:E6").setNote("Utility Start Value")
  newEntrySheet.getRange("F4:F6").setNote("Utility End Value")
  newEntrySheet.getRange("G4:G6").setNote("Utility Consumed")
  newEntrySheet.getRange("H4:H6").setNote("Utility Calculated Cost")
  newEntrySheet.getRange("I4:I6").setNote("Utility Charged")
  newEntrySheet.getRange("E7:E14").setNote("Transaction Value" )
  newEntrySheet.getRange("F7:F14").setNote("Text Note" )
  newEntrySheet.getRangeList(["C4:C14"]).setValue("None")
 // makeTotal ()
 }

function buildMainPickList()
{
  ss = SpreadsheetApp.getActiveSpreadsheet();
  var newEntrySheet = ss.getSheetByName("DataEntry")
  var utilList = ["Electricity","Water","Gas","None"]   // Make Pick List for utils
  var arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Choose from Menu 2 Only").requireValueInList(utilList).build()
  newEntrySheet.getRange('C4:C6').setDataValidation(arule)

  var accList =["Debit (+)","Credit (-)", "Note","None"]  // Make pick list for Debits /Credits
  arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Choose from Menu 3 Only").requireValueInList(accList).build()
  newEntrySheet.getRange('C7:C14').setDataValidation(arule)

}
/*
function makeTotal(){
  getEntrySheet()
 // myLog("G3 is " + entrySheet.getRange(3,7,1,1).getValue())
  entrySheet.getRange("G3").setNote("Invoice Total For Month")//Calculate Invoice total formulae
//  entrySheet.getRange("M4:M14").setValue("1")
  entrySheet.getRange("N4:N14").setValues([["=I4*M4"],["=I5*M5"],["=I6*M6"],["=E7*M7"],["=E8*M8"],["=E9*M9"],["=E10*M10"],["=E11*M11"],["=E12*M12"],["=E13*M13"],["=E14*M14"]])
  entrySheet.getRange("G3").setValue("=SUM(N4:N14)+G1")
  entrySheet.hideColumns(13,2)
}

*/

function calcTotal(){
  getEntrySheet()
  c1Stat = entrySheet.getRange("C1").getValue()
  if (c1Stat != "Not Saved"){
    writeSaved(false)
  }
  invTotal = entrySheetData[0][6]
  myArLog(entrySheetData)
  for (var i = 4; i <= 6; i++){
    uVal = entrySheet.getRange(i,9,1,1).getValue()
   myLog("Util Value " +uVal)
    if (uVal != ""){
      invTotal += uVal
    }
   myLog("Row number " +invTotal)
  }
  
  depPlus = 0   //Deposit Required
  depMinus = 0 //Deposit paid
  myLog(invTotal)
  for (var i = 6; i <= 13; i++) {
    dcType =  entrySheetData[i][2]
    dcSubType = entrySheetData[i][3]
    dcVal = entrySheetData[i][4]
    if (dcType == "Debit (+)"){
      invTotal = dcVal + invTotal
    } else if (dcType == "Credit (-)"){
      invTotal = invTotal - dcVal
    } else if (dcType == "Note"){
      invTotal = dcVal + invTotal
    }

    if (dcSubType == "Deposit Reqd"){
       depPlus +=  dcVal
    }
    if (dcSubType == "Deposit Paid"){
      depMinus +=  dcVal
      myLog("Dep minus " + depMinus)
    }
    if (dcSubType == "Deposit Refund"){
      depMinus +=  dcVal
    }
    myLog(invTotal)
  }
//  if ((depPlus == 0) && (depMinus > 0)) {
//       invTotal +=  depMinus
//  }
  
  entrySheet.getRange("G3").setValue(invTotal)
}

function testDepPaid(){
  checkIfDepPaid(8)
}

function checkIfDepPaid(theRow){
  getEntrySheetOnce()
  bCell = entrySheet.getRange(theRow,4,1,1).getValue()
  if (bCell == "Deposit Paid"){
    entrySheet.getRange(theRow,13,1,1).setValue("0") // Dont count deposits paid in invoice total update deposit
  }
}

// All the live stuff below Dynamic piclist and clearing

function makePickList(theRow){
  crList = ["Payments","Discount","Deposit Paid","Deposit Refund","Other:"]
  drList = ["Rent","Fibre", "Rates","Service Fee", "Interest Charge","Penalty","Deposit Reqd","Other:" ]
  noteList = ["Type Note /Message","Other:"]
  utilPayList = ["PostPaid","PrePaid","Fix","New","Other:"]


  getEntrySheetOnce()
  
  var options = {
   formatOnly: false,
   contentsOnly: true,
   validationsOnly : true
  };

  aCell = entrySheet.getRange(theRow,3,1,1).getValue()
//  if(entrySheet.getRange(3,7,1,1).getValue() == ""){
//   makeTotal()
//  }
  if (aCell == "Debit (+)"){
    subListVal = entrySheet.getRange(theRow,4,1,1).getValue()
    arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Choose from Menu 4 Only").requireValueInList(drList).build()
    entrySheet.getRange(theRow,4,1,1).setDataValidation(arule)
//    entrySheet.getRange(theRow,13,1,1).setValue("1")
  }
  else if (aCell == "Credit (-)"){
    subListVal = entrySheet.getRange(theRow,4,1,1).getValue()
    arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Choose from Menu 5 Only").requireValueInList(crList).build()
    entrySheet.getRange(theRow,4,1,1).setDataValidation(arule)
 //   entrySheet.getRange(theRow,13,1,1).setValue("-1")
  }
  else if (aCell == "Note"){
    subListVal = entrySheet.getRange(theRow,4,1,1).getValue()
    arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Choose from Menu 6 Only").requireValueInList(noteList).build()
    entrySheet.getRange(theRow,4,1,1).setDataValidation(arule)
  //  entrySheet.getRange(theRow,13,1,1).setValue("1")
  }
  else if (aCell == "None"){
    if (theRow < 7) {
      entrySheet.getRange(theRow,4,1,6).clear(options)
    }else
      entrySheet.getRange(theRow,4,1,3).clear(options)
  }
  else if  ((aCell == "Electricity") || (aCell == "Water") || (aCell == "Gas")) {
    arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Choose from Menu 7 Only").requireValueInList(utilPayList).build()
    entrySheet.getRange(theRow,4,1,1).setDataValidation(arule)
    entrySheet.getRange(theRow,7,1,1).setValue("=(F"+theRow+"-E"+theRow+")") // Util used formulae
    entrySheet.getRange(theRow,9,1,1).setValue('=if(D'+theRow+'="PrePaid",0,H'+theRow+')')    // Make a copy of cost that can be changed
   // Insert relavant formulea in appropriate cell
    if (aCell == "Electricity") {
      entrySheet.getRange(theRow,8,1,1).setValue("=calcElec(G"+theRow+")")

    } else if (aCell == "Water") {
      entrySheet.getRange(theRow,8,1,1).setValue("=calcWater(G"+theRow+")")
    } else if (aCell == "Gas") {
         entrySheet.getRange(theRow,8,1,1).setValue("=calcGas(G"+theRow+")")
    }
  }

}

//      entrySheet.getRange(theRow,8,1,1).setValue('=if(D'+theRow+'="PrePaid",0,calcElec(G'+theRow+'))') // this works !!, but not needed
//.                                               ('=if(D'+theRow+'="PrePaid",0,'H'+theRow+')')

/// ---- MakeEntrySheet ------ RS

function makeEntryDataSheet(){
  ss = SpreadsheetApp.getActiveSpreadsheet();
  newEntrySheet = ss.getSheetByName("DataEntry")
  if (newEntrySheet == null){
    ss.insertSheet("DataEntry")
    newEntrySheet = ss.getSheetByName("DataEntry")
  }
  maxCol = newEntrySheet.getMaxColumns()
  if (maxCol <10){
    newEntrySheet.insertColumnsAfter(1,10-maxCol)
  }
  glSheet = ss.getSheetByName("Globals")
  prRange = glSheet.getRange('B2:B32')
  // Wite static info
  newEntrySheet.getRange('A1:I14').setValues([
  ["Last Inv Date","","","Last Inv Number","", "Last Inv Balance","","Deposit Reqd",""],
  ["Pick Property","","Current Tenant","","","","","Deposit Paid",""],
  ["Invoice Date","","","Invoice Number","","Invoice Total","","Calculated Util","Charged Util"],
  ["Date","","","","","","","",""],
  ["Date","","","","","","","",""],
  ["Date","","","","","","","",""],
  ["Date","","","","","","","",""],
  ["Date","","","","","","","",""],
  ["Date","","","","","","","",""],
  ["Date","","","","","","","",""],
  ["Date","","","","","","","",""],
  ["Date","","","","","","","",""],
  ["Date","","","","","","","","Pick Old Invoice"],
  ["Date","","","","","","","",""]
  ])

  // Write utils colours
  newEntrySheet.getRange('C18:C19').setValues([["Utility - Electricity/Water/Gas"],["Debits/Credits"]])
 
 // Legend cell colours
  newEntrySheet.getRange(['C18:D18']).setBackground([["#ffe599"]])  // Set beige colour
  newEntrySheet.getRange(['C19:D19']).setBackground([["#a2c4c9"]])  // Set grey colour

  // Write last inv cell colours
  newEntrySheet.getRangeList(['B1','E1','G1','I1','I2']).setBackground([["#ffff00"]])  // Set yellow colour
  
  newEntrySheet.getRange('B3:C3').setBackground([["#00ff00"]])    // Set Green colour enter invoice date
  newEntrySheet.getRangeList(['E3','G3']).setBackground([["#f19292"]])  // Set Pink Red for invoice number and Total

  
  newEntrySheet.getRange(['A4:G6']).setBackground([["#ffe599"]])  // Set beige colour - Util Entry
  newEntrySheet.getRange(['A7:F14']).setBackground([["#a2c4c9"]])  // Set grey colour - Debit/Credit/ Note entry
  newEntrySheet.getRange(['H3:H6']).setBackground([["#d7e38a"]])  // Calc Util
  newEntrySheet.getRange(['I3:I6']).setBackground([["#befa99"]]) //Charged Util



  arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Pick a Date").requireDate().build()  // Set validation for all dates
  newEntrySheet.getRange('B1').setDataValidation(arule)
  newEntrySheet.getRange('B3:B14').setDataValidation(arule)
  newEntrySheet.getRangeList(['B3:B14','B1']).setNumberFormat("dd-mm-yyyy")
  newEntrySheet.getRange('C3').setDataValidation(arule)
  newEntrySheet.getRange("C3").setNumberFormat('MMM-yyyy');

  newEntrySheet.getRangeList(["G1","G3","E4:E14","H4:I6"]).setNumberFormat("0.00");
 
  
  arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Choose from Menu 1 Only").requireValueInRange(prRange).build() // from items in the sheet
  newEntrySheet.getRange('B2').setDataValidation(arule) // Choose property

  newEntrySheet.getRange('D2').setValue("=index((Globals!C2:C40),match(B2,Globals!B2:B40,0),)") // Display Current tenat code
  newEntrySheet.getRange('E2').setValue("=INDEX(Globals!C53:C100,MATCH(D2,Globals!B53:B100))") // Display Current Tenant name

//  arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireNumberLessThan(10000000).build();
//  newEntrySheet.getRange("E4:E14").setDataValidation(arule)
//  newEntrySheet.getRange("F4:I6").setDataValidation(arule)

  buildMainPickList()
/*
  utilList = ["Electricity","Water","Gas","None"]   // Make Pick List for utils
  arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Choose from Menu 2 Only").requireValueInList(utilList).build()
  newEntrySheet.getRange('C4:C6').setDataValidation(arule)

  accList =["Debit (+)","Credit (-)", "Note","None"]  // Make pick list for Debits /Credits
  arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Choose from Menu 3 Only").requireValueInList(accList).build()
  newEntrySheet.getRange('C7:C14').setDataValidation(arule)
*/

  // Set all notes in cell
  newEntrySheet.getRange("C3").setNote("Invoice For Which Month")
  newEntrySheet.getRange("E4:E6").setNote("Utility Start Value")
  newEntrySheet.getRange("F4:F6").setNote("Utility End Value")
  newEntrySheet.getRange("G4:G6").setNote("Utility Consumed")
  newEntrySheet.getRange("H4:H6").setNote("Utility Calculated Cost")
  newEntrySheet.getRange("I4:I6").setNote("Utility Charged")
  newEntrySheet.getRange("E7:E14").setNote("Transaction Value" )
  newEntrySheet.getRange("F7:F14").setNote("Text Note" )
  newEntrySheet.getRangeList(["C4:C14"]).setValue("None")
 // makeTotal ()
 }

function buildMainPickList()
{
  ss = SpreadsheetApp.getActiveSpreadsheet();
  var newEntrySheet = ss.getSheetByName("DataEntry")
  var utilList = ["Electricity","Water","Gas","None"]   // Make Pick List for utils
  var arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Choose from Menu 2 Only").requireValueInList(utilList).build()
  newEntrySheet.getRange('C4:C6').setDataValidation(arule)

  var accList =["Debit (+)","Credit (-)", "Note","None"]  // Make pick list for Debits /Credits
  arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Choose from Menu 3 Only").requireValueInList(accList).build()
  newEntrySheet.getRange('C7:C14').setDataValidation(arule)

}
/*
function makeTotal(){
  getEntrySheet()
 // myLog("G3 is " + entrySheet.getRange(3,7,1,1).getValue())
  entrySheet.getRange("G3").setNote("Invoice Total For Month")//Calculate Invoice total formulae
//  entrySheet.getRange("M4:M14").setValue("1")
  entrySheet.getRange("N4:N14").setValues([["=I4*M4"],["=I5*M5"],["=I6*M6"],["=E7*M7"],["=E8*M8"],["=E9*M9"],["=E10*M10"],["=E11*M11"],["=E12*M12"],["=E13*M13"],["=E14*M14"]])
  entrySheet.getRange("G3").setValue("=SUM(N4:N14)+G1")
  entrySheet.hideColumns(13,2)
}

*/

function calcTotal(){
  getEntrySheet()
  c1Stat = entrySheet.getRange("C1").getValue()
  if (c1Stat != "Not Saved"){
    writeSaved(false)
  }
  invTotal = entrySheetData[0][6]
  myArLog(entrySheetData)
  for (var i = 4; i <= 6; i++){
    uVal = entrySheet.getRange(i,9,1,1).getValue()
   myLog("Util Value " +uVal)
    if (uVal != ""){
      invTotal += uVal
    }
   myLog("Row number " +invTotal)
  }
  
  depPlus = 0   //Deposit Required
  depMinus = 0 //Deposit paid
  myLog(invTotal)
  for (var i = 6; i <= 13; i++) {
    dcType =  entrySheetData[i][2]
    dcSubType = entrySheetData[i][3]
    dcVal = entrySheetData[i][4]
    if (dcType == "Debit (+)"){
      invTotal = dcVal + invTotal
    } else if (dcType == "Credit (-)"){
      invTotal = invTotal - dcVal
    } else if (dcType == "Note"){
      invTotal = dcVal + invTotal
    }

    if (dcSubType == "Deposit Reqd"){
       depPlus +=  dcVal
    }
    if (dcSubType == "Deposit Paid"){
      depMinus +=  dcVal
      myLog("Dep minus " + depMinus)
    }
    if (dcSubType == "Deposit Refund"){
      depMinus +=  dcVal
    }
    myLog(invTotal)
  }
//  if ((depPlus == 0) && (depMinus > 0)) {
//       invTotal +=  depMinus
//  }
  
  entrySheet.getRange("G3").setValue(invTotal)
}

function testDepPaid(){
  checkIfDepPaid(8)
}

function checkIfDepPaid(theRow){
  getEntrySheetOnce()
  bCell = entrySheet.getRange(theRow,4,1,1).getValue()
  if (bCell == "Deposit Paid"){
    entrySheet.getRange(theRow,13,1,1).setValue("0") // Dont count deposits paid in invoice total update deposit
  }
}

// All the live stuff below Dynamic piclist and clearing

function makePickList(theRow){
  crList = ["Payments","Discount","Deposit Paid","Deposit Refund","Other:"]
  drList = ["Rent","Fibre", "Rates","Service Fee", "Interest Charge","Penalty","Deposit Reqd","Other:" ]
  noteList = ["Type Note /Message","Other:"]
  utilPayList = ["PostPaid","PrePaid","Fix","New","Other:"]


  getEntrySheetOnce()
  
  var options = {
   formatOnly: false,
   contentsOnly: true,
   validationsOnly : true
  };

  aCell = entrySheet.getRange(theRow,3,1,1).getValue()
//  if(entrySheet.getRange(3,7,1,1).getValue() == ""){
//   makeTotal()
//  }
  if (aCell == "Debit (+)"){
    subListVal = entrySheet.getRange(theRow,4,1,1).getValue()
    arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Choose from Menu 4 Only").requireValueInList(drList).build()
    entrySheet.getRange(theRow,4,1,1).setDataValidation(arule)
//    entrySheet.getRange(theRow,13,1,1).setValue("1")
  }
  else if (aCell == "Credit (-)"){
    subListVal = entrySheet.getRange(theRow,4,1,1).getValue()
    arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Choose from Menu 5 Only").requireValueInList(crList).build()
    entrySheet.getRange(theRow,4,1,1).setDataValidation(arule)
 //   entrySheet.getRange(theRow,13,1,1).setValue("-1")
  }
  else if (aCell == "Note"){
    subListVal = entrySheet.getRange(theRow,4,1,1).getValue()
    arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Choose from Menu 6 Only").requireValueInList(noteList).build()
    entrySheet.getRange(theRow,4,1,1).setDataValidation(arule)
  //  entrySheet.getRange(theRow,13,1,1).setValue("1")
  }
  else if (aCell == "None"){
    if (theRow < 7) {
      entrySheet.getRange(theRow,4,1,6).clear(options)
    }else
      entrySheet.getRange(theRow,4,1,3).clear(options)
  }
  else if  ((aCell == "Electricity") || (aCell == "Water") || (aCell == "Gas")) {
    arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Choose from Menu 7 Only").requireValueInList(utilPayList).build()
    entrySheet.getRange(theRow,4,1,1).setDataValidation(arule)
    entrySheet.getRange(theRow,7,1,1).setValue("=(F"+theRow+"-E"+theRow+")") // Util used formulae
    entrySheet.getRange(theRow,9,1,1).setValue('=if(D'+theRow+'="PrePaid",0,H'+theRow+')')    // Make a copy of cost that can be changed
   // Insert relavant formulea in appropriate cell
    if (aCell == "Electricity") {
      entrySheet.getRange(theRow,8,1,1).setValue("=calcElec(G"+theRow+")")

    } else if (aCell == "Water") {
      entrySheet.getRange(theRow,8,1,1).setValue("=calcWater(G"+theRow+")")
    } else if (aCell == "Gas") {
         entrySheet.getRange(theRow,8,1,1).setValue("=calcGas(G"+theRow+")")
    }
  }

}

//      entrySheet.getRange(theRow,8,1,1).setValue('=if(D'+theRow+'="PrePaid",0,calcElec(G'+theRow+'))') // this works !!, but not needed
//.                                               ('=if(D'+theRow+'="PrePaid",0,'H'+theRow+')')



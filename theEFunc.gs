/// ---- theEFunc ------ RS


function getLastEditInvoice()
{
 getPropSheet()
  colIndex1 = getPropSheetLastDataCol()
  if (colIndex1 == 1) {
    entrySheet.getRange(2,2,1,1).setValue("None")
     entrySheet.getRange(2,5,1,1).setValue("None")
     entrySheet.getRange(2,7,1,1).setValue("None")
    return
  }
  lastInvDate = propSheetData[2][colIndex1-1]
  entrySheet.getRange(2,2,1,1).setValue(lastInvDate)
  lastInvNumber = propSheetData[3][colIndex1-1]
  entrySheet.getRange(2,5,1,1).setValue(lastInvNumber)
  lastInvTotal = propSheetData[4][colIndex1-1]
  entrySheet.getRange(2,7,1,1).setValue(lastInvTotal)
  /*
  utilA_Name =  propSheetData[7][colIndex1-1]
  entrySheet.getRange(7,3,1,1).setValue(utilA_Name)
  utilA_Type =  propSheetData[8][colIndex1-1]
  entrySheet.getRange(7,4,1,1).setValue(utilA_Type)
  utilA_Last =  propSheetData[10][colIndex1-1]
  entrySheet.getRange(7,5,1,1).setValue(utilA_Last)

  utilB_Name =  propSheetData[17][colIndex1-1]
  entrySheet.getRange(8,3,1,1).setValue(utilB_Name)
  utilB_Type =  propSheetData[18][colIndex1-1]
  entrySheet.getRange(8,4,1,1).setValue(utilB_Type)
  utilB_Last =  propSheetData[20][colIndex1-1]
  entrySheet.getRange(8,5,1,1).setValue(utilB_Last)

  utilC_Name =  propSheetData[27][colIndex1-1]
  entrySheet.getRange(9,3,1,1).setValue(utilC_Name)
  utilC_Type =  propSheetData[28][colIndex1-1]
  entrySheet.getRange(9,4,1,1).setValue(utilC_Type)
  utilC_Last =  propSheetData[30][colIndex1-1]
  entrySheet.getRange(9,5,1,1).setValue(utilC_Last)
*/


}





function onEdit(e) {
  const dataEntrySheet = "DataEntry"   // for example//
  const statEntrySheet = "Statement"
  const specificCell = "C16"  //"B3"       // for example
  var range = e.range;
  var theCell = [0,0]

  let dataESheet = (range.getSheet().getName() == dataEntrySheet)
  let statESheet = (range.getSheet().getName() == statEntrySheet)
  if (dataESheet) {
      
    theCell[0] = range.getRow()
    theCell[1] = range.getColumn()
    pickListRange = [4,5,6,7,8,9,10,11,12,13,14]  // Rows in data entry sheet
  
    drCrIn= pickListRange.indexOf(theCell[0])

    if ((theCell[0] == 19) && (theCell[1] == 1)){     // Test cell to see if onEdit is running A19
      range.setNote("Back Groung Script is Running :- Date " + new Date );
      return
    }
    else if (range.getA1Notation() == "B2"){    // Change Property
      menuOne()
      loadUp()  // BaseToEntry
      return
  //    range.getSheet().getRange(3,5,1,1).setValue("Temp")
    }
    else if (range.getA1Notation() == "I14"){   // Change Invoice number
      getPickInvoice()
      return
    }
    else if ((drCrIn != -1) && (theCell[1] == 3)) {   // Make entry configurations & picks
      makePickList(theCell[0])
    }
    calcTotal()
  }
  else if (statESheet) {
   if (range.getA1Notation() == "A2"){
    var myProp = range.getSheet().getRange(2,1,1,1).getValue()
    populateInvNumber(myProp)
   // range.setNote("Back Groung Script is Running Statement :- Date " + new Date + "Prop " + myProp);
   }
  }
  else
  {
    return
  }

}


/*

function onEdit(e) {
  const specificSheet = "DataEntry"   // for example//
  const specificCell = "C16"  //"B3"       // for example
  var theCell = [0,0]
  var range = e.range;
    
    theCell[0] = range.getRow()
    theCell[1] = range.getColumn()

  
    let sheetCheck = (range.getSheet().getName() == specificSheet)
    if (sheetCheck)
    {
      range.setNote('Last modified: ' + new Date + " Sheet: " + sheetCheck + "Cell: " + theCell[0] + "," + theCell[1]);
    }
}

*/
/*
function onEdit(e) {
  const specificSheet = "DataEntry"   // for example//
  const specificCell = "C16"  //"B3"       // for example
  var range = e.range;

  let sheetCheck = (range.getSheet().getName() == specificSheet)
  if (sheetCheck) {
    if (range.getA1Notation() == "C16"){
     // getLastInv()
     // entrySheet.getRange(16,2,1,1).setValue("getLastInv")
        num = range.getSheet().getRange(18,3,1,1).getValue()
        range.getSheet().getRange(18,5,1,1).setValue("The num is " + num)
        range.setNote('Last modified: ' + new Date+ " Sheet: " + sheetCheck + "Cell: " + cellCheck + "Num: " + num );
    }
    else if (range.getA1Notation() == "B2"){
        range.getSheet().getRange(1,5,1,1).setValue("Got last inv data")
        num = range.getSheet().getRange(20,3,1,1).getValue()
        range.getSheet().getRange(20,5,1,1).setValue("The num is " + num)
        range.setNote('Last modified: ' + new Date+ " Sheet: " + sheetCheck + "Cell: " + cellCheck + "Num: " + num );
    }
    else if (range.getA1Notation() == "C5") {
      range.getSheet().getRange(1,5,1,1).setValue("Got last inv data")
      range.getSheet().getRange(19,5,1,1).setValue("The num is " + num)

    }
    else if (range.getA1Notation() == "C7") {
      makePickList("C7")
    }
    else{
      return
    }
  }
  else
  {
    return
  }

}

*/



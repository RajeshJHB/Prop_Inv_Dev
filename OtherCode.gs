/// ---- OtherCode ------ RS

/*
ss = SpreadsheetApp.getActiveSpreadsheet();
newEntrySheet = ss.getSheetByName("DataEntry")
var range = ss.getRange('A1:B10');
var protection = range.protect().setDescription('Sample protected range');

// Ensure the current user is an editor before removing others. Otherwise, if the user's edit
// permission comes from a group, the script throws an exception upon removing the group.
var me = Session.getEffectiveUser();
protection.addEditor(me);
protection.removeEditors(protection.getEditors());
if (protection.canDomainEdit()) {
  protection.setDomainEdit(true);
}


function testA1() {

  //2d array
  var wholeValues = [];


  for (var i = 0; i < 5; i++){

    //create a 1D array first with pushing 0,1,2 elements with a for loop
    var value = [];
    for (var j = 0; j < 3; j++) {
      value.push(j);
    }
    //pushing the value array with [0,1,2] to thw wholeValues array.
    wholeValues.push(value);
  } // the outer for loop runs five times , so five the 0,1,2 with be pushed in to thewholevalues array by creating wholeValues[0][0],wholeValues[0][1]...till..wholeValues[4][2]

  myArLog(wholeValues);
}

function testArr(){
  stArr = [["A","B","C"],["1","2","3"]]
  outArr = []
  inArr =[]
  in2Arr = []
  inArr.push(stArr[0][1])
  inArr.push(stArr[0][2])
  inArr.push(stArr[1][0])
  myArLog(inArr)
  outArr.push(inArr)
  myArLog(outArr)
  in2Arr.push(stArr[1][1])
  outArr.push(in2Arr)
  myArLog(outArr)
  outArr.push(inArr)
  myArLog(outArr)

}


function tdtest() {
  var wholeValues = [[],[]];
  var value = [];
    for (var i = 0; i < 5; i++){
       for (var j = 0; j < 3; j++) {
           wholeValues[[i],[j]] = value[j];
       }
    }

  Logger.log(wholeValues[[0],[1]]);
}


    if (invDate.getDate() == propSheetData[2][colIndex-1].getDate()) {
      colIndex = colIndex // overwrite
      myLog("Overwrite Index @ 1 " + colIndex)
    } else {
      colIndex = colIndex + 1// nex col
      myLog("Overwrite Index @ 2 " + colIndex)

    }

function genEntryToPropBase (colIndex){
  proCode = getPropSheet()
  colIndex = getPropSheetLastCol()
  myLog("Entry to base " + colIndex)
  //packPutEntryData(colIndex)
 // packPutFullEntryData(colIndex)
}

function getFullEntryData(){
  proCode = getPropSheet()
  colIndex = getPropSheetLastCol()
  var dEntry = []
  dEntryIndex = 0
  for (var i =0 ; i <=13; i++ ){    // Row 6 to 8
    for (var j= 0 ; j<=8; j++) {
          dEntry[dEntryIndex] = entrySheetData[i][j]
          dEntryIndex ++
    }
  }
  myArLog(dEntry)
  dEntry2D = onetoTwod(dEntry,1)
  propSheet.getRange(1,colIndex+1,dEntryIndex,1).setValues(dEntry2D)
}



function packPutEntryData(colNum)
{
  var dEntry = []
  dEntry[0] = entrySheetData[0][2]  //  Last Inv Date
  dEntry[1] = entrySheetData[1][5] //  Last Inv Total
  dEntry[2] = entrySheetData[1][7]  // Last Balance
  dEntry[4] = entrySheetData[2][2] // Property Code
  dEntry[4] = entrySheetData[2][4] // Tenant code
  dEntry[3] = entrySheetData[2][3] // Tenant code
  dEntry[4] = entrySheetData[4][1] // Invoice Date
  dEntry[5] = entrySheetData[4][4] // Invoice number
  dEntry[6] = entrySheetData[4][6]  // Invoice Total
  dEntryIndex = 5
  for (var i = 6 ; i <= 8; i++ ){    // Row 5 (blank) start from 6 to 9
    for (var j= 0 ; j<=9; j++) {
          dEntry[dEntryIndex] = entrySheetData[i][j]
          dEntryIndex ++
    }
  }
  for (var i =9 ; i <=16; i++ ){    // Row 6 to 8
    for (var j= 0 ; j<=9; j++) {
          dEntry[dEntryIndex] = entrySheetData[i][j]
          dEntryIndex ++
    }
  }
  dEntry2D = onetoTwod(dEntry,1)
  propSheet.getRange(1,colNum+1,dEntryIndex,1).setValues(dEntry2D)
}


 //newEntrySheet.getRange('A1').setValue("Make Invoice Tye")
  //arule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Choose from Menu Only").requireValueInRange(range).build(). // from items in the sheet
 
//  rule1 = SpreadsheetApp.newDataValidation().setAllowInvalid(false).setHelpText("Check value").requireValueInList(['Open', 'Closed']).build();
//  cell1.setDataValidation(rule1);
   
  cell = newEntrySheet.getRange('D1')
 // cell  = SpreadsheetApp.getActive().getRange('D1');

  range = newEntrySheet.getRange('B1:B10');
  myLog(range)
  //var range = SpreadsheetApp.getActive().getRange('B1:B10');
  
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(range).build()
 
  cell.setDataValidation(rule);



var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheets()[0];

var cell = sheet.getRange("B2");

// Sets the background to white
cell.setBackgroundRGB(255, 255, 255);

// Sets the background to red
cell.setBackgroundRGB(255, 0, 0);


=gethex("G1") // get colour code

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheets()[0];

var colors = [
  ["red", "white", "blue"],
  ["#FF0000", "#FFFFFF", "#0000FF"] // These are the hex equivalents
];

var cell = sheet.getRange("B5:D6");
cell.setBackgrounds(colors);


  cell = newEntrySheet.getRange('D1')
  rule = cell.getDataValidation();
  criteria = rule.getCriteriaType();
  args = rule.getCriteriaValues();
  myLog("Citeria " + criteria + "Args " + args )


  */
/*
function utilDataEntryUpdate() {
  getPropSheet()
  nextReadRow = 6
  for (var i = 0; i <= 2; i++ )  // 3 rows
  {
    utilType = entrySheetData[nextReadRow+i][2]
    utilStartReading = entrySheetData[nextReadRow+i][4]
    myLog("Start "+utilStartReading)
    utilEndReading = entrySheetData[nextReadRow+i][5]
    myLog("End "+utilEndReading)
    utilConsumed = utilEndReading - utilStartReading
    myLog(utilConsumed)
    if (utilType == "None") {

    }
    else if (utilType == "Electricity") {

    }
    else if (utilType == "Water") {

    }
     else if (utilType == "Gas") {

    }
  }
}
 
  }
  else
  {
    utilChargeType = entrySheetData[nextReadRow-1+i][3]
    myLog (utilType + " ->  "  + utilChargeType)
    utilDate = entrySheetData[nextReadRow-1+i][1]
    utilStartReading = entrySheetData[nextReadRow-1+i][4]
    utilEndReading = entrySheetData[nextReadRow-1+i][5]
    utilConsumed = entrySheetData[nextReadRow-1+i][6]
    utilCalcCharge = entrySheetData[nextReadRow-1+i][7]
    utilActCharged = entrySheetData[nextReadRow-1+i][8]
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

}

function testDate(){
    prop =   getPropSheet() // from data entry screen
    myLog(prop)
    nexCol = getPropSheetLastCol() //next new column
    lastDate = propSheet.getRange(6,colIndex-1,1,1).getValue()
    myLog(lastDate)
    dateonly = lastDate.getDate()
    myLog(dateonly)

  //  comRes = dates.compare(lastDate,lastDate)
  //  myLog(comRes)
}

*/


/// ---- ClearEntry ------ RS

function clearDataEntry(clearPick) {
  getEntrySheet()
  var options = {
   formatOnly: false,
   contentsOnly: true,
   validationsOnly : true
  };
  //myArLog(entrySheetData)
  entrySheet.getRangeList(["B1:C1","E1","G1","B3:C3","E3","G3","I1:I2","B4:F14","G4:I6","H13:H14"]).clearContent()
  entrySheet.getRange("C1").setBackground("#fafafa")
  entrySheet.getRangeList(["C4:C14"]).setValue("None")
  entrySheet.getRangeList(["C4:D14","I14"]).clear(options)  //"D4:D14"
  if (clearPick) {
    entrySheet.getRangeList(["B2","D2:E2"]).clearContent()
  }
  entrySheet.getRange("A21:I34").clearContent()
  entrySheet.getRange("E3").setValue("Temp")
}


function clearInvoice() {
  var invSheet = SpreadsheetApp.getActive().getSheetByName("Invoice");
  invSheet.getRangeList(["B3:B7","C8:C15","D9:D13","E4:E16","B18:E29","E30","C34:C39"]).clearContent()  // last Invoice
}

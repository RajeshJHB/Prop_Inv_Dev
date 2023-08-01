/// ---- Make_Invoice_Sheet ------ RS

function makeInvoiceSheet(){
  ss = SpreadsheetApp.getActiveSpreadsheet();
  invEntrySheet = ss.getSheetByName("Invoice")
  if (invEntrySheet == null){
    ss.insertSheet("Invoice")
    invEntrySheet = ss.getSheetByName("Invoice")
  }
  maxCol = invEntrySheet.getMaxColumns()
  if (maxCol <10){
    invEntrySheet.insertColumnsAfter(1,10-maxCol)
  }
  
  invEntrySheet.getRange('A1:E40').setValues([
  ["","","","",""],
  ["","SolProp Investments","","INVOICE",""],
  ["","","","",""],
  ["","","","INVOICE No.",""],
  ["","","","INVOICE DATE",""],
  ["","","","FOR MONTH OF",""],
  ["","","","",""],
  ["","TO:","","Property",""],
  ["","","","",""],
  ["","","","",""],
  ["","","","",""],
  ["","","","",""],
  ["","","","",""],
  ["","","","",""],
  ["","","","Deposit Paid",""],
  ["","","","Deposit Required",""],
  ["","Date","Description","","Amount"],
  ["","","","",""],
  ["","","","",""],
  ["","","","",""],
  ["","","","",""],
  ["","","","",""],
  ["","","","",""],
  ["","","","",""],
  ["","","","",""],
  ["","","","",""],
  ["","","","",""],
  ["","","","",""],
  ["","","","",""],
  ["","Total Due","","",""],
  ["","All amounts are due and payable on the first day of each month. Interest will be charged on overdue accounts. Payments will be credited against arrears if any.","","",""],
  ["","","","",""],
  ["","Banking Details: ","","",""],
  ["","Bank","","",""],
  ["","Acc Name","","",""],
  ["","Branch Code","","",""],
  ["","Account No.","","",""],
  ["","Deposit Ref.","","",""],
  ["","Amount","","",""],
  ["","","","",""],
  ])

//----------Set Col Widths------------------------
  invEntrySheet.setColumnWidth(1,100);
  invEntrySheet.setColumnWidth(2,90);
  invEntrySheet.setColumnWidth(3,299);
  invEntrySheet.setColumnWidth(4,119);
  invEntrySheet.setColumnWidth(5,100);
  invEntrySheet.setRowHeight(2,21)
  invEntrySheet.setRowHeight(30,37)
//______________________________________


var myRange = invEntrySheet.getRange("B2:C2")
myRange.mergeAcross()

myRange = invEntrySheet.getRange("D2:E2")
myRange.mergeAcross()

myRange = invEntrySheet.getRange("C17:D17")
myRange.mergeAcross()

myRange = invEntrySheet.getRange("B30:C30")
myRange.mergeAcross()

myRange = invEntrySheet.getRange("B31:E32")
myRange.merge()

myRange = invEntrySheet.getRange("B33:E33")
myRange.mergeAcross()

myRange = invEntrySheet.getRange("C18:D18")
myRange.mergeAcross()
myRange = invEntrySheet.getRange("C19:D19")
myRange.mergeAcross()
myRange = invEntrySheet.getRange("C20:D20")
myRange.mergeAcross()
myRange = invEntrySheet.getRange("C21:D21")
myRange.mergeAcross()
myRange = invEntrySheet.getRange("C22:D22")
myRange.mergeAcross()
myRange = invEntrySheet.getRange("C23:D23")
myRange.mergeAcross()
myRange = invEntrySheet.getRange("C24:D24")
myRange.mergeAcross()
myRange = invEntrySheet.getRange("C25:D25")
myRange.mergeAcross()
myRange = invEntrySheet.getRange("C26:D26")
myRange.mergeAcross()
myRange = invEntrySheet.getRange("C27:D27")
myRange.mergeAcross()
myRange = invEntrySheet.getRange("C28:D28")
myRange.mergeAcross()
myRange = invEntrySheet.getRange("C29:D29")
myRange.mergeAcross()



myRange = invEntrySheet.getRange("C34:D34")
myRange.mergeAcross()
myRange = invEntrySheet.getRange("C35:D35")
myRange.mergeAcross()
myRange = invEntrySheet.getRange("C36:D36")
myRange.mergeAcross()
myRange = invEntrySheet.getRange("C37:D37")
myRange.mergeAcross()
myRange = invEntrySheet.getRange("C38:D38")
myRange.mergeAcross()
myRange = invEntrySheet.getRange("C39:D39")
myRange.mergeAcross()
myRange = invEntrySheet.getRange("C40:D40")
myRange.mergeAcross()



myRange = invEntrySheet.getRangeList(["B8:E16","B17:B29","C17:D29","E17:E30","B30:C30","D30:E30","B31:E32","B33:E33","B34:B40","E34:E40"])
myRange.setBorder(true, true, true, true, false, false);

myRange = invEntrySheet.getRange("E30")
myRange.setBorder(null, true, null, null, null, null);

myRange = invEntrySheet.getRangeList(["B17","C17:D17","E17"])
myRange.setBorder(null, null, true, null, false, false);

myRange = invEntrySheet.getRange("B34:E40")
myRange.setBorder(true, true, true, true, true, true);


//____________SolProp Title_____
  myRange = invEntrySheet.getRange("B2")
  var myFontStyle = SpreadsheetApp.newTextStyle()
    .setBold(true)
    .setFontSize(16)
    .setItalic(false)
    .setFontFamily("Arial")
    .build();
 
  myRange.setTextStyle(myFontStyle)
 //-------------------------------------


//____________INVOICE_____
  myRange = invEntrySheet.getRange("D2")
  var myFontStyle = SpreadsheetApp.newTextStyle()
    .setBold(true)
    .setFontSize(20)
    .setItalic(false)
    .setFontFamily("Arial")
    .build();
 
  myRange.setTextStyle(myFontStyle)
 //-------------------------------------

//____________INV No...Date etc_____
  myRange = invEntrySheet.getRange("D4:D6")
  var myFontStyle = SpreadsheetApp.newTextStyle()
    .setBold(true)
    .setFontSize(10)
    .setItalic(true)
    .setFontFamily("Arial")
    .build();
 
  myRange.setTextStyle(myFontStyle)
 //-------------------------------------

//____________TO: & Property_____
  myRange = invEntrySheet.getRange("B8:D8")
  var myFontStyle = SpreadsheetApp.newTextStyle()
    .setBold(true)
    .setFontSize(11)
    .setItalic(true)
    .setFontFamily("Arial")
    .build();
 
  myRange.setTextStyle(myFontStyle)
 //-------------------------------------

//____________Date, Desc & Amount_____
  myRange = invEntrySheet.getRange("B17:E17")
  var myFontStyle = SpreadsheetApp.newTextStyle()
    .setBold(true)
    .setFontSize(12)
    .setItalic(false)
    .setFontFamily("Arial")
    .build();
 
  myRange.setTextStyle(myFontStyle)
 //-------------------------------------

//____________Total Due _____
  myRange = invEntrySheet.getRange("B30")
  var myFontStyle = SpreadsheetApp.newTextStyle()
    .setBold(true)
    .setFontSize(14)
    .setItalic(false)
    .setFontFamily("Arial")
    .build();

  myRange.setTextStyle(myFontStyle)
  myRange.setHorizontalAlignment("center")
 //-------------------------------------


//---------------------------All Amounts are due...

  myFontStyle = SpreadsheetApp.newTextStyle()
    .setBold(false)
    .setFontSize(12)
    .setItalic(false)
    .setFontFamily("Arial")
    .build();
  myRange = invEntrySheet.getRange("B31");
  myRange.setWrap(true)
  myRange.setHorizontalAlignment("center")
  myRange.setTextStyle(myFontStyle)
//---------------------------------------
  
//____________Banking Details Heading _____
  myRange = invEntrySheet.getRange("B33")
  var myFontStyle = SpreadsheetApp.newTextStyle()
    .setBold(false)
    .setFontSize(16)
    .setItalic(false)
    .setFontFamily("Arial")
    .setUnderline(true)
    .build();

  myRange.setTextStyle(myFontStyle)
  myRange.setHorizontalAlignment("center")
 //-------------------------------------

//____________Banking Details _____
  myRange = invEntrySheet.getRange("C34:C40")
  var myFontStyle = SpreadsheetApp.newTextStyle()
    .setBold(true)
    .setFontSize(11)
    .setItalic(false)
    .setFontFamily("Arial")
    .build();

  myRange.setTextStyle(myFontStyle)
  myRange.setHorizontalAlignment("center")
 //-------------------------------------

//____________Currency Format Values _____
  myRange = invEntrySheet.getRange("E15:E30")
  myRange.setNumberFormat('R 0.00')
 
  myRange = invEntrySheet.getRange("C39")
  myRange.setNumberFormat('R 0.00')

 //-------------------------------------
}

function countRows() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadSheet.getSheets();
  var toViewSheet = sheets[SHEET_IDX_TO_VIEW];
  
  var nRet = 0;
  for(nRows = ROW_START; nRows < sheet.getMaxRows(); nRows++ ) {
    var cell = toViewSheet.getRange(nRows, 1);
    
    if(!cell.getValue()) {
      break;
    }
    nRet = nRows
  }
  return nRet;
}

function loadGmailFromAdresses() {
  var aliases = GmailApp.getAliases()
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(aliases).build();
  
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadSheet.getSheets();
  var toviewSheet = sheets[SHEET_IDX_TO_VIEW];
  
  for(i = ROW_START; i <= countRows(); i++) {
    var cell = toViewSheet.getRange(i, COL_FROM_ADDR);
    cell.setDataValidation(rule);
  }
  
  Browser.msgBox("Load from addresses completed!")
}

function loadGmailDrafts() {
  
  /* Build dropdown list (Value In List Rule) */
  var draftMsgs = GmailApp.getDraftMessages();
  var msgSubjects = [];
  for (i = 0; i < draftMsgs.length; i++) {
    msgSubjects.push(draftMsgs[i].getSubject())
  }
  
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(msgSubjects).build();
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var toViewSheet = spreadSheet.getSheets[SHEET_IDX_TO_VIEW];
  
  for(i = ROW_START; i <= countRows(); i++) {
    var cell = toViewSheet.getRange(i, COL_TEMPLATE);
    cell.setDataValidation(rule);
  }
  Browser.msgBox("Load templates completed!")
}



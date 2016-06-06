 function sendAllPayslips(){
    var originalSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = originalSpreadsheet.getSheets();
    
    var instructorInfoSheet = sheets[SHEET_IDX_INSTRUCTOR_INFO];
    var instructorInfoRange = instructorInfoSheet.getDataRange();
    var instructorInfoValue = instructorInfoRange.getValues();
   
    var toSendSheet = sheets[SHEET_IDX_TO_SEND];
    
    var toViewSheet = sheets[SHEET_IDX_TO_VIEW];
    var toViewRange = toViewSheet.getDataRange(); 
    var toViewCell = toViewRange.getCell(5, 10);
    var progressCell = toViewRange.getCell(6,10);

    for (var r = 1; r < instructorInfoValue.length; r++){
        var row = instructorInfoValue[r];
        var code = row[2];
        var name = row[1];
        var email = row[3];
        toViewCell.setValue(code);
        if(toViewCell.getValue() == code){
          progressCell.setValue("Loading")
          var instrInfo = loadInstrInfo(code);
          var salary = loadSalary(code);
          if(salary[0] != null){           
            printInstructorInfo(instrInfo,toViewSheet);
            printInstructorInfo(instrInfo,toSendSheet);
            
            printSalary(salary,toViewSheet);
            printSalary(salary,toSendSheet);

            var draftMsgs = GmailApp.getDraftMessages();
  
            var newSpreadsheet = SpreadsheetApp.create("Spreadsheet to export");
                        
            toSendSheet.copyTo(newSpreadsheet);
            
            newSpreadsheet.getSheetByName('Sheet1').activate();
            newSpreadsheet.deleteActiveSheet();
            
            var pdf = DriveApp.getFileById(newSpreadsheet.getId()).getAs('application/pdf').getBytes();
            var attach = {fileName:'Payslips.pdf',content:pdf, mimeType:'application/pdf'};
            for(row = ROW_START; row <= countRows(); row++) {
                var range = toViewSheet.getRange(row, 1, 1, COL_TOTAL);
                var values = range.getValues();
      
                var from_addr = values[0][COL_FROM_ADDR - 1];
                var subject = values[0][COL_TEMPLATE - 1];
                if(subject) {
                    for(draftMsgIdx = 0; draftMsgIdx < draftMsgs.length; draftMsgIdx++) {
                        var draftMsg = draftMsgs[draftMsgIdx];
                        if(draftMsg.getSubject() == subject) {
                        var body = draftMsg.getBody();
                        body = body.replace(STR_NAME_PATTERN, name);
                        GmailApp.sendEmail(
                            "dracarys1312@gmail.com",         
                            subject,                 
                            '', {                        
                            htmlBody: body,
                            from: from_addr,
                            attachments:[attach]
                            }); 
                        }
                    }
                }   
            }
          }
          DriveApp.getFileById(newSpreadsheet.getId()).setTrashed(true);  
        }
    }       
    progressCell.setValue("Done")
}

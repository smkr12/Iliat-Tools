COL_TOTAL = 100
ROW_START = 2


// Sheet index
var SHEET_IDX_INSTRUCTOR_INFO = 0;
var SHEET_IDX_INSTRUCTOR_RATE = 1;
var SHEET_IDX_CLASS = 2;
var SHEET_IDX_SALARY = 3;
var SHEET_IDX_TO_VIEW = 5;
var SHEET_IDX_TO_SEND = 6;

// COL TOTAL in Salary sheet
var COL_TOTAL_INSTRUCTOR = 7;
var COL_TOTAL_RATE = 5;
var COL_TOTAL_SALARY = 12;

// Instructor Info sheet index
var INSTRUCTOR_SHEET_COL_NAME = 1;
var INSTRUCTOR_SHEET_COL_TEAM = 6;

// Salary sheet index
var SALARY_SHEET_COL_NO = 0;
var SALARY_SHEET_COL_CODE = 1;
var SALARY_SHEET_COL_NAME = 2;
var SALARY_SHEET_COL_TEAM = 3;
var SALARY_SHEET_COL_COURSE = 4;
var SALARY_SHEET_COL_CLASS = 5;
var SALARY_SHEET_COL_SALARY_RATE1 = 7;
var SALARY_SHEET_COL_SALARY_RATE2 = 10;
var SALARY_SHEET_COL_SALARY1 = 8
var SALARY_SHEET_COL_SALARY2 = 11
var SALARY_SHEET_COL_TOTAL_SALARY = 13;

// Total ROW
var INSTRUCTOR_INFO_TOTAL_ROWS = 66;
var SALARY_TOTAL_ROWS = 78;

var KPI_COL_IN_TARGET_SHEET = 6;
var ROW_NUM_14 = 14;
var ROW_NUM_18 = 18;


function onOpen() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var fEasyMenuEntries = [ 
    {name: "Clear all", functionName: "clearAll"},
    {name: "Load Instructor Codes & Rates", functionName: "loadInstructorCodesAndRates"},
    {name: "Send All Payslips", functionName: "sendAllPayslips"},
    {name: "Send This Payslips", functionName: "sendThisPaySlips"}
  ];
  spreadSheet.addMenu("FEasy2", fEasyMenuEntries);
}
    

function clearAll() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    
  var allSheets = spreadSheet.getSheets();

  var salarySheet = allSheets[SHEET_IDX_SALARY];
  
  for(var row = 3; row < salarySheet.getMaxRows(); row++) {
    var values = salarySheet.getRange(row, 1, 1, COL_TOTAL_SALARY).getValues();
    if(!values[0][1])
      break;
    values[0][0] = "";
    values[0][6] = "";
    values[0][9] = "";
    salarySheet.getRange(row, 1, 1, COL_TOTAL_SALARY).setValues(values);
    salarySheet.getRange(row, 1, 1, COL_TOTAL_SALARY).setBackground('white');
  }
}


function loadInstructorRates(instructorRateSheet) {
  
  var instructorRateList = [];
  for (var row = 2; row < instructorRateSheet.getMaxRows(); row++) {
    var values = instructorRateSheet.getRange(row, 1, 1, COL_TOTAL_RATE).getValues();
    if(!values[0][1])
      break;
    var name = values[0][1];
    var course = values[0][2];
    var salary1 = values[0][3];
    var salary2  = values[0][4];
    instructorRateList.push({
      name:name.toString(),
      course:course.toString(),
      salary1:salary1,
      salary2:salary2
    });
  }
  
  return instructorRateList;
}

function loadInstructorCodesAndRates() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var allSheets = spreadSheet.getSheets();
  
  var instructorSheet = allSheets[SHEET_IDX_INSTRUCTOR_INFO];
  var salarySheet = allSheets[SHEET_IDX_SALARY];
  var instructorRateSheet = allSheets[SHEET_IDX_INSTRUCTOR_RATE];
  
  instructorRateList = loadInstructorRates(instructorRateSheet);
  
  var instructorRateMap = {};
  for(var idx = 0; idx < instructorRateList.length; idx++) {
    var name = instructorRateList[idx].name;
    var course = instructorRateList[idx].course;
    instructorRateMap[name + course] = instructorRateList[idx];
  }
 
  var instructorList = loadInstructorInfo(instructorSheet);
  
  var instructorMap = {};
  for (var idx = 0; idx < instructorList.length; idx++) {
    name = instructorList[idx].name;
    team = instructorList[idx].team;
    instructorMap[name + team] = instructorList[idx];
  }
  
  Logger.log(salarySheet.getMaxRows());
  
  var notFoundInstructors = [];
  
  for(var row = 3; row < salarySheet.getMaxRows(); row++) {
    var range = salarySheet.getRange(row, 1, 1, COL_TOTAL_SALARY);
    var values = range.getValues();
    var formulas = range.getFormulas();
    
    if(!values[0][SALARY_SHEET_COL_NAME])
      break;
    var name = values[0][SALARY_SHEET_COL_NAME].toString();
    var team = values[0][SALARY_SHEET_COL_TEAM].toString(); 
    var course = values[0][SALARY_SHEET_COL_COURSE].toString();

    var instructor = instructorMap[name + team];
    var instructorRate = instructorRateMap[name + course];
    
    if(instructor) {
           
      salarySheet.getRange(row, SALARY_SHEET_COL_CODE + 1, 1, 1).setValue(instructor.code);
      
      if(instructorRate) {
        if(instructorRate.salary1) {
          salarySheet.getRange(row, SALARY_SHEET_COL_SALARY1 + 1, 1, 1).setValue(instructorRate.salary1);
        } else {
          salarySheet.getRange(row, SALARY_SHEET_COL_SALARY1 + 1, 1, 1).setValue("");
        }
        if(instructorRate.salary2) {
          salarySheet.getRange(row, SALARY_SHEET_COL_SALARY2 + 1, 1, 1).setValue(instructorRate.salary2);
        } else {
          salarySheet.getRange(row, SALARY_SHEET_COL_SALARY2 + 1, 1, 1).setValue("");
        }
        
        if(!instructorRate.salary1 && !instructorRate.salary2) {
          salarySheet.getRange(row, 1, 1, COL_TOTAL_SALARY).setBackground('orange');
        } else {
          salarySheet.getRange(row, 1, 1, COL_TOTAL_SALARY).setBackground('white');
        }
      } 
    } else {
      
      salarySheet.getRange(row, 1, 1, COL_TOTAL_SALARY).setBackground('red');
      notFoundInstructors.push(
        {
          name:name,
          team:team
        });
    }
  } 
}

// function showAlert(title, message) {
//   var ui = SpreadsheetApp.getUi(); // Same variations.
//
//   var result = ui.alert(
//      title,
//      message,
//     ui.ButtonSet.OK);
// }

function loadInstructorInfo(instructorSheet) {
  var instructorList = [];
  var instructorMap = {};
  for(var row = ROW_START; row < instructorSheet.getMaxRows(); row++) {
    var range = instructorSheet.getRange(row, 1, 1, COL_TOTAL_INSTRUCTOR);
    var values = range.getValues();
    Logger.log(values);
    if(!values[0][0])
      break;

    var name=values[0][1];
    var code=values[0][2];
    var email=values[0][3];
    var team=values[0][6];
    
    instructorList.push({
      code: code.toString().toUpperCase().trim(),
      name: name,
      email: email,
      team: team
    });
  }

  Logger.log(instructorList);
  return instructorList;
}

function onEdit() {
    var originalSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = originalSpreadsheet.getSheets();

    var instructorInfoSheet = sheets[SHEET_IDX_INSTRUCTOR_INFO];
    var instructorInfoRange = instructorInfoSheet.getDataRange();
    var instructorInfoValue = instructorInfoRange.getValues(); 
    
    var toViewSheet = sheets[SHEET_IDX_TO_VIEW];
    var toViewRange = toViewSheet.getDataRange();
    var toViewCell = toViewRange.getCell(5, 10); //J5
    var progressCell = toViewRange.getCell(6,10); //J6
    
    var toSendSheet = sheets[SHEET_IDX_TO_SEND];
    
    for (var r = 1; r < instructorInfoValue.length; r++){
        var row = instructorInfoValue[r];
        var code = row[2];
        if(toViewCell.getValue() == code){
            progressCell.setValue("Loading")
            var instrInfo = loadInstrInfo(code);
            var salary = loadSalary(code);
            
            printInstructorInfo(instrInfo,toViewSheet);
            printInstructorInfo(instrInfo,toSendSheet);
            
            printSalary(salary,toViewSheet);
            printSalary(salary,toSendSheet);
            progressCell.setValue("Done");
    }
  }
}

function loadInstrInfo(codeCheck){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets();
    var instructorInfoSheet = sheet[SHEET_IDX_INSTRUCTOR_INFO];
    var range = instructorInfoSheet.getDataRange();
    var values = range.getValues();
    for (var i = 2; i < values.length; i++){
        var code = values[i][2];
        if (code == codeCheck){
            var rowValues = values[i];
        }
    }
    return rowValues;
}

function loadSalary(codeCheck){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets();
    var salarySheet = sheet[SHEET_IDX_SALARY];
    var range = salarySheet.getDataRange();
    var values = range.getValues();
    var rowValuesArr = [];
    for (var i = 2; i < values.length ;i++){
        var code = values[i][1];
        if (code == codeCheck){
           rowValues = values[i];
           rowValuesArr.push(rowValues);
        }
    }
    return rowValuesArr
}

function printInstructorInfo(instructor, sheet){
    var cellName = sheet.getRange("B8");
    var cellTeam = sheet.getRange("B9");
    cellName.setValue(instructor[INSTRUCTOR_SHEET_COL_NAME]);
    cellTeam.setValue(instructor[INSTRUCTOR_SHEET_COL_TEAM]);
}

function printSalary(salary, sheet){
    var tempArr = [];
    var class_range = sheet.getRange("A14:A17");
    var salary_rate1_range = sheet.getRange("C14:C17");
    var salary_rate2_range = sheet.getRange("F14:F17");
    var salary1_range = sheet.getRange("D14:D17");
    var salary2_range = sheet.getRange("G14:G17");
    for (var i = 0; i < salary.length; i++){
        var temp = new Object();
        temp.class = salary[i][SALARY_SHEET_COL_CLASS];
        temp.salaryRate1 = salary[i][SALARY_SHEET_COL_SALARY_RATE1];
        temp.salaryRate2 = salary[i][SALARY_SHEET_COL_SALARY_RATE2];
        temp.salary1 = salary[i][SALARY_SHEET_COL_SALARY1];
        temp.salary2 = salary[i][SALARY_SHEET_COL_SALARY2];
        tempArr.push(temp);
    }

    class_range.clearContent();
    salary_rate1_range.clearContent();
    salary_rate2_range.clearContent();
    salary1_range.clearContent();
    salary2_range.clearContent();
    var i = ROW_NUM_14;
    var j = 0;
    while(i < ROW_NUM_18 && j < tempArr.length){
        var classRange = sheet.getRange(i,1);
        var salaryRate1Range = sheet.getRange(i,3);
        var salaryRate2Range = sheet.getRange(i,6);
        var salary1Range = sheet.getRange(i,4);
        var salary2Range = sheet.getRange(i,7);

        classRange.setValue(tempArr[j].class);
        salaryRate1Range.setValue(tempArr[j].salaryRate1);
        salaryRate2Range.setValue(tempArr[j].salaryRate2);
        salary1Range.setValue(tempArr[j].salary1);
        salary2Range.setValue(tempArr[j].salary2);
        i++;
        j++;
  }
 }
 
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
//          
//          var message = messageNoName + ": " +  instrInfo[1]; 
//          
//          var emailTo = "duytechkids@gmail.com";
//
//          // Create a new Spreadsheet and copy the current sheet into it.
//          var newSpreadsheet = SpreadsheetApp.create("Spreadsheet to export");
//          var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
//          
//          toSendSheet.copyTo(newSpreadsheet);
//          
//          // Find and delete the default "Sheet 1", after the copy to avoid triggering an apocalypse
//          newSpreadsheet.getSheetByName('Sheet1').activate();
//          newSpreadsheet.deleteActiveSheet();
//          
//        
//          var pdf = DriveApp.getFileById(newSpreadsheet.getId()).getAs('application/pdf').getBytes();
//          var attach = {fileName:'Payslips.pdf',content:pdf, mimeType:'application/pdf'};
//  
//            // Send the freshly constructed email 
//          MailApp.sendEmail(emailTo, subject, message, {attachments:[attach]});
//          
//           // Delete the wasted sheet we created, so our Drive stays tidy.
//          DriveApp.getFileById(newSpreadsheet.getId()).setTrashed(true);
//          Logger.log(newSpreadsheet.getUrl())
        }
    }
  }
  progressCell.setValue("Done")
}

 function sendThisPaySlips(){
   var ui = SpreadsheetApp.getUi();
   var uiSubjectPrompt = ui.prompt("Subject:",ui.ButtonSet.OK_CANCEL);
   var uiMsgPrompt = ui.prompt("Tin nhắn:", ui.ButtonSet.OK_CANCEL);
   var uiEmailPrompt = ui.prompt("Email:" , ui.ButtonSet.OK_CANCEL);
   var message = uiMsgPrompt.getResponseText();
   var subject = uiSubjectPrompt.getResponseText();
   var email = uiEmailPrompt.getResponseText();
   ui.alert('Subject: \n'+ subject + '\n'
              + 'Message : \n' + message +'\n'
              + 'Email: \n' + email);
   var response = ui.alert('Are you sure you want to continue?', ui.ButtonSet.OK_CANCEL);
   if (response == ui.Button.OK) {
     // User clicked "OK".
     var originalSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
     var sheets = originalSpreadsheet.getSheets();
     
     var toViewSheet = sheets[SHEET_IDX_TO_VIEW];
     var toViewRange = toViewSheet.getDataRange();
     
     var toSendSheet = sheets[SHEET_IDX_TO_SEND];
     
     
     // Create a new Spreadsheet and copy the current sheet into it.
     var newSpreadsheet = SpreadsheetApp.create("Spreadsheet to export");
     var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
     
     toSendSheet.copyTo(newSpreadsheet);
     
     // Find and delete the default "Sheet 1", after the copy to avoid triggering an apocalypse
     newSpreadsheet.getSheetByName('Sheet1').activate();
     newSpreadsheet.deleteActiveSheet();
     
     
     var pdf = DriveApp.getFileById(newSpreadsheet.getId()).getAs('application/pdf').getBytes();
     var attach = {fileName:'Payslips.pdf',content:pdf, mimeType:'application/pdf'};
     
     // Send the freshly constructed email 
     MailApp.sendEmail(email, subject, message, {attachments:[attach]});
     
     // Delete the wasted sheet we created, so our Drive stays tidy.
     DriveApp.getFileById(newSpreadsheet.getId()).setTrashed(true);
   } 
   else {
   //Users click Cancel 
   //Dimiss all
   }
}   
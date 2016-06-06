COL_TOTAL = 100;
ROW_START = 2;
ROW_START_1 = 9;

PRODUCTION = 0;

var COL_FROM_ADDR = 9;
var COL_TEMPLATE = 10;
var COL_CODE = 10;

var STR_NAME = "Name";

var STR_NAME_PATTERN = "[" + STR_NAME + "]"

// Sheet index
var SHEET_IDX_INSTRUCTOR_INFO = 0;
var SHEET_IDX_INSTRUCTOR_RATE = 1;
var SHEET_IDX_CLASS = 2;
var SHEET_IDX_SALARY = 3;
var SHEET_IDX_TO_VIEW = 4;
var SHEET_IDX_TO_SEND = 5;
var SHEET_IDX_CC = 7;

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
var SALARY_SHEET_COL_SB1 = 6;
var SALARY_SHEET_COL_SB2 = 9;
var SALARY_SHEET_COL_SALARY_RATE1 = 7;
var SALARY_SHEET_COL_SALARY_RATE2 = 10;
var SALARY_SHEET_COL_SALARY1 = 8
var SALARY_SHEET_COL_SALARY2 = 11
var SALARY_SHEET_COL_TOTAL_SALARY = 13;

// Total ROW
var INSTRUCTOR_INFO_TOTAL_ROWS = 66;
var SALARY_TOTAL_ROWS = 78;
var CC_SHEET_TEAM_ROW_IDX = 0

var KPI_COL_IN_TARGET_SHEET = 6;
var ROW_NUM_14 = 14;
var ROW_NUM_18 = 18;
var ROW_NUM_8 = 8;
var ROW_NUM_19 = 19;


function onOpen() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var fEasyMenuEntries = [ 
    {name: "Clear all", functionName: "clearAll"},
    {name: "Load Instructor Codes & Rates", functionName: "loadInstructorCodesAndRates"},
    {name: "Prepare Payslips", functionName: "preparePaySlips"}, 
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
          salarySheet.getRange(row, SALARY_SHEET_COL_SALARY_RATE1 + 1, 1, 1).setValue(instructorRate.salary1);
        } else {
          salarySheet.getRange(row, SALARY_SHEET_COL_SALARY_RATE1 + 1, 1, 1).setValue("");
        }
        if(instructorRate.salary2) {
          salarySheet.getRange(row, SALARY_SHEET_COL_SALARY_RATE2 + 1, 1, 1).setValue(instructorRate.salary2);
        } else {
          salarySheet.getRange(row, SALARY_SHEET_COL_SALARY_RATE2 + 1, 1, 1).setValue("");
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
  for(var row = 1; row < instructorSheet.getMaxRows(); row++) {
    var range = instructorSheet.getRange(row, 1, 1, COL_TOTAL_INSTRUCTOR);
    var values = range.getValues();
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
    for (var i = 1; i < values.length; i++){
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
    var sb1_range = sheet.getRange("B14:B17");
    var sb2_range = sheet.getRange("E14:E17");
    var salary_rate1_range = sheet.getRange("C14:C17");
    var salary_rate2_range = sheet.getRange("F14:F17");
    var salary1_range = sheet.getRange("D14:D17");
    var salary2_range = sheet.getRange("G14:G17");
    for (var i = 0; i < salary.length; i++){
        var temp = new Object();
        temp.class = salary[i][SALARY_SHEET_COL_CLASS];
        temp.sb1 = salary[i][SALARY_SHEET_COL_SB1];
        temp.sb2 = salary[i][SALARY_SHEET_COL_SB2];
        temp.salaryRate1 = salary[i][SALARY_SHEET_COL_SALARY_RATE1];
        temp.salaryRate2 = salary[i][SALARY_SHEET_COL_SALARY_RATE2];
        temp.salary1 = salary[i][SALARY_SHEET_COL_SALARY1];
        temp.salary2 = salary[i][SALARY_SHEET_COL_SALARY2];
        tempArr.push(temp);
    }

    class_range.clearContent();
    sb1_range.clearContent();
    sb2_range.clearContent();
    salary_rate1_range.clearContent();
    salary_rate2_range.clearContent();
    salary1_range.clearContent();
    salary2_range.clearContent();
    var i = ROW_NUM_14;
    var j = 0;
    while(i < ROW_NUM_18 && j < tempArr.length){
        var classRange = sheet.getRange(i,1);
        var sb1_range = sheet.getRange(i,2);
        var sb2_range = sheet.getRange(i,5);
        var salaryRate1Range = sheet.getRange(i,3);
        var salaryRate2Range = sheet.getRange(i,6);
        var salary1Range = sheet.getRange(i,4);
        var salary2Range = sheet.getRange(i,7);

        classRange.setValue(tempArr[j].class);
        sb1_range.setValue(tempArr[j].sb1);
        sb2_range.setValue(tempArr[j].sb2);
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
   
    var ccSheet = sheets[SHEET_IDX_CC];
    var ccRange = ccSheet.getDataRange();
    var ccValues = ccRange.getValues(); 
    
    var toViewSheet = sheets[SHEET_IDX_TO_VIEW];
    var toViewRange = toViewSheet.getDataRange(); 
    var toViewCell = toViewRange.getCell(5, 10);
    var progressCell = toViewRange.getCell(6,10);
   
    var ccMap = loadCCEmailList(ccSheet);
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
  
            var newSpreadsheet = SpreadsheetApp.create("Lương tháng 5/2016 của " + name );
                        
            toSendSheet.copyTo(newSpreadsheet);
            
            newSpreadsheet.getSheetByName('Sheet1').activate();
            newSpreadsheet.deleteActiveSheet();
            
            var pdf = DriveApp.getFileById(newSpreadsheet.getId()).getAs('application/pdf').getBytes();
            var attach = {fileName:'Payslips.pdf',content:pdf, mimeType:'application/pdf'};
            for(row = 8; row <= 19; row++) {
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
                        var ccList = ccMap[team].toString();
                        GmailApp.sendEmail(email,subject,'', {htmlBody: body,from: from_addr ,attachments:[attach], cc:ccList});
                        }
                    }
                }   
            } 
          }
          
        }
    }       
    progressCell.setValue("Done")
}


function sendThisPaySlips(){
    var originalSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = originalSpreadsheet.getSheets();

    var instructorInfoSheet = sheets[SHEET_IDX_INSTRUCTOR_INFO];
    var instructorInfoRange = instructorInfoSheet.getDataRange();
    var instructorInfoValue = instructorInfoRange.getValues(); 
  
    var ccSheet = sheets[SHEET_IDX_CC];
    var ccRange = ccSheet.getDataRange();
    var ccValues = ccRange.getValues(); 
    
    var toViewSheet = sheets[SHEET_IDX_TO_VIEW];
    var toViewRange = toViewSheet.getDataRange();
    var toViewCell = toViewRange.getCell(5, 10); //J5
    var progressCell = toViewRange.getCell(6,10); //J6
    
    var toSendSheet = sheets[SHEET_IDX_TO_SEND];
  
    var ccMap = loadCCEmailList(ccSheet);
    for (var r = 1; r < instructorInfoValue.length; r++){
        var row = instructorInfoValue[r];
        var code = row[2];
        var name = row[1];
        var email = row[3];
        var team = row[6];
        if(toViewCell.getValue() == code){
          progressCell.setValue("Loading")
          var draftMsgs = GmailApp.getDraftMessages();
          
          var newSpreadsheet = SpreadsheetApp.create("Lương tháng 5/2016 của " + name);
          
          toSendSheet.copyTo(newSpreadsheet);
          
          newSpreadsheet.getSheetByName('Sheet1').activate();
          newSpreadsheet.deleteActiveSheet();
          
          var pdf = DriveApp.getFileById(newSpreadsheet.getId()).getAs('application/pdf').getBytes();
          var attach = {fileName:'Payslips.pdf',content:pdf, mimeType:'application/pdf'};
          for(row = ROW_NUM_8; row <= ROW_NUM_19; row++) {
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
                  var ccList = ccMap[team].toString();
                  GmailApp.sendEmail(email,subject,'', {htmlBody: body,from: from_addr ,attachments:[attach], cc:ccList});
                }
              }
            }   
          }
          progressCell.setValue("Done");  
        }
    }
}
      
function preparePaySlips() {
  var aliases = GmailApp.getAliases()
  var rule1 = SpreadsheetApp.newDataValidation().requireValueInList(aliases).build();
  
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadSheet.getSheets();
  var toViewSheet = sheets[SHEET_IDX_TO_VIEW];
  
  var cell1 = toViewSheet.getRange(9, COL_FROM_ADDR);
  cell1.setDataValidation(rule1);
  
  //Browser.msgBox("Load from addresses completed!")
  
  /* Build dropdown list (Value In List Rule) */
  var draftMsgs = GmailApp.getDraftMessages();
  var msgSubjects = [];
  for (i = 0; i < draftMsgs.length; i++) {
    msgSubjects.push(draftMsgs[i].getSubject())
  }
  
  var rule2 = SpreadsheetApp.newDataValidation().requireValueInList(msgSubjects).build();
  var toViewSheet = sheets[SHEET_IDX_TO_VIEW];
 
  var cell2 = toViewSheet.getRange(9, COL_TEMPLATE);
  cell2.setDataValidation(rule2);
  
  /* Load instructor codes */
  var instructorInfoSheet = sheets[SHEET_IDX_INSTRUCTOR_INFO];
  var instructorInfo = loadInstructorInfo(instructorInfoSheet);
  var instructorInfoList = [];
  var cell3 = toViewSheet.getRange(5,COL_CODE);
  for (var i = 1; i < instructorInfo.length;i++){
    instructorInfoList.push(instructorInfo[i].code);
  }
  var rule3 = SpreadsheetApp.newDataValidation().requireValueInList(instructorInfoList).build();
  cell3.setDataValidation(rule3);
  
  Browser.msgBox("Load completed!")
}

function loadCCEmailList(ccSheet) {
  var ccRange = ccSheet.getDataRange();
  var ccValues = ccRange.getValues();
  
  var ccEmailMap = {};
  
  for(var col = 0; col < ccValues.length; col++) {
    var team = ccValues[CC_SHEET_TEAM_ROW_IDX][col];
    var email_list = []
    for(var row = 1; row < ccValues[0].length; row++){
      var email = ccValues[row][col];
      email_list.push(email);
    }
    ccEmailMap[team] = email_list;
    if(!team) break;
  } 
  //Logger.log(ccEmailMap['CFA'])
  return ccEmailMap
}
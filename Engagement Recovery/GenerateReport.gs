function generateReport() {
  var engagementCode = showPrompt()
  if(engagementCode == null){
    return;
  }
  var EC = getTaskData(engagementCode)
  if(EC == null){
    SpreadsheetApp.getUi().alert('No such engagement code!');
  }
  else{
    showDialog(EC, engagementCode)
  }
}

function showPrompt() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  
  var result = ui.prompt(
    'Generate Report',
    'Please enter the engagement code:',
    ui.ButtonSet.OK_CANCEL);
  
  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    // User clicked "OK".
    return text
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
    return null;
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
    Logger.log("Close Pressed");
  }
}

function showDialog(EC, engagementCode) {
  var output = createTables(EC,engagementCode)
  output.setWidth(600);
  output.setHeight(1500);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .showModalDialog(output, engagementCode.toString() + ' Report');
}

function getTaskData(engagementCode){
  var EC = {}
  var engagementCodes2D = erSheet.getRange(2,erEngagementCodeColumn,erLastRow-1,1).getValues()
  var engagementCodes = []
  for(var i = 0; i < engagementCodes2D.length; i++){
    engagementCodes = engagementCodes.concat(engagementCodes2D[i]);
  }
  var engagementEmployees = null;
  for(var i = 0; i < engagementCodes.length; i++){
    if(engagementCodes[i] == engagementCode){
      var ecRow = i+2;
      engagementEmployees = erSheet.getRange(i+2, erEmployeesColumn).getValue().split(", ")
      }
  }
  if(engagementEmployees == null){
    return null
  }
  for(var i=0;i<engagementEmployees.length;i++){
    var task = {}
    var employeeRate = employees[engagementEmployees[i]].Rate
    var employeePosition = employees[engagementEmployees[i]].Position
    var employeeSSID = employees[engagementEmployees[i]].ID
    var employeeSS = SpreadsheetApp.openByUrl(employeeSSID)
    var engagementSheet = employeeSS.getSheetByName(engagementCode);
    var taskDescriptions = engagementSheet.getRange(timesheetTaskStart,timesheetDescriptionColumn,engagementSheet.getLastRow()-timesheetTaskStart, 1).getValues().filter(String);
    //Logger.log("Task Descriptions: "+ taskDescriptions)
    var taskHours = engagementSheet.getRange(timesheetTaskStart,timesheetYearTotalColumn,engagementSheet.getLastRow()-timesheetTaskStart, 1).getValues().filter(String);
    //Logger.log("Task Hours: "+taskHours)
    var taskCount = engagementSheet.getRange(timesheetTaskStart,timesheetHTColumn,engagementSheet.getLastRow()-timesheetTaskStart, 1).getValues().filter(String).length;
    //Logger.log("Task Count: "+taskCount)
    for(var j=0;j<taskCount;j++){
      task[j+1] = {
        Description: taskDescriptions[j],
        Hours: taskHours[j]
      }
    }
    //Logger.log(engagementEmployees[i] +": "+task)
    EC[engagementEmployees[i]] = {
      Task: task,
      TaskCount: Object.keys(task).length,
      Rate: employeeRate,
      Position: employeePosition,
      Row: ecRow,
      Column: erEngagementCodeColumn
    }
  }
  return EC;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
  .getContent();
}


function createTables(EC,engagementCode) {
  var output = HtmlService.createHtmlOutput('<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:Calibri, Candara, Segoe, "Segoe UI", Optima, Arial, sans-serif;}table, td {border: 1px solid black;border-spacing: 0px;}td{padding:5px;}h3,h4 {margin-top: 5px;margin-bottom: 5px;} h3{text-decoration: underline;}</style></head><body>');
  var employeesInvolved = Object.keys(EC)
  var cost = 0
  for(var k =0; k<employeesInvolved.length;k++){
    var employee = employeesInvolved[k]
    output.append('<h3>'+ employee + '</h3>');
    output.append('<h4>'+ EC[employee].Position + ' : $' + EC[employee].Rate.toString() + '/hr</h4>');
    
    if(EC[employee].TaskCount != 0){
      var hours = 0
      output.append('<table><tbody><tr><td></td><td><b>Description</b></td><td><b>Hours</b></td></tr>'); 
      for (var i = 0; i < EC[employee].TaskCount; i++) {
        output.append('<tr>')
        for (var j = 0; j < 3; j++) {
          if(j == 0){
            output.append('<td>  '+(i+1).toString()+'  </td>')            
          }
          else if(j == 1){
            output.append('<td>'+EC[employee].Task[i+1].Description+'</td>')
          }
          else{
			output.append('<td>'+EC[employee].Task[i+1].Hours+'</td>')
            hours += Number(EC[employee].Task[i+1].Hours)
          }
        }
        output.append('</tr>')
      }
      output.append('</tbody></table>')
      output.append('<p>Total Hours: '+hours.toString()+'</p>')
      output.append('<p>Cost: $'+(hours*EC[employee].Rate).toString()+'</p>')
      output.append('<p>________________________________________________________________</p>')
      cost += Number(hours*EC[employee].Rate)
    }
    else{
      output.append('<p>Nothing Recorded</p>')
    }
  }
  output.append('<p>Total Cost: $'+cost.toString()+'</p>')
  output.append('<input type="button" value="Close" onclick="google.script.host.close()">')
  output.append('<input type="button" value="Save"  onclick="google.script.run.createReport('+EC[employee].Row.toString()+','+EC[employee].Column.toString()+')"></body></html>')
  return output
}


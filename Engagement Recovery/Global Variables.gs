var startTime = Date.now()
var KYCSS = SpreadsheetApp.openById("1_-yXCzeHl_OOrjPYsKCr4JGwTxcM8ScXGtIk5YHzeeM");
var KYCSheet = KYCSS.getSheetByName("KYC Form Responses");
var KYCValues = KYCSheet.getDataRange().getValues()
var KYCLastColumn = KYCSheet.getLastColumn();
var clientInfo = {}
var count = 0
for(i=0;i<KYCLastColumn+1;i++){
  if(count == 4){
    break;
  }
  if(KYCValues[0][i] == "Client Name"){
    var KYCClientNameColumn = i+1;
    count++;
  }
  else if(KYCValues[0][i] == "Client Type"){
    var KYCClientTypeColumn = i+1;
    count++;
  }
  else if(KYCValues[0][i] == "Client Identification"){
    var KYCClientIDColumn = i+1;
    count++;
  }
  else if(KYCValues[0][i] == "Financial Year End"){
    var KYCFYColumn = i+1;
    count++;
  }
}
var KYCLastRow = KYCSheet.getRange("A1:A").getValues().filter(String).length;
for(i=2;i<KYCLastRow+1;i++){
  clientInfo[KYCSheet.getRange(i,KYCClientNameColumn).getValue()] = {Type: KYCSheet.getRange(i,KYCClientTypeColumn).getValue(),
                                                                     ID: KYCSheet.getRange(i,KYCClientIDColumn).getValue(),
                                                                     FY: KYCSheet.getRange(i,KYCFYColumn).getValue()
                                                                       };
}
var KYCDuration = ((Date.now() - startTime)/1000)
Logger.log('KYC Duration: ' + KYCDuration.toString() + 's' )
startTime = Date.now()

var erSS = SpreadsheetApp.openById("1fENbME3jdio-kIkuHp9H6rLr_9tzXL39jNMjFx70lxc");
var erSheet = erSS.getSheetByName("Engagement Recovery");
var erValues = erSheet.getDataRange().getValues();
var erLastRow = erSheet.getRange("A1:A").getValues().filter(String).length;
var erLastColumn = erSheet.getLastColumn()
var count = 0;
for(i=0;i<erLastColumn+1;i++){
  if(count == 10){
    break;
  }
  if(erValues[0][i] == "Client Name"){
    var erClientNameColumn = i+1
  }
  else if(erValues[0][i] == "Engagement Type"){
    var erEngagementTypeColumn = i+1
  }
  else if(erValues[0][i] == "Engagement Code"){
    var erEngagementCodeColumn = i+1
  }
  else if(erValues[0][i] == "Check one or more options"){
    var erRoutineEngagementColumn = i+1
  }
  else if(erValues[0][i] == "Contract Value"){
    var erContractValueColumn = i+1
  }
  else if (erValues[0][i] == "Projected Budget"){
    var erProjectedBudgetColumn = i+1
  }
  else if (erValues[0][i] == "Time Charge"){
    var erTimeChargeColumn = i+1
  }
  else if (erValues[0][i] == "Profit/Loss"){
    var erPNLColumn = i+1
  }
  else if(erValues[0][i] == "Check the employees involved"){
    var erEmployeesColumn = i+1
  }
  else if(erValues[0][i] == "Status"){
    var erStatusColumn = i+1
  }
}
var erDuration = ((Date.now() - startTime)/1000)
Logger.log('ER Duration: ' + erDuration.toString() + 's' )
startTime = Date.now()

var employeesSheet = erSS.getSheetByName("List of Employees");
var employeesLastRow = employeesSheet.getRange("A1:A").getValues().filter(String).length;
var employeesLastColumn = employeesSheet.getLastColumn();
count = 0
for(i=1;i<employeesLastColumn+1;i++){
  if(count == 5){
    break;
  }
  if(employeesSheet.getRange(1,i).getValue() == "Name"){
    var employeesNameColumn = i
    count++;
  }
  else if(employeesSheet.getRange(1,i).getValue() == "Email"){
    var employeesEmailColumn = i
    count++;
  }
  else if(employeesSheet.getRange(1,i).getValue() == "Position"){
    var employeesPositionColumn = i
    count++;
  }
  else if(employeesSheet.getRange(1,i).getValue() == "Hourly Rate"){
    var employeesRateColumn = i
    count++;
  }
  else if(employeesSheet.getRange(1,i).getValue() == "ID"){
    var employeesIDColumn = i
    count++;
  }
}
var employees = {}
for(i=2;i<employeesLastRow+1;i++){
  employees[employeesSheet.getRange(i,employeesNameColumn).getValue()] = {Email: employeesSheet.getRange(i,employeesEmailColumn).getValue(), Position: employeesSheet.getRange(i,employeesPositionColumn).getValue(), Rate: employeesSheet.getRange(i,employeesRateColumn).getValue(), ID: employeesSheet.getRange(i,employeesIDColumn).getValue()}
}
var employeeNames = Object.keys(employees);

var employeesDuration = ((Date.now() - startTime)/1000)
Logger.log('Employees Duration: ' + employeesDuration.toString() + 's' )
startTime = Date.now()

var raSS = SpreadsheetApp.openById("1Xe4Njmz7REO3sEdR9u3c0MKSxIuN6JzIF4LlTQ3hRJ8");
var raSheet = raSS.getSheetByName("FY2018");
var raValues = raSheet.getDataRange().getValues()
var raLastColumn = raSheet.getLastColumn();
var raRowValues2D = raSheet.getRange("A1:A").getValues();
var raRowValues = []
for(var i = 0; i < raRowValues2D.length; i++){
  raRowValues = raRowValues.concat(raRowValues2D[i]);
  if(raRowValues[i].indexOf("FEBRUARY x FEBRUARY") != -1){
    var janRow = i
    }
  else if(raRowValues[i].indexOf("MARCH x MARCH") != -1){
    var febRow = i
    }
  else if(raRowValues[i].indexOf("APRIL x APRIL") != -1){
    var marRow = i
    }
  else if(raRowValues[i].indexOf("MAY x MAY") != -1){
    var aprRow = i
    }
  else if(raRowValues[i].indexOf("JUNE x JUNE") != -1){
    var mayRow = i
    }
  else if(raRowValues[i].indexOf("JULY x JULY") != -1){
    var junRow = i
    }
  else if(raRowValues[i].indexOf("AUGUST x") != -1){
    var julRow = i
    }
  else if(raRowValues[i].indexOf("SEPTEMBER x SEPTEMBER") != -1){
    var augRow = i
    }
  else if(raRowValues[i].indexOf("OCTOBER x OCTOBER") != -1){
    var sepRow = i
    }
  else if(raRowValues[i].indexOf("NOVEMBER x NOVEMBER") != -1){
    var octRow = i
    }
  else if(raRowValues[i].indexOf("DECEMBER x DECEMBER") != -1){
    var novRow = i
    }
  else if(raRowValues[i].indexOf("END x END") != -1){
    var decRow = i
    }
}
var raLastRow = decRow + 1
/*Logger.log(janRow)
Logger.log(febRow)
Logger.log(marRow)
Logger.log(aprRow)
Logger.log(mayRow)
Logger.log(junRow)
Logger.log(julRow)
Logger.log(augRow)
Logger.log(sepRow)
Logger.log(octRow)
Logger.log(novRow)
Logger.log(decRow)*/

var count = 0
for(i=1;i<3;i++){
  if(count == 5){
    break;
  }
  for(j=0;j<raValues[i].length;j++){
    if(count == 6){
      break;
    }
    if(raValues[i][j] == "Name"){
      var raNameColumn = j+1
    }
    /*if(raValues[i][j] == "Contract Value"){
      var raContractValueColumn = j+1
    }*/
    else if(raValues[i][j] == "Engagement Code"){
      var raEngagementCodeColumn = j+1
    }
    else if(raValues[i][j] == "Accounting"){
      var raAccountingStartColumn = j+1
      var raAccountingEndColumn = j+8
      var raAccountingColumnCount = raAccountingEndColumn - raAccountingStartColumn + 1
    }
    else if(raValues[i][j] == "Company Secretariat"){
      var raCSRangeStartColumn = j+1
      var raCSRangeEndColumn = j+4
      var raCSColumnCount = raCSRangeEndColumn - raCSRangeStartColumn + 1
    }
    else if (raValues[i][j] == "Tax"){
      var raTaxRangeStartColumn = j+1
      var raTaxRangeEndColumn = j+5
      var raTaxColumnCount = raTaxRangeEndColumn - raTaxRangeStartColumn + 1
    }
  }
}
var raDuration = ((Date.now() - startTime)/1000)
Logger.log('RA Duration: ' + raDuration.toString() + 's' )
startTime = Date.now()

var timesheetFolder = DriveApp.getFolderById('1rBtwgsfUJ2NSeKSSpBP79_M558_CPSem')
var timesheetTemplate = DriveApp.getFileById('1riCcAwe8LQ9nXSIJ_5wZhQuvjhwGkA_4RfQCeTnmrIM')
var timesheetTemplateSS = SpreadsheetApp.openById("1riCcAwe8LQ9nXSIJ_5wZhQuvjhwGkA_4RfQCeTnmrIM")
var timesheetTemplateSheet = timesheetTemplateSS.getSheetByName("Template Sheet");
var timesheetTemplateValues = timesheetTemplateSheet.getDataRange().getValues()
var timesheetHeaderRow = 8
var timesheetTaskStart = 9
var timesheetLastColumn = timesheetTemplateSheet.getLastColumn();
var timesheetHeaderValues2D = timesheetTemplateSheet.getRange(timesheetHeaderRow, 1, 1, timesheetLastColumn).getValues()
var timesheetHeaderValues = []
for(var i = 0; i < timesheetHeaderValues2D.length; i++){
  timesheetHeaderValues = timesheetHeaderValues.concat(timesheetHeaderValues2D[i]);
}
var count = 0
for(i=0;i<timesheetHeaderValues.length;i++){
  if(count == 3){
    break
  }
  if(timesheetHeaderValues[i] == "#"){
    var timesheetHTColumn = i+1
    count++;
  }
  else if(timesheetHeaderValues[i] == "Description"){
    var timesheetDescriptionColumn = i+1
    count++;
  }
  else if(timesheetHeaderValues[i] == "Year Total"){
    var timesheetYearTotalColumn = i+1
    count++;
  } 
}

var count = 0
for(i=0;i<timesheetTemplateValues.length;i++){
  if(count == 2){
    break;
  }
  for(j=0;j<timesheetTemplateValues[i].length;j++){
    if(count == 2){
      break;
    }
    if(timesheetTemplateValues[i][j] == "Name of Employee:"){
      var timesheetEmployeeTBFRow = i+1;
      var timesheetEmployeeTBFColumn = j+3;
      count++;
    }
    else if(timesheetTemplateValues[i][j] == "Name of Company:"){
      var timesheetCompanyTBFRow = i+1;
      var timesheetCompanyTBFColumn = j+3;
      count++;
    }
  }
}

var timesheetDuration = ((Date.now() - startTime)/1000)
Logger.log('Timesheet Duration: ' + timesheetDuration.toString() + 's' )
startTime = Date.now()

var reportFolder = DriveApp.getFolderById('1YE5oYJiwrNYww31nHPNL9rVg91MEv2xE')
var reportTemplate = DriveApp.getFileById('1gGmPhny-Os6_Qn3r4Q2-1xcnM-OuEdMoZSTlKjN8rsA')

function onFormSubmit(){
  Logger.log('onFormSubmit')
  var clientName = erSheet.getRange(erLastRow,erClientNameColumn).getValue();
  formattingAndFormula();
  createEngagementCode();
  var engagementType = erSheet.getRange(erLastRow,erEngagementTypeColumn).getValue();
  if(engagementType == "Routine"){
    addRA(clientName)
    writeRoutineEngagementCode(clientName);
  }
  else if(engagementType == "Non-routine"){
    writeNonRoutineEngagementCode(clientName);
  }
  createEmployeeSS(clientName)
}

function updateEmployeesOnForm(){
  var form = FormApp.openById("1vG_SnkgV3NkQm-v-ttkO3vozVxslslm4RszglLmotdE");
  var employeeList = form.getItemById("1177902667").asCheckboxItem();
  employeeList.setChoiceValues(employeeNames);  
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Mace')
  .addItem('Calculate Time Charge', 'calculateTimeCharge')
  .addItem('Generate Report', 'generateReport')  
  .addToUi();
}
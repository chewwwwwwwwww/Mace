function calculateTimeCharge() {
  var totalTimeCharge = 0
  var statuses2D = erSheet.getRange(2,erStatusColumn,erLastRow-1,1).getValues()
  var statuses = []
  for(var i = 0; i < statuses2D.length; i++){
    statuses = statuses.concat(statuses2D[i]);
  }
  var engagementCodes2D = erSheet.getRange(2,erEngagementCodeColumn,erLastRow-1,1).getValues()
  var engagementCodes = []
  for(var i = 0; i < engagementCodes2D.length; i++){
    engagementCodes = engagementCodes.concat(engagementCodes2D[i]);
  }
  for(var i = 0; i < statuses.length; i++){
    if(statuses[i] == "In Progress" || statuses[i] == "Completed"){
      var engagementEmployees = erSheet.getRange(i+2, erEmployeesColumn).getValue().split(", ")
      var engagementCode = engagementCodes[i]
      for(var j=0;j<engagementEmployees.length;j++){
        var employeeRate = employees[engagementEmployees[j]].Rate
        var employeeHours = 0
        var employeeSSID = employees[engagementEmployees[j]].ID
        var employeeSS = SpreadsheetApp.openByUrl(employeeSSID)
        var engagementSheet = employeeSS.getSheetByName(engagementCode);
        var engagementLastColumn = engagementSheet.getLastColumn();
        var engagementSheetValues = engagementSheet.getDataRange().getValues()
        var count = 0
        for(k=0;k<engagementSheetValues.length;k++){
          if(count == 1){
            break;
          }
          if(engagementSheetValues[k][engagementLastColumn-1] == "Final Total"){
            var totalsRow = k
            count++;
          }
        }
        employeeHours = engagementSheet.getRange(totalsRow,engagementLastColumn).getValue()
        Logger.log(employeeHours)
        var employeeTimeCharge = employeeHours * employeeRate
        Logger.log(employeeTimeCharge)
        totalTimeCharge += employeeTimeCharge
      }
      erSheet.getRange(i+2,erTimeChargeColumn).setValue(totalTimeCharge)
      Logger.log(totalTimeCharge)
    }
    
  }
  Logger.log('Calculate Time Charge Duration: ' + ((Date.now() - startTime)/1000).toString() + 's' )
}
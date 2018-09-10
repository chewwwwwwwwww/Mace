function createEmployeeSS(clientName){
  Logger.log('createEmployeeSS')
  var employeesInvolved = erSheet.getRange(erLastRow, erEmployeesColumn).getValue().split(", ")
  for(i=0;i<employeesInvolved.length;i++){
    var employeeSS = '';
    var files = timesheetFolder.getFiles();
    var employeeName = employeesInvolved[i]
    Logger.log(employeeName)
    while (files.hasNext()) {
      var file = files.next();
      if(file.getName() == employeeName){
        Logger.log('fileExists')
        employeeSS = SpreadsheetApp.openByUrl(file.getUrl());
        employeeFile = file
      }
    }
    if(employeeSS == ''){
      Logger.log('fileDoesNotExist')
      employeeFile = timesheetTemplate.makeCopy();
      employeeFile.setName(employeeName)
      for(j=0;j<employeeNames.length;j++){
        if(employeeNames[j] == employeeName){
          employeesSheet.getRange(j+2, employeesIDColumn).setValue(employeeFile.getUrl())
        }
      }
      employeeSS = SpreadsheetApp.openByUrl(employeeFile.getUrl())
    }
    employeeTemplateSheet = employeeSS.getSheetByName('Template Sheet');
    employeeSS.setActiveSheet(employeeTemplateSheet);
    var newEngagementSheet = employeeSS.duplicateActiveSheet();
    newEngagementSheet.setName(clientInfo[clientName].EC);
    newEngagementSheet.getRange(timesheetEmployeeTBFRow,timesheetEmployeeTBFColumn).setValue(employeeName)
    newEngagementSheet.getRange(timesheetCompanyTBFRow,timesheetCompanyTBFColumn).setValue(clientName)
    employeeFile.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.EDIT)
    employeeFile.setShareableByEditors(false)
    //employeeFile.removeEditor('chewhtbrandon@gmail.com')
    //employeeFile.setOwner(employees[employeeName].Email)
  }
}

function dummyFunction(){
}
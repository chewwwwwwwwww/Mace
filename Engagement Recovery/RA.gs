function addRA(clientName){
  clientInfo[clientName].CV = erSheet.getRange(erLastRow, erContractValueColumn).getValue()
  var routineEngagement = erSheet.getRange(erLastRow, erRoutineEngagementColumn).getValue().split(", ");
  var rowNumber = determineRow(clientInfo[clientName].FY)
  raSheet.insertRowAfter(rowNumber)
  raSheet.getRange(rowNumber, raNameColumn).setValue(clientName)
  raSheet.getRange(rowNumber, raContractValueColumn).setValue(clientInfo[clientName].CV)
  raSheet.getRange(rowNumber, raEngagementCodeColumn).setValue(clientInfo[clientName].EC)
  if(routineEngagement.indexOf('Accounting') == -1){
    raSheet.getRange(rowNumber, raAccountingStartColumn, 1, raAccountingColumnCount).setBackground('black')
  }
  if(routineEngagement.indexOf('Company Secretarial Work') == -1){
    raSheet.getRange(rowNumber, raCSRangeStartColumn, 1, raCSColumnCount).setBackground('black')
  }
  if(routineEngagement.indexOf('Taxation') == -1){
    raSheet.getRange(rowNumber, raTaxRangeStartColumn, 1, raTaxColumnCount).setBackground('black')
  }
  raSheet.getRange(rowNumber, 1, 1, raLastColumn).setBorder(true,true,true,true,true,true)
}
    
function determineRow(a) {
  Logger.log(typeof a)
  switch(a) {
    case 'January':
      return janRow;
    case 'February':
      return febRow;
    case 'March':
      return marRow;
    case 'April':
      return aprRow;
    case 'May':
      return mayRow;
    case 'June':
      return junRow;
    case 'July':
      return julRow;
    case 'August':
      return augRow;
    case 'September':
      return sepRow;
    case 'October':
      return octRow;
    case 'November':
      return novRow;
    case 'December':
      return decRow;
  }
}

function dummyFunction(){
}
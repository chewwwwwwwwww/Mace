function formattingAndFormula(){
  Logger.log('FNF')
  erSheet.getRange(1,erContractValueColumn,erSheet.getMaxRows(),1).setNumberFormat("$#,##0.00;$(#,##0.00)");
  erSheet.getRange(1,erProjectedBudgetColumn,erSheet.getMaxRows(),1).setNumberFormat("$#,##0.00;$(#,##0.00)");
  for(k=2;k<erLastRow+1;k++){
    var projectedBudgetString = "=(I"+k+"*150 + J"+k+"*80 + K"+k+"*50 + L"+k+"*20)";
    erSheet.getRange(k, erProjectedBudgetColumn, 1, 1).setValue(projectedBudgetString);
    var PNLString = "=(H"+k+"-N"+k+")";
    erSheet.getRange(k, erPNLColumn, 1, 1).setValue(PNLString);
  }
  erSheet.getRange(1,erTimeChargeColumn,erSheet.getMaxRows(),1).setNumberFormat("$#,##0.00;$(#,##0.00)"); 
  erSheet.getRange(1,erPNLColumn,erSheet.getMaxRows(),1).setNumberFormat("$#,##0.00;$(#,##0.00)");
}

function createEngagementCode(){
  Logger.log('CEC')
  var startTime = Date.now()
  var clientName = erSheet.getRange(erLastRow, erClientNameColumn).getValue().toString()
  var trimmedClientName = clientName.replace("."," ")
  trimmedClientName = trimmedClientName.replace("  "," ")
  trimmedClientName.replace(/[^a-zA-Z0-9 ]/g, '');
  var clientNameArrT = trimmedClientName.split(" ");
  var clientNameLength = clientNameArrT.length;
  var engagementCode = ""
  if(clientInfo[clientName].Type == "Individual"){
    var clientNameArr = []
    for(i=0;i<3;i++){
      if(clientNameArrT[i] != null){
        clientNameArr.push(clientNameArrT[i]);
      }
      else{
        break;
      }
    }
    var engagementCode = ''
    for(i=0;i<clientNameArr.length;i++){
      engagementCode += clientNameArr[i][0]
    }
    engagementCode += clientInfo[clientName].ID.substr(clientInfo[clientName].ID.length - 2)
  }
  else if(clientInfo[clientName].Type == "Corporate"){
    var clientNameArr = []
    for(i=0;i<2;i++){
      if(clientNameArrT[i] != null){
        clientNameArr.push(clientNameArrT[i]);
      }
      else{
        break;
      }
    }
    var clientNameWords = clientNameArr.length;
    if(isNaN(clientNameArr[0]) && clientNameArr[0] == clientNameArr[0].toUpperCase()) {
      //Logger.log("ALLCAPS")
      engagementCode = clientNameArr[0].substr(0,5)
    }
    /*else if(!isNaN(clientNameArr[0]) && clientNameWords == 1.0){
    Logger.log("NUMBER 1")
    engagementCode = clientNameArr[0].substr(0,5)
    }
    else if(!isNaN(clientNameArr[0]) && clientNameWords == 2.0){
    Logger.log("NUMBER 2")
    engagementCode = clientNameArr[0].substr(0,5)
    var charLeft = 5 - engagementCode.length
    Logger.log(charLeft)
    engagementCode += clientNameArr[1].substr(0,charLeft)
    }*/
    else if(clientNameWords == 1){
      //Logger.log("1 WORD")
      engagementCode = clientNameArr[0].substr(0,5)
    }
    else if(clientNameWords == 2){
      //Logger.log("2 WORDS")
      engagementCode = clientNameArr[0].substr(0,4)
      var charLeft = 5 - engagementCode.length
      engagementCode += clientNameArr[1].substr(0,charLeft)
    }
  }
  engagementCode += "."
  var clientNameValues = erSheet.getRange(2, erClientNameColumn, erLastRow-1).getValues()
  var clientCount = 0;
  for(i=0;i<clientNameValues.length;i++){
    if(clientNameValues[i] == clientName){
      clientCount++;
    }
  }
  if(clientCount < 10){
    engagementCode += "00" + clientCount.toString()
  }
  else if(clientCount >= 10){
    engagementCode += "0" + clientCount.toString()
  }
  else{
    engagementCode += clientCount.toString()
  }
  engagementCode = engagementCode.toUpperCase()
  erSheet.getRange(erLastRow, erEngagementCodeColumn).setValue(engagementCode);
  clientInfo[clientName].EC = engagementCode
  Logger.log('Duration: ' + ((Date.now() - startTime)/1000).toString() + 's' )
}

function writeRoutineEngagementCode(clientName){
  Logger.log('writeRoutine')
  //raSheet.getRange(?, ?).setValue(clientInfo[clientName].engagementCode);
  erSheet.getRange(erLastRow, erEngagementCodeColumn).setValue(clientInfo[clientName].EC);
}

function writeNonRoutineEngagementCode(clientName){
  Logger.log('writeNonRoutine')
  erSheet.getRange(erLastRow, erEngagementCodeColumn).setValue(clientInfo[clientName].EC);
}

//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

function dummyFunction(){
}

function writeAllEngagementCodes(){
  var startTime = Date.now()
  for(z=2;z<erLastRow+1;z++){
    var clientName = erSheet.getRange(z, erClientNameColumn).getValue().toString()
    var trimmedClientName = clientName.replace("."," ")
    trimmedClientName = trimmedClientName.replace("  "," ")
    trimmedClientName.replace(/[^a-zA-Z0-9 ]/g, '');
    var clientNameArrT = trimmedClientName.split(" ");
    var clientNameLength = clientNameArrT.length;
    var engagementCode = ""
    if(clientInfo[clientName].Type == "Individual"){
      var clientNameArr = []
      for(i=0;i<3;i++){
        if(clientNameArrT[i] != null){
          clientNameArr.push(clientNameArrT[i]);
        }
        else{
          break;
        }
      }
      var engagementCode = ''
      for(i=0;i<clientNameArr.length;i++){
        engagementCode += clientNameArr[i][0]
      }
      engagementCode += clientInfo[clientName].ID.substr(clientInfo[clientName].ID.length - 2)
    }
    else if(clientInfo[clientName].Type == "Corporate"){
      var clientNameArr = []
      for(i=0;i<2;i++){
        if(clientNameArrT[i] != null){
          clientNameArr.push(clientNameArrT[i]);
        }
        else{
          break;
        }
      }
      var clientNameWords = clientNameArr.length;
      if(isNaN(clientNameArr[0]) && clientNameArr[0] == clientNameArr[0].toUpperCase()) {
        //Logger.log("ALLCAPS")
        engagementCode = clientNameArr[0].substr(0,5)
      }
      /*else if(!isNaN(clientNameArr[0]) && clientNameWords == 1.0){
      Logger.log("NUMBER 1")
      engagementCode = clientNameArr[0].substr(0,5)
      }
      else if(!isNaN(clientNameArr[0]) && clientNameWords == 2.0){
      Logger.log("NUMBER 2")
      engagementCode = clientNameArr[0].substr(0,5)
      var charLeft = 5 - engagementCode.length
      Logger.log(charLeft)
      engagementCode += clientNameArr[1].substr(0,charLeft)
      }*/
      else if(clientNameWords == 1){
        //Logger.log("1 WORD")
        engagementCode = clientNameArr[0].substr(0,5)
      }
      else if(clientNameWords == 2){
        //Logger.log("2 WORDS")
        engagementCode = clientNameArr[0].substr(0,4)
        var charLeft = 5 - engagementCode.length
        engagementCode += clientNameArr[1].substr(0,charLeft)
      }
    }
    engagementCode += "."
    var clientNameValues = erSheet.getRange(2, erClientNameColumn, erLastRow-1).getValues()
    var clientCount = 0;
    for(i=0;i<clientNameValues.length;i++){
      if(clientNameValues[i] == clientName){
        clientCount++;
      }
    }
    if(clientCount < 10){
      engagementCode += "00" + clientCount.toString()
    }
    else if(clientCount >= 10){
      engagementCode += "0" + clientCount.toString()
    }
    else{
      engagementCode += clientCount.toString()
    }
    engagementCode = engagementCode.toUpperCase()
    erSheet.getRange(z, erEngagementCodeColumn).setValue(engagementCode);
    clientInfo[clientName].EC = engagementCode
  }
  Logger.log('Duration: ' + ((Date.now() - startTime)/1000).toString() + 's' )
}
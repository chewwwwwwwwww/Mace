var ss = SpreadsheetApp.openById("1_-yXCzeHl_OOrjPYsKCr4JGwTxcM8ScXGtIk5YHzeeM");
var clientsSheet = ss.getSheetByName("KYC Form Responses");
var values = ss.getDataRange().getValues()
var KYCLastRow = ss.getRange("A1:A").getValues().filter(String).length;

for(i=0;i<values.length;i++){
    if(values[0][i] == "Client Name"){
      var clientNamesCol = i+1;
    }
    if(values[0][i] == "Form Responses"){
      var formResponseCol = i+1
    }
  }

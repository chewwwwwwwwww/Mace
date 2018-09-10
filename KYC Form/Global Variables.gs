var ss = SpreadsheetApp.openById("1_-yXCzeHl_OOrjPYsKCr4JGwTxcM8ScXGtIk5YHzeeM");
var clientsSheet = ss.getSheetByName("KYC Form Responses");
var values = ss.getDataRange().getValues()

for(i=0;i<values.length;i++){
  for(j=0;j<values[i].length;j++){
    if(values[i][j] == "Client Name"){
      var clientNamesCol = j+1;
    }
  }
}
function updateEngagementForm(){
  // call your form
  var form = FormApp.openById("INSERT FORM ID HERE");
   
  //connect to the drop-down item 
  var clientList = form.getItemById("INSERT LIST ITEM HERE").asListItem();

  // grab the values in the first column of the sheet - use 2 to skip header row 
  var clientValues = clientsSheet.getRange(2, clientNamesCol, clientsSheet.getMaxRows() - 1).getValues();

  var finalClients = [];

  // convert the array ignoring empty cells (removing all the empty rows)
  for(var i = 0; i < clientValues.length; i++)    
    if(clientValues[i][0] != "")
      finalClients[i] = clientValues[i][0];

  // populate the drop-down item with the array data
  clientList.setChoiceValues(finalClients);  
}

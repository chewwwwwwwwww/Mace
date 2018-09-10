function updateEngagementForm(){
  // call your form
  var form = FormApp.openById("1vG_SnkgV3NkQm-v-ttkO3vozVxslslm4RszglLmotdE");
   
  //connect to the drop-down item 
  var clientList = form.getItemById("742955637").asListItem();

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
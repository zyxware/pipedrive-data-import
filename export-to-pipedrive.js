const DATA_SHEET_NAME = "Person data";
const CONFIG_SHEET_NAME = "Config";
const PIPEDRIVE_API = "B1";
const FIELDS_TO_BE_MERGED = "B2"
const EMAIL_ID = "B3";
const ORGANIZATION = "B4";
const TIME_ZONE = "B5";

const FIELD_KEYS_SHEET_NAME = "PersonField keys"

var person_data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
var config_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
var field_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FIELD_KEYS_SHEET_NAME);

var pipedriveAPIKey = config_sheet.getRange(PIPEDRIVE_API).getValue();
var fieldsToBeMerged = config_sheet.getRange(FIELDS_TO_BE_MERGED).getValue();
var columnAlphabet = config_sheet.getRange(EMAIL_ID).getValue();
var timezone = config_sheet.getRange(TIME_ZONE).getValue();
var fieldsToBeMergedArray = fieldsToBeMerged.split(",");
var headerRow = person_data.getRange(1, 1, 1, person_data.getLastColumn()).getValues()[0];

var columnIndex = columnAlphabet.charCodeAt(0) - 65 + 1;

//Fetching key to the corresponding label
var labelsInSheet = field_sheet.getRange(2, 1, field_sheet.getLastRow(), 1).getValues().flat();
var keysInSheet = field_sheet.getRange(2, 2, field_sheet.getLastRow(), 1).getValues().flat();
var typesInSheet = field_sheet.getRange(2, 3, field_sheet.getLastRow(), 1).getValues().flat();
var organizations = {};
var labelToKeyMap = {};
for (var i = 0; i < labelsInSheet.length; i++) {
  labelToKeyMap[labelsInSheet[i]] = keysInSheet[i];
}


function checkAndUpdateContacts() {
  var dataRange = person_data.getDataRange();
  var dataValues = dataRange.getValues();

  for (var i = 1; i < dataValues.length; i++) { //dataValues = Object that contains all the rows arrays
    
    var columnLabelsWithValues = [];
    var row = dataValues[i];
    var isUpdated = row[1]; // "Record uploaded?" column is the second column (index 1)
    
    if (isUpdated !== "Done") {
      
      var rowRange = person_data.getRange(i+1, 1, 1, person_data.getLastColumn());
      var rowValues = rowRange.getValues()[0];
      for (var j = 2; j < rowValues.length; j++) { //rowValues = Array of all columns in that row
        if (rowValues[j]) {
          columnLabelsWithValues.push(headerRow[j]);
        }
      }

      var rowData = row.slice(0, 28); // Extract the first 28 columns of data
      
      var email = person_data.getRange(i + 1, columnIndex).getValue(); //Fetching the email to check whether that contact exist
      
      // Check if the email exists as a contact in Pipedrive
      var contactData = checkContactExistsInPipedrive(email);
      

      try{

        if (contactData) { //If contact exist, then update
          updateContactInPipedrive(contactData,rowData,columnLabelsWithValues);
          person_data.getRange(i+1 ,2).setValue("Done");
        } else {
          addContactInPipedrive(rowData,columnLabelsWithValues);
          person_data.getRange(i+1 ,2).setValue("Done");
        }

      } catch (error){
        
        Logger.log(`Error adding/updating contact in Pipedrive: ${error}`);
        person_data.getRange(i+1 ,2).setValue("Failed");
      
      }
    }
  }
}



function checkContactExistsInPipedrive(email) {
  
  var url = `https://api.pipedrive.com/v1/persons/search?term=${email}&exact_match=1&api_token=${pipedriveAPIKey}`;
  var response = UrlFetchApp.fetch(url);
  var data = JSON.parse(response.getContentText());

  if (data.data && data.data.items && data.data.items.length > 0) {
    return data.data.items[0].item.id;
  } else {
    return false; // Person with the given email not found
  }
}



function updateContactInPipedrive(person_id, rowData, columnLabelsWithValues) {//Updating an existing contact in Pipedrive

  let keysForLabels = fetchKeysForLabels(columnLabelsWithValues);
  var updatedValues = [];

  var payload = {};
  for(k = 2; k < rowData.length; k++){
    if(rowData[k]){
      updatedValues.push(rowData[k])
    }
  }
  
  for (var i = 0; i < keysForLabels.length; i++) {
    if(fetchTypeForFieldKey(keysForLabels[i]) === "Multiple options" || fetchTypeForFieldKey(keysForLabels[i]) === "Single option"){ //Need to check whether the current field type is 'Multiple options' or 'Single option'

    var optionValues = updatedValues[i].split(",");

      for (var a = 0; a < optionValues.length; a++){
        let customFieldOptions = checkForExistingOption(optionValues[a],keysForLabels[i]); //if the provided value exist
        var noOptionExist = customFieldOptions[0];
        var customFieldId = customFieldOptions[1];
        if(noOptionExist){
          addOptionToCustomField(keysForLabels[i],optionValues[a],customFieldId);
          console.log(noOptionExist);
        }
      }
    }

    if(fetchTypeForFieldKey(keysForLabels[i]) === "Date"){
      updatedValues[i] = convertDateFormat(updatedValues[i]);
    }

    if(fetchTypeForFieldKey(keysForLabels[i]) === "Organization"){
      if(!organizations){
        fetchAllOrganization();
      }
      checkOrganizationExist
    }

    for(x = 0; x < fieldsToBeMergedArray.length; x++){
      if(columnLabelsWithValues[i]==fieldsToBeMergedArray[x]){ //Checks whether the label current column label matches any of the field name mentioned in the config sheet
        var personData = getContactCustomFieldData(person_id);
        var previousOptions = personData[labelToKeyMap[columnLabelsWithValues[i]]];
        updatedValues[i] = previousOptions + ',' + updatedValues[i]
        console.log(updatedValues[i]);
      }
    }
    payload[keysForLabels[i]] = updatedValues[i];
  }
  
  var url = `https://api.pipedrive.com/v1/persons/${person_id}?api_token=${pipedriveAPIKey}`;
  //console.log(payload);

  var options = {
    method: "PUT",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };

  var response = UrlFetchApp.fetch(url, options);
  console.log(response);

}

function addContactInPipedrive(rowData,columnLabelsWithValues){//Adding new contacts inside Pipedrive

  let keysForLabels = fetchKeysForLabels(columnLabelsWithValues);
  var updatedValues = [];

  var payload = {};
  for(k = 2; k < rowData.length; k++){
    if(rowData[k]){
      updatedValues.push(rowData[k])
    }
  }
  
  for (var i = 0; i < keysForLabels.length; i++) {//Need to add an IF condition for multiple options
    if(keysForLabels[i] == "first_name"){
      payload["name"] = updatedValues[i] + ' ';
      //console.log(updatedValues[i]);
    }

    if(keysForLabels[i] == "last_name"){
      payload["name"] += updatedValues[i]
      //console.log(updatedValues[i]);
    }

    if(fetchTypeForFieldKey(keysForLabels[i]) === "Date"){
      updatedValues[i] = convertDateFormat(updatedValues[i]);
    }

    if(fetchTypeForFieldKey(keysForLabels[i]) === "Multiple options" || fetchTypeForFieldKey(keysForLabels[i]) === "Single option"){ //Need to check whether the current field type is 'Multiple options' or 'Single option'

    var optionValues = updatedValues[i].split(",");

      for (var a = 0; a < optionValues.length; a++){
        let customFieldOptions = checkForExistingOption(optionValues[a],keysForLabels[i]); //if the provided value exist
        var noOptionExist = customFieldOptions[0];
        var customFieldId = customFieldOptions[1];
        if(noOptionExist){
          addOptionToCustomField(keysForLabels[i],optionValues[a],customFieldId);
          console.log(noOptionExist);
        }
      }
    }

    
    payload[keysForLabels[i]] = updatedValues[i];
  }

  var url = `https://api.pipedrive.com/v1/persons?api_token=${pipedriveAPIKey}`;

  var options = {
    method: "POST",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };

  var response = UrlFetchApp.fetch(url, options);
  console.log(response);

}



function fetchKeysForLabels(columnLabelsWithValues) {//Fetching field_keys from  PersonField Key sheet

  var keysForLabels = [];
  for (var i = 0; i < columnLabelsWithValues.length; i++) {
    var label = columnLabelsWithValues[i];
    if (labelToKeyMap[label]) {
      keysForLabels.push(labelToKeyMap[label]);
    }
  }
  return keysForLabels;
}


function fetchTypeForFieldKey(fieldKey){
  for (var i = 1; i < typesInSheet.length; i++){
    var fieldkey = keysInSheet[i];
    var type = typesInSheet[i];

    if (fieldkey === fieldKey) {
      return type;
    }
  }
}


function checkForExistingOption(option,fieldkey){ //Option ID can be fetched from here as well
  var url = `https://api.pipedrive.com/v1/personFields?api_token=${pipedriveAPIKey}`;

  var pfResponse = UrlFetchApp.fetch(url);
  var pfData = JSON.parse(pfResponse.getContentText());
  var noOptionExist = true;
  var customFieldId = '';
  
  for(i = 0; i < pfData.data.length ; i++){
    if(pfData.data[i].key == fieldkey){
      for(j = 0; j < pfData.data[i].options.length ; j++){
        if(pfData.data[i].options[j].label == option){ //option exist
          noOptionExist = false;
        }
        customFieldId = pfData.data[i].id;
      }
    }
  }
  var returnValue = [noOptionExist,customFieldId];
  return returnValue;
}


function addOptionToCustomField(customFieldKey,optionName,customFieldId){
  
  var dataArray = getAllCustomFieldsData();
  var customFieldObject = dataArray.find(function(data) {return data.key === customFieldKey;});
  var optionArray = optionName.split(',');
  var outputArray = [];

  for (var i = 0; i < optionArray.length; i++) {
    var obj = {
      label: optionArray[i]
    };
    outputArray.push(obj);
  }

  var newAppendedOptions = customFieldObject.options.concat(outputArray);
  
  var apiurl = `https://api.pipedrive.com/v1/personFields/${customFieldId}?api_token=${pipedriveAPIKey}`

  let payload = {
    name: customFieldObject.name,
    add_visible_flag: true
  };

  payload.options = newAppendedOptions;

  var options = {
    method: "PUT",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(apiurl, options);
  const responseData = JSON.parse(response.getContentText());
  console.log("addOptionToCustomFieldsData" + responseData);
}


function getAllCustomFieldsData(){
  var url = `https://api.pipedrive.com/v1/personFields?api_token=${pipedriveAPIKey}`;

  try {
    var response = UrlFetchApp.fetch(url);
    var responseData = JSON.parse(response.getContentText());

    if (responseData.data) {
      return responseData.data;
    } else {
      Logger.log("Failed to fetch person data.");
      return null;
    }
  } catch (error) {
    Logger.log(`Error fetching person data: ${error}`);
    return null;
  }
}


function getContactCustomFieldData(person_id) {
  var apiUrl = `https://api.pipedrive.com/v1/persons/${person_id}?api_token=${pipedriveAPIKey}`;

  try {
    var response = UrlFetchApp.fetch(apiUrl);
    var responseData = JSON.parse(response.getContentText());

    if (responseData.data) {
      return responseData.data;
    } else {
      Logger.log("Failed to fetch person data.");
      return null;
    }
  } catch (error) {
    Logger.log(`Error fetching person data: ${error}`);
    return null;
  }
}


function fetchAllOrganization(){
  var url = `https://api.pipedrive.com/v1/organizations?api_token=${pipedriveAPIKey}`;

  var response = UrlFetchApp.fetch(url);
  var data = JSON.parse(response.getContentText());

  for(i = 0; i < data.data.length; i++){
    organizations[data.data[i].name]=data.data[i].id;
  }

}

function checkOrganizationExist(org_name){
  if(organizations[org_name]){
    return organizations[org_name];
  }
}


function importPipedriveCustomFieldData(){
  var url = `https://api.pipedrive.com/v1/personFields?api_token=${pipedriveAPIKey}`;

  var response = UrlFetchApp.fetch(url);
  var data = JSON.parse(response.getContentText());

  for(i = 0; i < data.data.length; i++){
    var personFieldName = data.data[i].name;
    var personFieldKey = data.data[i].key;
    var newRow = [personFieldName,personFieldKey];

    var personFieldkeys = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PersonField keys');
    personFieldkeys.appendRow(newRow);
  }
}


function convertDateFormat(date) {
  
  // Split the input date into day, month, and year
  var dateParts = date.split('/');
  
  // Create a new Date object with the parsed components
  var date = new Date(dateParts[2], dateParts[1] - 1, dateParts[0]);
  
  // Format the date as yyyy-MM-dd HH:mm:ss
  var formattedDate = Utilities.formatDate(date, timezone, "yyyy-MM-dd HH:mm:ss");
  
  return formattedDate;
}


function onOpen(){
  const ui = SpreadsheetApp.getUi()
  ui.createMenu('Pipedrive')
    .addItem('Download Custom Field', 'importPipedriveCustomFieldData')
    .addItem('Export to Pipedrive', 'checkAndUpdateContacts')
    .addToUi();
}

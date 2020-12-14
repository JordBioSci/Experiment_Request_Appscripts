function updateForm(){
  // call your form and connect to the drop-down item
  var form = FormApp.openById("1aVPNjei3xYpMSf6mdNeOLb7OPUJmcf6bPMeZeGIU-hY");
  
  var experimentorsNameList = form.getItemById("846978913").asListItem();
  var requestorList = form.getItemById("1505636757").asListItem();
// identify the sheet where the data resides needed to populate the drop-down
  var ss = SpreadsheetApp.openById("1TMFJop6w_3ciLNdC7FJzadkx0dXcGLDGk9F3Byx5Ao8");
  var names = ss.getSheetByName("Experimentors_details");

  // grab the values in the first column of the sheet - use 2 to skip header row
  var namesValues = names.getRange(2, 1, names.getMaxRows(),3).getValues();

  Logger.log(namesValues);
  var experimentors = [];

  // convert the array ignoring empty cells
  for(var i = 0; i < namesValues.length; i++)
    if(namesValues[i][0] != "" && namesValues[i][2]=="Active")
    {
      experimentors.push(namesValues[i][0]);
    }
  
  Logger.log(experimentors);

  // populate the drop-down with the array data
  experimentorsNameList.setChoiceValues(experimentors);
  requestorList.setChoiceValues(experimentors);
  //populate the experiment names
 
}


function getEmailIDs(){
  // call your form and connect to the drop-down item
  var form = FormApp.openById("1aVPNjei3xYpMSf6mdNeOLb7OPUJmcf6bPMeZeGIU-hY");
  
  var experimentorsNameList = form.getItemById("846978913").asListItem();
  var requestorList = form.getItemById("1505636757").asListItem();
  // identify the sheet where the data resides needed to populate the drop-down
  var ss = SpreadsheetApp.openById("1TMFJop6w_3ciLNdC7FJzadkx0dXcGLDGk9F3Byx5Ao8");
  var names = ss.getSheetByName("Experimentors_details");
  // grab the values in the first column of the sheet - use 2 to skip header row
  var namesValues = names.getRange(2, 1, names.getMaxRows(),3).getValues();
  Logger.log(namesValues)
  return namesValues;
}

function getEmailIDForName(namesValues,name){
  // convert the array ignoring empty cells
  for(var i = 0; i < namesValues.length; i++)
    if(namesValues[i][0] != "" && namesValues[i][0]==name)
    {
      return namesValues[i][1];
    }
}

function testEmailID() {
  var emails = getEmailIDs();
  var email = getEmailIDForName(emails,"Sarija Janardhan");
  Logger.log(email);
}



function populateExperimentNames(){
  // call your form and connect to the drop-down item
  var form = FormApp.openById("1aVPNjei3xYpMSf6mdNeOLb7OPUJmcf6bPMeZeGIU-hY");
  //do a inspect source and search of data-item-id to get this id
  var experimentNames = form.getItemById("541116349").asListItem();
  
  // identify the sheet where the data resides needed to populate the drop-down
  var ss = SpreadsheetApp.openById("1TMFJop6w_3ciLNdC7FJzadkx0dXcGLDGk9F3Byx5Ao8");
  var names = ss.getSheetByName("Form Responses 1");//this is the list of experiments

  // grab the values in the first column of the sheet - use 2 to skip header row
  var namesValues = names.getRange(2, 2, names.getMaxRows() ,3).getValues();

  var experimentors = [];

  // convert the array ignoring empty cells
  for(var i = 0; i < namesValues.length; i++)
    if(namesValues[i][0] != "")
    {
      experimentors.push(namesValues[i][0]);
    }

  // populate the drop-down with the array data
  experimentNames.setChoiceValues(experimentors);
  
}

function onFormSubmit(e){
  //create a sheet on submit
  //Get information from form and set as variables
  var spreadsheetresponse = SpreadsheetApp.openById("1yxMH3I0D1KDPXrn_cebhrgDb8LcnUUSUdjihv1VIzUA");
  var timestampsheet = spreadsheetresponse.getSheetByName("Form Responses 1");//this is the list of experiments
  
  var timevalues = timestampsheet.getDataRange().getValues();
  var lastRow = timestampsheet.getLastRow();
  var lastColumn = timestampsheet.getLastColumn();
  var timestamp = timestampsheet.getRange(lastRow, lastColumn-4);
  //var timestamp = sheet.getRange(lastRow, lastColumn-4);
  // Logger.log(timestamp.getValue());
  
  var frm = FormApp.getActiveForm().getItems();
  var requsterName = e.response.getResponseForItem(frm[0]).getResponse();
  var experimenterName = e.response.getResponseForItem(frm[1]).getResponse();
  var typeOfExperiment = e.response.getResponseForItem(frm[2]).getResponse();
  var startDate = e.response.getResponseForItem(frm[3]).getResponse();
  var strainsUsed = e.response.getResponseForItem(frm[4]).getResponse();
  var experimentalRationale = e.response.getResponseForItem(frm[5]).getResponse();
  
  var experiment_column_names = []
  //add default experiment columns that need to be before the specified columns
   //1. to get the columns needed for the new sheet
    var config_spread_sheet = SpreadsheetApp.openById("1TMFJop6w_3ciLNdC7FJzadkx0dXcGLDGk9F3Byx5Ao8");
    var sheet_1 = config_spread_sheet.getSheetByName("Form Responses 1");
    //grab all the experiment names from the sheet
    var all_experiment_names = sheet_1.getRange(2,2,sheet_1.getMaxRows(),1).getValues();
     
    for(var i = 0; i < all_experiment_names.length; i++){
      if(all_experiment_names[i][0] != "" && all_experiment_names[i][0]==typeOfExperiment)
      {
        //get the column values for that specific experiment
        var experiment_columnNames = sheet_1.getRange(i+2,2,1,2).getValues();
        var col_names = experiment_columnNames[0][1].split(",");
        for (var col_index = 0; col_index < col_names.length ; col_index++){
          experiment_column_names.push(col_names[col_index])
        }
        
        break;
      }
    }
  
  //add default columns that need to be after speific columns
  experiment_column_names.push("Comments")

  experimenterNameForFileName = experimenterName.replace(/\s/g, "_");
  typeOfExperimentForFileName = typeOfExperiment.replace(/\s/g, "_");
  
  var sheetName = getFileName(typeOfExperimentForFileName , experimenterNameForFileName);
  
  Logger.log("Total Column Length") 
  Logger.log(experiment_column_names.length)
  
  Logger.log("Printing columns names") 
  Logger.log(experiment_column_names)
  var wendyID = "1yc1Yf9fGTyB7MUVs4gvv3z0h8ALud_Fz";
  var saraID = "1eo6SwLBnO6DtEo_0AgIHhvBr_fNa95V5";
  var benID = "1Y-pPlLJoR_saCW_XqnJFgP_ULuY3A5ie";
  var michaelID = "15iEYvoinQv9XT-irILJ8drnfT3prveHt";
  var sarijaID = "1jbwQgnoSKp91PW-qybyjWm9oW2p65cZv";
  var nicholeID = "1E-YTePw0Zq8NmS9Zj9WLTpEwbxxOLDcb";
//  //
//  //////************************************************************************************
  if(experimenterName == 'Wendy Hughes'){ 
    Logger.log("creating file for Wendy")
    folder = DriveApp.getFolderById(wendyID)
  }
  else if(experimenterName == 'Sara Bachmeier'){ 
    Logger.log("creating file for Sara")
    folder = DriveApp.getFolderById(saraID)
  }
  else if(experimenterName == 'Ben Clasen'){ 
    Logger.log("creating file for Ben")
    folder = DriveApp.getFolderById(benID)
  }
  else if(experimenterName == 'Michael Millican'){ 
    Logger.log("creating file for Micheal")
    folder = DriveApp.getFolderById(michaelID)
  }
  else if(experimenterName == 'Nichole Dopkins'){ 
    Logger.log("creating file for Nicole")
    folder = DriveApp.getFolderById(nicholeID)
  }
  else{
    Logger.log("creating file for Sarija")
    folder = DriveApp.getFolderById(sarijaID)
  }
  
  var emailIDs = getEmailIDs();
  
  var experimenterEmailID = getEmailIDForName(emailIDs , experimenterName);
  
  var requesterEmailID = getEmailIDForName(emailIDs, requsterName);
  
  var emailSubject  = "Experiment Requested to "+ experimenterName + "-  File Name " + sheetName
  
  var emailBody = "Hi " + experimenterName + "," + "\n\n The following experiment has been requested by " + requsterName + "."  
  + "\n\nRequester Name - " +  requsterName
  + "\nExperiment Type -" + typeOfExperiment
  + "\nRequested Start Date - " + startDate
  + "\nStrains used - " + strainsUsed
  + "\n\n A file named "+ sheetName +" has been created for your experiment at " + folder.getUrl();
  + "\n\n Thanks."
  
  
  Logger.log("Requester email id " + requesterEmailID);
  Logger.log("Experimenter email id " + experimenterEmailID);
  
  Logger.log("Sheet name "+ sheetName)
  Logger.log("Number of experiment columns "+ experiment_column_names.Length)
  Logger.log("ColumnsNames name "+ experiment_column_names)
  //create a new spread sheet with the name
  var spreadSheet = SpreadsheetApp.create(sheetName)
  var mainSheet = spreadSheet.getSheetByName("Sheet1");
  mainSheet.setName("Experiment_Data");
  mainSheet.getRange(1,1,1,experiment_column_names.length).setBackgroundRGB(119, 247, 153);
  mainSheet.getRange(1,1,1,experiment_column_names.length).setFontWeight(3)
  mainSheet.getRange(1,1,1,experiment_column_names.length).setValues([experiment_column_names])
  var statusSheet = spreadSheet.insertSheet("Status", 1)
  statusSheet.getRange(1, 1).setValue("Requester Name");
  statusSheet.getRange(1, 2).setValue(requsterName);
  statusSheet.getRange(2, 1).setValue("Experimenter Name");
  statusSheet.getRange(2, 2).setValue(experimenterName);
  statusSheet.getRange(3, 1).setValue("Type of Experiment");
  statusSheet.getRange(3, 2).setValue(typeOfExperiment);
  statusSheet.getRange(4, 1).setValue("StartDate");
  statusSheet.getRange(4, 2).setValue(startDate);
  statusSheet.getRange(5, 1).setValue("Strains used");
  statusSheet.getRange(5, 2).setValue(strainsUsed);
  statusSheet.getRange(6, 1).setValue("Status (Start/In-Progress/Failed/Complete)");
  //statusSheet.getRange(6, 2).setValue("Start");
  statusSheet.getRange(7, 1).setValue("ExperimentalObjective");
  statusSheet.getRange(7,2).setValue(experimentalRationale);
   
   var choiceSheet = spreadSheet.insertSheet("Choices", 2)
   choiceSheet.getRange(1, 1).setValue("Started");
   choiceSheet.getRange(2, 1).setValue("In-Progress");
   choiceSheet.getRange(3, 1).setValue("Completed");
   choiceSheet.getRange(4, 1).setValue("Failed");
   //var dynamicList = choiceSheet.getSheetByName('Choices').getRange('A1:A4');   // set to your sheet and range
   var dynamicList = choiceSheet.getRange('A1:A4'); // set to your sheet and range
   var arrayValues = dynamicList.getValues();
   //define the dropdown/validation rules
   var rangeRule = SpreadsheetApp.newDataValidation().requireValueInList(arrayValues);
  // set the dropdown validation for the row
   statusSheet.getRange(6,2).setDataValidation(rangeRule); // set range to your range
   statusSheet.getRange(6, 2).setValue("Started");

  DriveApp.getFileById(spreadSheet.getId()).moveTo(folder);
  
  //send the email notification
  sendemail(experimenterEmailID,requesterEmailID + ",michael.millican@jordbioscience.com,sarija@jordbioscience.com" ,emailSubject,emailBody)
}

function getFileName(typeOfExperiment, experimentName){
  var getDateFormatForFileName = Utilities.formatDate(new Date(),"GMT", "yyyy-MM-dd-HH-mm-ss")
  var sheetName = "EXP_" + experimentName + "_" + typeOfExperiment + "_" + getDateFormatForFileName
  return sheetName;
}

function sendemail(to,cc,subject,body){
    MailApp.sendEmail({
    to:to,
    cc:cc,
    subject: subject,
    body: body
  });



}

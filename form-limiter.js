/**
 *  ----------------------
 *   Google Forms Automatisation and Notification
 *  ----------------------
 * 
 *  license: MIT
 *  language: Google Apps Script
 *  author: Emanuel Lorenz, Amit Agarwal
 *  email: emanuellorenz@yahoo.de, amit@labnol.org
 * 
 */


/*SIGN UP FORM SCRIPT - SKI OG FJELLSPORT*/
/*ONLY CHANGE THE NEXT FOUR LINES*/
RESPONSE_COUNT   =  "100";  /*MAX NUMBER OF SIGN UPS*/
FORM_OPEN_DATE   =  "2020-01-22 11:41";  /*DATE WHEN THE FORM IS OPEN - FORMAT: YYYY-MM-DD HH:MM */
FORM_CLOSE_DATE  =  "2020-01-22 11:43";  /*DATE WHEN FORM IS CLOSED AND SEND TO THE TRIP COORDINATOR - FORMAT: YYYY-MM-DD HH:MM */
TRIP_COORDINATOR_EMAIL =  "YOUR.EMAIL@HERE.com";  /*CHANGE TO EMAIL OF ONE TRIP COORDINATOR*/

/* Don not change anything below this line*/ 

/* Create a spreadsheet for the form*/  
FORM_NAME = FormApp.getActiveForm().getTitle();
SPREADSHEET = SpreadsheetApp.create(FORM_NAME);
form = FormApp.openById(FormApp.getActiveForm().getId());
formURL = FormApp.getActiveForm().getPublishedUrl();

/* Initialize the form, setup time based triggers */
function Initialize() {
  
deleteTriggers_();
form.setDestination(FormApp.DestinationType.SPREADSHEET, SPREADSHEET.getId()); 
  
  if ((FORM_OPEN_DATE !== "") && 
      ((new Date()).getTime() < parseDate_(FORM_OPEN_DATE).getTime())) { 
    ScriptApp.newTrigger("openForm")
    .timeBased()
    .at(parseDate_(FORM_OPEN_DATE))
    .create();
  }
  
  if (FORM_CLOSE_DATE !== "") { 
    ScriptApp.newTrigger("closeForm")
    .timeBased()
    .at(parseDate_(FORM_CLOSE_DATE))
    .create(); 
  }
  
  if (RESPONSE_COUNT !== "") { 
    ScriptApp.newTrigger("checkLimit")
    .forForm(FormApp.getActiveForm())
    .onFormSubmit()
    .create();
  }
  
  /* Send email to trip coordinator with sign-up-link and start time */
  MailApp.sendEmail(TRIP_COORDINATOR_EMAIL, "Sign-up form for your Ski og Fjellsport trip: " + FORM_NAME, "Dear trip coordinator, \nwe just created a sign-up link for your trip. It will open at " + FORM_OPEN_DATE + " and close at " + FORM_CLOSE_DATE + ". \nThe link for the sign-up form is: " + formURL + " \nPlease include that link together with the opening time and date in your final trip email to our email system .\n\nAfter the form opened you will get an email that confirms that. If not, please contact us. \nLater, right after the link is closed, you will get an excel spreadsheet with the details of all participants via email. \n\nFor more information check out our trip coordinator lounge (https://ntnui.no/skiogfjellsport/tripcoordinator-lounge/). \n\nAll the best, \nSki og Fjellsport");  
}

/* Delete all existing Script Triggers */
function deleteTriggers_() {  
  var triggers = ScriptApp.getProjectTriggers();  
  for (var i in triggers) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

/* Send a mail to the form owner when the form status changes */
function informUser_(subject) {
  var formURL = FormApp.getActiveForm().getPublishedUrl();
  MailApp.sendEmail(TRIP_COORDINATOR_EMAIL, subject,  "Dear Trip coordinator, \nthe sign-up form for your trip just opened: "+ formURL + "\n\nAll the best, \nSki og Fjellsport");  
}

/* Allow Google Form to Accept Responses */
function openForm() {
  var form = FormApp.getActiveForm();
  form.setAcceptingResponses(true);
  informUser_("Your Ski og Fjellsport Sign-up form is now accepting responses");
}

/* Close the Google Form, Stop Accepting Responses */
function closeForm() {  
  var FORM_NAME = FormApp.getActiveForm().getTitle();
  var form = FormApp.getActiveForm();
  form.setAcceptingResponses(false);
  deleteTriggers_();
  var EMAIL_BODY = "Dear trip coordinator, \nthe sign-up form for your trip "+ FORM_NAME + " just closed. \nAttached is the participant list for your trip. \n\nAll the best, \nSki og Fjellsport,"
  MailApp.sendEmail(TRIP_COORDINATOR_EMAIL, "Participant list for your Ski og Fjellsport Trip: " + FORM_NAME, EMAIL_BODY, {attachments: convertSpreadsheet()});
}

/* If Total # of Form Responses >= Limit, Close Form */
function checkLimit() {
  if (FormApp.getActiveForm().getResponses().length >= RESPONSE_COUNT ) {
    closeForm();
  }  
}

/* Parse the Date for creating Time-Based Triggers */
function parseDate_(d) {
  return new Date(d.substr(0,4), d.substr(5,2)-1, 
                  d.substr(8,2), d.substr(11,2), d.substr(14,2));
}

function convertSpreadsheet() {
  var spreadsheetId = SPREADSHEET.getId();
  var file          = DriveApp.getFileById(spreadsheetId);
  var url           = 'https://docs.google.com/spreadsheets/d/'+spreadsheetId+'/export?format=xlsx';
  var token         = ScriptApp.getOAuthToken();
  var response      = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' +  token
    }
  });

 var fileName = (SPREADSHEET.getName()) + '.xlsx';
 return blobs   = [response.getBlob().setName(fileName)];
  }

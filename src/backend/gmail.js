
/*
TODO:
Add welcome email
Add favicon


*/

var userProperties = PropertiesService.getUserProperties(); // This is to allow retrieval of stored info

function doGet() {
    return HtmlService.createHtmlOutputFromFile('Index')
        // This setting may expose to to cross site scripting but will allow your app to work anywhere
        //.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        ///.setFaviconUrl(iconUrl) 
        .setTitle("Easy Gmail Scheduler")
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getDraftSubjects() { //Returns array of draft subjects and ids
    var subjects = [];
    var drafts = GmailApp.getDraftMessages();
    for (var i in drafts) {
        subjects.push(drafts[i].getSubject());
    }
    return subjects;
}

function sendScheduledEmails(){
    var scheduledEmails = JSON.parse(userProperties.getProperty('scheduledData'));
    var nowPlus15 = Date.parse(new Date()) + (15*60*1000); /// 15 minutes extra may not be nessicary
    // Email use a 15 minute buffer because Google's Scheduling is not percise
    // See: https://developers.google.com/apps-script/guides/triggers/installable#time-driven_triggers

    for (var i in scheduledEmails){
        var email = scheduledEmails[i];
        if (Date.parse(email.date) < nowPlus15){
            sendEmailBySubject(email.subject);
        }
    }
}


function removeEmailFromSchedule(subject){
    var scheduledEmails = JSON.parse(userProperties.getProperty('scheduledData'));
    for (var i in scheduledEmails) {
        if (scheduledEmails[i].subject == subject) {
            scheduledEmails.splice(i,1);
        }
    }
    setScheduledEmails(scheduledEmails);
}

function getCurrentUser(){
    return Session.getActiveUser().getEmail();
}

function sendErrorEmail(originalSubject, body){
    var header = "Sorry the email with the following subject could not be sent: <b>" + originalSubject + "</b><br>";
    var footer = "<br><br><br>Sent by easy gmail scheduler"
    var html =  header + body + footer;
    GmailApp.sendEmail(Session.getActiveUser().getEmail(), "Gmail Scheduler Error", 'bodySpace', {htmlBody: html} );
}


function sendEmailBySubject(subject) {
    var drafts = GmailApp.getDraftMessages();

    for (var i in drafts) {
        if (drafts[i].getSubject() == subject) {
            var result = dispatchDraft(drafts[i].getId());
            removeEmailFromSchedule(subject)
            updateTriggers();

            if (result === "Delivered"){
                return true;
            } else {
                sendErrorEmail(subject, "Error: " + result)
            }
        } else if(i == drafts.length-1){ // If we reach the last draft and the message wasn't found
            sendErrorEmail(subject, "Error: No email was found with the subject: " + subject);
            removeEmailFromSchedule(subject);
        }
    }
    return false;
}


function setScheduledEmails(scheduleInfo) { // expects non-stringified JSON
    var data = JSON.stringify(scheduleInfo);
    userProperties.setProperty('scheduledData', data);
}

function deleteProperties(){
    Logger.log("Before " + userProperties.getProperty('scheduledData'));
    userProperties.deleteAllProperties();
    Logger.log("After " + userProperties.getProperty('scheduledData'));
    getScheduledEmails();
}

function getScheduledEmails() {
    //Returns parsed email schedule
    var data = userProperties.getProperty('scheduledData');
    Logger.log(data);
    if (data === null) { //TODO: This is only needed because we don't have an initalizer for the script
        // initialize the script here!
        data = [];
        setScheduledEmails(data);
    } else {
        data = JSON.parse(data);
    }
    return data;
}

function removeOldTriggers() {
    // Because there is no way to get a trigger's time
    // This deletes all triggers once there are no emails left in schedule
    var triggers = ScriptApp.getProjectTriggers();

    for (var i in triggers) {
        ScriptApp.deleteTrigger(triggers[i]);
    }

}

function sortEmailsByDate(emails){
    var sorted = emails.sort(function(a,b){
        return new Date(b.date) - new Date(a.date);
    });
    var sorted = sorted.reverse();
    return sorted;
}

function updateTriggers() {
    removeOldTriggers(); // Deletes all old triggers
    var emails = getScheduledEmails();
    var sortedEmails = sortEmailsByDate(emails);
    var triggerLimit = 5;
    /* 20 the limit set by google https://developers.google.com/apps-script/guides/services/quotas
    Using 5 instead to speed up the call
     */

    for (var i = 0; i < sortedEmails.length && i < triggerLimit; i++){
        var date = new Date(sortedEmails[i].date);
        ScriptApp.newTrigger("sendScheduledEmails")
            .timeBased()
            .at(date)
            .create();
    }
}


function addEmailToSchedule(subject, date) {
    // NOTE ABOUT FUNCTION: Data should be converted to date here but can't be due to issue with Google Scripts
    // https://code.google.com/p/google-apps-script-issues/issues/detail?id=4426
    var newEmail = {
        subject: subject,
        date: JSON.parse(date)
    };
    var scheduledEmails = getScheduledEmails();
    var emailAlreadyScheduled = isScheduled(scheduledEmails, newEmail.subject);

    if (emailAlreadyScheduled > -1) {
        scheduledEmails[emailAlreadyScheduled] = newEmail;
    } else {
        scheduledEmails.push(newEmail);
    }

    setScheduledEmails(scheduledEmails);
    updateTriggers();

    function isScheduled(scheduledEmails, subject) {
        // Returns index of desired email subject
        for (var i in scheduledEmails) {
            if (scheduledEmails[i].subject == subject) {
                return i;
            }
        }
        return -1;
    }
}

// All the follows replaces the original dispatchDraft and uses the GMail API in order
// to send drafts within the their thread. See
// https://stackoverflow.com/questions/27206595/how-to-send-a-draft-email-using-google-apps-script

/**
*/
function dispatchDraft(msgId){
  // Get draft message.
  var draftMsg = getDraftMsg(msgId,"json");
  if (!getDraftMsg(msgId)) throw new Error( "Unable to get draft with msgId '"+msgId+"'" );

  // see https://developers.google.com/gmail/api/v1/reference/users/drafts/send
  var url = 'https://www.googleapis.com/gmail/v1/users/me/drafts/send'
  var headers = {
    Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
  };
  var params = {
    method: "post",
    contentType: "application/json",
    headers: headers,
    muteHttpExceptions: true,
    payload: JSON.stringify(draftMsg)
  };
  var check = UrlFetchApp.getRequest(url, params)
  var response = UrlFetchApp.fetch(url, params);

  var result = response.getResponseCode();
  if (result == '200') {  // OK
    return "Delivered";
    //return JSON.parse(response.getContentText());
  }
  else {
    // This is only needed when muteHttpExceptions == true
    var err = JSON.parse(response.getContentText());
    throw new Error( 'Error (' + result + ") " + err.error.message );
  }
}

/**
 * https://stackoverflow.com/questions/27206595/how-to-send-a-draft-email-using-google-apps-script
 * Gets the draft message content that corresponds to a given Gmail Message ID.
 * Throws if unsuccessful.
 * See https://developers.google.com/gmail/api/v1/reference/users/drafts/get.
 *
 * @param {String}     messageId   Immutable Gmail Message ID to search for
 * @param {String}     optFormat   Optional format; "object" (default) or "json"
 *
 * @returns {Object or String}     If successful, returns a Users.drafts resource.
 */
function getDraftMsg( messageId, optFormat ) {
  var draftId = getDraftId( messageId );

  var url = 'https://www.googleapis.com/gmail/v1/users/me/drafts'+"/"+draftId;
  var headers = {
    Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
  };
  var params = {
    headers: headers,
    muteHttpExceptions: true
  };
  var check = UrlFetchApp.getRequest(url, params)
  var response = UrlFetchApp.fetch(url, params);

  var result = response.getResponseCode();
  if (result == '200') {  // OK
    if (optFormat && optFormat == "JSON") {
      return response.getContentText();
    }
    else {
      return JSON.parse(response.getContentText());
    }
  }
  else {
    // This is only needed when muteHttpExceptions == true
    var error = JSON.parse(response.getContentText());
    throw new Error( 'Error (' + result + ") " + error.message );
  }
}

/**
 * Gets the draft message ID that corresponds to a given Gmail Message ID.
 *
 * @param {String}     messageId   Immutable Gmail Message ID to search for
 *
 * @returns {String}               Immutable Gmail Draft ID, or null if not found
 */
function getDraftId( messageId ) {
  if (messageId) {
    var drafts = getDrafts();

    for (var i=0; i<drafts.length; i++) {
      if (drafts[i].message.id === messageId) {
        return drafts[i].id;
      }
    }
  }

  // Didn't find the requested message
  return null;
}

/**
 * Gets the current user's draft messages.
 * Throws if unsuccessful.
 * See https://developers.google.com/gmail/api/v1/reference/users/drafts/list.
 *
 * @returns {Object[]}             If successful, returns an array of 
 *                                 Users.drafts resources.
 */
function getDrafts() {
  var url = 'https://www.googleapis.com/gmail/v1/users/me/drafts';
  var headers = {
    Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
  };
  var params = {
    headers: headers,
    muteHttpExceptions: true
  };
  var check = UrlFetchApp.getRequest(url, params)
  var response = UrlFetchApp.fetch(url, params);

  var result = response.getResponseCode();
  if (result == '200') {  // OK
    return JSON.parse(response.getContentText()).drafts;
  }
  else {
    // This is only needed when muteHttpExceptions == true
    var error = JSON.parse(response.getContentText());
    //throw new Error( 'Error (' + result + ") " + error.message );
    return "Message Not Sent: " + error;
  }
}

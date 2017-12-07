var userProperties = PropertiesService.getUserProperties(); // This is to allow retrieval of stored info

function sendScheduledEmails() {
    // Sends all emails that are scheduled to be sent 15 minutes from now or sooner
    var scheduledEmails = JSON.parse(userProperties.getProperty('scheduledData'));
    var nowPlus15 = Date.parse(new Date()) + (15 * 60 * 1000); /// 15 minutes extra may not be nessicary
    // Email use a 15 minute buffer because Google's Scheduling is not percise
    // See: https://developers.google.com/apps-script/guides/triggers/installable#time-driven_triggers

    for (var i in scheduledEmails) {
        var email = scheduledEmails[i];
        if (Date.parse(email.date) < nowPlus15) {
            sendEmailBySubject(email.subject);
        }
    }
    function sendEmailBySubject(subject) {
        var drafts = GmailApp.getDraftMessages();
        try {
            for (var i in drafts) {
                if (drafts[i].getSubject() == subject) {
                    // This will throw if there are any errors
                    dispatchDraft(drafts[i].getId());
                    break; // Because of this break it will only send the first email with the selected subject
                } else if (i == drafts.length - 1) { // throws if we reach end of array
                    throw "Error: No email was found with the subject: " + subject;
                }
            }
        } catch (e) {
            sendErrorEmail(subject, e);
        }
        removeEmailFromSchedule(subject);
        updateTriggers();
    }
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
function setScheduledEmails(scheduleInfo) {
    // expects non-stringified JSON
    var data = JSON.stringify(scheduleInfo);
    userProperties.setProperty('scheduledData', data);
}
function removeEmailFromSchedule(subject) {
    var scheduledEmails = JSON.parse(userProperties.getProperty('scheduledData'));
    for (var i in scheduledEmails) {
        if (scheduledEmails[i].subject == subject) {
            scheduledEmails.splice(i, 1);
        }
    }
    setScheduledEmails(scheduledEmails);
}
function updateTriggers() {
    removeOldTriggers(); // Deletes all old triggers
    var emails = getScheduledEmails();
    var sortedEmails = sortEmailsByDate(emails);
    var triggerLimit = 5;
    /* 20 the limit set by google https://developers.google.com/apps-script/guides/services/quotas
    Using 5 instead to speed up the call */

    for (var i = 0; i < sortedEmails.length && i < triggerLimit; i++) {
        var date = new Date(sortedEmails[i].date);
        ScriptApp.newTrigger("sendScheduledEmails")
            .timeBased()
            .at(date)
            .create();
    }

    function removeOldTriggers() {
        // There is no way to get a trigger's time or identify it, so this deletes all triggers once
        var triggers = ScriptApp.getProjectTriggers();
        for (var i in triggers) {
            ScriptApp.deleteTrigger(triggers[i]);
        }
    }
    function sortEmailsByDate(emails) {
        var sorted = emails.sort(function (a, b) {
            return new Date(b.date) - new Date(a.date);
        });
        var sorted = sorted.reverse();
        return sorted;
    }
}
function sendErrorEmail(originalSubject, body) {
    var header = "Sorry the email with the following subject could not be sent: <b>" + originalSubject + "</b><br>";
    var footer = "<br><br><br>Sent by easy gmail scheduler"
    var html = header + body + footer;
    GmailApp.sendEmail(getCurrentUser(), "Gmail Scheduler Error", 'bodySpace', { htmlBody: html });
}
function deleteProperties() {
    // This is only used for testing purposes
    Logger.log("Before " + userProperties.getProperty('scheduledData'));
    userProperties.deleteAllProperties();
    Logger.log("After " + userProperties.getProperty('scheduledData'));
    getScheduledEmails();
}

// Below are functions used by web interface
function doGet() {
    // Serves html page for web interface
    return HtmlService.createHtmlOutputFromFile('index')
        // This setting may expose to to cross site scripting but will allow your app to work anywhere
        //.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        ///.setFaviconUrl(iconUrl) 
        .setTitle("Easy Gmail Scheduler")
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function getCurrentUser() {
    // Necessicary for frontend
    return Session.getActiveUser().getEmail();
}
function getDraftSubjects() {
    // Returns array of draft subjects and ids - used by web interface
    var subjects = [];
    var drafts = GmailApp.getDraftMessages();
    for (var i in drafts) {
        subjects.push(drafts[i].getSubject());
    }
    return subjects;
}
function addEmailToSchedule(subject, date) {
    // This is used by the web interface
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
// End functions used by web interface
function checkAuthorization() {
    // Checks that the script owner has authorized The app 
    var url = 'https://www.googleapis.com/gmail/v1/users/me/drafts';
    try {
        var response = GmailAPIRequest(url);
        Logger.log("Authorized!");
        return { authorized: true };
    } catch (e) {
        Logger.log("Not authorized! Error: " + e);
        return { authorized: false, error: e };
    }
}

// All the follows replaces the original dispatchDraft and uses the GMail API in order
// to send drafts within the their thread. See
// https://stackoverflow.com/questions/27206595/how-to-send-a-draft-email-using-google-apps-script

function dispatchDraft(msgId) {
    // Get draft message.
    var draftMsg = getDraftMsg(msgId);
    if (!getDraftMsg(msgId)) throw new Error("Unable to get draft with msgId '" + msgId + "'");

    // see https://developers.google.com/gmail/api/v1/reference/users/drafts/send
    var url = 'https://www.googleapis.com/gmail/v1/users/me/drafts/send';
    var params = {
        method: "post",
        contentType: "application/json",
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true,
        payload: draftMsg
    };
    var response = GmailAPIRequest(url, params);
    if (response.getResponseCode() == '200') {
        Logger.log("Delivered successfully!");
        return true;
    }
}

/**
 * https://stackoverflow.com/questions/27206595/how-to-send-a-draft-email-using-google-apps-script
 * Gets the draft message content that corresponds to a given Gmail Message ID.
 * Throws if unsuccessful.
 * See https://developers.google.com/gmail/api/v1/reference/users/drafts/get.
 *
 * @param {String}     messageId   Immutable Gmail Message ID to search for
 *
 * @returns {Object or String}     If successful, returns a Users.drafts resource.
 */
function getDraftMsg(messageId) {
    var draftId = getDraftId(messageId);
    var url = 'https://www.googleapis.com/gmail/v1/users/me/drafts' + "/" + draftId;
    var response = GmailAPIRequest(url);
    return response.getContentText();
}

/**
 * Gets the draft message ID that corresponds to a given Gmail Message ID.
 *
 * @param {String}     messageId   Immutable Gmail Message ID to search for
 *
 * @returns {String}               Immutable Gmail Draft ID, or null if not found
 */
function getDraftId(messageId) {
    var drafts = getDrafts();

    if (!Array.isArray(drafts)) {
        throw new Error("Unable to retrieve drafts: " + drafts);
    }

    for (var i = 0; i < drafts.length; i++) {
        if (drafts[i].message.id === messageId) {
            return drafts[i].id;
        }
    }

    // Didn't find the requested message
    return null;
}

/**
 * Gets the current user's draft messages.
 * See https://developers.google.com/gmail/api/v1/reference/users/drafts/list.
 *
 * @returns {Object[]}             If successful, returns an array of 
 *                                 Users.drafts resources.
 */
function getDrafts() {
    var url = 'https://www.googleapis.com/gmail/v1/users/me/drafts';
    var response = GmailAPIRequest(url);
    return JSON.parse(response.getContentText()).drafts;
}

function GmailAPIRequest(url, optionalParameters) {
    var headers = {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken() // This token should probably be stored somewhere so we don't have to keep getting it
    };
    var params = optionalParameters || {
        headers: headers,
        muteHttpExceptions: true
    };
    var response = UrlFetchApp.fetch(url, params);
    var result = response.getResponseCode();
    if (result == '200') {  // OK
        return response;
    }
    else if (result == '429') { // too many requests
        // should wait a second and then try again
        if (firstTry != false) {
            Utilities.sleep(1 * 1000)
            return GmailAPIRequest(url, optionalParameters, false);
        } else {
            throw "Error 429 - too many requested. Tried waiting 1 second and trying again but didn't work."
        }
    }
    else if (result == '403') { /// Shouldn't need this because we'll check authorization before hand
        throw "Error 403: Check that Gmail API has been enabled in Resources->Advanced Google Services->Gmail to On. Also enable Gmail API in Google Cloud Console.";
    }
    else { // This is only needed when muteHttpExceptions == true
        var error = JSON.parse(response.getContentText());
        throw "Message Not Sent: " + error;
    }
}

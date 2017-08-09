
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

// ** Credit to Gmail Scheduler - https://github.com/webdigi/GmailScheduler/** //
function dispatchDraft(id) {

    try {

        var message = GmailApp.getMessageById(id);

        if (message) {

            var body = message.getBody();
            var raw  = message.getRawContent();

            /* Credit - YetAnotherMailMerge */

            var regMessageId = new RegExp(id, "g");
            if (body.match(regMessageId) != null) {
                var inlineImages = {};
                var nbrOfImg = body.match(regMessageId).length;
                var imgVars = body.match(/<img[^>]+>/g);
                var imgToReplace = [];
                if(imgVars != null){
                    for (var i = 0; i < imgVars.length; i++) {
                        if (imgVars[i].search(regMessageId) != -1) {
                            var id = imgVars[i].match(/realattid=([^&]+)&/);
                            if (id != null) {
                                id = id[1];
                                var temp = raw.split(id)[1];
                                temp = temp.substr(temp.lastIndexOf('Content-Type'));
                                var imgTitle = temp.match(/name="([^"]+)"/);
                                var contentType = temp.match(/Content-Type: ([^;]+);/);
                                contentType = (contentType != null) ? contentType[1] : "image/jpeg";
                                var b64c1 = raw.lastIndexOf(id) + id.length + 3; // first character in image base64
                                var b64cn = raw.substr(b64c1).indexOf("--") - 3; // last character in image base64
                                var imgb64 = raw.substring(b64c1, b64c1 + b64cn + 1); // is this fragile or safe enough?
                                var imgblob = Utilities.newBlob(Utilities.base64Decode(imgb64), contentType, id); // decode and blob
                                if (imgTitle != null) imgToReplace.push([imgTitle[1], imgVars[i], id, imgblob]);
                            }
                        }
                    }
                }

                for (var i = 0; i < imgToReplace.length; i++) {
                    inlineImages[imgToReplace[i][2]] = imgToReplace[i][3];
                    var newImg = imgToReplace[i][1].replace(/src="[^\"]+\"/, "src=\"cid:" + imgToReplace[i][2] + "\"");
                    body = body.replace(imgToReplace[i][1], newImg);
                }
            }

            var options = {
                from        : message.getFrom(),
                cc          : message.getCc(),
                bcc         : message.getBcc(),
                htmlBody    : body,
                replyTo     : message.getReplyTo(),
                inlineImages: inlineImages,
                name        : message.getFrom().match(/[^<]*/)[0].trim(),
                attachments : message.getAttachments()
            }

            GmailApp.sendEmail(message.getTo(), message.getSubject(), body, options);
            message.moveToTrash();
            return "Delivered";
        } else {
            return "Message not found in Drafts";
        }
    } catch (e) {
        return e.toString();
    }
}

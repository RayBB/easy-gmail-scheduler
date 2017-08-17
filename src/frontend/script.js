'use strict';
const demoMode = (window == top);
let scheduledEmails;

window.onload = function () {
    if (demoMode) setupDemoMode();
    displaySchedule();
    displayDrafts();
    displayEmailAddress();

    // Set dates to today
    const curTime = moment().format('HH:mm');
    const curDate = moment().format('YYYY-MM-DD');
    document.querySelector('#date').value = curDate;
    document.querySelector('#time').value = curTime;

    // Adds date picker for browsers without date support
    loadPicker();

    setTimeout(function () {
        document.getElementById('myHTML').style.height = '100%';
        document.querySelector('#myHTML > body').style.height = '100%';
    }, 100)
    /* 
    NOTE 1: This is very hacky but what is happening is #myHTML must have height set to over 100% otherwise you can't scroll on mobile (if needed)
    This may be caused by this all being in an iframe and a lot of the height of the document rendering later.
    This basically sets the height back to 100% like normal
    */
};

function getEmailAddress(callback) {
    google.script.run.withSuccessHandler(callback).getCurrentUser();
}
function displayEmailAddress() {
    getEmailAddress(function (address) {
        document.querySelector('#userEmail').innerText = address;
    })
}
function getDrafts(callback) {
    google.script.run.withSuccessHandler(callback).getDraftSubjects();
}
function displayDrafts() {
    getDrafts(function (drafts) {
        let options = '';
        for (let draft of drafts) {
            options += '<option value="' + draft + '">' + draft + '<\/option>';
        }
        document.querySelector('#subjects').innerHTML = options;
        showLoading(false);
    })
}
function getSchedule(callback) {
    showLoading(true);
    google.script.run.withSuccessHandler(callback).getScheduledEmails();
}
function displaySchedule() {
    getSchedule(
        function (s) {
            const template = document.querySelector('#productrow');
            let newContents = document.createDocumentFragment();

            for (let message of s) {
                const formattedDate = moment(new Date(message.date)).format('MMM Do YYYY, hh:mm a');
                const clone = document.importNode(template.content, true);
                let td = clone.querySelectorAll("td");
                td[0].textContent = message.subject;
                td[1].textContent = formattedDate;
                td[2].addEventListener('click', function () { removeScheduledMessage(message.subject); });
                newContents.appendChild(clone);
            };

            //Deletes old table-items
            document.querySelectorAll('.table-item').forEach(function (e) { e.remove() });
            document.querySelector("tbody").appendChild(newContents);
            showLoading(false);
        }
    )
}
function addScheduledMessage() {
    const date = document.querySelector('#date').value;
    const time = parseTime(document.querySelector('#time').value);
    const subject = document.querySelector('#subjects').value;
    let sendDate = new Date(date + " " + time);

    if (dateIsInFuture(sendDate)) {
        showError(false);
        google.script.run.addEmailToSchedule(subject, JSON.stringify(sendDate));
        displaySchedule();
    } else {
        showError(true);
    }
}
function removeScheduledMessage(subject) {
    document.querySelectorAll('td').forEach(e => { if (e.innerText == subject) { e.parentNode.remove() } });
    google.script.run.removeEmailFromSchedule(subject);
}
function loadPicker() {
    // Only loads if browser doesn't support input type="date"
    let dateField = document.createElement("input");
    dateField.setAttribute("type", "date");

    if (dateField.type != "date") {
        let head = document.querySelector('#myHTML > head');
        let script = document.createElement("script");
        script.src = "https://cdnjs.cloudflare.com/ajax/libs/pikaday/1.6.1/pikaday.js";
        head.appendChild(script);

        let css = document.createElement("link");
        css.setAttribute("rel", "stylesheet");
        css.setAttribute("type", "text/css");
        css.href = "https://cdnjs.cloudflare.com/ajax/libs/pikaday/1.6.1/css/pikaday.min.css";
        head.appendChild(css);

        setTimeout(function () {
            const picker = new Pikaday({ field: document.getElementById('date') });
        }, 2000)
    }
}
function dateIsInFuture(formDate) {
    return moment(formDate).isAfter(Date.now());
}
function parseTime(string) {
    //Necessary to avoid adding another library for cross browser time support
    const isPM = string.toLowerCase().includes("p");
    const numbers = string.match(/\d+/g);
    let hours = parseInt(numbers[0]);
    const minutes = parseInt(numbers[1] | 0);
    if (isPM && hours != 12) hours += 12;
    return hours + ":" + minutes;
}
function showLoading(boolean) {
    document.querySelector('#loader').classList.toggle('loader', boolean);
}
function showError(boolean) {
    document.querySelector('#dateTimeRow').classList.toggle("has-error", boolean);
    document.querySelector('.help-block').classList.toggle("hidden", !boolean);
}


function setupDemoMode() {
    console.log("Running in demo mode. No data is sent");
    scheduledEmails = [
        { "subject": "Resume Inspiration", "date": "2018-08-01T20:11:00.000Z" },
        { "subject": "iPod Print", "date": "2018-07-31T21:14:00.000Z" },
        { "subject": "Happy Birthday!", "date": "2018-08-01T20:11:00.000Z" }
    ];

    class demoFuncs {
        getCurrentUser() {
            this._callBack("email@email.com");
        }
        getDraftSubjects() {
            const draftSubjects = ["Happy Birthday!", "Follow up on meeting", "Medium Article Examples", "Resume Inspiration",
                "iPod Print", "Lo And Behold: Reveries Of The Connected World"];
            this._callBack(draftSubjects);
        }
        getScheduledEmails() {
            this._callBack(scheduledEmails);
        }
        addEmailToSchedule(subject, sendDate) {
            sendDate = JSON.parse(sendDate);
            console.log("Submitting the following to Google: " + subject + " " + sendDate);
            const newEmail = { "subject": subject, "date": sendDate };
            const index = scheduledEmails.findIndex(e => { return e.subject == subject });
            index > -1 ? scheduledEmails[index] = newEmail : scheduledEmails.push(newEmail);
        }
        removeEmailFromSchedule(subject) {
            console.log("Removing: " + subject);
        }
        withSuccessHandler(callback) {
            let obj = new demoFuncs();
            obj._callBack = callback;
            return obj;
        }
    }
    window.google = { script: { run: new demoFuncs() } }
}

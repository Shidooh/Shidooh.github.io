'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            getPhishingItem(Office.context.mailbox.item);
            composemail();
        });
    });

    // defining global variables to pass them to the composeMail function
    var phishItemId;
    var phishSubject;

    // this function has to run before composing a new mail to retrieve the details of the current selected email. 
    function getPhishingItem(item) {
        phishItemId = item.itemId
        phishSubject = item.subject
    }

     //function to open a new 'compose message' form with predefined information
    function composemail() {
        office.context.mailbox.displayNewMessageForm({
            torecipients: ["mathis.merme@gmail.com","test"],
            subject: "phishing report: \"" + phishsubject + "\"",
            htmlbody:
                'test',
            attachments: [
                { type: "item", itemid: phishitemid, name: phishsubject }
            ],
        });
    }
})();

function hideShowSettings() {
    if (document.getElementById("settings").style.display === "none") {
        document.getElementById("settings").style.display = "block";
    } else {
        document.getElementById("settings").style.display = "none";
    };
};
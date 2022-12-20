'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            securityTeamMailAddress();
            loadCurrentMailAddress()
            getPhishingItem(Office.context.mailbox.item);
            composeMail();
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

    // function to open a new 'compose message' form with predefined information
    function composeMail() {
        Office.context.mailbox.displayNewMessageForm({
            toRecipients: ["test"],
            // ccRecipients: ["sam@contoso.com"], Send to more mailaddresses if necessary
            subject: "Phishing report: \"" + phishSubject + "\"",
            htmlBody:
                'testttttttttttttttttttttttttttttttttttttttttttttt',
            attachments: [
                { type: "item", itemId: phishItemId, name: phishSubject }
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
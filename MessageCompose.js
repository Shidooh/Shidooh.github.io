'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
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
   // function composeMail() {
       // Office.context.mailbox.displayNewMessageForm({
           // toRecipients: ["mathis.merme@gmail.com","test"],
            // ccRecipients: ["sam@contoso.com"], Send to more mailaddresses if necessary
          //  subject: "Phishing report: \"" + phishSubject + "\"",
           // htmlBody:
              //  'testttttttttttttttttttttttttttttttttttttttttttttt',
           // attachments: [
                //{ type: "item", itemId: phishItemId, name: phishSubject }
            //],
       // });
    //}

    function sendMail() {
        Office.context.mailbox.item.to.setAsync(["mathis.merme@gmail.com", "test"]);
        Office.context.mailbox.item.subject.setAsync("Phishing report: \"" + phishSubject + "\"");
        Office.context.mailbox.item.body.setAsync('testttttttttttttttttttttttttttttttttttttttttttttt', { coercionType: "html" });
        Office.context.mailbox.item.addFileAttachmentAsync(phishItemId, phishSubject, function (result) {
            if (result.status === "succeeded") {
                Office.context.mailbox.item.sendAsync();
            } else {
                console.error("Error adding attachment: " + result.error.message);
            }
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
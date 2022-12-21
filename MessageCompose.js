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
    var email = 'mathis.merme@mail.com';



function getPhishingItem(item) {
        phishItemId = item.itemId
        phishSubject = item.subject
    }

    function composeMail() {
        Office.context.mailbox.displayNewMessageForm.send({
            toRecipients: ["mathis.merme@gmail.com"],
            subject: "Phishing report: \"" + phishSubject + "\"",
            htmlBody:'testttttttttttt',
            attachments: [ { type: "item", itemId: phishItemId, name: phishSubject } ],
        });
    }

})();
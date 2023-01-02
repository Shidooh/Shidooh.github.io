'use strict';
(function () {
    Office.onReady(function () {
        $(document).ready(function () {
            getPhishingItem(Office.context.mailbox.item);
            composemail();
        });
    });

    var phishItemId;
    var phishSubject;
    var mail;

    function getPhishingItem(item) {
        phishItemId = item.itemId
        phishSubject = item.subject
    }

    function composemail() {
        Office.context.mailbox.displayNewMessageForm({
            toRecipients: ["mathis.merme@gmail.com"],
            ccRecipients: ["test"],
            subject: "phishing reportttttttttttt: \"" + phishSubject + "\"",
            htmlBody: 'testttttttttttttttttttttttt',
            attachments: [
                { type: "item", itemId: phishItemId, name: phishSubject }
            ],
        },
            function (asyncResult) {
                var newMessage = asyncResult.value;
                newMessage.sendAsync();
            }
        );
    }


   
})();
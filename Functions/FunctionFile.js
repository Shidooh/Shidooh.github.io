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

    function getPhishingItem(item) {
        phishItemId = item.itemId
        phishSubject = item.subject
    }

    function composemail() {
        Office.context.mailbox.displayNewMessageForm({
            torecipients: ["mathis.merme@gmail.com"],
            subject: "phishing report: \"" + phishSubject + "\"",
            htmlbody: 'test',
            attachments: [
                { type: "item", itemId: phishItemId, name: phishSubject }
            ],
        });
    }


   
})();
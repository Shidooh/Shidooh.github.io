'use strict';
(function () {
    Office.onReady(function () {
        $(document).ready(function () {
            getPhishingItem(Office.context.mailbox.item);
            composeMail();
        });
    });

    var phishItemId;
    var phishSubject;

    function getPhishingItem(item) {
        phishItemId = item.itemId
        phishSubject = item.subject
    }

    function composeMail() {
        Office.context.mailbox.displayNewMessageFormAsync(
            {
                toRecipients: ["mathis.merme@gmail.com"],
                subject: "Phishing report: \"" + phishSubject + "\"",
                htmlBody: 'test',
                attachments: [{ type: "item", itemId: phishItemId, name: phishSubject }],
                saveToSentItems: true
            },
        );
    }

})();
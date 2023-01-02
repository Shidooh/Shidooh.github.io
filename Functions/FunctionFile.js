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
                subject: "Phishing reportttttttttttttttttttt: \"" + phishSubject + "\"",
                htmlBody: 'test',
                attachments: [{ type: "item", itemId: phishItemId, name: phishSubject }],
                removeButtons: ['Send']
            },
            function (asyncResult) {
                var newMessage = asyncResult.value;
                newMessage.addHandlerAsync(Office.EventType.ItemSend, sendMessage);
                newMessage.sendAsync();
            }
        );
    }

    function sendMessage(_event) {
        // Suppression de l'élément courant (le message phish)
        Office.context.mailbox.item.deleteAsync();
    }

})();
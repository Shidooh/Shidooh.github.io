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
        Office.context.mailbox.displayNewMessageFormAsync(
            {
                toRecipients: ["mathis.merme@gmail.com"],
                subject: "Phishing report: \"" + phishSubject + "\"",
                htmlBody: 'test',
                attachments: [{ type: "item", itemId: phishItemId, name: phishSubject }],
            },
            function (asyncResult) {
                var newMessage = asyncResult.value;
                newMessage.addHandlerAsync(Office.EventType.ItemSend, sendMessage);
            }
        );
    }

    function sendMessage(event) {
        // Suppression de l'élément courant (le message phish)
        Office.context.mailbox.item.deleteAsync();
    }

})();
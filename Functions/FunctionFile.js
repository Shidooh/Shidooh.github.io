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
                htmlBody: 'nope',
                attachments: [{ type: "item", itemId: phishItemId, name: phishSubject }],
            },
            function (asyncResult) {
                var newMessage = asyncResult.value;
                // Ajout de l'événement "send" sur le nouveau message
                newMessage.addHandlerAsync(Office.EventType.ItemSend, sendMessage);
            }
        );
    }

    // Fonction de rappel qui sera appelée lorsque le nouveau message est envoyé
    function sendMessage(event) {
        // Suppression de l'élément courant (le message phish)
        Office.context.mailbox.item.deleteAsync();
    }

})();

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
                // Ajout de l'�v�nement "send" sur le nouveau message
                newMessage.addHandlerAsync(Office.EventType.ItemSend, sendMessage);
            }
        );
    }

    // Fonction de rappel qui sera appel�e lorsque le nouveau message est envoy�
    function sendMessage(event) {
        // Suppression de l'�l�ment courant (le message phish)
        Office.context.mailbox.item.deleteAsync();
    }

})();

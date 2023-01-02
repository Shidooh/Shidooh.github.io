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

    //function composemail() {
    //    Office.context.mailbox.displayNewMessageForm({
    //        toRecipients: ["mathis.merme@gmail.com"],
    //        subject: "phishing report: \"" + phishSubject + "\"",
    //        htmlBody: 'test',
    //        attachments: [
    //            { type: "item", itemId: phishItemId, name: phishSubject }
    //        ],
    //    },
    //        function (asyncResult) {
    //            var newMessage = asyncResult.value;
    //            newMessage.sendAsync();
    //        }
    //    );
    //}

    function composemail() {
        Office.context.mailbox.displayNewMessageFormAsync(
            {
                toRecipients: ["mathis.merme@gmail.com"],
                subject: "Phishing report: \"" + phishSubject + "\"",
                htmlBody: 'testtttttttttttttttttttt',
                attachments: [{ type: "item", itemId: phishItemId, name: phishSubject }],
            },
            function (asyncResult) {
                var newMessage = asyncResult.value;
                // Ajout de l'événement "send" sur le nouveau message
                newMessage.addHandlerAsync(Office.EventType.ItemSend, sendMessage);
                // Envoi du message
                newMessage.sendAsync();
            }
        );
    }



   
})();
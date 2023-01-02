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
                removeButtons: ['Send']
            },
            function (asyncResult) {
                var newMessage = asyncResult.value;
                newMessage.addHandlerAsync(Office.EventType.ItemSend, sendMessage);
                newMessage.sendAsync();
            }
        );
    }

    //function composeMail() {
    //    Office.context.mailbox.createNewMessageAsync(
    //        {
    //            toRecipients: ["mathis.merme@gmail.com"],
    //            subject: "Phishing report: \"" + phishSubject + "\"",
    //            htmlBody: 'testt',
    //            attachments: [{ type: "item", itemId: phishItemId, name: phishSubject }],
    //        },
    //        function (asyncResult) {
    //            var newMessage = asyncResult.value;
    //            newMessage.sendAsync();
    //            // Suppression de l'élément courant (le message phish) une fois que le nouveau message a été envoyé
    //            Office.context.mailbox.item.deleteAsync();
    //        }
    //    );
    //} 

})();
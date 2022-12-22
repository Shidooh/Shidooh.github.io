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
        Office.context.mailbox.item.subject.setAsync(subject);
        Office.context.mailbox.item.to.setAsync(

        {
                toRecipients: ["mathis.merme@gmail.com"],
                subject: "Phishing report: \"" + phishSubject + "\"",
                htmlBody: 'test',
                attachments: [{ type: "item", itemId: phishItemId, name: phishSubject }],
            },
            
        );
        Office.context.mailbox.item.sendAsync();
        Office.context.mailbox.item.deleteAsync();
    }

})();



//function composeMail() {
//    Office.context.mailbox.displayNewMessageFormAsync(
//        {
//            toRecipients: ["mathis.merme@gmail.com"],
//            subject: "Phishing report: \"" + phishSubject + "\"",
//            htmlBody: 'test',
//            attachments: [{ type: "item", itemId: phishItemId, name: phishSubject }],
//        },
//        function (asyncResult) {
//            var newMessage = asyncResult.value;
//            // Ajout de l'événement "send" sur le nouveau message
//            newMessage.addHandlerAsync(Office.EventType.ItemSend, sendMessage);
//        }
//    );
//}
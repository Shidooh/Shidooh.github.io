'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            getPhishingItem(Office.context.mailbox.item);
            composeMail();
            //run();
            //runAsync();
        });
    });

    // defining global variables to pass them to the composeMail function
    var phishItemId;
    var phishSubject;


function getPhishingItem(item) {
        phishItemId = item.itemId
        phishSubject = item.subject
    }

    function composeMail() {
        Office.context.mailbox.displayNewMessageForm ({
            toRecipients: ["mathis.merme@gmail.com"],
            subject: "Phishing report: \"" + phishSubject + "\"",
            htmlBody:'test. test',
            attachments: [ { type: "item", itemId: phishItemId, name: phishSubject } ],
        });
    }

})();

//function run() {
//    Office.context.mailbox.displayNewMessageForm({
//        toRecipients: ["mathis.merme@gmail.com"],
//        subject: "Outlook add-ins are cool!",
//        htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
//        attachments: [{ type: "item", itemId: phishItemId, name: phishSubject }],
//    });
//}

//    function runAsync() {
//        Office.context.mailbox.displayNewMessageFormAsync(
//            {
//                toRecipients: ["mathis.merme@gmail.com"],
//                subject: "Outlook add-ins are cool!",
//                htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
//                attachments: [{ type: "item", itemId: phishItemId, name: phishSubject }],
//            },
//            function (asyncResult) {
//                console.log(JSON.stringify(asyncResult));
//            }
//        );
//    }
//})();

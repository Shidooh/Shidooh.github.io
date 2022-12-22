'use strict';

//(function () {
//    office.initialize = function (reason) { };
//function myentrypoint(event) {
//        office.context.mailbox.displaynewmessageformasync({
//            torecipients: ["mathis.merme@gmail.com"],
//            subject: "test subject",
//            htmlbody: "internet headers: ",
//        },
//            { asynccontext: event },
//            function (result) {
//                console.log(result);
//                event.completed();
//            }
//        );
//    }

//    }) ();


(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            getPhishingItem(Office.context.mailbox.item);
            composeMail();
            //deleteItem();
        });
    });

    // defining global variables to pass them to the composeMail function
    var phishItemId;
    var phishSubject;
    //var item = office.context.mailbox.item;


    function getPhishingItem(item) {
        phishItemId = item.itemId
        phishSubject = item.subject
    }

    function composeMail() {
        Office.context.mailbox.displayNewMessageFormAsync(
            {
                toRecipients: Office.context.mailbox.item.to, // Copies the To line from current item
                ccRecipients: ["mathis.merme@gmail.com"],
                subject: "Outlook add-ins are cool!",
                htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
                attachments: [
                    {
                        type: "item",
                        itemId: phishItemId,
                        name: phishSubject,
                        isInline: true
                    }
                ]
            }
        )
    }

    function (asyncResult) {
        console.log(JSON.stringify(asyncResult));
    }

}) ();

//Office.context.mailbox.displayNewMessageFormAsync({
//                toRecipients: ["mathis.merme@gmail.com"],
//                subject: "Phishing report: \"" + phishSubject + "\"",
//                htmlBody: 'nope',
//                attachments: [{ type: "item", itemId: phishItemId, name: phishSubject }],
//            });
//        }


//    Office.initialize = function () {
//        sendSelectedEmail();
//    }

//    function sendSelectedEmail() {
//        // R�cup�ration de l'objet mail s�lectionn�
//        var item = Office.context.mailbox.item;
//        // Envoi du mail � l'adresse sp�cifique
//        item.forwardAsync("mathis.merme@mail.com", { asyncContext: item }, function (asyncResult) {
//            if (asyncResult.status == "failed") {
//                console.error("L'envoi du mail a �chou� : " + asyncResult.error.message);
//            } else {
//                // Suppression du mail de la bo�te de r�ception
//                item.deleteAsync(function (asyncResult) {
//                    if (asyncResult.status == "failed") {
//                        console.error("La suppression du mail a �chou� : " + asyncResult.error.message);
//                    } else {
//                        console.log("Le mail a �t� envoy� et supprim� avec succ�s.");
//                    }
//                });
//            }
//        });
//    }
//})();
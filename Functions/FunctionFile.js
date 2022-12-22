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
            composeMail(event);
        });
    });

    // defining global variables to pass them to the composeMail function
    var phishItemId;
    var phishSubject;


    function getPhishingItem(item) {
        phishItemId = item.itemId
        phishSubject = item.subject
    }

    function composeMail(event) {

        Office.context.mailbox.displayNewMessageFormAsync({
                toRecipients: ["mathis.merme@gmail.com"],
                subject: "Phishing report: \"" + phishSubject + "\"",
                htmlBody: 'testtttttttttttttt',
                attachments: [{ type: "item", itemId: phishItemId, name: phishSubject }],
        });
        event.completed({ allowEvent: true });

        }
}) ();




//    Office.initialize = function () {
//        sendSelectedEmail();
//    }

//    function sendSelectedEmail() {
//        // Récupération de l'objet mail sélectionné
//        var item = Office.context.mailbox.item;
//        // Envoi du mail à l'adresse spécifique
//        item.forwardAsync("mathis.merme@mail.com", { asyncContext: item }, function (asyncResult) {
//            if (asyncResult.status == "failed") {
//                console.error("L'envoi du mail a échoué : " + asyncResult.error.message);
//            } else {
//                // Suppression du mail de la boîte de réception
//                item.deleteAsync(function (asyncResult) {
//                    if (asyncResult.status == "failed") {
//                        console.error("La suppression du mail a échoué : " + asyncResult.error.message);
//                    } else {
//                        console.log("Le mail a été envoyé et supprimé avec succès.");
//                    }
//                });
//            }
//        });
//    }
//})();
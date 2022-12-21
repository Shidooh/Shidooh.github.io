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
        Office.context.mailbox.displayNewMessageForm({
            toRecipients: ["mathis.merme@gmail.com"],
            subject: "Phishing report: \"" + phishSubject + "\"",
            htmlBody: 'TEST',
            attachments: [{ type: "item", itemId: phishItemId, name: phishSubject }],
        });
    }

})();


//Office.initialize = function () {
//}

//function sendSelectedEmail() {
//    // Récupération de l'objet mail sélectionné
//    var item = Office.context.mailbox.item;
//    // Envoi du mail à l'adresse spécifique
//    item.forwardAsync("mathis.merme@mail.com", { asyncContext: item }, function (asyncResult) {
//        if (asyncResult.status == "failed") {
//            console.error("L'envoi du mail a échoué : " + asyncResult.error.message);
//        } else {
//            // Suppression du mail de la boîte de réception
//            item.deleteAsync(function (asyncResult) {
//                if (asyncResult.status == "failed") {
//                    console.error("La suppression du mail a échoué : " + asyncResult.error.message);
//                } else {
//                    console.log("Le mail a été envoyé et supprimé avec succès.");
//                }
//            });
//        }
//    });
//}
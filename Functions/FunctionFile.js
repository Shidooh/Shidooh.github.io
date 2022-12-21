'use strict';

(function () {

        Office.onReady(function () {
            // Office is ready
            $(document).ready(function () {
                getPhishingItem(Office.context.mailbox.item);
                composeMail();
                deleteItem();
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
            Office.context.mailbox.displayNewMessageForm({
                toRecipients: ["mathis.merme@gmail.com"],
                subject: "Phishing report: \"" + phishSubject + "\"",
                htmlBody: 'test',
                attachments: [{ type: "item", itemId: phishItemId, name: phishSubject }],
            });
        }

        function deleteItem() {
            //Office.context.mailbox.item.deleteAsync
            phishItemId.deleteAsync(function (result) {
                if (result.status === "succeeded") {
                    console.log("Item deleted successfully");
                } else {
                    console.log("Error deleting item: " + result.error.message);
                }
            });
        }
})();




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
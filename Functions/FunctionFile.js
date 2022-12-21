Office.initialize = function () {
}

function sendSelectedEmail() {
    // R�cup�ration de l'objet mail s�lectionn�
    var item = Office.context.mailbox.item;
    // Envoi du mail � l'adresse sp�cifique
    item.forwardAsync("mathis.merme@mail.com", { asyncContext: item }, function (asyncResult) {
        if (asyncResult.status == "failed") {
            console.error("L'envoi du mail a �chou� : " + asyncResult.error.message);
        } else {
            // Suppression du mail de la bo�te de r�ception
            item.deleteAsync(function (asyncResult) {
                if (asyncResult.status == "failed") {
                    console.error("La suppression du mail a �chou� : " + asyncResult.error.message);
                } else {
                    console.log("Le mail a �t� envoy� et supprim� avec succ�s.");
                }
            });
        }
    });
}
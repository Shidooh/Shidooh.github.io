Office.initialize = function () {
    // Fonction exécutée lorsque l'add-in est chargé
};

function sendAndDeleteMessage() {
    // Récupère le message sélectionné
    var item = Office.context.mailbox.item;

    // Envoie le message à l'adresse e-mail spécifiée
    item.forwardAsync("mathis.merme@gmail.com", {
        asyncContext: { message: item }
    }, function (asyncResult) {
        if (asyncResult.status == "failed") {
            console.error("Error sending message: " + asyncResult.error.message);
            return;
        }

        // Supprime le message une fois l'envoi terminé
        asyncResult.value.deleteAsync(function (asyncResult) {
            if (asyncResult.status == "failed") {
                console.error("Error deleting message: " + asyncResult.error.message);
            }
        });
    });
}

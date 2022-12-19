Office.initialize = function () {
    // Fonction ex�cut�e lorsque l'add-in est charg�
};

function sendAndDeleteMessage() {
    // R�cup�re le message s�lectionn�
    var item = Office.context.mailbox.item;

    // Envoie le message � l'adresse e-mail sp�cifi�e
    item.forwardAsync("mathis.merme@gmail.com", {
        asyncContext: { message: item }
    }, function (asyncResult) {
        if (asyncResult.status == "failed") {
            console.error("Error sending message: " + asyncResult.error.message);
            return;
        }

        // Supprime le message une fois l'envoi termin�
        asyncResult.value.deleteAsync(function (asyncResult) {
            if (asyncResult.status == "failed") {
                console.error("Error deleting message: " + asyncResult.error.message);
            }
        });
    });
}

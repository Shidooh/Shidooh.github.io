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

    function composemail() {
    office.context.mailbox.displaynewmessageformasync(
        {
            torecipients: ["mathis.merme@gmail.com"],
            subject: "phishing report: \"" + phishsubject + "\"",
            htmlbody: 'test',
            attachments: [{ type: "item", itemid: phishitemid, name: phishsubject }],
        },
        function (asyncresult) {
            var newmessage = asyncresult.value;
            // ajout de l'événement "send" sur le nouveau message
            newmessage.addhandlerasync(office.eventtype.itemsend, sendmessage);
        }
    );
}

})();
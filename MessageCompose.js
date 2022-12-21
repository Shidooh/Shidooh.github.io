'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            //test();
            // The document is ready
            //securityTeamMailAddress();
            //loadCurrentMailAddress()
            getPhishingItem(Office.context.mailbox.item);
            composeMail();
        });
    });

    // defining global variables to pass them to the composeMail function
    var phishItemId;
    var phishSubject;
    var receipentMailAddress;
    var email = 'mathis.merme@mail.com';

    //function test() {
    //    Office.context.mailbox.item.create(
    //    "message",
    //    {
    //        toRecipients: ["mathis.merme@grmail.com"],
    //        subject: "Outlook add-ins are cool!",
    //        body: {
    //            contentType: "html",
    //            content: 'Hello <b>World</b>!'
    //        }
    //    },
    //    function (result) {
    //        // The create callback function
    //        if (result.status === "created") {
    //            var item = result.value;
    //            item.send();
    //        }
    //    }
    //);
    //}

    // get the reciepent or ask to enter value
    // check if there is an email address set to send the mail to
    //function securityTeamMailAddress() {
    //     check if email is already set
    //    if (Office.context.roamingSettings.get("email")) {
    //        receipentMailAddress = Office.context.roamingSettings.get("email")
    //    }
    //     show popup to enter the email address to report phishing to.
    //    else {
    //         TODO: Create popup to enter email at first run

    //         Office.context.roamingSettings.set("email", "j.vdvelden99@gmail.com")
    //        receipentMailAddress = Office.context.roamingSettings.get("email")
    //        saveRoamingSettings()
    //    }
    //}


function getPhishingItem(item) {
        phishItemId = item.itemId
        phishSubject = item.subject
    }

    
    // this function has to run before composing a new mail to retrieve the details of the current selected email. 
    


    // function to open a new 'compose message' form with predefined information
    function composeMail() {
        Office.context.mailbox.displayNewMessageForm({
            toRecipients: ["mathis.merme@gmail.com", "test"],
            // ccRecipients: ["sam@contoso.com"], Send to more mailaddresses if necessary
            subject: "Phishing report: \"" + phishSubject + "\"",
            htmlBody:
                'nopetttttttttttttt',
            attachments: [
                { type: "item", itemId: phishItemId, name: phishSubject }
            ],
        });
    }
})();

function hideShowSettings() {
    if (document.getElementById("settings").style.display === "none") {
        document.getElementById("settings").style.display = "block";
    } else {
        document.getElementById("settings").style.display = "none";
    };
};

//function loadCurrentMailAddress() {
//    // Write message property values to the task pane
//    document.getElementById("currentMailAddress").innerHTML = Office.context.roamingSettings.get("email");
//}
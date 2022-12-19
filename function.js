function SendToSpecificEmail() {
    // Make sure the script is running in Outlook on the web
    if (Office.context.mailbox.diagnostics.hostName !== 'OutlookIOS' && Office.context.mailbox.diagnostics.platform !== 'OutlookWebApp') {
        UI.notify('This script can only be run in Outlook on the web.', 'error');
        return;
    }

    // Get the current mailbox and selected message
    var mailbox = Office.context.mailbox;
    var item = mailbox.item;

    // Get the value of the saveToSentItems option
    var saveToSentItems = document.getElementById('saveToSentItems').value;

    // Forward the message to the specified email address and delete it
    try {
        item.forwardAsync("mathis.merme@gmail.com", { saveToSentItems: saveToSentItems },
            function (result) {
                if (result.status === 'failed') {
                    UI.notify(`An error occurred while forwarding the message: ${result.error.message}`, 'error');
                } else {
                    item.deleteAsync(function (deleteResult) {
                        if (deleteResult.status === 'failed') {
                            UI.notify(`An error occurred while deleting the message: ${deleteResult.error.message}`, 'error');
                        } else {
                            UI.notify('The message was forwarded and deleted successfully.', 'success');
                        }
                    });
                }
            });
    } catch (error) {
        UI.notify(`An unexpected error occurred: ${error.message}`, 'error');
    }
}
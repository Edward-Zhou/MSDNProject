/// <reference path="../App.js" />

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            displayItemDetails();
        });
    };

    // Displays the "Subject" and "From" fields, based on the current mail item
    function displayItemDetails() {
        var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
        $('#subject').text(item.subject);

        var from;
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            from = Office.cast.item.toMessageRead(item).from;
        } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            from = Office.cast.item.toAppointmentRead(item).organizer;
        }

        if (from) {
            $('#from').text(from.displayName);
            $('#from').click(function () {
                app.showNotification(from.displayName, from.emailAddress);
            });
        }
    }
})();

//Office.cast.item
function CastTest()
{
    var message = Office.cast.item.toItemRead(Office.context.mailbox.item);
    var messagebak = Office.context.mailbox.item;
    Office.context.mailbox.item.body.getAsync( function (result)
    {
        //var item = result.;
    });

    //messagebak.
}

function GetBodyTest() {
    //Office.cast.item.toMessageCompose(Office.context.mailbox.item).body.getAsync(function (result) {
    //    app.showNotification('The current body is', result.value)
    //});

    Office.context.mailbox.item.body.getAsync(function (result) {
        app.showNotification('The current body is', result.value)
    });
}

function getUserIdentityTokenAsync() {
    function getUserIdentityTokenCallback(asyncResult) {
        var token = asyncResult.value;
    }
}

function getBody() {
    Office.context.mailbox.item.body.getAsync("text", function (result) {
        app.showNotification(result.value);

    })

}

function openURL() {
    window.open("http://www.w3schools.com");
}

function runScript() {
    var shell = new ActiveXObject("WScript.Shell");
    shell.run("https://www.microsoft.com");
}

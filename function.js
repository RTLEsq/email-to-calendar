// Email to Calendar - Outlook Add-in
// Creates a pre-filled appointment form from the current email.
// The user adds date/time and saves.

Office.initialize = function (reason) {
    // Office is ready
};

/**
 * Main function: called when the user clicks the "Create Task" button.
 * Reads the current email's details and opens a new appointment form
 * pre-filled with subject, sender info, and a link to the email.
 */
function createCalendarEvent(event) {
    var item = Office.context.mailbox.item;

    // Gather email details
    var subject = item.subject || "(No Subject)";
    var sender = item.from ? item.from.displayName : "Unknown";
    var senderEmail = item.from ? item.from.emailAddress : "";
    var receivedDate = item.dateTimeCreated
        ? item.dateTimeCreated.toLocaleString()
        : "Unknown";

    // Build the Outlook Web App link to the original email
    var itemId = item.itemId;

    // Build the appointment body as HTML so links are clickable
    // Use outlook: protocol to open email in desktop client
    var desktopLink = "";
    var webLink = "";
    if (itemId) {
        desktopLink = "outlook:" + itemId;
        webLink = "https://outlook.office365.com/mail/id/" + encodeURIComponent(itemId);
    }

    // Get the body preview for context
    item.body.getAsync(Office.CoercionType.Text, function (result) {
        var emailBody = "";
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            emailBody = result.value;
            if (emailBody.length > 500) {
                emailBody = emailBody.substring(0, 500) + "...";
            }
            // Escape HTML characters in the email preview
            emailBody = emailBody.replace(/&/g, "&amp;")
                                 .replace(/</g, "&lt;")
                                 .replace(/>/g, "&gt;")
                                 .replace(/\n/g, "<br>");
        }

        var body = "<div style='font-family: Segoe UI, Arial, sans-serif;'>";
        body += "<h3 style='color: #0078d4;'>TASK FROM EMAIL</h3>";

        if (desktopLink) {
            body += "<p><strong>Open Original Email:</strong><br>";
            body += "<a href='" + desktopLink + "'>Open in Outlook Desktop</a>";
            body += " &nbsp;|&nbsp; ";
            body += "<a href='" + webLink + "'>Open in Outlook Web</a>";
            body += "</p>";
        }

        body += "<hr>";
        body += "<p>";
        body += "<strong>From:</strong> " + sender;
        if (senderEmail && senderEmail !== sender) {
            body += " &lt;" + senderEmail + "&gt;";
        }
        body += "<br>";
        body += "<strong>Received:</strong> " + receivedDate + "<br>";
        body += "<strong>Original Subject:</strong> " + subject;
        body += "</p>";
        body += "<hr>";

        if (emailBody) {
            body += "<p><strong>Email Preview:</strong><br>";
            body += emailBody;
            body += "</p>";
        }

        body += "</div>";

        // Open the new appointment form.
        // requiredAttendees and optionalAttendees are empty arrays
        // so it creates an APPOINTMENT, not a meeting.
        Office.context.mailbox.displayNewAppointmentForm({
            requiredAttendees: [],
            optionalAttendees: [],
            subject: subject,
            body: body,
            location: ""
        });

        // Signal that the add-in has finished processing
        event.completed();
    });
}

// Register the function with Office
Office.actions = Office.actions || {};
Office.actions.associate("createCalendarEvent", createCalendarEvent);

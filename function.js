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
    // This uses the itemId to create a deep link
    var itemId = item.itemId;
    var mailboxUrl = Office.context.mailbox.restUrl || "";

    // Construct OWA deep link
    // Format: https://outlook.office365.com/mail/deeplink/compose/<itemId>
    // For reading: https://outlook.office365.com/mail/id/<itemId>
    var emailLink = "";
    if (itemId) {
        // Use the EWS item ID to build an OWA link
        emailLink = "https://outlook.office365.com/mail/id/" + encodeURIComponent(itemId);
    }

    // Build the appointment body
    var body = "";
    body += "========================================\n";
    body += "TASK FROM EMAIL\n";
    body += "========================================\n\n";

    if (emailLink) {
        body += "OPEN ORIGINAL EMAIL (with attachments):\n";
        body += emailLink + "\n\n";
    }

    body += "----------------------------------------\n";
    body += "From: " + sender;
    if (senderEmail && senderEmail !== sender) {
        body += " <" + senderEmail + ">";
    }
    body += "\n";
    body += "Received: " + receivedDate + "\n";
    body += "Original Subject: " + subject + "\n";
    body += "----------------------------------------\n\n";

    // Get the body preview for context
    item.body.getAsync(Office.CoercionType.Text, function (result) {
        var emailBody = "";
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            // Include first 500 characters of the email body
            emailBody = result.value;
            if (emailBody.length > 500) {
                emailBody = emailBody.substring(0, 500) + "...";
            }
        }

        if (emailBody) {
            body += "EMAIL PREVIEW:\n";
            body += emailBody + "\n";
        }

        // Open the new appointment form
        // This opens Outlook's native appointment window with fields pre-filled
        // The user fills in date, time, and duration, then saves
        Office.context.mailbox.displayNewAppointmentForm(
            "[TASK] " + subject,  // Subject
            body,                  // Body (plain text)
            [],                    // No attendees (appointment, not meeting)
            "",                    // No location
            {                      // Start/end left to defaults
                // When start/end are not specified or invalid,
                // Outlook uses the next default time slot
            }
        );

        // Signal that the add-in has finished processing
        event.completed();
    });
}

// Register the function with Office
Office.actions = Office.actions || {};
Office.actions.associate("createCalendarEvent", createCalendarEvent);

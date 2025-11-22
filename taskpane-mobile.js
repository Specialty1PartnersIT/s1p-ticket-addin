/* -------------------------------------------------------
   S1P — Outlook Mobile IT Support Add-in (Command-based)
   This file runs ONLY on Outlook Mobile (iOS/Android)
---------------------------------------------------------*/

/*
 HOW IT WORKS:

 - Outlook Mobile does NOT support taskpanes.
 - When the user presses the “IT Mobile Ticket” button,
   Outlook calls `onMobileTicket(event)` below.

 - This function opens a NEW MESSAGE window with:
      • Pre-filled To: IT support email
      • Pre-filled Subject
      • Pre-filled Body template
      • Critical toggle text
      • Quick prompts for location, issue, callback

 - User can edit & send normally.
*/

/* -------------------------------------------------------
   MAIN ENTRY POINT
---------------------------------------------------------*/

function onMobileTicket(event) {
  try {
    console.log("Mobile IT ticket command triggered.");

    // Open a new email form with pre-populated values
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: ["support@specialty1partners.com"],

      subject: "Mobile IT Support Ticket",

      body:
        "Please provide the following details:\n\n" +
        "🔹 **Location Code:** \n" +
        "🔹 **Issue Category:** (Network / Hardware / Server / Power / Other)\n" +
        "🔹 **Callback Number:** \n" +
        "🔹 **Is this critical?** (Yes/No)\n\n" +
        "———————————————\n" +
        "Additional Details:\n"
    });

    event.completed(); // Required for mobile commands

  } catch (e) {
    console.log("Mobile Ticket Error:", e);
    event.completed();
  }
}

/* -------------------------------------------------------
   EXPORT COMMANDS (required for iOS/Android)
---------------------------------------------------------*/

if (typeof module !== "undefined") {
  module.exports = {
    onMobileTicket
  };
}

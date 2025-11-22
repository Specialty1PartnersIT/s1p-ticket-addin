/* 
  ============================================================
  S1P IT MOBILE SUPPORT ADD-IN (Outlook Mobile Safe)
  - No DOM
  - No Taskpane
  - No HTML
  - Uses ExecuteFunction only
  - Works on iOS, Android, Desktop, and OWA
  ============================================================
*/

/*
  MAIN ENTRY POINT
  This is called by ExecuteFunction from the manifest.
*/
function createMobileTicket(event) {
    try {
        const mailbox = Office.context.mailbox;

        // Prepopulate subject
        const subject = buildSubject();

        // Build the default email body
        const body = buildBody();

        // Recipient (same shared inbox you use in desktop version)
        const toAddress = "itsupport@specialty1partners.com";

        // Open a new ticket email
        mailbox.displayNewMessageForm({
            toRecipients: [toAddress],
            subject: subject,
            htmlBody: body
        });

    } catch (err) {
        console.log("Error in createMobileTicket():", err);
    }

    // ALWAYS call event.completed() to prevent "stuck" ribbon button
    if (event && event.completed) {
        event.completed();
    }
}

/*
  ============================================================
  SUBJECT BUILDER
  Mobile version uses:
  "Mobile IT Ticket – [Critical/Normal] – (Location)"
  ============================================================
*/
function buildSubject() {
    const isCritical = askCritical();
    const location = askLocation();

    let subject = "Mobile IT Ticket";

    if (isCritical) subject += " – CRITICAL";
    if (location) subject += ` – (${location})`;

    return subject;
}

/*
  ============================================================
  BODY BUILDER
  ============================================================
*/
function buildBody() {
    const isCritical = askCritical();
    const location = askLocation();

    const user = safeUserName();
    const deviceInfo = getDeviceInfo();

    return `
        <p><strong>New Mobile IT Support Ticket Submitted</strong></p>
        
        <p><strong>Submitted By:</strong> ${user}</p>
        <p><strong>Location Code:</strong> ${location || "Not provided"}</p>
        <p><strong>Critical Issue:</strong> ${isCritical ? "YES – Immediate attention required" : "No"}</p>

        <p><strong>Device Info:</strong><br>
        ${deviceInfo}
        </p>

        <hr>

        <p><strong>Description of Issue:</strong></p>
        <p>Please describe the issue here...</p>

        <p><em>Attach screenshots or videos if available.</em></p>
    `;
}

/*
  ============================================================
  PROMPTS
  (Mobile-safe input methods)
  ============================================================
*/

function askCritical() {
    try {
        return confirm(
            "Is this a CRITICAL issue (system down, no network, outage, preventing patient care)?"
        );
    } catch {
        return false;
    }
}

function askLocation() {
    try {
        const loc = prompt("Enter your Location Code (e.g., END073, PO0806, OS0609):");
        return loc ? loc.trim() : "";
    } catch {
        return "";
    }
}

/*
  ============================================================
  SAFE USER NAME
  Handles mobile where profile may return null
  ============================================================
*/
function safeUserName() {
    try {
        const user = Office.context.mailbox.userProfile;
        if (user && user.displayName) return user.displayName;
    } catch {}
    return "Unknown User";
}

/*
  ============================================================
  DEVICE IDENTIFICATION
  (Helps IT troubleshoot mobile issues)
  ============================================================
*/
function getDeviceInfo() {
    const ua = navigator.userAgent || "";

    let device = "Unknown Device";

    if (/iPhone/i.test(ua)) device = "iPhone";
    else if (/iPad/i.test(ua)) device = "iPad";
    else if (/Android/i.test(ua)) device = "Android Device";

    return `
        Device: ${device}<br>
        User Agent: ${ua}
    `;
}

/*
  ============================================================
  EXPORT FOR EXECUTEFUNCTION
  (Mandatory for Outlook Mobile)
  ============================================================
*/
if (typeof module !== "undefined") {
    module.exports = { createMobileTicket };
}

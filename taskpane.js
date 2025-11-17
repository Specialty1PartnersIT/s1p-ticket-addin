Office.onReady(() => {
  document.getElementById("submitBtn").onclick = sendTicket;

  const categorySelect = document.getElementById("category");
  const subcategorySelect = document.getElementById("subcategory");

  categorySelect.addEventListener("change", () => {
    updateSubcategories(categorySelect.value, subcategorySelect);
  });
});

// --------------------------
// FULL SUBCATEGORY MAP
// --------------------------
const SUBCATEGORY_MAP = {

  "PMS / Practice Management System": [
    "EagleSoft",
    "EndoVision/Clinivision/OMSVision",
    "WinOMS",
    "TDO",
    "WinDent",
    "Dentrix",
    "PBS Endo",
    "DSN",
    "Denticon",
    "Other (PMS)"
  ],

  "Imaging Software": [
    "CRD Dicom",
    "Romexis",
    "EZ3D-i",
    "Dexis",
    "Sidexis",
    "Carestream",
    "XDR",
    "Vatech EzDent",
    "VistaScan",
    "Other (Imaging)"
  ],

  "Email/Outlook": [
    "Outlook Desktop (Windows)",
    "Outlook Desktop (Mac)",
    "Outlook Mobile (iOS/Android)",
    "Shared Mailbox",
    "Distribution List",
    "Send/Receive Problems",
    "Calendar Issues",
    "Add-ins Not Loading",
    "Authentication / MFA",
    "Other (Email)"
  ],

  "Teams": [
    "Chat Issues",
    "Teams Meetings",
    "Devices (Headsets/Webcams)",
    "Teams Phone",
    "Channels / Permissions",
    "File Sharing Issues",
    "Notifications",
    "Login Errors",
    "Other (Teams)"
  ],

  "SharePoint": [
    "Permissions",
    "Missing Files",
    "Sync Issues (OneDrive)",
    "Site Access Issues",
    "Document Versioning",
    "SharePoint Lists / Forms",
    "Other (SharePoint)"
  ],

  "Hardware": [
    "Workstation / PC Issue",
    "Laptop",
    "Printer",
    "Scanner",
    "Camera",
    "Label Printer",
    "Card Reader",
    "Signature Pad",
    "Monitor",
    "Keyboard / Mouse",
    "Other (Hardware)"
  ],

  "Network/Internet": [
    "Internet Down",
    "WiFi Issues",
    "Firewall / Security",
    "VPN",
    "DNS Issues",
    "Slowness / Latency",
    "Cannot Reach Cloud Services",
    "Local Server Unreachable",
    "Other (Network)"
  ],

  "Other": [
    "New Employee Setup",
    "User Access Request",
    "Password Reset",
    "Software Install",
    "General Question",
    "Other (Misc)"
  ]
};

// --------------------------
// Populate Subcategories
// --------------------------
function updateSubcategories(category, subSelect) {
  subSelect.innerHTML = "";
  subSelect.disabled = true;

  if (!category || !SUBCATEGORY_MAP[category]) {
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = "-- Select a category first --";
    subSelect.appendChild(opt);
    return;
  }

  subSelect.disabled = false;

  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "-- Select a subcategory --";
  subSelect.appendChild(placeholder);

  SUBCATEGORY_MAP[category].forEach(sc => {
    const opt = document.createElement("option");
    opt.value = sc;
    opt.textContent = sc;
    subSelect.appendChild(opt);
  });
}

// --------------------------
// Send Ticket Email
// --------------------------
function sendTicket() {
  const category = document.getElementById("category").value;
  const subcategory = document.getElementById("subcategory").value;
  const workstation = document.getElementById("workstation").value.trim();
  const callback = document.getElementById("callback").value.trim();
  const location = document.getElementById("location").value.trim();
  const description = document.getElementById("description").value.trim();

  // REQUIRED FIELDS VALIDATION
  if (!workstation) {
    alert("Please enter the workstation name.");
    return;
  }

  if (!callback) {
    alert("Please enter a callback phone number.");
    return;
  }

  const user = Office.context.mailbox.userProfile.displayName;
  const email = Office.context.mailbox.userProfile.emailAddress;

  let subject = `New IT Support Ticket - ${category}`;
  if (subcategory) subject += ` - ${subcategory}`;

  const htmlBody = `
    <p><b>User:</b> ${user} (${email})</p>
    <p><b>Category:</b> ${category}</p>
    <p><b>Subcategory:</b> ${subcategory}</p>
    <p><b>Workstation:</b> ${workstation}</p>
    <p><b>Callback:</b> ${callback}</p>
    <p><b>Location Code:</b> ${location}</p>
    <p><b>Description:</b><br>${description}</p>
  `;

  Office.context.mailbox.displayNewMessageForm({
    toRecipients: ["support@specialty1partners.com"],
    subject: subject,
    htmlBody: htmlBody
  });
}

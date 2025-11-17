// ---------------------------
// CATEGORY → SUBCATEGORY MAP
// ---------------------------
const categoryMap = {
  "PMS / Practice Management System": [
    "Login Issues",
    "Data Issues",
    "Software Error",
    "Updates Needed"
  ],
  "Imaging Software": [
    "Sensor Issues",
    "Acquisition Errors",
    "Software Crash",
    "Integration Problems"
  ],
  "Email/Outlook": [
    "Send/Receive Issues",
    "Login Problems",
    "Add-in Issues",
    "Calendar Problems"
  ],
  "Teams": [
    "Chat Issues",
    "Meeting Problems",
    "Audio/Video Issues",
    "Login Issues"
  ],
  "SharePoint": [
    "Access Issue",
    "Sync Problems",
    "Permissions",
    "Document Issues"
  ],
  "Hardware": [
    "Printer Issue",
    "Scanner Issue",
    "Computer Won't Boot",
    "Peripheral Issue"
  ],
  "Network/Internet": [
    "Connectivity",
    "Slowness",
    "VPN Issues",
    "Drops/Outages"
  ],
  "Other": [
    "General Issue",
    "Unknown Issue"
  ],

  // ---------------------------
  // NEW — RingCentral
  // ---------------------------
  "RingCentral": [
    "Phone hardware",
    "Call quality",
    "Fax send/receive",
    "Text send/receive",
    "Early office closure / Schedule changes"
  ],

  // ---------------------------
  // NEW — Security
  // ---------------------------
  "Security": [
    "Breach",
    "Incident",
    "Non-Critical"
  ],

  // ---------------------------
  // NEW — Acumen
  // ---------------------------
  "Acumen": [
    "Access/Permissions",
    "Data Integrity",
    "New Reports",
    "Report Edits"
  ]
};

// ---------------------------
// INITIALIZE SUBCATEGORY LOGIC
// ---------------------------
window.onload = function () {
  const categoryDropdown = document.getElementById("category");
  const subcategoryDropdown = document.getElementById("subcategory");

  categoryDropdown.addEventListener("change", () => {
    const selected = categoryDropdown.value;

    // Clear previous options
    subcategoryDropdown.innerHTML = "";
    subcategoryDropdown.disabled = true;

    if (!selected || !categoryMap[selected]) {
      subcategoryDropdown.innerHTML =
        `<option value="">-- Select a category first --</option>`;
      return;
    }

    // Populate new subcategories
    categoryMap[selected].forEach(sub => {
      const opt = document.createElement("option");
      opt.value = sub;
      opt.textContent = sub;
      subcategoryDropdown.appendChild(opt);
    });

    subcategoryDropdown.disabled = false;
  });

  // ---------------------------
  // SUBMIT BUTTON HANDLER
  // ---------------------------
  document.getElementById("submitBtn").addEventListener("click", submitTicket);
};

// ---------------------------
// TICKET SUBMISSION HANDLER
// ---------------------------
function submitTicket() {
  const category = document.getElementById("category").value;
  const subcategory = document.getElementById("subcategory").value;
  const workstation = document.getElementById("workstation").value;
  const callback = document.getElementById("callback").value;
  const location = document.getElementById("location").value;
  const description = document.getElementById("description").value;

  // Basic validation
  if (!category) return alert("Please select a category.");
  if (!subcategory) return alert("Please select a subcategory.");
  if (!workstation) return alert("Workstation name is required.");
  if (!callback) return alert("Callback number is required.");

  const body = `
Category: ${category}
Subcategory: ${subcategory}
Workstation: ${workstation}
Callback: ${callback}
Location: ${location}
Description:
${description}
  `;

  // Insert into email or send to an API — whatever your workflow is
  Office.context.mailbox.item.body.setAsync(body, { coercionType: "text" }, () => {
    alert("Ticket information added to the email body.");
  });
}

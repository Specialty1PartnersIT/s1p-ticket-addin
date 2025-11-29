/* ------------------------------------------------
   CONSTANTS
--------------------------------------------------*/
const IT_NEW_HIRE_CATEGORY = "User account creation / new hire setup";

/* ------------------------------------------------
   DEPARTMENT → EMAIL ROUTING
--------------------------------------------------*/
const DEPARTMENT_EMAIL_MAP = {
  "IT": "support@specialty1partners.com",
  "CBO - Full Service": "cbo@specialty1partners.com",
  "Payor Relations": "payorrelations@specialty1partners.com",
  "RCM - Non Full Service": "rcm@specialty1partners.com"
};

/* ------------------------------------------------
   CATEGORY MAP (Per Department)
--------------------------------------------------*/
const CATEGORY_MAP = {
  IT: [
    "PMS / Practice Management System",
    "Imaging Software",
    "Email/Outlook",
    "Teams",
    "SharePoint",
    "Hardware",
    "Network/Internet",
    "Security",
    "RingCentral",
    "Acumen",
    IT_NEW_HIRE_CATEGORY,
    "Other"
  ],

  "CBO - Full Service": [
    "Adjustment Correction Needed",
    "Arestin Xray Request",
    "Authorization Review Request",
    "Billing / Claims Inquiry",
    "Charge Correction Needed",
    "Clinical Info Request",
    "COB Info Needed",
    "EFT Enrollment Request",
    "EOB Request",
    "General / Other RCM Questions",
    "Insurance Paid To Patient",
    "Insurance Refund Request",
    "Office Check Search Request",
    "Office EOB Search",
    "Patient AR Request",
    "Patient Insurance Info Needed",
    "Payment Correction Needed",
    "Payment Posting Request",
    "Secondary Claim Needed",
    "Other"
  ],

  "Payor Relations": [
    "Add Payor",
    "Claim Review",
    "Fee Review Request",
    "Fee Schedule discrepancy",
    "Leasing Inquiries",
    "Medical Request",
    "Meeting Request",
    "Miscellaneous",
    "New Affiliation",
    "New Provider",
    "Opt-Out",
    "Ownership Change",
    "Participation Status",
    "Payor News (Add to Intranet)",
    "Payor - Add Network",
    "Payor - Document Request",
    "Potential Participation",
    "Request Fee Schedule",
    "Term Networkedit",
    "Term Provider",
    "TIN Change",
    "Update Practice Data",
    "W-9 Request",
    "Other"
  ],

  "RCM - Non Full Service": [
    "Collection Placement",
    "RCM General questions",
    "Refund",
    "Other"
  ]
};

/* ------------------------------------------------
   IT SOFTWARE (2nd level)
--------------------------------------------------*/
const IT_SOFTWARE_MAP = {
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
    "CareStream",
    "Sidexis",
    "Romexis",
    "EZ3D-i",
    "Dexis",
    "Other (Imaging)"
  ]
};

/* ------------------------------------------------
   IT ISSUE TYPES (3rd level)
--------------------------------------------------*/
const IT_ISSUE_TYPE_MAP = {
  "PMS / Practice Management System": [
    "Access / Permissions",
    "Installation / Upgrade",
    "Performance / Slowness",
    "Data / Charting Issue",
    "Integration with other systems",
    "Other (PMS Issue)"
  ],
  "Imaging Software": [
    "Image acquisition / capture",
    "Viewer / workstation issue",
    "Integration with PMS",
    "Export / sharing issue",
    "Other (Imaging Issue)"
  ]
};

/* ------------------------------------------------
   IT OTHER SUBCATEGORIES (single level)
--------------------------------------------------*/
const IT_OTHER_SUBCATEGORY_MAP = {
  "Email/Outlook": [
    "Cannot send",
    "Cannot receive",
    "Sync issues",
    "Calendar issue",
    "Shared mailbox issue",
    "Add-in issue",
    "Other (Email)"
  ],

  "Teams": [
    "Chat/Channels",
    "Meetings/Calls",
    "Screen sharing",
    "Teams/Channels access",
    "Notifications",
    "Other (Teams)"
  ],

  "SharePoint": [
    "Permissions/Access",
    "File sync/OneDrive",
    "Broken link",
    "Page not loading",
    "Version history/Restore",
    "Other (SharePoint)"
  ],

  "Hardware": [
    "Desktop/Laptop",
    "Docking station",
    "Monitor",
    "Printer/Scanner",
    "Phone/Headset",
    "Other (Hardware)"
  ],

  "Network/Internet": [
    "No connectivity",
    "Slow network",
    "VPN",
    "Wi-Fi",
    "Other (Network)"
  ],

  "Security": [
    "Breach",
    "Incident",
    "Non-Critical"
  ],

  "RingCentral": [
    "Phone hardware",
    "Call quality",
    "Fax send/receive",
    "Text send/receive",
    "Early office closure / schedule change"
  ],

  "Acumen": [
    "Access/Permissions",
    "Data Integrity",
    "New Reports",
    "Report Edits"
  ],

  "Other": ["General / Not specified"]
};

/* ------------------------------------------------
   HOST DETECTION
--------------------------------------------------*/
function getHostEnvironment() {
  try {
    if (window.Office && Office.context && Office.context.mailbox) {
      return "outlook";
    }
  } catch (e) {
    // ignore
  }
  try {
    if (window.microsoftTeams && microsoftTeams.app) {
      return "teams";
    }
  } catch (e) {
    // ignore
  }
  return "web";
}

/* ------------------------------------------------
   HELPER FUNCTIONS
--------------------------------------------------*/
function clearSelect(select, placeholder) {
  if (!select) return;
  select.innerHTML = "";
  const opt = document.createElement("option");
  opt.value = "";
  opt.textContent = placeholder;
  select.appendChild(opt);
}

function fillSelect(select, values) {
  if (!select || !values) return;
  values.forEach(v => {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = v;
    select.appendChild(opt);
  });
}

function show(el, yes) {
  if (!el) return;
  el.classList[yes ? "remove" : "add"]("hidden");
}

/* ------------------------------------------------
   FORM WIRING
--------------------------------------------------*/
let formInitialized = false;

function initFormWiring() {
  if (formInitialized) return;
  formInitialized = true;

  const dept = document.getElementById("department");
  const cat = document.getElementById("category");
  const sub = document.getElementById("subcategory");
  const subLabel = document.getElementById("subcategory-label");
  const subsub = document.getElementById("subsubcategory");
  const subsubLabel = document.getElementById("subsubcategory-label");
  const ws = document.getElementById("ws-section");

  const newHireSection = document.getElementById("newhire-section");
  const equipmentSection = document.getElementById("equipment-section");
  const workerTypeSelect = document.getElementById("worker-type");
  const callbackSection = document.getElementById("callback-section");
  const locationSection = document.getElementById("location-section");

  clearSelect(cat, "-- Select a category --");

  // Department change
  if (dept) {
    dept.addEventListener("change", () => {
      clearSelect(cat, "-- Select a category --");
      clearSelect(sub, "-- Select a subcategory --");
      clearSelect(subsub, "-- Select an issue type --");

      show(sub, false);
      show(subLabel, false);
      show(subsub, false);
      show(subsubLabel, false);

      // reset special sections
      show(newHireSection, false);
      show(equipmentSection, false);
      show(callbackSection, true);
      show(locationSection, true);

      if (CATEGORY_MAP[dept.value]) {
        fillSelect(cat, CATEGORY_MAP[dept.value]);
      }

      show(ws, dept.value === "IT");
    });
  }

  // Category change
  if (cat) {
    cat.addEventListener("change", () => {
      clearSelect(sub, "-- Select a subcategory --");
      clearSelect(subsub, "-- Select an issue type --");

      show(sub, false);
      show(subLabel, false);
      show(subsub, false);
      show(subsubLabel, false);

      const deptVal = dept.value;
      const catVal = cat.value;
      const isIT = deptVal === "IT";
      const isNewHire = isIT && catVal === IT_NEW_HIRE_CATEGORY;

      // Reset new hire / visibility
      show(newHireSection, false);
      show(equipmentSection, false);

      if (isNewHire) {
        // New hire: hide workstation, callback, location; show new hire fields
        show(ws, false);
        show(callbackSection, false);
        show(locationSection, false);
        show(newHireSection, true);
        // No subcategories or issue type for new hire
        return;
      }

      // Non–new-hire: restore callback/location and ws as usual
      show(callbackSection, true);
      show(locationSection, true);
      show(ws, isIT);

      // Existing IT logic
      if (isIT) {
        if (IT_SOFTWARE_MAP[catVal]) {
          fillSelect(sub, IT_SOFTWARE_MAP[catVal]);
          show(sub, true);
          show(subLabel, true);

          fillSelect(subsub, IT_ISSUE_TYPE_MAP[catVal]);
          show(subsub, true);
          show(subsubLabel, true);
        } else if (IT_OTHER_SUBCATEGORY_MAP[catVal]) {
          fillSelect(sub, IT_OTHER_SUBCATEGORY_MAP[catVal]);
          show(sub, true);
          show(subLabel, true);
        }
      }
    });
  }

  // Worker Type change (for new hire) – show equipment only for Remote worker
  if (workerTypeSelect) {
    workerTypeSelect.addEventListener("change", () => {
      show(equipmentSection, workerTypeSelect.value === "Remote worker");
    });
  }

  // Submit handler
  const submitBtn = document.getElementById("submitBtn");
  if (submitBtn) {
    submitBtn.addEventListener("click", submitTicket);
  }
}

/* ------------------------------------------------
   CONTACT NAME AUTOFILL (Outlook Only)
--------------------------------------------------*/
function initContactAutofill() {
  try {
    const info = Office.context && Office.context.mailbox && Office.context.mailbox.userProfile;
    const nameField = document.getElementById("contactName");
    if (info && info.displayName && nameField) {
      nameField.value = info.displayName;
      localStorage.setItem("lastContactName", info.displayName);
    }

    const saved = localStorage.getItem("lastContactName");
    if (saved && nameField && !nameField.value) {
      nameField.value = saved;
    }
  } catch (e) {
    console.log("Name autofill failed:", e);
  }
}

/* ------------------------------------------------
   SUBMIT LOGIC
--------------------------------------------------*/
function submitTicket() {
  const dept = (document.getElementById("department") || {}).value || "";
  const cat = (document.getElementById("category") || {}).value || "";
  const sub = (document.getElementById("subcategory") || {}).value || "";
  const subsub = (document.getElementById("subsubcategory") || {}).value || "";
  const loc = (document.getElementById("location") || {}).value || "";
  const contact = (document.getElementById("contactName") || {}).value || "";
  const callback = (document.getElementById("callback") || {}).value || "";
  const workstationEl = document.getElementById("workstation");
  const workstation = workstationEl ? workstationEl.value : "";
  const desc = (document.getElementById("description") || {}).value || "";

  const newHireName = (document.getElementById("newhire-name") || {}).value?.trim() || "";
  const newHireTitle = (document.getElementById("newhire-title") || {}).value?.trim() || "";
  const workerType = (document.getElementById("worker-type") || {}).value || "";
  const equipment = (document.getElementById("equipment") || {}).value || "";

  const isNewHire = dept === "IT" && cat === IT_NEW_HIRE_CATEGORY;
  const detail = [sub, subsub].filter(v => v).join(" – ");

  const email = DEPARTMENT_EMAIL_MAP[dept];

  if (!dept || !cat || !email) {
    alert("Please select a department and category before submitting.");
    return;
  }

  if (!contact) {
    alert("Contact Name is required.");
    return;
  }

  if (!isNewHire && !callback) {
    alert("Callback Number is required.");
    return;
  }

  if (dept === "IT" && !isNewHire && !workstation) {
    alert("Please provide the workstation name for IT tickets.");
    return;
  }

  let subject;
  let body;

  if (isNewHire) {
    // New Hire subject
    subject = `New Hire – ${newHireName || "Name TBA"} – ${workerType || "Worker type"}`;
    if (loc) subject += ` – (${loc})`;

    const equipmentHtml = (workerType === "Remote worker" && equipment)
      ? `<b>Equipment Needed:</b><br>${equipment.replace(/\n/g, "<br>")}<br><br>`
      : "";

    body = `
<b>New Hire Name:</b> ${newHireName}<br>
<b>New Hire Title:</b> ${newHireTitle}<br>
<b>Worker Type:</b> ${workerType}<br>
${equipmentHtml}
<b>Requested by (Contact):</b> ${contact}<br>
<b>Department:</b> ${dept}<br>
<b>Category:</b> ${cat}<br>
${loc ? `<b>Location Code:</b> ${loc}<br>` : ""}
<br>
<b>Description / Additional Details:</b><br>
${(desc || "").replace(/\n/g, "<br>")}
    `;
  } else {
    // Standard ticket
    subject = `Ticket – ${dept}: ${cat}`;
    if (detail) subject += ` – ${detail}`;
    if (loc) subject += ` – (${loc})`;

    body = `
<b>Contact Name:</b> ${contact}<br>
<b>Callback Number:</b> ${callback}<br>
${dept === "IT" ? `<b>Workstation:</b> ${workstation}<br>` : ""}
<b>Location:</b> ${loc}<br>
<b>Department:</b> ${dept}<br>
<b>Category:</b> ${cat}<br>
${sub ? `<b>Subcategory:</b> ${sub}<br>` : ""}
${subsub ? `<b>Issue Type:</b> ${subsub}<br>` : ""}
<br>
<b>Description:</b><br>
${(desc || "").replace(/\n/g, "<br>")}
    `;
  }

  const host = getHostEnvironment();
  console.log("Submitting ticket; detected host:", host);

  if (host === "outlook") {
    try {
      Office.context.mailbox.displayNewMessageForm({
        toRecipients: [email],
        subject: subject,
        htmlBody: body
      });

      try {
        if (Office.context.ui && Office.context.ui.messageParent) {
          Office.context.ui.messageParent("close");
        }
      } catch (e) {
        // ignore
      }
    } catch (err) {
      console.error("displayNewMessageForm failed:", err);
      alert("Unable to open a new message in Outlook. Please try again.");
    }
    return;
  }

  if (host === "teams") {
    // Future: send to webhook/API; for now just log
    alert("✔ Ticket captured in Teams context.\n(Next step: wire up backend / webhook.)");
    console.log("Teams submission payload:", {
      dept, cat, sub, subsub, loc, contact, callback, workstation,
      newHireName, newHireTitle, workerType, equipment, desc
    });
    return;
  }

  // Standalone / web test
  alert("Form submitted in web/standalone mode.\n(Outlook-specific email compose is disabled here.)");
  console.log("Web submission payload:", {
    dept, cat, sub, subsub, loc, contact, callback, workstation,
    newHireName, newHireTitle, workerType, equipment, desc
  });
}

/* ------------------------------------------------
   INITIALIZATION
--------------------------------------------------*/

// DOM wiring (works in all hosts)
document.addEventListener("DOMContentLoaded", () => {
  initFormWiring();
});

// Outlook-specific initialization
if (window.Office && Office.onReady) {
  Office.onReady().then(info => {
    console.log("Office.onReady:", info);
    if (info.host === Office.HostType.Outlook) {
      initContactAutofill();
    }
  }).catch(err => {
    console.log("Office.onReady failed:", err);
  });
}

// Optional Teams init for future use; safe even when not in Teams
try {
  if (window.microsoftTeams && microsoftTeams.app) {
    microsoftTeams.app.initialize().then(() => {
      console.log("S1P Support: Teams SDK initialized");
      microsoftTeams.app.getContext().then(ctx => {
        console.log("Teams context:", ctx);
      });
    }).catch(err => {
      console.log("Teams initialization failed:", err);
    });
  }
} catch (e) {
  // ignore
}

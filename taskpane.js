/* ------------------------------------------------
   HOST DETECTION (Teams vs Outlook vs Standalone)
--------------------------------------------------*/
function getHostEnvironment() {
  if (window.microsoftTeams && microsoftTeams.app) return "teams";
  if (window.Office && Office.context && Office.context.mailbox) return "outlook";
  return "web";
}

let S1P_HOST = getHostEnvironment();

/* ------------------------------------------------
   CONSTANTS
--------------------------------------------------*/
const NEW_HIRE_CATEGORY = "User account creations / New hire setup";

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
    NEW_HIRE_CATEGORY,
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
   HELPERS
--------------------------------------------------*/
function clearSelect(select, placeholder) {
  select.innerHTML = "";
  const opt = document.createElement("option");
  opt.value = "";
  opt.textContent = placeholder;
  select.appendChild(opt);
}

function fillSelect(select, values) {
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
   FORM INITIALIZATION
--------------------------------------------------*/
document.addEventListener("DOMContentLoaded", () => {
  const dept = document.getElementById("department");
  const cat = document.getElementById("category");
  const sub = document.getElementById("subcategory");
  const subLabel = document.getElementById("subcategory-label");
  const subsub = document.getElementById("subsubcategory");
  const subsubLabel = document.getElementById("subsubcategory-label");
  const ws = document.getElementById("ws-section");

  const callbackSection = document.getElementById("callback-section");
  const locationSection = document.getElementById("location-section");
  const newHireSection = document.getElementById("newhire-section");
  const workerTypeSelect = document.getElementById("worker-type");
  const equipmentSection = document.getElementById("equipment-section");

  clearSelect(cat, "-- Select a category --");

  // Department change
  dept.addEventListener("change", () => {
    clearSelect(cat, "-- Select a category --");
    clearSelect(sub, "-- Select a subcategory --");
    clearSelect(subsub, "-- Select an issue type --");

    show(sub, false);
    show(subLabel, false);
    show(subsub, false);
    show(subsubLabel, false);

    show(newHireSection, false);
    show(callbackSection, true);
    show(locationSection, true);
    show(ws, dept.value === "IT");

    if (CATEGORY_MAP[dept.value]) {
      fillSelect(cat, CATEGORY_MAP[dept.value]);
    }
  });

  // Category change
  cat.addEventListener("change", () => {
    clearSelect(sub, "-- Select a subcategory --");
    clearSelect(subsub, "-- Select an issue type --");

    show(sub, false);
    show(subLabel, false);
    show(subsub, false);
    show(subsubLabel, false);

    const deptVal = dept.value;
    const catVal = cat.value;

    // Default: hide new hire section, show normal fields
    show(newHireSection, false);
    show(callbackSection, true);
    show(locationSection, true);
    show(ws, deptVal === "IT");

    // Special case: IT New Hire
    if (deptVal === "IT" && catVal === NEW_HIRE_CATEGORY) {
      // Hide workstation, callback, location
      show(ws, false);
      show(callbackSection, false);
      show(locationSection, false);

      // Hide subcategory controls
      show(sub, false);
      show(subLabel, false);
      show(subsub, false);
      show(subsubLabel, false);

      // Show new hire section
      show(newHireSection, true);
      return;
    }

    // Normal IT mapping
    if (deptVal === "IT") {
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

  // Worker type change → show equipment only for Remote Worker
  if (workerTypeSelect && equipmentSection) {
    workerTypeSelect.addEventListener("change", () => {
      show(equipmentSection, workerTypeSelect.value === "Remote Worker");
    });
  }

  document.getElementById("submitBtn").addEventListener("click", submitTicket);
});

/* ------------------------------------------------
   CONTACT NAME AUTOFILL (Outlook Only)
--------------------------------------------------*/
Office.onReady(info => {
  if (info.host !== Office.HostType.Outlook) return;

  try {
    const profile = Office.context.mailbox.userProfile;
    const nameField = document.getElementById("contactName");

    if (profile && profile.displayName) {
      nameField.value = profile.displayName;
      localStorage.setItem("lastContactName", profile.displayName);
    }

    const saved = localStorage.getItem("lastContactName");
    if (saved) nameField.value = saved;
  } catch (e) {
    console.log("Name autofill failed:", e);
  }
});

/* ------------------------------------------------
   SUBMIT LOGIC
--------------------------------------------------*/
function submitTicket() {
  const dept = document.getElementById("department").value;
  const cat = document.getElementById("category").value;
  const sub = document.getElementById("subcategory").value;
  const subsub = document.getElementById("subsubcategory").value;
  const loc = document.getElementById("location").value;
  const contact = document.getElementById("contactName").value;
  const callback = document.getElementById("callback").value;
  const workstationEl = document.getElementById("workstation");
  const workstation = workstationEl ? workstationEl.value : "";
  const desc = document.getElementById("description").value;

  // New hire fields
  const newHireName = (document.getElementById("newhire-name")?.value || "").trim();
  const newHireTitle = (document.getElementById("newhire-title")?.value || "").trim();
  const workerType = document.getElementById("worker-type")?.value || "";
  const equipmentNeeded = (document.getElementById("equipment-needed")?.value || "").trim();

  const isNewHire = (dept === "IT" && cat === NEW_HIRE_CATEGORY);
  const email = DEPARTMENT_EMAIL_MAP[dept];

  let subject = "";
  let body = "";

  if (isNewHire) {
    // New Hire–style subject
    subject = `New Hire – ${newHireName || "Name TBD"} – ${workerType || "Worker Type"}`;
    if (loc) subject += ` – (${loc})`;

    body = `
<b>Request Type:</b> New Hire Setup<br>
<b>New Hire Name:</b> ${newHireName}<br>
<b>New Hire Title:</b> ${newHireTitle}<br>
<b>Worker Type:</b> ${workerType}<br>
${equipmentNeeded ? `<b>Equipment Needed:</b><br>${equipmentNeeded.replace(/\n/g,"<br>")}<br>` : ""}
<br>
<b>Requester:</b> ${contact}<br>
<b>Callback Number:</b> ${callback}<br>
<b>Location:</b> ${loc}<br>
<b>Department:</b> ${dept}<br>
<b>Category:</b> ${cat}<br>
<br>
<b>Additional Details:</b><br>
${desc.replace(/\n/g, "<br>")}
    `;
  } else {
    const detail = [sub, subsub].filter(v => v).join(" – ");

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
${desc.replace(/\n/g, "<br>")}
    `;
  }

  /* ---------- Outlook Flow ---------- */
  if (S1P_HOST === "outlook") {
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: [email],
      subject: subject,
      htmlBody: body
    });

    try { Office.context.ui.messageParent("close"); } catch {}
    return;
  }

  /* ---------- Teams Flow ---------- */
  if (S1P_HOST === "teams") {
    alert("✔ Ticket form submitted in Teams.\n(Backend integration TBD).");

    console.log("Teams submission payload:", {
      dept, cat, sub, subsub, loc, contact, callback,
      workstation, desc,
      newHireName, newHireTitle, workerType, equipmentNeeded, isNewHire
    });

    return;
  }

  /* ---------- Web/Standalone ---------- */
  alert("Form submitted — not in Outlook or Teams.");
}

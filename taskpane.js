/* ------------------------------------------------
   DATA MAPS
--------------------------------------------------*/

// Categories per department
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

// For IT: software lists for PMS & Imaging (subcategory level)
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
    "Romexis",
    "EZ3D-i",
    "Dexis",
    "Other (Imaging)"
  ]
};

// For IT: issue types (sub-subcategory) for PMS & Imaging
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

// For IT: normal one-level subcategories for *other* IT categories
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
  Teams: [
    "Chat/Channels",
    "Meetings/Calls",
    "Screen sharing",
    "Teams/Channels access",
    "Notifications",
    "Other (Teams)"
  ],
  SharePoint: [
    "Permissions/Access",
    "File sync/OneDrive",
    "Broken link",
    "Page not loading",
    "Version history/Restore",
    "Other (SharePoint)"
  ],
  Hardware: [
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
  Security: [
    "Breach",
    "Incident",
    "Non-Critical"
  ],
  RingCentral: [
    "Phone hardware",
    "Call quality",
    "Fax send/receive",
    "Text send/receive",
    "Early office closure / schedule change"
  ],
  Acumen: [
    "Access/Permissions",
    "Data Integrity",
    "New Reports",
    "Report Edits"
  ],
  Other: [
    "General / Not specified"
  ]
};

/* ------------------------------------------------
   HELPERS
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
  if (!select || !Array.isArray(values)) return;
  values.forEach(v => {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = v;
    select.appendChild(opt);
  });
}

function showElement(el, show) {
  if (!el) return;
  if (show) {
    el.classList.remove("hidden");
  } else {
    el.classList.add("hidden");
  }
}

/* ------------------------------------------------
   FORM INITIALISATION
--------------------------------------------------*/

function initForm() {
  const deptSelect = document.getElementById("department");
  const categorySelect = document.getElementById("category");
  const subSelect = document.getElementById("subcategory");
  const subLabel = document.getElementById("subcategory-label");
  const subSubSelect = document.getElementById("subsubcategory");
  const subSubLabel = document.getElementById("subsubcategory-label");
  const wsSection = document.getElementById("ws-section");
  const submitBtn = document.getElementById("submitBtn");

  if (!deptSelect || !categorySelect || !submitBtn) {
    console.log("Ticket form elements not found – check taskpane.html.");
    return;
  }

  // Department change
  deptSelect.addEventListener("change", () => {
    const dept = deptSelect.value;

    clearSelect(categorySelect, "-- Select a category --");
    clearSelect(subSelect, "-- Select a subcategory --");
    clearSelect(subSubSelect, "-- Select an issue type --");

    showElement(subSelect, false);
    showElement(subLabel, false);
    showElement(subSubSelect, false);
    showElement(subSubLabel, false);

    if (CATEGORY_MAP[dept]) {
      fillSelect(categorySelect, CATEGORY_MAP[dept]);
    }

    showElement(wsSection, dept === "IT");
  });

  // Category change
  categorySelect.addEventListener("change", () => {
    const dept = deptSelect.value;
    const cat = categorySelect.value;

    clearSelect(subSelect, "-- Select a subcategory --");
    clearSelect(subSubSelect, "-- Select an issue type --");

    showElement(subSelect, false);
    showElement(subLabel, false);
    showElement(subSubSelect, false);
    showElement(subSubLabel, false);

    if (dept === "IT") {
      // PMS / Imaging special logic (software + issue type)
      if (cat === "PMS / Practice Management System" || cat === "Imaging Software") {
        fillSelect(subSelect, IT_SOFTWARE_MAP[cat]);
        fillSelect(subSubSelect, IT_ISSUE_TYPE_MAP[cat]);

        showElement(subSelect, true);
        showElement(subLabel, true);
        showElement(subSubSelect, true);
        showElement(subSubLabel, true);
        return;
      }

      // Other IT categories – one-level list
      const list = IT_OTHER_SUBCATEGORY_MAP[cat];
      if (list) {
        fillSelect(subSelect, list);
        showElement(subSelect, true);
        showElement(subLabel, true);
      }
    }
  });

  submitBtn.addEventListener("click", onSubmitClick);
}

/* ------------------------------------------------
   OFFICE HELPERS
--------------------------------------------------*/

function setSubjectSafe(subjectText) {
  try {
    if (
      typeof Office !== "undefined" &&
      Office.context?.mailbox?.item?.subject &&
      typeof Office.context.mailbox.item.subject.setAsync === "function"
    ) {
      Office.context.mailbox.item.subject.setAsync(subjectText);
    }
  } catch (e) {
    console.log("Subject set failed:", e);
  }
}

function closeTaskpaneSafe() {
  try {
    Office.context?.ui?.closeContainer();
  } catch (e) {}
}

/* ------------------------------------------------
   UPDATED SUBMIT HANDLER  (FULL BODY + SUBJECT)
--------------------------------------------------*/

function onSubmitClick() {
  const dept = document.getElementById("department")?.value || "";
  const category = document.getElementById("category")?.value || "";
  const sub = document.getElementById("subcategory")?.value || "";
  const subsub = document.getElementById("subsubcategory")?.value || "";
  const contact = document.getElementById("contactName")?.value || "";
  const callback = document.getElementById("callback")?.value || "";
  const workstation = document.getElementById("workstation")?.value || "";
  const location = document.getElementById("location")?.value || "";
  const description = document.getElementById("description")?.value || "";

  /* BUILD SUBJECT */
  let detailParts = [];

  if (dept === "IT") {
    if (sub) detailParts.push(sub);
    if (subsub) detailParts.push(subsub);
  } else {
    if (sub) detailParts.push(sub);
  }

  let subject = `Ticket – ${dept}: ${category}`;
  if (detailParts.length) subject += ` – ${detailParts.join(" – ")}`;
  if (location) subject += ` – (${location})`;

  setSubjectSafe(subject);

  /* BUILD BODY */
  let body = `
S1P Support Ticket

Department: ${dept}
Category: ${category}
${sub ? `Subcategory: ${sub}` : ""}
${subsub ? `Issue Type: ${subsub}` : ""}

Contact Name: ${contact}
Callback Number: ${callback}
${dept === "IT" ? `Workstation: ${workstation}\n` : ""}
Location Code: ${location}

Description:
${description}

------------------------------
(This ticket was generated using the S1P Outlook Add-in)
`;

  /* INSERT BODY */
  try {
    const item = Office.context?.mailbox?.item;

    if (item?.body && typeof item.body.setAsync === "function") {
      item.body.setAsync(body, { coercionType: Office.CoercionType.Text }, () => {
        closeTaskpaneSafe();
      });
    } else if (Office.context.mailbox.displayNewMessageForm) {
      Office.context.mailbox.displayNewMessageForm({
        subject,
        body
      });
      closeTaskpaneSafe();
    } else {
      console.log("Could not set body (no compose mode).");
      closeTaskpaneSafe();
    }
  } catch (e) {
    console.log("Body insertion failed:", e);
    closeTaskpaneSafe();
  }
}

/* ------------------------------------------------
   CONTACT NAME AUTOFILL
--------------------------------------------------*/

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(info => {
    if (info.host !== Office.HostType.Outlook) return;

    try {
      const profile = Office.context.mailbox.userProfile;
      const nameBox = document.getElementById("contactName");

      if (profile?.displayName && nameBox) {
        nameBox.value = profile.displayName;
        localStorage.setItem("lastContactName", profile.displayName);
      }

      const saved = localStorage.getItem("lastContactName");
      if (saved && nameBox) nameBox.value = saved;
    } catch (e) {
      console.log("Profile load error:", e);
    }
  });
}

/* ------------------------------------------------
   DOM READY
--------------------------------------------------*/

document.addEventListener("DOMContentLoaded", initForm);

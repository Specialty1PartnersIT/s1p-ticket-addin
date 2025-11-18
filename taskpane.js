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

    // Reset category + sub levels
    clearSelect(categorySelect, "-- Select a category --");
    clearSelect(subSelect, "-- Select a subcategory --");
    clearSelect(subSubSelect, "-- Select an issue type --");

    showElement(subSelect, false);
    showElement(subLabel, false);
    showElement(subSubSelect, false);
    showElement(subSubLabel, false);

    // Populate categories for chosen dept
    if (CATEGORY_MAP[dept]) {
      fillSelect(categorySelect, CATEGORY_MAP[dept]);
    }

    // Workstation section only for IT
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
      // Special 3-level flow for PMS/Imaging
      if (cat === "PMS / Practice Management System" || cat === "Imaging Software") {
        const softwareList = IT_SOFTWARE_MAP[cat] || [];
        const issueList = IT_ISSUE_TYPE_MAP[cat] || [];

        if (softwareList.length) {
          fillSelect(subSelect, softwareList);
          showElement(subSelect, true);
          showElement(subLabel, true);
        }

        if (issueList.length) {
          fillSelect(subSubSelect, issueList);
          showElement(subSubSelect, true);
          showElement(subSubLabel, true);
        }
      } else {
        // All other IT categories – single subcategory level
        const subList = IT_OTHER_SUBCATEGORY_MAP[cat];
        if (subList && subList.length) {
          fillSelect(subSelect, subList);
          showElement(subSelect, true);
          showElement(subLabel, true);
        }
      }
    } else {
      // Non-IT departments – no subcategories
      // (leave both hidden)
    }
  });

  // Submit
  submitBtn.addEventListener("click", onSubmitClick);
}

/* ------------------------------------------------
   OFFICE HELPERS
--------------------------------------------------*/

function setSubjectSafe(subjectText) {
  try {
    if (
      typeof Office === "undefined" ||
      !Office.context ||
      !Office.context.mailbox ||
      !Office.context.mailbox.item
    ) {
      return;
    }

    const item = Office.context.mailbox.item;
    const subj = item.subject;

    // Compose mode: subject is an object with setAsync
    if (subj && typeof subj.setAsync === "function") {
      subj.setAsync(subjectText, { asyncContext: null }, result => {
        if (result && result.status !== Office.AsyncResultStatus.Succeeded) {
          console.log("Subject setAsync failed:", result.error);
        }
      });
    } else if (Office.context.mailbox.displayNewMessageForm) {
      // Read mode – open a new message with the subject instead
      Office.context.mailbox.displayNewMessageForm({ subject: subjectText });
    } else {
      console.log("Subject is not writeable in this context.");
    }
  } catch (e) {
    console.log("Failed to set subject via Office.js:", e);
  }
}

function closeTaskpaneSafe() {
  try {
    if (typeof Office !== "undefined" && Office.context && Office.context.ui && Office.context.ui.closeContainer) {
      Office.context.ui.closeContainer();
    }
  } catch (e) {
    console.log("Failed to close taskpane:", e);
  }
}

/* ------------------------------------------------
   SUBMIT HANDLER
--------------------------------------------------*/

function onSubmitClick() {
  const deptSelect = document.getElementById("department");
  const categorySelect = document.getElementById("category");
  const subSelect = document.getElementById("subcategory");
  const subSubSelect = document.getElementById("subsubcategory");
  const locationInput = document.getElementById("location");

  const dept = deptSelect ? deptSelect.value : "";
  const category = categorySelect ? categorySelect.value : "";
  const locationCode = locationInput ? locationInput.value : "";

  let detailParts = [];

  if (dept === "IT") {
    const softwareOrSub = subSelect ? subSelect.value : "";
    const issueType = subSubSelect ? subSubSelect.value : "";

    if (softwareOrSub) detailParts.push(softwareOrSub);
    if (issueType) detailParts.push(issueType);
  } else {
    // For non-IT we *currently* don’t use subcategories, but we could later
    const maybeSub = subSelect ? subSelect.value : "";
    if (maybeSub) detailParts.push(maybeSub);
  }

  let subject = `Ticket – ${dept || "Dept"}: ${category || "Category"}`;
  if (detailParts.length) {
    subject += " – " + detailParts.join(" – ");
  }
  if (locationCode) {
    subject += ` – (${locationCode})`;
  }

  setSubjectSafe(subject);
  closeTaskpaneSafe();
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

      if (profile && profile.displayName && nameBox) {
        nameBox.value = profile.displayName;
        localStorage.setItem("lastContactName", profile.displayName);
      }

      const saved = localStorage.getItem("lastContactName");
      if (saved && nameBox) {
        nameBox.value = saved;
      }
    } catch (e) {
      console.log("Could not load profile name:", e);
    }
  });
}

/* ------------------------------------------------
   DOM READY
--------------------------------------------------*/

document.addEventListener("DOMContentLoaded", initForm);

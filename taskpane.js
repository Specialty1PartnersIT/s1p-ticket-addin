// ------------------------------------------------
//  DATA MAPS
// ------------------------------------------------

// High-level categories per department
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

// IT software lists (shown as Subcategory when Category is PMS/Imaging)
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
    "Other PMS"
  ],
  "Imaging Software": [
    "CRD Dicom",
    "Romexis",
    "EZ3D-i",
    "Dexis",
    "Other Imaging Software"
  ]
};

// IT issue-type lists for PMS/Imaging (shown as Issue Type)
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

// IT subcategories for the *other* IT categories
const IT_SUBCATEGORY_MAP = {
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

// ------------------------------------------------
//  INIT – run after page load
// ------------------------------------------------

function initTaskpane() {
  const deptEl = document.getElementById("department");
  const categoryEl = document.getElementById("category");
  const subcatEl = document.getElementById("subcategory");
  const subcatLabelEl = document.getElementById("subcategory-label");
  const subsubEl = document.getElementById("subsubcategory");
  const subsubLabelEl = document.getElementById("subsubcategory-label");
  const wsSectionEl = document.getElementById("ws-section");
  const contactNameEl = document.getElementById("contactName");
  const submitBtnEl = document.getElementById("submitBtn");

  // Basic sanity check
  if (!deptEl || !categoryEl || !submitBtnEl) {
    console.error("S1P ticket add-in: form elements not found – init aborted.");
    return;
  }

  // --- Office profile & saved name ---
  if (window.Office && Office.onReady) {
    Office.onReady().then(() => {
      try {
        const profile = Office.context.mailbox.userProfile;
        if (profile && profile.displayName && contactNameEl) {
          contactNameEl.value = profile.displayName;
          localStorage.setItem("lastContactName", profile.displayName);
        }
      } catch (e) {
        console.log("Could not load Outlook profile name:", e);
      }

      const savedName = localStorage.getItem("lastContactName");
      if (savedName && contactNameEl && !contactNameEl.value) {
        contactNameEl.value = savedName;
      }
    });
  }

  // --- Department change ---
  deptEl.addEventListener("change", () => {
    const dept = deptEl.value;

    // Reset dropdowns
    categoryEl.innerHTML = '<option value="">-- Select a category --</option>';
    subcatEl.innerHTML = '<option value="">-- Select a subcategory --</option>';
    subsubEl.innerHTML = '<option value="">-- Select an issue type --</option>';

    subcatEl.classList.add("hidden");
    subcatLabelEl.classList.add("hidden");
    subsubEl.classList.add("hidden");
    subsubLabelEl.classList.add("hidden");

    // Populate categories for this department
    if (CATEGORY_MAP[dept]) {
      CATEGORY_MAP[dept].forEach(cat => {
        const opt = document.createElement("option");
        opt.value = cat;
        opt.textContent = cat;
        categoryEl.appendChild(opt);
      });
    }

    // Workstation section only for IT
    if (dept === "IT") {
      wsSectionEl.classList.remove("hidden");
    } else {
      wsSectionEl.classList.add("hidden");
    }
  });

  // --- Category change ---
  categoryEl.addEventListener("change", () => {
    const dept = deptEl.value;
    const cat = categoryEl.value;

    // Reset sub-levels
    subcatEl.innerHTML = '<option value="">-- Select a subcategory --</option>';
    subsubEl.innerHTML = '<option value="">-- Select an issue type --</option>';
    subcatEl.classList.add("hidden");
    subcatLabelEl.classList.add("hidden");
    subsubEl.classList.add("hidden");
    subsubLabelEl.classList.add("hidden");

    // Only IT has subcategories
    if (dept !== "IT" || !cat) return;

    // PMS / Imaging: Software + Issue Type
    if (IT_SOFTWARE_MAP[cat]) {
      // Subcategory = software list
      IT_SOFTWARE_MAP[cat].forEach(soft => {
        const opt = document.createElement("option");
        opt.value = soft;
        opt.textContent = soft;
        subcatEl.appendChild(opt);
      });
      subcatEl.classList.remove("hidden");
      subcatLabelEl.classList.remove("hidden");

      // Issue Type dropdown
      if (IT_ISSUE_TYPE_MAP[cat]) {
        IT_ISSUE_TYPE_MAP[cat].forEach(issue => {
          const opt = document.createElement("option");
          opt.value = issue;
          opt.textContent = issue;
          subsubEl.appendChild(opt);
        });
        subsubEl.classList.remove("hidden");
        subsubLabelEl.classList.remove("hidden");
      }
      return;
    }

    // Other IT categories: single subcategory list (no Issue Type)
    if (IT_SUBCATEGORY_MAP[cat]) {
      IT_SUBCATEGORY_MAP[cat].forEach(sub => {
        const opt = document.createElement("option");
        opt.value = sub;
        opt.textContent = sub;
        subcatEl.appendChild(opt);
      });
      subcatEl.classList.remove("hidden");
      subcatLabelEl.classList.remove("hidden");
    }
  });

  // --- Submit ticket ---
  submitBtnEl.addEventListener("click", () => {
    const dept = deptEl.value || "";
    const category = categoryEl.value || "";
    const subcat = subcatEl.classList.contains("hidden") ? "" : subcatEl.value || "";
    const subsub = subsubEl.classList.contains("hidden") ? "" : subsubEl.value || "";
    const locationEl = document.getElementById("location");
    const location = locationEl ? (locationEl.value || "") : "";

    // Keep last contact name
    if (contactNameEl && contactNameEl.value) {
      localStorage.setItem("lastContactName", contactNameEl.value);
    }

    // Build subject parts
    let subject = `Ticket – ${dept}: ${category}`;
    if (subcat) subject += ` – ${subcat}`;
    if (subsub) subject += ` – ${subsub}`;
    if (location) subject += ` – (${location})`;

    // Apply subject in Outlook (if we're actually in Outlook)
    try {
      if (window.Office && Office.context && Office.context.mailbox && Office.context.mailbox.item) {
        Office.context.mailbox.item.subject.setAsync(subject);
        Office.context.ui.closeContainer();
      } else {
        console.log("Preview subject (not in Outlook):", subject);
      }
    } catch (e) {
      console.error("Failed to set subject via Office.js:", e);
    }
  });
}

// Wire init AFTER everything is loaded so getElementById works
window.addEventListener("load", initTaskpane);

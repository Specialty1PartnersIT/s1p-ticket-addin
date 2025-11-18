// ----------------------------------------------------
// Auto-fill Contact Name + Local Backup
// ----------------------------------------------------
Office.onReady(() => {
  try {
    const profile = Office.context.mailbox.userProfile;
    if (profile && profile.displayName) {
      document.getElementById("contactName").value = profile.displayName;
      localStorage.setItem("lastContactName", profile.displayName);
    }
  } catch (e) {}

  const saved = localStorage.getItem("lastContactName");
  if (saved) document.getElementById("contactName").value = saved;
});

// ----------------------------------------------------
// CATEGORY MAP (BY DEPARTMENT)
// ----------------------------------------------------
const CATEGORY_MAP = {
  "IT": [
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

// ----------------------------------------------------
// IT SUBCATEGORY MAP (ISSUE LISTS)
// ----------------------------------------------------
const IT_SUBCATEGORY_MAP = {
  "Email/Outlook": [
    "Cannot send","Cannot receive","Sync issues",
    "Calendar issue","Shared mailbox issue","Add-in issue",
    "Other (Email)"
  ],

  "Teams": [
    "Chat/Channels","Meetings/Calls","Screen sharing",
    "Teams/Channels access","Notifications","Other (Teams)"
  ],

  "SharePoint": [
    "Permissions/Access","File sync/OneDrive","Broken link",
    "Page not loading","Version history/Restore","Other (SharePoint)"
  ],

  "Hardware": [
    "Desktop/Laptop","Docking station","Monitor",
    "Printer/Scanner","Phone/Headset","Other (Hardware)"
  ],

  "Network/Internet": [
    "No connectivity","Slow network","VPN","Wi-Fi","Other (Network)"
  ],

  "Security": ["Breach","Incident","Non-Critical"],

  "RingCentral": [
    "Phone hardware","Call quality","Fax send/receive",
    "Text send/receive","Early office closure / schedule change"
  ],

  "Acumen": [
    "Access/Permissions","Data Integrity","New Reports","Report Edits"
  ],

  "Other": ["General / Not specified"]
};

// ----------------------------------------------------
// IT SOFTWARE MAP (PMS/Imaging Subcategory)
// ----------------------------------------------------
const IT_SOFTWARE_MAP = {
  "PMS / Practice Management System": [
    "EagleSoft","EndoVision","Clinivision","OMSVision",
    "WinOMS","TDO","WinDent","Dentrix","PBS Endo",
    "DSN","Denticon","Other PMS"
  ],

  "Imaging Software": [
    "CRD Dicom","Romexis","EZ3D-i","Dexis","Sidexis",
    "Carestream","XDR","Vatech","Other Imaging"
  ]
};

// ----------------------------------------------------
// IT ISSUE TYPE MAP (PMS/Imaging SUB-SUBCATEGORY)
// ----------------------------------------------------
const IT_ISSUE_TYPE_MAP = {
  "PMS / Practice Management System": [
    "Access / Permissions","Installation / Upgrade",
    "Performance / Slowness","Data / Charting Issue",
    "Integration with other systems","Other (PMS Issue)"
  ],

  "Imaging Software": [
    "Image acquisition / capture","Viewer / workstation issue",
    "Integration with PMS","Export / sharing issue",
    "Other (Imaging Issue)"
  ]
};

// ----------------------------------------------------
// DEPARTMENT CHANGE HANDLER
// ----------------------------------------------------
document.getElementById("department").addEventListener("change", function () {
  const dept = this.value;

  const category = document.getElementById("category");
  const subcat = document.getElementById("subcategory");
  const subsub = document.getElementById("subsubcategory");
  const subcatLabel = document.getElementById("subcategory-label");
  const subsubLabel = document.getElementById("subsubcategory-label");
  const wsSection = document.getElementById("ws-section");

  category.innerHTML = `<option value="">-- Select a category --</option>`;
  subcat.classList.add("hidden");
  subsub.classList.add("hidden");
  subcatLabel.classList.add("hidden");
  subsubLabel.classList.add("hidden");

  if (CATEGORY_MAP[dept]) {
    CATEGORY_MAP[dept].forEach(c => {
      category.innerHTML += `<option value="${c}">${c}</option>`;
    });
  }

  wsSection.classList.toggle("hidden", dept !== "IT");
});

// ----------------------------------------------------
// CATEGORY CHANGE HANDLER
// ----------------------------------------------------
document.getElementById("category").addEventListener("change", function () {
  const cat = this.value;
  const dept = document.getElementById("department").value;

  const subcat = document.getElementById("subcategory");
  const subsub = document.getElementById("subsubcategory");
  const subcatLabel = document.getElementById("subcategory-label");
  const subsubLabel = document.getElementById("subsubcategory-label");

  subcat.classList.add("hidden");
  subsub.classList.add("hidden");
  subcatLabel.classList.add("hidden");
  subsubLabel.classList.add("hidden");

  subcat.innerHTML = `<option value="">-- Select a subcategory --</option>`;
  subsub.innerHTML = `<option value="">-- Select an issue type --</option>`;

  // PMS/Imaging — 2-level
  if (dept === "IT" && IT_SOFTWARE_MAP[cat]) {
    subcat.classList.remove("hidden");
    subcatLabel.classList.remove("hidden");

    IT_SOFTWARE_MAP[cat].forEach(s => {
      subcat.innerHTML += `<option value="${s}">${s}</option>`;
    });
    return;
  }

  // Other IT categories — 1-level
  if (dept === "IT" && IT_SUBCATEGORY_MAP[cat]) {
    subcat.classList.remove("hidden");
    subcatLabel.classList.remove("hidden");

    IT_SUBCATEGORY_MAP[cat].forEach(s => {
      subcat.innerHTML += `<option value="${s}">${s}</option>`;
    });
  }
});

// ----------------------------------------------------
// SUBCATEGORY CHANGE HANDLER (PMS/Imaging Only)
// ----------------------------------------------------
document.getElementById("subcategory").addEventListener("change", function () {
  const cat = document.getElementById("category").value;
  const dept = document.getElementById("department").value;

  const subsub = document.getElementById("subsubcategory");
  const subsubLabel = document.getElementById("subsubcategory-label");

  subsub.classList.add("hidden");
  subsubLabel.classList.add("hidden");

  subsub.innerHTML = `<option value="">-- Select an issue type --</option>`;

  if (dept === "IT" && IT_ISSUE_TYPE_MAP[cat]) {
    subsub.classList.remove("hidden");
    subsubLabel.classList.remove("hidden");

    IT_ISSUE_TYPE_MAP[cat].forEach(i => {
      subsub.innerHTML += `<option value="${i}">${i}</option>`;
    });
  }
});

// ----------------------------------------------------
// SUBMIT BUTTON — SUBJECT BUILDER
// ----------------------------------------------------
document.getElementById("submitBtn").addEventListener("click", function () {
  const dept = document.getElementById("department").value;
  const category = document.getElementById("category").value;
  const subcat = document.getElementById("subcategory").value;
  const subsub = document.getElementById("subsubcategory").value;
  const location = document.getElementById("location").value;

  let detail = subsub || subcat;

  let subject = `Ticket – ${dept}: ${category}`;
  if (detail) subject += ` – ${detail}`;
  if (location) subject += ` – (${location})`;

  Office.context.mailbox.item.subject.setAsync(subject);
  Office.context.ui.closeContainer();
});

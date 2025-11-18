// Auto-fill Contact Name from Outlook
Office.onReady(() => {
  try {
    const profile = Office.context.mailbox.userProfile;
    if (profile && profile.displayName) {
      document.getElementById("contactName").value = profile.displayName;

      // Save to localStorage
      localStorage.setItem("lastContactName", profile.displayName);
    }
  } catch (e) {
    console.log("Could not load profile name:", e);
  }

  // Load saved contact name if available
  const savedName = localStorage.getItem("lastContactName");
  if (savedName) {
    document.getElementById("contactName").value = savedName;
  }
});

// ------------------------------
// CATEGORY + SUBCATEGORY MAPS
// ------------------------------
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
    "Leasing Inquires",
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

// IT Subcategories
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
    "Other (Imaging)"
  ],

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
  ]
};

// PMS / Imaging Systems
const SYSTEM_MAP = {
  "PMS / Practice Management System": [
    "EagleSoft", "Dentrix", "DSN", "TDO", "PBS Endo",
    "WinOMS", "WinDent", "Clinivision", "Endovision", "Other PMS"
  ],
  "Imaging Software": [
    "Romexis", "CRD Dicom", "EZ3D-i", "Dexis", "Other Imaging"
  ]
};

// ------------------------------
// Department Change Handler
// ------------------------------
document.getElementById("department").addEventListener("change", function () {
  const dept = this.value;

  const category = document.getElementById("category");
  const subcat = document.getElementById("subcategory");
  const subcatLabel = document.getElementById("subcategory-label");
  const system = document.getElementById("system");
  const systemLabel = document.getElementById("system-label");
  const wsSection = document.getElementById("ws-section");

  // Reset all dependent fields
  category.innerHTML = `<option value="">-- Select a category --</option>`;
  subcat.innerHTML = `<option value="">-- Select a subcategory --</option>`;
  subcat.classList.add("hidden");
  subcatLabel.classList.add("hidden");
  system.classList.add("hidden");
  systemLabel.classList.add("hidden");

  // Populate categories
  if (CATEGORY_MAP[dept]) {
    CATEGORY_MAP[dept].forEach(c => {
      category.innerHTML += `<option value="${c}">${c}</option>`;
    });
  }

  // Show workstation IF IT
  if (dept === "IT") {
    wsSection.classList.remove("hidden");
  } else {
    wsSection.classList.add("hidden");
  }
});

// ------------------------------
// Category Change Handler
// ------------------------------
document.getElementById("category").addEventListener("change", function () {
  const cat = this.value;
  const subcat = document.getElementById("subcategory");
  const subcatLabel = document.getElementById("subcategory-label");
  const system = document.getElementById("system");
  const systemLabel = document.getElementById("system-label");

  // Reset first
  subcat.innerHTML = `<option value="">-- Select a subcategory --</option>`;
  subcat.classList.add("hidden");
  subcatLabel.classList.add("hidden");
  system.classList.add("hidden");
  systemLabel.classList.add("hidden");

  // If this category has IT subcategories
  if (SUBCATEGORY_MAP[cat]) {
    subcat.classList.remove("hidden");
    subcatLabel.classList.remove("hidden");

    SUBCATEGORY_MAP[cat].forEach(s => {
      subcat.innerHTML += `<option value="${s}">${s}</option>`;
    });
  }

  // If this category has a system selector
  if (SYSTEM_MAP[cat]) {
    system.classList.remove("hidden");
    systemLabel.classList.remove("hidden");

    SYSTEM_MAP[cat].forEach(s => {
      system.innerHTML += `<option value="${s}">${s}</option>`;
    });
  }
});

// ------------------------------
// Submit Ticket
// ------------------------------
document.getElementById("submitBtn").addEventListener("click", function () {

  Office.context.ui.closeContainer(); // Auto-close add-in taskpane

  // Actual ticketing logic handled in back-end (email, system, etc.)
});

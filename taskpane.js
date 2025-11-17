/* global Office */

(() => {

  function byId(id) {
    return document.getElementById(id);
  }

  function clearSelect(select, placeholderText) {
    select.innerHTML = "";
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = placeholderText;
    select.appendChild(opt);
  }

  function populateSelect(select, options) {
    options.forEach(value => {
      const opt = document.createElement("option");
      opt.value = value;
      opt.textContent = value;
      select.appendChild(opt);
    });
  }

  // ---------------------------
  // Department → Email mapping
  // ---------------------------
  const DEPARTMENT_EMAIL = {
    "IT": "Support@specialty1partners.com",
    "CBO - Full Service": "cbo@specialty1partners.com",
    "Payor Relations": "payorrelations@specialty1partners.com",
    "RCM - Non Full Service": "rcm@specialty1partners.com"
  };

  // -----------------------------------
  // IT Categories → Subcategories
  // -----------------------------------
  const IT_SUBCATEGORY_MAP = {
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
      "Chat / Channels",
      "Meetings / Calls",
      "Screen sharing",
      "Teams / Channels access",
      "Notifications",
      "Other (Teams)"
    ],
    "SharePoint": [
      "Permissions / Access",
      "File sync / OneDrive",
      "Broken link",
      "Page not loading",
      "Version history / Restore",
      "Other (SharePoint)"
    ],
    "Hardware": [
      "Desktop / Laptop",
      "Docking station",
      "Monitor",
      "Printer / Scanner",
      "Phone / Headset",
      "Other (Hardware)"
    ],
    "Network/Internet": [
      "No connectivity",
      "Slow network",
      "VPN",
      "Wi-Fi",
      "Other (Network)"
    ],
    "RingCentral": [
      "Phone hardware",
      "Call quality",
      "Fax send/receive",
      "Text send/receive",
      "Early office closure / schedule changes",
      "Other (RingCentral)"
    ],
    "Security": [
      "Breach",
      "Incident",
      "Non-Critical",
      "Other (Security)"
    ],
    "Acumen": [
      "Access / Permissions",
      "Data integrity",
      "New reports",
      "Report edits",
      "Other (Acumen)"
    ],
    "Other": [
      "General / Not specified"
    ]
  };

  // -----------------------------
  // PMS / Imaging System Lists
  // -----------------------------
  const PMS_SYSTEMS = [
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
  ];

  const IMAGING_SYSTEMS = [
    "CRD Dicom",
    "Romexis",
    "EZ3D-i",
    "Other (Imaging)"
  ];

  // -------------------------
  // CBO Categories
  // -------------------------
  const CBO_CATEGORIES = [
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
  ];

  // -------------------------
  // Payor Relations Categories
  // -------------------------
  const PAYOR_RELATIONS_CATEGORIES = [
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
    "Term Network",
    "Term Provider",
    "TIN Change",
    "Update Practice Data",
    "W-9 Request",
    "Other"
  ];

  // -------------------------
  // RCM Categories
  // -------------------------
  const RCM_CATEGORIES = [
    "Collection Placement",
    "RCM General questions",
    "Refund",
    "Other"
  ];

  // -------------------------
  // Handle Department Change
  // -------------------------
  function onDepartmentChange() {
    const dept = byId("department").value;

    const categorySelect = byId("category");
    const subcategorySelect = byId("subcategory");
    const systemSelect = byId("system");

    const subLabel = byId("subcategory-label");
    const systemLabel = byId("system-label");

    const wsLabel = byId("ws-label");
    const wsInput = byId("workstation");
    const wsHint = byId("ws-hint");

    clearSelect(categorySelect, dept ? "-- Select a category --" : "-- Select a department first --");
    clearSelect(subcategorySelect, "-- Select a category first --");
    clearSelect(systemSelect, "-- Select a system --");

    subcategorySelect.classList.add("hidden");
    subLabel.classList.add("hidden");
    systemSelect.classList.add("hidden");
    systemLabel.classList.add("hidden");

    // Workstation field visible only for IT
    if (dept === "IT") {
      wsLabel.classList.remove("hidden");
      wsInput.classList.remove("hidden");
      wsHint.classList.remove("hidden");
    } else {
      wsLabel.classList.add("hidden");
      wsInput.classList.add("hidden");
      wsHint.classList.add("hidden");
    }

    if (!dept) return;

    if (dept === "IT") {
      populateSelect(categorySelect, Object.keys(IT_SUBCATEGORY_MAP));
    } else if (dept === "CBO - Full Service") {
      populateSelect(categorySelect, CBO_CATEGORIES);
    } else if (dept === "Payor Relations") {
      populateSelect(categorySelect, PAYOR_RELATIONS_CATEGORIES);
    } else if (dept === "RCM - Non Full Service") {
      populateSelect(categorySelect, RCM_CATEGORIES);
    }
  }

  // -------------------------
  // Handle Category Change
  // -------------------------
  function onCategoryChange() {
    const dept = byId("department").value;
    const category = byId("category").value;

    const subcategorySelect = byId("subcategory");
    const systemSelect = byId("system");
    const subLabel = byId("subcategory-label");
    const systemLabel = byId("system-label");

    clearSelect(subcategorySelect, "-- Select a category first --");
    clearSelect(systemSelect, "-- Select a system --");

    subcategorySelect.classList.add("hidden");
    subLabel.classList.add("hidden");
    systemSelect.classList.add("hidden");
    systemLabel.classList.add("hidden");

    if (dept !== "IT" || !category) return;

    // Subcategories for IT categories
    const subcats = IT_SUBCATEGORY_MAP[category] || [];
    if (subcats.length > 0) {
      clearSelect(subcategorySelect, "-- Select a subcategory --");
      populateSelect(subcategorySelect, subcats);
      subLabel.classList.remove("hidden");
      subcategorySelect.classList.remove("hidden");
    }

    // PMS / Imaging system list
    if (category === "PMS / Practice Management System") {
      populateSelect(systemSelect, PMS_SYSTEMS);
      systemSelect.classList.remove("hidden");
      systemLabel.classList.remove("hidden");
    } else if (category === "Imaging Software") {
      populateSelect(systemSelect, IMAGING_SYSTEMS);
      systemSelect.classList.remove("hidden");
      systemLabel.classList.remove("hidden");
    }
  }

  // -------------------------
  // Submit Ticket
  // -------------------------
  function onSubmit() {
    const dept = byId("department").value;
    const category = byId("category").value;
    const subcategory = byId("subcategory").classList.contains("hidden")
      ? ""
      : byId("subcategory").value;

    const system = byId("system").classList.contains("hidden")
      ? ""
      : byId("system").value;

    const contactName = byId("contactName").value.trim();
    const workstation = byId("workstation").value.trim();
    const callback = byId("callback").value.trim();
    const location = byId("location").value.trim();
    const description = byId("description").value.trim();

    // Required fields
    if (!dept) return alert("Please select a Department.");
    if (!category) return alert("Please select a Category.");
    if (!contactName) return alert("Please enter the Contact Name.");
    if (dept === "IT" && !workstation) return alert("Please enter the Workstation Name.");
    if (!callback) return alert("Please enter a Callback Number.");

    const toEmail = DEPARTMENT_EMAIL[dept] || DEPARTMENT_EMAIL["IT"];

    let subject = `Ticket – ${dept}: ${category}`;
    if (subcategory) subject += ` – ${subcategory}`;
    if (location) subject += ` – (${location})`;

    let body = "";
    body += `<b>Department:</b> ${dept}<br/>`;
    body += `<b>Category:</b> ${category}<br/>`;

    if (subcategory) body += `<b>Subcategory:</b> ${subcategory}<br/>`;
    if (system) body += `<b>PMS / Imaging System:</b> ${system}<br/>`;

    body += `<b>Contact Name:</b> ${contactName}<br/>`;

    if (dept === "IT") {
      body += `<b>Workstation:</b> ${workstation}<br/>`;
    }

    if (location) body += `<b>Location Code:</b> ${location}<br/>`;
    body += `<b>Callback Number:</b> ${callback}<br/><br/>`;
    body += `<b>Description:</b><br/>${description.replace(/\n/g,"<br/>")}`;

    Office.context.mailbox.displayNewMessageForm({
      toRecipients: [toEmail],
      subject,
      htmlBody: body
    });

    // Close the taskpane after sending
    if (Office.context.ui && Office.context.ui.closeContainer) {
      Office.context.ui.closeContainer();
    }
  }

  // -------------------------
  // Wire events + Auto-fill Contact Name + Auto-save Contact Name
  // -------------------------
  function wireEvents() {
    byId("department").addEventListener("change", onDepartmentChange);
    byId("category").addEventListener("change", onCategoryChange);
    byId("submitBtn").addEventListener("click", onSubmit);

    const contactInput = byId("contactName");

    // 1. Load saved name
    const savedName = localStorage.getItem("contactName");
    if (savedName) {
      contactInput.value = savedName;
    } else {
      // 2. Fallback → load from Outlook profile
      const userProfile = Office.context.mailbox?.userProfile;
      if (userProfile && userProfile.displayName) {
        contactInput.value = userProfile.displayName;
      }
    }

    // 3. Save automatically as user types
    contactInput.addEventListener("input", () => {
      localStorage.setItem("contactName", contactInput.value.trim());
    });
  }

  // -------------------------
  // Initialize Add-in
  // -------------------------
  Office.onReady(() => {
    if (document.readyState === "complete") wireEvents();
    else document.addEventListener("DOMContentLoaded", wireEvents);
  });

})();

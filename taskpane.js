// taskpane.js

/* global Office */

(() => {
  // Map department -> email address
  const DEPARTMENT_EMAIL = {
    "IT": "Support@specialty1partners.com",
    "CBO - Full Service": "cbo@specialty1partners.com",
    "Payor Relations": "payorrelations@specialty1partners.com",
    "RCM - Non Full Service": "rcm@specialty1partners.com"
  };

  // IT category -> subcategories (issue type level)
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

  // PMS & Imaging systems (sub-subcategory)
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

  // Department-specific category lists (no subcategories for these three)
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

  const PAYOR_RELATIONS_CATEGORIES = [
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
    "Term Network",
    "Term Provider",
    "TIN Change",
    "Update Practice Data",
    "W-9 Request",
    "Other"
  ];

  const RCM_CATEGORIES = [
    "Collection Placement",
    "RCM General questions",
    "Refund",
    "Other"
  ];

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

  // Handle Department change
  function onDepartmentChange() {
    const deptSelect = byId("department");
    const dept = deptSelect.value;

    const categorySelect = byId("category");
    const subcategorySelect = byId("subcategory");
    const systemSelect = byId("system");
    const subLabel = byId("subcategory-label");
    const systemLabel = byId("system-label");

    // Reset category & below
    clearSelect(categorySelect, dept ? "-- Select a category --" : "-- Select a department first --");

    // Hide subcategory & system by default
    subcategorySelect.classList.add("hidden");
    subLabel.classList.add("hidden");
    systemSelect.classList.add("hidden");
    systemLabel.classList.add("hidden");
    clearSelect(subcategorySelect, "-- Select a category first --");
    clearSelect(systemSelect, "-- Select a system --");

    if (!dept) {
      return;
    }

    if (dept === "IT") {
      // IT uses categories + subcategories
      const itCategories = Object.keys(IT_SUBCATEGORY_MAP);
      populateSelect(categorySelect, itCategories);
      categorySelect.disabled = false;
    } else if (dept === "CBO - Full Service") {
      populateSelect(categorySelect, CBO_CATEGORIES);
      categorySelect.disabled = false;
    } else if (dept === "Payor Relations") {
      populateSelect(categorySelect, PAYOR_RELATIONS_CATEGORIES);
      categorySelect.disabled = false;
    } else if (dept === "RCM - Non Full Service") {
      populateSelect(categorySelect, RCM_CATEGORIES);
      categorySelect.disabled = false;
    }
  }

  // Handle Category change (primarily for IT)
  function onCategoryChange() {
    const dept = byId("department").value;
    const category = byId("category").value;

    const subcategorySelect = byId("subcategory");
    const systemSelect = byId("system");
    const subLabel = byId("subcategory-label");
    const systemLabel = byId("system-label");

    // Reset
    clearSelect(subcategorySelect, "-- Select a category first --");
    clearSelect(systemSelect, "-- Select a system --");
    subcategorySelect.classList.add("hidden");
    subLabel.classList.add("hidden");
    systemSelect.classList.add("hidden");
    systemLabel.classList.add("hidden");

    if (!dept || !category) {
      return;
    }

    if (dept !== "IT") {
      // Non-IT departments: no subcategories
      return;
    }

    // IT department: show subcategories for this category
    const subcats = IT_SUBCATEGORY_MAP[category] || [];
    if (subcats.length > 0) {
      clearSelect(subcategorySelect, "-- Select a subcategory --");
      populateSelect(subcategorySelect, subcats);
      subcategorySelect.classList.remove("hidden");
      subLabel.classList.remove("hidden");
    }

    // PMS / Imaging specific system dropdown
    if (category === "PMS / Practice Management System") {
      clearSelect(systemSelect, "-- Select a PMS system --");
      populateSelect(systemSelect, PMS_SYSTEMS);
      systemSelect.classList.remove("hidden");
      systemLabel.classList.remove("hidden");
    } else if (category === "Imaging Software") {
      clearSelect(systemSelect, "-- Select an imaging system --");
      populateSelect(systemSelect, IMAGING_SYSTEMS);
      systemSelect.classList.remove("hidden");
      systemLabel.classList.remove("hidden");
    }
  }

  // Build and open the email
  function onSubmit() {
    const dept = byId("department").value;
    const category = byId("category").value;
    const subcategory = byId("subcategory").classList.contains("hidden")
      ? ""
      : byId("subcategory").value;
    const system = byId("system").classList.contains("hidden")
      ? ""
      : byId("system").value;

    const workstation = byId("workstation").value.trim();
    const callback = byId("callback").value.trim();
    const location = byId("location").value.trim();
    const description = byId("description").value.trim();

    // Basic validation
    if (!dept) {
      alert("Please select a Department.");
      return;
    }
    if (!category) {
      alert("Please select a Category.");
      return;
    }
    if (!workstation) {
      alert("Please enter the Workstation Name.");
      return;
    }
    if (!callback) {
      alert("Please enter a Callback Number.");
      return;
    }

    const toEmail = DEPARTMENT_EMAIL[dept] || DEPARTMENT_EMAIL["IT"];

    // Build subject: Ticket – {Department}: {Category} – {Subcategory} – (Location)
    const subjectParts = [];
    subjectParts.push(`Ticket – ${dept}: ${category}`);
    if (subcategory) {
      subjectParts.push(subcategory);
    }

    let subject = subjectParts.join(" – ");
    if (location) {
      subject += ` – (${location})`;
    }

    // Build HTML body
    let body = "";
    body += `<b>Department:</b> ${dept}<br/>`;
    body += `<b>Category:</b> ${category}<br/>`;
    if (subcategory) {
      body += `<b>Subcategory:</b> ${subcategory}<br/>`;
    }
    if (system) {
      body += `<b>PMS / Imaging System:</b> ${system}<br/>`;
    }
    if (location) {
      body += `<b>Location Code:</b> ${location}<br/>`;
    }
    body += `<b>Workstation:</b> ${workstation}<br/>`;
    body += `<b>Callback Number:</b> ${callback}<br/><br/>`;
    body += `<b>Description:</b><br/>`;
    body += (description ? description.replace(/\n/g, "<br/>") : "(no description provided)");

    Office.context.mailbox.displayNewMessageForm({
      toRecipients: [toEmail],
      subject: subject,
      htmlBody: body
    });
  }

  function wireEvents() {
    const deptSelect = byId("department");
    const categorySelect = byId("category");
    const submitBtn = byId("submitBtn");

    deptSelect.addEventListener("change", onDepartmentChange);
    categorySelect.addEventListener("change", onCategoryChange);
    submitBtn.addEventListener("click", onSubmit);
  }

  Office.onReady(() => {
    if (document.readyState === "complete" || document.readyState === "interactive") {
      wireEvents();
    } else {
      document.addEventListener("DOMContentLoaded", wireEvents);
    }
  });
})();

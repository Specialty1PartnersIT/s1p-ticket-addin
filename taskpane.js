// Run when Office is ready
Office.onReady(() => {
  const submitBtn = document.getElementById("submitBtn");
  const categorySelect = document.getElementById("category");
  const subcategorySelect = document.getElementById("subcategory");

  if (submitBtn) {
    submitBtn.onclick = sendTicket;
  }

  if (categorySelect && subcategorySelect) {
    categorySelect.addEventListener("change", () => {
      updateSubcategories(categorySelect.value, subcategorySelect);
    });
  }
});

// Hard-coded subcategory map (Option A)
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
  "Other": [
    "General / Not specified"
  ]
};

/**
 * Populate the subcategory dropdown based on the selected main category.
 */
function updateSubcategories(category, subcategorySelect) {
  // Clear existing options
  subcategorySelect.innerHTML = "";

  if (!category) {
    // No category selected
    subcategorySelect.disabled = true;
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = "-- Select a category first --";
    subcategorySelect.appendChild(opt);
    return;
  }

  const subcats = SUBCATEGORY_MAP[category];

  if (!subcats || subcats.length === 0) {
    // Fallback if somehow category has no mapping
    subcategorySelect.disabled = false;
    const opt = document.createElement("option");
    opt.value = "General / Not specified";
    opt.textContent = "General / Not specified";
    subcategorySelect.appendChild(opt);
    return;
  }

  // Build subcategory list
  subcategorySelect.disabled = false;

  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "-- Select a subcategory --";
  subcategorySelect.appendChild(placeholder);

  subcats.forEach(sc => {
    const opt = document.createElement("option");
    opt.value = sc;
    opt.textContent = sc;
    subcategorySelect.appendChild(opt);
  });
}

/**
 * Build and open the email to the support mailbox.
 */
function sendTicket() {
  const category = document.getElementById("category")?.value || "";
  const subcategory = document.getElementById("subcategory")?.value || "";
  const location = document.getElementById("location")?.value || "";
  const description = document.getElementById("description")?.value || "";

  const user = Office.context.mailbox.userProfile.displayName;
  const email = Office.context.mailbox.userProfile.emailAddress;

  // Build subject with category + subcategory if present
  let subject = "New IT Support Ticket";
  if (category) {
    subject += " - " + category;
  }
  if (subcategory) {
    subject += " - " + subcategory;
  }

  const htmlBody = `
    <p><b>User:</b> ${user} (${email})</p>
    <p><b>Category:</b> ${category || "N/A"}</p>
    <p><b>Subcategory:</b> ${subcategory || "N/A"}</p>
    <p><b>Location Code:</b> ${location || "N/A"}</p>
    <p><b>Description:</b><br>${description || ""}</p>
  `;

  Office.context.mailbox.displayNewMessageForm({
    toRecipients: ["support@specialty1partners.com"],
    subject: subject,
    htmlBody: htmlBody
  });
}

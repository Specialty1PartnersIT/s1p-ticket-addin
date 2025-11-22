/* ------------------------------------------------
   CONSTANTS / CONFIG
--------------------------------------------------*/

// IT mailbox only for mobile version
const IT_SUPPORT_ADDRESS = "support@specialty1partners.com";

/*
  Categories are simplified, outage-focused, and IT-only:
  - Network/Internet
  - Server / System Down
  - Workstation / Computer
  - Power / Facility
  - Other
*/
const MOBILE_CATEGORY_MAP = {
  "Network/Internet": [
    "Entire office offline",
    "Some users offline",
    "Slow network",
    "VPN down",
    "Wi-Fi down",
    "Other network issue"
  ],
  "Server / System Down": [
    "Practice management system down",
    "Imaging system down",
    "File shares / drives unavailable",
    "Remote desktop / app server down",
    "Other server/system issue"
  ],
  "Workstation / Computer": [
    "Won’t power on",
    "Won’t boot to Windows",
    "Login issue",
    "Very slow / freezing",
    "Other workstation issue"
  ],
  "Power / Facility": [
    "Full office power outage",
    "Partial outage (some rooms)",
    "Electrical event / breaker",
    "Other facility issue"
  ],
  "Other": [
    "General / not listed"
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

/**
 * Open a new email with subject/body prefilled.
 * Uses Outlook's displayNewMessageForm when available.
 */
function openTicketEmail(subject, body) {
  try {
    if (
      typeof Office !== "undefined" &&
      Office.context &&
      Office.context.mailbox &&
      typeof Office.context.mailbox.displayNewMessageForm === "function"
    ) {
      Office.context.mailbox.displayNewMessageForm({
        toRecipients: [IT_SUPPORT_ADDRESS],
        subject: subject,
        body: body
      });
    } else {
      // Browser/testing fallback
      console.log("Would open email:", {
        to: IT_SUPPORT_ADDRESS,
        subject,
        body
      });
      alert("Ticket composed.\n\nSubject:\n" + subject + "\n\nBody:\n" + body);
    }
  } catch (e) {
    console.log("Failed to open ticket email:", e);
  }
}

/**
 * Close taskpane safely when running in Outlook.
 */
function closeTaskpaneSafe() {
  try {
    if (
      typeof Office !== "undefined" &&
      Office.context &&
      Office.context.ui &&
      typeof Office.context.ui.closeContainer === "function"
    ) {
      Office.context.ui.closeContainer();
    }
  } catch (e) {
    console.log("Failed to close taskpane:", e);
  }
}

/* ------------------------------------------------
   MAIN FORM LOGIC
--------------------------------------------------*/

function initMobileForm() {
  const contactName = document.getElementById("contactName");
  const callback = document.getElementById("callback");
  const locationInput = document.getElementById("location");
  const categorySelect = document.getElementById("category");
  const subcategorySelect = document.getElementById("subcategory");
  const criticalCheckbox = document.getElementById("critical");
  const description = document.getElementById("description");
  const submitBtn = document.getElementById("submitBtn");

  if (!categorySelect || !subcategorySelect || !submitBtn) {
    console.log("Mobile ticket form elements not found. Check taskpane_mobile.html.");
    return;
  }

  // Populate categories for mobile IT-only view
  clearSelect(categorySelect, "-- Select a category --");
  fillSelect(categorySelect, Object.keys(MOBILE_CATEGORY_MAP));

  // Category change -> populate subcategories
  categorySelect.addEventListener("change", () => {
    const cat = categorySelect.value;
    clearSelect(subcategorySelect, "-- Select a subcategory (optional) --");
    const subList = MOBILE_CATEGORY_MAP[cat] || [];
    fillSelect(subcategorySelect, subList);
  });

  // Submit handler
  submitBtn.addEventListener("click", () => {
    const nameVal = (contactName && contactName.value || "").trim();
    const callbackVal = (callback && callback.value || "").trim();
    const locationVal = (locationInput && locationInput.value || "").trim();
    const categoryVal = (categorySelect && categorySelect.value || "").trim();
    const subcategoryVal = (subcategorySelect && subcategorySelect.value || "").trim();
    const isCritical = criticalCheckbox && criticalCheckbox.checked;
    const descVal = (description && description.value || "").trim();

    // Basic validation (keep strict only for key fields)
    if (!nameVal) {
      alert("Please enter your contact name.");
      return;
    }
    if (!callbackVal) {
      alert("Please enter a callback number.");
      return;
    }
    if (!categoryVal) {
      alert("Please select a category.");
      return;
    }

    // Build subject
    let subject = `Ticket – IT (Mobile): ${categoryVal}`;
    if (subcategoryVal) subject += ` – ${subcategoryVal}`;
    if (locationVal) subject += ` – (${locationVal})`;
    if (isCritical) subject = `[CRITICAL] ${subject}`;

    // Build body
    let bodyLines = [];
    bodyLines.push("S1P IT Mobile Ticket");
    bodyLines.push("====================");
    bodyLines.push("");
    bodyLines.push(`Contact Name: ${nameVal}`);
    bodyLines.push(`Callback Number: ${callbackVal}`);
    bodyLines.push(`Location Code: ${locationVal || "N/A"}`);
    bodyLines.push("");
    bodyLines.push(`Category: ${categoryVal}`);
    bodyLines.push(`Subcategory: ${subcategoryVal || "N/A"}`);
    bodyLines.push(`Critical: ${isCritical ? "Yes" : "No"}`);
    bodyLines.push("");
    bodyLines.push("Description:");
    bodyLines.push(descVal || "(No additional description provided.)");
    bodyLines.push("");
    bodyLines.push("Submitted via: Outlook Mobile Add-In");

    const body = bodyLines.join("\n");

    // Open email & close pane
    openTicketEmail(subject, body);
    closeTaskpaneSafe();
  });
}

/* ------------------------------------------------
   CONTACT NAME AUTOFILL (Outlook)
--------------------------------------------------*/

function initContactNameAutofill() {
  if (typeof Office === "undefined" || !Office.onReady) return;

  Office.onReady(info => {
    if (info.host !== Office.HostType.Outlook) return;

    try {
      const profile = Office.context.mailbox.userProfile;
      const nameBox = document.getElementById("contactName");

      if (profile && profile.displayName && nameBox) {
        nameBox.value = profile.displayName;
        // Optionally store last used name
        localStorage.setItem("s1p_mobile_lastContactName", profile.displayName);
      } else {
        // Fallback to last saved name
        const saved = localStorage.getItem("s1p_mobile_lastContactName");
        if (saved && nameBox) {
          nameBox.value = saved;
        }
      }
    } catch (e) {
      console.log("Could not load profile name in mobile add-in:", e);
    }
  });
}

/* ------------------------------------------------
   DOM READY
--------------------------------------------------*/

document.addEventListener("DOMContentLoaded", () => {
  initMobileForm();
  initContactNameAutofill();
});

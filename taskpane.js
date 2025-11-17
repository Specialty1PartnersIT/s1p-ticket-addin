Office.onReady(() => {
    document.getElementById("submitBtn").onclick = sendTicket;
});

/* ------------------------------------------
   Sub-Category Lists
------------------------------------------- */
const subcategories = {
    pms: [
        "Dentrix",
        "Eaglesoft",
        "OpenDental",
        "DentiMax",
        "Practice-Web"
    ],
    imaging: [
        "Dexis",
        "Sidexis",
        "Carestream",
        "XDR",
        "Vatech EzDent",
        "VistaScan"
    ]
};

/* ------------------------------------------
   Handle Category → Sub-Category logic
------------------------------------------- */
document.getElementById("category").addEventListener("change", function () {
    const subSelect = document.getElementById("subcategory");
    const selected = this.value;

    // Reset first
    subSelect.innerHTML = `<option value="">-- Select a Sub-Category --</option>`;
    subSelect.disabled = true;

    // If this category has no subcategories, stop here
    if (!subcategories[selected]) return;

    // Populate subcategory dropdown
    subcategories[selected].forEach(item => {
        const opt = document.createElement("option");
        opt.value = item;
        opt.textContent = item;
        subSelect.appendChild(opt);
    });

    subSelect.disabled = false;
});

/* ------------------------------------------
   Send Ticket Email
------------------------------------------- */
function sendTicket() {
    const category = document.getElementById("category").value || "Not selected";
    const subcategory = document.getElementById("subcategory").value || "Not selected";
    const location = document.getElementById("location").value || "Not provided";
    const description = document.getElementById("description").value || "No description";

    const user = Office.context.mailbox.userProfile.displayName;
    const email = Office.context.mailbox.userProfile.emailAddress;

    const htmlBody = `
        <p><b>User:</b> ${user} (${email})</p>
        <p><b>Category:</b> ${category}</p>
        <p><b>Sub-Category:</b> ${subcategory}</p>
        <p><b>Location Code:</b> ${location}</p>
        <p><b>Description:</b><br>${description}</p>
    `;

    Office.context.mailbox.displayNewMessageForm({
        toRecipients: ["support@specialty1partners.com"],
        subject: `New IT Support Ticket - ${category}`,
        htmlBody: htmlBody
    });
}

Office.onReady(() => {
  document.getElementById("submitBtn").onclick = sendTicket;
});

function sendTicket() {
  const category = document.getElementById("category").value;
  const location = document.getElementById("location").value;
  const description = document.getElementById("description").value;

  const user = Office.context.mailbox.userProfile.displayName;
  const email = Office.context.mailbox.userProfile.emailAddress;

  const htmlBody = `
    <p><b>User:</b> ${user} (${email})</p>
    <p><b>Category:</b> ${category}</p>
    <p><b>Location Code:</b> ${location}</p>
    <p><b>Description:</b><br>${description}</p>
  `;

  Office.context.mailbox.displayNewMessageForm({
    toRecipients: ["support@specialty1partners.com"],
    subject: `New IT Support Ticket - ${category}`,
    htmlBody: htmlBody
  });
}

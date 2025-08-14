const { EmailClient } = require("@azure/communication-email");

/**
 * Helper to safely get a string value from an object. If the key
 * does not exist or the value is undefined/null, returns the fallback.
 *
 * @param {Object} obj Source object
 * @param {string} key Key to retrieve
 * @param {string} fallback Fallback value if key is missing
 */
function getString(obj, key, fallback = "") {
  const val = obj[key];
  return val !== undefined && val !== null ? String(val) : fallback;
}

/**
 * Azure Function entry point. Receives JSON data from the SWMS form,
 * composes a plain text and HTML email summarising the submission and
 * sends it via Azure Communication Services. No PDF is generated in
 * this variant – the goal is to keep dependencies light and ensure
 * deployment succeeds in Azure Static Web Apps environments.
 *
 * Expected keys include taskName, location, name, email, company,
 * phone, date, swmsId and notes. Alternative keys used by the front
 * end (like job_description or job_location) are also checked to
 * populate missing values.
 *
 * Required environment variables:
 *  - COMMUNICATION_SERVICES_CONNECTION_STRING: connection string for ACS
 *  - EMAIL_SENDER: verified sender address
 *  - EMAIL_RECIPIENT (optional): default recipient when form does not provide an email
 */
module.exports = async function (context, req) {
  try {
    const connStr = process.env.COMMUNICATION_SERVICES_CONNECTION_STRING;
    const sender = process.env.EMAIL_SENDER;
    const defaultRecipient = process.env.EMAIL_RECIPIENT;

    if (!connStr) {
      throw new Error("Missing COMMUNICATION_SERVICES_CONNECTION_STRING app setting");
    }
    if (!sender) {
      throw new Error("Missing EMAIL_SENDER app setting");
    }

    const body = req.body || {};

    // Extract fields with fallbacks to alternate names and defaults
    const taskName = getString(body, "taskName") || getString(body, "job_description") || getString(body, "task") || "SWMS Task";
    const location = getString(body, "location") || getString(body, "job_location") || "Unknown location";
    const name = getString(body, "name") || getString(body, "site_manager") || getString(body, "pc_representative") || getString(body, "submitted_by") || "User";
    const userEmail = getString(body, "email") || getString(body, "manager_email") || defaultRecipient;
    const company = getString(body, "company");
    const phone = getString(body, "phone");
    const date = getString(body, "date") || getString(body, "date_provided") || new Date().toISOString().slice(0, 10);
    const swmsId = getString(body, "swmsId");
    const notes = getString(body, "notes");

    if (!userEmail) {
      throw new Error("An email address must be provided either in the form or via EMAIL_RECIPIENT");
    }

    // Compose plain text body
    const plainText = [
      `A SWMS form has been submitted:`,
      `Task: ${taskName}`,
      `Location: ${location}`,
      `Name: ${name}`,
      `Email: ${userEmail}`,
      `Company: ${company}`,
      `Phone: ${phone}`,
      `Date: ${date}`,
      `SWMS ID: ${swmsId}`,
      `Notes: ${notes}`
    ].join("\n");

    // Compose HTML body for richer formatting
    const htmlBody = `\n      <h3>SWMS form submission</h3>\n      <ul>\n        <li><strong>Task:</strong> ${taskName}</li>\n        <li><strong>Location:</strong> ${location}</li>\n        <li><strong>Name:</strong> ${name}</li>\n        <li><strong>Email:</strong> ${userEmail}</li>\n        <li><strong>Company:</strong> ${company}</li>\n        <li><strong>Phone:</strong> ${phone}</li>\n        <li><strong>Date:</strong> ${date}</li>\n        <li><strong>SWMS ID:</strong> ${swmsId}</li>\n        <li><strong>Notes:</strong> ${notes}</li>\n      </ul>\n    `;

    // Build the email message
    const emailClient = new EmailClient(connStr);
    const message = {
      sender,
      content: {
        subject: `SWMS Form Submission: ${taskName}`,
        plainText,
        html: htmlBody
      },
      toRecipients: [
        {
          address: userEmail,
          displayName: name
        },
        {
          address: defaultRecipient || sender,
          displayName: "SWMS Admin"
        }
      ]
    };

    // Send the email via Azure Communication Services
    const result = await emailClient.send(message);

    context.res = {
      status: 200,
      body: {
        message: "Email sent successfully",
        id: result?.messageId || null
      }
    };
  } catch (err) {
    context.log.error(err);
    context.res = {
      status: 500,
      body: {
        error: err.message || "An error occurred"
      }
    };
  }
};
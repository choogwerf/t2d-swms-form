const { EmailClient } = require("@azure/communication-email");
const { chromium } = require("playwright");
const mustache = require("mustache");
const fs = require("fs");
const path = require("path");

// Helper to pick certain keys (unused but provided for future extension)
const pick = (obj, keys) =>
  keys.reduce((acc, key) => {
    if (obj && obj[key] != null) acc[key] = obj[key];
    return acc;
  }, {});

module.exports = async function (context, req) {
  // Environment variables from Azure Function configuration
  const connectionString = process.env['COMMUNICATION_SERVICES_CONNECTION_STRING'];
  const sender = process.env['EMAIL_SENDER'];
  const recipient = process.env['EMAIL_RECIPIENT'];

  if (!connectionString || !sender || !recipient) {
    context.res = {
      status: 500,
      body: { error: 'Missing environment variables: COMMUNICATION_SERVICES_CONNECTION_STRING, EMAIL_SENDER, or EMAIL_RECIPIENT' },
    };
    return;
  }

  const data = req.body || {};
  // Allow alternative field names from the form submission
  const taskName = data.taskName || data.job_description || "N/A";
  const location = data.location || data.job_location || "N/A";
  const name = data.name || data.site_name || "N/A";
  const date = data.date || new Date().toISOString().split('T')[0];
  const notes = data.notes || data.description || "";
  const signature = data.signature || data.signatureData || "";

  // Read the HTML template for the SWMS PDF
  const templatePath = path.join(__dirname, "swms.html");
  const template = fs.readFileSync(templatePath, "utf8");

  const html = mustache.render(template, {
    taskName,
    location,
    name,
    date,
    notes,
    signature,
    now: new Date().toLocaleString('en-AU', { timeZone: 'Australia/Sydney' }),
  });

  try {
    // Generate PDF using Playwright
    const browser = await chromium.launch();
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: 'load' });
    const pdfBuffer = await page.pdf({ format: 'A4' });
    await browser.close();

    // Compose and send email via Azure Communication Services
    const emailClient = new EmailClient(connectionString);
    const message = {
      sender,
      content: {
        subject: `SWMS Submission: ${taskName}`,
        plainText: 'A SWMS form has been submitted. See attached PDF.',
        html: '<p>A SWMS form has been submitted. See attached PDF.</p>',
      },
      toRecipients: [ { address: recipient } ],
      attachments: [
        {
          name: `swms-${Date.now()}.pdf`,
          contentType: 'application/pdf',
          contentBytes: pdfBuffer.toString('base64'),
        },
      ],
    };

    const result = await emailClient.send(message);

    context.res = {
      status: 200,
      body: { message: 'Email sent successfully', messageId: result.messageId },
    };
  } catch (err) {
    context.res = {
      status: 500,
      body: { error: err.message || 'Error generating PDF or sending email' },
    };
  }
};

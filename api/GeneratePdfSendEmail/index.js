const { EmailClient } = require("@azure/communication-email");
const { PDFDocument, rgb, StandardFonts } = require('pdf-lib');

// Note: we no longer use Playwright to render HTML to PDF.  Instead we
// construct a simple PDF using the pdf-lib library.  This approach avoids
// heavy headless browser dependencies that often fail to build in Azure
// Static Web App environments.


/**
 * Helper to pick a subset of keys from an object and coerce values to strings.
 * Undefined or null values are ignored.
 * @param {Object} obj Source object
 * @param {string[]} keys Keys to pick
 * @returns {Object}
 */
function pick(obj, keys) {
  const out = {};
  keys.forEach(key => {
    const val = obj[key];
    if (val !== undefined && val !== null) {
      out[key] = String(val);
    }
  });
  return out;
}

/**
 * HTTP-triggered Azure Function which accepts form data, renders an HTML
 * template, converts it to a PDF and emails it using Azure Communication
 * Services. Expected JSON body keys include taskName, location, name,
 * email, company, phone, date, swmsId, notes and signatureDataUrl.
 *
 * Required environment variables:
 *  - COMMUNICATION_SERVICES_CONNECTION_STRING: connection string for ACS
 *  - EMAIL_SENDER: verified email address used as the sender
 *  - EMAIL_RECIPIENT (optional): default recipient address
 */
module.exports = async function (context, req) {
  try {
    // Ensure required settings are present
    const connStr = process.env.COMMUNICATION_SERVICES_CONNECTION_STRING;
    const sender = process.env.EMAIL_SENDER;
    if (!connStr) {
      throw new Error("Missing COMMUNICATION_SERVICES_CONNECTION_STRING app setting");
    }
    if (!sender) {
      throw new Error("Missing EMAIL_SENDER app setting");
    }

    // Extract and validate body
    const body = req.body || {};
    const data = pick(body, [
      "taskName",
      "location",
      "name",
      "email",
      "company",
      "phone",
      "date",
      "swmsId",
      "notes",
      "signatureDataUrl"
    ]);
    // Provide sensible defaults for important fields.  Many of the input
    // field names in the client form differ from the API expectations, so we
    // attempt to derive them from alternate keys before falling back to
    // placeholders.  Only the sender email is truly mandatory.
    data.taskName = data.taskName || body.job_description || body.task || 'SWMS Task';
    data.location = data.location || body.job_location || body.location || 'Unknown location';
    data.name = data.name || body.site_manager || body.pc_representative || body.submitted_by || 'User';
    data.email = data.email || body.email || body.manager_email || process.env.EMAIL_RECIPIENT;
    data.date = data.date || body.date_provided || new Date().toISOString().slice(0, 10);

    if (!data.email) {
      throw new Error('An email address must be provided either in the form or via EMAIL_RECIPIENT');
    }

    // Generate a simple PDF using pdf-lib.  We draw each field on the
    // document, and embed the signature image (if present).  This avoids
    // relying on Playwright or a headless browser.
    const pdfDoc = await PDFDocument.create();
    // A4 page size in points (72 points per inch).  A4 is 8.27 x 11.69 inches.
    const page = pdfDoc.addPage([595.28, 841.89]);
    const { width, height } = page.getSize();
    const margin = 40;
    // Load a standard font
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontSize = 12;
    let yPos = height - margin;
    // Helper to draw a label and value
    function drawLine(label, value) {
      const text = `${label}: ${value || ''}`;
      page.drawText(text, { x: margin, y: yPos, size: fontSize, font, color: rgb(0, 0, 0) });
      yPos -= fontSize + 6;
    }
    drawLine('Task', data.taskName);
    drawLine('Location', data.location);
    drawLine('Name', data.name);
    drawLine('Email', data.email);
    drawLine('Company', data.company);
    drawLine('Phone', data.phone);
    drawLine('Date', data.date);
    drawLine('SWMS ID', data.swmsId);
    drawLine('Notes', data.notes);

    // Draw a blank line before signature
    yPos -= 10;
    page.drawText('Signature:', { x: margin, y: yPos, size: fontSize, font, color: rgb(0, 0, 0) });
    yPos -= fontSize + 6;

    // If a signature data URL is provided, embed it
    if (data.signatureDataUrl && typeof data.signatureDataUrl === 'string') {
      try {
        const commaIndex = data.signatureDataUrl.indexOf(',');
        const base64 = data.signatureDataUrl.slice(commaIndex + 1);
        const mime = data.signatureDataUrl.slice(5, commaIndex);
        const imageBytes = Buffer.from(base64, 'base64');
        let image;
        if (mime.includes('png')) {
          image = await pdfDoc.embedPng(imageBytes);
        } else {
          image = await pdfDoc.embedJpg(imageBytes);
        }
        const sigWidth = 200;
        const sigHeight = (image.height / image.width) * sigWidth;
        page.drawImage(image, { x: margin, y: yPos - sigHeight, width: sigWidth, height: sigHeight });
        yPos -= sigHeight + 10;
      } catch (e) {
        // If embedding fails, skip the signature silently
        context.log.warn('Failed to embed signature image:', e);
      }
    } else {
      // Draw a line for signature if not provided
      page.drawLine({ start: { x: margin, y: yPos }, end: { x: margin + 200, y: yPos }, color: rgb(0, 0, 0), thickness: 1 });
      yPos -= 20;
    }

    // Save PDF to a Uint8Array
    const pdfBytes = await pdfDoc.save();

    // Build the email message with PDF attachment
    const fileName = `swms-${Date.now()}.pdf`;
    const emailClient = new EmailClient(connStr);
    const message = {
      sender: sender,
      content: {
        subject: `SWMS Form Submission: ${data.taskName}`,
        plainText: `A SWMS form has been submitted. See attached PDF.`,
        html: `<p>A SWMS form has been submitted.</p>`
      },
      toRecipients: [
        {
          address: data.email,
          displayName: data.name
        },
        {
          address: process.env.EMAIL_RECIPIENT || sender,
          displayName: "SWMS Admin"
        }
      ],
      attachments: [
        {
          name: fileName,
          contentType: "application/pdf",
          contentBytes: Buffer.from(pdfBytes).toString('base64')
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
    // Log and return an error
    context.log.error(err);
    context.res = {
      status: 500,
      body: {
        error: err.message || "An error occurred"
      }
    };
  }
};

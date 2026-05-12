/**
 * Google Apps Script - RARS Form Automation
 *
 * Install this script on the MASTER Google Sheet (the one SheetMonkey writes to).
 *
 * When a new form submission arrives, this script:
 *   1. Reads the "Form Type" to determine which form was submitted
 *   2. Copies the row to the appropriate separate Google Sheet
 *   3. Reads the "Email" tab in that sheet to get the recipient list
 *   4. Sends email notifications to those recipients
 *   5. Generates a filled PDF for membership applications
 *   6. Posts to Discord
 *
 * A daily health check runs at noon and posts a 24-hour summary to Discord.
 *
 * SETUP: See form-automation-setup.md for full instructions.
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

// Each form type maps to a separate Google Sheet by spreadsheet ID.
// Each of these spreadsheets must have a tab called "Email" with
// email addresses in column A, starting at row 2 (row 1 is the header).
var FORM_ROUTING = {
  // "Form Type" value must match the hidden Form Type field in each HTML form exactly.
  "Contact": {
    spreadsheetId: "",  // TODO: Paste the spreadsheet ID from the Contact Us Google Sheet URL
    subject: "RARS Website: New Contact Form Submission",
    generatePdf: false
  },

  "Membership Application": {
    spreadsheetId: "",  // TODO: Paste the spreadsheet ID from the Membership Google Sheet URL
    subject: "RARS Website: New Membership Application",
    generatePdf: true
  },

  "Elmer Request": {
    spreadsheetId: "",  // TODO: Paste the spreadsheet ID from the Elmer Request Google Sheet URL
    subject: "RARS Website: New Elmer Request",
    generatePdf: false
  }
};

// Google Doc template ID for membership application PDF.
// Open the template doc and copy the ID from the URL:
//   https://docs.google.com/document/d/PASTE_THIS_PART/edit
var MEMBERSHIP_TEMPLATE_ID = "";  // TODO: Paste your Google Doc template ID here

// Google Drive folder ID to save generated membership PDFs.
// Open the folder in Drive and copy the ID from the URL:
//   https://drive.google.com/drive/folders/PASTE_THIS_PART
var PDF_FOLDER_ID = "";  // TODO: Paste your Drive folder ID here

// From name shown in outgoing notification emails
var FROM_NAME = "RARS Website";

// Discord webhook URL. In your Discord server:
//   Server Settings → Integrations → Webhooks → New Webhook → Copy Webhook URL
// Leave empty to disable Discord notifications.
var DISCORD_WEBHOOK_URL = "";  // TODO: Paste your Discord webhook URL here

// Name of the tab in the master SheetMonkey sheet where submissions are written.
// Check the tab name at the bottom of your master Google Sheet.
var MASTER_SHEET_NAME = "Sheet1";  // TODO: Update if your tab is named differently

// Name of the tab in each destination sheet that holds notification email addresses.
// Each destination sheet must have a tab with this exact name.
// Row 1 = "Email" header, Row 2+ = one email address per row.
var EMAIL_TAB_NAME = "Email";  // TODO: Update if your tab is named differently

// ============================================================================
// TRIGGER SETUP - Run these functions ONCE to install triggers
// ============================================================================

function installTrigger() {
  // Remove any existing onChange triggers to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "onSheetChange") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger("onSheetChange")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onChange()
    .create();

  Logger.log("onChange trigger installed successfully.");
}

function installDailyTrigger() {
  // Remove any existing daily health check triggers to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "dailyHealthCheck") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger("dailyHealthCheck")
    .timeBased()
    .everyDays(1)
    .atHour(12)
    .create();

  Logger.log("Daily health check trigger installed (noon every day).");
}

// ============================================================================
// MAIN HANDLER
// ============================================================================

function onSheetChange(e) {
  // Acquire a lock to prevent concurrent execution.
  // When we set "Notified" = TRUE, it triggers another onChange event.
  // Without a lock, the second run could overlap and send duplicates.
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log("Could not acquire lock. Another instance is running. Exiting.");
    return;
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);

    if (!masterSheet) {
      Logger.log("Master sheet '" + MASTER_SHEET_NAME + "' not found.");
      return;
    }

    var lastRow = masterSheet.getLastRow();
    if (lastRow < 2) return;

    // Read headers
    var lastCol = masterSheet.getLastColumn();
    var headers = masterSheet.getRange(1, 1, 1, lastCol).getValues()[0];

    // Find or create the "Notified" column
    var notifiedCol = headers.indexOf("Notified") + 1; // 1-indexed
    if (notifiedCol === 0) {
      notifiedCol = lastCol + 1;
      masterSheet.getRange(1, notifiedCol).setValue("Notified").setFontWeight("bold");
      headers.push("Notified");
      SpreadsheetApp.flush(); // Ensure the column is written before continuing
    }

    // Scan ALL rows for any that haven't been notified yet (not just the last row).
    // This catches submissions that were missed or added in quick succession.
    for (var row = 2; row <= lastRow; row++) {
      var notifiedValue = masterSheet.getRange(row, notifiedCol).getValue().toString().trim().toUpperCase();
      if (notifiedValue === "TRUE" || notifiedValue === "YES") {
        continue; // Already processed
      }

      // Mark as notified FIRST to prevent duplicates from concurrent triggers
      masterSheet.getRange(row, notifiedCol).setValue("TRUE");
      SpreadsheetApp.flush();

      var rowData = masterSheet.getRange(row, 1, 1, lastCol).getValues()[0];

      // Build a key-value object from the row
      var submission = {};
      for (var i = 0; i < headers.length; i++) {
        var header = headers[i].toString().trim();
        var value = rowData[i] !== undefined ? rowData[i].toString().trim() : "";
        if (header && header !== "Notified") {
          submission[header] = value;
        }
      }

      // Skip if this row looks empty
      if (!submission["Name"] && !submission["Email"]) continue;

      // Determine the form type
      var formType = submission["Form Type"] || "";
      var routing = FORM_ROUTING[formType];

      if (!routing) {
        Logger.log("Row " + row + ": Unknown form type '" + formType + "'. Skipping.");
        sendDiscordError(formType, submission, row);
        continue;
      }

      // --- Step 1: Copy row to the destination spreadsheet ---
      var destLastRow = copyToDestinationSheet(routing.spreadsheetId, headers, rowData);

      // --- Step 2: Send Discord notification (always fires, regardless of email config) ---
      sendDiscordNotification(formType, submission);

      // --- Step 3: Read recipients from the "Email" tab in the destination sheet ---
      var recipients = getRecipientsFromSheet(routing.spreadsheetId);

      if (recipients.length === 0) {
        Logger.log("Row " + row + ": No recipients in Email tab for " + formType + ". Skipping email.");
        markDestinationNotified(routing.spreadsheetId, destLastRow);
        continue;
      }

      // --- Step 4: Generate PDF (membership only) ---
      var pdfBlob = null;
      if (routing.generatePdf && MEMBERSHIP_TEMPLATE_ID) {
        pdfBlob = generateMembershipPdf(submission);
      }

      // --- Step 5: Send email notifications ---
      var sheetUrl = "https://docs.google.com/spreadsheets/d/" + routing.spreadsheetId;
      sendEmailNotifications(recipients, routing.subject, formType, submission, pdfBlob, sheetUrl);

      // --- Step 6: Mark destination sheet as notified ---
      markDestinationNotified(routing.spreadsheetId, destLastRow);

      Logger.log("Processed row " + row + ": " + formType + " from " + (submission["Name"] || "unknown") +
                 " → notified: " + recipients.join(", "));
    }

  } catch (error) {
    Logger.log("Error in onSheetChange: " + error.toString());
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// DAILY HEALTH CHECK
// ============================================================================

/**
 * Runs at noon every day. Posts a Discord message confirming the script is
 * running and summarizing form submissions from the last 24 hours.
 */
function dailyHealthCheck() {
  if (!DISCORD_WEBHOOK_URL) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  var masterUrl = ss.getUrl();

  var summary = { total: 0, byType: {} };
  var recentNames = [];

  if (masterSheet && masterSheet.getLastRow() >= 2) {
    var lastCol = masterSheet.getLastColumn();
    var headers = masterSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var data = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, lastCol).getValues();

    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - 1); // 24 hours ago

    var submittedCol = headers.indexOf("Submitted");
    var nameCol = headers.indexOf("Name");
    var formTypeCol = headers.indexOf("Form Type");

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var submittedRaw = submittedCol >= 0 ? row[submittedCol] : null;
      var submittedDate = submittedRaw ? new Date(submittedRaw) : null;

      if (!submittedDate || submittedDate < cutoff) continue;

      summary.total++;
      var formType = formTypeCol >= 0 ? row[formTypeCol].toString().trim() : "Unknown";
      summary.byType[formType] = (summary.byType[formType] || 0) + 1;

      var name = nameCol >= 0 ? row[nameCol].toString().trim() : "";
      if (name) recentNames.push(name + " (" + formType + ")");
    }
  }

  var typeLines = [];
  var typeEmojis = {
    "Contact": "✉️",
    "Membership Application": "📋",
    "Elmer Request": "🎓"
  };
  for (var type in summary.byType) {
    var emoji = typeEmojis[type] || "📄";
    typeLines.push(emoji + " " + type + ": " + summary.byType[type]);
  }
  if (typeLines.length === 0) typeLines.push("No submissions in the last 24 hours.");

  var description = "**Last 24 hours:** " + summary.total + " submission" + (summary.total !== 1 ? "s" : "") + "\n" +
                    typeLines.join("\n");

  if (recentNames.length > 0) {
    description += "\n\n**Submitters:** " + recentNames.join(", ");
  }

  var payload = {
    embeds: [{
      title: "🟢 RARS Website — Daily Health Check",
      description: description,
      color: 3066993, // green
      fields: [{
        name: "Master Sheet",
        value: "[View Submissions](" + masterUrl + ")",
        inline: true
      }],
      footer: { text: "www.rowanars.net" },
      timestamp: new Date().toISOString()
    }]
  };

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(DISCORD_WEBHOOK_URL, options);
  Logger.log("Daily health check posted to Discord. Response: " + response.getResponseCode());
}

// ============================================================================
// SHEET ROUTING
// ============================================================================

/**
 * Copies a submission row to the first tab of the destination spreadsheet.
 * Creates a header row if the sheet is empty.
 * Returns the row number of the newly added row.
 */
function copyToDestinationSheet(spreadsheetId, headers, rowData) {
  var destSS = SpreadsheetApp.openById(spreadsheetId);
  var destSheet = destSS.getSheets()[0]; // First tab (the data tab, not "Email")

  // If the sheet is empty, add headers first
  if (destSheet.getLastRow() === 0) {
    // Include "Notified" column in headers
    var destHeaders = headers.slice();
    if (destHeaders.indexOf("Notified") === -1) {
      destHeaders.push("Notified");
    }
    destSheet.appendRow(destHeaders);
    destSheet.getRange(1, 1, 1, destHeaders.length).setFontWeight("bold");
  }

  destSheet.appendRow(rowData);
  var newRow = destSheet.getLastRow();
  Logger.log("Copied submission to spreadsheet: " + spreadsheetId + " (row " + newRow + ")");
  return newRow;
}

/**
 * Marks a row as notified in the destination spreadsheet.
 * Finds the "Notified" column by header name.
 */
function markDestinationNotified(spreadsheetId, row) {
  var destSS = SpreadsheetApp.openById(spreadsheetId);
  var destSheet = destSS.getSheets()[0];
  var headers = destSheet.getRange(1, 1, 1, destSheet.getLastColumn()).getValues()[0];

  var notifiedCol = headers.indexOf("Notified") + 1;
  if (notifiedCol === 0) {
    // Add the column if it doesn't exist
    notifiedCol = destSheet.getLastColumn() + 1;
    destSheet.getRange(1, notifiedCol).setValue("Notified").setFontWeight("bold");
  }

  destSheet.getRange(row, notifiedCol).setValue("TRUE");
}

// ============================================================================
// RECIPIENT LOOKUP
// ============================================================================

/**
 * Reads email addresses from the "Email" tab in a destination spreadsheet.
 * Expects email addresses in column A, rows 2 through the last row.
 * Returns an array of non-empty email strings.
 */
function getRecipientsFromSheet(spreadsheetId) {
  var destSS = SpreadsheetApp.openById(spreadsheetId);
  var emailSheet = destSS.getSheetByName(EMAIL_TAB_NAME);

  if (!emailSheet) {
    Logger.log("No '" + EMAIL_TAB_NAME + "' tab found in spreadsheet: " + spreadsheetId);
    return [];
  }

  var lastRow = emailSheet.getLastRow();
  if (lastRow < 2) return []; // Only header or empty

  // Read column A from row 2 to the last row
  var emailRange = emailSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var recipients = [];

  for (var i = 0; i < emailRange.length; i++) {
    var email = emailRange[i][0].toString().trim();
    if (email && email.indexOf("@") > -1) {
      recipients.push(email);
    }
  }

  return recipients;
}

// ============================================================================
// PDF GENERATION (Membership Applications)
// ============================================================================

/**
 * Fills a Google Doc template with submission data and converts to PDF.
 * Returns the PDF as a Blob (for email attachment).
 * Also saves a copy to the configured Google Drive folder.
 *
 * Template placeholders: {{Name}}, {{Callsign}}, {{Email}}, etc.
 */
function generateMembershipPdf(submission) {
  try {
    var templateFile = DriveApp.getFileById(MEMBERSHIP_TEMPLATE_ID);
    var callsign = submission["Callsign"] || (submission["Name"] || "Unknown").replace(/[^a-zA-Z0-9]/g, "_");
    var dateStamp = getDayMonthYear(); // DD-MM-YYYY
    var copyName = callsign + "_" + dateStamp + "_Application";
    var copyFile = templateFile.makeCopy(copyName);
    var copyDoc = DocumentApp.openById(copyFile.getId());
    var body = copyDoc.getBody();

    // Replace all placeholders with submission data
    var placeholders = {
      "{{Name}}": submission["Name"] || "",
      "{{Callsign}}": submission["Callsign"] || "N/A",
      "{{License Class}}": submission["License Class"] || "N/A",
      "{{ARRL Member}}": submission["ARRL Member"] || "N/A",
      "{{Email}}": submission["Email"] || "",
      "{{Street Address}}": submission["Street Address"] || "N/A",
      "{{City State Zip}}": submission["City State Zip"] || "N/A",
      "{{Home Phone}}": submission["Home Phone"] || "N/A",
      "{{Cell Phone}}": submission["Cell Phone"] || "N/A",
      "{{Radio History}}": submission["Radio History"] || "",
      "{{Operating Habits}}": submission["Operating Habits"] || "N/A",
      "{{Frequencies and Modes}}": submission["Frequencies and Modes"] || "N/A",
      "{{Emergency Capabilities}}": submission["Emergency Capabilities"] || "N/A",
      "{{Goals}}": submission["Goals"] || "N/A",
      "{{Planned Activities}}": submission["Planned Activities"] || "N/A",
      "{{Date}}": submission["Submitted"] || new Date().toLocaleDateString()
    };

    for (var placeholder in placeholders) {
      body.replaceText(escapeRegex(placeholder), placeholders[placeholder]);
    }

    copyDoc.saveAndClose();

    // Convert the filled doc to PDF
    var pdfBlob = DriveApp.getFileById(copyFile.getId()).getAs("application/pdf");
    pdfBlob.setName(copyName + ".pdf");

    // Save PDF to Drive folder
    if (PDF_FOLDER_ID) {
      var folder = DriveApp.getFolderById(PDF_FOLDER_ID);
      folder.createFile(pdfBlob);
      Logger.log("PDF saved to Drive: " + copyName + ".pdf");
    }

    // Delete the temporary Google Doc copy
    DriveApp.getFileById(copyFile.getId()).setTrashed(true);

    return pdfBlob;

  } catch (error) {
    Logger.log("Error generating PDF: " + error.toString());
    return null;
  }
}

// ============================================================================
// EMAIL NOTIFICATIONS
// ============================================================================

/**
 * Sends formatted email notifications to the recipients list.
 * Attaches PDF if one was generated.
 */
function sendEmailNotifications(recipients, subject, formType, submission, pdfBlob, sheetUrl) {
  var body = buildPlainTextBody(formType, submission, sheetUrl);
  var htmlBody = buildHtmlBody(formType, submission, sheetUrl);

  for (var j = 0; j < recipients.length; j++) {
    var recipient = recipients[j];

    var emailOptions = {
      to: recipient,
      subject: subject,
      body: body,
      htmlBody: htmlBody,
      name: FROM_NAME
    };

    if (pdfBlob) {
      emailOptions.attachments = [pdfBlob];
    }

    MailApp.sendEmail(emailOptions);
  }

  Logger.log("Email sent for " + formType + " to " + recipients.join(", "));
}

function buildPlainTextBody(formType, submission, sheetUrl) {
  var lines = [];
  lines.push("New " + formType + " submission from the RARS website:");
  lines.push("");
  lines.push("----------------------------------------");

  for (var key in submission) {
    if (submission[key] && key !== "Form Type") {
      lines.push(key + ": " + submission[key]);
    }
  }

  lines.push("----------------------------------------");
  lines.push("");

  if (formType === "Membership Application") {
    lines.push("A filled membership application PDF is attached to this email.");
    lines.push("");
  }

  lines.push("View all submissions: " + sheetUrl);
  lines.push("");
  lines.push("This is an automated notification from www.rowanars.net");

  return lines.join("\n");
}

function buildHtmlBody(formType, submission, sheetUrl) {
  var html = [];
  html.push("<div style='font-family: Arial, sans-serif; max-width: 600px;'>");
  html.push("<h2 style='color: #1A478A; border-bottom: 2px solid #E87A1E; padding-bottom: 8px;'>");
  html.push("New " + formType + "</h2>");

  if (formType === "Membership Application") {
    html.push("<p style='background: #f0f7ff; padding: 10px; border-left: 4px solid #1A478A; margin: 16px 0;'>");
    html.push("A filled membership application PDF is attached to this email.");
    html.push("</p>");
  }

  html.push("<table style='width: 100%; border-collapse: collapse; margin-top: 16px;'>");

  for (var key in submission) {
    if (submission[key] && key !== "Form Type") {
      html.push("<tr>");
      html.push("<td style='padding: 8px 12px; border-bottom: 1px solid #eee; font-weight: bold; vertical-align: top; width: 35%; color: #333;'>");
      html.push(escapeHtml(key));
      html.push("</td>");
      html.push("<td style='padding: 8px 12px; border-bottom: 1px solid #eee; color: #555;'>");
      html.push(escapeHtml(submission[key]).replace(/\n/g, "<br>"));
      html.push("</td>");
      html.push("</tr>");
    }
  }

  html.push("</table>");
  html.push("<p style='margin-top: 16px;'>");
  html.push("<a href='" + sheetUrl + "' style='display: inline-block; padding: 8px 16px; background: #1A478A; color: #fff; text-decoration: none; border-radius: 4px;'>View All Submissions</a>");
  html.push("</p>");
  html.push("<p style='margin-top: 20px; font-size: 12px; color: #999;'>");
  html.push("Automated notification from <a href='https://www.rowanars.net'>www.rowanars.net</a>");
  html.push("</p>");
  html.push("</div>");

  return html.join("");
}

// ============================================================================
// DISCORD NOTIFICATIONS
// ============================================================================

function sendDiscordNotification(formType, submission) {
  if (!DISCORD_WEBHOOK_URL) return;

  var fields = [];
  for (var key in submission) {
    if (submission[key] && key !== "Form Type" && key !== "Submitted") {
      fields.push({
        name: key,
        value: submission[key].substring(0, 1024),
        inline: submission[key].length < 50
      });
    }
  }

  var typeColors = {
    "Contact": 3447003,            // blue
    "Membership Application": 15844367, // gold
    "Elmer Request": 10181046      // purple
  };

  var payload = {
    embeds: [{
      title: "New " + formType,
      color: typeColors[formType] || 1722250,
      fields: fields,
      footer: { text: "www.rowanars.net" },
      timestamp: new Date().toISOString()
    }]
  };

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(DISCORD_WEBHOOK_URL, options);
  Logger.log("Discord notification sent. Response: " + response.getResponseCode());
}

function sendDiscordError(formType, submission, row) {
  if (!DISCORD_WEBHOOK_URL) return;

  var fields = [];
  for (var key in submission) {
    if (submission[key]) {
      fields.push({
        name: key,
        value: submission[key].substring(0, 1024),
        inline: submission[key].length < 50
      });
    }
  }

  var unknownType = formType || "(empty)";

  var payload = {
    embeds: [{
      title: "⚠️ Unrecognized Form Submission — Action Required",
      description: "A submission arrived with Form Type **\"" + unknownType + "\"** which does not match any known form type (Contact, Membership Application, Elmer Request).\n\n**This submission was NOT saved to any destination sheet.** It is only in the master sheet (row " + row + "). Please review and file it manually.",
      color: 15158332, // red
      fields: fields,
      footer: { text: "www.rowanars.net" },
      timestamp: new Date().toISOString()
    }]
  };

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(DISCORD_WEBHOOK_URL, options);
  Logger.log("Discord error notification sent. Response: " + response.getResponseCode());
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

function escapeHtml(text) {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function getDateStamp() {
  var d = new Date();
  var month = ("0" + (d.getMonth() + 1)).slice(-2);
  var day = ("0" + d.getDate()).slice(-2);
  return d.getFullYear() + "-" + month + "-" + day;
}

function getDayMonthYear() {
  var d = new Date();
  var day = ("0" + d.getDate()).slice(-2);
  var month = ("0" + (d.getMonth() + 1)).slice(-2);
  return day + "-" + month + "-" + d.getFullYear();
}

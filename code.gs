// ---- Contains certificate template & google drive folder ----
const TEMPLATE_ID = "1m5_tQqCefX5XPLnrVOmmpmsOlHOMmlcF-dAGgkVynZE";
const DEST_FOLDER_ID = "1SMDCK6JaSrQXe6PyME1Mozs7yTzrvfV1";
const TEMPLATE_FOLDER_ID = "1UMH19RP5V6IH_-f4Z7GiJvJuXpfJvSla";

//---- Add custom menu ----
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Certificate Generation")
    .addItem("Generate & Email PDFs", "generateAndSendEmails")
    .addItem("Open Sidebar", "openSidebar")
    .addToUi();
}

function openSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar").setTitle(
    "Certificate Generator"
  );
  SpreadsheetApp.getUi().showSidebar(html);
}

//---Save Email Templates----
function saveEmailTemplate(name, subject, body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("EmailTemplates");

  if (!sheet) {
    sheet = ss.insertSheet("EmailTemplates");
    sheet.appendRow(["Template Name", "Subject", "Body"]);
  }

  const data = sheet.getDataRange().getValues();

  // Check if template with the same name exists, update if yes
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name) {
      sheet.getRange(i + 1, 2).setValue(subject);
      sheet.getRange(i + 1, 3).setValue(body);
      return;
    }
  }

  // Otherwise, append new template row
  sheet.appendRow([name, subject, body]);
}

//---- MISSING FUNCTIONS - These were referenced in your HTML but not defined ----
function getSavedEmailTemplateList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("EmailTemplates");

  if (!sheet) {
    return [];
  }

  const data = sheet.getDataRange().getValues();
  const templates = [];

  // Skip header row
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      // Only if template name exists
      templates.push({
        name: data[i][0],
        subject: data[i][1] || "",
        body: data[i][2] || "",
      });
    }
  }

  return templates;
}

function getEmailTemplateByName(templateName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("EmailTemplates");

  if (!sheet) {
    return null;
  }

  const data = sheet.getDataRange().getValues();

  // Skip header row
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === templateName) {
      return {
        name: data[i][0],
        subject: data[i][1] || "",
        body: data[i][2] || "",
      };
    }
  }

  return null;
}

//----Gets available templates-----
function getAvailableTemplates() {
  try {
    const folder = DriveApp.getFolderById(TEMPLATE_FOLDER_ID);
    const files = folder.getFilesByType(MimeType.GOOGLE_DOCS);
    const templates = [];

    while (files.hasNext()) {
      const file = files.next();
      templates.push({
        id: file.getId(),
        name: file.getName(),
      });
    }

    return templates;
  } catch (error) {
    Logger.log("Error getting templates: " + error.toString());
    return [];
  }
}

//----Email Templates-----
function generateAndSendEmailsFromTemplateWithEmail(
  templateId,
  emailSubject,
  emailBody
) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const nameCol = headers.indexOf("Name");
  const emailCol = headers.indexOf("Email");
  const courseCol = headers.indexOf("Course");
  const dateCol = headers.indexOf("Date");
  const statusCol = headers.indexOf("Status");
  const pdfCol = headers.indexOf("PDF Link");

  if (
    [nameCol, emailCol, courseCol, dateCol, statusCol, pdfCol].some(
      (i) => i === -1
    )
  ) {
    throw new Error(
      "One or more required columns are missing: Name, Email, Course, Date, Status, PDF Link"
    );
  }

  const folder = DriveApp.getFolderById(DEST_FOLDER_ID);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[statusCol];

    if (status === "Sent") continue;
    if (!row[emailCol] || !row[emailCol].includes("@")) {
      sheet.getRange(i + 1, statusCol + 1).setValue("Invalid Email");
      continue;
    }

    try {
      const copy = DriveApp.getFileById(templateId).makeCopy(
        `${row[nameCol]} - Certificate`
      );
      const doc = DocumentApp.openById(copy.getId());
      const body = doc.getBody();

      // Replace placeholders (case insensitive, trim spaces)
      ["Name", "Course", "Date"].forEach((field) => {
        const colIndex = headers.indexOf(field);
        if (colIndex > -1) {
          body.replaceText(
            new RegExp(`{{\\s*${field}\\s*}}`, "gi"),
            row[colIndex]
          );
        }
      });

      doc.saveAndClose();

      // Move file to target folder (replaces parents)
      DriveApp.getFileById(copy.getId()).moveTo(folder);

      // Create PDF in folder
      const pdfBlob = copy
        .getAs(MimeType.PDF)
        .setName(`${row[nameCol]} - Certificate.pdf`);
      const pdfFile = folder.createFile(pdfBlob);

      // Update link formula in sheet
      const fileId = pdfFile.getId();
      const formula = `=HYPERLINK("https://drive.google.com/file/d/${fileId}/view", "View PDF")`;
      sheet.getRange(i + 1, pdfCol + 1).setFormula(formula);

      // Replace placeholders in email
      const personalizedBody = fillPlaceholders(emailBody, row, headers);
      const personalizedSubject = fillPlaceholders(emailSubject, row, headers);

      // Send email
      GmailApp.sendEmail(row[emailCol], personalizedSubject, personalizedBody, {
        attachments: [pdfBlob],
        name: "Training Team",
      });

      sheet.getRange(i + 1, statusCol + 1).setValue("Sent");

      // Optional: Sleep 500 ms to avoid quota issues
      Utilities.sleep(500);
    } catch (e) {
      Logger.log(`Error on row ${i + 1} (${row[emailCol]}): ${e.toString()}`);
      sheet.getRange(i + 1, statusCol + 1).setValue("Error");
    }
  }
}

// Helper to fill in placeholders like {{Name}}, {{Course}}, etc.
function fillPlaceholders(template, row, headers) {
  let filled = template;
  headers.forEach((header, index) => {
    const placeholder = new RegExp(`{{\\s*${header}\\s*}}`, "gi");
    filled = filled.replace(placeholder, row[index] || "");
  });
  return filled;
}

//attachment_downloader.js code till 1:00 pm 03-07-2025


/**
 * @fileoverview This script processes Gmail attachments from specific labels,
 * saves them to corresponding Google Drive folders, and logs file details
 * to a Google Sheet. It includes functionality for cancellation, progress updates,
 * and special handling for a 'buffer' sheet to track original and changed filenames.
 *
 * Updated to work with dynamic folder structure:
 * - Uses only main company folder paths (aligned with Reroute.js)
 * - Dynamically navigates through FY → Accruals → Buffer/Bills and Invoices → Month → Inflow/Outflow
 * - Automatically detects financial year and month from file date
 *
 * NEW FUNCTIONALITY:
 * - When 'Status' in buffer sheet is 'Delete', deletes file from Inflow/Outflow folders and associated log entries from Main, Inflow, Outflow sheets.
 * - When 'Status' in buffer sheet is 'Active', copies file from buffer to Inflow/Outflow folders and re-logs entries in Main, Inflow, Outflow sheets.
 */

// Global variable to track cancellation status
var CANCELLATION_TOKEN = null;

// Add this line below your other global variables:
var isScriptEdit = false;

// --- IMPORTANT: CONFIGURE YOUR DRIVE FOLDER IDs HERE ---
// Aligned with Reroute.js - only main company folders needed
// Using different variable name to avoid conflict with Reroute.js
var ATTACHMENT_COMPANY_FOLDER_MAP = {
  "analogy": "160pN2zDCb9UQbwIXqgggdTjLUrFM2cM3", // Main Analogy Folder
  "humane": "1E6ijhWhdYykymN0MEUINd9jETmdM2sAt"   // Main Humane Folder
};

// --------------------------------------------------------

// --- Sheet Header Definitions ---
const BUFFER_SHEET_HEADERS = ['OriginalFileName', 'ChangedFilename', 'Invoice ID', 'Drive File ID', 'Gmail Message ID', 'Reason', 'Status'];
const MAIN_SHEET_HEADERS = [
  'File Name', 'File ID', 'File URL',
  'Date Created (Drive)', 'Last Updated (Drive)', 'Size (bytes)', 'Mime Type',
  'Email Subject', 'Gmail Message ID', 'invoice status'
];

/**
 * Gets the financial year from a date (aligned with Reroute.js)
 * @param {Date} date - The date to get financial year for
 * @returns {string} Financial year in format "FY-XX-XX"
 */
function calculateFinancialYear(date) {
  const year = date.getFullYear();
  const month = date.getMonth() + 1; // getMonth() returns 0-11

  // Financial year starts from April (month 4)
  if (month >= 4) {
    // Current year to next year
    return `FY-${year.toString().slice(-2)}-${(year + 1).toString().slice(-2)}`;
  } else {
    // Previous year to current year
    return `FY-${(year - 1).toString().slice(-2)}-${year.toString().slice(-2)}`;
  }
}

/**
 * Gets the month name from a date (aligned with Reroute.js)
 * @param {Date} date - The date to get month for
 * @returns {string} Month name (January, February, etc.)
 */
function getMonthFromDate(date) {
  const months = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  return months[date.getMonth()];
}

/**
 * Get existing folder or create new one (aligned with Reroute.js)
 * @param {GoogleAppsScript.Drive.DriveFolder} parentFolder - Parent folder
 * @param {string} folderName - Name of the folder to get or create
 * @returns {GoogleAppsScript.Drive.DriveFolder} The folder object
 */
function getOrCreateFolder(parentFolder, folderName) {
  try {
    const folders = parentFolder.getFoldersByName(folderName);

    if (folders.hasNext()) {
      return folders.next(); // Returns existing folder
    } else {
      return parentFolder.createFolder(folderName); // Creates new folder
    }
  } catch (error) {
    console.error('Error creating/getting folder:', folderName, error);
    throw error;
  }
}

/**
 * Create folder structure: Company/FY-XX-XX/Accruals/Buffer
 * @param {string} companyName - Company name (analogy/humane)
 * @param {string} financialYear - Financial year (FY-XX-XX)
 * @returns {GoogleAppsScript.Drive.DriveFolder} Buffer folder
 */
function createBufferFolderStructure(companyName, financialYear) {
  try {
    const companyFolder = DriveApp.getFolderById(ATTACHMENT_COMPANY_FOLDER_MAP[companyName]);
    const financialYearFolder = getOrCreateFolder(companyFolder, financialYear);
    const accrualsFolder = getOrCreateFolder(financialYearFolder, "Accruals");
    const bufferFolder = getOrCreateFolder(accrualsFolder, "Buffer");
    return bufferFolder;
  } catch (error) {
    console.error('Error creating buffer folder structure:', error);
    throw error;
  }
}

/**
 * Create folder structure: Company/FY-XX-XX/Accruals/Bills and Invoices/Month/FlowType
 * @param {string} companyName - Company name (analogy/humane)
 * @param {string} financialYear - Financial year (FY-XX-XX)
 * @param {string} month - Month name (January, February, etc.)
 * @param {string} flowType - Flow type (inflow/outflow)
 * @returns {GoogleAppsScript.Drive.DriveFolder} Flow folder
 */
function createFlowFolderStructure(companyName, financialYear, month, flowType) {
  try {
    const companyFolder = DriveApp.getFolderById(ATTACHMENT_COMPANY_FOLDER_MAP[companyName]);
    const financialYearFolder = getOrCreateFolder(companyFolder, financialYear);
    const accrualsFolder = getOrCreateFolder(financialYearFolder, "Accruals");
    const billsFolder = getOrCreateFolder(accrualsFolder, "Bills and Invoices");
    const monthFolder = getOrCreateFolder(billsFolder, month);
    const flowFolderName = flowType.charAt(0).toUpperCase() + flowType.slice(1); // Capitalize first letter
    const flowFolder = getOrCreateFolder(monthFolder, flowFolderName);
    return flowFolder;
  } catch (error) {
    console.error('Error creating flow folder structure:', error);
    throw error;
  }
}


/**
 * Helper to determine the target inflow/outflow folder based on file date and invoice status.
 * @param {string} companyName - The company name.
 * @param {Date} fileDate - The date of the file creation.
 * @param {string} invoiceStatus - 'inflow' or 'outflow'.
 * @returns {GoogleAppsScript.Drive.DriveFolder} The target flow folder.
 */
function findOrCreateFlowFolder(companyName, fileDate, invoiceStatus) {
  if (invoiceStatus !== "inflow" && invoiceStatus !== "outflow") {
    throw new Error("Invalid invoice status for flow folder creation: " + invoiceStatus);
  }
  const financialYear = calculateFinancialYear(fileDate);
  const month = getMonthFromDate(fileDate);
  return createFlowFolderStructure(companyName, financialYear, month, invoiceStatus);
}


/**
 * Sets the cancellation token to stop processing
 * @param {string} token - Unique token for the current process
 */
function setCancellationToken(token) {
  CANCELLATION_TOKEN = token;
  Logger.log(`Cancellation token set: ${token}`);
}

/**
 * Checks if the current process should be cancelled
 * @param {string} currentToken - Token for the current process
 * @returns {boolean} True if process should be cancelled
 */
function shouldCancel(currentToken) {
  const cancelled = CANCELLATION_TOKEN === currentToken;
  if (cancelled) {
    Logger.log(`Process cancelled with token: ${currentToken}`);
  }
  return cancelled;
}

/**
 * Clears the cancellation token
 */
function clearCancellationToken() {
  CANCELLATION_TOKEN = null;
  Logger.log('Cancellation token cleared');
}

/**
 * Retrieves all Gmail labels and formats them for the dropdown.
 * @returns {Array} An array of objects, each with 'text' and 'value' for the dropdown.
 */
function getGmailLabels() {
  var labels = GmailApp.getUserLabels();
  var labelData = [];
  labels.forEach(function(label) {
    labelData.push({ text: label.getName(), value: label.getName() });
  });
  return labelData;
}

/**
 * Gets processed IDs from a specific column of the log sheet to avoid duplicates.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} logSheet The sheet containing processed file logs.
 * @param {number} columnIndex The 0-indexed column number to read IDs from.
 * @returns {Set<string>} A set of processed IDs.
 */
function getProcessedLogEntryIds(logSheet, columnIndex) {
  const data = logSheet.getDataRange().getValues();
  const processedIds = new Set();
  // Skip header row (index 0)
  for (let i = 1; i < data.length; i++) {
    const id = data[i][columnIndex];
    if (id) {
      processedIds.add(id);
    }
  }
  return processedIds;
}

/**
 * Extracts invoice ID from filename (third underscore-separated part, or 'NA').
 * @param {string} filename The name of the file.
 * @returns {string} The extracted invoice ID or 'NA'.
 */
function extractInvoiceIdFromFilename(filename) {
  const parts = filename.split('_');
  if (parts.length >= 3) {
    const id = parts[2].split('.')[0];
    return id && id.trim() ? id : 'NA';
  }
  return 'NA';
}

/**
 * Generates a new filename based on extracted data and original filename.
 * This is a helper function to ensure consistent filename generation logic.
 * @param {Object} extractedData An object potentially containing date, vendorName, invoiceNumber, amount.
 * @param {string} originalFileName The original filename including extension.
 * @returns {string} The new, sanitized filename.
 */
function generateNewFilename(extractedData, originalFileName) {
  const lastDotIndex = originalFileName.lastIndexOf('.');
  const extension = lastDotIndex > 0 ? originalFileName.substring(lastDotIndex) : '';
  const date = String(extractedData.date || "YYYY-MM-DD");
  const vendor = String(extractedData.vendorName || "UnknownVendor");
  const invoice = String(extractedData.invoiceNumber || "INV-Unknown");
  const amount = String(extractedData.amount || "0.00");

  const sanitizedDate = date.replace(/[^0-9\-]/g, '').trim() || "YYYY-MM-DD";
  // Allow spaces in vendor name, replace underscores with hyphens
  const sanitizedVendor = vendor.replace(/[_]/g, '-')
    .replace(/[/\\:*?"<>|]/g, '').replace(/\s+/g, ' ').trim() || "UnknownVendor";
  const sanitizedInvoice = invoice.replace(/[_]/g, '-')
    .replace(/[/\\:*?"<>|]/g, '').replace(/\s+/g, '').trim() || "INV-Unknown";
  const sanitizedAmount = amount.replace(/[_]/g, '')
    .replace(/[/\\:*?"<>|]/g, '').trim() || "0.00";

  return `${sanitizedDate}_${sanitizedVendor}_${sanitizedInvoice}_${sanitizedAmount}${extension}`;
}

/**
 * Retrieves all existing 'ChangedFilename' values from the buffer sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} bufferSheet The buffer sheet.
 * @returns {Set<string>} A Set of existing changed filenames.
 */
function getExistingChangedFilenames(bufferSheet) {
  const existingChangedFilenames = new Set();
  const data = bufferSheet.getDataRange().getValues();
  // ChangedFilename is in column 2 (index 1)
  for (let i = 1; i < data.length; i++) { // Start from 1 to skip header row
    const changedFilename = data[i][1];
    if (changedFilename) {
      existingChangedFilenames.add(changedFilename);
    }
  }
  return existingChangedFilenames;
}

/**
 * Sets up data validation for the Status dropdown in buffer sheets.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The buffer sheet to apply validation to.
 */
function setStatusDropdownValidation(sheet) {
  const range = sheet.getRange("G:G"); // Status column is G (index 6, assuming 0-indexed column 5)
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(['Active', 'Delete'], true).build();
  range.setDataValidation(rule);
  Logger.log(`Status dropdown validation applied to ${sheet.getName()}`);
}

/**
 * Finds a Gmail message by its ID.
 * @param {string} messageId The ID of the Gmail message.
 * @returns {GoogleAppsScript.Gmail.GmailMessage|null} The Gmail message object, or null if not found.
 */
function getGmailMessageById(messageId) {
  try {
    const message = GmailApp.getMessageById(messageId);
    return message;
  } catch (e) {
    Logger.log(`Error retrieving Gmail message ID ${messageId}: ${e.toString()}`);
    return null;
  }
}

/**
 * Finds a specific attachment within a Gmail message by its name.
 * @param {GoogleAppsScript.Gmail.GmailMessage} message The Gmail message object.
 * @param {string} attachmentName The name of the attachment to find.
 * @returns {GoogleAppsScript.Gmail.GmailAttachment|null} The attachment object, or null if not found.
 */
function findAttachmentInMessage(message, attachmentName) {
  const attachments = message.getAttachments();
  for (let i = 0; i < attachments.length; i++) {
    if (attachments[i].getName() === attachmentName) {
      return attachments[i];
    }
  }
  return null;
}

/**
 * Logs a file's details to a specified log sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} logSheet The sheet to log to.
 * @param {GoogleAppsScript.Drive.File} driveFile The Drive file object.
 * @param {string} emailSubject The subject of the original email.
 * @param {string} gmailMessageId The ID of the original Gmail message.
 * @param {string} invoiceStatus The determined invoice status (inflow, outflow, unknown).
 */
function logFileToMainSheet(logSheet, driveFile, emailSubject, gmailMessageId, invoiceStatus) {
  logSheet.appendRow([
    driveFile.getName(),
    driveFile.getId(),
    driveFile.getUrl(),
    driveFile.getDateCreated(),
    driveFile.getLastUpdated(),
    driveFile.getSize(),
    driveFile.getMimeType(),
    emailSubject,
    gmailMessageId,
    invoiceStatus
  ]);
  Logger.log(`Logged file ${driveFile.getName()} (ID: ${driveFile.getId()}) to ${logSheet.getName()}`);
}

/**
 * Gets the email subject for a given Gmail Message ID.
 * @param {string} gmailMessageId The Gmail message ID.
 * @returns {string} The email subject, or 'N/A' if not found.
 */
function getEmailSubjectForMessageId(gmailMessageId) {
  try {
    const message = GmailApp.getMessageById(gmailMessageId);
    return message ? message.getSubject() : 'N/A';
  } catch (e) {
    Logger.log(`Could not get subject for Gmail Message ID ${gmailMessageId}: ${e.toString()}`);
    return 'N/A';
  }
}

/**
 * Processes attachments from Gmail labels and save them to Google Drive folders
 * @param {string} labelName - The Gmail label name (e.g., 'analogy', 'humane')
 * @param {string} processToken - Unique token for this process
 * @returns {Object} An object with status and message.
 */
function processAttachments(labelName, processToken) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  clearCancellationToken(); // Clear any existing token and set a new one for this run.

  // Example: labelName = "analogy/accruals/bills&invoices"
  const companyName = labelName.split('/')[0]; // "analogy" or "humane"
const companyFolder = DriveApp.getFolderById(ATTACHMENT_COMPANY_FOLDER_MAP[companyName]);
  const bufferLabelName = `${companyName}-buffer`;

  // Check if the company exists in our folder mapping
  if (!ATTACHMENT_COMPANY_FOLDER_MAP[companyName]) {
    return { status: 'error', message: `Error: No Drive folder configured for company '${companyName}'. Please update the script.` };
  }

  let bufferSheet;
  try {
    bufferSheet = ss.getSheetByName(bufferLabelName);
    if (!bufferSheet) {
      bufferSheet = ss.insertSheet(bufferLabelName);
      Logger.log(`Created new sheet: ${bufferLabelName}`);
    }
    // Ensure headers are correct for buffer sheet
    const firstRowRange = bufferSheet.getRange(1, 1, 1, BUFFER_SHEET_HEADERS.length);
    const currentHeaders = firstRowRange.getValues()[0];

    if (currentHeaders.join(',') !== BUFFER_SHEET_HEADERS.join(',')) {
      bufferSheet.clear();
      bufferSheet.appendRow(BUFFER_SHEET_HEADERS);
      bufferSheet.getRange(1, 1, 1, BUFFER_SHEET_HEADERS.length).setFontWeight('bold').setBackground('#E8F0FE').setBorder(true, true, true, true, true, true);
      bufferSheet.setFrozenRows(1);
      bufferSheet.setColumnWidth(1, 200); // OriginalFilename
      bufferSheet.setColumnWidth(2, 200); // ChangedFilename
      bufferSheet.setColumnWidth(3, 120); // Invoice ID
      bufferSheet.setColumnWidth(4, 180); // Drive File ID (New)
      bufferSheet.setColumnWidth(5, 200); // Gmail Message ID (New)
      bufferSheet.setColumnWidth(6, 250); // Reason
      bufferSheet.setColumnWidth(7, 80);  // Status
    }
    // Remove any extra columns if present
    if (bufferSheet.getLastColumn() > BUFFER_SHEET_HEADERS.length) {
      bufferSheet.deleteColumns(BUFFER_SHEET_HEADERS.length + 1, bufferSheet.getLastColumn() - BUFFER_SHEET_HEADERS.length);
    }
    setStatusDropdownValidation(bufferSheet); // Apply dropdown validation
  } catch (e) {
    return { status: 'error', message: `Error setting up buffer sheet for ${bufferLabelName}: ${e.toString()}` };
  }

try {
  // Get the main company folder
  const companyName = labelName.split('/')[0]; // "analogy" or "humane"
  const companyFolder = DriveApp.getFolderById(ATTACHMENT_COMPANY_FOLDER_MAP[companyName]);
  const gmailLabel = GmailApp.getUserLabelByName(labelName); // Original Gmail label (analogy/humane)

    if (!gmailLabel) {
      return { status: 'error', message: `Error: Gmail label '${labelName}' not found.` };
    }

    if (shouldCancel(processToken)) {
      return { status: 'cancelled', message: "Process cancelled by user before starting." };
    }

    const processedGmailMessageIds = getProcessedLogEntryIds(bufferSheet, 4); // Gmail Message ID is at index 4 (E column)
    const existingChangedFilenamesInCurrentBuffer = getExistingChangedFilenames(bufferSheet); // For duplicate checking in buffer sheet
    let totalNewAttachments = 0;
    let processedAttachments = 0;
    let skippedAttachments = 0;

    // First pass to count total NEW attachments for accurate progress
    const threads = gmailLabel.getThreads();
    for (let t = 0; t < threads.length; t++) {
      if (shouldCancel(processToken)) {
        return { status: 'cancelled', message: "Process cancelled during attachment counting." };
      }
      const messages = threads[t].getMessages();
      for (let m = 0; m < messages.length; m++) {
        const message = messages[m];
        if (processedGmailMessageIds.has(message.getId())) {
          skippedAttachments += message.getAttachments().filter(a => !a.isGoogleType() && !a.getName().startsWith('ATT')).length;
          continue;
        }
        totalNewAttachments += message.getAttachments().filter(a => !a.isGoogleType() && !a.getName().startsWith('ATT')).length;
      }
      Utilities.sleep(50);
    }

    if (totalNewAttachments === 0) {
      let message = `No new email attachments found for '${labelName}'.`;
      if (skippedAttachments > 0) {
        message += ` (${skippedAttachments} email attachments skipped as their messages were previously processed)`;
      }
      clearCancellationToken();
      return { status: 'success', message: message };
    }

    // Second pass to process and log NEW attachments only
    for (let t = 0; t < threads.length; t++) {
      if (shouldCancel(processToken)) {
        clearCancellationToken();
        return { status: 'cancelled', message: `Process cancelled. Processed ${processedAttachments} of ${totalNewAttachments} attachments before cancellation.` };
      }

      const messages = threads[t].getMessages();
      for (let m = 0; m < messages.length; m++) {
        const message = messages[m];
        const messageId = message.getId();

        if (processedGmailMessageIds.has(messageId)) {
          continue;
        }

        const attachments = message.getAttachments();
        let attachmentsProcessedForThisMessage = 0;

        for (let a = 0; a < attachments.length; a++) {
          if (shouldCancel(processToken)) {
            clearCancellationToken();
            return { status: 'cancelled', message: `Process cancelled. Processed ${processedAttachments} of ${totalNewAttachments} attachments before cancellation.` };
          }

          const attachment = attachments[a];
          if (!attachment.isGoogleType() && !attachment.getName().startsWith('ATT')) {
            try {
              const originalFilename = attachment.getName();
              const invoiceId = extractInvoiceIdFromFilename(originalFilename);
              const changedFilename = generateNewFilename({ invoiceNumber: invoiceId }, originalFilename);

              const isDuplicateChangedFilename = existingChangedFilenamesInCurrentBuffer.has(changedFilename);
              let driveFile = null;

              if (!isDuplicateChangedFilename) {
                // Store in Buffer/Active with the correct name
                const now = new Date();
                const month = getMonthFromDate(now);
                const financialYear = calculateFinancialYear(now);
                const bufferActiveFolder = getOrCreateBufferSubfolder(companyName, financialYear, "Active");

                // Create the file with the changedFilename
                const renamedBlob = attachment.copyBlob().setName(changedFilename);
                driveFile = bufferActiveFolder.createFile(renamedBlob);
                processedAttachments++;
                attachmentsProcessedForThisMessage++;
                existingChangedFilenamesInCurrentBuffer.add(changedFilename);

                // AI logic
                const emailSubject = message.getSubject ? message.getSubject() : '';
                const blob = driveFile.getBlob();
                const aiResult = callGeminiAPIInternal(blob, changedFilename);
                const invoiceStatus = aiResult.invoiceStatus || "unknown";

                // Log to main sheet
                let mainSheet = ss.getSheetByName(companyName);
                if (!mainSheet) {
                  mainSheet = ss.insertSheet(companyName);
                  mainSheet.appendRow(MAIN_SHEET_HEADERS);
                }
                logFileToMainSheet(mainSheet, driveFile, emailSubject, messageId, invoiceStatus);

                // If inflow/outflow, create a true copy in the inflow/outflow folder
                if (invoiceStatus === "inflow" || invoiceStatus === "outflow") {
                  const now = new Date();
                  const month = getMonthFromDate(now);
                  const financialYear = calculateFinancialYear(now);
                  const flowFolder = createFlowFolderStructure(companyName, financialYear, month, invoiceStatus);

                  let flowFile = null;
                  try {
                    flowFile = driveFile.makeCopy(changedFilename, flowFolder);
                    Logger.log(`Copied file ${changedFilename} to ${invoiceStatus} folder for ${financialYear}/${month}.`);
                  } catch (copyErr) {
                    Logger.log(`Error copying file ${changedFilename} to ${invoiceStatus} folder: ${copyErr}`);
                  }

                  // Log to inflow/outflow sheet (log the copy, not the buffer file)
                  const flowSheetName = `${companyName}-${invoiceStatus}`;
                  let flowSheet = ss.getSheetByName(flowSheetName);
                  if (!flowSheet) {
                    flowSheet = ss.insertSheet(flowSheetName);
                    flowSheet.appendRow(MAIN_SHEET_HEADERS);
                  }
                  if (flowFile) {
                    logFileToMainSheet(flowSheet, flowFile, emailSubject, messageId, invoiceStatus);
                  }
                }
              } else {
                // If duplicate, don't upload again, but log in buffer
                Logger.log(`Skipping upload: Duplicate changed filename '${changedFilename}' already exists in buffer folder for '${labelName}'.`);
                skippedAttachments++;
                attachmentsProcessedForThisMessage++;
              }

              // Append to buffer sheet
              const rowData = [
                originalFilename,
                changedFilename,
                invoiceId,
                driveFile ? driveFile.getId() : '', // Drive File ID
                messageId,                               // Gmail Message ID
                '',                                      // Reason (blank for new entries)
                'Active'                                 // Default Status (Active)
              ];
              bufferSheet.appendRow(rowData);
              const newRowIndex = bufferSheet.getLastRow();

              // If it's a duplicate, color the entire row yellow
              if (isDuplicateChangedFilename) {
                bufferSheet.getRange(newRowIndex, 1, 1, BUFFER_SHEET_HEADERS.length).setBackground('#FFFF00'); // Yellow
              }

              try {
                google.script.run.withSuccessHandler(function(){})
                  .updateProgress(processedAttachments, totalNewAttachments, labelName, processToken);
              } catch (progressError) {
                Logger.log("Error sending progress update: " + progressError.toString());
              }
              Utilities.sleep(100);

              Logger.log("Created file: " + (driveFile ? driveFile.getName() : originalFilename) + " in buffer folder for: " + companyName);

            } catch (fileError) {
              Logger.log(`Error processing attachment '${attachment.getName()}': ${fileError.toString()}`);
            }
          }
        }
        if (attachmentsProcessedForThisMessage > 0 || attachments.length === 0) {
          processedGmailMessageIds.add(messageId);
        }
      }
    }

    clearCancellationToken();
    let resultMessage = `Completed processing for '${labelName}'. `;
    resultMessage += `Processed ${processedAttachments} new email attachments uploaded to company folder. `;
    resultMessage += `Skipped ${skippedAttachments} attachments (already processed or duplicate in buffer).`;

    // After processing into company folder, automatically trigger the buffer processing for this company
    processBufferFilesAndLog(labelName);

    return { status: 'success', message: resultMessage };

  } catch (e) {
    clearCancellationToken();
    return { status: 'error', message: "An error occurred: " + e.toString() };
  }
}

/**
 * Processes files from a specific company's buffer sheet: renames in Drive,
 * logs to the main sheet (sheet only, no drive storage), and copies to inflow/outflow sheets and folders.
 * This function should be called after `processAttachments` for a given company.
 * @param {string} companyName The company name (e.g., 'analogy', 'humane').
 */
function processBufferFilesAndLog(companyName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const bufferSheet = ss.getSheetByName(`${companyName}-buffer`);
  if (!bufferSheet) {
    Logger.log(`Buffer sheet '${companyName}-buffer' not found.`);
    return;
  }

  let mainSheet = ss.getSheetByName(companyName);

  // Ensure main sheet headers are correct
  if (!mainSheet) {
    mainSheet = ss.insertSheet(companyName);
  }
  const mainSheetFirstRowRange = mainSheet.getRange(1, 1, 1, MAIN_SHEET_HEADERS.length);
  if (mainSheet.getLastRow() === 0 || mainSheetFirstRowRange.getValues()[0].join(',') !== MAIN_SHEET_HEADERS.join(',')) {
    mainSheet.clear();
    mainSheet.getRange(1, 1, 1, MAIN_SHEET_HEADERS.length).setValues([MAIN_SHEET_HEADERS]);
    mainSheet.getRange(1, 1, 1, MAIN_SHEET_HEADERS.length).setFontWeight('bold').setBackground('#E8F0FE').setBorder(true, true, true, true, true, true);
    mainSheet.setFrozenRows(1);
  }
  if (mainSheet.getLastColumn() > MAIN_SHEET_HEADERS.length) {
    mainSheet.deleteColumns(MAIN_SHEET_HEADERS.length + 1, mainSheet.getLastColumn() - MAIN_SHEET_HEADERS.length);
  }

  // Prepare inflow/outflow sheets
  const inflowSheetName = `${companyName}-inflow`;
  const outflowSheetName = `${companyName}-outflow`;

  let inflowSheet = ss.getSheetByName(inflowSheetName);
  if (!inflowSheet) {
    inflowSheet = ss.insertSheet(inflowSheetName);
    inflowSheet.appendRow(MAIN_SHEET_HEADERS);
    inflowSheet.getRange(1, 1, 1, MAIN_SHEET_HEADERS.length).setFontWeight('bold').setBackground('#E8F0FE').setBorder(true, true, true, true, true, true);
    inflowSheet.setFrozenRows(1);
  }
  if (inflowSheet.getLastColumn() > MAIN_SHEET_HEADERS.length) {
    inflowSheet.deleteColumns(MAIN_SHEET_HEADERS.length + 1, inflowSheet.getLastColumn() - MAIN_SHEET_HEADERS.length);
  }

  let outflowSheet = ss.getSheetByName(outflowSheetName);
  if (!outflowSheet) {
    outflowSheet = ss.insertSheet(outflowSheetName);
    outflowSheet.appendRow(MAIN_SHEET_HEADERS);
    outflowSheet.getRange(1, 1, 1, MAIN_SHEET_HEADERS.length).setFontWeight('bold').setBackground('#E8F0FE').setBorder(true, true, true, true, true, true);
    outflowSheet.setFrozenRows(1);
  }
  if (outflowSheet.getLastColumn() > MAIN_SHEET_HEADERS.length) {
    outflowSheet.deleteColumns(MAIN_SHEET_HEADERS.length + 1, outflowSheet.getLastColumn() - MAIN_SHEET_HEADERS.length);
  }

  const bufferData = bufferSheet.getDataRange().getValues();
  const bufferRanges = bufferSheet.getDataRange();

  for (let i = 1; i < bufferData.length; i++) { // Start from 1 to skip header row
    const row = bufferData[i];
    const originalFilename = row[0];
    const changedFilename = row[1];
    let driveFileId = row[3]; // This can be updated
    const gmailMessageId = row[4];
    const status = row[6]; // Status column
    const bufferRowIndex = i + 1; // 1-indexed row number in the sheet

    // Only process 'Active' files that have a valid Drive File ID (or a placeholder 'DELETED' that needs to be re-created)
    if (status === 'Active' && (driveFileId && driveFileId !== '')) {
      try {
        let file;
        let fileNeedsCopyingFromBuffer = false; // Flag to check if we need to copy from buffer folder

        if (driveFileId === 'DELETED') {
          // File was marked deleted and now needs to be restored from buffer
          fileNeedsCopyingFromBuffer = true;
          const companyFolder = DriveApp.getFolderById(ATTACHMENT_COMPANY_FOLDER_MAP[companyName]);
          const financialYear = calculateFinancialYear(new Date()); // Use current date for initial buffer folder search
          const bufferFolder = createBufferFolderStructure(companyName, financialYear);

          const filesInFolder = bufferFolder.getFilesByName(changedFilename);
          if (filesInFolder.hasNext()) {
            file = filesInFolder.next(); // Get the file from buffer
            driveFileId = file.getId(); // Update DriveFileId with the actual ID from buffer
            bufferSheet.getRange(bufferRowIndex, 4).setValue(driveFileId); // Update the buffer sheet with the new ID
            Logger.log(`Found file in buffer folder for restoration: ${changedFilename} (${driveFileId})`);
          } else {
            Logger.log(`File '${changedFilename}' (ID: ${driveFileId}) not found in buffer folder. Cannot restore.`);
            bufferSheet.getRange(bufferRowIndex, 6).setValue('Error: File not found in buffer for restore.');
            continue; // Skip to next row
          }
        } else {
          // File should exist, verify it
          try {
            file = DriveApp.getFileById(driveFileId);
            // Verify file name consistency
            if (file.getName() !== changedFilename) {
              file.setName(changedFilename);
              Logger.log(`Renamed file ${originalFilename} to ${changedFilename} in buffer folder.`);
            }
          } catch (e) {
            Logger.log(`File with ID ${driveFileId} not found in Drive. Attempting to find in buffer for re-creation.`);
            // If file is not found in Drive, it means it was likely deleted manually or not created properly.
            // We'll treat this as if it needs to be restored from buffer if a file with changedFilename exists there.
            fileNeedsCopyingFromBuffer = true;
            const companyFolder = DriveApp.getFolderById(ATTACHMENT_COMPANY_FOLDER_MAP[companyName]);
            const financialYear = calculateFinancialYear(new Date()); // Use current date for initial buffer folder search
            const bufferFolder = createBufferFolderStructure(companyName, financialYear);
            const filesInFolder = bufferFolder.getFilesByName(changedFilename);
            if (filesInFolder.hasNext()) {
              file = filesInFolder.next(); // Get the file from buffer
              driveFileId = file.getId(); // Update DriveFileId with the actual ID from buffer
              bufferSheet.getRange(bufferRowIndex, 4).setValue(driveFileId); // Update the buffer sheet with the new ID
              Logger.log(`Found missing file in buffer folder for re-creation: ${changedFilename} (${driveFileId})`);
            } else {
              Logger.log(`File '${changedFilename}' (ID: ${driveFileId}) not found in buffer folder either. Cannot restore/re-process.`);
              bufferSheet.getRange(bufferRowIndex, 6).setValue('Error: File not found in Drive or buffer for re-process.');
              continue; // Skip to next row
            }
          }
        }

        // Get file creation date to determine financial year and month
        const fileDate = file.getDateCreated();
        const financialYear = calculateFinancialYear(fileDate);
        const month = getMonthFromDate(fileDate);

        // 1. Move file to correct buffer folder based on financial year (only if it needs to be moved)
        const bufferFolder = createBufferFolderStructure(companyName, financialYear);
        const currentParents = file.getParents();
        let isAlreadyInBufferFolder = false;
        while(currentParents.hasNext()){
          if(currentParents.next().getId() === bufferFolder.getId()){
            isAlreadyInBufferFolder = true;
            break;
          }
        }

        if (!isAlreadyInBufferFolder) {
          file.moveTo(bufferFolder);
          Logger.log(`Moved file ${changedFilename} to ${companyName} ${financialYear} buffer folder.`);
        }

        // 2. Use AI to determine inflow/outflow/unknown
        const emailSubject = getEmailSubjectForMessageId(gmailMessageId);
        const blob = file.getBlob();
        const aiResult = callGeminiAPIInternal(blob, changedFilename);
        const invoiceStatus = aiResult.invoiceStatus || "unknown";
        Logger.log(`AI invoiceStatus for file ${changedFilename}: ${invoiceStatus}`);

        // 3. Delete existing log entries from Main, Inflow, Outflow sheets before re-logging
        deleteLogEntries(mainSheet, driveFileId, gmailMessageId);
        deleteLogEntries(inflowSheet, driveFileId, gmailMessageId);
        deleteLogEntries(outflowSheet, driveFileId, gmailMessageId);


        // 4. Log to main sheet (sheet only, no drive storage)
        logFileToMainSheet(mainSheet, file, emailSubject, gmailMessageId, invoiceStatus);

        // 5. Copy to inflow or outflow if appropriate (both sheet and drive)
        if (invoiceStatus === "inflow" || invoiceStatus === "outflow") {
          const targetFlowFolder = findOrCreateFlowFolder(companyName, fileDate, invoiceStatus);

          // Check if a file with the same name already exists in the target flow folder
          let copiedFile = null;
          const existingFilesInFlow = targetFlowFolder.getFilesByName(changedFilename);
          if (existingFilesInFlow.hasNext()) {
            copiedFile = existingFilesInFlow.next();
            Logger.log(`File ${changedFilename} already exists in ${invoiceStatus} folder. Reusing existing.`);
          } else {
            copiedFile = file.makeCopy(changedFilename, targetFlowFolder);
            Logger.log(`Copied file ${changedFilename} to ${companyName} ${financialYear} ${month} ${invoiceStatus} folder.`);
          }

          const flowSheet = (invoiceStatus === "inflow") ? inflowSheet : outflowSheet;
          logFileToMainSheet(flowSheet, copiedFile, emailSubject, gmailMessageId, invoiceStatus);
        }

        // Clear any previous "Reason" or yellow background if successfully processed as Active
        bufferSheet.getRange(bufferRowIndex, 6).setValue('');
        bufferSheet.getRange(bufferRowIndex, 1, 1, BUFFER_SHEET_HEADERS.length).setBackground(null); // Remove background color


      } catch (e) {
        Logger.log(`Error processing buffer row for file ID ${driveFileId} (Row ${bufferRowIndex}): ${e.toString()}`);
        ui.alert('Error', `Could not process file "${originalFilename}" from buffer (Row ${bufferRowIndex}): ${e.message}`, ui.ButtonSet.OK);
        bufferSheet.getRange(bufferRowIndex, 6).setValue(`Error: ${e.message}`); // Log error reason
        bufferSheet.getRange(bufferRowIndex, 1, 1, BUFFER_SHEET_HEADERS.length).setBackground('#FF0000'); // Red background for error
      }
    } else if (status === 'Delete' && driveFileId && driveFileId !== 'DELETED') {
      // Handle deletion when status is set to 'Delete'
      Logger.log(`Processing 'Delete' status for file: ${changedFilename} (ID: ${driveFileId})`);
      try {
        const fileToDelete = DriveApp.getFileById(driveFileId);
        const emailSubject = getEmailSubjectForMessageId(gmailMessageId); // Get subject before deleting logs

        // Delete from Inflow/Outflow Drive folders (if it exists there)
        // Find all parents of the file and check if any are inflow/outflow
        const parents = fileToDelete.getParents();
        let isFileInFlowFolder = false;
        while(parents.hasNext()){
          const parent = parents.next();
          // Check if parent path contains "Bills and Invoices"
          // This is a heuristic, a more robust solution might involve storing parent folder IDs
          if (parent.getName().toLowerCase() === "inflow" || parent.getName().toLowerCase() === "outflow") {
            try {
              fileToDelete.setTrashed(true); // Move to trash
              Logger.log(`Trashed file ${changedFilename} (ID: ${driveFileId}) from flow folder.`);
              isFileInFlowFolder = true;
              break;
            } catch (trashError) {
              Logger.log(`Could not trash file ${driveFileId}: ${trashError.toString()}. It might already be trashed or moved.`);
            }
          }
        }
        if(!isFileInFlowFolder) {
            Logger.log(`File ${changedFilename} (ID: ${driveFileId}) was not found in an Inflow/Outflow folder or could not be trashed.`);
        }


        // Delete corresponding log entries from Main, Inflow, Outflow sheets
        deleteLogEntries(mainSheet, driveFileId, gmailMessageId);
        deleteLogEntries(inflowSheet, driveFileId, gmailMessageId);
        deleteLogEntries(outflowSheet, driveFileId, gmailMessageId);

        // Mark the Drive File ID in the buffer sheet as 'DELETED'
        bufferSheet.getRange(bufferRowIndex, 4).setValue('DELETED');
        bufferSheet.getRange(bufferRowIndex, 6).setValue('Deleted from flow/logs'); // Add reason
        bufferSheet.getRange(bufferRowIndex, 1, 1, BUFFER_SHEET_HEADERS.length).setBackground('#FFD966'); // Orange background for deleted


      } catch (e) {
        Logger.log(`Error deleting file from flow folder/logs for ID ${driveFileId} (Row ${bufferRowIndex}): ${e.toString()}`);
        ui.alert('Error', `Could not delete file "${changedFilename}" from flow folder/logs (Row ${bufferRowIndex}): ${e.message}`, ui.ButtonSet.OK);
        bufferSheet.getRange(bufferRowIndex, 6).setValue(`Deletion Error: ${e.message}`); // Log error reason
        bufferSheet.getRange(bufferRowIndex, 1, 1, BUFFER_SHEET_HEADERS.length).setBackground('#FF0000'); // Red background for error
      }
    }
  }

  Logger.log(`Completed processing of '${companyName}-buffer'.`);
}

/**
 * Deletes log entries from a given sheet based on Drive File ID or Gmail Message ID.
 * It iterates from bottom to top to avoid issues with row index changes during deletion.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to delete from.
 * @param {string} driveFileId The Drive File ID to match.
 * @param {string} gmailMessageId The Gmail Message ID to match.
 */
function deleteLogEntries(sheet, driveFileId, gmailMessageId) {
  if (!sheet) {
    Logger.log("Sheet not found for deletion. Skipping.");
    return;
  }
  const data = sheet.getDataRange().getValues();
  // Find column indexes dynamically, assuming headers are always present
  const fileIdColIndex = MAIN_SHEET_HEADERS.indexOf('File ID');
  const gmailMessageIdColIndex = MAIN_SHEET_HEADERS.indexOf('Gmail Message ID');

  if (fileIdColIndex === -1 || gmailMessageIdColIndex === -1) {
    Logger.log(`Required headers for deletion not found in sheet: ${sheet.getName()}`);
    return;
  }

  let deletedCount = 0;
  // Iterate backwards to avoid issues with row index changes during deletion
  for (let i = data.length - 1; i >= 1; i--) { // Start from last row, skip header
    const row = data[i];
    // Check if either file ID or Gmail Message ID matches
    if (row[fileIdColIndex] === driveFileId || row[gmailMessageIdColIndex] === gmailMessageId) {
      sheet.deleteRow(i + 1); // +1 because sheet rows are 1-indexed
      deletedCount++;
    }
  }
  if (deletedCount > 0) {
    Logger.log(`Deleted ${deletedCount} log entries from ${sheet.getName()} for Drive ID: ${driveFileId} or Gmail ID: ${gmailMessageId}`);
  }
}


/**
 * Duplicates log entries and files from a source label to target labels.
 * This function seems to be for initial setup/copying, not part of the active deletion/restoration.
 * It's kept here as it was in the original code, but note its separate purpose.
 * @param {string} sourceLabel - The label whose data to duplicate (e.g., 'analogy').
 * @param {Array<string>} targetLabels - The labels to copy data into (e.g., ['analogy-inflow', 'analogy-outflow']).
 * @param {Object} companyFolderMap - The mapping of label names to Drive folder IDs.
 */
function duplicateLogAndFiles(sourceLabel, targetLabels, companyFolderMap) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(sourceLabel);
  if (!sourceSheet) return;

  var data = sourceSheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log(`No data to duplicate from ${sourceLabel}.`);
    return;
  }

  var headers = data[0];
  var rows = data.slice(1);

  targetLabels.forEach(function(targetLabel) {
    var targetSheet = ss.getSheetByName(targetLabel);
    if (!targetSheet) {
      targetSheet = ss.insertSheet(targetLabel);
      targetSheet.appendRow(headers); // Add headers
      targetSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#E8F0FE').setBorder(true, true, true, true, true, true);
      targetSheet.setFrozenRows(1);
    }
    // Remove any extra columns from target sheet if present
    if (targetSheet.getLastColumn() > headers.length) {
      targetSheet.deleteColumns(headers.length + 1, targetSheet.getLastColumn() - headers.length);
    }

    // This `targetFolderId` logic is problematic.
    // The targetLabel here (e.g., 'analogy-inflow') is not a direct key in ATTACHMENT_COMPANY_FOLDER_MAP.
    // This `duplicateLogAndFiles` function needs to be re-evaluated or clarified
    // if it's meant to copy files to inflow/outflow *folders*.
    // For now, assuming it's meant to duplicate within the main company folder if no specific flow folder ID is mapped.
    var targetFolder = null;
    if (companyFolderMap[targetLabel]) {
      targetFolder = DriveApp.getFolderById(companyFolderMap[targetLabel]);
    } else if (companyFolderMap[sourceLabel]) {
      // Fallback: use the source company's main folder if no specific target folder is mapped.
      targetFolder = DriveApp.getFolderById(companyFolderMap[sourceLabel]);
      Logger.log(`Warning: No specific folder for '${targetLabel}'. Using main folder for '${sourceLabel}'.`);
    } else {
      Logger.log(`Target folder ID not found for ${targetLabel} or ${sourceLabel}. Skipping duplication for ${targetLabel}.`);
      return;
    }


    // Get existing file IDs in the target sheet to avoid re-logging
    const existingTargetFileIds = getProcessedLogEntryIds(targetSheet, 1); // File ID is at index 1

    rows.forEach(function(row) {
      try {
        var fileId = row[1]; // File ID column in main sheet
        var fileName = row[0]; // File Name column in main sheet
        var gmailMessageId = row[8]; // Gmail Message ID in main sheet
        var invoiceStatus = row[9]; // Invoice Status in main sheet

        if (!fileId || existingTargetFileIds.has(fileId)) {
          // Skip if no fileId or already logged in target sheet
          return;
        }

        var file = DriveApp.getFileById(fileId);
        var copiedFile = null;

        // Determine the correct subfolder (Inflow/Outflow) if the targetLabel suggests it
        if (targetLabel.endsWith('-inflow') || targetLabel.endsWith('-outflow')) {
          const companyName = targetLabel.split('-')[0];
          const flowType = targetLabel.split('-')[1];
          const fileDate = file.getDateCreated();
          const specificFlowFolder = createFlowFolderStructure(companyName, calculateFinancialYear(fileDate), getMonthFromDate(fileDate), flowType);
          copiedFile = file.makeCopy(fileName, specificFlowFolder);
          Logger.log(`Duplicated file ${fileName} to ${specificFlowFolder.getName()} folder.`);
        } else {
          // Default to the main company folder if not inflow/outflow specific
          copiedFile = file.makeCopy(fileName, targetFolder); // Copy with its current name
          Logger.log(`Duplicated file ${fileName} to ${targetFolder.getName()} folder.`);
        }

        var newRow = row.slice();
        newRow[1] = copiedFile.getId();       // Update with new copied file ID
        newRow[2] = copiedFile.getUrl();      // Update with new URL
        newRow[3] = copiedFile.getDateCreated(); // Update with new date created
        newRow[4] = copiedFile.getLastUpdated(); // Update with new last updated

        // Ensure invoice status is consistent with target label if applicable
        if (targetLabel.endsWith('-inflow')) {
          newRow[9] = 'inflow';
        } else if (targetLabel.endsWith('-outflow')) {
          newRow[9] = 'outflow';
        }

        targetSheet.appendRow(newRow);
        Logger.log(`Logged duplicated entry for ${fileName} in ${targetSheet.getName()}.`);

      } catch (e) {
        Logger.log(`Error duplicating log entry for file ID ${row[1]} to ${targetLabel}: ${e.toString()}`);
      }
    });
  });
  Logger.log(`Completed duplication for ${sourceLabel}.`);
}


/**
 * Placeholder for your Gemini API call.
 * This function needs to be implemented to interact with your AI model.
 * It should return an object with at least an `invoiceStatus` property.
 * @param {GoogleAppsScript.Base.Blob} fileBlob The content of the file.
 * @param {string} fileName The name of the file.
 * @returns {Object} An object containing the extracted invoice status (e.g., {invoiceStatus: "inflow"}).
 */
function callGeminiAPIInternal(fileBlob, fileName) {
  // --- IMPORTANT: REPLACE THIS WITH YOUR ACTUAL GEMINI API INTEGRATION ---
  Logger.log(`Simulating Gemini API call for ${fileName}`);

  // Simulate AI logic based on filename for demonstration
  let simulatedInvoiceStatus = "unknown";
  const lowerFileName = fileName.toLowerCase();

  if (lowerFileName.includes("invoice") && !lowerFileName.includes("payment")) {
    simulatedInvoiceStatus = "outflow"; // Assuming invoices typically represent money going out
  } else if (lowerFileName.includes("receipt") || lowerFileName.includes("deposit")) {
    simulatedInvoiceStatus = "inflow";
  } else if (lowerFileName.includes("credit")) {
    simulatedInvoiceStatus = "inflow";
  }

  // You would typically send the fileBlob content to your Gemini API here
  // const API_KEY = "YOUR_GEMINI_API_KEY";
  // const GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-pro-vision:generateContent?key=" + API_KEY;
  //
  // const imageData = Utilities.base64Encode(fileBlob.getBytes());
  //
  // const payload = {
  //   contents: [
  //     {
  //       parts: [
  //         {text: "Analyze this document and determine if it represents an 'inflow' (money coming in) or 'outflow' (money going out) for a business. If unsure, return 'unknown'. Respond with a single word: inflow, outflow, or unknown."},
  //         {inlineData: {mimeType: fileBlob.getContentType(), data: imageData}}
  //       ]
  //     }
  //   ]
  // };
  //
  // const options = {
  //   method: 'post',
  //   contentType: 'application/json',
  //   payload: JSON.stringify(payload),
  //   muteHttpExceptions: true
  // };
  //
  // try {
  //   const response = UrlFetchApp.fetch(GEMINI_URL, options);
  //   const jsonResponse = JSON.parse(response.getContentText());
  //   Logger.log("Gemini API Response: " + JSON.stringify(jsonResponse));
  //
  //   if (jsonResponse.candidates && jsonResponse.candidates[0] && jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts) {
  //     const aiText = jsonResponse.candidates[0].content.parts[0].text.toLowerCase().trim();
  //     if (aiText.includes("inflow")) {
  //       simulatedInvoiceStatus = "inflow";
  //     } else if (aiText.includes("outflow")) {
  //       simulatedInvoiceStatus = "outflow";
  //     } else {
  //       simulatedInvoiceStatus = "unknown";
  //     }
  //   }
  // } catch (e) {
  //   Logger.log("Error calling Gemini API: " + e.toString());
  //   simulatedInvoiceStatus = "unknown_api_error";
  // }

  return { invoiceStatus: simulatedInvoiceStatus };
}

// --- Add these helper functions near the top of your script, after global variables ---

function setScriptEditFlag(value) {
  PropertiesService.getScriptProperties().setProperty('isScriptEdit', value ? 'true' : 'false');
}

function getScriptEditFlag() {
  return PropertiesService.getScriptProperties().getProperty('isScriptEdit') === 'true';
}

// --- New onEdit Trigger Function ---
/**
 * Handles changes in the spreadsheet, specifically for the 'Status' column in buffer sheets.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The event object.
 */
function onEdit(e) {
  // Prevent multiple triggers from the same edit
  if (getScriptEditFlag()) {
    setScriptEditFlag(false);
    return;
  }

  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  // Only handle Status column changes in buffer sheets
  if (
    sheetName.endsWith('-buffer') &&
    range.getColumn() === BUFFER_SHEET_HEADERS.indexOf('Status') + 1 &&
    range.getNumRows() === 1 && // Only single cell edits
    range.getNumColumns() === 1
  ) {
    const companyName = sheetName.replace('-buffer', '');
    const editedRow = range.getRow();
    const newStatus = e.value;
    const oldStatus = e.oldValue;
    
    // Skip header row and prevent recursive calls
    if (editedRow === 1 || newStatus === oldStatus) return;

    const rowData = sheet.getRange(editedRow, 1, 1, BUFFER_SHEET_HEADERS.length).getValues()[0];
    const changedFilename = rowData[1];
    let driveFileId = rowData[3];
    const gmailMessageId = rowData[4];
    const ui = SpreadsheetApp.getUi();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheetByName(companyName);
    const inflowSheet = ss.getSheetByName(`${companyName}-inflow`);
    const outflowSheet = ss.getSheetByName(`${companyName}-outflow`);

    // --- DELETE ACTION ---
    if (newStatus === 'Delete' && oldStatus !== 'Delete') {
      if (!shouldShowPrompt()) {
        return; // Prevent multiple prompts
      }
      
      const response = ui.prompt(
        'Reason for Deletion',
        `Please provide a reason for deleting "${changedFilename}".`,
        ui.ButtonSet.OK_CANCEL
      );
      
      if (response.getSelectedButton() === ui.Button.OK) {
        const reasonText = response.getResponseText().trim();
        
        // Only proceed if reason is provided
        if (!reasonText) {
          setScriptEditFlag(true);
          sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('Status') + 1).setValue(oldStatus || 'Active');
          ui.alert('Error', 'Reason is required. Status reverted.', ui.ButtonSet.OK);
          return;
        }
        
        setScriptEditFlag(true);
        sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('Reason') + 1).setValue(reasonText);

        if (driveFileId && driveFileId !== 'DELETED') {
          try {
            const file = DriveApp.getFileById(driveFileId);
            const fileDate = file.getDateCreated();
            const financialYear = calculateFinancialYear(fileDate);
            const bufferActiveFolder = getOrCreateBufferSubfolder(companyName, financialYear, "Active");
            const bufferDeletedFolder = getOrCreateBufferSubfolder(companyName, financialYear, "Deleted");

            // Move from Buffer/Active to Buffer/Deleted
            moveFileWithDriveApi(file.getId(), bufferDeletedFolder.getId(), bufferActiveFolder.getId());

            // Move files from inflow/outflow folders to Buffer/Deleted folder
            const aiResult = callGeminiAPIInternal(file.getBlob(), changedFilename);
            const invoiceStatus = aiResult.invoiceStatus || "unknown";
            if (invoiceStatus === "inflow" || invoiceStatus === "outflow") {
              const month = getMonthFromDate(fileDate);
              const flowFolder = createFlowFolderStructure(companyName, financialYear, month, invoiceStatus);
              const filesInFlow = flowFolder.getFilesByName(changedFilename);
              while (filesInFlow.hasNext()) {
                const flowFile = filesInFlow.next();
                // Move the file from inflow/outflow to Buffer/Deleted folder
                moveFileWithDriveApi(flowFile.getId(), bufferDeletedFolder.getId(), flowFolder.getId());
                Logger.log(`Moved file ${changedFilename} from ${invoiceStatus} folder to Buffer/Deleted folder.`);
              }
            }

            // Remove log entries
            deleteLogEntries(mainSheet, driveFileId, gmailMessageId);
            deleteLogEntries(inflowSheet, driveFileId, gmailMessageId);
            deleteLogEntries(outflowSheet, driveFileId, gmailMessageId);

            // Mark the Drive File ID in the buffer sheet as 'DELETED'
            setScriptEditFlag(true);
            sheet.getRange(editedRow, 4).setValue('DELETED');
            sheet.getRange(editedRow, 6).setValue(reasonText);
            sheet.getRange(editedRow, 1, 1, BUFFER_SHEET_HEADERS.length).setBackground('#FFD966');
          } catch (e) {
            Logger.log(`Error moving file to buffer: ${e.toString()}`);
            ui.alert('Error', `Failed to move file "${changedFilename}" to buffer: ${e.message}`, ui.ButtonSet.OK);
            setScriptEditFlag(true);
            sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('Status') + 1).setValue(oldStatus || 'Active');
          }
        }
      } else {
        setScriptEditFlag(true);
        sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('Status') + 1).setValue(oldStatus || 'Active');
        return;
      }
    }

    // --- ACTIVATE ACTION ---
    else if (newStatus === 'Active' && oldStatus !== 'Active') {
      if (!shouldShowPrompt()) {
        return; // Prevent multiple prompts
      }
      
      const response = ui.prompt(
        'Reason for Activation',
        `Please provide a reason for activating "${changedFilename}".`,
        ui.ButtonSet.OK_CANCEL
      );
      
      if (response.getSelectedButton() === ui.Button.OK) {
        const reasonText = response.getResponseText().trim();
        
        // Only proceed if reason is provided
        if (!reasonText) {
          setScriptEditFlag(true);
          sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('Status') + 1).setValue(oldStatus || 'Delete');
          ui.alert('Error', 'Reason is required. Status reverted.', ui.ButtonSet.OK);
          return;
        }
        
        setScriptEditFlag(true);
        sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('Reason') + 1).setValue(reasonText);

        if (driveFileId && driveFileId !== 'DELETED') {
          try {
            // First check if the file is in Buffer/Deleted folder
            let file = null;
            let fileDate = null;
            let financialYear = null;
            let bufferActiveFolder = null;
            let bufferDeletedFolder = null;
            
            // Try to find file in buffer deleted folder first
            const currentFinancialYear = calculateFinancialYear(new Date());
            bufferDeletedFolder = getOrCreateBufferSubfolder(companyName, currentFinancialYear, "Deleted");
            const filesInDeleted = bufferDeletedFolder.getFilesByName(changedFilename);
            
            if (filesInDeleted.hasNext()) {
              // File found in deleted folder
              file = filesInDeleted.next();
              fileDate = file.getDateCreated();
              financialYear = calculateFinancialYear(fileDate);
              bufferActiveFolder = getOrCreateBufferSubfolder(companyName, financialYear, "Active");
              bufferDeletedFolder = getOrCreateBufferSubfolder(companyName, financialYear, "Deleted");
              
              // Move from Buffer/Deleted to Buffer/Active
              moveFileWithDriveApi(file.getId(), bufferActiveFolder.getId(), bufferDeletedFolder.getId());

              // Determine if the file should be moved to inflow/outflow
              const aiResult = callGeminiAPIInternal(file.getBlob(), changedFilename);
              const invoiceStatus = aiResult.invoiceStatus || "unknown";
              const emailSubject = getEmailSubjectForMessageId(gmailMessageId);
              
              if (invoiceStatus === "inflow" || invoiceStatus === "outflow") {
                const month = getMonthFromDate(fileDate);
                const flowFolder = createFlowFolderStructure(companyName, financialYear, month, invoiceStatus);

                // Copy from Buffer/Active to the appropriate inflow/outflow folder
                const copiedFile = file.makeCopy(changedFilename, flowFolder);
                Logger.log(`Copied file ${changedFilename} from Buffer/Active to ${invoiceStatus} folder.`);

                // Log in main and inflow/outflow sheets
                logFileToMainSheet(mainSheet, copiedFile, emailSubject, gmailMessageId, invoiceStatus);
                const flowSheet = (invoiceStatus === "inflow") ? inflowSheet : outflowSheet;
                logFileToMainSheet(flowSheet, copiedFile, emailSubject, gmailMessageId, invoiceStatus);
              }
              
              // Update the driveFileId in the buffer sheet
              setScriptEditFlag(true);
              sheet.getRange(editedRow, 4).setValue(file.getId());
            } else {
              // File not in deleted folder, try to get by ID
              try {
                file = DriveApp.getFileById(driveFileId);
                fileDate = file.getDateCreated();
                financialYear = calculateFinancialYear(fileDate);
                bufferActiveFolder = getOrCreateBufferSubfolder(companyName, financialYear, "Active");
                
                // Use AI to determine inflow/outflow/unknown
                const aiResult = callGeminiAPIInternal(file.getBlob(), changedFilename);
                const invoiceStatus = aiResult.invoiceStatus || "unknown";
                const emailSubject = getEmailSubjectForMessageId(gmailMessageId);

                // Copy to inflow/outflow folders if applicable
                if (invoiceStatus === "inflow" || invoiceStatus === "outflow") {
                  const month = getMonthFromDate(fileDate);
                  const flowFolder = createFlowFolderStructure(companyName, financialYear, month, invoiceStatus);
                  
                  const copiedFile = file.makeCopy(changedFilename, flowFolder);
                  Logger.log(`Copied file ${changedFilename} from Buffer/Active to ${invoiceStatus} folder.`);

                  // Log in main and inflow/outflow sheets
                  logFileToMainSheet(mainSheet, copiedFile, emailSubject, gmailMessageId, invoiceStatus);
                  const flowSheet = (invoiceStatus === "inflow") ? inflowSheet : outflowSheet;
                  logFileToMainSheet(flowSheet, copiedFile, emailSubject, gmailMessageId, invoiceStatus);
                } else {
                  // Log in main sheet only
                  logFileToMainSheet(mainSheet, file, emailSubject, gmailMessageId, invoiceStatus);
                }
              } catch (fileNotFoundError) {
                Logger.log(`File with ID ${driveFileId} not found. Cannot activate.`);
                ui.alert('Error', `File not found. Cannot activate "${changedFilename}".`, ui.ButtonSet.OK);
                setScriptEditFlag(true);
                sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('Status') + 1).setValue(oldStatus || 'Delete');
                return;
              }
            }

            // Clear any previous "Reason" or yellow background if successfully processed as Active
            setScriptEditFlag(true);
            sheet.getRange(editedRow, 6).setValue(reasonText);
            sheet.getRange(editedRow, 1, 1, BUFFER_SHEET_HEADERS.length).setBackground(null);
            
            // Update Drive File ID in buffer sheet if it was 'DELETED'
            if (sheet.getRange(editedRow, 4).getValue() === 'DELETED') {
              setScriptEditFlag(true);
              sheet.getRange(editedRow, 4).setValue(file.getId());
            }
          } catch (e) {
            Logger.log(`Error activating file: ${e.toString()}`);
            ui.alert('Error', `Failed to activate file "${changedFilename}": ${e.message}`, ui.ButtonSet.OK);
            setScriptEditFlag(true);
            sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('Status') + 1).setValue(oldStatus || 'Delete');
            return;
          }
        }
      } else {
        setScriptEditFlag(true);
        sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('Status') + 1).setValue(oldStatus || 'Delete');
        return;
      }
    }
  }
}


function shouldShowPrompt() {
  var props = PropertiesService.getScriptProperties();
  var lastPrompt = Number(props.getProperty('lastPromptTime') || 0);
  var now = Date.now();
  if (now - lastPrompt < 3000) { // 3 seconds to prevent rapid fire
    return false;
  }
  props.setProperty('lastPromptTime', String(now));
  return true;
}

/**
 * Moves a file from one folder to another using the Drive API (UrlFetchApp).
 * @param {string} fileId - The ID of the file to move.
 * @param {string} addParentId - The folder ID to add as parent.
 * @param {string} removeParentId - The folder ID to remove as parent.
 */
function moveFileWithDriveApi(fileId, addParentId, removeParentId) {
  var token = ScriptApp.getOAuthToken();
  var url = 'https://www.googleapis.com/drive/v3/files/' + fileId + '?addParents=' + addParentId + '&removeParents=' + removeParentId + '&fields=id,parents';
  var options = {
    method: 'patch',
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + token
    }
  };
  var response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    throw new Error('Failed to move file: ' + response.getContentText());
  }
  return JSON.parse(response.getContentText());
}

/**
 * Get or create the buffer subfolder for a company and financial year.
 * @param {string} companyName - The name of the company (e.g., 'analogy', 'humane').
 * @param {string} financialYear - The financial year (e.g., 'FY-23-24').
 * @param {string} subfolderName - The name of the subfolder to create or get (e.g., 'Active' or 'Deleted').
 * @returns {GoogleAppsScript.Drive.DriveFolder} The buffer subfolder.
 */
function getOrCreateBufferSubfolder(companyName, financialYear, subfolderName) {
  const bufferFolder = createBufferFolderStructure(companyName, financialYear);
  return getOrCreateFolder(bufferFolder, subfolderName); // "Active" or "Deleted"
}
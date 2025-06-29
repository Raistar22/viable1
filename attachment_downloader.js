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
 */

// Global variable to track cancellation status
var CANCELLATION_TOKEN = null;

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
 * Process attachments from Gmail labels and save them to Google Drive folders
 * @param {string} labelName - The Gmail label name (e.g., 'analogy', 'humane')
 * @param {string} processToken - Unique token for this process
 * @returns {Object} An object with status and message.
 */
function processAttachments(labelName, processToken) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  clearCancellationToken(); // Clear any existing token and set a new one for this run.

  // Determine the target buffer sheet and folder based on the labelName
  const bufferLabelName = `${labelName}-buffer`;
  
  // Check if the company exists in our folder mapping
  if (!ATTACHMENT_COMPANY_FOLDER_MAP[labelName]) {
    return { status: 'error', message: `Error: No Drive folder configured for company '${labelName}'. Please update the script.` };
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
    const companyFolder = DriveApp.getFolderById(ATTACHMENT_COMPANY_FOLDER_MAP[labelName]);
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
                // For now, save to the main company folder
                // The processBufferFilesAndLog function will move it to the correct buffer folder later
                driveFile = companyFolder.createFile(attachment);
                processedAttachments++;
                attachmentsProcessedForThisMessage++;
                existingChangedFilenamesInCurrentBuffer.add(changedFilename); // Add to set for current run's duplicate checking
              } else {
                // If it's a duplicate by changed filename, we don't upload to Drive again, but still log the entry in buffer
                Logger.log(`Skipping upload: Duplicate changed filename '${changedFilename}' already exists in buffer folder for '${labelName}'.`);
                skippedAttachments++;
                attachmentsProcessedForThisMessage++; // Still counted as handled
              }

              // Append to buffer sheet
              const rowData = [
                originalFilename,
                changedFilename,
                invoiceId,
                driveFile ? driveFile.getId() : '', // Drive File ID
                messageId,                          // Gmail Message ID
                '',                                 // Reason (blank for new entries)
                'Active'                            // Default Status (Active)
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

              Logger.log("Created file: " + driveFile.getName() + " in folder: " + companyFolder.getName() + " (" + companyFolder.getId() + ")");

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

  for (let i = 1; i < bufferData.length; i++) { // Start from 1 to skip header row
    const row = bufferData[i];
    const originalFilename = row[0];
    const changedFilename = row[1];
    const driveFileId = row[3];
    const gmailMessageId = row[4];
    const status = row[6]; // Status column

    // Only process 'Active' files that have a valid Drive File ID and are not already marked 'DELETED'
    if (status === 'Active' && driveFileId && driveFileId !== 'DELETED') {
      try {
        const file = DriveApp.getFileById(driveFileId);
        
        // Get file creation date to determine financial year and month
        const fileDate = file.getDateCreated();
        const financialYear = calculateFinancialYear(fileDate);
        const month = getMonthFromDate(fileDate);

        // 1. Move file to correct buffer folder based on financial year
        const bufferFolder = createBufferFolderStructure(companyName, financialYear);
        
        // Move file to buffer folder if it's not already there
        const currentParent = file.getParents().next();
        if (currentParent.getId() !== bufferFolder.getId()) {
          file.moveTo(bufferFolder);
          Logger.log(`Moved file ${changedFilename} to ${companyName} ${financialYear} buffer folder.`);
        }

        // 2. Rename the file in the buffer folder if necessary
        if (file.getName() !== changedFilename) {
          file.setName(changedFilename);
          Logger.log(`Renamed file ${originalFilename} to ${changedFilename} in buffer folder.`);
        }

        // 3. Use AI to determine inflow/outflow/unknown
        const emailSubject = getEmailSubjectForMessageId(gmailMessageId);
        const blob = file.getBlob();
        const aiResult = callGeminiAPIInternal(blob, changedFilename);
        const invoiceStatus = aiResult.invoiceStatus || "unknown";
        Logger.log(`AI invoiceStatus for file ${changedFilename}: ${invoiceStatus}`);

        // 4. Log to main sheet, including invoice status (sheet only, no drive storage)
        logFileToMainSheet(mainSheet, file, emailSubject, gmailMessageId, invoiceStatus);

        // 5. Copy to inflow or outflow if appropriate (both sheet and drive)
        if (invoiceStatus === "inflow") {
          const inflowFolder = createFlowFolderStructure(companyName, financialYear, month, "inflow");
          const copiedFile = file.makeCopy(changedFilename, inflowFolder);
          logFileToMainSheet(inflowSheet, copiedFile, emailSubject, gmailMessageId, "inflow");
          Logger.log(`Copied file ${changedFilename} to ${companyName} ${financialYear} ${month} inflow folder.`);
        } else if (invoiceStatus === "outflow") {
          const outflowFolder = createFlowFolderStructure(companyName, financialYear, month, "outflow");
          const copiedFile = file.makeCopy(changedFilename, outflowFolder);
          logFileToMainSheet(outflowSheet, copiedFile, emailSubject, gmailMessageId, "outflow");
          Logger.log(`Copied file ${changedFilename} to ${companyName} ${financialYear} ${month} outflow folder.`);
        }
        // If "unknown", do nothing extra (already logged in main sheet)

      } catch (e) {
        Logger.log(`Error processing buffer row for file ID ${driveFileId}: ${e.toString()}`);
        ui.alert('Error', `Could not process file "${originalFilename}" from buffer: ${e.message}`, ui.ButtonSet.OK);
      }
    }
  }

  Logger.log(`Completed processing of '${companyName}-buffer'.`);
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
 * Duplicates log entries and files from a source label to target labels.
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

    var targetFolderId = companyFolderMap[targetLabel];
    if (!targetFolderId) {
      Logger.log(`Target folder ID not found for ${targetLabel}. Skipping duplication.`);
      return;
    }
    var targetFolder = DriveApp.getFolderById(targetFolderId);

    // Get existing file IDs in the target sheet to avoid re-logging
    const existingTargetFileIds = getProcessedLogEntryIds(targetSheet, 1); // File ID is at index 1

    rows.forEach(function(row) {
      try {
        var fileId = row[1]; // File ID column in main sheet
        var fileName = row[0]; // File Name column in main sheet
        var gmailMessageId = row[8]; // Gmail Message ID in main sheet

        if (!fileId || existingTargetFileIds.has(fileId)) {
          // Skip if no fileId or already logged in target sheet
          return;
        }

        var file = DriveApp.getFileById(fileId);
        var copiedFile = file.makeCopy(fileName, targetFolder); // Copy with its current name

        var newRow = row.slice();
        newRow[1] = copiedFile.getId();       // Update with new copied file ID
        newRow[2] = copiedFile.getUrl();      // Update with new URL
        newRow[3] = copiedFile.getDateCreated();
        newRow[4] = copiedFile.getLastUpdated();
        newRow[5] = copiedFile.getSize();
        newRow[6] = copiedFile.getMimeType();

        targetSheet.appendRow(newRow);
        existingTargetFileIds.add(copiedFile.getId()); // Add the new ID to the set
        Logger.log(`Duplicated file ${fileName} (ID: ${copiedFile.getId()}) to ${targetLabel} sheet and folder.`);
      } catch (e) {
        Logger.log(`Error duplicating file ${fileName} for ${targetLabel}: ${e.toString()}`);
      }
    });
  });
}

/**
 * An installable onEdit trigger function to handle changes in buffer sheets.
 * This is the core logic for deletion and restoration.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The event object.
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  const row = range.getRow();
  const col = range.getColumn();
  const ui = SpreadsheetApp.getUi();

  // Check if it's one of the buffer sheets and the Status column (column G, index 6)
  if ((sheetName === 'analogy-buffer' || sheetName === 'humane-buffer') && col === 7 && row > 1) { // col 7 is 'G'
    const status = e.value;
    const oldStatus = e.oldValue;

    // Get row data (assuming headers are 1st row)
    const rowData = sheet.getRange(row, 1, 1, BUFFER_SHEET_HEADERS.length).getValues()[0];
    const originalFilename = rowData[0];
    const changedFilename = rowData[1];
    let driveFileId = rowData[3]; // Drive File ID column (D)
    const gmailMessageId = rowData[4]; // Gmail Message ID column (E)
    let reason = rowData[5];      // Reason column (F)

    const companyName = sheetName.split('-')[0]; // 'analogy' or 'humane'

    if (status === 'Delete') {
      // Prompt for reason if not already provided or if changing from Active to Delete
      if (!reason || oldStatus !== 'Delete') {
        const response = ui.prompt(
          'Reason for Deletion',
          `Please provide a reason for deleting '${originalFilename}'.`,
          ui.ButtonSet.OK_CANCEL
        );
        if (response.getSelectedButton() === ui.Button.OK) {
          reason = response.getResponseText();
          sheet.getRange(row, 6).setValue(reason); // Update Reason column
        } else {
          // User cancelled, revert status
          sheet.getRange(row, col).setValue(oldStatus || 'Active'); // Revert to old status or 'Active'
          ui.alert('Deletion Cancelled', 'The deletion was cancelled. Status reverted.', ui.ButtonSet.OK);
          return; // Stop execution
        }
      }

      // Delete files from inflow/outflow folders by finding them by name
      // (since they are copies, they have different file IDs but same names)
      try {
        const inflowFolderId = ATTACHMENT_COMPANY_FOLDER_MAP[`${companyName}-inflow`];
        const outflowFolderId = ATTACHMENT_COMPANY_FOLDER_MAP[`${companyName}-outflow`];
        
        // Delete from inflow folder
        if (inflowFolderId) {
          const inflowFolder = DriveApp.getFolderById(inflowFolderId);
          const inflowFiles = inflowFolder.getFilesByName(changedFilename);
          while (inflowFiles.hasNext()) {
            const inflowFile = inflowFiles.next();
            inflowFolder.removeFile(inflowFile);
            Logger.log(`Removed file '${changedFilename}' from inflow folder.`);
          }
        }
        
        // Delete from outflow folder
        if (outflowFolderId) {
          const outflowFolder = DriveApp.getFolderById(outflowFolderId);
          const outflowFiles = outflowFolder.getFilesByName(changedFilename);
          while (outflowFiles.hasNext()) {
            const outflowFile = outflowFiles.next();
            outflowFolder.removeFile(outflowFile);
            Logger.log(`Removed file '${changedFilename}' from outflow folder.`);
          }
        }
        
        Logger.log(`Deleted copies of '${originalFilename}' from inflow/outflow folders. Original file remains in buffer folder.`);
      } catch (fileError) {
        Logger.log(`Error deleting files from inflow/outflow folders: ${fileError.toString()}`);
        ui.alert('Drive Deletion Error', `Could not delete copies of '${originalFilename}' from inflow/outflow folders: ${fileError.message}`, ui.ButtonSet.OK);
        // Don't halt, proceed to delete from sheets even if Drive failed
      }

      // Mark the Drive File ID in the buffer sheet as 'DELETED'
      sheet.getRange(row, 4).setValue('DELETED'); // Column D (index 3)

      // Delete corresponding rows from main and inflow/outflow sheets
      deleteRowFromConnectedSheets(companyName, driveFileId);

    } else if (status === 'Active' && oldStatus === 'Delete') {
      ui.alert('Restoration Initiated', `Attempting to restore '${originalFilename}' from buffer folder... This may take a moment.`, ui.ButtonSet.OK);

      try {
        // Try to find the file in the buffer folder by filename
        const bufferFolder = DriveApp.getFolderById(ATTACHMENT_COMPANY_FOLDER_MAP[`${companyName}-buffer`]);
        const files = bufferFolder.getFilesByName(changedFilename);
        if (!files.hasNext()) {
          ui.alert('Restoration Error', `File '${changedFilename}' not found in buffer folder. Cannot restore.`, ui.ButtonSet.OK);
          sheet.getRange(row, col).setValue('Delete'); // Revert status
          return;
        }

        const restoredFile = files.next();

        // Update buffer sheet with the original Drive File ID (restore it)
        sheet.getRange(row, 4).setValue(restoredFile.getId()); // Update Drive File ID (Column D)
        sheet.getRange(row, 6).setValue(''); // Clear Reason for restoration

        // Re-log the file to the main sheet (sheet only, no drive storage)
        const mainSheet = sheet.getParent().getSheetByName(companyName);
        logFileToMainSheet(mainSheet, restoredFile, getEmailSubjectForMessageId(gmailMessageId), gmailMessageId, "unknown");

        // Use AI to determine inflow/outflow/unknown
        const blob = restoredFile.getBlob();
        const aiResult = callGeminiAPIInternal(blob, changedFilename);
        const invoiceStatus = aiResult.invoiceStatus || "unknown";
        Logger.log(`AI invoiceStatus for file ${changedFilename}: ${invoiceStatus}`);

        // Copy to inflow or outflow if appropriate (both sheet and drive)
        if (invoiceStatus === "inflow") {
          const inflowSheet = sheet.getParent().getSheetByName(`${companyName}-inflow`);
          const inflowFolder = DriveApp.getFolderById(ATTACHMENT_COMPANY_FOLDER_MAP[`${companyName}-inflow`]);
          const copiedFile = restoredFile.makeCopy(changedFilename, inflowFolder);
          logFileToMainSheet(inflowSheet, copiedFile, getEmailSubjectForMessageId(gmailMessageId), gmailMessageId, "inflow");
        } else if (invoiceStatus === "outflow") {
          const outflowSheet = sheet.getParent().getSheetByName(`${companyName}-outflow`);
          const outflowFolder = DriveApp.getFolderById(ATTACHMENT_COMPANY_FOLDER_MAP[`${companyName}-outflow`]);
          const copiedFile = restoredFile.makeCopy(changedFilename, outflowFolder);
          logFileToMainSheet(outflowSheet, copiedFile, getEmailSubjectForMessageId(gmailMessageId), gmailMessageId, "outflow");
        }

        ui.alert('Restoration Complete', `'${originalFilename}' has been successfully restored from buffer folder and re-logged.`, ui.ButtonSet.OK);

      } catch (restoreError) {
        Logger.log(`Error during restoration of '${originalFilename}': ${restoreError.toString()}`);
        ui.alert('Restoration Failed', `An error occurred during restoration of '${originalFilename}': ${restoreError.message}. Status reverted to 'Delete'.`, ui.ButtonSet.OK);
        sheet.getRange(row, col).setValue('Delete'); // Revert status if restoration fails
      }
    }
  }
}

/**
 * Deletes a row from main and inflow/outflow sheets based on Drive File ID.
 * @param {string} companyName The company name (analogy or humane).
 * @param {string} driveFileId The Drive File ID of the row to delete.
 */
function deleteRowFromConnectedSheets(companyName, driveFileId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToUpdate = [
    ss.getSheetByName(companyName),
    ss.getSheetByName(`${companyName}-inflow`),
    ss.getSheetByName(`${companyName}-outflow`)
  ];

  sheetsToUpdate.forEach(targetSheet => {
    if (!targetSheet) {
      Logger.log(`Sheet '${targetSheet}' not found for deletion.`);
      return;
    }
    const data = targetSheet.getDataRange().getValues();
    let rowsToDelete = [];
    for (let i = 1; i < data.length; i++) { // Start from 1 to skip header
      if (data[i][1] === driveFileId) { // File ID is column B (index 1)
        rowsToDelete.push(i + 1); // Store 1-indexed row number
      }
    }
    // Delete rows from bottom up to avoid index issues
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      targetSheet.deleteRow(rowsToDelete[i]);
      Logger.log(`Deleted row with File ID ${driveFileId} from ${targetSheet.getName()}.`);
    }
  });
}

/**
 * Function to handle cancellation requests from the UI
 * @param {string} processToken The token of the process to cancel
 * @returns {Object} Cancellation status
 */
function cancelProcess(processToken) {
  if (!processToken) {
    return { status: 'error', message: 'No process token provided for cancellation.' };
  }
  setCancellationToken(processToken);
  return { status: 'success', message: 'Cancellation request sent. Process will stop at the next safe checkpoint.' };
}

/**
 * Updated progress function that accepts the process token for cancellation checks
 * @param {number} current
 * @param {number} total
 * @param {string} labelName
 * @param {string} processToken
 */
function updateProgress(current, total, labelName, processToken) {
  // This function doesn't actually do anything on the server-side,
  // it just serves as a target for google.script.run from the client
  // to allow the client to update its own UI.
  Logger.log(`Progress for ${labelName}: ${current}/${total} (Token: ${processToken})`);
}

function ensureSheetHeaders(sheet, headers) {
  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (firstRow.join(',') !== headers.join(',')) {
    sheet.clear();
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#E8F0FE')
      .setBorder(true, true, true, true, true, true);
    sheet.setFrozenRows(1);
  }
  if (sheet.getLastColumn() > headers.length) {
    sheet.deleteColumns(headers.length + 1, sheet.getLastColumn() - headers.length);
  }
}

function testDrivePermission() {
  // This will try to list the first file in your Drive (safe, just for permission)
  var files = DriveApp.getFiles();
  if (files.hasNext()) {
    var file = files.next();
    Logger.log("Found file: " + file.getName());
  } else {
    Logger.log("No files found in Drive.");
  }
}

/**
 * Test function to trigger Drive permissions
 */
function testDrivePermissions() {
  try {
    // Try to access Drive to trigger permission request
    const testFolder = DriveApp.getRootFolder();
    Logger.log("Drive access successful: " + testFolder.getName());
    return "Drive permissions are working correctly!";
  } catch (error) {
    Logger.log("Drive permission error: " + error.toString());
    return "Drive permission error: " + error.toString();
  }
}

function createOnEditTrigger() {
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}
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
const BUFFER_SHEET_HEADERS = ['Date', 'OriginalFileName', 'ChangedFilename', 'Invoice ID', 'Drive File ID', 'Gmail Message ID', 'Reason', 'Status', 'UI'];
const BUFFER2_SHEET_HEADERS = [
  'Date',
  'OriginalFileName',
  'ChangedFilename',
  'Invoice ID',
  'Drive File ID',
  'Gmail Message ID',
  'Relevance',
  'UI'
];
const MAIN_SHEET_HEADERS = [
  'File Name', 'File ID', 'File URL',
  'Date Created (Drive)', 'Last Updated (Drive)', 'Size (bytes)', 'Mime Type',
  'Email Subject', 'Gmail Message ID', 'invoice status', 'UI'
];

// Global counter for unique identifiers
var uniqueIdentifierCounter = 0;

// Global object to store file ID to unique identifier mappings
var FILE_IDENTIFIER_MAP = {};

// Edge case tracking objects
var PROCESSED_EMAILS_LOG = {}; // Track emails by message ID
var ATTACHMENT_PROCESSING_LOG = {}; // Track attachments by message ID + attachment name
var ERROR_RECOVERY_LOG = {}; // Track failed operations for retry


/**
 * Generates a unique identifier for a file - ONLY USED FOR INITIAL ASSIGNMENT IN BUFFER SHEET
 * @param {string} fileId - The file ID to generate a unique identifier for
 * @returns {string} Unique identifier like V001, V002, V003
 */
function generateUniqueIdentifierForFile(fileId) {
  // Check if we already have a unique identifier for this file
  if (FILE_IDENTIFIER_MAP[fileId]) {
    return FILE_IDENTIFIER_MAP[fileId];
  }
  
  // Generate a unique identifier
  uniqueIdentifierCounter++;
  const uniqueId = `V${String(uniqueIdentifierCounter).padStart(6, '0')}`;
  
  // Store the mapping
  FILE_IDENTIFIER_MAP[fileId] = uniqueId;
  
  Logger.log(`Generated unique identifier '${uniqueId}' for file ID: ${fileId}`);
  return uniqueId;
}

/**
 * Gets the UI (unique identifier) for a file from the buffer sheet - this is the ONLY source of truth
 * @param {string} companyName - The company name
 * @param {string} changedFilename - The changed filename to look for
 * @param {string} driveFileId - The drive file ID to look for
 * @returns {string} The unique identifier from buffer sheet or empty string if not found
 */
function getUIFromBufferSheet(companyName, changedFilename, driveFileId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bufferSheet = ss.getSheetByName(`${companyName}-buffer`);
  
  if (!bufferSheet) {
    Logger.log(`Buffer sheet not found for company: ${companyName}`);
    return '';
  }
  
  const bufferData = bufferSheet.getDataRange().getValues();
  
  // Look for the file by filename or file ID
  for (let i = 1; i < bufferData.length; i++) {
    const row = bufferData[i];
    const bufferChangedFilename = row[1]; // ChangedFilename column
    const bufferDriveFileId = row[3]; // Drive File ID column
    const bufferUI = row[7]; // UI column (Unique Identifier)
    
    // Match by either filename or file ID
    if ((changedFilename && bufferChangedFilename === changedFilename) || 
        (driveFileId && bufferDriveFileId === driveFileId)) {
      Logger.log(`Found unique identifier '${bufferUI}' for file ${changedFilename || driveFileId} from buffer sheet`);
      return bufferUI || '';
    }
  }
  
  Logger.log(`No UI found in buffer sheet for file: ${changedFilename || driveFileId}`);
  return '';
}

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
 * Create folder structure: Company/FY-XX-XX/Accruals/Buffer2 (for non-invoice files)
 * @param {string} companyName - Company name (analogy/humane)
 * @param {string} financialYear - Financial year (FY-XX-XX)
 * @returns {GoogleAppsScript.Drive.DriveFolder} Buffer2 folder
 */
function createBuffer2FolderStructure(companyName, financialYear) {
  try {
    const companyFolder = DriveApp.getFolderById(ATTACHMENT_COMPANY_FOLDER_MAP[companyName]);
    const financialYearFolder = getOrCreateFolder(companyFolder, financialYear);
    const accrualsFolder = getOrCreateFolder(financialYearFolder, "Accruals");
    const buffer2Folder = getOrCreateFolder(accrualsFolder, "Buffer2");
    return buffer2Folder;
  } catch (error) {
    console.error('Error creating buffer2 folder structure:', error);
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
 * Enhanced tracking functions for edge case handling
 */

/**
 * Gets all processed Gmail message IDs from all relevant sheets to prevent email duplication
 * @param {string} companyName - The company name
 * @returns {Set<string>} Set of all processed Gmail message IDs
 */
function getAllProcessedGmailMessageIds(companyName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const processedIds = new Set();
  
  // Check buffer sheet
  const bufferSheet = ss.getSheetByName(`${companyName}-buffer`);
  if (bufferSheet && bufferSheet.getLastRow() > 1) {
    const bufferData = bufferSheet.getDataRange().getValues();
    for (let i = 1; i < bufferData.length; i++) {
      const gmailId = bufferData[i][4]; // Gmail Message ID column
      if (gmailId) processedIds.add(gmailId);
    }
  }
  
  // Check main sheet
  const mainSheet = ss.getSheetByName(companyName);
  if (mainSheet && mainSheet.getLastRow() > 1) {
    const mainData = mainSheet.getDataRange().getValues();
    const gmailIdIndex = MAIN_SHEET_HEADERS.indexOf('Gmail Message ID');
    if (gmailIdIndex !== -1) {
      for (let i = 1; i < mainData.length; i++) {
        const gmailId = mainData[i][gmailIdIndex];
        if (gmailId) processedIds.add(gmailId);
      }
    }
  }
  
  // Check inflow and outflow sheets
  ['inflow', 'outflow'].forEach(flowType => {
    const flowSheet = ss.getSheetByName(`${companyName}-${flowType}`);
    if (flowSheet && flowSheet.getLastRow() > 1) {
      const flowData = flowSheet.getDataRange().getValues();
      const gmailIdIndex = MAIN_SHEET_HEADERS.indexOf('Gmail Message ID');
      if (gmailIdIndex !== -1) {
        for (let i = 1; i < flowData.length; i++) {
          const gmailId = flowData[i][gmailIdIndex];
          if (gmailId) processedIds.add(gmailId);
        }
      }
    }
  });
  
  Logger.log(`Found ${processedIds.size} already processed Gmail message IDs for company: ${companyName}`);
  return processedIds;
}

/**
 * Tracks email processing with metadata to prevent omission
 * @param {string} messageId - Gmail message ID
 * @param {Object} emailData - Email metadata
 */
function trackEmailProcessing(messageId, emailData) {
  PROCESSED_EMAILS_LOG[messageId] = {
    processedAt: new Date(),
    subject: emailData.subject,
    attachmentCount: emailData.attachmentCount,
    attachmentsProcessed: emailData.attachmentsProcessed || 0,
    status: emailData.status || 'processing'
  };
}

/**
 * Tracks attachment processing to prevent omission
 * @param {string} messageId - Gmail message ID
 * @param {string} attachmentName - Name of the attachment
 * @param {string} status - Processing status
 * @param {string} reason - Reason for status (optional)
 */
function trackAttachmentProcessing(messageId, attachmentName, status, reason = '') {
  const key = `${messageId}_${attachmentName}`;
  ATTACHMENT_PROCESSING_LOG[key] = {
    messageId: messageId,
    attachmentName: attachmentName,
    status: status, // 'processed', 'skipped', 'failed'
    reason: reason,
    processedAt: new Date()
  };
}

/**
 * Gets processing status of an attachment
 * @param {string} messageId - Gmail message ID
 * @param {string} attachmentName - Name of the attachment
 * @returns {Object|null} Processing status or null if not found
 */
function getAttachmentProcessingStatus(messageId, attachmentName) {
  const key = `${messageId}_${attachmentName}`;
  return ATTACHMENT_PROCESSING_LOG[key] || null;
}

/**
 * Validates email completeness - ensures all attachments were processed
 * @param {string} messageId - Gmail message ID
 * @returns {Object} Validation result with status and details
 */
function validateEmailCompleteness(messageId) {
  const emailLog = PROCESSED_EMAILS_LOG[messageId];
  if (!emailLog) {
    return { isComplete: false, reason: 'Email not found in processing log' };
  }
  
  if (emailLog.attachmentCount === 0) {
    return { isComplete: true, reason: 'No attachments to process' };
  }
  
  if (emailLog.attachmentsProcessed < emailLog.attachmentCount) {
    return { 
      isComplete: false, 
      reason: `Only ${emailLog.attachmentsProcessed} of ${emailLog.attachmentCount} attachments processed` 
    };
  }
  
  return { isComplete: true, reason: 'All attachments processed successfully' };
}

/**
 * Logs processing errors for recovery attempts
 * @param {string} messageId - Gmail message ID
 * @param {string} attachmentName - Name of the attachment (optional)
 * @param {string} error - Error details
 * @param {string} operation - Operation that failed
 */
function logProcessingError(messageId, attachmentName, error, operation) {
  const key = attachmentName ? `${messageId}_${attachmentName}` : messageId;
  ERROR_RECOVERY_LOG[key] = {
    messageId: messageId,
    attachmentName: attachmentName,
    error: error,
    operation: operation,
    timestamp: new Date(),
    retryCount: (ERROR_RECOVERY_LOG[key]?.retryCount || 0) + 1
  };
}

/**
 * Creates a comprehensive processing report
 * @param {string} companyName - The company name
 * @returns {Object} Detailed processing report
 */
function createProcessingReport(companyName) {
  const report = {
    companyName: companyName,
    timestamp: new Date(),
    emailsProcessed: Object.keys(PROCESSED_EMAILS_LOG).length,
    attachmentsProcessed: 0,
    attachmentsSkipped: 0,
    attachmentsFailed: 0,
    incompleteEmails: [],
    errors: [],
    duplicatesDetected: 0
  };
  
  // Analyze attachment processing
  Object.values(ATTACHMENT_PROCESSING_LOG).forEach(log => {
    switch (log.status) {
      case 'processed':
        report.attachmentsProcessed++;
        break;
      case 'skipped':
        report.attachmentsSkipped++;
        if (log.reason.includes('duplicate')) {
          report.duplicatesDetected++;
        }
        break;
      case 'failed':
        report.attachmentsFailed++;
        break;
    }
  });
  
  // Check for incomplete emails
  Object.entries(PROCESSED_EMAILS_LOG).forEach(([messageId, emailLog]) => {
    const validation = validateEmailCompleteness(messageId);
    if (!validation.isComplete) {
      report.incompleteEmails.push({
        messageId: messageId,
        subject: emailLog.subject,
        reason: validation.reason
      });
    }
  });
  
  // Collect errors
  report.errors = Object.values(ERROR_RECOVERY_LOG);
  
  Logger.log(`Processing Report for ${companyName}: ${JSON.stringify(report, null, 2)}`);
  return report;
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
  const range = sheet.getRange("G:G"); // Status column is G (index 6, 0-indexed)
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
 * Logs a file's details to a specified log sheet with UI.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} logSheet The sheet to log to.
 * @param {GoogleAppsScript.Drive.File} driveFile The Drive file object.
 * @param {string} emailSubject The subject of the original email.
 * @param {string} gmailMessageId The ID of the original Gmail message.
 * @param {string} invoiceStatus The determined invoice status (inflow, outflow, unknown).
 * @param {string} companyName The company name to get UI from buffer sheet.
 * @param {string} providedUI Optional UI to use instead of looking up from buffer sheet.
 */
function logFileToMainSheet(logSheet, driveFile, emailSubject, gmailMessageId, invoiceStatus, companyName, providedUI = null) {
  // Use provided UI if available, otherwise get from buffer sheet
  let ui = providedUI;
  if (!ui) {
    ui = getUIFromBufferSheet(companyName, driveFile.getName(), driveFile.getId());
  }
  
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
    invoiceStatus,
    ui
  ]);
  Logger.log(`Logged file ${driveFile.getName()} (ID: ${driveFile.getId()}) with UI '${ui}' to ${logSheet.getName()}`);
  sortSheetByDateDesc(logSheet, 4); // Sort by 'Date Created (Drive)' (column 4)
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
      bufferSheet.setColumnWidth(8, 120); // UI
    }
    // Remove any extra columns if present
    if (bufferSheet.getLastColumn() > BUFFER_SHEET_HEADERS.length) {
      bufferSheet.deleteColumns(BUFFER_SHEET_HEADERS.length + 1, bufferSheet.getLastColumn() - BUFFER_SHEET_HEADERS.length);
    }
    setStatusDropdownValidation(bufferSheet); // Apply dropdown validation
  } catch (e) {
    return { status: 'error', message: `Error setting up buffer sheet for ${bufferLabelName}: ${e.toString()}` };
  }

  // Initialize Buffer2 sheet with headers even if no files will be processed
  const buffer2SheetName = `${companyName}-buffer2`;
  let buffer2Sheet;
  try {
    buffer2Sheet = ss.getSheetByName(buffer2SheetName);
    if (!buffer2Sheet) {
      buffer2Sheet = ss.insertSheet(buffer2SheetName);
      Logger.log(`Created new Buffer2 sheet: ${buffer2SheetName}`);
    }
    
    // Ensure headers are correct for Buffer2 sheet
    const buffer2FirstRowRange = buffer2Sheet.getRange(1, 1, 1, BUFFER2_SHEET_HEADERS.length);
    const currentBuffer2Headers = buffer2FirstRowRange.getValues()[0];
    
    if (currentBuffer2Headers.join(',') !== BUFFER2_SHEET_HEADERS.join(',')) {
      buffer2Sheet.clear();
      buffer2Sheet.appendRow(BUFFER2_SHEET_HEADERS);
      buffer2Sheet.getRange(1, 1, 1, BUFFER2_SHEET_HEADERS.length).setFontWeight('bold').setBackground('#E8F0FE').setBorder(true, true, true, true, true, true);
      buffer2Sheet.setFrozenRows(1);
      buffer2Sheet.setColumnWidth(1, 300); // File name
      buffer2Sheet.setColumnWidth(2, 200); // Gmail id
      setRelevanceDropdownValidation(buffer2Sheet);
      Logger.log(`Initialized Buffer2 sheet headers for ${buffer2SheetName}`);
    }
    
    // Remove any extra columns if present
    if (buffer2Sheet.getLastColumn() > BUFFER2_SHEET_HEADERS.length) {
      buffer2Sheet.deleteColumns(BUFFER2_SHEET_HEADERS.length + 1, buffer2Sheet.getLastColumn() - BUFFER2_SHEET_HEADERS.length);
    }
    setRelevanceDropdownValidation(buffer2Sheet);
  } catch (e) {
    Logger.log(`Warning: Could not initialize Buffer2 sheet for ${buffer2SheetName}: ${e.toString()}`);
    // Don't return error here as this is not critical for the main process
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

    // Enhanced tracking with comprehensive duplicate prevention
    const processedGmailMessageIds = getAllProcessedGmailMessageIds(companyName);
    const existingChangedFilenamesInCurrentBuffer = getExistingChangedFilenames(bufferSheet);
    let totalNewAttachments = 0;
    let processedAttachments = 0;
    let skippedAttachments = 0;
    let emailsAnalyzed = 0;
    let emailsSkipped = 0;
    
    // Clear processing logs for this session
    PROCESSED_EMAILS_LOG = {};
    ATTACHMENT_PROCESSING_LOG = {};
    ERROR_RECOVERY_LOG = {};

    // First pass: Enhanced email and attachment analysis with comprehensive tracking
    const threads = gmailLabel.getThreads();
    Logger.log(`Starting analysis of ${threads.length} email threads for label: ${labelName}`);
    
    for (let t = 0; t < threads.length; t++) {
      if (shouldCancel(processToken)) {
        return { status: 'cancelled', message: "Process cancelled during attachment counting." };
      }
      
      const messages = threads[t].getMessages();
      for (let m = 0; m < messages.length; m++) {
        const message = messages[m];
        const messageId = message.getId();
        const subject = message.getSubject();
        const attachments = message.getAttachments().filter(a => !a.isGoogleType() && !a.getName().startsWith('ATT'));
        
        emailsAnalyzed++;
        
        // Track email for comprehensive monitoring
        trackEmailProcessing(messageId, {
          subject: subject,
          attachmentCount: attachments.length,
          status: 'analyzing'
        });
        
        if (processedGmailMessageIds.has(messageId)) {
          emailsSkipped++;
          skippedAttachments += attachments.length;
          
          // Track skipped attachments
          attachments.forEach(attachment => {
            trackAttachmentProcessing(messageId, attachment.getName(), 'skipped', 'Email already processed');
          });
          
          // Update email status
          PROCESSED_EMAILS_LOG[messageId].status = 'already_processed';
          continue;
        }
        
        // Count valid attachments for new emails
        totalNewAttachments += attachments.length;
        
        // Pre-validate attachments for potential issues
        attachments.forEach(attachment => {
          const attachmentName = attachment.getName();
          const status = getAttachmentProcessingStatus(messageId, attachmentName);
          
          if (status) {
            Logger.log(`Warning: Attachment ${attachmentName} from message ${messageId} was previously processed with status: ${status.status}`);
          }
        });
      }
      
      // Periodic progress update
      if (t % 10 === 0) {
        Logger.log(`Analyzed ${t + 1}/${threads.length} threads. Found ${totalNewAttachments} new attachments in ${emailsAnalyzed - emailsSkipped} new emails.`);
      }
      
      Utilities.sleep(50);
    }
    
    Logger.log(`Analysis complete: ${emailsAnalyzed} emails analyzed, ${emailsSkipped} already processed, ${totalNewAttachments} new attachments found`);

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
            const originalFilename = attachment.getName();
            
            try {
              // Check if this specific attachment was already processed
              const existingAttachmentStatus = getAttachmentProcessingStatus(messageId, originalFilename);
              if (existingAttachmentStatus && existingAttachmentStatus.status === 'processed') {
                Logger.log(`Skipping attachment ${originalFilename} from message ${messageId}: already processed`);
                trackAttachmentProcessing(messageId, originalFilename, 'skipped', 'Already processed in previous run');
                skippedAttachments++;
                attachmentsProcessedForThisMessage++;
                continue;
              }
              
              // Use AI to extract comprehensive data for proper renaming
              const blob = attachment.copyBlob();
              const aiExtractedData = callGeminiAPIInternal(blob, originalFilename);
              
              Logger.log(`AI extracted data for ${originalFilename}: ${JSON.stringify(aiExtractedData)}`);
              
              // Handle multi-invoice files
              if (aiExtractedData.isMultiInvoice && aiExtractedData.invoiceCount > 1) {
                Logger.log(`Multi-invoice file detected: ${originalFilename} contains ${aiExtractedData.invoiceCount} invoices`);
                
                // Process each invoice separately
                for (let invoiceIndex = 0; invoiceIndex < aiExtractedData.invoices.length; invoiceIndex++) {
                  const invoiceData = aiExtractedData.invoices[invoiceIndex];
                  
                  // Generate unique filename for each invoice
                  const multiInvoiceFilename = generateMultiInvoiceFilename(invoiceData, originalFilename, invoiceIndex + 1, aiExtractedData.invoiceCount);
                  
                  Logger.log(`Processing invoice ${invoiceIndex + 1}/${aiExtractedData.invoiceCount}: ${multiInvoiceFilename}`);
                  
                  // Check for duplicates
                  const isDuplicateMultiFilename = existingChangedFilenamesInCurrentBuffer.has(multiInvoiceFilename);
                  let multiInvoiceFile = null;
                  let multiUniqueIdentifier = '';
                  
                  if (!isDuplicateMultiFilename) {
                    try {
                      // Create separate file for each invoice
                      const now = new Date();
                      const financialYear = calculateFinancialYear(now);
                      const bufferActiveFolder = getOrCreateBufferSubfolder(companyName, financialYear, "Active");
                      
                      const multiInvoiceBlob = attachment.copyBlob().setName(multiInvoiceFilename);
                      multiInvoiceFile = bufferActiveFolder.createFile(multiInvoiceBlob);
                      
                      multiUniqueIdentifier = generateUniqueIdentifierForFile(multiInvoiceFile.getId());
                      
                      processedAttachments++;
                      attachmentsProcessedForThisMessage++;
                      existingChangedFilenamesInCurrentBuffer.add(multiInvoiceFilename);
                      
                      // Track successful processing
                      trackAttachmentProcessing(messageId, `${originalFilename}_invoice_${invoiceIndex + 1}`, 'processed', `Multi-invoice file: ${invoiceIndex + 1}/${aiExtractedData.invoiceCount}`);
                      
                      // Log to main sheet
                      const emailSubject = message.getSubject ? message.getSubject() : '';
                      const invoiceStatus = invoiceData.invoiceStatus || "unknown";
                      
                      let mainSheet = ss.getSheetByName(companyName);
                      if (!mainSheet) {
                        mainSheet = ss.insertSheet(companyName);
                        mainSheet.appendRow(MAIN_SHEET_HEADERS);
                      }
                      logFileToMainSheet(mainSheet, multiInvoiceFile, emailSubject, messageId, invoiceStatus, companyName, multiUniqueIdentifier);
                      
                      // Handle inflow/outflow for each invoice
                      if (invoiceStatus === "inflow" || invoiceStatus === "outflow") {
                        const month = getMonthFromDate(now);
                        const flowFolder = createFlowFolderStructure(companyName, financialYear, month, invoiceStatus);
                        
                        let flowFile = null;
                        try {
                          flowFile = multiInvoiceFile.makeCopy(multiInvoiceFilename, flowFolder);
                          Logger.log(`Copied multi-invoice file ${multiInvoiceFilename} to ${invoiceStatus} folder.`);
                        } catch (copyErr) {
                          Logger.log(`Error copying multi-invoice file ${multiInvoiceFilename} to ${invoiceStatus} folder: ${copyErr}`);
                        }
                        
                        // Log to inflow/outflow sheet
                        const flowSheetName = `${companyName}-${invoiceStatus}`;
                        let flowSheet = ss.getSheetByName(flowSheetName);
                        if (!flowSheet) {
                          flowSheet = ss.insertSheet(flowSheetName);
                          flowSheet.appendRow(MAIN_SHEET_HEADERS);
                          flowSheet.getRange(1, 1, 1, MAIN_SHEET_HEADERS.length).setFontWeight('bold').setBackground('#E8F0FE').setBorder(true, true, true, true, true, true);
                          flowSheet.setFrozenRows(1);
                        }
                        if (flowFile) {
                          logFileToMainSheet(flowSheet, flowFile, emailSubject, messageId, invoiceStatus, companyName, multiUniqueIdentifier);
                        }
                      }
                      
                    } catch (multiFileError) {
                      Logger.log(`Error creating multi-invoice file ${invoiceIndex + 1}: ${multiFileError.toString()}`);
                      logProcessingError(messageId, `${originalFilename}_invoice_${invoiceIndex + 1}`, multiFileError.toString(), 'multi_invoice_creation');
                      trackAttachmentProcessing(messageId, `${originalFilename}_invoice_${invoiceIndex + 1}`, 'failed', `Multi-invoice creation failed: ${multiFileError.message}`);
                      continue;
                    }
                  } else {
                    // Duplicate multi-invoice file
                    multiUniqueIdentifier = generateUniqueIdentifierForFile(`multi_duplicate_${messageId}_${invoiceIndex}`);
                    Logger.log(`Skipping duplicate multi-invoice file: ${multiInvoiceFilename}`);
                    trackAttachmentProcessing(messageId, `${originalFilename}_invoice_${invoiceIndex + 1}`, 'skipped', `Duplicate multi-invoice filename`);
                    skippedAttachments++;
                    attachmentsProcessedForThisMessage++;
                  }
                  
                  // Log to buffer sheet for each invoice
                  const multiInvoiceRowData = [
                    `${originalFilename} (Invoice ${invoiceIndex + 1}/${aiExtractedData.invoiceCount})`, // Enhanced original filename
                    multiInvoiceFilename,
                    invoiceData.invoiceNumber || 'INV-Unknown',
                    multiInvoiceFile ? multiInvoiceFile.getId() : '',
                    messageId,
                    `Multi-invoice: ${invoiceIndex + 1}/${aiExtractedData.invoiceCount} | ${invoiceData.documentType || 'document'} | Amount: ${invoiceData.amount || '0.00'}`,
                    'Active',
                    multiUniqueIdentifier
                  ];
                  bufferSheet.appendRow(multiInvoiceRowData);
                  sortSheetByDateDesc(bufferSheet, 1);
                  const newMultiRowIndex = bufferSheet.getLastRow();
                  if (isDuplicateMultiFilename) {
                    bufferSheet.getRange(newMultiRowIndex, 1, 1, BUFFER_SHEET_HEADERS.length).setBackground('#FFFF00'); // Yellow for duplicate
                  }
                  
                  Utilities.sleep(50); // Small delay between invoice processing
                }
                
                // Skip the regular single-file processing since we handled it as multi-invoice
                continue;
              }
              
              // Regular single-invoice processing
              const changedFilename = generateNewFilename(aiExtractedData, originalFilename);
              Logger.log(`Generated filename: ${changedFilename}`);
              
              // Store AI-extracted data for later use
              const extractedInvoiceStatus = aiExtractedData.invoiceStatus || 'unknown';

              // Check if this is a non-invoice/bill file (unknown status)
              if (extractedInvoiceStatus === 'unknown') {
                // Move to Buffer2 folder and log to Buffer2 sheet
                try {
                  const now = new Date();
                  const financialYear = calculateFinancialYear(now);
                  const buffer2Folder = createBuffer2FolderStructure(companyName, financialYear);
                  
                  // Create the file in Buffer2
                  const renamedBlob = attachment.copyBlob().setName(changedFilename);
                  const buffer2File = buffer2Folder.createFile(renamedBlob);
                  
                  processedAttachments++;
                  attachmentsProcessedForThisMessage++;
                  
                  // Track successful processing
                  trackAttachmentProcessing(messageId, originalFilename, 'processed', `Non-invoice file moved to Buffer2: ${changedFilename}`);
                  
                  // Log to Buffer2 sheet
                  const buffer2SheetName = `${companyName}-buffer2`;
                  let buffer2Sheet = ss.getSheetByName(buffer2SheetName);
                  if (!buffer2Sheet) {
                    buffer2Sheet = ss.insertSheet(buffer2SheetName);
                    buffer2Sheet.appendRow(BUFFER2_SHEET_HEADERS);
                    buffer2Sheet.getRange(1, 1, 1, BUFFER2_SHEET_HEADERS.length).setFontWeight('bold').setBackground('#E8F0FE').setBorder(true, true, true, true, true, true);
                    buffer2Sheet.setFrozenRows(1);
                    buffer2Sheet.setColumnWidth(1, 300); // File name
                    buffer2Sheet.setColumnWidth(2, 200); // Gmail id
                    setRelevanceDropdownValidation(buffer2Sheet);
                  }
                  
                  // Ensure headers are correct
                  const buffer2FirstRowRange = buffer2Sheet.getRange(1, 1, 1, BUFFER2_SHEET_HEADERS.length);
                  const currentBuffer2Headers = buffer2FirstRowRange.getValues()[0];
                  if (currentBuffer2Headers.join(',') !== BUFFER2_SHEET_HEADERS.join(',')) {
                    buffer2Sheet.clear();
                    buffer2Sheet.appendRow(BUFFER2_SHEET_HEADERS);
                    buffer2Sheet.getRange(1, 1, 1, BUFFER2_SHEET_HEADERS.length).setFontWeight('bold').setBackground('#E8F0FE').setBorder(true, true, true, true, true, true);
                    buffer2Sheet.setFrozenRows(1);
                    buffer2Sheet.setColumnWidth(1, 300); // File name
                    buffer2Sheet.setColumnWidth(2, 200); // Gmail id
                    setRelevanceDropdownValidation(buffer2Sheet);
                  }
                  
                  // Log the non-invoice file to Buffer2 sheet
                  buffer2Sheet.appendRow([
                    new Date(),
                    originalFilename,
                    changedFilename,
                    aiExtractedData.invoiceNumber || 'INV-Unknown',
                    '', // Drive File ID (not available for Buffer2)
                    messageId,
                    '', // Relevance blank by default
                    ''  // UI blank by default
                  ]);
                  sortSheetByDateDesc(buffer2Sheet, 1);
                  setRelevanceDropdownValidation(buffer2Sheet);
                  
                  Logger.log(`Non-invoice file ${changedFilename} moved to Buffer2 folder and logged to ${buffer2SheetName}`);
                  
                  // Continue to next attachment since this was handled as Buffer2
                  continue;
                  
                } catch (buffer2Error) {
                  Logger.log(`Error moving file to Buffer2: ${buffer2Error.toString()}`);
                  logProcessingError(messageId, originalFilename, buffer2Error.toString(), 'buffer2_creation');
                  trackAttachmentProcessing(messageId, originalFilename, 'failed', `Buffer2 creation failed: ${buffer2Error.message}`);
                  // Fall through to regular processing if Buffer2 fails
                }
              }

              const isDuplicateChangedFilename = existingChangedFilenamesInCurrentBuffer.has(changedFilename);
              let driveFile = null;
              let uniqueIdentifier = '';

              if (!isDuplicateChangedFilename) {
                // Store in Buffer/Active with the correct name
                const now = new Date();
                const month = getMonthFromDate(now);
                const financialYear = calculateFinancialYear(now);
                const bufferActiveFolder = getOrCreateBufferSubfolder(companyName, financialYear, "Active");

                try {
                  // Create the file with the changedFilename
                  const renamedBlob = attachment.copyBlob().setName(changedFilename);
                  driveFile = bufferActiveFolder.createFile(renamedBlob);
                  
                  // Generate unique identifier for the file - IMMEDIATELY after file creation
                  uniqueIdentifier = generateUniqueIdentifierForFile(driveFile.getId());
                  
                  processedAttachments++;
                  attachmentsProcessedForThisMessage++;
                  existingChangedFilenamesInCurrentBuffer.add(changedFilename);
                  
                  // Track successful processing
                  trackAttachmentProcessing(messageId, originalFilename, 'processed', `Successfully saved as ${changedFilename}`);
                  
                } catch (fileCreationError) {
                  Logger.log(`Error creating file for attachment ${originalFilename}: ${fileCreationError.toString()}`);
                  logProcessingError(messageId, originalFilename, fileCreationError.toString(), 'file_creation');
                  trackAttachmentProcessing(messageId, originalFilename, 'failed', `File creation failed: ${fileCreationError.message}`);
                  continue; // Skip to next attachment
                }

                // Use previously extracted AI data for invoice status
                const emailSubject = message.getSubject ? message.getSubject() : '';
                const invoiceStatus = extractedInvoiceStatus || "unknown";

                // Log to main sheet
                let mainSheet = ss.getSheetByName(companyName);
                if (!mainSheet) {
                  mainSheet = ss.insertSheet(companyName);
                  mainSheet.appendRow(MAIN_SHEET_HEADERS);
                }
                logFileToMainSheet(mainSheet, driveFile, emailSubject, messageId, invoiceStatus, companyName, uniqueIdentifier);

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
                    flowSheet.getRange(1, 1, 1, MAIN_SHEET_HEADERS.length).setFontWeight('bold').setBackground('#E8F0FE').setBorder(true, true, true, true, true, true);
                    flowSheet.setFrozenRows(1);
                  } else {
                    // Ensure existing flow sheet has correct headers with UI column
                    const flowFirstRowRange = flowSheet.getRange(1, 1, 1, MAIN_SHEET_HEADERS.length);
                    if (flowSheet.getLastRow() === 0 || flowFirstRowRange.getValues()[0].join(',') !== MAIN_SHEET_HEADERS.join(',')) {
                      // Only clear and reset if headers don't match exactly
                      if (flowSheet.getLastRow() > 0) {
                        const currentHeaders = flowFirstRowRange.getValues()[0];
                        if (currentHeaders.length < MAIN_SHEET_HEADERS.length || currentHeaders[currentHeaders.length - 1] !== 'UI') {
                          // Add UI column if missing
                          if (flowSheet.getLastColumn() < MAIN_SHEET_HEADERS.length) {
                            flowSheet.getRange(1, flowSheet.getLastColumn() + 1).setValue('UI');
                            flowSheet.getRange(1, flowSheet.getLastColumn()).setFontWeight('bold').setBackground('#E8F0FE').setBorder(true, true, true, true, true, true);
                          }
                        }
                      } else {
                        flowSheet.appendRow(MAIN_SHEET_HEADERS);
                        flowSheet.getRange(1, 1, 1, MAIN_SHEET_HEADERS.length).setFontWeight('bold').setBackground('#E8F0FE').setBorder(true, true, true, true, true, true);
                        flowSheet.setFrozenRows(1);
                      }
                    }
                  }
                  if (flowFile) {
                    logFileToMainSheet(flowSheet, flowFile, emailSubject, messageId, invoiceStatus, companyName, uniqueIdentifier);
                  }
                }
              } else {
                // If duplicate, don't upload again, but still generate unique identifier for buffer logging
                uniqueIdentifier = generateUniqueIdentifierForFile(`duplicate_${messageId}_${a}`);
                Logger.log(`Skipping upload: Duplicate changed filename '${changedFilename}' already exists in buffer folder for '${labelName}'. Generated UI: ${uniqueIdentifier}`);
                
                // Track duplicate detection
                trackAttachmentProcessing(messageId, originalFilename, 'skipped', `Duplicate filename: ${changedFilename}`);
                skippedAttachments++;
                attachmentsProcessedForThisMessage++;
              }

              // Append to buffer sheet with AI-extracted data
              const rowData = [
                new Date(),
                originalFilename,
                changedFilename,
                aiExtractedData.invoiceNumber || 'INV-Unknown', // Use AI-extracted invoice number
                driveFile ? driveFile.getId() : '', // Drive File ID
                messageId,                               // Gmail Message ID
                `AI: ${aiExtractedData.documentType || 'document'} | Amount: ${aiExtractedData.amount || '0.00'}`, // Enhanced reason with AI data
                'Active',                                // Default Status (Active)
                uniqueIdentifier                         // UI (unique identifier)
              ];
              bufferSheet.appendRow(rowData);
              sortSheetByDateDesc(bufferSheet, 1);
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
              
              // Log error for recovery and tracking
              logProcessingError(messageId, originalFilename, fileError.toString(), 'attachment_processing');
              trackAttachmentProcessing(messageId, originalFilename, 'failed', `Processing error: ${fileError.message}`);
              
              // Continue processing other attachments
            }
          }
        }
        // Enhanced email completion tracking
        const emailLog = PROCESSED_EMAILS_LOG[messageId];
        if (emailLog) {
          emailLog.attachmentsProcessed = attachmentsProcessedForThisMessage;
          emailLog.status = 'completed';
          
          // Validate email completeness
          const validation = validateEmailCompleteness(messageId);
          if (!validation.isComplete) {
            Logger.log(`Warning: Email ${messageId} (${message.getSubject()}) incomplete: ${validation.reason}`);
            emailLog.status = 'incomplete';
            logProcessingError(messageId, null, validation.reason, 'email_completion_check');
          }
        }
        
        // Only mark as processed if we actually processed some attachments or if there were no valid attachments
        if (attachmentsProcessedForThisMessage > 0 || attachments.length === 0) {
          processedGmailMessageIds.add(messageId);
        }
      }
    }

    clearCancellationToken();
    
    // Generate comprehensive processing report
    const processingReport = createProcessingReport(companyName);
    
    let resultMessage = `Completed processing for '${labelName}'. `;
    resultMessage += `Emails analyzed: ${emailsAnalyzed}, New emails: ${emailsAnalyzed - emailsSkipped}. `;
    resultMessage += `Attachments processed: ${processedAttachments}, Skipped: ${skippedAttachments}. `;
    
    // Add warnings for incomplete processing
    if (processingReport.incompleteEmails.length > 0) {
      resultMessage += `⚠️ Warning: ${processingReport.incompleteEmails.length} emails had incomplete attachment processing. `;
    }
    
    if (processingReport.errors.length > 0) {
      resultMessage += `⚠️ ${processingReport.errors.length} errors occurred during processing. `;
    }
    
    if (processingReport.duplicatesDetected > 0) {
      resultMessage += `📋 ${processingReport.duplicatesDetected} duplicate files detected and skipped. `;
    }
    
    // Log detailed report
    Logger.log(`Processing Report Summary for ${labelName}:`);
    Logger.log(`- Total emails analyzed: ${processingReport.emailsProcessed}`);
    Logger.log(`- Attachments processed: ${processingReport.attachmentsProcessed}`);
    Logger.log(`- Attachments skipped: ${processingReport.attachmentsSkipped}`);
    Logger.log(`- Attachments failed: ${processingReport.attachmentsFailed}`);
    Logger.log(`- Incomplete emails: ${processingReport.incompleteEmails.length}`);
    Logger.log(`- Errors logged: ${processingReport.errors.length}`);
    
    // After processing into company folder, automatically trigger the buffer processing for this company
    processBufferFilesAndLog(labelName);

    return { 
      status: 'success', 
      message: resultMessage,
      report: processingReport
    };

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
  } else {
    // Ensure existing inflow sheet has correct headers with UI column
    const inflowFirstRowRange = inflowSheet.getRange(1, 1, 1, MAIN_SHEET_HEADERS.length);
    if (inflowSheet.getLastRow() === 0 || inflowFirstRowRange.getValues()[0].join(',') !== MAIN_SHEET_HEADERS.join(',')) {
      // Only clear and reset if headers don't match exactly
      if (inflowSheet.getLastRow() > 0) {
        const currentHeaders = inflowFirstRowRange.getValues()[0];
        if (currentHeaders.length < MAIN_SHEET_HEADERS.length || currentHeaders[currentHeaders.length - 1] !== 'UI') {
          // Add UI column if missing
          if (inflowSheet.getLastColumn() < MAIN_SHEET_HEADERS.length) {
            inflowSheet.getRange(1, inflowSheet.getLastColumn() + 1).setValue('UI');
            inflowSheet.getRange(1, inflowSheet.getLastColumn()).setFontWeight('bold').setBackground('#E8F0FE').setBorder(true, true, true, true, true, true);
          }
        }
      } else {
        inflowSheet.appendRow(MAIN_SHEET_HEADERS);
        inflowSheet.getRange(1, 1, 1, MAIN_SHEET_HEADERS.length).setFontWeight('bold').setBackground('#E8F0FE').setBorder(true, true, true, true, true, true);
        inflowSheet.setFrozenRows(1);
      }
    }
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
  } else {
    // Ensure existing outflow sheet has correct headers with UI column
    const outflowFirstRowRange = outflowSheet.getRange(1, 1, 1, MAIN_SHEET_HEADERS.length);
    if (outflowSheet.getLastRow() === 0 || outflowFirstRowRange.getValues()[0].join(',') !== MAIN_SHEET_HEADERS.join(',')) {
      // Only clear and reset if headers don't match exactly
      if (outflowSheet.getLastRow() > 0) {
        const currentHeaders = outflowFirstRowRange.getValues()[0];
        if (currentHeaders.length < MAIN_SHEET_HEADERS.length || currentHeaders[currentHeaders.length - 1] !== 'UI') {
          // Add UI column if missing
          if (outflowSheet.getLastColumn() < MAIN_SHEET_HEADERS.length) {
            outflowSheet.getRange(1, outflowSheet.getLastColumn() + 1).setValue('UI');
            outflowSheet.getRange(1, outflowSheet.getLastColumn()).setFontWeight('bold').setBackground('#E8F0FE').setBorder(true, true, true, true, true, true);
          }
        }
      } else {
        outflowSheet.appendRow(MAIN_SHEET_HEADERS);
        outflowSheet.getRange(1, 1, 1, MAIN_SHEET_HEADERS.length).setFontWeight('bold').setBackground('#E8F0FE').setBorder(true, true, true, true, true, true);
        outflowSheet.setFrozenRows(1);
      }
    }
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
    const existingAnimalName = row[7]; // UI column (unique identifier)
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

        // 2. Use existing unique identifier from buffer sheet - buffer sheet is source of truth
        let uniqueIdentifier = existingAnimalName;
        
        // If no unique identifier in buffer, generate one and update buffer sheet
        if (!uniqueIdentifier) {
          uniqueIdentifier = generateUniqueIdentifierForFile(driveFileId);
          bufferSheet.getRange(bufferRowIndex, 8).setValue(uniqueIdentifier); // UI column
          Logger.log(`Updated buffer sheet with new UI '${uniqueIdentifier}' for file ID: ${driveFileId}`);
        }
        
        // 3. Use AI to determine inflow/outflow/unknown
        const emailSubject = getEmailSubjectForMessageId(gmailMessageId);
        const blob = file.getBlob();
        const aiResult = callGeminiAPIInternal(blob, changedFilename);
        const invoiceStatus = aiResult.invoiceStatus || "unknown";
        Logger.log(`AI invoiceStatus for file ${changedFilename}: ${invoiceStatus}`);

        // 4. Delete existing log entries from Main, Inflow, Outflow sheets before re-logging
        deleteLogEntries(mainSheet, driveFileId, gmailMessageId);
        deleteLogEntries(inflowSheet, driveFileId, gmailMessageId);
        deleteLogEntries(outflowSheet, driveFileId, gmailMessageId);


        // 5. Log to main sheet (sheet only, no drive storage)
        logFileToMainSheet(mainSheet, file, emailSubject, gmailMessageId, invoiceStatus, companyName, uniqueIdentifier);

        // 6. Copy to inflow or outflow if appropriate (both sheet and drive)
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
          
          // Use logFileToMainSheet which gets UI from buffer sheet
          logFileToMainSheet(flowSheet, copiedFile, emailSubject, gmailMessageId, invoiceStatus, companyName, uniqueIdentifier);
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
  sortSheetByDateDesc(mainSheet, 4); // Sort by 'Date Created (Drive)' (column 4)
  sortSheetByDateDesc(inflowSheet, 4); // Sort by 'Date Created (Drive)' (column 4)
  sortSheetByDateDesc(outflowSheet, 4); // Sort by 'Date Created (Drive)' (column 4)
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
        
        // Copy UI from the existing row (buffer sheet is source of truth)
        const existingUI = row[10]; // Existing UI from main sheet
        if (newRow.length > 10) {
          newRow[10] = existingUI; // UI column
        } else {
          newRow.push(existingUI);
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
 * Enhanced Gemini AI integration for comprehensive file data extraction.
 * Detects single or multiple invoices and extracts data accordingly.
 * @param {GoogleAppsScript.Base.Blob} fileBlob The content of the file.
 * @param {string} fileName The name of the file.
 * @returns {Object} An object containing extracted data, invoice status, and multi-invoice info.
 */
function callGeminiAPIInternal(fileBlob, fileName) {
  Logger.log(`Calling Gemini AI for comprehensive data extraction from: ${fileName}`);
  
  // Get API key from script properties (set this in your Google Apps Script project)
  const API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  if (!API_KEY) {
    Logger.log('Warning: GEMINI_API_KEY not found in script properties. Using fallback extraction.');
    return fallbackDataExtraction(fileBlob, fileName);
  }
  
  const GEMINI_URL = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${API_KEY}`;
  
  try {
    // Convert file to base64
    const imageData = Utilities.base64Encode(fileBlob.getBytes());
    const mimeType = fileBlob.getContentType();
    
    // Enhanced prompt for multi-invoice detection and extraction
    const prompt = `Analyze this document carefully and determine if it contains single or multiple invoices/bills/receipts.

If SINGLE invoice/document, respond with:
{
  "isMultiInvoice": false,
  "invoiceData": {
    "date": "YYYY-MM-DD format date",
    "vendorName": "Company or vendor name (clean, no special characters)",
    "invoiceNumber": "Invoice/bill/reference number",
    "amount": "Total amount as number (no currency symbols)",
    "invoiceStatus": "inflow or outflow or unknown",
    "documentType": "type of document (invoice, receipt, bill, etc.)"
  }
}

If MULTIPLE invoices/documents, respond with:
{
  "isMultiInvoice": true,
  "invoiceData": [
    {
      "date": "YYYY-MM-DD format date",
      "vendorName": "Company or vendor name (clean, no special characters)",
      "invoiceNumber": "Invoice/bill/reference number",
      "amount": "Amount as number (no currency symbols)",
      "invoiceStatus": "inflow or outflow or unknown",
      "documentType": "type of document (invoice, receipt, bill, etc.)",
      "pageNumber": "Page number or position in document"
    },
    {
      "date": "YYYY-MM-DD format date",
      "vendorName": "Company or vendor name (clean, no special characters)",
      "invoiceNumber": "Invoice/bill/reference number",
      "amount": "Amount as number (no currency symbols)",
      "invoiceStatus": "inflow or outflow or unknown",
      "documentType": "type of document (invoice, receipt, bill, etc.)",
      "pageNumber": "Page number or position in document"
    }
  ]
}

IMPORTANT RULES:
1. Look for multiple invoice numbers, different vendor names, different dates, or separate line items that represent distinct transactions
2. For invoiceStatus: 'inflow' = money coming in (sales invoices, receipts), 'outflow' = money going out (purchase invoices, bills, expenses)
3. Clean vendor names (remove special characters, keep only alphanumeric and spaces)
4. Each invoice must have a unique invoice number - if numbers are the same, it's likely a single invoice
5. If any field cannot be determined, use appropriate defaults: date='${getCurrentDateString()}', vendorName='UnknownVendor', invoiceNumber='INV-Unknown', amount='0.00'
6. Amount should be just the number without currency symbols
7. For multi-invoice files, ensure each invoice has distinct data
8. Respond ONLY with valid JSON, no additional text`;
    
    const payload = {
      contents: [{
        parts: [
          { text: prompt },
          { 
            inlineData: { 
              mimeType: mimeType, 
              data: imageData 
            } 
          }
        ]
      }],
      generationConfig: {
        temperature: 0.1,
        topK: 32,
        topP: 1,
        maxOutputTokens: 2048, // Increased for multi-invoice responses
      }
    };
    
    const options = {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(GEMINI_URL, options);
    const responseText = response.getContentText();
    
    Logger.log(`Gemini API Response Code: ${response.getResponseCode()}`);
    
    if (response.getResponseCode() !== 200) {
      Logger.log(`Gemini API Error: ${responseText}`);
      return fallbackDataExtraction(fileBlob, fileName);
    }
    
    const jsonResponse = JSON.parse(responseText);
    
    if (jsonResponse.candidates && jsonResponse.candidates[0] && 
        jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts) {
      
      const aiText = jsonResponse.candidates[0].content.parts[0].text.trim();
      Logger.log(`Raw AI Response: ${aiText}`);
      
      // Parse JSON response from AI
      try {
        // Clean the response - remove any markdown formatting
        const cleanedText = aiText.replace(/```json\n?|```\n?/g, '').trim();
        const extractedData = JSON.parse(cleanedText);
        
        // Process single or multi-invoice data
        const processedData = processMultiInvoiceData(extractedData, fileName);
        Logger.log(`Processed multi-invoice data: ${JSON.stringify(processedData)}`);
        
        return processedData;
        
      } catch (parseError) {
        Logger.log(`Error parsing AI JSON response: ${parseError.toString()}`);
        Logger.log(`Attempting to extract data from text: ${aiText}`);
        
        // Try to extract data using regex if JSON parsing fails
        return extractDataFromText(aiText, fileName);
      }
    } else {
      Logger.log('Unexpected Gemini API response structure');
      return fallbackDataExtraction(fileBlob, fileName);
    }
    
  } catch (error) {
    Logger.log(`Error calling Gemini API: ${error.toString()}`);
    return fallbackDataExtraction(fileBlob, fileName);
  }
}

/**
 * Processes multi-invoice data from AI response
 * @param {Object} aiData - Raw data from AI
 * @param {string} fileName - Original filename for fallback
 * @returns {Object} Processed data with validation
 */
function processMultiInvoiceData(aiData, fileName) {
  try {
    if (aiData.isMultiInvoice) {
      // Multi-invoice processing
      const invoices = aiData.invoiceData || [];
      const processedInvoices = invoices.map((invoice, index) => {
        return validateAndSanitizeExtractedData(invoice, `${fileName}_invoice_${index + 1}`);
      });
      
      // Filter out invalid invoices (those with default values only)
      const validInvoices = processedInvoices.filter(invoice => 
        invoice.invoiceNumber !== 'INV-Unknown' || 
        invoice.vendorName !== 'UnknownVendor' ||
        invoice.amount !== '0.00'
      );
      
      if (validInvoices.length === 0) {
        Logger.log(`No valid invoices found in multi-invoice file: ${fileName}`);
        return fallbackDataExtraction(null, fileName);
      }
      
      return {
        isMultiInvoice: true,
        invoiceCount: validInvoices.length,
        invoices: validInvoices,
        // For backward compatibility, return first invoice data at root level
        ...validInvoices[0]
      };
    } else {
      // Single invoice processing
      const singleInvoiceData = validateAndSanitizeExtractedData(aiData.invoiceData || aiData, fileName);
      return {
        isMultiInvoice: false,
        invoiceCount: 1,
        invoices: [singleInvoiceData],
        ...singleInvoiceData
      };
    }
  } catch (error) {
    Logger.log(`Error processing multi-invoice data: ${error.toString()}`);
    return fallbackDataExtraction(null, fileName);
  }
}

/**
 * Validates and sanitizes data extracted by AI
 * @param {Object} data - Raw data from AI
 * @param {string} fileName - Original filename for fallback
 * @returns {Object} Validated and sanitized data
 */
function validateAndSanitizeExtractedData(data, fileName) {
  const currentDate = getCurrentDateString();
  
  return {
    date: validateDate(data.date) || currentDate,
    vendorName: sanitizeVendorName(data.vendorName) || 'UnknownVendor',
    invoiceNumber: sanitizeInvoiceNumber(data.invoiceNumber) || extractInvoiceIdFromFilename(fileName) || 'INV-Unknown',
    amount: validateAmount(data.amount) || '0.00',
    invoiceStatus: validateInvoiceStatus(data.invoiceStatus) || 'unknown',
    documentType: data.documentType || 'document'
  };
}

/**
 * Fallback data extraction when AI is not available
 * @param {GoogleAppsScript.Base.Blob} fileBlob - File content
 * @param {string} fileName - Original filename
 * @returns {Object} Extracted data using fallback methods
 */
function fallbackDataExtraction(fileBlob, fileName) {
  Logger.log(`Using fallback data extraction for: ${fileName}`);
  
  const lowerFileName = fileName.toLowerCase();
  let invoiceStatus = 'unknown';
  
  // Simple logic based on filename
  if (lowerFileName.includes('invoice') && !lowerFileName.includes('payment')) {
    invoiceStatus = 'outflow';
  } else if (lowerFileName.includes('receipt') || lowerFileName.includes('deposit') || lowerFileName.includes('credit')) {
    invoiceStatus = 'inflow';
  } else if (lowerFileName.includes('bill') || lowerFileName.includes('expense')) {
    invoiceStatus = 'outflow';
  }
  
  const fallbackData = {
    date: getCurrentDateString(),
    vendorName: 'UnknownVendor',
    invoiceNumber: extractInvoiceIdFromFilename(fileName) || 'INV-Unknown',
    amount: '0.00',
    invoiceStatus: invoiceStatus,
    documentType: 'document'
  };
  
  return {
    isMultiInvoice: false,
    invoiceCount: 1,
    invoices: [fallbackData],
    ...fallbackData
  };
}

/**
 * Extracts data from text using regex when JSON parsing fails
 * @param {string} text - AI response text
 * @param {string} fileName - Original filename for fallback
 * @returns {Object} Extracted data
 */
function extractDataFromText(text, fileName) {
  const data = {
    date: getCurrentDateString(),
    vendorName: 'UnknownVendor',
    invoiceNumber: extractInvoiceIdFromFilename(fileName) || 'INV-Unknown',
    amount: '0.00',
    invoiceStatus: 'unknown',
    documentType: 'document'
  };
  
  // Try to extract date (various formats)
  const dateMatch = text.match(/\b(\d{4}-\d{2}-\d{2}|\d{2}\/\d{2}\/\d{4}|\d{2}-\d{2}-\d{4})\b/);
  if (dateMatch) {
    data.date = standardizeDateFormat(dateMatch[1]);
  }
  
  // Try to extract vendor name
  const vendorMatch = text.match(/vendor[:\s]+([^\n,]+)/i) || text.match(/company[:\s]+([^\n,]+)/i);
  if (vendorMatch) {
    data.vendorName = sanitizeVendorName(vendorMatch[1].trim());
  }
  
  // Try to extract invoice number
  const invoiceMatch = text.match(/invoice[\s#:]+([A-Za-z0-9-]+)/i) || text.match(/\b(INV[0-9-]+)\b/i);
  if (invoiceMatch) {
    data.invoiceNumber = sanitizeInvoiceNumber(invoiceMatch[1]);
  }
  
  // Try to extract amount
  const amountMatch = text.match(/\$?([0-9,]+\.?[0-9]*)/g);
  if (amountMatch && amountMatch.length > 0) {
    // Get the largest amount (likely the total)
    const amounts = amountMatch.map(a => parseFloat(a.replace(/[$,]/g, ''))).filter(a => !isNaN(a));
    if (amounts.length > 0) {
      data.amount = Math.max(...amounts).toFixed(2);
    }
  }
  
  // Try to extract invoice status
  if (text.toLowerCase().includes('inflow')) {
    data.invoiceStatus = 'inflow';
  } else if (text.toLowerCase().includes('outflow')) {
    data.invoiceStatus = 'outflow';
  }
  
  return data;
}

/**
 * Validation helper functions
 */
function validateDate(dateStr) {
  if (!dateStr) return null;
  const date = new Date(dateStr);
  if (isNaN(date.getTime())) return null;
  return standardizeDateFormat(dateStr);
}

function standardizeDateFormat(dateStr) {
  const date = new Date(dateStr);
  if (isNaN(date.getTime())) return getCurrentDateString();
  return date.toISOString().split('T')[0]; // YYYY-MM-DD format
}

function getCurrentDateString() {
  return new Date().toISOString().split('T')[0];
}

function sanitizeVendorName(vendor) {
  if (!vendor) return null;
  return vendor.toString()
    .replace(/[^a-zA-Z0-9\s-]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 50); // Limit length
}

function sanitizeInvoiceNumber(invoice) {
  if (!invoice) return null;
  return invoice.toString()
    .replace(/[^a-zA-Z0-9-]/g, '')
    .trim()
    .substring(0, 20); // Limit length
}

function validateAmount(amount) {
  if (!amount) return null;
  const numAmount = parseFloat(amount.toString().replace(/[^0-9.]/g, ''));
  if (isNaN(numAmount)) return null;
  return numAmount.toFixed(2);
}

function validateInvoiceStatus(status) {
  if (!status) return null;
  const lowerStatus = status.toString().toLowerCase().trim();
  if (['inflow', 'outflow', 'unknown'].includes(lowerStatus)) {
    return lowerStatus;
  }
  return null;
}

/**
 * Generates filename for multi-invoice files
 * @param {Object} invoiceData - Data for specific invoice
 * @param {string} originalFilename - Original filename
 * @param {number} invoiceNumber - Invoice number (1, 2, 3, etc.)
 * @param {number} totalInvoices - Total number of invoices in file
 * @returns {string} Generated filename for specific invoice
 */
function generateMultiInvoiceFilename(invoiceData, originalFilename, invoiceNumber, totalInvoices) {
  const lastDotIndex = originalFilename.lastIndexOf('.');
  const extension = lastDotIndex > 0 ? originalFilename.substring(lastDotIndex) : '';
  
  const date = String(invoiceData.date || getCurrentDateString());
  const vendor = String(invoiceData.vendorName || "UnknownVendor");
  const invoice = String(invoiceData.invoiceNumber || `INV-${invoiceNumber}`);
  const amount = String(invoiceData.amount || "0.00");
  
  // Sanitization
  const sanitizedDate = date.replace(/[^0-9\-]/g, '').trim() || getCurrentDateString();
  const sanitizedVendor = vendor.replace(/[_]/g, '-')
    .replace(/[/\\:*?"<>|]/g, '').replace(/\s+/g, ' ').trim() || "UnknownVendor";
  const sanitizedInvoice = invoice.replace(/[_]/g, '-')
    .replace(/[/\\:*?"<>|]/g, '').replace(/\s+/g, '').trim() || `INV-${invoiceNumber}`;
  const sanitizedAmount = amount.replace(/[_]/g, '')
    .replace(/[/\\:*?"<>|]/g, '').trim() || "0.00";
  
  // Add multi-invoice identifier
  const multiInvoiceId = `Multi${invoiceNumber}of${totalInvoices}`;
  
  return `${sanitizedDate}_${sanitizedVendor}_${sanitizedInvoice}_${sanitizedAmount}_${multiInvoiceId}${extension}`;
}

/**
 * Enhanced tracking for multi-invoice files
 * @param {string} messageId - Gmail message ID
 * @param {string} originalFilename - Original attachment filename
 * @param {number} totalInvoices - Total number of invoices detected
 * @param {number} processedInvoices - Number of invoices successfully processed
 */
function trackMultiInvoiceProcessing(messageId, originalFilename, totalInvoices, processedInvoices) {
  const key = `${messageId}_multi_${originalFilename}`;
  ATTACHMENT_PROCESSING_LOG[key] = {
    messageId: messageId,
    attachmentName: originalFilename,
    status: processedInvoices === totalInvoices ? 'completed' : 'partial',
    reason: `Multi-invoice file: ${processedInvoices}/${totalInvoices} invoices processed`,
    totalInvoices: totalInvoices,
    processedInvoices: processedInvoices,
    processedAt: new Date(),
    isMultiInvoice: true
  };
}

/**
 * Validates multi-invoice processing completeness
 * @param {string} messageId - Gmail message ID
 * @param {string} originalFilename - Original attachment filename
 * @returns {Object} Validation result
 */
function validateMultiInvoiceCompleteness(messageId, originalFilename) {
  const key = `${messageId}_multi_${originalFilename}`;
  const multiLog = ATTACHMENT_PROCESSING_LOG[key];
  
  if (!multiLog || !multiLog.isMultiInvoice) {
    return { isComplete: false, reason: 'Multi-invoice log not found' };
  }
  
  if (multiLog.processedInvoices < multiLog.totalInvoices) {
    return {
      isComplete: false,
      reason: `Only ${multiLog.processedInvoices} of ${multiLog.totalInvoices} invoices processed`
    };
  }
  
  return { isComplete: true, reason: 'All invoices processed successfully' };
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

            // UI will be handled by logFileToMainSheet function which gets it from buffer sheet
            
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
              
              // UI will be handled by logFileToMainSheet function which gets it from buffer sheet
              
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
                logFileToMainSheet(mainSheet, copiedFile, emailSubject, gmailMessageId, invoiceStatus, companyName);
                const flowSheet = (invoiceStatus === "inflow") ? inflowSheet : outflowSheet;
                logFileToMainSheet(flowSheet, copiedFile, emailSubject, gmailMessageId, invoiceStatus, companyName);
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
                
                // UI will be handled by logFileToMainSheet function which gets it from buffer sheet
                
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
                  logFileToMainSheet(mainSheet, copiedFile, emailSubject, gmailMessageId, invoiceStatus, companyName);
                  const flowSheet = (invoiceStatus === "inflow") ? inflowSheet : outflowSheet;
                  logFileToMainSheet(flowSheet, copiedFile, emailSubject, gmailMessageId, invoiceStatus, companyName);
                } else {
                  // Log in main sheet only
                  logFileToMainSheet(mainSheet, file, emailSubject, gmailMessageId, invoiceStatus, companyName);
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

/**
 * Helper function to populate missing UI values in existing sheets
 * Run this once to populate UI values for existing entries
 * @param {string} companyName - The company name (e.g., 'analogy', 'humane')
 */
function populateMissingUIValues(companyName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get all relevant sheets
  const bufferSheet = ss.getSheetByName(`${companyName}-buffer`);
  const mainSheet = ss.getSheetByName(companyName);
  const inflowSheet = ss.getSheetByName(`${companyName}-inflow`);
  const outflowSheet = ss.getSheetByName(`${companyName}-outflow`);
  
  if (!bufferSheet) {
    Logger.log(`Buffer sheet not found for company: ${companyName}`);
    return;
  }
  
  // First, populate buffer sheet UI values if missing
  const bufferData = bufferSheet.getDataRange().getValues();
  for (let i = 1; i < bufferData.length; i++) {
    const row = bufferData[i];
    const driveFileId = row[3];
    const currentUI = row[7];
    
    if (driveFileId && !currentUI) {
      const newUI = generateUniqueIdentifierForFile(driveFileId);
      bufferSheet.getRange(i + 1, 8).setValue(newUI);
      Logger.log(`Added UI '${newUI}' to buffer sheet row ${i + 1}`);
    }
  }
  
  // Then populate other sheets by looking up UI from buffer sheet
  const sheets = [mainSheet, inflowSheet, outflowSheet].filter(sheet => sheet !== null);
  
  sheets.forEach(sheet => {
    if (sheet.getLastRow() <= 1) return; // Skip if only headers or empty
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const uiColumnIndex = headers.indexOf('UI');
    const fileIdColumnIndex = headers.indexOf('File ID');
    const fileNameColumnIndex = headers.indexOf('File Name');
    
    if (uiColumnIndex === -1 || fileIdColumnIndex === -1) {
      Logger.log(`Required columns not found in sheet: ${sheet.getName()}`);
      return;
    }
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const fileId = row[fileIdColumnIndex];
      const fileName = row[fileNameColumnIndex];
      const currentUI = row[uiColumnIndex];
      
      if (fileId && !currentUI) {
        const ui = getUIFromBufferSheet(companyName, fileName, fileId);
        if (ui) {
          sheet.getRange(i + 1, uiColumnIndex + 1).setValue(ui);
          Logger.log(`Added UI '${ui}' to ${sheet.getName()} row ${i + 1}`);
        }
      }
    }
  });
  
  Logger.log(`Completed populating missing UI values for company: ${companyName}`);
}

// Helper to set dropdown for 'Relevance' column in Buffer2
function setRelevanceDropdownValidation(sheet) {
  const relevanceCol = BUFFER2_SHEET_HEADERS.indexOf('Relevance') + 1;
  const range = sheet.getRange(2, relevanceCol, sheet.getMaxRows() - 1); // Exclude header
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build();
  range.setDataValidation(rule);
  Logger.log(`Relevance dropdown validation applied to ${sheet.getName()}`);
}

// Helper to sort a sheet by a date column (descending, latest first)
function sortSheetByDateDesc(sheet, dateCol) {
  if (!sheet || sheet.getLastRow() < 2) return;
  // Sort range: all rows except header
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
    .sort({column: dateCol, ascending: false});
}

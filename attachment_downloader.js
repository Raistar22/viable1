//attachment_downloader.js code working with buffer1 buffer2 classification ,invoice generation date time, vendor date you have to combine the vendor name and date for renaming the file only at this point 

//attachment_donwloader.js till 17-07-25

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

// Parent folder where all company folders are stored
var PARENT_COMPANIES_FOLDER_ID = "1Tqj8May8je0L1lET5PIHRJj4P8d9T3MT";

// Company folder mappings - new companies will be auto-created in parent folder
var ATTACHMENT_COMPANY_FOLDER_MAP = {
  "analogy": "160pN2zDCb9UQbwIXqgggdTjLUrFM2cM3", // Main Analogy Folder
  "humane": "1E6ijhWhdYykymN0MEUINd9jETmdM2sAt"   // Main Humane Folder
  // New companies will be automatically added here when first accessed
};

// Initialize company folder mappings on script load
(function initializeCompanyMappings() {
  try {
    loadCompanyFolderMappings();
    Logger.log('Company folder mappings initialized successfully.');
  } catch (error) {
    Logger.log(`Warning: Could not initialize company folder mappings: ${error.toString()}`);
  }
})();

// --------------------------------------------------------

// --- Sheet Header Definitions ---
const BUFFER_SHEET_HEADERS = ['Date', 'OriginalFileName', 'ChangedFilename', 'Invoice ID', 'Drive File ID', 'Gmail Message ID', 'Reason', 'Status', 'UI', 'Repeated', 'Invoice Count', 'Attachment ID', 'Email ID', 'Vendor Name'];
const BUFFER2_SHEET_HEADERS = [
  'Date',
  'OriginalFileName',
  'ChangedFilename',
  'Invoice ID',
  'Drive File ID',
  'Gmail Message ID',
  'Relevance',
  'UI',
  'Vendor Name'
];
const MAIN_SHEET_HEADERS = [
  'File Name', 'File ID', 'File URL',
  'Date Created (Drive)', 'Last Updated (Drive)', 'Size (bytes)', 'Mime Type',
  'Email Subject', 'Gmail Message ID', 'invoice status', 'UI',
  'Date', 'Month', 'FY', 'GST', 'TDS', 'OT', 'NA', 'Vendor Name'
];

// Global counter for unique identifiers
var uniqueIdentifierCounter = 0;

// Global object to store file ID to unique identifier mappings
var FILE_IDENTIFIER_MAP = {};

// Edge case tracking objects
var PROCESSED_EMAILS_LOG = {}; // Track emails by message ID
var ATTACHMENT_PROCESSING_LOG = {}; // Track attachments by message ID + attachment name
var ERROR_RECOVERY_LOG = {}; // Track failed operations for retry
var THREAD_CONTEXT_LOG = {}; // Track thread-level information for sender recognition (supplementary)


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
    const bufferChangedFilename = row[2]; // ChangedFilename column (index 2) - THE SOURCE OF TRUTH
    const bufferDriveFileId = row[4]; // Drive File ID column (index 4)
    const bufferUI = row[8]; // UI column (index 8)
    
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
 * Automatically creates a company folder structure in the parent companies folder
 * @param {string} companyName - The name of the company
 * @returns {string} The ID of the created company folder
 */
function createCompanyFolderStructure(companyName) {
  try {
    Logger.log(`Creating folder structure for new company: ${companyName}`);
    
    // Get the parent companies folder
    const parentFolder = DriveApp.getFolderById(PARENT_COMPANIES_FOLDER_ID);
    Logger.log(`Parent folder found: ${parentFolder.getName()}`);
    
    // Check if company folder already exists
    const existingFolders = parentFolder.getFoldersByName(companyName);
    let companyFolder;
    
    if (existingFolders.hasNext()) {
      // Company folder already exists
      companyFolder = existingFolders.next();
      Logger.log(`Company folder already exists: ${companyFolder.getName()} (ID: ${companyFolder.getId()})`);
    } else {
      // Create new company folder
      companyFolder = parentFolder.createFolder(companyName);
      Logger.log(`Created new company folder: ${companyFolder.getName()} (ID: ${companyFolder.getId()})`);
    }
    
    // Update the company folder mapping
    ATTACHMENT_COMPANY_FOLDER_MAP[companyName] = companyFolder.getId();
    
    // Store the mapping in script properties for persistence
    const props = PropertiesService.getScriptProperties();
    props.setProperty(`COMPANY_FOLDER_${companyName.toUpperCase()}`, companyFolder.getId());
    
    Logger.log(`Company folder mapping updated for ${companyName}: ${companyFolder.getId()}`);
    
    return companyFolder.getId();
    
  } catch (error) {
    Logger.log(`Error creating company folder structure for ${companyName}: ${error.toString()}`);
    throw new Error(`Failed to create company folder structure: ${error.message}`);
  }
}

/**
 * Loads company folder mappings from script properties (for persistence across runs)
 */
function loadCompanyFolderMappings() {
  try {
    const props = PropertiesService.getScriptProperties();
    const allProperties = props.getProperties();
    
    // Load any stored company folder mappings
    Object.keys(allProperties).forEach(key => {
      if (key.startsWith('COMPANY_FOLDER_')) {
        const companyName = key.replace('COMPANY_FOLDER_', '').toLowerCase();
        const folderId = allProperties[key];
        
        // Only add if not already in the mapping
        if (!ATTACHMENT_COMPANY_FOLDER_MAP[companyName]) {
          ATTACHMENT_COMPANY_FOLDER_MAP[companyName] = folderId;
          Logger.log(`Loaded company mapping from storage: ${companyName} -> ${folderId}`);
        }
      }
    });
  } catch (error) {
    Logger.log(`Error loading company folder mappings: ${error.toString()}`);
  }
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
 * Create folder structure: Company/FY-XX-XX/Accruals/Buffer2 (for non-invoice files and unknown types)
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
      const gmailId = bufferData[i][5]; // Gmail Message ID column (index 5)
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
/**
 * Generates a new filename with robust defaults to avoid errors.
 * @param {Object} extractedData Extracted data.
 * @param {string} originalFileName Original filename.
 * @returns {string} New filename.
 */
function generateNewFilename(extractedData, originalFileName) {
  const extension = originalFileName.split('.').pop() || 'pdf';
  
  // Use extracted or safe defaults
  const date = extractedData.date || getCurrentDateString();
  const vendor = extractedData.vendorName || extractVendorFromFilename(originalFileName) || 'UnknownVendor';
  const invoice = extractedData.invoiceNumber || extractInvoiceIdFromFilename(originalFileName) || generateInvoiceNumber();
  const amount = extractedData.amount || '0.00';

  // Sanitize (same as before, but ensure no empties)
  const sanitizedDate = date.replace(/[^0-9-]/g, '') || getCurrentDateString();
  const sanitizedVendor = vendor.replace(/[^a-zA-Z0-9\s-]/g, '').trim() || 'UnknownVendor';
  const sanitizedInvoice = invoice.replace(/[^a-zA-Z0-9-]/g, '').toUpperCase() || 'INV-UNKNOWN';
  const sanitizedAmount = amount.replace(/[^0-9.]/g, '') || '0.00';

  const newFilename = `${sanitizedDate}_${sanitizedVendor}_${sanitizedInvoice}_${sanitizedAmount}.${extension}`;
  Logger.log(`Generated robust filename: ${newFilename}`);
  return newFilename;
}

/**
 * Enhanced fallback filename generation using meaningful words from original filename
 * @param {string} originalFileName Original filename
 * @returns {string} New filename using meaningful words
 */
function generateFallbackFilenameFromOriginal(originalFileName) {
  Logger.log(`Generating fallback filename from original: ${originalFileName}`);
  
  const extension = originalFileName.split('.').pop() || 'pdf';
  const baseName = originalFileName.replace(/\.[^/.]+$/, ''); // Remove extension
  
  // Extract meaningful words (exclude common meaningless words)
  const meaninglessWords = ['the', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for', 'of', 'with', 'by', 'from', 'up', 'about', 'into', 'through', 'during', 'before', 'after', 'above', 'below', 'between', 'among', 'file', 'document', 'doc', 'img', 'image', 'scan', 'copy', 'attachment', 'att', 'untitled', 'new', 'temp', 'tmp'];
  
  // Split by common separators and filter meaningful words
  const words = baseName
    .toLowerCase()
    .split(/[\s\-_\.\(\)\[\]]+/)
    .filter(word => word.length > 2 && !meaninglessWords.includes(word))
    .filter(word => !/^[0-9]+$/.test(word) || word.length >= 4) // Keep numbers if they're 4+ digits (years, invoice numbers)
    .slice(0, 4); // Take first 4 meaningful words
  
  // Extract potential date from filename
  const dateMatch = baseName.match(/(\d{4}[-_]\d{2}[-_]\d{2}|\d{2}[-_]\d{2}[-_]\d{4}|\d{8})/);
  const extractedDate = dateMatch ? standardizeDateFormat(dateMatch[1]) : getCurrentDateString();
  
  // Extract potential amount from filename
  const amountMatch = baseName.match(/(\d+\.\d{2}|\d+[\-_]?\d{2})/);
  const extractedAmount = amountMatch ? amountMatch[1].replace(/[\-_]/g, '.') : '0.00';
  
  // Build filename with meaningful components
  let meaningfulParts = [];
  
  // Add date
  meaningfulParts.push(extractedDate.replace(/[^0-9-]/g, ''));
  
  // Add meaningful words as vendor/description
  if (words.length > 0) {
    const vendorPart = words.join('_').replace(/[^a-zA-Z0-9_]/g, '').substring(0, 20);
    meaningfulParts.push(vendorPart);
  } else {
    meaningfulParts.push('Document');
  }
  
  // Add invoice number if found, otherwise generate one
  const invoiceMatch = baseName.match(/\b(inv|invoice|bill|ref|no)[\-_]?([a-zA-Z0-9\-]+)\b/i);
  if (invoiceMatch) {
    meaningfulParts.push(invoiceMatch[2].toUpperCase());
  } else {
    meaningfulParts.push(generateInvoiceNumber());
  }
  
  // Add amount
  meaningfulParts.push(extractedAmount);
  
  const fallbackFilename = `${meaningfulParts.join('_')}.${extension}`;
  Logger.log(`Generated fallback filename: ${fallbackFilename}`);
  return fallbackFilename;
}

/**
 * Retrieves all existing 'ChangedFilename' values from the buffer sheet with row tracking.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} bufferSheet The buffer sheet.
 * @returns {Object} Object with filename set and row mapping.
 */
function getExistingChangedFilenames(bufferSheet) {
  const existingChangedFilenames = new Set();
  const filenameToRowMap = new Map(); // Maps filename to row number
  const data = bufferSheet.getDataRange().getValues();
  
  // ChangedFilename is in column 3 (index 2) - THE SOURCE OF TRUTH
  for (let i = 1; i < data.length; i++) { // Start from 1 to skip header row
    const changedFilename = data[i][2]; // Correct index for ChangedFilename
    if (changedFilename) {
      existingChangedFilenames.add(changedFilename);
      if (!filenameToRowMap.has(changedFilename)) {
        filenameToRowMap.set(changedFilename, i + 1); // Store 1-based row number
      }
    }
  }
  
  return {
    filenames: existingChangedFilenames,
    rowMap: filenameToRowMap
  };
}

/**
 * Colors duplicate files and updates the Repeated column with reference row numbers
 * @param {GoogleAppsScript.Spreadsheet.Sheet} bufferSheet The buffer sheet
 * @param {string} duplicateFilename The filename that's duplicated
 * @param {number} newRowIndex The row index of the new duplicate entry
 * @param {number} originalRowIndex The row index of the original file
 */
function handleDuplicateFileColoring(bufferSheet, duplicateFilename, newRowIndex, originalRowIndex) {
  try {
    // Color both the original and duplicate rows
    const duplicateColor = '#FFE6E6'; // Light red for duplicates
    const originalColor = '#E6F3FF'; // Light blue for original
    
    // Color the new duplicate row
    bufferSheet.getRange(newRowIndex, 1, 1, BUFFER_SHEET_HEADERS.length).setBackground(duplicateColor);
    
    // Color the original row
    bufferSheet.getRange(originalRowIndex, 1, 1, BUFFER_SHEET_HEADERS.length).setBackground(originalColor);
    
    // Update the Repeated column (index 9) for the duplicate row
    const repeatedColumnIndex = BUFFER_SHEET_HEADERS.indexOf('Repeated') + 1; // Convert to 1-based
    bufferSheet.getRange(newRowIndex, repeatedColumnIndex).setValue(`Row ${originalRowIndex}`);
    
    // Update the Repeated column for the original row to show it has duplicates
    bufferSheet.getRange(originalRowIndex, repeatedColumnIndex).setValue(`Duplicated in Row ${newRowIndex}`);
    
    Logger.log(`Duplicate file handling: ${duplicateFilename} - Original: Row ${originalRowIndex}, Duplicate: Row ${newRowIndex}`);
    
  } catch (error) {
    Logger.log(`Error handling duplicate file coloring: ${error.toString()}`);
  }
}

/**
 * Sets up data validation for the Status dropdown in buffer sheets.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The buffer sheet to apply validation to.
 */
function setStatusDropdownValidation(sheet) {
  const range = sheet.getRange("H:H"); // Status column is G (index 6, 0-indexed)
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
 * Logs a file's details to a specified log sheet with UI and vendor name.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} logSheet The sheet to log to.
 * @param {GoogleAppsScript.Drive.File} driveFile The Drive file object.
 * @param {string} emailSubject The subject of the original email.
 * @param {string} gmailMessageId The ID of the original Gmail message.
 * @param {string} invoiceStatus The determined invoice status (inflow, outflow, unknown).
 * @param {string} companyName The company name to get UI from buffer sheet.
 * @param {string} providedUI Optional UI to use instead of looking up from buffer sheet.
 * @param {Object} extraData Additional data including vendor name.
 */
function logFileToMainSheet(logSheet, driveFile, emailSubject, gmailMessageId, invoiceStatus, companyName, providedUI = null, extraData = {}) {
  // Ensure headers are present in the first row
  const firstRow = logSheet.getRange(1, 1, 1, MAIN_SHEET_HEADERS.length).getValues()[0];
  if (firstRow.join(',') !== MAIN_SHEET_HEADERS.join(',')) {
    logSheet.clear();
    logSheet.appendRow(MAIN_SHEET_HEADERS);
    logSheet.getRange(1, 1, 1, MAIN_SHEET_HEADERS.length)
      .setFontWeight('bold').setBackground('#E8F0FE').setBorder(true, true, true, true, true, true);
    logSheet.setFrozenRows(1);
  }
  let ui = providedUI;
  if (!ui) {
    ui = getUIFromBufferSheet(companyName, driveFile.getName(), driveFile.getId());
  }
  // Extract extra fields including vendor name
  const {
    date = '',
    month = '',
    fy = '',
    gst = '',
    tds = '',
    ot = '',
    na = '',
    vendorName = 'Unknown'
  } = extraData;
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
    ui,
    date,
    month,
    fy,
    gst,
    tds,
    ot,
    na,
    vendorName
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
  const bufferLabelName = `${companyName}-buffer`;

  // Load existing company folder mappings from script properties
  loadCompanyFolderMappings();

  // Check if the company exists in our folder mapping, create if not
  let isNewCompany = false;
  if (!ATTACHMENT_COMPANY_FOLDER_MAP[companyName]) {
    Logger.log(`Company '${companyName}' not found in folder mapping. Creating new folder structure...`);
    
    try {
      const newCompanyFolderId = createCompanyFolderStructure(companyName);
      Logger.log(`Successfully created company folder structure for '${companyName}' with ID: ${newCompanyFolderId}`);
      isNewCompany = true;
    } catch (createError) {
      const availableCompanies = Object.keys(ATTACHMENT_COMPANY_FOLDER_MAP).join(', ');
      return { 
        status: 'error', 
        message: `Error: Failed to create folder structure for company '${companyName}': ${createError.message}. Available companies: ${availableCompanies}. Please check the parent folder permissions.` 
      };
    }
  }

  const companyFolder = DriveApp.getFolderById(ATTACHMENT_COMPANY_FOLDER_MAP[companyName]);

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
      bufferSheet.setColumnWidth(9, 120); // Repeated
      bufferSheet.setColumnWidth(10, 120); // Invoice Count
      bufferSheet.setColumnWidth(11, 120); // Attachment ID
      bufferSheet.setColumnWidth(12, 120); // Email ID
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
    const existingFilenamesData = getExistingChangedFilenames(bufferSheet);
    const existingChangedFilenamesInCurrentBuffer = existingFilenamesData.filenames;
    const filenameToRowMap = existingFilenamesData.rowMap;
    let totalNewAttachments = 0;
    let processedAttachments = 0;
    let skippedAttachments = 0;
    let emailsAnalyzed = 0;
    let emailsSkipped = 0;
    
    // Clear processing logs for this session
    PROCESSED_EMAILS_LOG = {};
    ATTACHMENT_PROCESSING_LOG = {};
    ERROR_RECOVERY_LOG = {};
    THREAD_CONTEXT_LOG = {}; // Clear supplementary thread context

    // First pass: Enhanced email and attachment analysis with comprehensive tracking
    const threads = gmailLabel.getThreads();
    Logger.log(`Starting analysis of ${threads.length} email threads for label: ${labelName}`);
    
    for (let t = 0; t < threads.length; t++) {
      if (shouldCancel(processToken)) {
        return { status: 'cancelled', message: "Process cancelled during attachment counting." };
      }
      
      // Extract thread context for analysis (supplementary, doesn't affect main logic)
      const currentThread = threads[t];
      const threadContext = getThreadContext(currentThread);
      logThreadContext(threadContext.threadId, threadContext, 'analyzing');
      
      const messages = currentThread.getMessages();
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

      // Get thread context for this processing iteration (supplementary)
      const currentThread = threads[t];
      const threadContext = getThreadContext(currentThread);
      
      const messages = currentThread.getMessages();
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
                // Update thread stats (supplementary)
                updateThreadStats(threadContext.threadId, 'skipped');
                skippedAttachments++;
                attachmentsProcessedForThisMessage++;
                continue;
              }
              
              // Use AI to extract comprehensive data for proper renaming
              const blob = attachment.copyBlob();
              const aiExtractedData = callGeminiAPIInternal(blob, originalFilename);
              
              Logger.log(`=== AI PROCESSING FOR ${originalFilename} ===`);
              Logger.log(`AI extracted data: ${JSON.stringify(aiExtractedData)}`);
              Logger.log(`AI extracted invoice status: ${aiExtractedData.invoiceStatus}`);
              
              // Generate new filename using AI-extracted data
              const changedFilename = generateNewFilename(aiExtractedData, originalFilename);
              Logger.log(`Generated new filename: "${originalFilename}" → "${changedFilename}"`);
              
              // Store AI-extracted data for later use
              const extractedInvoiceStatus = aiExtractedData.invoiceStatus || 'outflow';

              // Check if this is explicitly marked as a non-invoice/bill file or unknown type
              // Files marked as 'irrelevant', 'not_relevant', or 'unknown' should go to Buffer2
              // Only clear 'inflow' and 'outflow' files go to regular buffer
              if (extractedInvoiceStatus === 'irrelevant' || extractedInvoiceStatus === 'not_relevant' || extractedInvoiceStatus === 'unknown') {
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
                  
                  // Update thread stats (supplementary)
                  updateThreadStats(threadContext.threadId, 'processed');
                  
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
                  // Use invoice date from AI extraction, fallback to current date
                  const buffer2InvoiceDate = aiExtractedData.date ? new Date(aiExtractedData.date) : new Date();
                  
                  buffer2Sheet.appendRow([
                    buffer2InvoiceDate,                   // Invoice Date (from AI extraction)
                    originalFilename,
                    changedFilename,
                    aiExtractedData.invoiceNumber || 'INV-Unknown',
                    buffer2File.getId(), // Use actual Buffer2 file ID
                    messageId,
                    '', // Relevance blank by default
                    '',  // UI blank by default
                    aiExtractedData.vendorName || 'Unknown' // Vendor Name
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
              let originalRowIndex = null;

              if (isDuplicateChangedFilename) {
                originalRowIndex = filenameToRowMap.get(changedFilename);
                Logger.log(`Duplicate filename detected: ${changedFilename} (original in row ${originalRowIndex})`);
              }

              if (!isDuplicateChangedFilename) {
                // Store in Buffer/Active with the correct name
                const now = new Date();
                const month = getMonthFromDate(now);
                const financialYear = calculateFinancialYear(now);
                const bufferActiveFolder = getOrCreateBufferSubfolder(companyName, financialYear, "Active");

                try {
                  Logger.log(`Creating file in buffer folder with name: "${changedFilename}"`);
                  
                  // Create the file with the NEW AI-generated filename
                  const renamedBlob = attachment.copyBlob().setName(changedFilename);
                  driveFile = bufferActiveFolder.createFile(renamedBlob);
                  
                  Logger.log(`✓ File created successfully: "${driveFile.getName()}" (ID: ${driveFile.getId()})`);
                  
                  // Verify the file was created with correct name
                  if (driveFile.getName() !== changedFilename) {
                    Logger.log(`Warning: File created with different name. Expected: "${changedFilename}", Actual: "${driveFile.getName()}"`);
                    // Rename the file to ensure it has the correct name
                    driveFile.setName(changedFilename);
                    Logger.log(`✓ File renamed to: "${changedFilename}"`);
                  }
                  
                  // Generate unique identifier for the file - IMMEDIATELY after file creation
                  uniqueIdentifier = generateUniqueIdentifierForFile(driveFile.getId());
                  
                  processedAttachments++;
                  attachmentsProcessedForThisMessage++;
                  existingChangedFilenamesInCurrentBuffer.add(changedFilename);
                  
                  // Track successful processing
                  trackAttachmentProcessing(messageId, originalFilename, 'processed', `Successfully renamed and saved: ${originalFilename} → ${changedFilename}`);
                  
                  // Update thread stats (supplementary)
                  updateThreadStats(threadContext.threadId, 'processed');
                  
                  Logger.log(`✓ File processing completed: ${changedFilename} with UI: ${uniqueIdentifier}`);
                  
                } catch (fileCreationError) {
                  Logger.log(`✗ Error creating file for attachment ${originalFilename}: ${fileCreationError.toString()}`);
                  logProcessingError(messageId, originalFilename, fileCreationError.toString(), 'file_creation');
                  trackAttachmentProcessing(messageId, originalFilename, 'failed', `File creation failed: ${fileCreationError.message}`);
                  continue; // Skip to next attachment
                }

                              // Use previously extracted AI data for invoice status
              const emailSubject = message.getSubject ? message.getSubject() : '';
              const invoiceStatus = extractedInvoiceStatus || "unknown";
              
              // Enhanced debug logging for invoice status determination
              Logger.log(`\n=== CLASSIFICATION DECISION FOR ${changedFilename} ===`);
              Logger.log(`Original filename: ${originalFilename}`);
              Logger.log(`New filename: ${changedFilename}`);
              Logger.log(`AI extracted status: ${extractedInvoiceStatus}`);
              Logger.log(`Final status: ${invoiceStatus}`);
              Logger.log(`Document type: ${aiExtractedData.documentType || 'unknown'}`);
              Logger.log(`Vendor: ${aiExtractedData.vendorName || 'unknown'}`);
              Logger.log(`Invoice number: ${aiExtractedData.invoiceNumber || 'unknown'}`);
              Logger.log(`Amount: ${aiExtractedData.amount || 'unknown'}`);
              Logger.log(`Is financial document: ${aiExtractedData.isFinancialDocument}`);
              
              if (invoiceStatus === 'irrelevant' || invoiceStatus === 'not_relevant' || invoiceStatus === 'unknown') {
                Logger.log(`✓ ROUTING: Buffer2 (NON-FINANCIAL/IRRELEVANT/UNKNOWN document)`);
                Logger.log(`  → Will be saved to Buffer2 folder for manual review`);
              } else if (invoiceStatus === 'inflow') {
                Logger.log(`✓ ROUTING: Regular Buffer → Inflow folder (MONEY COMING IN)`);
                Logger.log(`  → This appears to be a sales invoice or payment received`);
              } else if (invoiceStatus === 'outflow') {
                Logger.log(`✓ ROUTING: Regular Buffer → Outflow folder (MONEY GOING OUT)`);
                Logger.log(`  → This appears to be a purchase invoice or expense`);
              } else {
                Logger.log(`✓ ROUTING: Regular Buffer → Manual review needed (UNCLEAR STATUS)`);
                Logger.log(`  → Could not determine if inflow or outflow`);
              }
              Logger.log(`=== END CLASSIFICATION ===\n`);

                // Log to main sheet
                let mainSheet = ss.getSheetByName(companyName);
                if (!mainSheet) {
                  mainSheet = ss.insertSheet(companyName);
                  mainSheet.appendRow(MAIN_SHEET_HEADERS);
                }
                // Use invoice date from AI extraction, fallback to current date
                const mainSheetInvoiceDate = aiExtractedData.date ? new Date(aiExtractedData.date) : now;
                
                logFileToMainSheet(mainSheet, driveFile, emailSubject, messageId, invoiceStatus, companyName, uniqueIdentifier, {
                  date: mainSheetInvoiceDate.toISOString().split('T')[0],
                  month: getMonthFromDate(mainSheetInvoiceDate),
                  fy: calculateFinancialYear(mainSheetInvoiceDate),
                  gst: aiExtractedData.gst || '',
                  tds: aiExtractedData.tds || '',
                  ot: aiExtractedData.ot || '',
                  na: aiExtractedData.na || '',
                  vendorName: aiExtractedData.vendorName || 'Unknown'
                });

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
                    // Use invoice date from AI extraction, fallback to current date
                    const flowSheetInvoiceDate = aiExtractedData.date ? new Date(aiExtractedData.date) : now;
                    
                    logFileToMainSheet(flowSheet, flowFile, emailSubject, messageId, invoiceStatus, companyName, uniqueIdentifier, {
                      date: flowSheetInvoiceDate.toISOString().split('T')[0],
                      month: getMonthFromDate(flowSheetInvoiceDate),
                      fy: calculateFinancialYear(flowSheetInvoiceDate),
                      gst: aiExtractedData.gst || '',
                      tds: aiExtractedData.tds || '',
                      ot: aiExtractedData.ot || '',
                      na: aiExtractedData.na || '',
                      vendorName: aiExtractedData.vendorName || 'Unknown'
                    });
                  }
                }
              } else {
                // If duplicate, don't upload again, but still generate unique identifier for buffer logging
                uniqueIdentifier = generateUniqueIdentifierForFile(`duplicate_${messageId}_${a}`);
                Logger.log(`Skipping upload: Duplicate changed filename '${changedFilename}' already exists in buffer folder for '${labelName}'. Generated UI: ${uniqueIdentifier}`);
                
                // Track duplicate detection
                trackAttachmentProcessing(messageId, originalFilename, 'skipped', `Duplicate filename: ${changedFilename}`);
                // Update thread stats (supplementary)
                updateThreadStats(threadContext.threadId, 'skipped');
                skippedAttachments++;
                attachmentsProcessedForThisMessage++;
              }

              // Append to buffer sheet with comprehensive AI-extracted data
              const aiReason = `AI Classification: ${extractedInvoiceStatus.toUpperCase()} | Type: ${aiExtractedData.documentType || 'document'} | Vendor: ${aiExtractedData.vendorName || 'Unknown'} | Amount: ${aiExtractedData.amount || '0.00'}`;
              
              // Use invoice date from AI extraction, fallback to current date
              const invoiceDate = aiExtractedData.date ? new Date(aiExtractedData.date) : new Date();
              
              const rowData = [
                invoiceDate,                             // Invoice Date (from AI extraction)
                originalFilename,                        // Original filename
                changedFilename,                         // NEW AI-generated filename
                aiExtractedData.invoiceNumber || generateInvoiceNumber(), // AI-extracted or generated invoice number
                driveFile ? driveFile.getId() : '',      // Drive File ID
                messageId,                               // Gmail Message ID
                aiReason,                                // Comprehensive reason with AI classification
                'Active',                                // Default Status (Active)
                uniqueIdentifier,                        // UI (unique identifier)
                '',                                      // Repeated field (will be updated for duplicates)
                aiExtractedData.invoiceCount || 1,       // Invoice count
                attachment.getId ? attachment.getId() : '', // Attachment ID
                messageId,                               // Email ID
                aiExtractedData.vendorName || 'Unknown'  // Vendor Name
              ];
              
              Logger.log(`Buffer sheet row data: ${JSON.stringify(rowData)}`);
              bufferSheet.appendRow(rowData);
              sortSheetByDateDesc(bufferSheet, 1);
              const newRowIndex = bufferSheet.getLastRow();

              // Enhanced duplicate handling with coloring and row references
              if (isDuplicateChangedFilename && originalRowIndex) {
                handleDuplicateFileColoring(bufferSheet, changedFilename, newRowIndex, originalRowIndex);
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
    
    // Add new company creation message if applicable
    if (isNewCompany) {
      resultMessage += `🆕 New company folder structure created for '${companyName}' in parent directory. `;
    }
    
    resultMessage += `Emails analyzed: ${emailsAnalyzed}, New emails: ${emailsAnalyzed - emailsSkipped}. `;
    resultMessage += `Attachments processed: ${processedAttachments}, Skipped: ${skippedAttachments}. `;
    
    // Add supplementary thread insights to message
    const threadSummary = getThreadSummary();
    if (threadSummary.totalThreads > 0) {
      resultMessage += `📧 Thread recognition: ${threadSummary.totalThreads} conversations analyzed, ${threadSummary.singleSenderThreads} single-sender threads detected. `;
    }
    
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
    
    // Log supplementary thread analysis
    const detailedThreadSummary = getThreadSummary();
    if (detailedThreadSummary.totalThreads > 0) {
      Logger.log(`\n=== THREAD RECOGNITION ANALYSIS (Supplementary) ===`);
      Logger.log(`- Gmail threads processed: ${detailedThreadSummary.totalThreads}`);
      Logger.log(`- Single-sender threads: ${detailedThreadSummary.singleSenderThreads}`);
      Logger.log(`- Recurring conversations: ${detailedThreadSummary.recurringConversations}`);
      Logger.log(`- Top senders by attachments:`);
      detailedThreadSummary.topSenders.forEach((sender, index) => {
        Logger.log(`  ${index + 1}. ${sender.name}: ${sender.attachments} attachments, ${sender.emails} emails`);
      });
      Logger.log(`=== END THREAD ANALYSIS ===\n`);
    }
    
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

  // Load existing company folder mappings from script properties
  loadCompanyFolderMappings();

  // Ensure company folder exists (create if needed)
  if (!ATTACHMENT_COMPANY_FOLDER_MAP[companyName]) {
    Logger.log(`Company '${companyName}' not found in folder mapping. Creating new folder structure...`);
    
    try {
      const newCompanyFolderId = createCompanyFolderStructure(companyName);
      Logger.log(`Successfully created company folder structure for '${companyName}' with ID: ${newCompanyFolderId}`);
    } catch (createError) {
      Logger.log(`Error: Failed to create folder structure for company '${companyName}': ${createError.message}`);
      return;
    }
  }

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
    const originalFilename = row[1]; // OriginalFileName (column 2, index 1)
    const changedFilename = row[2];   // ChangedFilename (column 3, index 2) - THE SOURCE OF TRUTH
    let driveFileId = row[4];         // Drive File ID (column 5, index 4)
    const gmailMessageId = row[5];    // Gmail Message ID (column 6, index 5)
    const status = row[7];            // Status (column 8, index 7)
    const existingAnimalName = row[8]; // UI (column 9, index 8)
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
        // Use invoice date from AI extraction, fallback to file creation date
        const processBufferInvoiceDate = aiResult.date ? new Date(aiResult.date) : fileDate;
        
        logFileToMainSheet(mainSheet, file, emailSubject, gmailMessageId, invoiceStatus, companyName, uniqueIdentifier, {
          date: processBufferInvoiceDate.toISOString().split('T')[0],
          month: getMonthFromDate(processBufferInvoiceDate),
          fy: calculateFinancialYear(processBufferInvoiceDate),
          gst: aiResult.gst || '',
          tds: aiResult.tds || '',
          ot: aiResult.ot || '',
          na: aiResult.na || '',
          vendorName: aiResult.vendorName || 'Unknown'
        });

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
          // Use invoice date from AI extraction, fallback to file creation date
          const flowBufferInvoiceDate = aiResult.date ? new Date(aiResult.date) : fileDate;
          
          logFileToMainSheet(flowSheet, copiedFile, emailSubject, gmailMessageId, invoiceStatus, companyName, uniqueIdentifier, {
            date: flowBufferInvoiceDate.toISOString().split('T')[0],
            month: getMonthFromDate(flowBufferInvoiceDate),
            fy: calculateFinancialYear(flowBufferInvoiceDate),
            gst: aiResult.gst || '',
            tds: aiResult.tds || '',
            ot: aiResult.ot || '',
            na: aiResult.na || '',
            vendorName: aiResult.vendorName || 'Unknown'
          });
        }

        // Clear any previous "Reason" or yellow background if successfully processed as Active
        bufferSheet.getRange(bufferRowIndex, 6).setValue('');
        bufferSheet.getRange(bufferRowIndex, 1, 1, BUFFER_SHEET_HEADERS.length).setBackground(null); // Remove background color


      } catch (e) {
        Logger.log(`Error processing buffer row for file ID ${driveFileId} (Row ${bufferRowIndex}): ${e.toString()}`);
        ui.alert('Error', `Could not process file "${changedFilename}" from buffer (Row ${bufferRowIndex}): ${e.message}`, ui.ButtonSet.OK);
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
 * Enhanced Gemini AI integration with retries and text fallback for robust extraction.
 * @param {GoogleAppsScript.Base.Blob} fileBlob The content of the file.
 * @param {string} fileName The name of the file.
 * @returns {Object} Extracted data.
 */
function callGeminiAPIInternal(fileBlob, fileName) {
  Logger.log(`Calling enhanced Gemini AI for: ${fileName}`);
  
  const API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!API_KEY) {
    Logger.log('Warning: GEMINI_API_KEY not found. Using fallback.');
    return fallbackDataExtraction(fileBlob, fileName);
  }
  
  const GEMINI_URL = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${API_KEY}`;
  const maxRetries = 2;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const imageData = Utilities.base64Encode(fileBlob.getBytes());
      const mimeType = fileBlob.getContentType();
      
      // Enhanced prompt: Emphasize extraction even if partial, with fallbacks
      const prompt = `Analyze this document and extract key information. Be robust: If it's not a clear invoice/bill, still attempt to extract any visible date, vendor, number, and amount. Classify strictly but provide best-guess data.

STRICT RULES:
- If partial data (e.g., missing invoice number but has amount/date/vendor), classify as "irrelevant" but STILL extract available fields.
- For unclear docs, use filename hints if needed.
- NEVER leave fields empty - use "NA" or "0.00" as defaults.

CLASSIFICATION RULES:
1. INVOICE/BILL CRITERIA (MUST HAVE ALL for inflow/outflow):
   - Explicit "Invoice" or "Bill" label.
   - Invoice/Reference number.
   - Total billed amount.
   - Date (issue or due).
   - Billed from/to parties (vendor/customer).

2. FINANCIAL but NON-INVOICE → "irrelevant".
3. NON-FINANCIAL → "irrelevant".
4. INFLOW/OUTFLOW (only for true invoices/bills):
   - INFLOW: You are the seller.
   - OUTFLOW: You are the buyer.
   - If unclear but meets criteria → default to "outflow".

TASK: Respond with VALID JSON ONLY. If extraction fails for a field, use "NA" but NEVER empty strings:
{
  "documentType": "invoice|bill|statement|report|contract|email|other",
  "date": "YYYY-MM-DD (extract or use current if missing)",
  "vendorName": "Vendor/company name (cleaned)",
  "invoiceNumber": "Invoice/ref number (or 'NA' if not an invoice)",
  "amount": "Total as number (or '0.00' if not an invoice)",
  "invoiceStatus": "inflow|outflow|irrelevant",
  "isFinancialDocument": true|false,
  "reason": "Brief explanation of classification",
  "isMultiInvoice": true|false,
  "totalInvoices": number,
  "gst": "GST amount or empty",
  "tds": "TDS amount or empty",
  "ot": "Other taxes or empty",
  "na": "Additional notes or empty"
}`;
      
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
          topK: 10,
          topP: 0.8,
          maxOutputTokens: 2048,
        }
      };
      
      const options = {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(GEMINI_URL, options);
      if (response.getResponseCode() !== 200) {
        throw new Error(`API error: ${response.getContentText()}`);
      }
      
      const jsonResponse = JSON.parse(response.getContentText());
      const aiText = jsonResponse.candidates[0].content.parts[0].text.trim();
      const cleanedText = aiText.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
      
      try {
        const extractedData = JSON.parse(cleanedText);
        const validatedData = validateAIClassification(extractedData, fileName);
        return processAIExtractedData(validatedData, fileName);
      } catch (parseError) {
        Logger.log(`JSON parse failed on attempt ${attempt}: ${parseError}. Trying text extraction.`);
        // Fallback: Extract from raw text if JSON fails
        const textExtracted = extractDataFromText(aiText, fileName);
        return validateAIClassification(textExtracted, fileName);  // Treat text-extracted as AI data
      }
    } catch (error) {
      Logger.log(`Attempt ${attempt} failed for ${fileName}: ${error}`);
      if (attempt === maxRetries) {
        Logger.log(`All retries failed. Using fallback extraction.`);
        return fallbackDataExtraction(fileBlob, fileName);
      }
      Utilities.sleep(1000);  // Delay before retry
    }
  }
}

/**
 * Validates and potentially reclassifies AI output using rule-based logic.
 * Ensures only true invoices/bills are classified as inflow/outflow.
 * @param {Object} aiData - Raw data from AI.
 * @param {string} fileName - Original filename for additional checks.
 * @returns {Object} Validated data.
 */
function validateAIClassification(aiData, fileName) {
  const lowerFileName = fileName.toLowerCase();
  
  // Required elements for a true invoice/bill
  const hasInvoiceNumber = aiData.invoiceNumber && aiData.invoiceNumber !== 'NA' && aiData.invoiceNumber.trim() !== '';
  const hasAmount = parseFloat(aiData.amount) > 0;
  const hasDate = validateDate(aiData.date) !== null;
  const hasVendor = aiData.vendorName && aiData.vendorName.trim() !== '';
  
  // If missing any required elements, reclassify as 'irrelevant'
  if (!hasInvoiceNumber || !hasAmount || !hasDate || !hasVendor) {
    Logger.log(`Reclassifying ${fileName} as 'irrelevant': Missing required invoice elements (number=${hasInvoiceNumber}, amount=${hasAmount}, date=${hasDate}, vendor=${hasVendor})`);
    aiData.invoiceStatus = 'irrelevant';
    aiData.documentType = aiData.documentType || 'other'; // Preserve type if set
    aiData.reason = (aiData.reason || '') + ' | Reclassified: Missing core invoice elements';
    aiData.invoiceNumber = 'NA';
    aiData.amount = '0.00';
    return aiData;
  }
  
  // Additional checks: If documentType isn't invoice/bill, reclassify
  const isTrueInvoiceType = ['invoice', 'bill'].includes(aiData.documentType.toLowerCase());
  if (!isTrueInvoiceType) {
    Logger.log(`Reclassifying ${fileName} as 'irrelevant': Document type '${aiData.documentType}' is not a true invoice/bill`);
    aiData.invoiceStatus = 'irrelevant';
    aiData.reason = (aiData.reason || '') + ' | Reclassified: Not a true invoice/bill type';
    return aiData;
  }
  
  // Filename pattern overrides: If filename suggests non-invoice (e.g., 'statement.pdf'), reclassify
  const nonInvoicePatterns = ['statement', 'report', 'contract', 'email', 'photo', 'image', 'screenshot'];
  if (nonInvoicePatterns.some(pattern => lowerFileName.includes(pattern))) {
    Logger.log(`Reclassifying ${fileName} as 'irrelevant': Filename suggests non-invoice pattern`);
    aiData.invoiceStatus = 'irrelevant';
    aiData.reason = (aiData.reason || '') + ' | Reclassified: Filename pattern mismatch';
    return aiData;
  }
  
  // If all checks pass, keep AI's classification
  Logger.log(`Validated ${fileName} as true invoice/bill: Status=${aiData.invoiceStatus}`);
  return aiData;
}

/**
 * Processes AI extracted data and validates it
 * @param {Object} aiData - Raw data from AI
 * @param {string} fileName - Original filename for fallback
 * @returns {Object} Processed data with validation
 */
function processAIExtractedData(aiData, fileName) {
  try {
    Logger.log(`Processing AI data for ${fileName}: ${JSON.stringify(aiData)}`);
    
    // Simple single document processing (no multi-invoice for now)
    const processedData = {
      isMultiInvoice: false,
      invoiceCount: 1,
      date: validateDate(aiData.date) || getCurrentDateString(),
      vendorName: sanitizeVendorName(aiData.vendorName) || extractVendorFromFilename(fileName),
      invoiceNumber: sanitizeInvoiceNumber(aiData.invoiceNumber) || extractInvoiceIdFromFilename(fileName) || generateInvoiceNumber(),
      amount: validateAmount(aiData.amount) || '0.00',
      invoiceStatus: validateAndNormalizeInvoiceStatus(aiData.invoiceStatus, aiData.isFinancialDocument),
      documentType: aiData.documentType || 'document',
      gst: aiData.gst || '',
      tds: aiData.tds || '',
      ot: aiData.ot || '',
      na: aiData.na || ''
    };
    
    Logger.log(`Processed AI data result: Status=${processedData.invoiceStatus}, Vendor=${processedData.vendorName}, Invoice=${processedData.invoiceNumber}`);
    
    return {
      ...processedData,
      invoices: [processedData]
    };
    
  } catch (error) {
    Logger.log(`Error processing AI extracted data: ${error.toString()}`);
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
  
  const validatedData = {
    date: validateDate(data.date) || currentDate,
    vendorName: sanitizeVendorName(data.vendorName) || 'UnknownVendor',
    invoiceNumber: sanitizeInvoiceNumber(data.invoiceNumber) || extractInvoiceIdFromFilename(fileName) || 'INV-Unknown',
    amount: validateAmount(data.amount) || '0.00',
    invoiceStatus: validateInvoiceStatus(data.invoiceStatus) || 'outflow',
    documentType: data.documentType || 'document',
    gst: data.gst || '',
    tds: data.tds || '',
    ot: data.ot || '',
    na: data.na || ''
  };
  
  Logger.log(`Validated data for ${fileName}: Status=${validatedData.invoiceStatus}, Type=${validatedData.documentType}`);
  return validatedData;
}

/**
 * Improved fallback data extraction when AI is not available
 * @param {GoogleAppsScript.Base.Blob} fileBlob - File content
 * @param {string} fileName - Original filename
 * @returns {Object} Extracted data using fallback methods
 */
/**
 * Enhanced fallback with direct text extraction from file content.
 * @param {GoogleAppsScript.Base.Blob} fileBlob File content.
 * @param {string} fileName Original filename.
 * @returns {Object} Extracted data.
 */
function fallbackDataExtraction(fileBlob, fileName) {
  Logger.log(`Enhanced fallback for: ${fileName}`);
  
  let text = '';
  try {
    // Convert to Google Doc for text extraction (handles PDFs/images via OCR)
    const tempFile = DriveApp.createFile(fileBlob);
    const doc = DocumentApp.openById(tempFile.getId());
    text = doc.getBody().getText().toLowerCase();
    tempFile.setTrashed(true);  // Cleanup
  } catch (error) {
    Logger.log(`Text extraction failed: ${error}. Using filename-only fallback.`);
  }

  const data = {
    date: getCurrentDateString(),
    vendorName: 'UnknownVendor',
    invoiceNumber: 'NA',
    amount: '0.00',
    invoiceStatus: 'irrelevant',
    documentType: 'other',
    isFinancialDocument: false,
    reason: 'Fallback extraction',
    isMultiInvoice: false,
    totalInvoices: 1
  };

  // Extract from text (priority) or filename
  const dateMatch = text.match(/date[:\s]*(\d{4}-\d{2}-\d{2}|\d{2}\/\d{2}\/\d{4})/i) || fileName.match(/(\d{4}-\d{2}-\d{2})/);
  if (dateMatch) data.date = standardizeDateFormat(dateMatch[1]);

  const vendorMatch = text.match(/(?:vendor|from|company)[:\s]*([A-Za-z\s&-]+)/i) || fileName.match(/^([A-Za-z\s&-]+)/);
  if (vendorMatch) data.vendorName = sanitizeVendorName(vendorMatch[1]);

  const invMatch = text.match(/(?:invoice|bill|ref)[\s#]*([A-Za-z0-9-]+)/i) || fileName.match(/inv-?([A-Za-z0-9-]+)/i);
  if (invMatch) data.invoiceNumber = sanitizeInvoiceNumber(invMatch[1]);

  const amtMatch = text.match(/total[:\s]*\$?([\d,.]+)/i) || fileName.match(/(\d+\.\d{2})/);
  if (amtMatch) data.amount = validateAmount(amtMatch[1]);

  // Status from keywords
  if (text.match(/invoice|bill/i)) {
    data.documentType = 'invoice';
    data.isFinancialDocument = true;
    data.invoiceStatus = text.match(/sales|inflow/i) ? 'inflow' : 'outflow';
  } else if (text.match(/statement|report/i)) {
    data.documentType = 'statement';
    data.isFinancialDocument = true;
  }

  // If all extraction methods fail, use meaningful words from original filename
  if (data.vendorName === 'UnknownVendor' && data.invoiceNumber === 'NA' && data.amount === '0.00') {
    Logger.log(`All extraction methods failed, using filename-based fallback for: ${fileName}`);
    data.reason = 'Filename-based fallback (all extraction methods failed)';
    
    // Extract meaningful components from original filename
    const meaningfulFilename = generateFallbackFilenameFromOriginal(fileName);
    const parts = meaningfulFilename.replace(/\.[^/.]+$/, '').split('_');
    
    if (parts.length >= 4) {
      data.date = parts[0] || getCurrentDateString();
      data.vendorName = parts[1] || 'Document';
      data.invoiceNumber = parts[2] || 'REF-UNKNOWN';
      data.amount = parts[3] || '0.00';
    }
  }

  Logger.log(`Fallback extracted: ${JSON.stringify(data)}`);
  return data;
}

/**
 * Extracts data from raw AI text response using regex when JSON fails.
 * @param {string} text Raw AI response text.
 * @param {string} fileName Original filename for fallbacks.
 * @returns {Object} Extracted data.
 */
function extractDataFromText(text, fileName) {
  const data = {
    documentType: 'other',
    date: getCurrentDateString(),
    vendorName: extractVendorFromFilename(fileName) || 'UnknownVendor',
    invoiceNumber: extractInvoiceIdFromFilename(fileName) || 'NA',
    amount: '0.00',
    invoiceStatus: 'irrelevant',
    isFinancialDocument: false,
    reason: 'Extracted from text (JSON failed)',
    isMultiInvoice: false,
    totalInvoices: 1
  };

  // Extract date (various formats)
  const dateMatch = text.match(/date[:\s]*(\d{4}-\d{2}-\d{2}|\d{2}\/\d{2}\/\d{4})/i);
  if (dateMatch) data.date = standardizeDateFormat(dateMatch[1]);

  // Extract vendor
  const vendorMatch = text.match(/vendorName[:\s]*"([^"]+)"/i) || text.match(/(?:vendor|company|from)[:\s]*([A-Za-z\s&-]+)/i);
  if (vendorMatch) data.vendorName = sanitizeVendorName(vendorMatch[1]);

  // Extract invoice number
  const invMatch = text.match(/invoiceNumber[:\s]*"([^"]+)"/i) || text.match(/(?:invoice|ref|bill)[\s#]*([A-Za-z0-9-]+)/i);
  if (invMatch) data.invoiceNumber = sanitizeInvoiceNumber(invMatch[1]);

  // Extract amount
  const amtMatch = text.match(/amount[:\s]*"([^"]+)"/i) || text.match(/total[:\s]*\$?([\d,.]+)/i);
  if (amtMatch) data.amount = validateAmount(amtMatch[1]);

  // Determine status from keywords
  if (text.match(/inflow|sales|revenue/i)) data.invoiceStatus = 'inflow';
  else if (text.match(/outflow|expense|bill/i)) data.invoiceStatus = 'outflow';

  // Financial check
  if (text.match(/financial|invoice|bill|amount/i)) data.isFinancialDocument = true;

  Logger.log(`Text-extracted data for ${fileName}: ${JSON.stringify(data)}`);
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
  if (!status) return 'outflow'; // Default to outflow if no status
  const lowerStatus = status.toString().toLowerCase().trim();
  
  // Map various status terms to standard values
  if (['inflow', 'income', 'revenue', 'sales', 'receipt', 'credit'].includes(lowerStatus)) {
    return 'inflow';
  } else if (['outflow', 'expense', 'purchase', 'bill', 'payment', 'cost'].includes(lowerStatus)) {
    return 'outflow';
  } else if (['irrelevant', 'not_relevant', 'non_financial', 'invalid'].includes(lowerStatus)) {
    return 'irrelevant';
  } else {
    // If unclear, default to outflow (most business documents are expenses)
    Logger.log(`Uncertain invoice status '${status}' - defaulting to outflow`);
    return 'outflow';
  }
}

/**
 * Enhanced invoice status validation that considers document type
 * @param {string} status - The status from AI
 * @param {boolean} isFinancialDocument - Whether this is a financial document
 * @returns {string} Validated status
 */
function validateAndNormalizeInvoiceStatus(status, isFinancialDocument) {
  if (!status) {
    return isFinancialDocument === false ? 'irrelevant' : 'outflow';
  }
  
  const lowerStatus = status.toString().toLowerCase().trim();
  
  // Handle non-financial documents first
  if (isFinancialDocument === false || lowerStatus === 'irrelevant' || lowerStatus === 'non_financial') {
    Logger.log(`Document classified as irrelevant/non-financial`);
    return 'irrelevant';
  }
  
  // Handle financial documents
  if (['inflow', 'income', 'revenue', 'sales', 'receipt', 'credit', 'sale'].includes(lowerStatus)) {
    Logger.log(`Document classified as INFLOW (money coming in)`);
    return 'inflow';
  } else if (['outflow', 'expense', 'purchase', 'bill', 'payment', 'cost', 'expenditure'].includes(lowerStatus)) {
    Logger.log(`Document classified as OUTFLOW (money going out)`);
    return 'outflow';
  } else {
    // Default for financial documents is outflow (most business docs are expenses)
    Logger.log(`Uncertain financial document status '${status}' - defaulting to OUTFLOW`);
    return 'outflow';
  }
}

/**
 * Extracts vendor name from filename as fallback
 * @param {string} fileName - Original filename
 * @returns {string} Extracted vendor name or default
 */
function extractVendorFromFilename(fileName) {
  if (!fileName) return 'UnknownVendor';
  
  // Remove extension
  const nameWithoutExt = fileName.replace(/\.[^.]+$/, '');
  
  // Try to extract company/vendor name patterns
  const patterns = [
    /^([A-Za-z\s&]+?)[-_\s]*(invoice|bill|receipt)/i,
    /^([A-Za-z\s&]{3,20})/,
  ];
  
  for (const pattern of patterns) {
    const match = nameWithoutExt.match(pattern);
    if (match) {
      return sanitizeVendorName(match[1]) || 'UnknownVendor';
    }
  }
  
  return 'UnknownVendor';
}

/**
 * Generates a unique invoice number as fallback
 * @returns {string} Generated invoice number
 */
function generateInvoiceNumber() {
  const timestamp = new Date().getTime().toString().slice(-6);
  return `INV-${timestamp}`;
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

  // --- Handle Buffer2 sheet 'Relevance' column edits (Yes/No) ---
  // Yes: Move file from Buffer2 to Buffer/Active and log in buffer & main sheets
  // No: Move file from Buffer2 to Buffer/Active (existing functionality)
  if (
    sheetName.endsWith('-buffer2') &&
    range.getColumn() === BUFFER2_SHEET_HEADERS.indexOf('Relevance') + 1 &&
    range.getNumRows() === 1 &&
    range.getNumColumns() === 1
  ) {
    let companyName = sheetName.replace('-buffer2', '');
    const editedRow = range.getRow();
    const newRelevance = e.value;
    const oldRelevance = e.oldValue;
    if (editedRow === 1 || newRelevance === oldRelevance) return;
    
    // Handle "Yes" selection - move file from Buffer2 to Buffer and log in main sheet
    if (newRelevance === 'Yes') {
      const rowData = sheet.getRange(editedRow, 1, 1, BUFFER2_SHEET_HEADERS.length).getValues()[0];
      const date = rowData[0];
      const originalFilename = rowData[1];
      const changedFilename = rowData[2];
      const invoiceId = rowData[3];
      let driveFileId = rowData[4];
      const gmailMessageId = rowData[5];
      const ui = SpreadsheetApp.getUi();
      const ss = SpreadsheetApp.getActiveSpreadsheet();

      try {
        // Validate input data
        if (!changedFilename || changedFilename.trim() === '') {
          ui.alert('Error', 'Filename is missing or empty. Cannot process.', ui.ButtonSet.OK);
          return;
        }
        
        Logger.log(`Processing Buffer2 relevance=Yes for file: ${changedFilename}`);
        Logger.log(`Initial company from sheet name: ${companyName}`);
        
        // Try to extract company name from changed filename if sheet name doesn't work
        const filenameCompany = extractCompanyFromFilename(changedFilename);
        if (filenameCompany) {
          Logger.log(`Company extracted from filename: ${filenameCompany}`);
          companyName = filenameCompany;
        }
        
        Logger.log(`Final company name: ${companyName}, Date: ${date}, Gmail ID: ${gmailMessageId}`);
        
        // Find the correct financial year (try from date, else current)
        let fileDate = new Date(date);
        if (isNaN(fileDate.getTime())) {
          Logger.log(`Invalid date found: ${date}, using current date`);
          fileDate = new Date();
        }
        const financialYear = calculateFinancialYear(fileDate);
        Logger.log(`Using financial year: ${financialYear}`);
        
        // Ensure company folder mappings are loaded
        loadCompanyFolderMappings();
        
        // Get or create company mapping
        const validCompanyName = getOrCreateCompanyMapping(companyName);
        if (!validCompanyName) {
          Logger.log(`Failed to get or create company mapping for: ${companyName}`);
          ui.alert('Error', `Could not find or create company folder for: ${companyName}. Available companies: ${Object.keys(ATTACHMENT_COMPANY_FOLDER_MAP).join(', ')}. Please contact administrator.`, ui.ButtonSet.OK);
          return;
        }
        companyName = validCompanyName;
        
        // Create/get Buffer2 folder structure
        const buffer2Folder = createBuffer2FolderStructure(companyName, financialYear);
        Logger.log(`Buffer2 folder ID: ${buffer2Folder.getId()}`);
        
        // Validate Drive access to Buffer2 folder
        if (!validateDriveAccess(buffer2Folder.getId())) {
          ui.alert('Error', 'Insufficient permissions to access Buffer2 folder. Please contact administrator to grant Drive permissions.', ui.ButtonSet.OK);
          return;
        }
        
        // Find the file in Buffer2 folder
        const files = buffer2Folder.getFilesByName(changedFilename);
        if (!files.hasNext()) {
          Logger.log(`File not found in Buffer2 folder: ${changedFilename}`);
          ui.alert('Error', `File ${changedFilename} not found in Buffer2 folder. It may have been moved or deleted.`, ui.ButtonSet.OK);
          return;
        }
        
        const file = files.next();
        driveFileId = file.getId();
        Logger.log(`Found file in Buffer2, ID: ${driveFileId}`);
        
        // Validate file ID
        if (!driveFileId || driveFileId.trim() === '') {
          Logger.log(`Invalid file ID: ${driveFileId}`);
          ui.alert('Error', 'Invalid file ID. Cannot move file.', ui.ButtonSet.OK);
          return;
        }

        // Create/get Buffer/Active folder
        const bufferActiveFolder = getOrCreateBufferSubfolder(companyName, financialYear, 'Active');
        Logger.log(`Buffer/Active folder ID: ${bufferActiveFolder.getId()}`);
        
        // Validate Drive access to Buffer/Active folder
        if (!validateDriveAccess(bufferActiveFolder.getId())) {
          ui.alert('Error', 'Insufficient permissions to access Buffer/Active folder. Please contact administrator to grant Drive permissions.', ui.ButtonSet.OK);
          return;
        }
        
        // Validate folder IDs before moving
        if (!bufferActiveFolder.getId() || !buffer2Folder.getId()) {
          Logger.log(`Invalid folder IDs - Buffer/Active: ${bufferActiveFolder.getId()}, Buffer2: ${buffer2Folder.getId()}`);
          ui.alert('Error', 'Invalid folder IDs. Cannot move file.', ui.ButtonSet.OK);
          return;
        }
        
        Logger.log(`Moving file ${changedFilename} from Buffer2 to Buffer/Active (Relevance=Yes)`);
        Logger.log(`File ID: ${driveFileId}, Target: ${bufferActiveFolder.getId()}, Source: ${buffer2Folder.getId()}`);
        
        // Move the file
        moveFileWithDriveApp(driveFileId, bufferActiveFolder.getId(), buffer2Folder.getId());

        // Add log to buffer sheet if not already present
        const bufferSheet = ss.getSheetByName(`${companyName}-buffer`);
        let alreadyLogged = false;
        if (bufferSheet) {
          const bufferData = bufferSheet.getDataRange().getValues();
          for (let i = 1; i < bufferData.length; i++) {
            if (bufferData[i][2] === changedFilename || bufferData[i][4] === driveFileId) {
              alreadyLogged = true;
              break;
            }
          }
        }
        if (!alreadyLogged) {
          // Generate UI if not present
          let uniqueIdentifier = rowData[7];
          if (!uniqueIdentifier) {
            uniqueIdentifier = generateUniqueIdentifierForFile(driveFileId);
          }
          // Append to buffer sheet
          bufferSheet.appendRow([
            new Date(),
            originalFilename,
            changedFilename,
            invoiceId,
            driveFileId,
            gmailMessageId,
            'Moved from Buffer2 (Relevance=Yes)',
            'Active',
            uniqueIdentifier,
            '',
            1, // Invoice count default
            '', // Attachment ID
            gmailMessageId
          ]);
          sortSheetByDateDesc(bufferSheet, 1);
          Logger.log(`Added file to buffer sheet: ${changedFilename}`);
        }

        // Add log to main sheet
        const mainSheet = ss.getSheetByName(companyName);
        if (mainSheet) {
          // Check if already logged in main sheet
          const mainData = mainSheet.getDataRange().getValues();
          let alreadyInMain = false;
          for (let i = 1; i < mainData.length; i++) {
            if (mainData[i][1] === driveFileId || mainData[i][0] === changedFilename) {
              alreadyInMain = true;
              break;
            }
          }
          if (!alreadyInMain) {
            // Get email subject for main sheet logging
            const emailSubject = getEmailSubjectForMessageId(gmailMessageId) || 'Unknown Subject';
            
            // Default to outflow status for files moved from Buffer2
            const invoiceStatus = 'outflow';
            
            // Get unique identifier
            let uniqueIdentifier = rowData[7];
            if (!uniqueIdentifier) {
              uniqueIdentifier = generateUniqueIdentifierForFile(driveFileId);
            }
            
            // Log to main sheet
            logFileToMainSheet(mainSheet, file, emailSubject, gmailMessageId, invoiceStatus, companyName, uniqueIdentifier, {
              date: fileDate.toISOString().split('T')[0],
              month: getMonthFromDate(fileDate),
              fy: calculateFinancialYear(fileDate),
              gst: '',
              tds: '',
              ot: '',
              na: ''
            });
            Logger.log(`Added file to main sheet: ${changedFilename}`);
          }
        }

        // Remove the row from buffer2 sheet
        setScriptEditFlag(true);
        sheet.deleteRow(editedRow);
        ui.alert('Success', `File ${changedFilename} moved to Buffer/Active and logged in buffer and main sheets.`, ui.ButtonSet.OK);
        
             } catch (err) {
         Logger.log(`Error moving file from Buffer2 to Buffer (Relevance=Yes): ${err.toString()}`);
         
         // Provide specific error messages based on error type
         let errorMessage = `Failed to move file from Buffer2 to Buffer: ${err.message}`;
         
         if (err.toString().includes('Invalid argument: id')) {
           errorMessage = 'Invalid file or folder ID. The file may have been moved or deleted. Please refresh and try again.';
         } else if (err.toString().includes('permissions') || err.toString().includes('auth')) {
           errorMessage = 'Insufficient permissions to access Google Drive. Please contact administrator to grant proper Drive permissions.';
         } else if (err.toString().includes('not found')) {
           errorMessage = 'File or folder not found. The file may have been moved or deleted.';
         }
         
         ui.alert('Error', errorMessage, ui.ButtonSet.OK);
       }
      return;
    }
    
    // Handle "No" selection - move file from Buffer2 to Buffer/Active (existing functionality)
    if (newRelevance !== 'No') return;

    const rowData = sheet.getRange(editedRow, 1, 1, BUFFER2_SHEET_HEADERS.length).getValues()[0];
    const date = rowData[0];
    const originalFilename = rowData[1];
    const changedFilename = rowData[2];
    const invoiceId = rowData[3];
    let driveFileId = rowData[4];
    const gmailMessageId = rowData[5];
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Try to extract company name from changed filename if sheet name doesn't work
    const filenameCompany = extractCompanyFromFilename(changedFilename);
    if (filenameCompany) {
      Logger.log(`Company extracted from filename for 'No' selection: ${filenameCompany}`);
      companyName = filenameCompany;
    }

    // Find the file in Buffer2 folder (by changedFilename)
    try {
      // Find the correct financial year (try from date, else current)
      let fileDate = new Date(date);
      if (isNaN(fileDate.getTime())) fileDate = new Date();
      const financialYear = calculateFinancialYear(fileDate);
      
      // Get or create company mapping
      const validCompanyName = getOrCreateCompanyMapping(companyName);
      if (!validCompanyName) {
        Logger.log(`Failed to get or create company mapping for 'No' selection: ${companyName}`);
        ui.alert('Error', `Could not find or create company folder for: ${companyName}. Please contact administrator.`, ui.ButtonSet.OK);
        return;
      }
      companyName = validCompanyName;
      
      const buffer2Folder = createBuffer2FolderStructure(companyName, financialYear);
      const files = buffer2Folder.getFilesByName(changedFilename);
      if (!files.hasNext()) {
        ui.alert('Error', `File ${changedFilename} not found in Buffer2 folder.`, ui.ButtonSet.OK);
        return;
      }
      const file = files.next();
      driveFileId = file.getId();

      // Move file to Buffer/Active
      const bufferActiveFolder = getOrCreateBufferSubfolder(companyName, financialYear, 'Active');
      Logger.log(`Moving file ${changedFilename} from Buffer2 to Buffer/Active`);
      moveFileWithDriveApp(file.getId(), bufferActiveFolder.getId(), buffer2Folder.getId());

      // Add log to buffer sheet if not already present
      const bufferSheet = ss.getSheetByName(`${companyName}-buffer`);
      let alreadyLogged = false;
      if (bufferSheet) {
        const bufferData = bufferSheet.getDataRange().getValues();
        for (let i = 1; i < bufferData.length; i++) {
          if (bufferData[i][1] === changedFilename || bufferData[i][3] === driveFileId) {
            alreadyLogged = true;
            break;
          }
        }
      }
      if (!alreadyLogged) {
        // Generate UI if not present
        let uniqueIdentifier = rowData[7];
        if (!uniqueIdentifier) {
          uniqueIdentifier = generateUniqueIdentifierForFile(driveFileId);
        }
        // Append to buffer sheet
        bufferSheet.appendRow([
          new Date(),
          originalFilename,
          changedFilename,
          invoiceId,
          driveFileId,
          gmailMessageId,
          'Moved from Buffer2 (Relevance=No)',
          'Active',
          uniqueIdentifier,
          '',
          1, // Invoice count default
          '', // Attachment ID
          gmailMessageId
        ]);
        sortSheetByDateDesc(bufferSheet, 1);
      }

      // Remove the row from buffer2 sheet
      setScriptEditFlag(true);
      sheet.deleteRow(editedRow);
      ui.alert('File moved to Buffer/Active and logged in buffer sheet.');
    } catch (err) {
      Logger.log(`Error moving file from Buffer2 to Buffer: ${err.toString()}`);
      ui.alert('Error', `Failed to move file from Buffer2 to Buffer: ${err.message}`, ui.ButtonSet.OK);
    }
    return;
  }

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
    const originalFilename = rowData[1]; // OriginalFileName (column 2, index 1)
    let changedFilename = rowData[2];     // ChangedFilename (column 3, index 2) - THE SOURCE OF TRUTH (let allows reassignment)
    let driveFileId = rowData[4];         // Drive File ID (column 5, index 4)
    const gmailMessageId = rowData[5];    // Gmail Message ID (column 6, index 5)
    const ui = SpreadsheetApp.getUi();

    // Validate row data
    Logger.log(`Processing buffer sheet row ${editedRow}:`);
    Logger.log(`Original Filename: ${originalFilename}`);
    Logger.log(`Changed Filename: ${changedFilename}`);
    Logger.log(`Drive File ID: ${driveFileId}`);
    Logger.log(`Gmail Message ID: ${gmailMessageId}`);
    
    if (!changedFilename || changedFilename.trim() === '') {
      ui.alert('Error', 'Changed filename is missing in buffer sheet row. Cannot proceed.', ui.ButtonSet.OK);
      setScriptEditFlag(true);
      sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('Status') + 1).setValue(oldStatus || 'Active');
      return;
    }

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
            Logger.log(`Attempting to delete file with ID: ${driveFileId}`);
            
            // Use the new file recovery function to find the file
            const recoveryResult = findAndRecoverFile(companyName, changedFilename, driveFileId);
            
            if (!recoveryResult.file) {
              throw new Error(`File "${changedFilename}" not found in any location. It may have already been moved or deleted manually.`);
            }
            
            const file = recoveryResult.file;
            Logger.log(`File recovery successful: Found "${file.getName()}" in ${recoveryResult.foundIn}`);
            
            // Update the buffer sheet with the correct file ID if it was recovered
            if (recoveryResult.foundIn !== 'current_id') {
              setScriptEditFlag(true);
              sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('Drive File ID') + 1).setValue(file.getId());
              Logger.log(`Updated buffer sheet with recovered file ID: ${file.getId()}`);
            }
            
            // If the actual filename is different from what's in the buffer sheet, update it
            const actualFilename = file.getName();
            if (actualFilename !== changedFilename) {
              Logger.log(`Filename mismatch detected. Buffer: "${changedFilename}", Actual: "${actualFilename}"`);
              setScriptEditFlag(true);
              sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('ChangedFilename') + 1).setValue(actualFilename);
              Logger.log(`Updated buffer sheet with correct filename: ${actualFilename}`);
              // Update the local variable to use the correct filename for subsequent operations
              changedFilename = actualFilename;
            }
            
            const fileDate = file.getDateCreated();
            const financialYear = calculateFinancialYear(fileDate);
            const month = getMonthFromDate(fileDate);
            const bufferActiveFolder = getOrCreateBufferSubfolder(companyName, financialYear, "Active");
            const bufferDeletedFolder = getOrCreateBufferSubfolder(companyName, financialYear, "Deleted");

            Logger.log(`=== DELETE OPERATION STARTED for ${changedFilename} ===`);
            Logger.log(`File date: ${fileDate}, Financial Year: ${financialYear}, Month: ${month}`);
            Logger.log(`Buffer Active folder: ${bufferActiveFolder.getName()} (ID: ${bufferActiveFolder.getId()})`);
            Logger.log(`Buffer Deleted folder: ${bufferDeletedFolder.getName()} (ID: ${bufferDeletedFolder.getId()})`);

            // Debug: Check current file locations
            debugFileFinding(companyName, financialYear, changedFilename);

            // Use AI to determine inflow/outflow status for proper folder targeting
            const aiResult = callGeminiAPIInternal(file.getBlob(), changedFilename);
            const invoiceStatus = aiResult.invoiceStatus || "unknown";
            Logger.log(`AI determined invoice status: ${invoiceStatus}`);
            
            // 1. First, find and move files from inflow/outflow folders to Buffer/Deleted
            if (invoiceStatus === "inflow" || invoiceStatus === "outflow") {
              try {
                const flowFolder = createFlowFolderStructure(companyName, financialYear, month, invoiceStatus);
                Logger.log(`Looking for files in ${invoiceStatus} folder: ${flowFolder.getName()}`);
                
                const filesInFlow = flowFolder.getFilesByName(changedFilename);
                let filesMovedFromFlow = 0;
                
                while (filesInFlow.hasNext()) {
                  const flowFile = filesInFlow.next();
                  try {
                    Logger.log(`Found file in ${invoiceStatus} folder: ${flowFile.getName()} (ID: ${flowFile.getId()})`);
                    // Move the file from inflow/outflow to Buffer/Deleted folder
                    moveFileWithDriveApp(flowFile.getId(), bufferDeletedFolder.getId(), flowFolder.getId());
                    Logger.log(`Successfully moved file ${changedFilename} from ${invoiceStatus} folder to Buffer/Deleted folder.`);
                    filesMovedFromFlow++;
                  } catch (moveError) {
                    Logger.log(`Error moving file from ${invoiceStatus} folder: ${moveError.toString()}`);
                    // Continue with other operations even if this fails
                  }
                }
                
                if (filesMovedFromFlow === 0) {
                  Logger.log(`No files found to move from ${invoiceStatus} folder for filename: ${changedFilename}`);
                }
              } catch (flowFolderError) {
                Logger.log(`Error accessing ${invoiceStatus} folder: ${flowFolderError.toString()}`);
              }
            }
            
            // 2. Move the main file from Buffer/Active to Buffer/Deleted
            try {
              Logger.log(`Moving main file ${changedFilename} from Buffer/Active to Buffer/Deleted`);
              moveFileWithDriveApp(file.getId(), bufferDeletedFolder.getId(), bufferActiveFolder.getId());
              Logger.log(`Successfully moved main file ${changedFilename} from Buffer/Active to Buffer/Deleted.`);
            } catch (moveError) {
              Logger.log(`Error moving main file to Buffer/Deleted: ${moveError.toString()}`);
              throw moveError; // Re-throw to trigger the catch block below
            }

            // 3. Remove log entries from all relevant sheets
            deleteLogEntries(mainSheet, driveFileId, gmailMessageId);
            deleteLogEntries(inflowSheet, driveFileId, gmailMessageId);
            deleteLogEntries(outflowSheet, driveFileId, gmailMessageId);
            Logger.log(`Removed log entries for file ${changedFilename} from all sheets.`);

            // 4. Update buffer sheet - mark as DELETED and set background color
            setScriptEditFlag(true);
            sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('Drive File ID') + 1).setValue('DELETED');
            sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('Reason') + 1).setValue(reasonText);
            sheet.getRange(editedRow, 1, 1, BUFFER_SHEET_HEADERS.length).setBackground('#FFD966'); // Orange for deleted
            
            Logger.log(`=== DELETE OPERATION COMPLETED SUCCESSFULLY for ${changedFilename} ===`);
            
          } catch (e) {
            Logger.log(`=== DELETE OPERATION FAILED for ${changedFilename}: ${e.toString()} ===`);
            ui.alert('Error', `Failed to delete file "${changedFilename}": ${e.message}`, ui.ButtonSet.OK);
            setScriptEditFlag(true);
            sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('Status') + 1).setValue(oldStatus || 'Active');
          }
        } else {
          Logger.log(`No valid Drive File ID found for ${changedFilename}. Attempting file recovery...`);
          
          // Try to find the file using recovery function even without a file ID
          const recoveryResult = findAndRecoverFile(companyName, changedFilename, '');
          
          if (recoveryResult.file) {
            Logger.log(`File recovered without file ID: ${recoveryResult.file.getName()} in ${recoveryResult.foundIn}`);
            
            // Update buffer sheet with found file ID and retry the delete operation
            setScriptEditFlag(true);
            sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('Drive File ID') + 1).setValue(recoveryResult.file.getId());
            sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('Reason') + 1).setValue(`${reasonText} (File ID recovered automatically)`);
            
            ui.alert('Success', `File "${changedFilename}" was found and its ID has been updated. Please try the delete operation again.`, ui.ButtonSet.OK);
            
            // Revert status to allow user to try again
            setScriptEditFlag(true);
            sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('Status') + 1).setValue('Active');
          } else {
            Logger.log(`File ${changedFilename} not found anywhere. Marking as permanently missing.`);
            
            setScriptEditFlag(true);
            sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('Drive File ID') + 1).setValue('NOT_FOUND');
            sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('Reason') + 1).setValue(`${reasonText} (File not found in any location)`);
            sheet.getRange(editedRow, 1, 1, BUFFER_SHEET_HEADERS.length).setBackground('#FF9999'); // Light red for missing
            
            // Still remove any existing log entries since the file is gone
            deleteLogEntries(mainSheet, driveFileId || '', gmailMessageId);
            deleteLogEntries(inflowSheet, driveFileId || '', gmailMessageId);
            deleteLogEntries(outflowSheet, driveFileId || '', gmailMessageId);
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

        // Handle both DELETED files and existing files
        let file = null;
        let fileDate = null;
        let financialYear = null;
        let isRestoredFromDeleted = false;

        try {
          Logger.log(`=== ACTIVATE OPERATION STARTED for ${changedFilename} ===`);
          Logger.log(`Drive File ID: ${driveFileId}, Old Status: ${oldStatus}`);
          
          // Use the robust file recovery function to find the file
          const recoveryResult = findAndRecoverFile(companyName, changedFilename, driveFileId);
          
          if (!recoveryResult.file) {
            throw new Error(`File "${changedFilename}" not found in any location. It may have been permanently deleted.`);
          }
          
          file = recoveryResult.file;
          fileDate = file.getDateCreated();
          financialYear = recoveryResult.financialYear || calculateFinancialYear(fileDate);
          
          // Determine if this is a restoration from deleted folder
          isRestoredFromDeleted = (recoveryResult.location === 'buffer_deleted');
          
                     Logger.log(`File recovery successful: Found "${file.getName()}" in ${recoveryResult.foundIn}`);
           Logger.log(`File will be restored from deleted: ${isRestoredFromDeleted}`);
           
           // Update the buffer sheet with the correct file ID if it was recovered
           if (recoveryResult.foundIn !== 'current_id') {
             setScriptEditFlag(true);
             sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('Drive File ID') + 1).setValue(file.getId());
             Logger.log(`Updated buffer sheet with recovered file ID: ${file.getId()}`);
           }
           
           // If the actual filename is different from what's in the buffer sheet, update it
           const actualFilename = file.getName();
           if (actualFilename !== changedFilename) {
             Logger.log(`Filename mismatch detected during activation. Buffer: "${changedFilename}", Actual: "${actualFilename}"`);
             setScriptEditFlag(true);
             sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('ChangedFilename') + 1).setValue(actualFilename);
             Logger.log(`Updated buffer sheet with correct filename: ${actualFilename}`);
             // Update the local variable to use the correct filename for subsequent operations
             changedFilename = actualFilename;
           }

           const month = getMonthFromDate(fileDate);
           const bufferActiveFolder = getOrCreateBufferSubfolder(companyName, financialYear, "Active");
           const bufferDeletedFolder = getOrCreateBufferSubfolder(companyName, financialYear, "Deleted");
          
                     // 1. Move file from Buffer/Deleted to Buffer/Active (if it was in deleted)
           if (isRestoredFromDeleted) {
             try {
               Logger.log(`Moving file ${changedFilename} from Buffer/Deleted to Buffer/Active`);
               moveFileWithDriveApp(file.getId(), bufferActiveFolder.getId(), bufferDeletedFolder.getId());
               Logger.log(`Successfully moved file ${changedFilename} from Buffer/Deleted to Buffer/Active.`);
             } catch (moveError) {
               Logger.log(`Error moving file from deleted to active: ${moveError.toString()}`);
               throw moveError; // Re-throw to trigger the catch block below
             }
           }

          // 2. Use AI to determine inflow/outflow status
          const aiResult = callGeminiAPIInternal(file.getBlob(), changedFilename);
          const invoiceStatus = aiResult.invoiceStatus || "unknown";
          const emailSubject = getEmailSubjectForMessageId(gmailMessageId);
          
          // 3. Get UI from buffer sheet for logging
          const existingUI = sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('UI') + 1).getValue();
          const uniqueIdentifier = existingUI || generateUniqueIdentifierForFile(file.getId());
          
          // 4. Always log to main sheet first
          // Use invoice date from AI extraction, fallback to file creation date
          const onEditInvoiceDate = aiResult.date ? new Date(aiResult.date) : fileDate;
          
          logFileToMainSheet(mainSheet, file, emailSubject, gmailMessageId, invoiceStatus, companyName, uniqueIdentifier, {
            date: onEditInvoiceDate.toISOString().split('T')[0],
            month: getMonthFromDate(onEditInvoiceDate),
            fy: calculateFinancialYear(onEditInvoiceDate),
            gst: aiResult.gst || '',
            tds: aiResult.tds || '',
            ot: aiResult.ot || '',
            na: aiResult.na || '',
            vendorName: aiResult.vendorName || 'Unknown'
          });
          
          // 5. Copy to inflow/outflow folder and log if applicable
          if (invoiceStatus === "inflow" || invoiceStatus === "outflow") {
            const flowFolder = createFlowFolderStructure(companyName, financialYear, month, invoiceStatus);
            
            // Check if file already exists in flow folder
            const existingFilesInFlow = flowFolder.getFilesByName(changedFilename);
            let copiedFile = null;
            
            if (existingFilesInFlow.hasNext()) {
              copiedFile = existingFilesInFlow.next();
              Logger.log(`File ${changedFilename} already exists in ${invoiceStatus} folder.`);
            } else {
              copiedFile = file.makeCopy(changedFilename, flowFolder);
              Logger.log(`Copied file ${changedFilename} to ${invoiceStatus} folder.`);
            }

            // Log to inflow/outflow sheet
            const flowSheet = (invoiceStatus === "inflow") ? inflowSheet : outflowSheet;
            // Use invoice date from AI extraction, fallback to file creation date
            const onEditFlowInvoiceDate = aiResult.date ? new Date(aiResult.date) : fileDate;
            
            logFileToMainSheet(flowSheet, copiedFile, emailSubject, gmailMessageId, invoiceStatus, companyName, uniqueIdentifier, {
              date: onEditFlowInvoiceDate.toISOString().split('T')[0],
              month: getMonthFromDate(onEditFlowInvoiceDate),
              fy: calculateFinancialYear(onEditFlowInvoiceDate),
              gst: aiResult.gst || '',
              tds: aiResult.tds || '',
              ot: aiResult.ot || '',
              na: aiResult.na || '',
              vendorName: aiResult.vendorName || 'Unknown'
            });
          }
          
          // 6. Update buffer sheet
          setScriptEditFlag(true);
          sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('Drive File ID') + 1).setValue(file.getId());
          sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('Reason') + 1).setValue(reasonText);
          sheet.getRange(editedRow, 1, 1, BUFFER_SHEET_HEADERS.length).setBackground(null); // Remove background color
          
          // Update UI if it was missing
          if (!existingUI) {
            sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('UI') + 1).setValue(uniqueIdentifier);
          }
          
          Logger.log(`=== ACTIVATE OPERATION COMPLETED SUCCESSFULLY for ${changedFilename} ===`);
          
        } catch (e) {
          Logger.log(`=== ACTIVATE OPERATION FAILED for ${changedFilename}: ${e.toString()} ===`);
          ui.alert('Error', `Failed to activate file "${changedFilename}": ${e.message}`, ui.ButtonSet.OK);
          setScriptEditFlag(true);
          sheet.getRange(editedRow, BUFFER_SHEET_HEADERS.indexOf('Status') + 1).setValue(oldStatus || 'Delete');
          return;
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
 * Moves a file from one folder to another using Google Apps Script DriveApp.
 * This is more reliable than using the REST API directly.
 * @param {string} fileId - The ID of the file to move.
 * @param {string} addParentId - The folder ID to add as parent.
 * @param {string} removeParentId - The folder ID to remove as parent (optional).
 */
function moveFileWithDriveApp(fileId, addParentId, removeParentId = null) {
  try {
    // Validate inputs
    if (!fileId || fileId.trim() === '') {
      throw new Error('File ID is required and cannot be empty');
    }
    if (!addParentId || addParentId.trim() === '') {
      throw new Error('Target folder ID is required and cannot be empty');
    }
    
    Logger.log(`Attempting to move file with ID: ${fileId}`);
    Logger.log(`Target folder ID: ${addParentId}`);
    Logger.log(`Source folder ID: ${removeParentId || 'Not specified'}`);
    
    // Get the file with better error handling
    let file;
    try {
      file = DriveApp.getFileById(fileId.trim());
      Logger.log(`Successfully retrieved file: ${file.getName()}`);
    } catch (fileError) {
      Logger.log(`Failed to retrieve file by ID ${fileId}: ${fileError.toString()}`);
      throw new Error(`File with ID ${fileId} not found or not accessible. It may have been moved or deleted.`);
    }
    
    // Get the target folder
    let targetFolder;
    try {
      targetFolder = DriveApp.getFolderById(addParentId.trim());
      Logger.log(`Target folder: ${targetFolder.getName()}`);
    } catch (folderError) {
      Logger.log(`Failed to retrieve target folder by ID ${addParentId}: ${folderError.toString()}`);
      throw new Error(`Target folder with ID ${addParentId} not found or not accessible.`);
    }
    
    Logger.log(`Moving file "${file.getName()}" to folder "${targetFolder.getName()}"`);
    
    // Check if file is already in the target folder
    const currentParents = file.getParents();
    let alreadyInTarget = false;
    while (currentParents.hasNext()) {
      const parent = currentParents.next();
      if (parent.getId() === addParentId.trim()) {
        alreadyInTarget = true;
        Logger.log(`File is already in target folder: ${parent.getName()}`);
        break;
      }
    }
    
    // Remove from current parents if specified
    if (removeParentId && removeParentId.trim() !== '') {
      try {
        const sourceFolder = DriveApp.getFolderById(removeParentId.trim());
        
        // Check if file is actually in this source folder
        const filesInSource = sourceFolder.getFilesByName(file.getName());
        let fileFoundInSource = false;
        while (filesInSource.hasNext()) {
          const fileInSource = filesInSource.next();
          if (fileInSource.getId() === file.getId()) {
            fileFoundInSource = true;
            break;
          }
        }
        
        if (fileFoundInSource) {
          sourceFolder.removeFile(file);
          Logger.log(`Removed file from source folder: ${sourceFolder.getName()}`);
        } else {
          Logger.log(`File not found in specified source folder: ${sourceFolder.getName()}`);
        }
      } catch (removeError) {
        Logger.log(`Warning: Could not remove file from source folder ${removeParentId}: ${removeError.toString()}`);
        // Continue with adding to new folder even if removal fails
      }
    } else {
      // Remove from all current parents except target
      const parents = file.getParents();
      const parentsToRemove = [];
      
      while (parents.hasNext()) {
        const parent = parents.next();
        if (parent.getId() !== addParentId.trim()) {
          parentsToRemove.push(parent);
        }
      }
      
      for (const parent of parentsToRemove) {
        try {
          parent.removeFile(file);
          Logger.log(`Removed file from parent folder: ${parent.getName()}`);
        } catch (removeError) {
          Logger.log(`Warning: Could not remove file from parent folder ${parent.getName()}: ${removeError.toString()}`);
        }
      }
    }
    
    // Add to new parent if not already there
    if (!alreadyInTarget) {
      try {
        targetFolder.addFile(file);
        Logger.log(`Successfully added file to target folder: ${targetFolder.getName()}`);
      } catch (addError) {
        Logger.log(`Error adding file to target folder: ${addError.toString()}`);
        throw new Error(`Failed to add file to target folder: ${addError.message}`);
      }
    }
    
    Logger.log(`File move operation completed successfully`);
    return { success: true, fileId: fileId, fileName: file.getName() };
    
  } catch (error) {
    Logger.log(`Error in moveFileWithDriveApp: ${error.toString()}`);
    throw new Error(`Failed to move file: ${error.message}`);
  }
}

/**
 * Legacy function for backward compatibility - redirects to new implementation
 * @param {string} fileId - The ID of the file to move.
 * @param {string} addParentId - The folder ID to add as parent.
 * @param {string} removeParentId - The folder ID to remove as parent.
 */
function moveFileWithDriveApi(fileId, addParentId, removeParentId) {
  return moveFileWithDriveApp(fileId, addParentId, removeParentId);
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
 * Enhanced file finder that searches across multiple possible locations
 * @param {string} filename - The filename to search for
 * @param {Array<GoogleAppsScript.Drive.DriveFolder>} foldersToSearch - Array of folders to search in
 * @returns {Array<GoogleAppsScript.Drive.File>} Array of found files
 */
function findFilesInFolders(filename, foldersToSearch) {
  const foundFiles = [];
  
  for (const folder of foldersToSearch) {
    try {
      const filesInFolder = folder.getFilesByName(filename);
      while (filesInFolder.hasNext()) {
        const file = filesInFolder.next();
        foundFiles.push(file);
        Logger.log(`Found file "${filename}" in folder: ${folder.getName()} (ID: ${file.getId()})`);
      }
    } catch (folderError) {
      Logger.log(`Error searching in folder ${folder.getName()}: ${folderError.toString()}`);
    }
  }
  
  return foundFiles;
}

/**
 * Attempts to find and recover the correct file ID for a given filename
 * Enhanced to handle data inconsistencies where buffer sheet might have wrong filename
 * @param {string} companyName - Company name
 * @param {string} changedFilename - The filename to search for
 * @param {string} currentFileId - The current file ID (may be invalid)
 * @returns {Object} Object containing found file and its location info
 */
function findAndRecoverFile(companyName, changedFilename, currentFileId) {
  Logger.log(`=== ATTEMPTING FILE RECOVERY for "${changedFilename}" ===`);
  Logger.log(`Current file ID: ${currentFileId}`);
  
  // Load existing company folder mappings from script properties
  loadCompanyFolderMappings();

  // Ensure company folder exists (create if needed)
  if (!ATTACHMENT_COMPANY_FOLDER_MAP[companyName]) {
    Logger.log(`Company '${companyName}' not found in folder mapping. Creating new folder structure...`);
    
    try {
      const newCompanyFolderId = createCompanyFolderStructure(companyName);
      Logger.log(`Successfully created company folder structure for '${companyName}' with ID: ${newCompanyFolderId}`);
    } catch (createError) {
      Logger.log(`Error: Failed to create folder structure for company '${companyName}': ${createError.message}`);
      return {
        file: null,
        fileId: null,
        location: null,
        financialYear: null,
        foundIn: null
      };
    }
  }
  
  const result = {
    file: null,
    fileId: null,
    location: null,
    financialYear: null,
    foundIn: null
  };
  
  try {
    // First, try the current file ID if it exists and looks valid
    if (currentFileId && currentFileId !== 'DELETED' && currentFileId.trim() !== '') {
      try {
        const file = DriveApp.getFileById(currentFileId.trim());
        // Accept the file even if the name doesn't match exactly (data inconsistency handling)
        result.file = file;
        result.fileId = file.getId();
        result.foundIn = 'current_id';
        Logger.log(`File found using current ID: ${file.getName()} (searched for: ${changedFilename})`);
        return result;
      } catch (idError) {
        Logger.log(`Current file ID ${currentFileId} is invalid: ${idError.toString()}`);
      }
    }
    
    // Search across multiple financial years
    const currentFinancialYear = calculateFinancialYear(new Date());
    const years = [currentFinancialYear];
    
    // Add previous and next financial years
    const currentYearNum = parseInt(currentFinancialYear.split('-')[1]);
    years.push(`FY-${(currentYearNum - 1).toString().padStart(2, '0')}-${currentYearNum.toString().padStart(2, '0')}`);
    years.push(`FY-${currentYearNum.toString().padStart(2, '0')}-${(currentYearNum + 1).toString().padStart(2, '0')}`);
    
    // Search in Buffer/Active folders
    for (const year of years) {
      try {
        const bufferActiveFolder = getOrCreateBufferSubfolder(companyName, year, "Active");
        
        // First try exact filename match
        const files = bufferActiveFolder.getFilesByName(changedFilename);
        if (files.hasNext()) {
          const file = files.next();
          result.file = file;
          result.fileId = file.getId();
          result.financialYear = year;
          result.location = 'buffer_active';
          result.foundIn = `buffer_active_${year}`;
          Logger.log(`File found in Buffer/Active for ${year}: ${file.getId()}`);
          return result;
        }
        
        // If exact match fails and changedFilename looks like an original filename,
        // try to find files with AI-generated patterns
        if (isOriginalFilenameFormat(changedFilename)) {
          Logger.log(`Searching for AI-generated filename alternatives for: ${changedFilename}`);
          const allFiles = bufferActiveFolder.getFiles();
          while (allFiles.hasNext()) {
            const file = allFiles.next();
            const fileName = file.getName();
            // Look for files with AI-generated pattern (YYYY-MM-DD_Vendor_Invoice_Amount.ext)
            if (isAIGeneratedFilenameFormat(fileName)) {
              Logger.log(`Found potential AI-generated file: ${fileName}`);
              result.file = file;
              result.fileId = file.getId();
              result.financialYear = year;
              result.location = 'buffer_active';
              result.foundIn = `buffer_active_${year}_pattern_match`;
              Logger.log(`File found by pattern matching in Buffer/Active for ${year}: ${file.getId()}`);
              return result;
            }
          }
        }
        
      } catch (yearError) {
        Logger.log(`Error checking Buffer/Active for ${year}: ${yearError.toString()}`);
      }
    }
    
    // Search in Buffer/Deleted folders
    for (const year of years) {
      try {
        const bufferDeletedFolder = getOrCreateBufferSubfolder(companyName, year, "Deleted");
        
        // First try exact filename match
        const files = bufferDeletedFolder.getFilesByName(changedFilename);
        if (files.hasNext()) {
          const file = files.next();
          result.file = file;
          result.fileId = file.getId();
          result.financialYear = year;
          result.location = 'buffer_deleted';
          result.foundIn = `buffer_deleted_${year}`;
          Logger.log(`File found in Buffer/Deleted for ${year}: ${file.getId()}`);
          return result;
        }
        
        // If exact match fails and changedFilename looks like an original filename,
        // try to find files with AI-generated patterns
        if (isOriginalFilenameFormat(changedFilename)) {
          Logger.log(`Searching for AI-generated filename alternatives in Deleted for: ${changedFilename}`);
          const allFiles = bufferDeletedFolder.getFiles();
          while (allFiles.hasNext()) {
            const file = allFiles.next();
            const fileName = file.getName();
            // Look for files with AI-generated pattern (YYYY-MM-DD_Vendor_Invoice_Amount.ext)
            if (isAIGeneratedFilenameFormat(fileName)) {
              Logger.log(`Found potential AI-generated file in Deleted: ${fileName}`);
              result.file = file;
              result.fileId = file.getId();
              result.financialYear = year;
              result.location = 'buffer_deleted';
              result.foundIn = `buffer_deleted_${year}_pattern_match`;
              Logger.log(`File found by pattern matching in Buffer/Deleted for ${year}: ${file.getId()}`);
              return result;
            }
          }
        }
        
      } catch (yearError) {
        Logger.log(`Error checking Buffer/Deleted for ${year}: ${yearError.toString()}`);
      }
    }
    
    // Search in inflow/outflow folders
    for (const year of years) {
      try {
        const currentDate = new Date();
        const month = getMonthFromDate(currentDate);
        
        ['inflow', 'outflow'].forEach(flowType => {
          try {
            const flowFolder = createFlowFolderStructure(companyName, year, month, flowType);
            const files = flowFolder.getFilesByName(changedFilename);
            
            if (files.hasNext()) {
              const file = files.next();
              result.file = file;
              result.fileId = file.getId();
              result.financialYear = year;
              result.location = flowType;
              result.foundIn = `${flowType}_${year}_${month}`;
              Logger.log(`File found in ${flowType} for ${year}/${month}: ${file.getId()}`);
              return result;
            }
          } catch (flowError) {
            Logger.log(`Error checking ${flowType} for ${year}: ${flowError.toString()}`);
          }
        });
      } catch (yearError) {
        Logger.log(`Error checking flow folders for ${year}: ${yearError.toString()}`);
      }
    }
    
    Logger.log(`File "${changedFilename}" not found in any location`);
    return result;
    
  } catch (error) {
    Logger.log(`Error in findAndRecoverFile: ${error.toString()}`);
    return result;
  }
}

/**
 * Debug function to log folder structure and file locations
 * @param {string} companyName - Company name
 * @param {string} financialYear - Financial year
 * @param {string} filename - Filename to search for
 */
function debugFileFinding(companyName, financialYear, filename) {
  Logger.log(`=== DEBUG: Searching for file "${filename}" in ${companyName} ${financialYear} ===`);
  
  try {
    // Check Buffer/Active
    const bufferActive = getOrCreateBufferSubfolder(companyName, financialYear, "Active");
    Logger.log(`Buffer/Active folder: ${bufferActive.getName()} (ID: ${bufferActive.getId()})`);
    const activeFiles = bufferActive.getFilesByName(filename);
    let activeCount = 0;
    while (activeFiles.hasNext()) {
      activeFiles.next();
      activeCount++;
    }
    Logger.log(`Files in Buffer/Active: ${activeCount}`);
    
    // Check Buffer/Deleted
    const bufferDeleted = getOrCreateBufferSubfolder(companyName, financialYear, "Deleted");
    Logger.log(`Buffer/Deleted folder: ${bufferDeleted.getName()} (ID: ${bufferDeleted.getId()})`);
    const deletedFiles = bufferDeleted.getFilesByName(filename);
    let deletedCount = 0;
    while (deletedFiles.hasNext()) {
      deletedFiles.next();
      deletedCount++;
    }
    Logger.log(`Files in Buffer/Deleted: ${deletedCount}`);
    
    // Check inflow/outflow folders
    const currentDate = new Date();
    const month = getMonthFromDate(currentDate);
    
    ['inflow', 'outflow'].forEach(flowType => {
      try {
        const flowFolder = createFlowFolderStructure(companyName, financialYear, month, flowType);
        Logger.log(`${flowType} folder: ${flowFolder.getName()} (ID: ${flowFolder.getId()})`);
        const flowFiles = flowFolder.getFilesByName(filename);
        let flowCount = 0;
        while (flowFiles.hasNext()) {
          flowFiles.next();
          flowCount++;
        }
        Logger.log(`Files in ${flowType}: ${flowCount}`);
      } catch (flowError) {
        Logger.log(`Error checking ${flowType} folder: ${flowError.toString()}`);
      }
    });
    
  } catch (debugError) {
    Logger.log(`Debug error: ${debugError.toString()}`);
  }
  
  Logger.log(`=== END DEBUG ===`);
}

/**
 * SUPPLEMENTARY THREAD RECOGNITION FUNCTIONS
 * These functions add thread context without modifying existing functionality
 */

/**
 * Analyzes Gmail thread to extract sender and conversation context (supplementary)
 * @param {GoogleAppsScript.Gmail.GmailThread} thread - Gmail thread object
 * @returns {Object} Thread analysis with sender info and context
 */
function getThreadContext(thread) {
  try {
    const messages = thread.getMessages();
    const threadId = thread.getId();
    const firstMessage = messages[0];
    
    // Extract sender information
    const primarySender = firstMessage.getFrom();
    const senders = new Set();
    let totalAttachments = 0;
    
    // Analyze all messages in thread
    messages.forEach(message => {
      senders.add(message.getFrom());
      totalAttachments += message.getAttachments().filter(a => !a.isGoogleType() && !a.getName().startsWith('ATT')).length;
    });
    
    // Extract sender details
    const senderEmail = primarySender.match(/<(.+?)>/)?.[1] || primarySender;
    const senderName = primarySender.replace(/<.+?>/, '').trim().replace(/['"]/g, '');
    
    // Create thread context (supplementary information)
    return {
      threadId: threadId,
      messageCount: messages.length,
      senderName: senderName,
      senderEmail: senderEmail,
      isSingleSender: senders.size === 1,
      isRecurringConversation: messages.length > 1,
      totalAttachments: totalAttachments,
      threadLabel: `${senderName.replace(/\s+/g, '')}_${threadId.substring(0, 6)}`
    };
    
  } catch (error) {
    Logger.log(`Thread context extraction error: ${error.toString()}`);
    return {
      threadId: thread.getId(),
      messageCount: 0,
      senderName: 'Unknown',
      senderEmail: 'unknown@example.com',
      isSingleSender: false,
      isRecurringConversation: false,
      totalAttachments: 0,
      threadLabel: `Unknown_${thread.getId().substring(0, 6)}`
    };
  }
}

/**
 * Logs thread context for analysis (supplementary)
 * @param {string} threadId - Thread ID
 * @param {Object} context - Thread context
 * @param {string} status - Processing status
 */
function logThreadContext(threadId, context, status) {
  THREAD_CONTEXT_LOG[threadId] = {
    ...context,
    status: status,
    processedAt: new Date(),
    attachmentsProcessed: 0,
    attachmentsSkipped: 0
  };
}

/**
 * Updates thread processing stats (supplementary)
 * @param {string} threadId - Thread ID
 * @param {string} action - Action type
 */
function updateThreadStats(threadId, action) {
  if (THREAD_CONTEXT_LOG[threadId]) {
    if (action === 'processed') {
      THREAD_CONTEXT_LOG[threadId].attachmentsProcessed++;
    } else if (action === 'skipped') {
      THREAD_CONTEXT_LOG[threadId].attachmentsSkipped++;
    }
  }
}

/**
 * Gets thread summary for logging (supplementary)
 * @returns {Object} Thread processing summary
 */
function getThreadSummary() {
  const threads = Object.values(THREAD_CONTEXT_LOG);
  const singleSenderThreads = threads.filter(t => t.isSingleSender);
  const recurringConversations = threads.filter(t => t.isRecurringConversation);
  
  return {
    totalThreads: threads.length,
    singleSenderThreads: singleSenderThreads.length,
    recurringConversations: recurringConversations.length,
    topSenders: threads.map(t => ({ name: t.senderName, emails: t.messageCount, attachments: t.totalAttachments }))
      .sort((a, b) => b.attachments - a.attachments)
      .slice(0, 5)
  };
}

/**
 * Checks if a filename appears to be in original format (not AI-generated)
 * @param {string} filename - The filename to check
 * @returns {boolean} True if it looks like an original filename
 */
function isOriginalFilenameFormat(filename) {
  if (!filename) return false;
  
  // AI-generated filenames follow pattern: YYYY-MM-DD_Vendor_Invoice_Amount.ext
  // Original filenames typically don't follow this strict pattern
  const aiPattern = /^\d{4}-\d{2}-\d{2}_[^_]+_[^_]+_[\d,.]+\.[a-zA-Z]+$/;
  
  // If it matches AI pattern, it's NOT an original filename
  if (aiPattern.test(filename)) {
    return false;
  }
  
  // If it has spaces, mixed case, or doesn't follow the strict AI pattern, 
  // it's likely an original filename
  return true;
}

/**
 * Checks if a filename appears to be AI-generated
 * @param {string} filename - The filename to check  
 * @returns {boolean} True if it looks like an AI-generated filename
 */
function isAIGeneratedFilenameFormat(filename) {
  if (!filename) return false;
  
  // AI-generated filenames follow pattern: YYYY-MM-DD_Vendor_Invoice_Amount.ext
  const aiPattern = /^\d{4}-\d{2}-\d{2}_[^_]+_[^_]+_[\d,.]+\.[a-zA-Z]+$/;
  
  return aiPattern.test(filename);
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

/**
 * Test function to verify AI classification and file renaming is working
 * Use this to test with a sample file or debug issues
 * @param {string} testFileName - Name of test file
 * @returns {Object} Test result with classification and filename generation
 */
function testAIClassificationAndRenaming(testFileName = "Sample Invoice.pdf") {
  Logger.log(`=== TESTING AI CLASSIFICATION AND RENAMING ===`);
  Logger.log(`Test filename: ${testFileName}`);
  
  try {
    // Test fallback extraction (simulates when AI is not available)
    const fallbackResult = fallbackDataExtraction(null, testFileName);
    Logger.log(`\n--- FALLBACK CLASSIFICATION ---`);
    Logger.log(`Status: ${fallbackResult.invoiceStatus}`);
    Logger.log(`Document Type: ${fallbackResult.documentType}`);
    Logger.log(`Vendor: ${fallbackResult.vendorName}`);
    Logger.log(`Invoice Number: ${fallbackResult.invoiceNumber}`);
    
    // Test filename generation
    const generatedFilename = generateNewFilename(fallbackResult, testFileName);
    Logger.log(`\n--- FILENAME GENERATION ---`);
    Logger.log(`Original: ${testFileName}`);
    Logger.log(`Generated: ${generatedFilename}`);
    
    // Test status validation
    const validatedStatus = validateAndNormalizeInvoiceStatus(fallbackResult.invoiceStatus, fallbackResult.isFinancialDocument);
    Logger.log(`\n--- STATUS VALIDATION ---`);
    Logger.log(`Original Status: ${fallbackResult.invoiceStatus}`);
    Logger.log(`Validated Status: ${validatedStatus}`);
    Logger.log(`Is Financial: ${fallbackResult.isFinancialDocument}`);
    
    // Create summary
    const summary = {
      originalFilename: testFileName,
      generatedFilename: generatedFilename,
      classification: validatedStatus,
      vendor: fallbackResult.vendorName,
      invoiceNumber: fallbackResult.invoiceNumber,
      amount: fallbackResult.amount,
      isRenamed: testFileName !== generatedFilename,
      isClassified: validatedStatus !== 'unknown'
    };
    
    Logger.log(`\n--- TEST SUMMARY ---`);
    Logger.log(`✓ File will be renamed: ${summary.isRenamed}`);
    Logger.log(`✓ File will be classified: ${summary.isClassified}`);
    Logger.log(`✓ Final classification: ${summary.classification}`);
    Logger.log(`✓ Routing: ${summary.classification === 'irrelevant' || summary.classification === 'unknown' ? 'Buffer2' : 'Main Buffer → ' + summary.classification + ' folder'}`);
    
    Logger.log(`=== END TEST ===`);
    
    return summary;
    
  } catch (error) {
    Logger.log(`Test failed: ${error.toString()}`);
    return { error: error.toString() };
  }
}

/**
 * Debugging function to check recent processed files
 * @param {string} companyName - Company to check
 */
function debugRecentFiles(companyName) {
  Logger.log(`=== DEBUGGING RECENT FILES FOR ${companyName.toUpperCase()} ===`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bufferSheet = ss.getSheetByName(`${companyName}-buffer`);
  
  if (!bufferSheet) {
    Logger.log(`No buffer sheet found for ${companyName}`);
    return;
  }
  
  const data = bufferSheet.getDataRange().getValues();
  const headers = data[0];
  Logger.log(`Buffer sheet headers: ${headers.join(', ')}`);
  
  if (data.length <= 1) {
    Logger.log(`No files processed yet for ${companyName}`);
    return;
  }
  
  // Show last 5 processed files
  const recentFiles = data.slice(-5).reverse(); // Last 5, newest first
  
  Logger.log(`\n--- RECENT PROCESSED FILES ---`);
  recentFiles.forEach((row, index) => {
    const originalFilename = row[1];
    const changedFilename = row[2];
    const invoiceId = row[3];
    const driveFileId = row[4];
    const reason = row[6];
    const status = row[7];
    const ui = row[8];
    
    Logger.log(`\n${index + 1}. ${originalFilename}`);
    Logger.log(`   → Renamed to: ${changedFilename}`);
    Logger.log(`   → Invoice ID: ${invoiceId}`);
    Logger.log(`   → Status: ${status}`);
    Logger.log(`   → UI: ${ui}`);
    Logger.log(`   → Classification: ${reason}`);
    Logger.log(`   → File ID: ${driveFileId}`);
    Logger.log(`   → Renamed: ${originalFilename !== changedFilename ? 'YES' : 'NO'}`);
  });
  
  Logger.log(`=== END DEBUG ===`);
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

/**
 * Helper function to validate Drive permissions and folder access
 * @param {string} folderId - Folder ID to test access
 * @returns {boolean} True if access is available
 */
function validateDriveAccess(folderId) {
  try {
    if (!folderId || folderId.trim() === '') {
      Logger.log('Invalid folder ID provided for validation');
      return false;
    }
    
    // Try to access the folder
    const folder = DriveApp.getFolderById(folderId);
    const folderName = folder.getName(); // This will throw if no access
    Logger.log(`Drive access validated for folder: ${folderName} (ID: ${folderId})`);
    return true;
  } catch (error) {
    Logger.log(`Drive access validation failed for folder ID ${folderId}: ${error.toString()}`);
    return false;
  }
}

/**
 * Extract company name from changed filename
 * Changed filenames typically start with company name (e.g., "analogy_2024_invoice.pdf")
 * @param {string} changedFilename - The AI-generated filename
 * @returns {string|null} Company name or null if not found
 */
function extractCompanyFromFilename(changedFilename) {
  if (!changedFilename || typeof changedFilename !== 'string') {
    return null;
  }
  
  // Convert to lowercase for matching
  const lowerFilename = changedFilename.toLowerCase();
  
  // Ensure company mappings are loaded
  loadCompanyFolderMappings();
  
  // Check against known company names in the mapping
  const knownCompanies = Object.keys(ATTACHMENT_COMPANY_FOLDER_MAP);
  for (const company of knownCompanies) {
    if (lowerFilename.startsWith(company.toLowerCase())) {
      Logger.log(`Found company '${company}' in filename: ${changedFilename}`);
      return company;
    }
  }
  
  // Try to extract company name from common patterns
  // Pattern 1: company_year_document.ext
  const underscorePattern = changedFilename.split('_')[0];
  if (underscorePattern && underscorePattern.length > 2) {
    // Check if this matches any known company
    for (const company of knownCompanies) {
      if (underscorePattern.toLowerCase() === company.toLowerCase()) {
        Logger.log(`Found company '${company}' from underscore pattern in: ${changedFilename}`);
        return company;
      }
    }
  }
  
  // Pattern 2: company-year-document.ext
  const dashPattern = changedFilename.split('-')[0];
  if (dashPattern && dashPattern.length > 2) {
    // Check if this matches any known company
    for (const company of knownCompanies) {
      if (dashPattern.toLowerCase() === company.toLowerCase()) {
        Logger.log(`Found company '${company}' from dash pattern in: ${changedFilename}`);
        return company;
      }
    }
  }
  
  Logger.log(`Could not extract company name from filename: ${changedFilename}`);
  Logger.log(`Available companies: ${knownCompanies.join(', ')}`);
  return null;
}

/**
 * Get or create company name mapping, with fallback to create new company
 * @param {string} companyName - Company name to find or create
 * @returns {string} Valid company name that exists in mappings
 */
function getOrCreateCompanyMapping(companyName) {
  if (!companyName || typeof companyName !== 'string') {
    return null;
  }
  
  // Ensure company mappings are loaded
  loadCompanyFolderMappings();
  
  // Check if company already exists (case-insensitive)
  const knownCompanies = Object.keys(ATTACHMENT_COMPANY_FOLDER_MAP);
  for (const company of knownCompanies) {
    if (company.toLowerCase() === companyName.toLowerCase()) {
      Logger.log(`Found existing company mapping: ${company}`);
      return company;
    }
  }
  
  // Company doesn't exist, try to create it
  try {
    Logger.log(`Creating new company mapping for: ${companyName}`);
    const newCompanyFolderId = createCompanyFolderStructure(companyName);
    Logger.log(`Successfully created company folder structure for: ${companyName} with ID: ${newCompanyFolderId}`);
    return companyName;
  } catch (createError) {
    Logger.log(`Failed to create company folder structure for ${companyName}: ${createError.toString()}`);
    return null;
  }
}

/**
 * Function to test and request Drive permissions
 * Call this function once to ensure proper permissions are granted
 */
function requestDrivePermissions() {
  try {
    Logger.log('Testing Drive permissions...');
    
    // Test basic Drive access
    const folders = DriveApp.getFolders();
    if (folders.hasNext()) {
      const testFolder = folders.next();
      Logger.log(`✓ Basic Drive access granted. Test folder: ${testFolder.getName()}`);
    }
    
    // Test creating a temporary folder (will be deleted)
    const tempFolder = DriveApp.createFolder('TempPermissionTest_' + Date.now());
    Logger.log(`✓ Drive write access granted. Created temp folder: ${tempFolder.getId()}`);
    
    // Clean up
    DriveApp.getFolderById(tempFolder.getId()).setTrashed(true);
    Logger.log('✓ Drive permissions test completed successfully');
    
    return true;
  } catch (error) {
    Logger.log(`✗ Drive permissions test failed: ${error.toString()}`);
    Logger.log('Please ensure the following permissions are granted:');
    Logger.log('- https://www.googleapis.com/auth/drive');
    Logger.log('- https://www.googleapis.com/auth/drive.file');
    return false;
  }
}
  
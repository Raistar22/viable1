// Global variable to track cancellation status
var CANCELLATION_TOKEN = null;

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
 * @param {Sheet} logSheet The sheet containing processed file logs.
 * @param {number} columnIndex The 0-indexed column number to read IDs from (e.g., 1 for File ID, 8 for Gmail Message ID).
 * @returns {Set} A set of processed IDs.
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
 * Scans a specific Gmail label for attachments and saves them to a Drive folder.
 * Also logs processed files to a sub-sheet.
 * Handles initial population from Drive folder and then processes new emails.
 * @param {string} labelName The name of the Gmail label (company name).
 * @param {string} processToken Unique token to identify this process for cancellation.
 * @returns {Object} An object with status and message.
 */
function processAttachments(labelName, processToken) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Clear any existing cancellation token and set new one
  clearCancellationToken();

  // --- IMPORTANT: CONFIGURE YOUR DRIVE FOLDER IDs HERE ---
  var companyFolderMap = {
    "analogy": "1QKZZ-WXTEK9e7UkEs0Xz4X-NgcBhjyaY", // Replace with your actual ID
    "humane": "1nD_ng51OUCAJKtW132O5bIabSp_bfBHW", // Replace with your actual ID
    "buffer": "1uUDIHEkyWfYPFMKNbMaehQJ_pBUmyBSn", // Replace with your actual ID
    // Add more mappings as needed: "Gmail Label Name": "Drive Folder Name",
  };
  // --------------------------------------------------------

  var folderId = companyFolderMap[labelName];

  if (!folderId) {
    return { status: 'error', message: "Error: No Drive folder configured for label '" + labelName + "'. Please update the script." };
  }

  let logSheet;
  // Standard headers reflecting Drive metadata + original email context
  const downloaderHeaders = [
    'File Name', 'File ID', 'File URL', 
    'Date Created (Drive)', 'Last Updated (Drive)', 'Size (bytes)', 'Mime Type',
    'Email Subject', 'Gmail Message ID','invoice status'
  ];
  // Reroute.js compatible headers
  const rerouteHeaders = [
    'Date', 'Month', 'Vendor Name', 'Financial Year', 'Document Link',
    'Document Number', 'Gross Amount', 'GST', 'TDS', 'Other Taxes', 'Net Amount', 'Redundant'
  ];
  // Combine all unique headers, preserving order: downloaderHeaders first, then any rerouteHeaders not already present
  const allHeaders = downloaderHeaders.concat(rerouteHeaders.filter(h => !downloaderHeaders.includes(h)));

  try {
    logSheet = ss.getSheetByName(labelName);

    // If sheet doesn't exist, create it
    if (!logSheet) {
      logSheet = ss.insertSheet(labelName);
      Logger.log(`Created new sheet: ${labelName}`);
    }

    // Check if headers need to be added (add only missing headers, do not remove any)
    let currentHeaders = [];
    if (logSheet.getLastRow() > 0) {
      const firstRowRange = logSheet.getRange(1, 1, 1, logSheet.getLastColumn());
      currentHeaders = firstRowRange.getValues()[0];
    }
    // Add any missing headers from allHeaders
    let headersToSet = currentHeaders.slice();
    allHeaders.forEach(h => {
      if (!headersToSet.includes(h)) headersToSet.push(h);
    });
    // Only set headers if there are new ones to add
    if (headersToSet.length > currentHeaders.length) {
      logSheet.getRange(1, 1, 1, headersToSet.length).setValues([headersToSet]);
      // Optionally format new headers
      const headerRange = logSheet.getRange(1, 1, 1, headersToSet.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#E8F0FE');
      headerRange.setBorder(true, true, true, true, true, true);
      logSheet.setFrozenRows(1);
    }

  } catch (e) {
    return { status: 'error', message: "Error setting up log sheet for " + labelName + ": " + e.toString() };
  }

  try {
    const folder = DriveApp.getFolderById(folderId);
    const label = GmailApp.getUserLabelByName(labelName);

    if (!label) {
      return { status: 'error', message: "Error: Gmail label '" + labelName + "' not found." };
    }

    // Check for cancellation before starting
    if (shouldCancel(processToken)) {
      return { status: 'cancelled', message: "Process cancelled by user before starting." };
    }

    // --- Step 1: Preemptive Scan of Drive Folder and Log Population ---
    const loggedDriveFileIds = getProcessedLogEntryIds(logSheet, 1); // File ID is at index 1
    let filesLoggedFromDrive = 0;

    // Collect existing file names in the target Drive folder for quick lookup
    const existingDriveFileNames = new Set();
    const driveFilesIterator = folder.getFiles(); // Get a fresh iterator for the Drive folder
    while (driveFilesIterator.hasNext()) {
      // Check for cancellation during Drive scan
      if (shouldCancel(processToken)) {
        return { status: 'cancelled', message: "Process cancelled during Drive folder scan." };
      }

      const driveFile = driveFilesIterator.next();
      if (!loggedDriveFileIds.has(driveFile.getId())) {
        // This file exists in Drive but is not in our log sheet, add it.
        logSheet.appendRow([
          driveFile.getName(),
          driveFile.getId(),
          driveFile.getUrl(),
          driveFile.getDateCreated(),
          driveFile.getLastUpdated(),
          driveFile.getSize(),
          driveFile.getMimeType(),
          '', // Email Subject (blank for Drive-scanned files)
          ''  // Gmail Message ID (blank for Drive-scanned files)
        ]);
        loggedDriveFileIds.add(driveFile.getId()); // Add to set to prevent re-logging in this run
        filesLoggedFromDrive++;
      }
      existingDriveFileNames.add(driveFile.getName()); // Add name for quick duplicate checking later
      Utilities.sleep(50); // Small pause
    }
    Logger.log(`Preemptive Drive scan for '${labelName}': Logged ${filesLoggedFromDrive} existing files. Found ${existingDriveFileNames.size} unique file names in Drive folder.`);

    // Check for cancellation after Drive scan
    if (shouldCancel(processToken)) {
      return { status: 'cancelled', message: "Process cancelled after Drive folder scan." };
    }

    // --- Step 2: Process New Gmail Attachments ---
    const processedGmailMessageIds = getProcessedLogEntryIds(logSheet, 8); // Gmail Message ID is at index 8
    let totalNewAttachments = 0;
    let processedAttachments = 0;
    let skippedAttachments = 0; // Skipped because email already processed or Drive duplicate

    // First pass to count total NEW attachments for accurate progress
    const threads = label.getThreads();
    for (let t = 0; t < threads.length; t++) {
      // Check for cancellation during counting
      if (shouldCancel(processToken)) {
        return { status: 'cancelled', message: "Process cancelled during attachment counting." };
      }

      const thread = threads[t];
      const messages = thread.getMessages();
      
      for (let m = 0; m < messages.length; m++) {
        const message = messages[m];
        const messageId = message.getId();
        
        // Skip if message was already processed (based on Gmail Message ID in log)
        if (processedGmailMessageIds.has(messageId)) {
          // This message has been processed before, so its attachments should be too.
          const attachments = message.getAttachments();
          for (let a = 0; a < attachments.length; a++) {
            const attachment = attachments[a];
            if (!attachment.isGoogleType() && !attachment.getName().startsWith('ATT')) {
              skippedAttachments++;
            }
          }
          continue; // Skip this entire message
        }

        // Only count attachments if the message itself is new to processing
        const attachments = message.getAttachments();
        for (let a = 0; a < attachments.length; a++) {
          const attachment = attachments[a];
          if (!attachment.isGoogleType() && !attachment.getName().startsWith('ATT')) {
            // Further check: Is a file with this name already in Drive?
            // This counts it as 'new' for total, but 'skipped' if it's a Drive duplicate later.
            totalNewAttachments++;
          }
        }
      }
      
      // Small pause during counting to allow cancellation checks
      if (t % 10 === 0) {
        Utilities.sleep(100);
      }
    }

    // If no new attachments, return early (after Drive scan)
    if (totalNewAttachments === 0) {
      let message = `No new email attachments found for '${labelName}'.`;
      if (filesLoggedFromDrive > 0) {
          message += ` (Logged ${filesLoggedFromDrive} files found during Drive scan)`;
      }
      if (skippedAttachments > 0) {
        message += ` (${skippedAttachments} email attachments skipped as their messages were previously processed)`;
      }
      clearCancellationToken();
      return { status: 'success', message: message };
    }

    // Second pass to process and log NEW attachments only
    for (let t = 0; t < threads.length; t++) {
      // Check for cancellation before processing each thread
      if (shouldCancel(processToken)) {
        clearCancellationToken();
        return { status: 'cancelled', message: `Process cancelled. Processed ${processedAttachments} of ${totalNewAttachments} attachments before cancellation.` };
      }

      const thread = threads[t];
      const messages = thread.getMessages();
      
      for (let m = 0; m < messages.length; m++) {
        const message = messages[m];
        const messageId = message.getId();
        
        // Skip if message was already processed (to avoid re-downloading)
        if (processedGmailMessageIds.has(messageId)) {
          continue;
        }

        const attachments = message.getAttachments();
        let attachmentsProcessedForThisMessage = 0;
        
        for (let a = 0; a < attachments.length; a++) {
          // Check for cancellation before processing each attachment
          if (shouldCancel(processToken)) {
            clearCancellationToken();
            return { status: 'cancelled', message: `Process cancelled. Processed ${processedAttachments} of ${totalNewAttachments} attachments before cancellation.` };
          }

          const attachment = attachments[a];
          if (!attachment.isGoogleType() && !attachment.getName().startsWith('ATT')) {
            try {
              let fileToLog; // This will be the actual Drive file object to be logged

              // NEW DUPLICATE CHECK: Does a file with this exact name already exist in the Drive folder?
              if (existingDriveFileNames.has(attachment.getName())) {
                Logger.log(`Skipping upload: Duplicate file name '${attachment.getName()}' already exists in Drive folder for '${labelName}'.`);
                skippedAttachments++;
                
                // If we skip upload, we should still log the existing file's details if not already logged
                const existingFiles = folder.getFilesByName(attachment.getName());
                if (existingFiles.hasNext()) {
                  fileToLog = existingFiles.next(); // Get the first found existing file
                  if (!loggedDriveFileIds.has(fileToLog.getId())) {
                     // Log the existing Drive file's details
                     logSheet.appendRow([
                       fileToLog.getName(),
                       fileToLog.getId(),
                       fileToLog.getUrl(),
                       fileToLog.getDateCreated(),
                       fileToLog.getLastUpdated(),
                       fileToLog.getSize(),
                       fileToLog.getMimeType(),
                       message.getSubject(), // Still link to email subject for context
                       messageId             // Still link to email message ID for context
                     ]);
                     loggedDriveFileIds.add(fileToLog.getId()); // Add to our set of logged Drive IDs
                     attachmentsProcessedForThisMessage++; // Count as processed for progress, even if skipped upload
                  } else {
                     Logger.log(`Existing file '${fileToLog.getName()}' (ID: ${fileToLog.getId()}) already logged. No action needed.`);
                  }
                }
              } else {
                // No duplicate found by name, proceed with uploading
                fileToLog = folder.createFile(attachment); // Upload the new attachment
                processedAttachments++;
                attachmentsProcessedForThisMessage++;

                // Log the newly uploaded file's details
                logSheet.appendRow([
                  fileToLog.getName(),
                  fileToLog.getId(),
                  fileToLog.getUrl(),
                  fileToLog.getDateCreated(),
                  fileToLog.getLastUpdated(),
                  fileToLog.getSize(),
                  fileToLog.getMimeType(),
                  message.getSubject(), // Email Subject
                  messageId             // Gmail Message ID
                ]);
                loggedDriveFileIds.add(fileToLog.getId()); // Add to our set of logged Drive IDs
                existingDriveFileNames.add(fileToLog.getName()); // Add to list of existing names for future checks
              }

              // Send progress update to the UI
              try {
                google.script.run.withSuccessHandler(function(){}) // Empty handler as we just send
                  .updateProgress(processedAttachments, totalNewAttachments, labelName, processToken);
              } catch (progressError) {
                Logger.log("Error sending progress update: " + progressError.toString());
              }

              // Small pause to prevent rate limiting and allow cancellation checks
              Utilities.sleep(100);
              
            } catch (fileError) {
              Logger.log("Error processing attachment '" + attachment.getName() + "': " + fileError.toString());
              // Continue processing other attachments even if one fails
            }
          }
        }

        // After attempting to process all attachments for a message, add its ID to the processed set.
        // This ensures the entire message isn't scanned again, even if some attachments were duplicates or failed.
        if (attachmentsProcessedForThisMessage > 0 || attachments.length === 0) { // If message had attachments we dealt with, or no attachments at all
            processedGmailMessageIds.add(messageId);
        }
      }
    }

    // Clear cancellation token on successful completion
    clearCancellationToken();

    let resultMessage = `Completed processing for '${labelName}'. `;
    if (filesLoggedFromDrive > 0) {
        resultMessage += `Logged ${filesLoggedFromDrive} existing files from Drive folder. `;
    }
    resultMessage += `Processed ${processedAttachments} new email attachments uploaded to Drive. `;
    resultMessage += `Skipped ${skippedAttachments} attachments (already processed or duplicate in Drive).`;

    return { status: 'success', message: resultMessage };

  } catch (e) {
    clearCancellationToken();
    return { status: 'error', message: "An error occurred: " + e.toString() };
  }
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

/**
 * Get all company names from sheet tabs (Reroute.js compatible)
 * @return {Array} Array of company names
 */
function getCompanies() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const companies = [];
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    // Skip inflow/outflow/system sheets
    if (!sheetName.includes(' - inflow') && !sheetName.includes(' - outflow') && sheetName !== 'Sheet1') {
      companies.push(sheetName);
    }
  });
  return companies.sort();
}

/**
 * Flood file details and create inflow/outflow sheets (Reroute.js compatible)
 * @param {string} companyName - Name of the company
 * @return {Object} Success response
 */
function floodFileDetails(companyName) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const companySheet = spreadsheet.getSheetByName(companyName);
    if (!companySheet) {
      throw new Error(`Company sheet '${companyName}' not found`);
    }
    // Get data from company sheet
    const data = companySheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    // Find required column indices
    const fileNameIndex = headers.indexOf('File Name');
    const fileUrlIndex = headers.indexOf('File URL');
    const invoiceStatusIndex = headers.indexOf('invoice status');
    if (fileNameIndex === -1 || fileUrlIndex === -1) {
      throw new Error('Required columns (File Name, File URL) not found in company sheet');
    }
    // Process each file and categorize by invoice status
    const inflowData = [];
    const outflowData = [];
    const seenInvoices = new Set();
    const inflowRedundant = [];
    const outflowRedundant = [];
    rows.forEach(row => {
      if (row[fileNameIndex]) {
        const fileName = row[fileNameIndex];
        const fileUrl = row[fileUrlIndex] || '';
        const invoiceStatus = row[invoiceStatusIndex] || 'inflow';
        const parsedData = parseFileName(fileName, fileUrl);
        const vendorName = parsedData[2] || '';
        const invoiceNumber = parsedData[5] || '';
        const totalAmount = parsedData[10] || '';
        const uniqueKey = invoiceNumber + '|' + vendorName + '|' + totalAmount;
        if (seenInvoices.has(uniqueKey)) {
          if (invoiceStatus.toLowerCase().includes('outflow')) {
            outflowRedundant.push([...parsedData, 'Yes']);
          } else {
            inflowRedundant.push([...parsedData, 'Yes']);
          }
          return;
        }
        seenInvoices.add(uniqueKey);
        if (invoiceStatus.toLowerCase().includes('outflow')) {
          outflowData.push([...parsedData, 'No']);
        } else {
          inflowData.push([...parsedData, 'No']);
        }
      }
    });
    inflowData.push(...inflowRedundant);
    outflowData.push(...outflowRedundant);
    createOrUpdateFlowSheet(companyName + ' - inflow', inflowData);
    createOrUpdateFlowSheet(companyName + ' - outflow', outflowData);
    return {
      success: true,
      message: `Processed ${inflowData.length} inflow and ${outflowData.length} outflow records`
    };
  } catch (error) {
    Logger.log('Error flooding file details: ' + error);
    throw new Error('Failed to flood file details: ' + error.message);
  }
}

/**
 * Parse file name and extract relevant information (Reroute.js compatible)
 * @param {string} fileName - File name in format Date_VendorName_InvoiceNumber_TotalAmount
 * @param {string} fileUrl - File URL
 * @return {Array} Parsed data array
 */
function parseFileName(fileName, fileUrl) {
  try {
    const parts = fileName.split('_');
    if (parts.length < 4) {
      return [
        '', '', fileName, '', fileUrl, '', '', '', '', '', ''
      ];
    }
    const date = parts[0];
    const vendorName = parts[1];
    const invoiceNumber = parts[2];
    const totalAmount = parts[3].replace(/\.[^.]*$/, '');
    const month = getMonthFromDate(date);
    const financialYear = calculateFinancialYear(date);
    return [
      date, month, vendorName, financialYear, fileUrl, invoiceNumber, '', '', '', '', totalAmount
    ];
  } catch (error) {
    Logger.log('Error parsing filename: ' + fileName + ' ' + error);
    return [ '', '', fileName, '', fileUrl, '', '', '', '', '', '' ];
  }
}

/**
 * Get month name from date string (YYYY-MM-DD)
 * @param {string} dateStr
 * @return {string} Month name
 */
function getMonthFromDate(dateStr) {
  try {
    const date = new Date(dateStr);
    const months = [
      'January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December'
    ];
    return months[date.getMonth()];
  } catch (error) {
    return '';
  }
}

/**
 * Calculate financial year from date (YYYY-MM-DD)
 * @param {string} dateStr
 * @return {string} Financial year in format YYYY-YYYY
 */
function calculateFinancialYear(dateStr) {
  try {
    const date = new Date(dateStr);
    const year = date.getFullYear();
    const month = date.getMonth() + 1;
    if (month >= 4) {
      return `${year}-${year + 1}`;
    } else {
      return `${year - 1}-${year}`;
    }
  } catch (error) {
    return '';
  }
}

/**
 * Create or update flow sheet with data (Reroute.js compatible)
 * @param {string} sheetName - Name of the sheet
 * @param {Array} data - Data to populate
 */
function createOrUpdateFlowSheet(sheetName, data) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
    } else {
      sheet.clear();
    }
    const headers = [
      'Date', 'Month', 'Vendor Name', 'Financial Year', 'Document Link',
      'Document Number', 'Gross Amount', 'GST', 'TDS', 'Other Taxes', 'Net Amount', 'Redundant'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    if (data && data.length > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
    }
    sheet.autoResizeColumns(1, headers.length);
  } catch (error) {
    Logger.log('Error creating/updating flow sheet: ' + error);
    throw error;
  }
}

/**
 * Shifts up all cells in the specified column if a cell is empty.
 * @param {string} sheetName - The name of the sheet to operate on.
 * @param {number} col - The 1-based column number to check (e.g., 1 for column A).
 */
function shiftUpIfCellEmpty(sheetName, col) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // No data

  var values = sheet.getRange(2, col, lastRow - 1, 1).getValues(); // skip header
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === "" || values[i][0] === null) {
      // Shift up all cells below
      for (var j = i + 1; j < values.length; j++) {
        values[j - 1][0] = values[j][0];
      }
      values[values.length - 1][0] = ""; // Clear last cell
      // Write back and exit (do one shift per call)
      sheet.getRange(2, col, values.length, 1).setValues(values);
      break;
    }
  }
}
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
  // New standard headers reflecting Drive metadata + original email context
  const desiredHeaders = [
    'File Name', 'File ID', 'File URL', 
    'Date Created (Drive)', 'Last Updated (Drive)', 'Size (bytes)', 'Mime Type',
    'Email Subject', 'Gmail Message ID','invoice status'
  ];

  try {
    logSheet = ss.getSheetByName(labelName);

    // If sheet doesn't exist, create it
    if (!logSheet) {
      logSheet = ss.insertSheet(labelName);
      Logger.log(`Created new sheet: ${labelName}`);
    }

    // Check if headers need to be added or updated
    let addHeaders = false;
    if (logSheet.getLastRow() === 0) {
      // Sheet is completely empty, definitely add headers
      addHeaders = true;
    } else {
      // Sheet has data, check if first row matches desired headers
      const firstRowRange = logSheet.getRange(1, 1, 1, desiredHeaders.length);
      const firstRowValues = firstRowRange.getValues()[0];
      
      // Compare arrays (simple string comparison for equality)
      if (firstRowValues.join(',') !== desiredHeaders.join(',')) {
        addHeaders = true;
        // Optionally, clear existing content if headers are truly malformed for a clean slate.
        // logSheet.clearContents(); 
        Logger.log(`Headers for '${labelName}' are incorrect or missing. Updating.`);
      }
    }

    if (addHeaders) {
      logSheet.getRange(1, 1, 1, desiredHeaders.length).setValues([desiredHeaders]);
      
      // Apply formatting to headers
      const headerRange = logSheet.getRange(1, 1, 1, desiredHeaders.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#E8F0FE');
      headerRange.setBorder(true, true, true, true, true, true);
      
      logSheet.setFrozenRows(1);
      // Adjust column widths for new headers
      logSheet.setColumnWidth(1, 200); // File Name
      logSheet.setColumnWidth(2, 180); // File ID
      logSheet.setColumnWidth(3, 250); // File URL
      logSheet.setColumnWidth(4, 180); // Date Created (Drive)
      logSheet.setColumnWidth(5, 180); // Last Updated (Drive)
      logSheet.setColumnWidth(6, 120); // Size (bytes)
      logSheet.setColumnWidth(7, 150); // Mime Type
      logSheet.setColumnWidth(8, 300); // Email Subject
      logSheet.setColumnWidth(9, 200); // Gmail Message ID
      Logger.log(`Headers added/updated for sheet: ${labelName}`);
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
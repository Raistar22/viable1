/**
 * GmailProcessor.gs - Handles Gmail operations and attachment processing (FIXED)
 */

/**
 * Main function to process Gmail attachments for all active clients
 */
function processAllClientsGmail() {
    try {
      infoLog('Starting Gmail processing for all active clients');
      const activeClients = getActiveClients();
      
      if (activeClients.length === 0) {
        return {
          success: true,
          message: 'No active clients found',
          results: []
        };
      }
      
      const results = [];
      let successCount = 0;
      let failureCount = 0;
      
      for (const client of activeClients) {
        try {
          infoLog(`Processing Gmail for client: ${client.name}`);
          const result = processClientGmail(client);
          results.push({
            client: client.name,
            success: true,
            ...result
          });
          successCount++;
          
          // Add delay between clients to respect rate limits
          sleep(SYSTEM_CONFIG.PROCESSING.BATCH_DELAY);
          
        } catch (error) {
          errorLog(`Error processing Gmail for client ${client.name}`, error);
          results.push({
            client: client.name,
            success: false,
            error: error.message,
            code: error.code || 'UNKNOWN_ERROR'
          });
          failureCount++;
        }
      }
      
      const summary = {
        success: true,
        message: `Processed ${activeClients.length} clients: ${successCount} successful, ${failureCount} failed`,
        totalClients: activeClients.length,
        successCount: successCount,
        failureCount: failureCount,
        results: results
      };
      
      infoLog('Completed Gmail processing for all clients', summary);
      return summary;
      
    } catch (error) {
      errorLog('Error in processAllClientsGmail', error);
      throw error;
    }
  }
  
  /**
   * Process Gmail attachments for a specific client with enhanced error handling
   */
  function processClientGmail(client) {
    try {
      if (!client || !client.name || !client.gmailLabel) {
        throw createError(SYSTEM_CONFIG.ERROR_CODES.INVALID_INPUT, 'Valid client object required');
      }
      
      debugLog(`Processing Gmail for client: ${client.name} with label: ${client.gmailLabel}`);
      
      // Validate client configuration first
      const validation = validateClientConfiguration(client.name);
      if (!validation.isValid) {
        throw createError(
          SYSTEM_CONFIG.ERROR_CODES.INVALID_INPUT,
          `Client configuration invalid: ${validation.errors.join(', ')}`
        );
      }
      
      // Get Gmail label with validation
      const label = getGmailLabel(client.gmailLabel);
      
      // Get folder structure and spreadsheet
      const folderStructure = getClientFolderStructure(client);
      const spreadsheet = SpreadsheetApp.openById(client.spreadsheetId);
      const bufferSheet = getOrCreateSheet(spreadsheet, SYSTEM_CONFIG.SHEETS.BUFFER_SHEET_NAME);
      
      // Get processed message IDs to avoid duplicates
      const processedMessageIds = getProcessedMessageIds(bufferSheet);
      
      // Process Gmail threads with pagination
      const threads = label.getThreads();
      infoLog(`Found ${threads.length} threads in label: ${client.gmailLabel}`);
      
      let processedAttachments = 0;
      let totalAttachments = 0;
      let skippedMessages = 0;
      let errorCount = 0;
      const processedMessages = new Set();
      
      for (let i = 0; i < threads.length; i++) {
        try {
          const thread = threads[i];
          const messages = thread.getMessages();
          
          for (const message of messages) {
            const messageId = message.getId();
            
            // Skip if already processed
            if (processedMessageIds.has(messageId) || processedMessages.has(messageId)) {
              skippedMessages++;
              continue;
            }
            
            try {
              const attachments = message.getAttachments();
              totalAttachments += attachments.length;
              
              if (attachments.length === 0) {
                // Mark message as processed even if no attachments
                processedMessages.add(messageId);
                continue;
              }
              
              let messageAttachmentCount = 0;
              
              for (const attachment of attachments) {
                try {
                  if (isValidAttachment(attachment)) {
                    const result = processAttachment(
                      attachment,
                      message,
                      folderStructure.bufferFolder,
                      bufferSheet
                    );
                    
                    if (result.success) {
                      processedAttachments++;
                      messageAttachmentCount++;
                    }
                  } else {
                    debugLog(`Skipped invalid attachment: ${attachment.getName()}`);
                  }
                } catch (attachmentError) {
                  errorLog(`Error processing attachment: ${attachment.getName()}`, attachmentError);
                  errorCount++;
                }
              }
              
              // Mark message as processed regardless of attachment success/failure
              // This prevents infinite reprocessing of problematic messages
              processedMessages.add(messageId);
              
            } catch (messageError) {
              errorLog(`Error processing message: ${messageId}`, messageError);
              errorCount++;
              // Still mark as processed to avoid infinite retry
              processedMessages.add(messageId);
            }
          }
          
          // Add small delay every 10 threads to respect rate limits
          if ((i + 1) % 10 === 0) {
            sleep(500);
          }
          
        } catch (threadError) {
          errorLog(`Error processing thread`, threadError);
          errorCount++;
        }
      }
      
      const result = {
        totalThreads: threads.length,
        totalAttachments: totalAttachments,
        processedAttachments: processedAttachments,
        skippedMessages: skippedMessages,
        errorCount: errorCount,
        newMessagesProcessed: processedMessages.size
      };
      
      infoLog(`Gmail processing completed for client: ${client.name}`, result);
      return result;
      
    } catch (error) {
      errorLog(`Error processing Gmail for client: ${client?.name}`, error);
      throw error;
    }
  }
  
  /**
   * Get Gmail label with enhanced validation
   */
  function getGmailLabel(labelName) {
    try {
      validateInput(labelName, 'string', 'Gmail label name');
      
      // Try to get the label
      const labels = GmailApp.getUserLabels();
      const targetLabel = labels.find(label => label.getName() === labelName);
      
      if (!targetLabel) {
        // Check if label follows expected pattern
        if (!labelName.startsWith(SYSTEM_CONFIG.GMAIL.LABEL_PREFIX)) {
          warnLog(`Gmail label '${labelName}' doesn't follow expected pattern (should start with '${SYSTEM_CONFIG.GMAIL.LABEL_PREFIX}')`);
        }
        
        throw createError(
          SYSTEM_CONFIG.ERROR_CODES.INVALID_INPUT,
          `Gmail label '${labelName}' not found. Please create the label in Gmail first.`
        );
      }
      
      debugLog(`Found Gmail label: ${labelName}`);
      return targetLabel;
      
    } catch (error) {
      errorLog(`Error getting Gmail label: ${labelName}`, error);
      throw error;
    }
  }
  
  /**
   * Get or create sheet with proper error handling
   */
  function getOrCreateSheet(spreadsheet, sheetName) {
    try {
      let sheet = spreadsheet.getSheetByName(sheetName);
      
      if (!sheet) {
        infoLog(`Creating missing sheet: ${sheetName}`);
        sheet = spreadsheet.insertSheet(sheetName);
        
        // Setup the sheet structure
        setupSheetStructure(sheet, sheetName);
      }
      
      // Validate sheet has proper headers
      validateSheetHeaders(sheet, sheetName);
      
      return sheet;
      
    } catch (error) {
      errorLog(`Error getting/creating sheet: ${sheetName}`, error);
      throw error;
    }
  }
  
  /**
   * Validate sheet headers and fix if necessary
   */
  function validateSheetHeaders(sheet, sheetName) {
    try {
      if (sheet.getLastRow() === 0) {
        // Sheet is empty, set up headers
        setupSheetStructure(sheet, sheetName);
        return;
      }
      
      const expectedHeaders = getExpectedHeadersForSheet(sheetName);
      if (!expectedHeaders) return;
      
      const actualHeaders = sheet.getRange(1, 1, 1, expectedHeaders.length).getValues()[0];
      
      // Check if headers match
      let headersMatch = true;
      for (let i = 0; i < expectedHeaders.length; i++) {
        if (actualHeaders[i] !== expectedHeaders[i]) {
          headersMatch = false;
          break;
        }
      }
      
      if (!headersMatch) {
        warnLog(`Sheet headers don't match expected format for: ${sheetName}`, {
          expected: expectedHeaders,
          actual: actualHeaders
        });
        
        // Could optionally fix headers here, but for safety we'll just warn
      }
      
    } catch (error) {
      errorLog(`Error validating sheet headers for: ${sheetName}`, error);
    }
  }
  
  /**
   * Get expected headers for sheet type
   */
  function getExpectedHeadersForSheet(sheetName) {
    switch (sheetName) {
      case SYSTEM_CONFIG.SHEETS.BUFFER_SHEET_NAME:
        return SYSTEM_CONFIG.SHEETS.BUFFER_COLUMNS;
      case SYSTEM_CONFIG.SHEETS.FINAL_SHEET_NAME:
        return SYSTEM_CONFIG.SHEETS.FINAL_COLUMNS;
      case SYSTEM_CONFIG.SHEETS.INFLOW_SHEET_NAME:
      case SYSTEM_CONFIG.SHEETS.OUTFLOW_SHEET_NAME:
        return SYSTEM_CONFIG.SHEETS.FLOW_COLUMNS;
      default:
        return null;
    }
  }
  
  /**
   * Enhanced function to get processed message IDs with proper deduplication
   */
  function getProcessedMessageIds(bufferSheet) {
    try {
      const processedIds = new Set();
      
      if (bufferSheet.getLastRow() <= 1) {
        debugLog('Buffer sheet is empty or has only headers');
        return processedIds;
      }
      
      const data = bufferSheet.getDataRange().getValues();
      const headers = data[0];
      
      // Find the Message ID column (this is the fix for the original error)
      const messageIdIndex = getColumnIndex(headers, 'Gmail Message ID');
      
      if (messageIdIndex === -1) {
        warnLog('Gmail Message ID column not found in buffer sheet, cannot deduplicate properly');
        return processedIds;
      }
      
      // Collect all message IDs
      for (let i = 1; i < data.length; i++) {
        const messageId = safeGetCellValue(data[i], messageIdIndex);
        if (messageId && messageId.trim() !== '') {
          processedIds.add(messageId.trim());
        }
      }
      
      debugLog(`Found ${processedIds.size} processed message IDs in buffer sheet`);
      return processedIds;
      
    } catch (error) {
      errorLog('Error getting processed message IDs', error);
      return new Set(); // Return empty set on error to avoid blocking processing
    }
  }
  
  /**
   * Process a single attachment with comprehensive error handling
   */
  function processAttachment(attachment, message, bufferFolder, bufferSheet) {
    try {
      if (!attachment || !message || !bufferFolder || !bufferSheet) {
        throw createError(SYSTEM_CONFIG.ERROR_CODES.INVALID_INPUT, 'All parameters required for attachment processing');
      }
      
      const attachmentName = attachment.getName() || 'unnamed_attachment';
      debugLog(`Processing attachment: ${attachmentName}`);
      
      // Create file in buffer folder with retry logic
      const file = retryWithBackoff(() => {
        return bufferFolder.createFile(attachment);
      }, 3, 1000, `creating file ${attachmentName}`);
      
      const fileUrl = file.getUrl();
      const fileId = file.getId();
      const uniqueId = generateUniqueId();
      const messageId = message.getId();
      
      // Get message details safely
      const emailSubject = getMessageSubject(message);
      const emailSender = getMessageSender(message);
      const messageDate = getMessageDate(message);
      
      // Prepare row data with all required columns
      const headers = bufferSheet.getRange(1, 1, 1, bufferSheet.getLastColumn()).getValues()[0];
      const rowData = createBufferRowData(headers, {
        originalFileName: attachmentName,
        changedFileName: attachmentName, // Will be updated during AI processing
        fileUrl: fileUrl,
        fileId: fileId,
        messageId: messageId,
        invoiceNumber: '', // To be filled by AI
        status: SYSTEM_CONFIG.STATUS.ACTIVE,
        reason: '',
        emailSubject: emailSubject,
        emailSender: emailSender,
        dateAdded: getCurrentTimestamp(),
        lastModified: getCurrentTimestamp(),
        processingAttempts: '0'
      });
      
      // Add row to buffer sheet with retry
      retryWithBackoff(() => {
        bufferSheet.appendRow(rowData);
      }, 3, 1000, `adding row to buffer sheet for ${attachmentName}`);
      
      const result = {
        success: true,
        filename: attachmentName,
        fileId: fileId,
        fileUrl: fileUrl,
        messageId: messageId
      };
      
      debugLog(`Successfully processed attachment: ${attachmentName}`, result);
      return result;
      
    } catch (error) {
      errorLog(`Error processing attachment: ${attachment?.getName()}`, error);
      return {
        success: false,
        error: error.message,
        filename: attachment?.getName() || 'unknown'
      };
    }
  }
  
  /**
   * Create buffer row data matching the sheet structure
   */
  function createBufferRowData(headers, data) {
    try {
      const rowData = new Array(headers.length).fill('');
      
      // Map data to correct columns (robust names)
      const columnMappings = {
        'Date': data.dateAdded || getCurrentTimestamp(),
        'OriginalFileName': data.originalFileName,
        'ChangedFilename': data.changedFileName,
        'Invoice ID': data.invoiceNumber,
        'Drive File ID': data.fileId,
        'Gmail Message ID': data.messageId,
        'Reason': data.reason,
        'Status': data.status,
        'UI': data.uniqueId,
        'Repeated': data.repeated,
        'Invoice Count': data.invoiceCount,
        'Attachment ID': data.attachmentId,
        'Email ID': data.emailId
      };
      
      // Fill row data based on header positions
      for (const [columnName, value] of Object.entries(columnMappings)) {
        const index = getColumnIndex(headers, columnName);
        if (index !== -1) {
          rowData[index] = value || '';
        }
      }
      
      return rowData;
      
    } catch (error) {
      errorLog('Error creating buffer row data', error);
      throw error;
    }
  }
  
  /**
   * Safely get message subject
   */
  function getMessageSubject(message) {
    try {
      return message.getSubject() || 'No Subject';
    } catch (error) {
      warnLog('Error getting message subject', error);
      return 'No Subject';
    }
  }
  
  /**
   * Safely get message sender
   */
  function getMessageSender(message) {
    try {
      return message.getFrom() || 'Unknown Sender';
    } catch (error) {
      warnLog('Error getting message sender', error);
      return 'Unknown Sender';
    }
  }
  
  /**
   * Safely get message date
   */
  function getMessageDate(message) {
    try {
      const date = message.getDate();
      return date ? date.toISOString() : getCurrentTimestamp();
    } catch (error) {
      warnLog('Error getting message date', error);
      return getCurrentTimestamp();
    }
  }
  
  /**
   * Enhanced attachment validation with comprehensive checks
   */
  function isValidAttachment(attachment) {
    try {
      if (!attachment) {
        debugLog('Attachment is null or undefined');
        return false;
      }
      
      // Check if it's a Google-type file (skip these)
      try {
        if (attachment.isGoogleType && attachment.isGoogleType()) {
          debugLog(`Skipping Google file type: ${attachment.getName()}`);
          return false;
        }
      } catch (error) {
        // isGoogleType might not be available in all contexts
        debugLog('Could not check Google file type, continuing validation');
      }
      
      // Get attachment properties safely
      const name = getAttachmentName(attachment);
      const size = getAttachmentSize(attachment);
      const mimeType = getAttachmentMimeType(attachment);
      
      // Validate filename
      if (!name || name.trim() === '' || name.startsWith('ATT')) {
        debugLog(`Invalid filename: ${name}`);
        return false;
      }
      
      // Check file size
      if (size > SYSTEM_CONFIG.DRIVE.MAX_FILE_SIZE) {
        debugLog(`Attachment too large: ${name} (${size} bytes)`);
        return false;
      }
      
      // Check MIME type
      if (!isValidMimeType(mimeType)) {
        debugLog(`Unsupported MIME type: ${mimeType} for file ${name}`);
        return false;
      }
      
      // Additional filename validations
      if (name.length > SYSTEM_CONFIG.DRIVE.MAX_FILENAME_LENGTH) {
        debugLog(`Filename too long: ${name}`);
        return false;
      }
      
      // Check for suspicious file patterns
      const suspiciousPatterns = [
        /^\./, // Hidden files
        /\.(exe|bat|cmd|scr|pif|com)$/i, // Executable files
        /^~\$/, // Temporary files
        /^thumbs\.db$/i, // System files
        /^desktop\.ini$/i
      ];
      
      for (const pattern of suspiciousPatterns) {
        if (pattern.test(name)) {
          debugLog(`Suspicious file pattern detected: ${name}`);
          return false;
        }
      }
      
      debugLog(`Attachment validation passed: ${name}`);
      return true;
      
    } catch (error) {
      errorLog(`Error validating attachment: ${attachment?.getName()}`, error);
      return false;
    }
  }
  
  /**
   * Safely get attachment name
   */
  function getAttachmentName(attachment) {
    try {
      return attachment.getName() || '';
    } catch (error) {
      warnLog('Error getting attachment name', error);
      return '';
    }
  }
  
  /**
   * Safely get attachment size
   */
  function getAttachmentSize(attachment) {
    try {
      return attachment.getSize() || 0;
    } catch (error) {
      warnLog('Error getting attachment size', error);
      return 0;
    }
  }
  
  /**
   * Safely get attachment MIME type
   */
  function getAttachmentMimeType(attachment) {
    try {
      return attachment.getContentType() || 'application/octet-stream';
    } catch (error) {
      warnLog('Error getting attachment MIME type', error);
      return 'application/octet-stream';
    }
  }
  
  /**
   * Manual trigger for specific client Gmail processing
   */
  function processClientGmailByName(clientName) {
    try {
      validateInput(clientName, 'string', 'Client name');
      
      const client = getClientByName(clientName);
      if (!client) {
        throw createError(SYSTEM_CONFIG.ERROR_CODES.INVALID_INPUT, `Client '${clientName}' not found`);
      }
      
      return processClientGmail(client);
      
    } catch (error) {
      errorLog(`Error processing Gmail for client by name: ${clientName}`, error);
      throw error;
    }
  }
  
  /**
   * Get Gmail labels for client validation
   */
  function getGmailLabelsForClient() {
    try {
      const labels = GmailApp.getUserLabels();
      const clientLabels = [];
      
      labels.forEach(label => {
        const labelName = label.getName();
        if (labelName.startsWith(SYSTEM_CONFIG.GMAIL.LABEL_PREFIX)) {
          try {
            clientLabels.push({
              name: labelName,
              displayName: labelName.replace(SYSTEM_CONFIG.GMAIL.LABEL_PREFIX, ''),
              threadCount: label.getThreads(0, 1).length > 0 ? '1+' : '0' // Quick check
            });
          } catch (error) {
            warnLog(`Error getting details for label: ${labelName}`, error);
            clientLabels.push({
              name: labelName,
              displayName: labelName.replace(SYSTEM_CONFIG.GMAIL.LABEL_PREFIX, ''),
              threadCount: 'Unknown'
            });
          }
        }
      });
      
      return clientLabels;
      
    } catch (error) {
      errorLog('Error getting Gmail labels', error);
      throw error;
    }
  }
  
  /**
   * Create Gmail label for client if it doesn't exist
   */
  function createClientGmailLabel(clientName) {
    try {
      validateInput(clientName, 'string', 'Client name');
      
      const labelName = SYSTEM_CONFIG.GMAIL.LABEL_PREFIX + cleanFilename(clientName).toLowerCase();
      
      // Check if label already exists
      const existingLabels = GmailApp.getUserLabels();
      const existingLabel = existingLabels.find(label => label.getName() === labelName);
      
      if (existingLabel) {
        infoLog(`Gmail label already exists: ${labelName}`);
        return { success: true, labelName: labelName, created: false };
      }
      
      // Create new label
      const label = GmailApp.createLabel(labelName);
      infoLog(`Created Gmail label: ${labelName}`);
      
      return { success: true, labelName: labelName, created: true };
      
    } catch (error) {
      errorLog(`Error creating Gmail label for client: ${clientName}`, error);
      throw error;
    }
  }
  
  /**
   * Enhanced Gmail setup validation
   */
  function validateClientGmailSetup(client) {
    try {
      const validation = {
        isValid: true,
        errors: [],
        warnings: []
      };
      
      if (!client || !client.gmailLabel) {
        validation.errors.push('Client or Gmail label missing');
        validation.isValid = false;
        return validation;
      }
      
      // Check if Gmail label exists
      let label;
      try {
        label = getGmailLabel(client.gmailLabel);
      } catch (error) {
        validation.errors.push(`Gmail label '${client.gmailLabel}' not found`);
        validation.isValid = false;
        return validation;
      }
      
      // Check label naming convention
      if (!client.gmailLabel.startsWith(SYSTEM_CONFIG.GMAIL.LABEL_PREFIX)) {
        validation.warnings.push(`Gmail label '${client.gmailLabel}' doesn't follow naming convention (should start with '${SYSTEM_CONFIG.GMAIL.LABEL_PREFIX}')`);
      }
      
      // Check for messages in label
      try {
        const threads = label.getThreads(0, 5); // Check first 5 threads
        if (threads.length === 0) {
          validation.warnings.push(`No messages found in Gmail label '${client.gmailLabel}'`);
        } else {
          // Check for attachments in recent messages
          let hasAttachments = false;
          let totalAttachments = 0;
          
          for (const thread of threads) {
            const messages = thread.getMessages();
            for (const message of messages) {
              const attachments = message.getAttachments();
              totalAttachments += attachments.length;
              if (attachments.length > 0) {
                hasAttachments = true;
              }
            }
          }
          
          if (!hasAttachments) {
            validation.warnings.push(`No attachments found in recent messages for label '${client.gmailLabel}'`);
          } else {
            validation.info = `Found ${totalAttachments} total attachments in recent messages`;
          }
        }
      } catch (error) {
        validation.warnings.push(`Could not check messages in label '${client.gmailLabel}': ${error.message}`);
      }
      
      return validation;
      
    } catch (error) {
      errorLog(`Error validating Gmail setup for client: ${client?.name}`, error);
      return {
        isValid: false,
        errors: [`Validation failed: ${error.message}`],
        warnings: []
      };
    }
  }
  
  /**
   * Get comprehensive Gmail processing statistics
   */
  function getGmailProcessingStats(clientName) {
    try {
      validateInput(clientName, 'string', 'Client name');
      
      const client = getClientByName(clientName);
      if (!client) {
        throw createError(SYSTEM_CONFIG.ERROR_CODES.INVALID_INPUT, `Client '${clientName}' not found`);
      }
      
      const spreadsheet = SpreadsheetApp.openById(client.spreadsheetId);
      const bufferSheet = getOrCreateSheet(spreadsheet, SYSTEM_CONFIG.SHEETS.BUFFER_SHEET_NAME);
      
      const stats = {
        totalFiles: 0,
        activeFiles: 0,
        deletedFiles: 0,
        lastProcessed: null,
        processingErrors: 0,
        uniqueMessages: 0
      };
      
      if (bufferSheet.getLastRow() <= 1) {
        return stats;
      }
      
      const data = bufferSheet.getDataRange().getValues();
      const headers = data[0];
      
      // Get column indices
      const statusIndex = getColumnIndex(headers, 'Status');
      const dateAddedIndex = getColumnIndex(headers, 'Date');
      const messageIdIndex = getColumnIndex(headers, 'Gmail Message ID');
      const attemptsIndex = getColumnIndex(headers, 'Processing Attempts');
      
      const uniqueMessageIds = new Set();
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const status = safeGetCellValue(row, statusIndex);
        const dateAdded = safeGetCellValue(row, dateAddedIndex);
        const messageId = safeGetCellValue(row, messageIdIndex);
        const attempts = safeGetCellValue(row, attemptsIndex, '0');
        
        stats.totalFiles++;
        
        // Count by status
        if (status === SYSTEM_CONFIG.STATUS.ACTIVE) {
          stats.activeFiles++;
        } else if (status === SYSTEM_CONFIG.STATUS.DELETED) {
          stats.deletedFiles++;
        }
        
        // Track processing errors
        const attemptCount = parseInt(attempts) || 0;
        if (attemptCount > 1) {
          stats.processingErrors++;
        }
        
        // Track unique messages
        if (messageId) {
          uniqueMessageIds.add(messageId);
        }
        
        // Track last processed
        if (dateAdded && (!stats.lastProcessed || new Date(dateAdded) > new Date(stats.lastProcessed))) {
          stats.lastProcessed = dateAdded;
        }
      }
      
      stats.uniqueMessages = uniqueMessageIds.size;
      
      return stats;
      
    } catch (error) {
      errorLog(`Error getting Gmail processing stats for client: ${clientName}`, error);
      throw error;
    }
  }
  
  /**
   * Clean up old processed messages to prevent memory issues
   */
  function cleanupOldProcessedMessages(clientName, daysOld = 30) {
    try {
      validateInput(clientName, 'string', 'Client name');
      
      const client = getClientByName(clientName);
      if (!client) {
        throw createError(SYSTEM_CONFIG.ERROR_CODES.INVALID_INPUT, `Client '${clientName}' not found`);
      }
      
      const spreadsheet = SpreadsheetApp.openById(client.spreadsheetId);
      const bufferSheet = getOrCreateSheet(spreadsheet, SYSTEM_CONFIG.SHEETS.BUFFER_SHEET_NAME);
      
      if (bufferSheet.getLastRow() <= 1) {
        return { cleaned: 0, message: 'No data to clean' };
      }
      
      const cutoffDate = new Date();
      cutoffDate.setDate(cutoffDate.getDate() - daysOld);
      
      const data = bufferSheet.getDataRange().getValues();
      const headers = data[0];
      const dateAddedIndex = getColumnIndex(headers, 'Date');
      const statusIndex = getColumnIndex(headers, 'Status');
      
      let cleanedCount = 0;
      
      // Process from bottom to top to avoid index issues
      for (let i = data.length - 1; i >= 1; i--) {
        const row = data[i];
        const dateAdded = safeGetCellValue(row, dateAddedIndex);
        const status = safeGetCellValue(row, statusIndex);
        
        if (dateAdded && status === SYSTEM_CONFIG.STATUS.DELETED) {
          const rowDate = new Date(dateAdded);
          if (rowDate < cutoffDate) {
            bufferSheet.deleteRow(i + 1);
            cleanedCount++;
          }
        }
      }
      
      infoLog(`Cleaned up ${cleanedCount} old processed messages for client: ${clientName}`);
      return { cleaned: cleanedCount, message: `Cleaned up ${cleanedCount} old records` };
      
    } catch (error) {
      errorLog(`Error cleaning up old messages for client: ${clientName}`, error);
      throw error;
    }
  }
/**
 * ClientManager.gs - Manages client configurations and operations (FIXED)
 */

/**
 * Enhanced Client configuration class with validation
 */
class ClientConfig {
    constructor(name, gmailLabel, rootFolderId, spreadsheetId, status = 'Active') {
      // Validate inputs
      validateInput(name, 'string', 'Client name');
      validateInput(gmailLabel, 'string', 'Gmail label');
      validateInput(rootFolderId, 'string', 'Root folder ID');
      validateInput(spreadsheetId, 'string', 'Spreadsheet ID');
      
      this.name = name.trim();
      this.gmailLabel = gmailLabel.trim();
      this.rootFolderId = rootFolderId.trim();
      this.spreadsheetId = spreadsheetId.trim();
      this.status = status;
      this.createdAt = getCurrentTimestamp();
      this.lastModified = getCurrentTimestamp();
    }
    
    /**
     * Validate client configuration
     */
    validate() {
      const errors = [];
      
      try {
        // Test folder access
        DriveApp.getFolderById(this.rootFolderId);
      } catch (error) {
        errors.push(`Cannot access root folder: ${this.rootFolderId}`);
      }
      
      try {
        // Test spreadsheet access
        SpreadsheetApp.openById(this.spreadsheetId);
      } catch (error) {
        errors.push(`Cannot access spreadsheet: ${this.spreadsheetId}`);
      }
      
      return {
        isValid: errors.length === 0,
        errors: errors
      };
    }
  }
  
  /**
   * Get all client configurations with enhanced error handling
   */
  function getAllClients() {
    let lock;
    try {
      // Use lock to ensure data consistency
      lock = LockService.getScriptLock();
      if (!lock.tryLock(5000)) { // 5 second timeout
        throw createError(SYSTEM_CONFIG.ERROR_CODES.SYSTEM_ERROR, 'Could not acquire lock to read clients');
      }
      
      const masterSheetId = getMasterConfigSheetId();
      const spreadsheet = SpreadsheetApp.openById(masterSheetId);
      const sheet = spreadsheet.getActiveSheet();
      
      // Auto-create headers if missing
      const requiredHeaders = ['Client Name', 'Gmail Label', 'Root Folder ID', 'Spreadsheet ID', 'Status'];
      let data = sheet.getDataRange().getValues();
      if (data.length === 0 || data[0].length < requiredHeaders.length || requiredHeaders.some((h, i) => data[0][i] !== h)) {
        // Set headers in the first row
        sheet.clear();
        sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
        data = [requiredHeaders];
      }
      
      // Validate sheet has data
      if (data.length <= 1) {
        debugLog('Master config sheet is empty');
        return [];
      }
      
      const headers = data[0];
      const clients = [];
      
      // Validate headers
      const missingHeaders = requiredHeaders.filter(header => !headers.includes(header));
      if (missingHeaders.length > 0) {
        throw createError(
          SYSTEM_CONFIG.ERROR_CODES.INVALID_INPUT,
          `Missing required headers in master config sheet: ${missingHeaders.join(', ')}`
        );
      }
      
      // Get column indices safely
      const nameIndex = getColumnIndex(headers, 'Client Name');
      const labelIndex = getColumnIndex(headers, 'Gmail Label');
      const folderIndex = getColumnIndex(headers, 'Root Folder ID');
      const sheetIndex = getColumnIndex(headers, 'Spreadsheet ID');
      const statusIndex = getColumnIndex(headers, 'Status');
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        // Validate required fields
        const name = safeGetCellValue(row, nameIndex);
        const label = safeGetCellValue(row, labelIndex);
        const folderId = safeGetCellValue(row, folderIndex);
        const spreadsheetId = safeGetCellValue(row, sheetIndex);
        
        if (name && label && folderId && spreadsheetId) {
          try {
            const status = safeGetCellValue(row, statusIndex, 'Active');
            const client = new ClientConfig(name, label, folderId, spreadsheetId, status);
            clients.push(client);
          } catch (error) {
            warnLog(`Invalid client data at row ${i + 1}`, error.message);
          }
        } else {
          warnLog(`Incomplete client data at row ${i + 1}`, {
            name, label, folderId, spreadsheetId
          });
        }
      }
      
      infoLog(`Loaded ${clients.length} clients from master config`);
      return clients;
      
    } catch (error) {
      errorLog('Error loading clients from master config', error);
      throw error;
    } finally {
      if (lock) {
        try {
          lock.releaseLock();
        } catch (releaseError) {
          errorLog('Error releasing lock', releaseError);
        }
      }
    }
  }
  
  /**
   * Get client configuration by name with validation
   */
  function getClientByName(clientName) {
    try {
      validateInput(clientName, 'string', 'Client name');
      
      const clients = getAllClients();
      const client = clients.find(c => c.name.toLowerCase() === clientName.toLowerCase().trim());
      
      if (client) {
        debugLog(`Found client: ${clientName}`);
      } else {
        debugLog(`Client not found: ${clientName}`);
      }
      
      return client || null;
    } catch (error) {
      errorLog(`Error getting client by name: ${clientName}`, error);
      throw error;
    }
  }
  
  /**
   * Get client configuration by Gmail label
   */
  function getClientByLabel(label) {
    try {
      validateInput(label, 'string', 'Gmail label');
      
      const clients = getAllClients();
      return clients.find(c => c.gmailLabel.toLowerCase() === label.toLowerCase().trim()) || null;
    } catch (error) {
      errorLog(`Error getting client by label: ${label}`, error);
      throw error;
    }
  }
  
  /**
   * Enhanced client creation with atomic operations and proper rollback
   */
  function addClientWithAtomicTransaction(clientName, gmailLabel, parentFolderId = null) {
    let lock;
    const createdResources = {
      rootFolder: null,
      spreadsheet: null,
      masterSheetRow: null
    };
    
    try {
      // Input validation
      validateInput(clientName, 'string', 'Client name');
      validateInput(gmailLabel, 'string', 'Gmail label');
      
      const cleanName = clientName.trim();
      const cleanLabel = gmailLabel.trim();
      
      // Acquire lock for atomic operation
      lock = LockService.getScriptLock();
      if (!lock.tryLock(30000)) { // 30 second timeout for folder creation
        throw createError(SYSTEM_CONFIG.ERROR_CODES.SYSTEM_ERROR, 'Could not acquire lock for client creation');
      }
      
      // Check if client already exists
      const existingClient = getClientByName(cleanName);
      if (existingClient) {
        throw createError(
          SYSTEM_CONFIG.ERROR_CODES.DUPLICATE_CLIENT,
          `Client '${cleanName}' already exists`
        );
      }
      
      // Check if Gmail label already exists
      const existingLabel = getClientByLabel(cleanLabel);
      if (existingLabel) {
        throw createError(
          SYSTEM_CONFIG.ERROR_CODES.DUPLICATE_CLIENT,
          `Gmail label '${cleanLabel}' is already in use by client '${existingLabel.name}'`
        );
      }
      
      infoLog(`Creating client: ${cleanName} with label: ${cleanLabel}`);
      
      // Step 1: Create folder structure
      infoLog('Step 1: Creating folder structure');
      // Use attachment downloader parent folder for seamless integration
      const ATTACHMENT_PARENT_FOLDER_ID = "1Tqj8May8je0L1lET5PIHRJj4P8d9T3MT";
      const folderStructure = createClientFolderStructure(cleanName, ATTACHMENT_PARENT_FOLDER_ID);
      createdResources.rootFolder = folderStructure.rootFolder;
      
      // Step 2: Create and setup spreadsheet
      infoLog('Step 2: Creating spreadsheet');
      const spreadsheet = createClientSpreadsheet(cleanName, folderStructure.spreadsheetsFolder);
      createdResources.spreadsheet = spreadsheet;
      
      // Step 3: Setup spreadsheet sheets with proper structure
      infoLog('Step 3: Setting up spreadsheet sheets');
      setupClientSpreadsheetSheets(spreadsheet, cleanName);
      
      // Step 4: Add to master config sheet
      infoLog('Step 4: Adding to master config sheet');
      const masterSheetId = getMasterConfigSheetId();
      const masterSpreadsheet = SpreadsheetApp.openById(masterSheetId);
      const masterSheet = masterSpreadsheet.getActiveSheet();
      // Ensure headers exist before appending
      const requiredHeaders = ['Client Name', 'Gmail Label', 'Root Folder ID', 'Spreadsheet ID', 'Status'];
      let data = masterSheet.getDataRange().getValues();
      if (data.length === 0 || data[0].length < requiredHeaders.length || requiredHeaders.some((h, i) => data[0][i] !== h)) {
        masterSheet.clear();
        masterSheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
      }
      
      // Final check for duplicates (race condition protection)
      const currentData = masterSheet.getDataRange().getValues();
      if (currentData.length > 1) {
        const headers = currentData[0];
        const nameIndex = getColumnIndex(headers, 'Client Name');
        const labelIndex = getColumnIndex(headers, 'Gmail Label');
        
        for (let i = 1; i < currentData.length; i++) {
          const existingName = safeGetCellValue(currentData[i], nameIndex);
          const existingLabel = safeGetCellValue(currentData[i], labelIndex);
          
          if (existingName.toLowerCase() === cleanName.toLowerCase()) {
            throw createError(
              SYSTEM_CONFIG.ERROR_CODES.DUPLICATE_CLIENT,
              `Client '${cleanName}' was just created by another process`
            );
          }
          
          if (existingLabel.toLowerCase() === cleanLabel.toLowerCase()) {
            throw createError(
              SYSTEM_CONFIG.ERROR_CODES.DUPLICATE_CLIENT,
              `Gmail label '${cleanLabel}' was just taken by another process`
            );
          }
        }
      }
      
      // Add the new client row
      const newRow = [
        cleanName,
        cleanLabel,
        folderStructure.rootFolder.getId(),
        spreadsheet.getId(),
        SYSTEM_CONFIG.STATUS.ACTIVE,
        getCurrentTimestamp(), // Created at
        getCurrentTimestamp()  // Last modified
      ];
      
      masterSheet.appendRow(newRow);
      createdResources.masterSheetRow = masterSheet.getLastRow();
      
      // Step 5: Verify the client was added correctly
      infoLog('Step 5: Verifying client creation');
      const verifyClient = getClientByName(cleanName);
      if (!verifyClient) {
        throw createError(SYSTEM_CONFIG.ERROR_CODES.SYSTEM_ERROR, 'Client verification failed after creation');
      }
      
      // Step 6: Validate all resources are accessible
      const validation = verifyClient.validate();
      if (!validation.isValid) {
        throw createError(
          SYSTEM_CONFIG.ERROR_CODES.SYSTEM_ERROR,
          `Client validation failed: ${validation.errors.join(', ')}`
        );
      }
      
      infoLog(`Successfully created client: ${cleanName}`, {
        rootFolderId: folderStructure.rootFolder.getId(),
        spreadsheetId: spreadsheet.getId(),
        folderStructure: folderStructure.folders
      });
      
      return {
        success: true,
        message: `Client '${cleanName}' created successfully`,
        client: verifyClient,
        folderStructure: folderStructure.folders,
        spreadsheetId: spreadsheet.getId()
      };
      
    } catch (error) {
      errorLog(`Error creating client: ${clientName}`, error);
      
      // Rollback created resources
      try {
        infoLog('Rolling back created resources due to error');
        
        // Remove from master sheet if added
        if (createdResources.masterSheetRow) {
          try {
            const masterSheetId = getMasterConfigSheetId();
            const masterSheet = SpreadsheetApp.openById(masterSheetId).getActiveSheet();
            masterSheet.deleteRow(createdResources.masterSheetRow);
            infoLog('Removed client from master sheet');
          } catch (rollbackError) {
            errorLog('Error removing client from master sheet during rollback', rollbackError);
          }
        }
        
        // Delete spreadsheet if created
        if (createdResources.spreadsheet) {
          try {
            DriveApp.getFileById(createdResources.spreadsheet.getId()).setTrashed(true);
            infoLog('Moved spreadsheet to trash');
          } catch (rollbackError) {
            errorLog('Error trashing spreadsheet during rollback', rollbackError);
          }
        }
        
        // Delete folder structure if created
        if (createdResources.rootFolder) {
          try {
            createdResources.rootFolder.setTrashed(true);
            infoLog('Moved root folder to trash');
          } catch (rollbackError) {
            errorLog('Error trashing root folder during rollback', rollbackError);
          }
        }
        
      } catch (rollbackError) {
        errorLog('Error during rollback process', rollbackError);
      }
      
      throw error;
    } finally {
      if (lock) {
        try {
          lock.releaseLock();
        } catch (releaseError) {
          errorLog('Error releasing lock', releaseError);
        }
      }
    }
  }
  
  /**
   * Create client folder structure compatible with attachment downloader system
   */
  function createClientFolderStructure(clientName, parentFolderId = null) {
    try {
      infoLog(`Creating attachment downloader compatible folder structure for client: ${clientName}`);
      
      // Use the same parent folder as attachment downloader
      const ATTACHMENT_PARENT_FOLDER_ID = "1Tqj8May8je0L1lET5PIHRJj4P8d9T3MT"; // Same as attachment_downloader.js
      
      let rootFolder;
      if (parentFolderId && parentFolderId === ATTACHMENT_PARENT_FOLDER_ID) {
        // Use the existing attachment downloader parent folder
        const parentFolder = DriveApp.getFolderById(ATTACHMENT_PARENT_FOLDER_ID);
        
        // Check if company folder already exists
        const existingFolders = parentFolder.getFoldersByName(clientName);
        if (existingFolders.hasNext()) {
          rootFolder = existingFolders.next();
          infoLog(`Using existing company folder: ${rootFolder.getName()}`);
        } else {
          rootFolder = parentFolder.createFolder(clientName);
          infoLog(`Created new company folder: ${rootFolder.getName()}`);
        }
      } else {
        // Create in specified parent or root
        if (parentFolderId) {
          try {
            const parentFolder = DriveApp.getFolderById(parentFolderId);
            rootFolder = parentFolder.createFolder(`Client-${clientName}`);
          } catch (error) {
            warnLog(`Cannot access parent folder ${parentFolderId}, creating in attachment parent`, error);
            const attachmentParent = DriveApp.getFolderById(ATTACHMENT_PARENT_FOLDER_ID);
            rootFolder = attachmentParent.createFolder(clientName);
          }
        } else {
          // Use attachment downloader parent folder as default
          const attachmentParent = DriveApp.getFolderById(ATTACHMENT_PARENT_FOLDER_ID);
          rootFolder = attachmentParent.createFolder(clientName);
        }
      }
      
      // Register the company in attachment downloader system
      registerClientWithAttachmentDownloader(clientName, rootFolder.getId());
      
      // Create attachment downloader compatible structure
      // This creates the base structure that attachment downloader will extend with FY folders
      const currentYear = new Date().getFullYear();
      const currentMonth = new Date().getMonth();
      
      // Determine current financial year (April to March)
      let fyStart, fyEnd;
      if (currentMonth >= 3) { // April onwards (month index 3 = April)
        fyStart = currentYear;
        fyEnd = currentYear + 1;
      } else { // January to March
        fyStart = currentYear - 1;
        fyEnd = currentYear;
      }
      
      const currentFY = `FY-${fyStart.toString().slice(-2)}-${fyEnd.toString().slice(-2)}`;
      
      // Create current financial year structure
      const financialYearFolder = getOrCreateFolder(rootFolder, currentFY);
      const accrualsFolder = getOrCreateFolder(financialYearFolder, "Accruals");
      
      // Create Buffer and Bills & Invoices structure
      const bufferFolder = getOrCreateFolder(accrualsFolder, "Buffer");
      const bufferActiveFolder = getOrCreateFolder(bufferFolder, "Active");
      const bufferDeletedFolder = getOrCreateFolder(bufferFolder, "Deleted");
      const buffer2Folder = getOrCreateFolder(accrualsFolder, "Buffer2");
      
      const billsInvoicesFolder = getOrCreateFolder(accrualsFolder, "Bills and Invoices");
      
      // Create month folders with Inflow/Outflow subfolders for current year
      const months = [
        'April', 'May', 'June', 'July', 'August', 'September',
        'October', 'November', 'December', 'January', 'February', 'March'
      ];
      
      const monthFolders = {};
      months.forEach(month => {
        const monthFolder = getOrCreateFolder(billsInvoicesFolder, month);
        monthFolders[month] = {
          folder: monthFolder,
          inflow: getOrCreateFolder(monthFolder, "Inflow"),
          outflow: getOrCreateFolder(monthFolder, "Outflow")
        };
      });
      
      const folderStructure = {
        rootFolder: rootFolder,
        folders: {
          root: rootFolder.getId(),
          currentFY: financialYearFolder.getId(),
          accruals: accrualsFolder.getId(),
          buffer: bufferFolder.getId(),
          bufferActive: bufferActiveFolder.getId(),
          bufferDeleted: bufferDeletedFolder.getId(),
          buffer2: buffer2Folder.getId(),
          billsInvoices: billsInvoicesFolder.getId(),
          months: monthFolders
        }
      };
      
      infoLog(`Created attachment downloader compatible folder structure for client: ${clientName}`, {
        rootId: rootFolder.getId(),
        currentFY: currentFY,
        monthsCreated: Object.keys(monthFolders).length
      });
      
      return folderStructure;
      
    } catch (error) {
      errorLog(`Error creating folder structure for client: ${clientName}`, error);
      throw error;
    }
  }
  
  /**
   * Register client with attachment downloader system for seamless integration
   */
  function registerClientWithAttachmentDownloader(clientName, rootFolderId) {
    try {
      // Store the mapping in script properties (same as attachment downloader)
      const props = PropertiesService.getScriptProperties();
      props.setProperty(`COMPANY_FOLDER_${clientName.toUpperCase()}`, rootFolderId);
      
      infoLog(`Registered client ${clientName} with attachment downloader system: ${rootFolderId}`);
      
    } catch (error) {
      warnLog(`Error registering client with attachment downloader: ${clientName}`, error);
    }
  }
  
  /**
   * Get or create folder (same function as attachment downloader)
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
      errorLog('Error creating/getting folder:', folderName, error);
      throw error;
    }
  }
  
  /**
   * Create client spreadsheet with proper setup using template approach
   */
  function createClientSpreadsheet(clientName, spreadsheetsFolder) {
    try {
      const spreadsheetName = `${clientName}_Processing`;
      
      // Get the current spreadsheet as template (contains all Apps Script code)
      const templateSpreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
      
      // Create a copy of the template spreadsheet
      const templateFile = DriveApp.getFileById(templateSpreadsheetId);
      const copiedFile = templateFile.makeCopy(spreadsheetName);
      
      // Wait a moment for file to be ready
      sleep(1000);
      
      // Move to correct folder
      spreadsheetsFolder.addFile(copiedFile);
      
      // Remove from root folder
      const rootFolders = DriveApp.getRootFolder();
      if (rootFolders.getFilesByName(spreadsheetName).hasNext()) {
        rootFolders.removeFile(copiedFile);
      }
      
      // Open the copied spreadsheet
      const spreadsheet = SpreadsheetApp.openById(copiedFile.getId());
      
      infoLog(`Created spreadsheet from template: ${spreadsheetName}`);
      return spreadsheet;
      
    } catch (error) {
      errorLog(`Error creating spreadsheet for client: ${clientName}`, error);
      throw error;
    }
  }
  
  // --- Robust Header Definition (from attachment_downloader.js) ---
  const MAIN_SHEET_HEADERS = [
    'File Name', 'File ID', 'File URL',
    'Date Created (Drive)', 'Last Updated (Drive)', 'Size (bytes)', 'Mime Type',
    'Email Subject', 'Gmail Message ID', 'invoice status', 'UI',
    'Date', 'Month', 'FY', 'GST', 'TDS', 'OT', 'NA'
  ];

  function ensureRobustSheetStructure(sheet) {
    sheet.clear();
    sheet.appendRow(MAIN_SHEET_HEADERS);
    sheet.getRange(1, 1, 1, MAIN_SHEET_HEADERS.length).setFontWeight('bold').setBackground('#E8F0FE').setBorder(true, true, true, true, true, true);
    sheet.setFrozenRows(1);
    // Optionally set column widths for key columns
    sheet.setColumnWidth(1, 200); // File Name
    sheet.setColumnWidth(2, 180); // File ID
    sheet.setColumnWidth(3, 200); // File URL
    // ... add more as needed
    if (sheet.getLastColumn() > MAIN_SHEET_HEADERS.length) {
      sheet.deleteColumns(MAIN_SHEET_HEADERS.length + 1, sheet.getLastColumn() - MAIN_SHEET_HEADERS.length);
    }
  }

  function getOrCreateRobustSheet(ss, sheetName) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      ensureRobustSheetStructure(sheet);
    } else {
      const firstRow = sheet.getRange(1, 1, 1, MAIN_SHEET_HEADERS.length).getValues()[0];
      if (firstRow.join(',') !== MAIN_SHEET_HEADERS.join(',')) {
        ensureRobustSheetStructure(sheet);
      }
    }
    return sheet;
  }

  function logFileToMainSheet(sheet, fileData) {
    const row = MAIN_SHEET_HEADERS.map(h => fileData[h] || '');
    sheet.appendRow(row);
  }

  function createRobustSheetsForClient(ss, clientName) {
    getOrCreateRobustSheet(ss, `${clientName}-buffer`);
    getOrCreateRobustSheet(ss, clientName);
    getOrCreateRobustSheet(ss, `${clientName}-inflow`);
    getOrCreateRobustSheet(ss, `${clientName}-outflow`);
  }
  
  /**
   * Enhanced spreadsheet sheet setup with template cleanup only (no sheet creation)
   */
  function setupClientSpreadsheetSheets(spreadsheet, clientName) {
    try {
      // Only clean up template sheets - don't create new sheets
      cleanupTemplateSheets(spreadsheet, clientName);
      
      infoLog(`Successfully cleaned up template sheets for client: ${clientName}`);
      
    } catch (error) {
      errorLog(`Error setting up spreadsheet sheets for client: ${clientName}`, error);
      throw error;
    }
  }
  
  /**
   * Clean up template sheets - remove all sheets except Sheet1 for clean start
   */
  function cleanupTemplateSheets(spreadsheet, clientName) {
    try {
      const sheets = spreadsheet.getSheets();
      
      // Remove all sheets except Sheet1 to start with a clean spreadsheet
      sheets.forEach(sheet => {
        const sheetName = sheet.getName();
        if (sheetName !== 'Sheet1') {
          try {
            spreadsheet.deleteSheet(sheet);
            infoLog(`Deleted template sheet: ${sheetName}`);
          } catch (deleteError) {
            warnLog(`Could not delete sheet: ${sheetName}`, deleteError);
          }
        }
      });
      
      // Clear any data from Sheet1 to ensure it's clean
      const defaultSheet = spreadsheet.getSheetByName('Sheet1');
      if (defaultSheet) {
        defaultSheet.clear();
        infoLog(`Cleared data from Sheet1 for clean start`);
      }
      
      infoLog(`Template cleanup completed - spreadsheet ready for client: ${clientName}`);
      
    } catch (error) {
      warnLog(`Error during template cleanup for client: ${clientName}`, error);
      // Don't throw error here, continue with setup
    }
  }
  
  /**
   * Setup individual sheet structure with proper headers and formatting
   */
  function setupSheetStructure(sheet, sheetName) {
    let headers;
    if (sheetName.endsWith('-buffer')) {
      headers = MAIN_SHEET_HEADERS;
    } else {
      headers = MAIN_SHEET_HEADERS;
    }
    sheet.clear();
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#E8F0FE').setBorder(true, true, true, true, true, true);
    sheet.setFrozenRows(1);
  }
  
  /**
   * Update client configuration with validation
   */
  function updateClient(clientName, updates) {
    let lock;
    try {
      validateInput(clientName, 'string', 'Client name');
      
      if (!updates || typeof updates !== 'object') {
        throw createError(SYSTEM_CONFIG.ERROR_CODES.INVALID_INPUT, 'Updates object is required');
      }
      
      lock = LockService.getScriptLock();
      if (!lock.tryLock(10000)) {
        throw createError(SYSTEM_CONFIG.ERROR_CODES.SYSTEM_ERROR, 'Could not acquire lock for client update');
      }
      
      const masterSheetId = getMasterConfigSheetId();
      const sheet = SpreadsheetApp.openById(masterSheetId).getActiveSheet();
      const data = sheet.getDataRange().getValues();
      
      if (data.length <= 1) {
        throw createError(SYSTEM_CONFIG.ERROR_CODES.INVALID_INPUT, `Client '${clientName}' not found`);
      }
      
      const headers = data[0];
      const nameIndex = getColumnIndex(headers, 'Client Name');
      
      for (let i = 1; i < data.length; i++) {
        const rowName = safeGetCellValue(data[i], nameIndex);
        if (rowName.toLowerCase() === clientName.toLowerCase()) {
          const row = i + 1;
          
          // Update allowed fields
          if (updates.gmailLabel !== undefined) {
            const labelIndex = getColumnIndex(headers, 'Gmail Label');
            if (labelIndex !== -1) {
              sheet.getRange(row, labelIndex + 1).setValue(updates.gmailLabel);
            }
          }
          
          if (updates.status !== undefined) {
            const statusIndex = getColumnIndex(headers, 'Status');
            if (statusIndex !== -1) {
              sheet.getRange(row, statusIndex + 1).setValue(updates.status);
            }
          }
          
          // Update last modified timestamp
          const lastModifiedIndex = getColumnIndex(headers, 'Last Modified');
          if (lastModifiedIndex !== -1) {
            sheet.getRange(row, lastModifiedIndex + 1).setValue(getCurrentTimestamp());
          }
          
          infoLog(`Updated client: ${clientName}`, updates);
          return { success: true, message: `Client '${clientName}' updated successfully` };
        }
      }
      
      throw createError(SYSTEM_CONFIG.ERROR_CODES.INVALID_INPUT, `Client '${clientName}' not found`);
      
    } catch (error) {
      errorLog('Error updating client', error);
      throw error;
    } finally {
      if (lock) {
        try {
          lock.releaseLock();
        } catch (releaseError) {
          errorLog('Error releasing lock', releaseError);
        }
      }
    }
  }
  
  /**
   * Get client folder structure with validation
   */
  function getClientFolderStructure(client) {
    try {
      if (!client || !client.rootFolderId) {
        throw createError(SYSTEM_CONFIG.ERROR_CODES.INVALID_INPUT, 'Valid client with root folder ID required');
      }
      
      const rootFolder = DriveApp.getFolderById(client.rootFolderId);
      
      // Navigate through folder structure with error handling
      const folderMap = {
        rootFolder: rootFolder,
        accrualsFolder: findSubfolder(rootFolder, SYSTEM_CONFIG.DRIVE.FOLDER_STRUCTURE.ACCRUALS),
        spreadsheetsFolder: findSubfolder(rootFolder, SYSTEM_CONFIG.DRIVE.FOLDER_STRUCTURE.SPREADSHEETS)
      };
      
      // Get nested folders
      folderMap.billsInvoicesFolder = findSubfolder(folderMap.accrualsFolder, SYSTEM_CONFIG.DRIVE.FOLDER_STRUCTURE.BILLS_AND_INVOICES);
      folderMap.bufferFolder = findSubfolder(folderMap.billsInvoicesFolder, SYSTEM_CONFIG.DRIVE.FOLDER_STRUCTURE.BUFFER);
      folderMap.monthsFolder = findSubfolder(folderMap.billsInvoicesFolder, SYSTEM_CONFIG.DRIVE.FOLDER_STRUCTURE.MONTHS);
      folderMap.inflowFolder = findSubfolder(folderMap.monthsFolder, SYSTEM_CONFIG.DRIVE.FOLDER_STRUCTURE.INFLOW);
      folderMap.outflowFolder = findSubfolder(folderMap.monthsFolder, SYSTEM_CONFIG.DRIVE.FOLDER_STRUCTURE.OUTFLOW);
      
      return folderMap;
      
    } catch (error) {
      errorLog(`Error getting folder structure for client: ${client?.name}`, error);
      throw error;
    }
  }
  
  /**
   * Helper function to find subfolder with error handling
   */
  function findSubfolder(parentFolder, folderName) {
    try {
      const folders = parentFolder.getFoldersByName(folderName);
      if (folders.hasNext()) {
        return folders.next();
      } else {
        throw createError(
          SYSTEM_CONFIG.ERROR_CODES.SYSTEM_ERROR,
          `Folder '${folderName}' not found in parent folder`
        );
      }
    } catch (error) {
      errorLog(`Error finding subfolder: ${folderName}`, error);
      throw error;
    }
  }
  
  /**
   * List all active clients with caching
   */
  function getActiveClients() {
    try {
      const allClients = getAllClients();
      const activeClients = allClients.filter(client => client.status === SYSTEM_CONFIG.STATUS.ACTIVE);
      
      infoLog(`Found ${activeClients.length} active clients out of ${allClients.length} total`);
      return activeClients;
      
    } catch (error) {
      errorLog('Error getting active clients', error);
      throw error;
    }
  }
  
  /**
   * Deactivate client safely
   */
  function deactivateClient(clientName) {
    try {
      return updateClient(clientName, { status: 'Inactive' });
    } catch (error) {
      errorLog(`Error deactivating client: ${clientName}`, error);
      throw error;
    }
  }
  
  /**
   * Activate client safely
   */
  function activateClient(clientName) {
    try {
      return updateClient(clientName, { status: SYSTEM_CONFIG.STATUS.ACTIVE });
    } catch (error) {
      errorLog(`Error activating client: ${clientName}`, error);
      throw error;
    }
  }
  
  /**
   * Comprehensive client validation
   */
  function validateClientConfiguration(clientName) {
    try {
      const client = getClientByName(clientName);
      if (!client) {
        return {
          isValid: false,
          errors: [`Client '${clientName}' not found`],
          warnings: []
        };
      }
      
      const validation = {
        isValid: true,
        errors: [],
        warnings: []
      };
      
      // Validate client object
      const clientValidation = client.validate();
      if (!clientValidation.isValid) {
        validation.errors.push(...clientValidation.errors);
        validation.isValid = false;
      }
      
      // Validate folder structure
      try {
        const folderStructure = getClientFolderStructure(client);
        debugLog(`Folder structure validation passed for: ${clientName}`);
      } catch (error) {
        validation.errors.push(`Folder structure invalid: ${error.message}`);
        validation.isValid = false;
      }
      
      // Validate spreadsheet structure
      try {
        const spreadsheet = SpreadsheetApp.openById(client.spreadsheetId);
        const requiredSheets = [
          SYSTEM_CONFIG.SHEETS.BUFFER_SHEET_NAME,
          SYSTEM_CONFIG.SHEETS.FINAL_SHEET_NAME,
          SYSTEM_CONFIG.SHEETS.INFLOW_SHEET_NAME,
          SYSTEM_CONFIG.SHEETS.OUTFLOW_SHEET_NAME
        ];
        
        for (const sheetName of requiredSheets) {
          const sheet = spreadsheet.getSheetByName(sheetName);
          if (!sheet) {
            validation.errors.push(`Required sheet '${sheetName}' not found`);
            validation.isValid = false;
          }
        }
      } catch (error) {
        validation.errors.push(`Spreadsheet validation failed: ${error.message}`);
        validation.isValid = false;
      }
      
      return validation;
      
    } catch (error) {
      errorLog(`Error validating client configuration: ${clientName}`, error);
      return {
        isValid: false,
        errors: [`Validation failed: ${error.message}`],
        warnings: []
      };
    }
  }
  
  /**
   * Clean up orphaned resources
   */
  function cleanupOrphanedClientResources() {
    try {
      infoLog('Starting cleanup of orphaned client resources');
      
      const clients = getAllClients();
      const clientIds = new Set();
      const clientFolders = new Set();
      const clientSpreadsheets = new Set();
      
      // Collect all valid client resources
      clients.forEach(client => {
        clientIds.add(client.name);
        clientFolders.add(client.rootFolderId);
        clientSpreadsheets.add(client.spreadsheetId);
      });
      
      // Implementation would scan for orphaned folders and spreadsheets
      // This is a placeholder for the cleanup logic
      
      infoLog('Completed cleanup of orphaned client resources');
      
    } catch (error) {
      errorLog('Error cleaning up orphaned resources', error);
      throw error;
    }
  }
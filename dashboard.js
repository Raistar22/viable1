//this dashboard code is able to create new sheet and copy the appscritp code from the parent sheet 

// Google Apps Script server-side code for FinTech Automation Dashboard
// Integrated with ClientManager functionality

// Configuration
const MASTER_SHEET_ID = '1Jv-kbz_zzV7GbumOun7cMtfVowVjzm4gND_rq-csI1c'; // Master sheet for client configurations
const PARENT_COMPANIES_FOLDER_ID = '1Tqj8May8je0L1lET5PIHRJj4P8d9T3MT'; // Parent folder for all client companies

// System Configuration (from ClientManager.js)
const SYSTEM_CONFIG = {
  ERROR_CODES: {
    SYSTEM_ERROR: 'SYSTEM_ERROR',
    INVALID_INPUT: 'INVALID_INPUT',
    DUPLICATE_CLIENT: 'DUPLICATE_CLIENT'
  },
  STATUS: {
    ACTIVE: 'Active',
    INACTIVE: 'Inactive'
  },
  DRIVE: {
    FOLDER_STRUCTURE: {
      ACCRUALS: 'Accruals',
      SPREADSHEETS: 'Spreadsheets',
      BILLS_AND_INVOICES: 'Bills and Invoices',
      BUFFER: 'Buffer',
      MONTHS: 'Months',
      INFLOW: 'Inflow',
      OUTFLOW: 'Outflow'
    }
  },
  SHEETS: {
    BUFFER_SHEET_NAME: 'buffer',
    FINAL_SHEET_NAME: 'final',
    INFLOW_SHEET_NAME: 'inflow',
    OUTFLOW_SHEET_NAME: 'outflow'
  }
};

// Standardized headers for all sheets
// const MAIN_SHEET_HEADERS = [
//   'File Name', 'File ID', 'File URL',
//   'Date Created (Drive)', 'Last Updated (Drive)', 'Size (bytes)', 'Mime Type',
//   'Email Subject', 'Gmail Message ID', 'invoice status', 'UI',
//   'Date', 'Month', 'FY', 'GST', 'TDS', 'OT', 'NA'
// ];

/**
 * Opens the dashboard HTML page
 */
function openDashboard() {
  const htmlOutput = HtmlService.createTemplateFromFile('Dashboard')
    .evaluate()
    .setWidth(1200)
    .setHeight(800)
    .setTitle('FinTech Automation Dashboard');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'FinTech Automation Dashboard');
}

/**
 * Add onOpen function to create menu in Google Sheets
 */
// function onOpen() {
//   const ui = SpreadsheetApp.getUi();
//   ui.createMenu('FinTech Automation')
//     .addItem('Open Dashboard', 'openDashboard')
//     .addToUi();
// }

/**
 * Include other HTML files (for templates)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ========== UTILITY FUNCTIONS ==========

/**
 * Get current timestamp
 */
function getCurrentTimestamp() {
  return new Date().toISOString();
}

/**
 * Validate input parameters
 */
function validateInput(value, type, fieldName) {
  if (value === null || value === undefined || (typeof value === 'string' && value.trim() === '')) {
    throw new Error(`${fieldName} is required`);
  }
  if (typeof value !== type) {
    throw new Error(`${fieldName} must be of type ${type}`);
  }
}

/**
 * Create error object
 */
function createError(code, message) {
  const error = new Error(message);
  error.code = code;
  return error;
}

/**
 * Get master config sheet ID
 */
function getMasterConfigSheetId() {
  return MASTER_SHEET_ID;
}

/**
 * Get column index from headers
 */
function getColumnIndex(headers, columnName) {
  return headers.findIndex(header => header === columnName);
}

/**
 * Safe get cell value
 */
function safeGetCellValue(row, index, defaultValue = '') {
  return (row && row[index] !== undefined && row[index] !== null) ? row[index].toString().trim() : defaultValue;
}

/**
 * Sleep function
 */
function sleep(ms) {
  Utilities.sleep(ms);
}

/**
 * Logging functions
 */
function debugLog(message, data = null) {
  console.log(`[DEBUG] ${message}`, data || '');
}

function infoLog(message, data = null) {
  console.log(`[INFO] ${message}`, data || '');
}

function warnLog(message, data = null) {
  console.log(`[WARN] ${message}`, data || '');
}

function errorLog(message, error = null) {
  console.error(`[ERROR] ${message}`, error || '');
}

// ========== CLIENT CONFIGURATION CLASS ==========

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

// ========== CLIENT MANAGEMENT FUNCTIONS ==========

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
          // Add ID field for dashboard tracking
          client.id = folderId; // Use folder ID as unique identifier
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
    return []; // Return empty array instead of throwing
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
    return null;
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
    return null;
  }
}

// ========== PROMPT-BASED CLIENT CREATION ==========

/**
 * Add new client using prompts instead of popup form
 */
function addNewClientWithPrompts() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Get client name
    const clientNameResponse = ui.prompt(
      'Add New Client - Step 1/2',
      'Enter the client name:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (clientNameResponse.getSelectedButton() !== ui.Button.OK) {
      return { success: false, message: 'Client creation cancelled' };
    }
    
    const clientName = clientNameResponse.getResponseText().trim();
    if (!clientName) {
      ui.alert('Error', 'Client name cannot be empty', ui.ButtonSet.OK);
      return { success: false, message: 'Client name cannot be empty' };
    }
    
    // Get Gmail label
    const gmailLabelResponse = ui.prompt(
      'Add New Client - Step 2/2',
      'Enter the Gmail label (e.g., client/accruals/bills&invoices):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (gmailLabelResponse.getSelectedButton() !== ui.Button.OK) {
      return { success: false, message: 'Client creation cancelled' };
    }
    
    const gmailLabel = gmailLabelResponse.getResponseText().trim();
    if (!gmailLabel) {
      ui.alert('Error', 'Gmail label cannot be empty', ui.ButtonSet.OK);
      return { success: false, message: 'Gmail label cannot be empty' };
    }
    
    // Show progress
    ui.alert('Creating Client', `Creating client "${clientName}" with Gmail label "${gmailLabel}". This may take a moment...`, ui.ButtonSet.OK);
    
    // Create client using integrated function
    const result = addClientWithAtomicTransaction(clientName, gmailLabel, PARENT_COMPANIES_FOLDER_ID);
    
    if (result.success) {
      ui.alert('Success', result.message, ui.ButtonSet.OK);
    } else {
      ui.alert('Error', result.message, ui.ButtonSet.OK);
    }
    
    return result;
    
  } catch (error) {
    errorLog('Error in addNewClient', error);
    SpreadsheetApp.getUi().alert('Error', `Error creating client: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    return { success: false, message: error.message };
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
    // Use the parent companies folder for all clients
    const folderStructure = createClientFolderStructure(cleanName, PARENT_COMPANIES_FOLDER_ID);
    createdResources.rootFolder = folderStructure.rootFolder;

    // Step 2: Create and setup spreadsheet
    infoLog('Step 2: Creating spreadsheet');
    const spreadsheet = createClientSpreadsheet(cleanName, folderStructure.rootFolder);
    createdResources.spreadsheet = spreadsheet;

    // Step 3: Setup spreadsheet sheets with proper structure
    infoLog('Step 3: Setting up spreadsheet sheets');
    setupClientSpreadsheetSheets(spreadsheet, cleanName);

    // Step 4: Set up folder mapping in the new client spreadsheet
    infoLog('Step 4: Setting up folder mapping in new client spreadsheet');
    setupClientSpreadsheetFolderMapping(spreadsheet, cleanName, folderStructure.rootFolder.getId());

    // Step 5: Add to master config sheet
    infoLog('Step 5: Adding to master config sheet');
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
    
    // Step 6: Verify the client was added correctly
    infoLog('Step 6: Verifying client creation');
    const verifyClient = getClientByName(cleanName);
    if (!verifyClient) {
      throw createError(SYSTEM_CONFIG.ERROR_CODES.SYSTEM_ERROR, 'Client verification failed after creation');
    }
    
    // Step 7: Validate all resources are accessible
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
    
    return {
      success: false,
      message: error.message || 'Error creating client'
    };
    
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

// ========== FOLDER STRUCTURE FUNCTIONS ==========

/**
 * Create client folder structure compatible with attachment downloader system
 */
function createClientFolderStructure(clientName, parentFolderId = null) {
  try {
    infoLog(`Creating attachment downloader compatible folder structure for client: ${clientName}`);
    
    // Use the parent companies folder for all clients
    const targetParentFolderId = parentFolderId || PARENT_COMPANIES_FOLDER_ID;
    
    let rootFolder;
    if (targetParentFolderId === PARENT_COMPANIES_FOLDER_ID) {
      // Use the parent companies folder
      const parentFolder = DriveApp.getFolderById(PARENT_COMPANIES_FOLDER_ID);
      
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
      // Create in specified parent or fall back to parent companies folder
      try {
        const parentFolder = DriveApp.getFolderById(targetParentFolderId);
        rootFolder = parentFolder.createFolder(clientName);
      } catch (error) {
        warnLog(`Cannot access parent folder ${targetParentFolderId}, using parent companies folder`, error);
        const parentCompaniesFolder = DriveApp.getFolderById(PARENT_COMPANIES_FOLDER_ID);
        rootFolder = parentCompaniesFolder.createFolder(clientName);
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
 * Helper function to create subfolder with error handling
 */
function createSubfolder(parentFolder, folderName) {
  try {
    // Check if folder already exists
    const existingFolders = parentFolder.getFoldersByName(folderName);
    if (existingFolders.hasNext()) {
      const existing = existingFolders.next();
      debugLog(`Folder '${folderName}' already exists, using existing folder`);
      return existing;
    }
    
    // Create new folder
    const newFolder = parentFolder.createFolder(folderName);
    debugLog(`Created folder: ${folderName}`);
    return newFolder;
    
  } catch (error) {
    errorLog(`Error creating subfolder: ${folderName}`, error);
    throw error;
  }
}

/**
 * Helper function to get or create a folder
 */
function getOrCreateFolder(parentFolder, folderName) {
  try {
    const existingFolders = parentFolder.getFoldersByName(folderName);
    if (existingFolders.hasNext()) {
      return existingFolders.next();
    }
    return parentFolder.createFolder(folderName);
  } catch (error) {
    errorLog(`Error getting or creating folder: ${folderName}`, error);
    throw error;
  }
}

/**
 * Register client with the attachment downloader system
 */
function registerClientWithAttachmentDownloader(clientName, rootFolderId) {
  try {
    const attachmentDownloaderScript = ScriptApp.getScriptId();
    const attachmentDownloaderFile = DriveApp.getFileById(attachmentDownloaderScript);
    
    const attachmentDownloaderFolder = DriveApp.getFolderById(rootFolderId);
    
    // Check if the script file is already in the folder
    const existingFiles = attachmentDownloaderFolder.getFilesByName(attachmentDownloaderFile.getName());
    if (!existingFiles.hasNext()) {
      attachmentDownloaderFolder.addFile(attachmentDownloaderFile);
      infoLog(`Registered client '${clientName}' with attachment downloader. Script added to folder.`);
    } else {
      infoLog(`Client '${clientName}' is already registered with attachment downloader. Script already in folder.`);
    }
  } catch (error) {
    errorLog(`Error registering client '${clientName}' with attachment downloader:`, error);
  }
}

// ========== SPREADSHEET FUNCTIONS ==========

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
function setupSheetStructure(sheet) {
  try {
    sheet.clear();
    sheet.appendRow(MAIN_SHEET_HEADERS);
    sheet.getRange(1, 1, 1, MAIN_SHEET_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#E8F0FE')
      .setBorder(true, true, true, true, true, true);
    sheet.setFrozenRows(1);
    
    // Set column widths for better readability
    sheet.setColumnWidth(1, 200); // File Name
    sheet.setColumnWidth(2, 180); // File ID
    sheet.setColumnWidth(3, 200); // File URL
    
  } catch (error) {
    errorLog('Error setting up sheet structure', error);
    throw error;
  }
}

// ========== OTHER DASHBOARD FUNCTIONS ==========

/**
 * Process Gmail for a specific client using attachment_downloader.js functionality
 */
function processClientGmail(clientId) {
  try {
    const clients = getAllClients();
    const client = clients.find(c => c.name === clientId || c.rootFolderId === clientId);
    
    if (!client) {
      return { success: false, message: 'Client not found' };
    }
    
    infoLog(`Starting Gmail processing for client: ${client.name} with label: ${client.gmailLabel}`);
    
    // Check if the required functions from attachment_downloader.js are available
    if (typeof processAttachments !== 'function') {
      errorLog('processAttachments function not found - attachment_downloader.js may not be loaded');
      return { 
        success: false, 
        message: 'Gmail processing functions not available. Please ensure attachment_downloader.js is properly loaded.' 
      };
    }
    
    // Generate a unique process token for this operation
    const processToken = `gmail_process_${client.name}_${Date.now()}`;
    
    // Use the Gmail label from the client configuration to process attachments
    // The label format should match what attachment_downloader.js expects
    const labelName = client.gmailLabel;
    
    infoLog(`Processing Gmail attachments for label: ${labelName}`);
    
    // Call the processAttachments function from attachment_downloader.js
    const result = processAttachments(labelName, processToken);
    
    if (result.status === 'success') {
      infoLog(`Successfully processed Gmail for client: ${client.name}`, result);
      return { 
        success: true, 
        message: `Gmail processed successfully for ${client.name}. ${result.message}`,
        details: result.report || {}
      };
    } else if (result.status === 'cancelled') {
      warnLog(`Gmail processing cancelled for client: ${client.name}`);
      return { 
        success: false, 
        message: `Gmail processing was cancelled for ${client.name}` 
      };
    } else {
      errorLog(`Gmail processing failed for client: ${client.name}`, result);
      return { 
        success: false, 
        message: `Gmail processing failed for ${client.name}: ${result.message}` 
      };
    }
    
  } catch (error) {
    errorLog(`Error processing Gmail for client: ${clientId}`, error);
    return { 
      success: false, 
      message: `Error processing Gmail: ${error.message}` 
    };
  }
}

/**
 * Process all clients' Gmail using attachment_downloader.js functionality
 */
function processAllGmail() {
  try {
    const clients = getAllClients();
    const activeClients = clients.filter(c => c.status === SYSTEM_CONFIG.STATUS.ACTIVE);
    
    if (activeClients.length === 0) {
      return { success: true, message: 'No active clients found to process' };
    }
    
    infoLog(`Starting Gmail processing for ${activeClients.length} active clients`);
    
    let processed = 0;
    let failed = 0;
    const results = [];
    
    for (const client of activeClients) {
      try {
        infoLog(`Processing Gmail for client: ${client.name}`);
        const result = processClientGmail(client.name);
        
        if (result.success) {
          processed++;
          results.push({
            client: client.name,
            success: true,
            message: result.message
          });
        } else {
          failed++;
          results.push({
            client: client.name,
            success: false,
            message: result.message
          });
        }
        
        // Add small delay between clients to avoid rate limits
        if (activeClients.length > 1) {
          Utilities.sleep(2000); // 2 second delay between clients
        }
        
      } catch (error) {
        failed++;
        errorLog(`Error processing Gmail for client ${client.name}`, error);
        results.push({
          client: client.name,
          success: false,
          message: `Error: ${error.message}`
        });
      }
    }
    
    const summary = `Processed Gmail for ${processed} clients successfully, ${failed} failed`;
    infoLog(summary, { processed, failed, results });
    
    return { 
      success: true, 
      message: summary,
      details: {
        totalClients: activeClients.length,
        processed: processed,
        failed: failed,
        results: results
      }
    };
  } catch (error) {
    errorLog('Error processing all Gmail:', error);
    return { success: false, message: 'Error processing Gmail: ' + error.message };
  }
}

/**
 * Run system diagnostics
 */
function runSystemDiagnostics() {
  try {
    const diagnostics = {
      masterSheetAccess: false,
      parentFolderAccess: false,
      gmailAccess: false,
      clientsCount: 0
    };
    
    // Check master sheet access
    try {
      SpreadsheetApp.openById(MASTER_SHEET_ID);
      diagnostics.masterSheetAccess = true;
    } catch (e) {
      errorLog('Master sheet access failed:', e);
    }
    
    // Check parent folder access
    try {
      DriveApp.getFolderById(PARENT_COMPANIES_FOLDER_ID);
      diagnostics.parentFolderAccess = true;
    } catch (e) {
      errorLog('Parent folder access failed:', e);
    }
    
    // Check Gmail access
    try {
      GmailApp.getInboxThreads(0, 1);
      diagnostics.gmailAccess = true;
    } catch (e) {
      errorLog('Gmail access failed:', e);
    }
    
    // Get clients count
    diagnostics.clientsCount = getAllClients().length;
    
    return { success: true, diagnostics: diagnostics };
  } catch (error) {
    errorLog('Error running diagnostics:', error);
    return { success: false, message: 'Error running diagnostics: ' + error.message };
  }
}

/**
 * Delete a client
 */
function deleteClient(clientId) {
  try {
    const clients = getAllClients();
    const client = clients.find(c => c.name === clientId || c.rootFolderId === clientId);
    
    if (!client) {
      return { success: false, message: 'Client not found' };
    }
    
    // This would implement client deletion logic
    // For now, just return success message
    return { success: true, message: `Client "${client.name}" would be deleted` };
    
  } catch (error) {
    errorLog('Error deleting client:', error);
    return { success: false, message: 'Error deleting client: ' + error.message };
  }
}

/**
 * Get dashboard metrics
 */
function getDashboardMetrics() {
  try {
    const clients = getAllClients();
    
    return {
      totalClients: clients.length,
      activeClients: clients.filter(c => c.status === SYSTEM_CONFIG.STATUS.ACTIVE).length,
      configuredClients: clients.filter(c => c.status === SYSTEM_CONFIG.STATUS.ACTIVE).length,
      systemStatus: 'healthy'
    };
  } catch (error) {
    errorLog('Error getting dashboard metrics:', error);
    return {
      totalClients: 0,
      activeClients: 0,
      configuredClients: 0,
      systemStatus: 'error'
    };
  }
}

/**
 * Open the Add Client popup
 */
function openAddClientPopup() {
  try {
    const htmlOutput = HtmlService.createTemplateFromFile('AddClientPopup')
      .evaluate()
      .setWidth(600)
      .setHeight(700)
      .setTitle('Add New Client');
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Add New Client');
    
    return { success: true, message: 'Popup opened successfully' };
  } catch (error) {
    errorLog('Error opening add client popup', error);
    return { success: false, message: 'Error opening popup: ' + error.message };
  }
}

/**
 * Get dashboard data (clients and metrics)
 */
function getDashboardData() {
  try {
    const clients = getAllClients();
    const metrics = getDashboardMetrics();
    
    return {
      success: true,
      clients: clients,
      metrics: metrics
    };
  } catch (error) {
    errorLog('Error getting dashboard data', error);
    return {
      success: false,
      clients: [],
      metrics: {
        totalClients: 0,
        activeClients: 0,
        configuredClients: 0,
        systemStatus: 'error'
      }
    };
  }
}

/**
 * Add new client (wrapper for the popup form)
 */
function addNewClient(formData) {
  try {
    if (!formData || !formData.name || !formData.gmailLabel) {
      return { success: false, message: 'Missing required fields' };
    }
    
    const result = addClientWithAtomicTransaction(formData.name, formData.gmailLabel, PARENT_COMPANIES_FOLDER_ID);
    
    if (result.success) {
      // Return additional data for the popup
      return {
        success: true,
        message: result.message,
        companyFolderId: result.folderStructure ? result.folderStructure.root : null,
        sheetId: result.spreadsheetId
      };
    } else {
      return result;
    }
  } catch (error) {
    errorLog('Error in addNewClient wrapper', error);
    return { success: false, message: 'Error adding client: ' + error.message };
  }
}

/**
 * Test function to verify the script is working
 */
function testDashboard() {
  console.log('Dashboard script is working!');
  return { success: true, message: 'Dashboard script is working!' };
}

/**
 * Sets up the folder mapping in the new client spreadsheet's Script Properties.
 * This function must be run in the context of the new spreadsheet.
 */
function setupClientSpreadsheetFolderMapping(spreadsheet, clientName, rootFolderId) {
  try {
    // Use the Apps Script API to run a function in the new spreadsheet context if needed.
    // If running in the same script, just set the property here:
    const props = PropertiesService.getScriptProperties();
    props.setProperty(`COMPANY_FOLDER_${clientName.toUpperCase()}`, rootFolderId);
    infoLog(`Set folder mapping for client ${clientName} in spreadsheet ${spreadsheet.getId()}`);
  } catch (error) {
    errorLog(`Error setting up folder mapping for client ${clientName}:`, error);
  }
}

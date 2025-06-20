//Reroute.gs

// Configuration
const SHARED_FOLDER_ID = '17g3DvSFb1qN8M9XWohyLK6ypQ2MyaHu3'; // Replace with your actual shared folder ID
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

/**
 * Get all company names from sheet tabs
 * @return {Array} Array of company names
 */
function getCompanies() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    const companies = [];
    
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      // Skip sheets that are already inflow/outflow or system sheets
      if (!sheetName.includes(' - inflow') && 
          !sheetName.includes(' - outflow') && 
          sheetName !== 'Sheet1') {
        companies.push(sheetName);
      }
    });
    
    return companies.sort();
  } catch (error) {
    console.error('Error getting companies:', error);
    throw new Error('Failed to load companies: ' + error.message);
  }
}

/**
 * Flood file details and create inflow/outflow sheets
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
    // Set to track unique invoices (invoiceNumber + vendorName + totalAmount)
    const seenInvoices = new Set();
    // Track redundant rows
    const inflowRedundant = [];
    const outflowRedundant = [];
    
    rows.forEach(row => {
      if (row[fileNameIndex]) {
        const fileName = row[fileNameIndex];
        const fileUrl = row[fileUrlIndex] || '';
        const invoiceStatus = row[invoiceStatusIndex] || 'inflow'; // default to inflow
        
        const parsedData = parseFileName(fileName, fileUrl);
        // parsedData: [date, month, vendorName, financialYear, fileUrl, invoiceNumber, '', '', '', '', totalAmount]
        const vendorName = parsedData[2] || '';
        const invoiceNumber = parsedData[5] || '';
        const totalAmount = parsedData[10] || '';
        const uniqueKey = invoiceNumber + '|' + vendorName + '|' + totalAmount;
        if (seenInvoices.has(uniqueKey)) {
          // Duplicate found, add as redundant
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
    // Merge redundant rows into main data arrays
    inflowData.push(...inflowRedundant);
    outflowData.push(...outflowRedundant);
    // Create or update inflow sheet
    createOrUpdateFlowSheet(companyName + ' - inflow', inflowData);
    // Create or update outflow sheet
    createOrUpdateFlowSheet(companyName + ' - outflow', outflowData);
    return {
      success: true,
      message: `Processed ${inflowData.length} inflow and ${outflowData.length} outflow records`
    };
  } catch (error) {
    console.error('Error flooding file details:', error);
    throw new Error('Failed to flood file details: ' + error.message);
  }
}

/**
 * Parse file name and extract relevant information
 * @param {string} fileName - File name in format Date_VendorName_InvoiceNumber_TotalAmount
 * @param {string} fileUrl - File URL
 * @return {Array} Parsed data array
 */
function parseFileName(fileName, fileUrl) {
  try {
    // Parse filename format: Date_VendorName_InvoiceNumber_TotalAmount
    const parts = fileName.split('_');
    
    if (parts.length < 4) {
      // If filename doesn't match expected format, use defaults
      return [
        '', // Date
        '', // Month
        fileName, // Vendor Name (use full filename)
        '', // Financial Year
        fileUrl, // Document Link
        '', // Document Number
        '', // Gross Amount
        '', // GST
        '', // TDS
        '', // Other Taxes
        '' // Net Amount
      ];
    }
    
    const date = parts[0];
    const vendorName = parts[1];
    const invoiceNumber = parts[2];
    const totalAmount = parts[3].replace(/\.[^.]*$/, ''); // Remove file extension
    
    // Calculate month and financial year
    const month = getMonthFromDate(date);
    const financialYear = calculateFinancialYear(date);
    
    return [
      date, // Date
      month, // Month
      vendorName, // Vendor Name
      financialYear, // Financial Year
      fileUrl, // Document Link
      invoiceNumber, // Document Number
      '', // Gross Amount (empty)
      '', // GST (empty)
      '', // TDS (empty)
      '', // Other Taxes (empty)
      totalAmount // Net Amount
    ];
    
  } catch (error) {
    console.error('Error parsing filename:', fileName, error);
    // Return default row if parsing fails
    return [
      '', '', fileName, '', fileUrl, '', '', '', '', '', ''
    ];
  }
}

/**
 * Get month name from date string
 * @param {string} dateStr - Date in YYYY-MM-DD format
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
 * Calculate financial year from date
 * @param {string} dateStr - Date in YYYY-MM-DD format
 * @return {string} Financial year in format YYYY-YYYY
 */
function calculateFinancialYear(dateStr) {
  try {
    const date = new Date(dateStr);
    const year = date.getFullYear();
    const month = date.getMonth() + 1; // getMonth() returns 0-11
    
    if (month >= 4) {
      // April to March of next year
      return `${year}-${year + 1}`;
    } else {
      // January to March of current year belongs to previous financial year
      return `${year - 1}-${year}`;
    }
  } catch (error) {
    return '';
  }
}

/**
 * Create or update flow sheet with data
 * @param {string} sheetName - Name of the sheet
 * @param {Array} data - Data to populate
 */
function createOrUpdateFlowSheet(sheetName, data) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
    } else {
      // Clear existing data
      sheet.clear();
    }
    
    // Set headers (add 'Redundant' column)
    const headers = [
      'Date', 'Month', 'Vendor Name', 'Financial Year', 'Document Link',
      'Document Number', 'Gross Amount', 'GST', 'TDS', 'Other Taxes', 'Net Amount', 'Redundant'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    
    // Add data if available
    if (data && data.length > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
    }
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, headers.length);
    
  } catch (error) {
    console.error('Error creating/updating flow sheet:', error);
    throw error;
  }
}

/**
 * Reroute files to appropriate folders
 * @param {string} flowSheetName - Name of the flow sheet (company - inflow/outflow)
 * @return {Object} Success response
 */
function rerouteFiles(flowSheetName) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const flowSheet = spreadsheet.getSheetByName(flowSheetName);
    
    if (!flowSheet) {
      throw new Error(`Flow sheet '${flowSheetName}' not found`);
    }
    
    // Parse company name and flow type
    const parts = flowSheetName.split(' - ');
    const companyName = parts[0];
    const flowType = parts[1]; // 'inflow' or 'outflow'
    
    // Get data from flow sheet
    const data = flowSheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    // Find required column indices
    const dateIndex = headers.indexOf('Date');
    const monthIndex = headers.indexOf('Month');
    const financialYearIndex = headers.indexOf('Financial Year');
    const documentLinkIndex = headers.indexOf('Document Link');
    
    if (dateIndex === -1 || monthIndex === -1 || financialYearIndex === -1 || documentLinkIndex === -1) {
      throw new Error('Required columns not found in flow sheet');
    }
    
    let processedCount = 0;
    const sharedFolder = DriveApp.getFolderById(SHARED_FOLDER_ID);
    
    // Process each row
    rows.forEach(row => {
      if (row[documentLinkIndex] && row[monthIndex] && row[financialYearIndex]) {
        try {
          const fileUrl = row[documentLinkIndex];
          const month = row[monthIndex];
          const financialYear = row[financialYearIndex];
          
          // Extract file ID from URL
          const fileId = extractFileIdFromUrl(fileUrl);
          if (!fileId) {
            console.warn('Could not extract file ID from URL:', fileUrl);
            return;
          }
          
          // Get the file
          const file = DriveApp.getFileById(fileId);
          
          // Create folder structure and move file
          const targetFolder = createFolderStructure(sharedFolder, companyName, financialYear, month, flowType);
          
          // Move file to target folder
          file.getParents().next().removeFile(file);
          targetFolder.addFile(file);
          
          processedCount++;
          
        } catch (fileError) {
          console.error('Error processing file:', row[documentLinkIndex], fileError);
        }
      }
    });
    
    return {
      success: true,
      message: `Successfully rerouted ${processedCount} files`
    };
    
  } catch (error) {
    console.error('Error rerouting files:', error);
    throw new Error('Failed to reroute files: ' + error.message);
  }
}

/**
 * Extract file ID from Google Drive URL
 * @param {string} url - Google Drive file URL
 * @return {string} File ID or null
 */
function extractFileIdFromUrl(url) {
  try {
    const regex = /\/file\/d\/([a-zA-Z0-9-_]+)/;
    const match = url.match(regex);
    return match ? match[1] : null;
  } catch (error) {
    console.error('Error extracting file ID:', error);
    return null;
  }
}

/**
 * Create folder structure: Company/FinancialYear/Month/Month-FlowType
 * @param {Folder} parentFolder - Parent folder
 * @param {string} companyName - Company name
 * @param {string} financialYear - Financial year (e.g., "2025-2026")
 * @param {string} month - Month name
 * @param {string} flowType - "inflow" or "outflow"
 * @return {Folder} Target folder
 */
function createFolderStructure(parentFolder, companyName, financialYear, month, flowType) {
  try {
    // Create or get company folder
    const companyFolder = getOrCreateFolder(parentFolder, companyName);
    
    // Create or get financial year folder
    const financialYearFolder = getOrCreateFolder(companyFolder, financialYear);
    
    // Create or get month folder
    const monthFolder = getOrCreateFolder(financialYearFolder, month);
    
    // Create or get flow type folder (e.g., "January-inflow")
    const flowFolderName = `${month}-${flowType}`;
    const flowFolder = getOrCreateFolder(monthFolder, flowFolderName);
    
    return flowFolder;
    
  } catch (error) {
    console.error('Error creating folder structure:', error);
    throw error;
  }
}

/**
 * Get existing folder or create new one
 * @param {Folder} parentFolder - Parent folder
 * @param {string} folderName - Name of folder to create/get
 * @return {Folder} The folder
 */
function getOrCreateFolder(parentFolder, folderName) {
  try {
    const folders = parentFolder.getFoldersByName(folderName);
    
    if (folders.hasNext()) {
      return folders.next();
    } else {
      return parentFolder.createFolder(folderName);
    }
  } catch (error) {
    console.error('Error creating/getting folder:', folderName, error);
    throw error;
  }
}

/**
 * Serve HTML page
 * @return {HtmlOutput} HTML page
 */
function doGet() {
  return HtmlService.createTemplateFromFile('RerouteIndex')
      .evaluate()
      .setTitle('File Rerouting Tool')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Include HTML file content
 * @param {string} filename - Name of HTML file
 * @return {string} HTML content
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
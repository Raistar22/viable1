// Configuration
const SHARED_FOLDER_ID = '1Tqj8May8je0L1lET5PIHRJj4P8d9T3MT'; // Replace with your actual shared folder ID
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// Minimal folder mappings - only main company folders needed
const REROUTE_COMPANY_FOLDER_MAP = {
  "analogy": "160pN2zDCb9UQbwIXqgggdTjLUrFM2cM3", // Main Analogy Folder
  "humane": "1E6ijhWhdYykymN0MEUINd9jETmdM2sAt"   // Main Humane Folder
};

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
  
   rows.forEach(row => {
     if (row[fileNameIndex]) {
       const fileName = row[fileNameIndex];
       const fileUrl = row[fileUrlIndex] || '';
       const invoiceStatus = row[invoiceStatusIndex] || 'inflow'; // default to inflow
      
       const parsedData = parseFileName(fileName, fileUrl);
      
       if (invoiceStatus.toLowerCase().includes('outflow')) {
         outflowData.push(parsedData);
       } else {
         inflowData.push(parsedData);
       }
     }
   });
  
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
* @return {string} Financial year in format FY-XX-XX
*/
function calculateFinancialYear(dateStr) {
 try {
   const date = new Date(dateStr);
   const year = date.getFullYear();
   const month = date.getMonth() + 1; // getMonth() returns 0-11
  
   if (month >= 4) {
     // April to March of next year
     return `FY-${year.toString().slice(-2)}-${(year + 1).toString().slice(-2)}`;
   } else {
     // January to March of current year belongs to previous financial year
     return `FY-${(year - 1).toString().slice(-2)}-${year.toString().slice(-2)}`;
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
  
   // Set headers
   const headers = [
     'Date', 'Month', 'Vendor Name', 'Financial Year', 'Document Link',
     'Document Number', 'Gross Amount', 'GST', 'TDS', 'Other Taxes', 'Net Amount'
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
* Reroute files to appropriate folders using existing folder structure
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
  
   // Get the main company folder ID
   const companyFolderId = REROUTE_COMPANY_FOLDER_MAP[companyName.toLowerCase()];
   if (!companyFolderId) {
     throw new Error(`No folder mapping found for company: ${companyName}`);
   }
  
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
     throw new Error('Required columns (Date, Month, Financial Year, Document Link) not found in flow sheet');
   }
  
   let processedCount = 0;
   const companyFolder = DriveApp.getFolderById(companyFolderId);
  
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
         const targetFolder = createFolderStructure(companyFolder, financialYear, month, flowType);
         
         // Check if file is already in the target folder
         const parents = file.getParents();
         let isAlreadyInTarget = false;
         while (parents.hasNext()) {
           const parent = parents.next();
           if (parent.getId() === targetFolder.getId()) {
             isAlreadyInTarget = true;
             break;
           }
         }
         
         if (!isAlreadyInTarget) {
           // Move file to target folder
           file.getParents().next().removeFile(file);
           targetFolder.addFile(file);
           processedCount++;
           console.log(`Moved file ${file.getName()} to ${financialYear}/${month}/${flowType} folder`);
         } else {
           console.log(`File ${file.getName()} is already in ${financialYear}/${month}/${flowType} folder`);
         }
         
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
* Create folder structure: Company/FY-XX-XX/Accruals/Bills and Invoices/Month/FlowType
* @param {Folder} companyFolder - Company folder
* @param {string} financialYear - Financial year (e.g., "FY-24-25")
* @param {string} month - Month name
* @param {string} flowType - "inflow" or "outflow"
* @return {Folder} Target folder
*/
function createFolderStructure(companyFolder, financialYear, month, flowType) {
 try {
   // Create or get financial year folder
   const financialYearFolder = getOrCreateFolder(companyFolder, financialYear);
   
   // Create or get Accruals folder
   const accrualsFolder = getOrCreateFolder(financialYearFolder, "Accruals");
   
   // Create or get Bills and Invoices folder
   const billsFolder = getOrCreateFolder(accrualsFolder, "Bills and Invoices");
   
   // Create or get month folder
   const monthFolder = getOrCreateFolder(billsFolder, month);
   
   // Create or get flow type folder (Inflow or Outflow)
   const flowFolderName = flowType.charAt(0).toUpperCase() + flowType.slice(1); // Capitalize first letter
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


// IngestInvoice.gs - Complete Code for Invoice Buffer and Ingestion

// Instructions:

// Create Two New Folders in Google Drive:

// Incoming Invoice Drop Zone: This is where files will initially land. Get its Folder ID.
// Raw Invoices Buffer: This is where your script will store its copies. Get its Folder ID.
// Open Your Apps Script Project: Go to your Google Sheet, then Extensions > Apps Script.

// Create a New Script File: In the Apps Script editor, click + next to "Files" and select Script. Name it IngestInvoice.gs.

// Paste the Code: Copy the entire code block below and paste it into the IngestInvoice.gs file.

// Update Configuration Variables:

// Replace 'YOUR_INCOMING_DROP_FOLDER_ID' with the actual ID of your Incoming Invoice Drop Zone folder.
// Replace 'YOUR_RAW_INVOICES_BUFFER_FOLDER_ID' with the actual ID of your Raw Invoices Buffer folder.
// Ensure Company Sheets Exist: Make sure you have sheets in your spreadsheet named after the companies you expect (e.g., "CompanyA", "CompanyB"). These sheets must have the headers File Name, File URL, and invoice status (and optionally Date, Month, Vendor Name, Financial Year, Document Number, Net Amount) in the first row.

// Set Up a Trigger:

// In the Apps Script editor, click the Triggers icon (looks like a clock) on the left sidebar.
// Click + Add Trigger.
// Choose function to run: ingestFilesFromDropZone
// Choose deployment where function is: Head
// Select event source: From Drive
// Select event type: On change
// Save.
// --- Configuration ---
const INCOMING_DROP_FOLDER_ID = 'YOUR_INCOMING_DROP_FOLDER_ID';     // <<< IMPORTANT: REPLACE WITH YOUR FOLDER ID
const RAW_INVOICES_BUFFER_FOLDER_ID = 'YOUR_RAW_INVOICES_BUFFER_FOLDER_ID'; // <<< IMPORTANT: REPLACE WITH YOUR FOLDER ID
const MAIN_SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId(); // The ID of the spreadsheet containing company sheets

// --- Helper Functions (Copied from Reroute.gs for self-containment) ---

/**
 * Parse file name and extract relevant information
 * @param {string} fileName - File name in format Date_VendorName_InvoiceNumber_TotalAmount
 * @param {string} fileUrl - File URL
 * @return {Array} Parsed data array
 */
function parseFileName(fileName, fileUrl) {
  try {
    // Expected filename format: Date_VendorName_InvoiceNumber_TotalAmount.extension
    const parts = fileName.split('_');

    if (parts.length < 4) {
      // If filename doesn't match expected format, use defaults
      // Remove file extension from fileName if used as vendor name
      const cleanFileName = fileName.substring(0, fileName.lastIndexOf('.')) || fileName;
      return [
        '', // Date (index 0)
        '', // Month (index 1)
        cleanFileName, // Vendor Name (use full filename, remove extension) (index 2)
        '', // Financial Year (index 3)
        fileUrl, // Document Link (index 4)
        '', // Document Number (index 5)
        '', // Gross Amount (index 6)
        '', // GST (index 7)
        '', // TDS (index 8)
        '', // Other Taxes (index 9)
        ''  // Net Amount (index 10)
      ];
    }

    const date = parts[0];
    const vendorName = parts[1];
    const invoiceNumber = parts[2];
    const totalAmount = parts[3].substring(0, parts[3].lastIndexOf('.')); // Remove file extension from amount part

    // Calculate month and financial year
    const month = getMonthFromDate(date);
    const financialYear = calculateFinancialYear(date);

    return [
      date,          // Date
      month,         // Month
      vendorName,    // Vendor Name
      financialYear, // Financial Year
      fileUrl,       // Document Link
      invoiceNumber, // Document Number
      '',            // Gross Amount (empty)
      '',            // GST (empty)
      '',            // TDS (empty)
      '',            // Other Taxes (empty)
      totalAmount    // Net Amount
    ];

  } catch (error) {
    console.error('Error parsing filename:', fileName, error);
    // Return default row if parsing fails
    const cleanFileName = fileName.substring(0, fileName.lastIndexOf('.')) || fileName;
    return [
      '', '', cleanFileName, '', fileUrl, '', '', '', '', '', ''
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
    if (isNaN(date.getTime())) { // Check for invalid date
        return '';
    }
    const months = [
      'January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December'
    ];
    return months[date.getMonth()];
  } catch (error) {
    console.error('Error getting month from date:', dateStr, error);
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
    if (isNaN(date.getTime())) { // Check for invalid date
        return '';
    }
    const year = date.getFullYear();
    const month = date.getMonth() + 1; // getMonth() returns 0-11

    if (month >= 4) {
      // April (4) to March (3) of next year, e.g., April 2025 -> 2025-2026
      return `${year}-${year + 1}`;
    } else {
      // January (1) to March (3) of current year belongs to previous financial year, e.g., Jan 2025 -> 2024-2025
      return `${year - 1}-${year}`;
    }
  } catch (error) {
    console.error('Error calculating financial year:', dateStr, error);
    return '';
  }
}

/**
 * Get existing folder or create new one
 * @param {GoogleAppsScript.Drive.Folder} parentFolder - Parent folder
 * @param {string} folderName - Name of folder to create/get
 * @return {GoogleAppsScript.Drive.Folder} The folder
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
    throw error; // Re-throw to indicate a critical failure
  }
}

/**
 * Get all company names from sheet tabs
 * This function assumes company sheets are not named ' - inflow' or ' - outflow' sheets or 'Sheet1'.
 * @return {Array} Array of company names (sheet names)
 */
function getCompanies() {
  try {
    const spreadsheet = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
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
    console.error('Error getting companies (from IngestInvoice.gs):', error);
    // Return empty array if there's an issue loading companies
    return [];
  }
}

// --- Main Ingestion Logic ---

/**
 * Main function to be called by a Google Drive "On change" trigger.
 * It processes new files in the INCOMING_DROP_FOLDER_ID, copies them to the buffer,
 * updates the main spreadsheet, and moves the original file.
 */
function ingestFilesFromDropZone() {
  try {
    const dropFolder = DriveApp.getFolderById(INCOMING_DROP_FOLDER_ID);
    const rawBufferFolder = DriveApp.getFolderById(RAW_INVOICES_BUFFER_FOLDER_ID);
    const spreadsheet = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);

    // Get files in the drop folder that are not in the 'Processed' subfolder
    // This assumes 'Processed' is a direct child of 'dropFolder'
    const files = dropFolder.getFiles();
    const processedFolderName = 'Processed'; // Name of the subfolder for processed files
    let processedFolder;
    try {
      processedFolder = getOrCreateFolder(dropFolder, processedFolderName);
    } catch (e) {
      console.warn(`Warning: Could not create 'Processed' folder in drop zone (${INCOMING_DROP_FOLDER_ID}). Original files will be trashed.`);
      // If processedFolder creation fails, it will remain null/undefined
    }


    while (files.hasNext()) {
      const file = files.next();

      // Skip folders or files that are already in the 'Processed' folder
      // (This check helps if the trigger fires before the move is complete,
      // though the 'move' logic later is the primary prevention)
      const parentFolders = file.getParents();
      let isAlreadyProcessed = false;
      while(parentFolders.hasNext()){
        const parent = parentFolders.next();
        if(parent.getName() === processedFolderName && parent.getId() === processedFolder?.getId()){
          isAlreadyProcessed = true;
          break;
        }
      }
      if (isAlreadyProcessed) {
        console.log(`Skipping file '${file.getName()}' as it's already in 'Processed' folder.`);
        continue;
      }

      // --- Determine Company Name ---
      let companyName = null;
      const fileDirectParents = file.getParents(); // Get iterators for parents
      if (fileDirectParents.hasNext()) {
        const directParentFolder = fileDirectParents.next();
        if (directParentFolder.getId() !== INCOMING_DROP_FOLDER_ID) {
          // If the file is in a subfolder directly under the drop zone,
          // assume the subfolder name is the company name.
          companyName = directParentFolder.getName();
        }
      }

      // If company name not derived from subfolder, try to infer or use default
      if (!companyName) {
        // Option 1: Try to get from filename if it contains company info (requires specific filename format)
        // This is complex to generalize. For simplicity, we'll try to find an existing company sheet.

        // Option 2: Attempt to find a suitable company sheet by matching file name part to company name
        // This is a heuristic and might need refinement for your specific naming conventions.
        const allCompanies = getCompanies();
        const fileNameWithoutExt = file.getName().substring(0, file.getName().lastIndexOf('.')) || file.getName();

        for (const comp of allCompanies) {
            // Check if the filename contains the company name, case-insensitive
            if (fileNameWithoutExt.toLowerCase().includes(comp.toLowerCase())) {
                companyName = comp;
                break;
            }
        }
      }

      // If no company name determined, log error and skip or use a fallback 'Uncategorized' sheet
      if (!companyName) {
        console.error(`Could not determine company for file: '${file.getName()}'. Please ensure it's in a company-named subfolder or filename contains company name.`);
        // You could uncomment the line below to route to a default 'Uncategorized' sheet
        // companyName = 'Uncategorized'; // Make sure this sheet exists in your spreadsheet!
        continue; // Skip this file if no company sheet can be determined
      }

      const targetCompanySheet = spreadsheet.getSheetByName(companyName);
      if (!targetCompanySheet) {
        console.error(`Company sheet '${companyName}' not found in spreadsheet '${MAIN_SPREADSHEET_ID}'. Skipping file: ${file.getName()}`);
        continue;
      }
      // --- End Determine Company Name ---

      // 1. Make a copy of the file into your controlled buffer folder
      const copiedFile = file.makeCopy(file.getName(), rawBufferFolder); // Preserve original filename
      const newFileUrl = copiedFile.getUrl(); // Get the URL of YOUR copy

      // 2. Extract metadata from the file name using the parseFileName function
      // parsedData: [date, month, vendorName, financialYear, fileUrl, invoiceNumber, '', '', '', '', totalAmount]
      const parsedData = parseFileName(file.getName(), newFileUrl);

      // 3. Store metadata in the 'company' sheet
      // Headers in the company sheet (first row) must be:
      // 'File Name', 'File URL', 'invoice status', 'Date', 'Month', 'Vendor Name', 'Financial Year', 'Document Number', 'Net Amount'
      const headers = targetCompanySheet.getDataRange().getValues()[0];
      const newRowData = new Array(headers.length).fill(''); // Initialize with empty strings

      // Map parsedData to corresponding header columns
      const headerMap = {
        'File Name': file.getName(),
        'File URL': newFileUrl,
        'invoice status': 'inflow', // Default status upon initial ingestion
        'Date': parsedData[0],
        'Month': parsedData[1],
        'Vendor Name': parsedData[2],
        'Financial Year': parsedData[3],
        'Document Number': parsedData[5],
        'Net Amount': parsedData[10]
      };

      // Populate the newRowData array based on header positions
      headers.forEach((header, index) => {
        if (headerMap.hasOwnProperty(header)) {
          newRowData[index] = headerMap[header];
        }
      });

      targetCompanySheet.appendRow(newRowData);
      console.log(`Ingested: '${file.getName()}'. Copied to buffer, added to sheet '${companyName}'.`);

      // 4. Move the original file from the drop folder to a "Processed" subfolder or trash it
      // This is crucial to avoid re-processing the same file on subsequent trigger runs.
      try {
        if (processedFolder) {
          // Remove from current parent (dropFolder) and add to 'Processed' subfolder
          file.getParents().next().removeFile(file); // This should be the direct parent (dropFolder or its subfolder)
          processedFolder.addFile(file);
          console.log(`Original file '${file.getName()}' moved to 'Processed' folder.`);
        } else {
          // Fallback to trashing if 'Processed' folder could not be created/found
          file.setTrashed(true);
          console.log(`Original file '${file.getName()}' trashed (Processed folder not available).`);
        }
      } catch (moveError) {
        console.error(`Failed to move/trash original file '${file.getName()}':`, moveError);
        // Log, but don't stop the overall ingestion process for other files
      }
    }
    console.log('Finished processing files in drop zone.');

  } catch (error) {
    console.error('CRITICAL ERROR in ingestFilesFromDropZone:', error);
    // Consider sending an email notification for critical errors
    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(), // Or a specific admin email
      subject: 'Apps Script Error: Invoice Ingestion Failed',
      body: `An error occurred during invoice ingestion: ${error.message}\nStack: ${error.stack}`
    });
  }
}
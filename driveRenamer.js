// this driveRenamer.js can handle renaming of the multiple vendors and multiple invoices

// ----- Backend Functions called from HTML -----


// Add this constant at the top of your script
const GEMINI_MODEL_ID_AIS = "gemini-2.0-flash"; // Updated to use Gemini 2.0 Flash


// Enhanced regex pattern for stricter naming convention validation
// Pattern: YYYY-MM-DD_VendorName_InvoiceNumber_Amount.extension
// - Date must be valid format YYYY-MM-DD
// - Vendor name: alphanumeric, spaces, hyphens, periods, ampersands (NO underscores)
// - Invoice number: alphanumeric, hyphens, periods (NO underscores)
// - Amount: numeric with optional decimal and currency symbols
const FILE_NAME_REGEX = /^[0-9]{4}-[0-9]{2}-[0-9]{2}_[a-zA-Z0-9 \-.&]+_[a-zA-Z0-9\-\.]+_[\d\.,\$€£¥₹]+(\.[a-zA-Z0-9]+)?$/i;


/**
* Enhanced function to validate if a filename follows the naming convention
* @param {string} filename - The filename to validate
* @returns {Object} - Validation result with details
*/
function validateFilenameConvention(filename) {
 if (!filename || typeof filename !== 'string') {
   return {
     isValid: false,
     reason: "Invalid filename provided",
     details: "Filename is empty or not a string"
   };
 }


 const trimmedName = filename.trim();


 // Primary check: Does it match the strict regex pattern?
 if (!FILE_NAME_REGEX.test(trimmedName)) {
   // If regex fails, it's definitely not valid. Now, provide detailed reasons.
   let reason = "Does not match naming convention";
   let details = `Expected format YYYY-MM-DD_VendorName_InvoiceNumber_Amount.extension. Actual: "${trimmedName}".`;


   const parts = trimmedName.split('_');


   // Check for insufficient parts first
   if (parts.length < 4) {
     reason = "Insufficient parts in filename";
     details += ` Expected at least 4 parts separated by underscores (Date, Vendor, Invoice, Amount), found ${parts.length}.`;
     return { isValid: false, reason: reason, details: details };
   }


   const datePart = parts[0];
   const vendorPart = parts[1];
   const invoicePart = parts[2];
   // Rejoin remaining parts for amount and extension, as regex failure might mean more underscores
   const amountWithExtension = parts.slice(3).join('_');
   const lastDotIndex = amountWithExtension.lastIndexOf('.');
   const amountPart = lastDotIndex > 0 ? amountWithExtension.substring(0, lastDotIndex) : amountWithExtension;
   const extensionPart = lastDotIndex > 0 ? amountWithExtension.substring(lastDotIndex) : '';


   // Specific checks for common violations
   const dateRegex = /^[0-9]{4}-[0-9]{2}-[0-9]{2}$/;
   if (!dateRegex.test(datePart)) {
     reason = "Invalid date format";
     details += ` Date part ("${datePart}") is not in YYYY-MM-DD format.`;
   } else {
     const [year, month, day] = datePart.split('-').map(Number);
     const date = new Date(year, month - 1, day);
     if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
       reason = "Invalid calendar date";
       details += ` "${datePart}" is not a valid calendar date.`;
     }
   }


   if (vendorPart.includes('_') || vendorPart.trim().length === 0 || !/^[a-zA-Z0-9 \-.&]+$/.test(vendorPart)) {
     if (reason === "Does not match naming convention") reason = "Invalid vendor name"; // Refine reason
     details += ` Vendor name ("${vendorPart}") contains invalid characters (e.g., underscores) or is empty.`;
   }


   if (invoicePart.includes('_') || invoicePart.trim().length === 0 || !/^[a-zA-Z0-9\-\.]+$/.test(invoicePart)) {
     if (reason === "Does not match naming convention") reason = "Invalid invoice number"; // Refine reason
     details += ` Invoice number ("${invoicePart}") contains invalid characters (e.g., underscores) or is empty.`;
   }


   // Basic check for amount (though regex is more thorough)
   if (amountPart.trim().length === 0) {
     if (reason === "Does not match naming convention") reason = "Missing amount";
     details += ` Amount part is empty.`;
   } else if (!/^[\d\.,\$€£¥₹]+$/.test(amountPart)) {
      if (reason === "Does not match naming convention") reason = "Invalid amount format";
      details += ` Amount part ("${amountPart}") contains invalid characters.`;
   }


   // This handles cases where parts might be fine individually but the overall structure isn't,
   // or if unexpected characters exist outside of the specifically checked parts.
   if (reason === "Does not match naming convention" && details.includes("Expected format")) {
       // Fallback to the general reason if no specific sub-violation was identified
       details += " General structure or character set violation.";
   }




   return {
     isValid: false,
     reason: reason,
     details: details
   };
 }


 // If regex test passes, then extract parts and return details
 const parts = trimmedName.split('_');
 const datePart = parts[0];
 const vendorPart = parts[1];
 const invoicePart = parts[2];
 const amountWithExtension = parts.slice(3).join('_'); // Rejoin in case there are more underscores (though regex should prevent this now)
 const lastDotIndex = amountWithExtension.lastIndexOf('.');
 const amountPart = lastDotIndex > 0 ? amountWithExtension.substring(0, lastDotIndex) : amountWithExtension;


 return {
   isValid: true,
   reason: "Valid naming convention",
   details: {
     date: datePart,
     vendor: vendorPart,
     invoice: invoicePart,
     amount: amountPart
   }
 };
}


/**
* Serves the HTML file for the web application.
* @returns {GoogleAppsScript.HTML.HtmlOutput} The HTML output for the web app.
*/
function doGet() {
 return HtmlService.createHtmlOutputFromFile('Index')
   .setTitle('Drive File Renamer')
   .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}


/**
* Retrieves the names of all sub-sheets in the active Google Spreadsheet.
* @returns {string[]} An array of sheet names.
* @throws {Error} If sheet names cannot be retrieved.
*/
function getSubsheetNamesGAS() {
 try {
   const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
   console.log("Successfully retrieved sheet names.");
   return sheets.map(sheet => sheet.getName());
 } catch (e) {
   console.error("Error in getSubsheetNamesGAS: %s", e.toString(), e.stack);
   throw new Error("Could not retrieve subsheet names. Please ensure you have access to the spreadsheet and the script is authorized.");
 }
}


/**
* Enhanced function to load file information from a specified sheet for AI processing.
* It identifies Google Drive file IDs/URLs in Column B, checks their naming convention,
* and compiles a list of files that need AI processing.
* @param {string} sheetName The name of the sheet to process.
* @returns {Object[]} An array of file objects to be processed.
* @throws {Error} If the sheet is not found or other loading errors occur.
*/
function loadFilesForProcessingGAS(sheetName) {
 console.log(`Attempting to load files from sheet: "${sheetName}"`);
 try {
   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
   if (!sheet) {
     throw new Error(`Sheet "${sheetName}" not found. Please ensure the sheet name is correct.`);
   }


   const lastRow = sheet.getLastRow();
   console.log(`Sheet "${sheetName}" has ${lastRow} rows.`);


   if (lastRow < 2) {
     console.log("No data rows found (sheet only has header row or is empty).");
     return [];
   }


   // Get values from Column A, B, and C, starting from the second row (skipping header)
   const dataRange = sheet.getRange("A2:C" + lastRow);
   const values = dataRange.getValues();
   const filesToProcess = [];
   const skippedEntries = [];


   for (let i = 0; i < values.length; i++) {
     const fileNameFromSheet = values[i][0]; // Column A: File Name (Potentially old name)
     const fileIdFromSheet = values[i][1];   // Column B: File ID
     const fileUrlFromSheet = values[i][2];  // Column C: File URL
     const sheetRowNumber = i + 2;
     console.log(`Processing row ${sheetRowNumber}: File Name (from sheet): "${fileNameFromSheet}", File ID: "${fileIdFromSheet}"`);


     if (!fileIdFromSheet || typeof fileIdFromSheet !== 'string' || fileIdFromSheet.trim() === '') {
       console.log(`Row ${sheetRowNumber} has no valid File ID. Skipping.`);
       skippedEntries.push(`Row ${sheetRowNumber}: No valid File ID found in Column B.`);
       continue;
     }


     try {
       const file = DriveApp.getFileById(fileIdFromSheet.trim());
       const actualFileName = file.getName(); // Get the actual name from Drive
       console.log(`Accessed file in Drive: "${actualFileName}" (ID: ${fileIdFromSheet})`);


       // Perform validation on the actual file name from Drive
       const validation = validateFilenameConvention(actualFileName);


       if (!validation.isValid) {
         console.log(`File "${actualFileName}" needs processing. Reason: ${validation.reason}`);
         console.log(`Validation details: ${validation.details}`);


         filesToProcess.push({
           fileId: fileIdFromSheet.trim(),
           originalName: actualFileName, // Use the actual name from Drive
           sheetRow: sheetRowNumber,
           mimeType: file.getMimeType(),
           fileUrl: file.getUrl(),
           validationResult: validation
         });
       } else {
         console.log(`File "${actualFileName}" already properly named. Skipping AI processing.`);
         console.log(`Validation details:`, validation.details);
       }
     } catch (e) {
       const errorMessage = `Could not access file ID "${fileIdFromSheet}" from sheet "${sheetName}", row ${sheetRowNumber}. Error: ${e.message}`;
       console.warn(errorMessage);
       skippedEntries.push(`Row ${sheetRowNumber}: ${errorMessage}`);
     }
   }


   console.log(`Found ${filesToProcess.length} files to process in sheet "${sheetName}".`);
   if (skippedEntries.length > 0) {
     console.log(`Skipped ${skippedEntries.length} entries:`, skippedEntries);
   }
   return filesToProcess;
 } catch (e) {
   console.error("Error in loadFilesForProcessingGAS: %s", e.toString(), e.stack);
   throw new Error("Could not load files for processing: " + e.message);
 }
}


/**
* Enhanced function to process a single file using the Gemini AI model to extract information.
* Also prepares preview data based on the file's MIME type.
* @param {string} fileId The Google Drive ID of the file.
* @param {string} originalName The original name of the file.
* @param {string} mimeType The MIME type of the file.
* @param {string} fileUrl The Google Drive URL of the file.
* @returns {Object} An object containing extracted AI data and preview data, or an error.
*/
function processFileWithAIGAS(fileId, originalName, mimeType, fileUrl) {
  console.log(`Starting AI processing for file ID: ${fileId}, Name: "${originalName}", MIME: ${mimeType}`);
  let previewData = { type: 'unsupported', content: 'Preview not available for this file type.' };
  let base64Content = null;
  let ocrText = '';

  try {
    // Validate file ID
    if (!fileId || typeof fileId !== 'string' || fileId.trim() === '') {
      throw new Error('Invalid file ID provided');
    }

    // Get the file from Drive
    const file = DriveApp.getFileById(fileId.trim());
    if (!file) {
      throw new Error(`File with ID ${fileId} not found or inaccessible`);
    }

    // Check if file is trashed
    if (file.isTrashed()) {
      throw new Error(`File with ID ${fileId} is in trash`);
    }

    // Get blob with validation
    let blob = null;
    
    // Handle Google Workspace files (Docs, Sheets, Slides) by exporting them
    const googleWorkspaceMimeTypes = [
      'application/vnd.google-apps.document',    // Google Docs
      'application/vnd.google-apps.spreadsheet', // Google Sheets  
      'application/vnd.google-apps.presentation' // Google Slides
    ];

    if (googleWorkspaceMimeTypes.includes(mimeType)) {
      console.log(`Converting Google Workspace file "${originalName}" to PDF for AI processing`);
      try {
        blob = Drive.Files.export(fileId, 'application/pdf');
        mimeType = 'application/pdf'; // Update MIME type for processing
      } catch (exportError) {
        console.error(`Failed to export Google Workspace file: ${exportError.message}`);
        return { 
          error: `Cannot process Google Workspace file "${originalName}". Export failed: ${exportError.message}`, 
          confidence: "Error", 
          previewData: previewData 
        };
      }
    } else {
      // Regular file - get blob directly
      blob = file.getBlob();
    }

    // Validate blob
    if (!blob) {
      throw new Error(`Unable to obtain blob for file ${fileId}. File may be corrupted or inaccessible.`);
    }

    // Validate blob has required methods
    if (typeof blob.getContentType !== 'function' || typeof blob.getBytes !== 'function') {
      throw new Error(`Invalid blob object for file ${fileId}. Expected Blob but got: ${typeof blob}`);
    }

    // OCR for images and PDFs
    if (mimeType.startsWith("image/") || mimeType === "application/pdf") {
      try {
        ocrText = extractTextWithOCR(fileId);
        console.log(`OCR completed for "${originalName}". Text length: ${ocrText.length} characters`);
      } catch (ocrError) {
        console.warn(`OCR failed for "${originalName}": ${ocrError.message}. Continuing without OCR text.`);
        ocrText = '';
      }
    }

    // Prepare preview data for frontend
    const commonOfficeMimeTypes = [
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document", // .docx
      "application/msword", // .doc
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", // .xlsx
      "application/vnd.ms-excel", // .xls
      "application/vnd.openxmlformats-officedocument.presentationml.presentation", // .pptx
      "application/vnd.ms-powerpoint" // .ppt
    ];

    if (mimeType.startsWith("image/")) {
      base64Content = Utilities.base64Encode(blob.getBytes());
      previewData = { type: 'image', content: `data:${mimeType};base64,${base64Content}` };
    } else if (mimeType === "application/pdf") {
      base64Content = Utilities.base64Encode(blob.getBytes());
      previewData = { type: 'pdf', content: fileId }; // For PDF, we pass the fileId for direct embedding
    } else if (commonOfficeMimeTypes.includes(mimeType)) {
      previewData = { type: 'office', content: fileUrl };
    }

    // Call Gemini API, passing OCR text
    console.log(`Calling Gemini API for file "${originalName}" with MIME type: ${mimeType}`);
    const extractedInfo = callGeminiAPIInternal(blob, originalName, base64Content, ocrText);

    if (extractedInfo.error) {
      console.warn(`AI extraction for "${originalName}" returned an error: ${extractedInfo.error}`);
      return { ...extractedInfo, previewData: previewData };
    }

    // Enhanced validation of extracted data
    const validatedInfo = validateAndCleanExtractedData(extractedInfo);
    console.log("AI Extraction successful. Validated result:", validatedInfo);
    return { ...validatedInfo, previewData: previewData };

  } catch (e) {
    const errorMessage = `Error processing file ${fileId} with AI: ${e.message}`;
    console.error(errorMessage);
    return { 
      error: `AI Processing failed for "${originalName}": ${e.message}`, 
      confidence: "Error", 
      previewData: previewData 
    };
  }
}


/**
* Validates and cleans the data extracted by AI
* @param {Object} extractedData - Raw data from AI
* @returns {Object} - Validated and cleaned data
*/
function validateAndCleanExtractedData(extractedData) {
 const cleaned = { ...extractedData };


 // Clean and validate date
 if (cleaned.date && cleaned.date !== "N/A") {
   cleaned.date = validateAndFormatDate(cleaned.date);
 }


 // Clean vendor name (remove special characters that would break filename)
 if (cleaned.vendorName && cleaned.vendorName !== "N/A") {
   cleaned.vendorName = cleaned.vendorName
     .replace(/[_]/g, '-') // Replace underscores with hyphens
     .replace(/[/\\:*?"<>|]/g, '') // Remove invalid filename characters
     .trim();
 }


 // Clean invoice number
 if (cleaned.invoiceNumber && cleaned.invoiceNumber !== "N/A") {
   cleaned.invoiceNumber = cleaned.invoiceNumber
     .replace(/[_]/g, '-') // Replace underscores with hyphens
     .replace(/[/\\:*?"<>|]/g, '') // Remove invalid filename characters
     .trim();
 }


 // Clean amount
 if (cleaned.amount && cleaned.amount !== "N/A") {
   cleaned.amount = cleaned.amount
     .replace(/[_]/g, '') // Remove underscores
     .replace(/[/\\:*?"<>|]/g, '') // Remove invalid filename characters
     .trim();
 }


 return cleaned;
}


/**
* Validates and formats a date string to YYYY-MM-DD format
* @param {string} dateStr - Date string in various formats
* @returns {string} - Formatted date string or original if invalid
*/
function validateAndFormatDate(dateStr) {
 try {
   // Handle common date formats
   let date;


   // Try parsing different formats
   if (dateStr.match(/^\d{4}-\d{2}-\d{2}$/)) {
     // Already in YYYY-MM-DD format
     date = new Date(dateStr);
   } else if (dateStr.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
     // MM/DD/YYYY format
     const [month, day, year] = dateStr.split('/');
     date = new Date(year, month - 1, day);
   } else if (dateStr.match(/^\d{2}-\d{2}-\d{4}$/)) {
     // MM-DD-YYYY format
     const [month, day, year] = dateStr.split('-');
     date = new Date(year, month - 1, day);
   } else if (dateStr.match(/^\d{4}\/\d{2}\/\d{2}$/)) {
     // YYYY/MM/DD format
     const [yearStr, monthStr, dayStr] = dateStr.split('/');
     date = new Date(parseInt(yearStr), parseInt(monthStr) - 1, parseInt(dayStr));
   } else {
     // Try generic Date parsing
     date = new Date(dateStr);
   }


   // Check if date is valid
   if (isNaN(date.getTime())) {
     console.warn(`Invalid date format for conversion: ${dateStr}`);
     return dateStr; // Return original if can't parse
   }


   // Format to YYYY-MM-DD
   const year = date.getFullYear();
   const month = String(date.getMonth() + 1).padStart(2, '0');
   const day = String(date.getDate()).padStart(2, '0');


   // Double-check if the formatted date is logically consistent with original year (e.g., prevents 2024-02-30)
   // This is a basic sanity check, Date object handles most invalid dates by rolling over.
   const recheckedDate = new Date(year, month - 1, day);
   if (recheckedDate.getFullYear() !== year || recheckedDate.getMonth() !== month - 1 || recheckedDate.getDate() !== parseInt(day)) {
       console.warn(`Formatted date ${year}-${month}-${day} is not a valid calendar date for original: ${dateStr}`);
       return dateStr; // Fallback if formatted date is somehow invalid
   }




   return `${year}-${month}-${day}`;
 } catch (e) {
   console.warn(`Error formatting date ${dateStr}: ${e.message}`);
   return dateStr; // Return original if error
 }
}


/**
* Enhanced function to generate a new filename based on extracted data
* @param {Object} extractedData - Data extracted from AI
* @param {string} originalFileName - Original filename to preserve extension
* @returns {string} - New filename following the convention
*/
function generateNewFilename(extractedData, originalFileName) {
 // Extract file extension
 const lastDotIndex = originalFileName.lastIndexOf('.');
 const extension = lastDotIndex > 0 ? originalFileName.substring(lastDotIndex) : '';


 // Use extracted data or fallback values, ensure values are strings and cleaned
 const date = String(extractedData.date || "YYYY-MM-DD");
 const vendor = String(extractedData.vendorName || "UnknownVendor");
 const invoice = String(extractedData.invoiceNumber || "INV-Unknown");
 const amount = String(extractedData.amount || "0.00");


 // Sanitize individual parts to prevent issues before forming the final name
 const sanitizedDate = validateAndFormatDate(date); // Re-validate and format date one last time
 const sanitizedVendor = vendor
   .replace(/[_]/g, '-') // Replace underscores with hyphens
   .replace(/[/\\:*?"<>|]/g, '') // Remove invalid filename characters
   .replace(/\s+/g, ' ') // Normalize multiple spaces to a single space
   .trim() || "UnknownVendor"; // Ensure not empty after cleaning


 const sanitizedInvoice = invoice
   .replace(/[_]/g, '-') // Replace underscores with hyphens
   .replace(/[/\\:*?"<>|]/g, '') // Remove invalid filename characters
   .replace(/\s+/g, '') // Remove spaces for invoice numbers
   .trim() || "INV-Unknown"; // Ensure not empty after cleaning


 const sanitizedAmount = amount
   .replace(/[_]/g, '') // Remove underscores
   .replace(/[/\\:*?"<>|]/g, '') // Remove invalid filename characters
   .trim() || "0.00"; // Ensure not empty after cleaning


 // Create new filename
 let newFileName = `${sanitizedDate}_${sanitizedVendor}_${sanitizedInvoice}_${sanitizedAmount}${extension}`;


 // Final validation on the generated filename (important for double-check)
 const validation = validateFilenameConvention(newFileName);
 if (!validation.isValid) {
   console.error(`ERROR: Generated filename "${newFileName}" doesn't pass final validation: ${validation.reason}. Details: ${validation.details}`);
   // Fallback to a safe, though less informative, name if generated one is bad
   return `ERROR_GENERATED_NAME_${Date.now()}${extension}`;
 }


 return newFileName;
}


/**
* Sanitizes a filename to ensure it's valid for file systems.
* Removes invalid characters and normalizes whitespace.
* This is a general sanitization, `generateNewFilename` does more specific cleaning for each part.
* @param {string} filename - Filename to sanitize
* @returns {string} - Sanitized filename
*/
function sanitizeFilename(filename) {
 return filename
   .replace(/[<>:"/\\|?*]/g, '') // Remove invalid characters for Windows/Unix filenames
   .replace(/\s+/g, ' ') // Normalize multiple spaces to a single space
   .trim(); // Trim leading/trailing whitespace
}


/**
* Internal function to call the Gemini Multimodal API.
* Uses Google AI Studio API Key and Gemini 2.0 Flash model.
* Enhanced to detect invoice status (inflow/outflow).
* @param {GoogleAppsScript.Base.Blob} fileBlob The blob of the file content.
* @param {string} fileName The name of the file.
* @param {string|null} preEncodedBase64 Optional: Pre-encoded base64 string of the file.
* @returns {Object} An object containing extracted data or an error message.
*/
function callGeminiAPIInternal(fileBlob, fileName, preEncodedBase64, ocrText) {
  // Validate blob parameter first
  if (!fileBlob) {
    console.error(`No blob received for "${fileName}". Check that the ID is a real file and that the script has Drive access.`);
    return { error: 'File blob is empty or inaccessible. Please ensure the file exists and the script has proper Drive permissions.', confidence: "Error" };
  }

  // Validate blob has required methods
  if (typeof fileBlob.getContentType !== 'function') {
    console.error(`Invalid blob object for "${fileName}". Expected Blob but got: ${typeof fileBlob}`);
    return { error: 'Invalid file blob object. Please check file permissions and try again.', confidence: "Error" };
  }

  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    console.error("Gemini API Key not found in User Properties. Please set 'GEMINI_API_KEY' in Project Settings -> Script Properties.");
    return { error: "API Key not configured. Please set GEMINI_API_KEY in Google Apps Script Properties.", confidence: "Error" };
  }

  const MOCK_AI_PROCESSING = false; // Set to true for testing without API calls

  if (MOCK_AI_PROCESSING) {
    console.warn("Using MOCK AI Processing for file: %s", fileName);
    Utilities.sleep(1500);
    const randomSuffix = Math.floor(Math.random() * 100);
    return {
      date: `2024-${String(Math.floor(Math.random() * 12) + 1).padStart(2, '0')}-${String(Math.floor(Math.random() * 28) + 1).padStart(2, '0')}`,
      invoiceNumber: `INV-MOCK-${randomSuffix}`,
      vendorName: `Mock Vendor ${String.fromCharCode(65 + randomSuffix % 26)}`,
      amount: `${(Math.random() * 500 + 50).toFixed(2)}`,
      invoiceStatus: ["inflow", "outflow", "unknown"][randomSuffix % 3],
      confidence: ["High", "Medium", "Low"][randomSuffix % 3] + " (Mocked)"
    };
  }

  // --- ACTUAL GEMINI API CALL LOGIC ---
  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL_ID_AIS}:generateContent?key=${apiKey}`;

  try {
    const mimeType = fileBlob.getContentType();
    console.log(`Processing file "${fileName}" with MIME type: ${mimeType}`);
    
    const fileBytesBase64 = preEncodedBase64 || Utilities.base64Encode(fileBlob.getBytes());

    // Add OCR text to the prompt if available
    let promptText = `
You are a document analysis AI system specialized in processing invoices and billing documents.
Analyze the attached file (PDF or image) and extract the following key data fields accurately and consistently.
`;

    if (ocrText && ocrText.length > 100) { // Only include if meaningful
      promptText += `
Here is the OCR-extracted text from the document (use this as your primary source if possible):

${ocrText}

---
`;
    }

    promptText += `
 INVOICE DATA EXTRACTION INSTRUCTIONS
 
 Analyze the provided document(s) and extract invoice information. Handle both single and multiple invoice scenarios carefully.
 
 CRITICAL: First determine if this document contains multiple invoices, then adjust your extraction strategy accordingly.
 
 Return your output as a valid JSON object with exactly the following structure:
 
 For SINGLE invoice files:
 {
   "date": "YYYY-MM-DD",
   "invoiceNumber": "string",
   "vendorName": "string", 
   "amount": "string",
   "invoiceStatus": "inflow | outflow | unknown",
   "numberofinvoices": "1"
 }
 
 For MULTIPLE invoice files:
 {
   "date": "COMBINED_DATES_OR_RANGE",
   "invoiceNumber": "MULTIPLE_INVOICES",
   "vendorName": "PRIMARY_OR_MULTIPLE_VENDORS",
   "amount": "TOTAL_OR_RANGE", 
   "invoiceStatus": "COMBINED_STATUS",
   "numberofinvoices": "actual_count_as_string"
 }
 
 STEP 1: INVOICE COUNT DETECTION
 Before extracting any data, determine the number of invoices using these indicators:
 
 Strong Indicators of Multiple Invoices:
 - Multiple unique invoice numbers (e.g., "INV-001", "INV-002", "BILL-123")
 - Repeated complete invoice headers/footers
 - Multiple "Invoice Date" or "Bill Date" entries with different values
 - Multiple "Total Due" or "Amount Payable" sections
 - Multiple "Bill To" or "Invoice To" blocks with different recipients
 - Page breaks followed by new invoice structures
 - Different vendor letterheads or logos appearing multiple times
 - Sequential invoice layouts (common in batch processing)
 
 Weak Indicators (verify carefully):
 - Multiple line items (could be one detailed invoice)
 - Multiple dates (could be order date, invoice date, due date for same invoice)
 - Long documents (could be one complex invoice with attachments)
 
 STEP 2: EXTRACTION RULES BY SCENARIO
 
 === FOR SINGLE INVOICE (numberofinvoices = "1") ===
 
 1. "date": 
    - Extract the primary invoice/bill date
    - Priority order: "Invoice Date" > "Bill Date" > "Date Issued" > "Created Date"
    - Format as YYYY-MM-DD
    - If multiple dates exist, choose the one closest to the main invoice header
 
 2. "invoiceNumber":
    - Look for: "Invoice #", "Invoice No", "Bill #", "Reference #", "Document #"
    - Clean: Keep alphanumerics, hyphens, periods, ampersands, spaces only
    - Remove: underscores, slashes, excessive whitespace
 
 3. "vendorName":
    - Extract the invoice issuer (company sending the bill)
    - Look in: document header, "From" section, company logo area, "Issued by"
    - Clean: Remove special chars except alphanumerics, periods, hyphens, ampersands, spaces
    - Prioritize official company name over individual names
 
 4. "amount":
    - Extract final payable amount
    - Priority: "Total Due" > "Amount Payable" > "Grand Total" > "Balance Due" > "Total"
    - Include currency symbol, format as string (e.g., "₹1025.50")
    - Remove thousands separators (use "₹1025.50" not "₹1,025.50")
 
 5. "invoiceStatus":
    - "outflow": You are paying (look for "Bill To: [Your Company]", "Invoice To: [You]")
    - "inflow": You are receiving payment (look for "Bill To: [Other Company]", you are the vendor)
    - "unknown": Cannot determine clearly
 
 === FOR MULTIPLE INVOICES (numberofinvoices > "1") ===
 
 1. "date":
    - If all invoices have same date: use that date
    - If dates span a range: use "YYYY-MM-DD to YYYY-MM-DD" format
    - If dates are scattered: use "MULTIPLE_DATES"
 
 2. "invoiceNumber":
    - Always set to "MULTIPLE_INVOICES"
    - Do not attempt to list all numbers
 
 3. "vendorName":
    - If same vendor for all: use that vendor name
    - If multiple vendors: use "MULTIPLE_VENDORS"
    - If one primary vendor with subsidiaries: use primary vendor name
 
 4. "amount":
    - If clear total across all invoices: calculate and include currency
    - If unclear or mixed currencies: use "MULTIPLE_AMOUNTS"
    - Format: "₹[total]" or "MULTIPLE_AMOUNTS"
 
 5. "invoiceStatus":
    - If all invoices same direction: use "inflow" or "outflow"
    - If mixed directions: use "mixed"
    - If unclear: use "unknown"
 
 6. "numberofinvoices":
    - Count distinct invoice documents
    - Must be string representation of integer
    - Only count invoices with at least invoice number OR amount OR vendor name
 
 STEP 3: VALIDATION AND ERROR HANDLING
 
 - Do NOT fabricate missing information
 - Use "N/A" for missing fields (except invoiceStatus and numberofinvoices)
 - Use "unknown" for unclear invoiceStatus
 - Ensure valid JSON (no trailing commas, proper quotes)
 - Double-check invoice count before finalizing
 
 STEP 4: COMMON MULTI-INVOICE SCENARIOS
 
 Scenario A: Batch Invoice Processing
 - Multiple invoices from same vendor to different customers
 - Extract: vendor name, set invoiceNumber to "MULTIPLE_INVOICES", count accurately
 
 Scenario B: Statement with Multiple Bills
 - One document containing several billing periods
 - Extract: primary vendor, date range, total if available
 
 Scenario C: Consolidated Invoice Pack
 - Different vendors, different time periods
 - Extract: "MULTIPLE_VENDORS", date range, "MULTIPLE_AMOUNTS"
 
 EXAMPLES:
 
 Single Invoice:
 {
   "date": "2024-11-01",
   "invoiceNumber": "INV-102938", 
   "vendorName": "Acme Technologies Pvt. Ltd.",
   "amount": "₹1750.00",
   "invoiceStatus": "outflow",
   "numberofinvoices": "1"
 }
 
 Multiple Invoices (Same Vendor):
 {
   "date": "2024-11-01 to 2024-11-15",
   "invoiceNumber": "MULTIPLE_INVOICES",
   "vendorName": "Acme Technologies Pvt. Ltd.", 
   "amount": "₹5250.00",
   "invoiceStatus": "outflow",
   "numberofinvoices": "3"
 }
 
 Multiple Invoices (Different Vendors):
 {
   "date": "MULTIPLE_DATES",
   "invoiceNumber": "MULTIPLE_INVOICES",
   "vendorName": "MULTIPLE_VENDORS",
   "amount": "MULTIPLE_AMOUNTS", 
   "invoiceStatus": "mixed",
   "numberofinvoices": "4"
 }
 
 Now analyze the provided document and return the appropriate JSON response.
 `;

    // Check supported file types for Gemini Vision model
    if (!mimeType.startsWith("image/") && mimeType !== "application/pdf") {
      console.error(`Unsupported MIME type for Gemini API: ${mimeType} for file ${fileName}.`);
      return { error: `File type ${mimeType} is not supported by the AI model for analysis.`, confidence: "Error" };
    }

    const requestBody = {
      "contents": [
        {
          "parts": [
            { "text": promptText },
            {
              "inline_data": {
                "mime_type": mimeType,
                "data": fileBytesBase64
              }
            }
          ]
        }
      ],
      "generationConfig": {
        "temperature": 0.1,
        "maxOutputTokens": 2048,
        "responseMimeType": "application/json"
      }
    };

    const options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(requestBody),
      'muteHttpExceptions': true,
      'followRedirects': true
    };

    console.log(`Sending request to Gemini API for file: "${fileName}" (MIME: ${mimeType}).`);

    try {
      const response = UrlFetchApp.fetch(endpoint, options);
      const responseCode = response.getResponseCode();
      const responseBody = response.getContentText();

      console.log(`Gemini API Response Code: ${responseCode}`);

      if (responseCode === 200) {
        console.log("Gemini API Raw Response (truncated):", responseBody.substring(0, 500) + "...");

        let jsonResponse;
        try {
          jsonResponse = JSON.parse(responseBody);
        } catch (e) {
          throw new Error(`Failed to parse AI response as JSON. Raw response: ${responseBody.substring(0, 200)}... Error: ${e.message}`);
        }

        let parsedData = {};
        if (jsonResponse.candidates && jsonResponse.candidates[0] &&
            jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts &&
            jsonResponse.candidates[0].content.parts[0]) {

          const part = jsonResponse.candidates[0].content.parts[0];
          if (part.text) {
            try {
              let cleanText = part.text.trim();
              // Remove markdown code block fences if present
              if (cleanText.startsWith('```json')) {
                cleanText = cleanText.replace(/^```json\s*/, '').replace(/\s*```$/, '');
              } else if (cleanText.startsWith('```')) {
                cleanText = cleanText.replace(/^```\s*/, '').replace(/\s*```$/, '');
              }
              parsedData = JSON.parse(cleanText);
            } catch (e) {
              throw new Error(`AI returned text that was not valid JSON. Text: "${part.text.substring(0, 200)}...". Error: ${e.message}`);
            }
          } else if (Object.keys(part).length > 0) {
            // If the part itself is a JSON object (e.g., if responseMimeType is set)
            parsedData = part;
          } else {
            throw new Error("AI response candidate content structure is not recognized or is empty.");
          }

        } else {
          throw new Error("Valid candidate with parsable content not found in AI response.");
        }

        // Ensure all expected keys are present, defaulting to "N/A" or "unknown"
        return {
          date: parsedData.date || "N/A",
          invoiceNumber: parsedData.invoiceNumber || "N/A",
          vendorName: parsedData.vendorName || "N/A",
          amount: parsedData.amount || "N/A",
          invoiceStatus: parsedData.invoiceStatus || "unknown",
          confidence: "High (Gemini 2.0 Flash)",
          numberofinvoices : parsedData.numberofinvoices
        };

      } else {
        let errorMessage = `AI API request failed with status ${responseCode}.`;
        try {
          const errorJson = JSON.parse(responseBody);
          if (errorJson.error && errorJson.error.message) {
            errorMessage += ` Message: ${errorJson.error.message}`;
          }
        } catch (e) {
          errorMessage += ` Raw response: ${responseBody.substring(0, 200)}...`;
        }
        console.error(`Gemini API Error: ${errorMessage}`);
        return { error: errorMessage, confidence: "Error" };
      }

    } catch (apiError) {
      console.error("Critical error during Gemini API call: %s", apiError.message);
      return { error: `An unexpected error occurred during AI processing: ${apiError.message}`, confidence: "Error" };
    }

  } catch (error) {
    console.error(`Error in callGeminiAPIInternal for ${fileName}:`, error);
    return { 
      error: `AI processing failed: ${error.message}`, 
      confidence: "Error" 
    };
  }
}


/**
* Enhanced function to rename a Google Drive file and update the corresponding cell in the spreadsheet.
* @param {string} fileId The ID of the file to rename.
* @param {string} newName The new desired name for the file.
* @param {string} sheetName The name of the sheet where the file is listed.
* @param {number} row The row number in the sheet where the file entry is located.
* @param {number} column The column number in the sheet to update (typically 1 for Column A).
* @param {string} invoiceStatus The status of the invoice.
* @param {string} vendorName The name of the vendor.
* @returns {Object} A result object with success status and message.
*/
function renameFileAndUpdateSheetGAS(fileId, newName, sheetName, row, column, invoiceStatus, vendorName) {
 console.log(`Attempting to rename file "${fileId}" to "${newName}" and update sheet with invoice status "${invoiceStatus}".`);
 try {
   if (!fileId || !newName || !sheetName || !row || !column) {
     throw new Error("Missing one or more required parameters for renaming operation.");
   }


   // Validate the new filename before proceeding. This is crucial now.
   const validation = validateFilenameConvention(newName);
   if (!validation.isValid) {
     // If the generated/user-edited name is invalid, we should ideally not proceed
     // or at least warn significantly. For this workflow, I'll return an error.
     const errorMessage = `Attempted rename with an invalid new filename: "${newName}". Reason: ${validation.reason}. Details: ${validation.details}`;
     console.error(errorMessage);
     return {
       success: false,
       message: errorMessage,
       error: `Generated filename does not conform to convention.`
     };
   }


   // Rename the file in Google Drive
   const file = DriveApp.getFileById(fileId);
   const oldName = file.getName();
   file.setName(newName);
   console.log(`File ID "${fileId}" successfully renamed from "${oldName}" to "${newName}" in Drive.`);


   // Update the sheet with the new filename and invoice status
   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
   if (!sheet) {
     throw new Error(`Sheet "${sheetName}" not found for updating.`);
   }


   // Update filename in the specified column (typically Column A)
   sheet.getRange(row, column).setValue(newName);
   console.log(`Sheet "${sheetName}", Cell[${row},${column}] successfully updated with "${newName}".`);


   // Update invoice status if provided and valid
   if (invoiceStatus && (invoiceStatus === "inflow" || invoiceStatus === "outflow")) {
     const invoiceStatusColumn = findInvoiceStatusColumn(sheet);
     if (invoiceStatusColumn > 0) {
       sheet.getRange(row, invoiceStatusColumn).setValue(invoiceStatus);
       console.log(`Sheet "${sheetName}", Invoice Status Column ${invoiceStatusColumn}, Row ${row} updated with "${invoiceStatus}".`);
     } else {
       console.warn(`Invoice Status column not found in sheet "${sheetName}". Skipping status update.`);
     }
   }


   return {
     success: true,
     message: `Successfully renamed to "${newName}" and updated sheet.`,
     oldName: oldName,
     newName: newName,
     invoiceStatusUpdated: invoiceStatus && (invoiceStatus === "inflow" || invoiceStatus === "outflow")
   };
 } catch (e) {
   const errorMessage = `Error renaming file or updating sheet for file ID "${fileId}": ${e.message}`;
   console.error(errorMessage);
   return {
     success: false,
     message: errorMessage,
     error: e.message
   };
 }
}

/*
* Helper function to find the "Invoice Status" column in a sheet.
* @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to search in.
* @returns {number} The column number (1-based) of the Invoice Status column, or 0 if not found.
*/
function findInvoiceStatusColumn(sheet) {
 try {
   const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];


   for (let i = 0; i < headerRow.length; i++) {
     const headerValue = String(headerRow[i]).toLowerCase().trim();
     if (headerValue.includes('invoice status') || headerValue.includes('invoicestatus')) {
       return i + 1; // Return 1-based column number
     }
   }


   console.warn("Invoice Status column not found in sheet headers:", headerRow);
   return 0; // Not found
 } catch (e) {
   console.error("Error finding Invoice Status column:", e.message);
   return 0;
 }
}

/**
 * Extracts text from an image or PDF file using Google Drive OCR.
 * @param {string} fileId - The ID of the file to OCR.
 * @returns {string} - The extracted text, or an error message.
 */
function extractTextWithOCR(fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var resource = {
      title: file.getName(),
      mimeType: MimeType.GOOGLE_DOCS
    };
    var ocrFile = Drive.Files.insert(resource, blob, {
      ocr: true,
      ocrLanguage: 'en'
    });
    var doc = DocumentApp.openById(ocrFile.id);
    var text = doc.getBody().getText();
    DriveApp.getFileById(ocrFile.id).setTrashed(true); // Clean up
    return text;
  } catch (e) {
    return '';
  }
}

/**
* Debug function to check spreadsheet for file ID issues
* @param {string} sheetName The sheet name to check
* @returns {Object} Debug information about the spreadsheet
*/
function debugSpreadsheetFileIds(sheetName) {
  console.log(`Debugging file IDs in sheet: "${sheetName}"`);
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      return { error: `Sheet "${sheetName}" not found` };
    }
    
    const lastRow = sheet.getLastRow();
    console.log(`Sheet has ${lastRow} rows`);
    
    if (lastRow < 2) {
      return { error: "Sheet has no data rows (only header or empty)" };
    }
    
    // Get all data from columns A, B, C
    const dataRange = sheet.getRange("A2:C" + lastRow);
    const values = dataRange.getValues();
    
    const results = {
      totalRows: values.length,
      validFileIds: [],
      invalidFileIds: [],
      emptyRows: [],
      errors: []
    };
    
    for (let i = 0; i < values.length; i++) {
      const rowNumber = i + 2;
      const fileName = values[i][0];
      const fileId = values[i][1];
      const fileUrl = values[i][2];
      
      console.log(`Row ${rowNumber}: FileName="${fileName}", FileID="${fileId}", FileURL="${fileUrl}"`);
      
      // Check if file ID is empty or invalid
      if (!fileId || typeof fileId !== 'string' || fileId.trim() === '') {
        results.emptyRows.push({
          row: rowNumber,
          fileName: fileName,
          fileId: fileId,
          fileUrl: fileUrl
        });
        continue;
      }
      
      // Test if file ID is valid
      try {
        const file = DriveApp.getFileById(fileId.trim());
        const actualFileName = file.getName();
        
        results.validFileIds.push({
          row: rowNumber,
          sheetFileName: fileName,
          actualFileName: actualFileName,
          fileId: fileId.trim(),
          fileUrl: fileUrl,
          mimeType: file.getMimeType()
        });
        
        console.log(`✓ Row ${rowNumber}: Valid file "${actualFileName}"`);
        
      } catch (error) {
        results.invalidFileIds.push({
          row: rowNumber,
          fileName: fileName,
          fileId: fileId.trim(),
          fileUrl: fileUrl,
          error: error.message
        });
        
        console.log(`✗ Row ${rowNumber}: Invalid file ID "${fileId}" - ${error.message}`);
      }
    }
    
    console.log(`Debug Results:`);
    console.log(`- Total rows: ${results.totalRows}`);
    console.log(`- Valid file IDs: ${results.validFileIds.length}`);
    console.log(`- Invalid file IDs: ${results.invalidFileIds.length}`);
    console.log(`- Empty rows: ${results.emptyRows.length}`);
    
    return results;
    
  } catch (error) {
    console.error(`Error debugging spreadsheet: ${error.message}`);
    return { error: `Debug failed: ${error.message}` };
  }
}

/**
* Test function with a sample file ID for demonstration
* Replace the fileId below with a real file ID from your spreadsheet
* @returns {Object} Test results
*/
function testWithSampleFile() {
  // Replace this with a real file ID from your spreadsheet (Column B)
  const sampleFileId = "1ABC123DEF456"; // ← REPLACE WITH REAL FILE ID
  
  console.log("Testing with sample file ID:", sampleFileId);
  return testFileAccess(sampleFileId);
}

/**
* Test function to debug file access issues
* @param {string} fileId The file ID to test
* @returns {Object} Test results
*/
function testFileAccess(fileId) {
  console.log(`Testing file access for ID: ${fileId}`);
  
  try {
    // Test 1: Basic file access
    const file = DriveApp.getFileById(fileId);
    console.log(`✓ File found: "${file.getName()}"`);
    
    // Test 2: Get blob
    const blob = file.getBlob();
    if (!blob) {
      console.log(`✗ getBlob() returned null/undefined`);
      return { success: false, error: 'getBlob() returned null/undefined' };
    }
    
    // Test 3: Check blob methods
    if (typeof blob.getContentType !== 'function') {
      console.log(`✗ blob.getContentType is not a function`);
      return { success: false, error: 'blob.getContentType is not a function' };
    }
    
    if (typeof blob.getBytes !== 'function') {
      console.log(`✗ blob.getBytes is not a function`);
      return { success: false, error: 'blob.getBytes is not a function' };
    }
    
    // Test 4: Get content type
    const mimeType = blob.getContentType();
    console.log(`✓ MIME type: ${mimeType}`);
    
    // Test 5: Get bytes
    const bytes = blob.getBytes();
    console.log(`✓ Bytes length: ${bytes.length}`);
    
    // Test 6: Check if it's a Google Workspace file
    const googleWorkspaceMimeTypes = [
      'application/vnd.google-apps.document',
      'application/vnd.google-apps.spreadsheet', 
      'application/vnd.google-apps.presentation'
    ];
    
    if (googleWorkspaceMimeTypes.includes(mimeType)) {
      console.log(`⚠ File is Google Workspace type: ${mimeType}`);
      try {
        const exportedBlob = Drive.Files.export(fileId, 'application/pdf');
        console.log(`✓ Export successful, exported blob size: ${exportedBlob.getBytes().length}`);
      } catch (exportError) {
        console.log(`✗ Export failed: ${exportError.message}`);
        return { success: false, error: `Export failed: ${exportError.message}` };
      }
    }
    
    console.log(`✓ All tests passed for file: "${file.getName()}"`);
    return { 
      success: true, 
      fileName: file.getName(),
      mimeType: mimeType,
      fileSize: bytes.length
    };
    
  } catch (error) {
    console.log(`✗ Test failed: ${error.message}`);
    return { success: false, error: error.message };
  }
}

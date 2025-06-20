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

  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();

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

    // Call Gemini API
    const extractedInfo = callGeminiAPIInternal(blob, originalName, base64Content);

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
    return { error: `AI Processing failed for "${originalName}": ${e.message}`, confidence: "Error", previewData: previewData };
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
function callGeminiAPIInternal(fileBlob, fileName, preEncodedBase64) {
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

  const mimeType = fileBlob.getContentType();
  const fileBytesBase64 = preEncodedBase64 || Utilities.base64Encode(fileBlob.getBytes());

  const promptText = `
You are a document analysis AI system specialized in processing invoices and billing documents. Analyze the attached file (PDF or image) and extract the following key data fields accurately and consistently:

Return your output as a valid JSON object with exactly the following keys:
{
  "date": "YYYY-MM-DD",
  "invoiceNumber": "string",
  "vendorName": "string",
  "amount": "string",
  "invoiceStatus": "inflow | outflow | unknown",
  "numberofinvoices": "integer"
}

Field-Specific Extraction Rules:

1. "date":
- Extract the main invoice or bill date.
- Look for terms such as "Invoice Date", "Bill Date", "Date Issued", etc.
- Convert all formats to YYYY-MM-DD.
- If multiple dates are present, prioritize the one closest to the invoice title or header.

2. "invoiceNumber":
- Look for labels like "Invoice #", "Bill #", "Reference #", "Ref #", etc.
- Clean the extracted value: retain only alphanumerics, hyphens, periods, ampersands, and spaces. Remove underscores or slashes.

3. "vendorName":
- Extract the name of the seller or issuer of the invoice.
- Typically located at the top of the document or labeled as "From", "Issued by", "Vendor", or part of the company logo/header.
- Clean the result: remove special characters except alphanumerics, periods, hyphens, ampersands, and spaces.

4. "amount":
- Extract the final amount payable or due.
- Prioritize values near labels such as "Total", "Amount Due", "Grand Total", or "Balance Due".
- Include currency symbol (e.g., ₹, $, €), and format as a plain string (e.g., "₹1025.50").
- Avoid commas in numbers (e.g., use "₹1025.50", not "₹1,025.50").

5. "invoiceStatus":
- Determine the nature of the invoice:
  • If phrases like "Invoice To", "Bill To", "Billed To" appear, followed by a recipient name, set as "outflow".
  • If phrases like "Invoice From", "Billed By", "Supplier", or "Issued By" point to someone invoicing you, set as "inflow".
  • If unclear, return "unknown".

6. "numberofinvoices":
- Estimate the number of distinct invoice or billing documents present in the file — especially across multi-page PDFs.

- Use structural and contextual cues to count invoices, including:
  • The appearance of multiple unique invoice numbers, e.g., more than one "Invoice #", "Ref #", or "Bill #" with different values.
  • Repeated structured patterns: headers like "Invoice Date", "Bill To", "Total", etc., occurring multiple times across the document.
  • Page-wise distribution: Detect invoice groupings across consecutive pages. For example, a 14-page PDF where:
    - Pages 1–3 have one invoice block (Invoice 1)
    - Pages 4–5 are blank or contain only separators
    - Pages 6–8 have a new invoice block (Invoice 2)
    - Pages 9–10 are blank or contain only logos
    - Pages 11–14 contain another invoice (Invoice 3)
  should be considered as 3 invoices.

- Use page breaks, whitespace patterns, blank pages, or layout resets as segmentation clues.

- Count an invoice block only if it contains at least a valid invoice number or total amount, and optionally, a date or vendor name.

- Return the count as a stringified integer:
  • If multiple distinct invoice blocks are confidently identified, return their count (e.g., "3").
  • If only one invoice is confidently identified across the document, return "1".
  • If the structure is ambiguous or incomplete, return "0".

Important Notes:
- Do NOT hallucinate or fabricate missing information. If a field is not found or unclear, use:
  - "N/A" for date, invoiceNumber, vendorName, amount
  - "unknown" for invoiceStatus
- Ensure that output is valid JSON (no trailing commas, no missing quotes).

Output Example:
{
  "date": "2024-11-01",
  "invoiceNumber": "INV-102938",
  "vendorName": "Acme Technologies Pvt. Ltd.",
  "amount": "₹1750.00",
  "invoiceStatus": "inflow",
  "numberofinvoices": "1"
}
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
}

/**
 * Enhanced function to rename a Google Drive file and update the corresponding cell in the spreadsheet.
 * @param {string} fileId The ID of the file to rename.
 * @param {string} newName The new desired name for the file.
 * @param {string} sheetName The name of the sheet where the file is listed.
 * @param {number} row The row number in the sheet where the file entry is located.
 * @param {number} column The column number in the sheet to update (typically 1 for Column A).
 * @returns {Object} A result object with success status and message.
 */
function renameFileAndUpdateSheetGAS(fileId, newName, sheetName, row, column, invoiceStatus) {
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
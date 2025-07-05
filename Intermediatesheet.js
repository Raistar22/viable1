/**
 * Returns all file entries from the main sheet (analogy or humane).
 * Each entry includes fileId, fileName, url, date, size, mimeType, gmailId, status, ui.
 */
function getMainSheetFiles(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const files = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    files.push({
      fileId: row[1], // adjust index as per your sheet structure
      fileName: row[0],
      url: row[2],
      date: row[3],
      size: row[5],
      mimeType: row[6],
      gmailId: row[8],
      status: row[9],
      ui: row[10]
    });
  }
  return files;
}

/**
 * Returns the details of a single file by fileId from the given sheet.
 */
function getFileDetailsById(sheetName, fileId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] === fileId) {
      return {
        fileId: row[1],
        fileName: row[0],
        url: row[2],
        date: row[3],
        size: row[5],
        mimeType: row[6],
        gmailId: row[8],
        status: row[9],
        ui: row[10]
      };
    }
  }
  return null;
}

/**
 * Verifies the data for a given fileId and field.
 * Returns true if the data matches, false otherwise.
 */
function verifyFieldData(sheetName, fileId, field) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIndex = headers.findIndex(h => h.toLowerCase().includes(field.toLowerCase()));
  if (colIndex === -1) return false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === fileId) {
      // Add your verification logic here (e.g., check if value is not empty)
      return !!data[i][colIndex];
    }
  }
  return false;
}

function getSubsheetNamesGAS() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  return sheets.map(sheet => sheet.getName());
}
  
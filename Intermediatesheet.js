/**
 * Returns the content of an HTML file for inclusion in other HTML files.
 * Useful for modular UI components.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Fetches structured data from the "analogy" sheet.
 * Returns an object with headers as keys and arrays of values.
 */
function getAnalogySheetData() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName('analogy');

    if (!sheet) {
      throw new Error('Sheet named "analogy" not found');
    }

    const values = sheet.getDataRange().getValues();
    if (values.length === 0) return { error: 'No data found in the analogy sheet' };

    const headers = values[0];
    const organizedData = {};

    for (let col = 0; col < headers.length; col++) {
      const header = headers[col];
      if (!header || header.toString().trim() === '') continue;

      const columnData = [];
      for (let row = 1; row < values.length; row++) {
        const cellValue = values[row][col];
        if (cellValue !== '' && cellValue !== null && cellValue !== undefined) {
          columnData.push(cellValue.toString().trim());
        }
      }

      organizedData[header.toString().trim()] = columnData;
    }

    return {
      success: true,
      data: organizedData,
      sheetName: sheet.getName(),
      totalRows: values.length,
      totalColumns: headers.length
    };
  } catch (error) {
    Logger.log('Error in getAnalogySheetData: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Called by frontend to refresh the sidebar data.
 */
function refreshData() {
  return getAnalogySheetData();
}

/**
 * Fetches basic metadata about the analogy sheet.
 */
function getSheetInfo() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName('analogy');

    if (!sheet) return { error: 'Sheet named "analogy" not found' };

    return {
      sheetName: sheet.getName(),
      lastRow: sheet.getLastRow(),
      lastColumn: sheet.getLastColumn(),
      spreadsheetName: spreadsheet.getName()
    };
  } catch (error) {
    return { error: error.toString() };
  }
}

/**
 * Updates the data in the analogy sheet.
 * Expects a JSON object with headers as keys and arrays of values.
 */
function updateAnalogySheetData(updatedData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('analogy');
    if (!sheet) throw new Error('Sheet named "analogy" not found');

    const currentValues = sheet.getDataRange().getValues();
    const headers = currentValues[0];

    let maxRows = 1;
    for (let key in updatedData) {
      if (Array.isArray(updatedData[key])) {
        maxRows = Math.max(maxRows, updatedData[key].length + 1);
      }
    }

    const newData = [headers];
    for (let row = 1; row < maxRows; row++) {
      const rowData = [];
      for (let col = 0; col < headers.length; col++) {
        const header = headers[col];
        const cellValue = updatedData[header]?.[row - 1] ?? '';
        rowData.push(cellValue);
      }
      newData.push(rowData);
    }

    sheet.clear();
    sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);

    return {
      success: true,
      message: 'Data updated successfully',
      rowsUpdated: newData.length,
      columnsUpdated: headers.length
    };
  } catch (error) {
    Logger.log('Error in updateAnalogySheetData: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Adds a new value to a specific column in the analogy sheet.
 */
function addItemToColumn(header, value) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('analogy');
    if (!sheet) throw new Error('Sheet named "analogy" not found');

    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const headerIndex = headers.indexOf(header);

    if (headerIndex === -1) throw new Error(`Header "${header}" not found`);

    let targetRow = sheet.getLastRow() + 1;
    for (let row = 1; row < values.length; row++) {
      if (!values[row][headerIndex]) {
        targetRow = row + 1;
        break;
      }
    }

    sheet.getRange(targetRow, headerIndex + 1).setValue(value);

    return {
      success: true,
      message: 'Item added successfully',
      addedTo: header,
      value,
      row: targetRow
    };
  } catch (error) {
    Logger.log('Error in addItemToColumn: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Removes a value from a specific column in the analogy sheet.
 */
function removeItemFromColumn(header, value) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('analogy');
    if (!sheet) throw new Error('Sheet named "analogy" not found');

    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const headerIndex = headers.indexOf(header);

    if (headerIndex === -1) throw new Error(`Header "${header}" not found`);

    let removed = false;
    for (let row = 1; row < values.length; row++) {
      if (values[row][headerIndex]?.toString().trim() === value.trim()) {
        sheet.getRange(row + 1, headerIndex + 1).setValue('');
        removed = true;
        break;
      }
    }

    if (!removed) throw new Error(`Value "${value}" not found in column "${header}"`);

    return {
      success: true,
      message: 'Item removed successfully',
      removedFrom: header,
      value
    };
  } catch (error) {
    Logger.log('Error in removeItemFromColumn: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

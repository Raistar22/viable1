/**
 * Opens the sidebar when the spreadsheet is opened
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Analogy Data')
    .addItem('Show Sidebar', 'showSidebar')
    .addToUi();
}

/**
 * Shows the sidebar
 */
function showSidebar() {
  var html = HtmlService.createTemplateFromFile('intermediatesheet');
  var htmlOutput = html.evaluate()
    .setTitle('Analogy Sheet Data')
    .setWidth(400);
  
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

/**
 * Include external HTML files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Fetches data from the analogy sheet and organizes it with headers as keys
 * @return {Object} Object with headers as keys and corresponding column data as values
 */
function getAnalogySheetData() {
  try {
    // Get the active spreadsheet
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get the "analogy" sheet
    var sheet = spreadsheet.getSheetByName('analogy');
    
    if (!sheet) {
      throw new Error('Sheet named "analogy" not found');
    }
    
    // Get all data from the sheet
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    
    if (values.length === 0) {
      return { error: 'No data found in the analogy sheet' };
    }
    
    // Get headers from the first row
    var headers = values[0];
    
    // Create an object to store the organized data
    var organizedData = {};
    
    // Process each column
    for (var col = 0; col < headers.length; col++) {
      var header = headers[col];
      
      // Skip empty headers
      if (!header || header.toString().trim() === '') {
        continue;
      }
      
      // Get all values in this column (excluding the header)
      var columnData = [];
      for (var row = 1; row < values.length; row++) {
        var cellValue = values[row][col];
        
        // Only add non-empty values
        if (cellValue !== null && cellValue !== undefined && cellValue.toString().trim() !== '') {
          columnData.push(cellValue.toString().trim());
        }
      }
      
      // Add to organized data
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
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Refreshes the data in the sidebar
 */
function refreshData() {
  return getAnalogySheetData();
}

/**
 * Gets basic sheet information
 */
function getSheetInfo() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName('analogy');
    
    if (!sheet) {
      return { error: 'Sheet named "analogy" not found' };
    }
    
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
 * Updates data in the analogy sheet
 * @param {Object} updatedData - Object with headers as keys and arrays of values
 * @return {Object} Success/error response
 */
function updateAnalogySheetData(updatedData) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName('analogy');
    
    if (!sheet) {
      throw new Error('Sheet named "analogy" not found');
    }
    
    // Get current data to understand the structure
    var dataRange = sheet.getDataRange();
    var currentValues = dataRange.getValues();
    
    if (currentValues.length === 0) {
      throw new Error('No data found in the analogy sheet');
    }
    
    var headers = currentValues[0];
    var maxRows = 1; // Start with 1 for headers
    
    // Find the maximum number of rows needed
    for (var header in updatedData) {
      if (updatedData[header] && Array.isArray(updatedData[header])) {
        maxRows = Math.max(maxRows, updatedData[header].length + 1);
      }
    }
    
    // Create new data array
    var newData = [];
    
    // Add headers as first row
    newData.push(headers);
    
    // Fill in the data
    for (var row = 1; row < maxRows; row++) {
      var rowData = [];
      
      for (var col = 0; col < headers.length; col++) {
        var header = headers[col];
        var headerData = updatedData[header];
        
        if (headerData && Array.isArray(headerData) && headerData[row - 1] !== undefined) {
          rowData.push(headerData[row - 1]);
        } else {
          rowData.push(''); // Empty cell
        }
      }
      
      newData.push(rowData);
    }
    
    // Clear the existing data
    sheet.clear();
    
    // Write the new data
    if (newData.length > 0 && newData[0].length > 0) {
      sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
    }
    
    return {
      success: true,
      message: 'Data updated successfully',
      rowsUpdated: newData.length,
      columnsUpdated: headers.length
    };
    
  } catch (error) {
    Logger.log('Error in updateAnalogySheetData: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Adds a new item to a specific column
 * @param {string} header - The column header
 * @param {string} value - The value to add
 * @return {Object} Success/error response
 */
function addItemToColumn(header, value) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName('analogy');
    
    if (!sheet) {
      throw new Error('Sheet named "analogy" not found');
    }
    
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    
    if (values.length === 0) {
      throw new Error('No data found in the analogy sheet');
    }
    
    var headers = values[0];
    var headerIndex = headers.indexOf(header);
    
    if (headerIndex === -1) {
      throw new Error('Header "' + header + '" not found');
    }
    
    // Find the first empty cell in this column
    var targetRow = sheet.getLastRow() + 1;
    
    // Check if there are empty cells in between
    for (var row = 1; row < values.length; row++) {
      if (!values[row][headerIndex] || values[row][headerIndex].toString().trim() === '') {
        targetRow = row + 1;
        break;
      }
    }
    
    // Add the new value
    sheet.getRange(targetRow, headerIndex + 1).setValue(value);
    
    return {
      success: true,
      message: 'Item added successfully',
      addedTo: header,
      value: value,
      row: targetRow
    };
    
  } catch (error) {
    Logger.log('Error in addItemToColumn: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Removes an item from a specific column
 * @param {string} header - The column header
 * @param {string} value - The value to remove
 * @return {Object} Success/error response
 */
function removeItemFromColumn(header, value) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName('analogy');
    
    if (!sheet) {
      throw new Error('Sheet named "analogy" not found');
    }
    
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    
    if (values.length === 0) {
      throw new Error('No data found in the analogy sheet');
    }
    
    var headers = values[0];
    var headerIndex = headers.indexOf(header);
    
    if (headerIndex === -1) {
      throw new Error('Header "' + header + '" not found');
    }
    
    // Find and remove the value
    var removed = false;
    for (var row = 1; row < values.length; row++) {
      if (values[row][headerIndex] && values[row][headerIndex].toString().trim() === value.toString().trim()) {
        sheet.getRange(row + 1, headerIndex + 1).setValue('');
        removed = true;
        break;
      }
    }
    
    if (!removed) {
      throw new Error('Value "' + value + '" not found in column "' + header + '"');
    }
    
    return {
      success: true,
      message: 'Item removed successfully',
      removedFrom: header,
      value: value
    };
    
  } catch (error) {
    Logger.log('Error in removeItemFromColumn: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}
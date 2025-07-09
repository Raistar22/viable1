/**
 * buffer1sidebar.gs
 * Google Apps Script functions for manual buffer entry count management
 */

/**
 * Updates the entry count for a specific buffer sheet by manually setting the count
 * @param {string} sheetName - Name of the buffer sheet (analogy-buffer or humane-buffer)
 * @param {number} entryCount - The number of entries to set
 * @returns {Object} Result object with success status
 */
function updateBufferEntryCount(sheetName, entryCount) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = spreadsheet.getSheetByName(sheetName);
      
      // Create the sheet if it doesn't exist
      if (!sheet) {
        sheet = spreadsheet.insertSheet(sheetName);
        console.log(`Created new sheet: ${sheetName}`);
        
        // Add header row for new sheet
        const headers = ['Entry ID', 'Content', 'Type', 'Date Added', 'Status'];
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        
        // Format header row
        const headerRange = sheet.getRange(1, 1, 1, headers.length);
        headerRange.setBackground('#4a90e2');
        headerRange.setFontColor('white');
        headerRange.setFontWeight('bold');
      }
      
      // Get current data to preserve existing entries
      const currentData = sheet.getDataRange().getValues();
      const headerRow = currentData.length > 0 ? currentData[0] : ['Entry ID', 'Content', 'Type', 'Date Added', 'Status'];
      
      // Clear the sheet
      sheet.clear();
      
      // Add header row back
      sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
      
      // Format header row
      const headerRange = sheet.getRange(1, 1, 1, headerRow.length);
      headerRange.setBackground('#4a90e2');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
      
      // Add placeholder entries to match the specified count
      if (entryCount > 0) {
        const placeholderData = [];
        const currentDate = new Date().toISOString().split('T')[0];
        
        for (let i = 1; i <= entryCount; i++) {
          placeholderData.push([
            i,
            `Entry ${i} - ${sheetName}`,
            sheetName.includes('analogy') ? 'Analogy' : 'Humane',
            currentDate,
            'Active'
          ]);
        }
        
        // Add the placeholder data
        if (placeholderData.length > 0) {
          sheet.getRange(2, 1, placeholderData.length, placeholderData[0].length).setValues(placeholderData);
        }
      }
      
      // Auto-resize columns
      sheet.autoResizeColumns(1, headerRow.length);
      
      console.log(`Updated ${sheetName} with ${entryCount} entries`);
      
      return {
        success: true,
        message: `${sheetName} updated with ${entryCount} entries`,
        count: entryCount,
        sheetName: sheetName
      };
      
    } catch (error) {
      console.error('Error updating buffer entry count:', error);
      throw new Error(`Failed to update ${sheetName}: ${error.message}`);
    }
  }
  
  /**
   * Gets the current number of entries (excluding header) in the specified buffer sheet
   * @param {string} sheetName - Name of the buffer sheet to count entries from
   * @returns {number} Number of entries in the sheet
   */
  function getEntryCount(sheetName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = spreadsheet.getSheetByName(sheetName);
      
      if (!sheet) {
        console.log(`Sheet ${sheetName} not found, returning 0`);
        return 0;
      }
      
      const lastRow = sheet.getLastRow();
      
      // Return 0 if no data, or subtract 1 for header row
      const entryCount = lastRow > 0 ? Math.max(0, lastRow - 1) : 0;
      
      console.log(`Sheet ${sheetName} has ${entryCount} entries`);
      return entryCount;
      
    } catch (error) {
      console.error('Error getting entry count:', error);
      return 0;
    }
  }
  
  /**
   * Adds a single entry to the specified buffer sheet
   * @param {string} sheetName - Name of the buffer sheet
   * @param {string} content - Content of the entry
   * @param {string} type - Type of entry (optional)
   * @returns {Object} Result object with success status
   */
  function addSingleEntry(sheetName, content, type = null) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = spreadsheet.getSheetByName(sheetName);
      
      // Create the sheet if it doesn't exist
      if (!sheet) {
        sheet = spreadsheet.insertSheet(sheetName);
        
        // Add header row for new sheet
        const headers = ['Entry ID', 'Content', 'Type', 'Date Added', 'Status'];
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        
        // Format header row
        const headerRange = sheet.getRange(1, 1, 1, headers.length);
        headerRange.setBackground('#4a90e2');
        headerRange.setFontColor('white');
        headerRange.setFontWeight('bold');
      }
      
      const lastRow = sheet.getLastRow();
      const nextRow = lastRow + 1;
      const entryId = lastRow; // ID will be the current count + 1
      const currentDate = new Date().toISOString().split('T')[0];
      const entryType = type || (sheetName.includes('analogy') ? 'Analogy' : 'Humane');
      
      // Add the new entry
      const newEntry = [entryId, content, entryType, currentDate, 'Active'];
      sheet.getRange(nextRow, 1, 1, newEntry.length).setValues([newEntry]);
      
      console.log(`Added entry to ${sheetName}: ${content}`);
      
      return {
        success: true,
        message: `Entry added to ${sheetName}`,
        entryId: entryId,
        content: content
      };
      
    } catch (error) {
      console.error('Error adding single entry:', error);
      throw new Error(`Failed to add entry to ${sheetName}: ${error.message}`);
    }
  }
  
  /**
   * Removes the last entry from the specified buffer sheet
   * @param {string} sheetName - Name of the buffer sheet
   * @returns {Object} Result object with success status
   */
  function removeLastEntry(sheetName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = spreadsheet.getSheetByName(sheetName);
      
      if (!sheet) {
        throw new Error(`Sheet ${sheetName} not found`);
      }
      
      const lastRow = sheet.getLastRow();
      
      if (lastRow <= 1) {
        return {
          success: false,
          message: `No entries to remove from ${sheetName}`,
          count: 0
        };
      }
      
      // Delete the last row
      sheet.deleteRow(lastRow);
      
      const newCount = Math.max(0, lastRow - 2); // -2 because we deleted one row and exclude header
      
      console.log(`Removed last entry from ${sheetName}. New count: ${newCount}`);
      
      return {
        success: true,
        message: `Last entry removed from ${sheetName}`,
        count: newCount
      };
      
    } catch (error) {
      console.error('Error removing last entry:', error);
      throw new Error(`Failed to remove entry from ${sheetName}: ${error.message}`);
    }
  }
  
  /**
   * Clears all entries from the specified buffer sheet (keeps header)
   * @param {string} sheetName - Name of the buffer sheet to clear
   * @returns {Object} Result object with success status
   */
  function clearBufferSheet(sheetName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = spreadsheet.getSheetByName(sheetName);
      
      if (!sheet) {
        throw new Error(`Sheet ${sheetName} not found`);
      }
      
      const lastRow = sheet.getLastRow();
      
      if (lastRow > 1) {
        // Delete all rows except header
        sheet.deleteRows(2, lastRow - 1);
      }
      
      console.log(`Cleared all entries from ${sheetName}`);
      
      return {
        success: true,
        message: `All entries cleared from ${sheetName}`,
        count: 0
      };
      
    } catch (error) {
      console.error('Error clearing buffer sheet:', error);
      throw new Error(`Failed to clear ${sheetName}: ${error.message}`);
    }
  }
  
  /**
   * Gets detailed information about both buffer sheets
   * @returns {Object} Information about both analogy-buffer and humane-buffer sheets
   */
  function getBufferSheetsInfo() {
    try {
      const analogyCount = getEntryCount('analogy-buffer');
      const humaneCount = getEntryCount('humane-buffer');
      
      return {
        success: true,
        analogyBuffer: {
          name: 'analogy-buffer',
          count: analogyCount,
          exists: analogyCount >= 0
        },
        humaneBuffer: {
          name: 'humane-buffer',
          count: humaneCount,
          exists: humaneCount >= 0
        },
        totalEntries: analogyCount + humaneCount
      };
      
    } catch (error) {
      console.error('Error getting buffer sheets info:', error);
      throw new Error(`Failed to get buffer sheets info: ${error.message}`);
    }
  }
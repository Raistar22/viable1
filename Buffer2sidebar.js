/**
 * Gets the number of entries (excluding header) in the specified sheet
 * @param {string} sheetName - Name of the sheet to count entries from
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
   * Buffer2sidebar.gs
   * Google Apps Script functions for managing buffer operations
   */
  
  /**
   * Updates the document count in a designated cell or range
   * @param {number} count - The number of documents to update
   * @returns {Object} Result object with success status
   */
  function updateDocumentCount(count) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      
      // You can modify this to target the specific sheet and cell where you want to store the document count
      // For example, if you want to store it in cell A1 of a sheet named "Dashboard"
      const targetSheet = spreadsheet.getSheetByName('Dashboard') || spreadsheet.getActiveSheet();
      
      // Update the document count in cell A1 (modify as needed)
      targetSheet.getRange('A1').setValue(count);
      
      // Optional: Add a label in B1
      targetSheet.getRange('B1').setValue('Total Documents');
      
      // Log the action
      console.log(`Document count updated to: ${count}`);
      
      return {
        success: true,
        message: `Document count updated to ${count}`,
        count: count
      };
      
    } catch (error) {
      console.error('Error updating document count:', error);
      throw new Error(`Failed to update document count: ${error.message}`);
    }
  }
  
  /**
   * Pushes relevant entries to the specified buffer sheet
   * @param {string} targetSheetName - Name of the target sheet (analogy-buffer2 or humane-buffer2)
   * @returns {Object} Result object with count of entries pushed
   */
  function pushRelevantEntries(targetSheetName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      
      // Get the source sheet (modify this based on where your data comes from)
      const sourceSheet = spreadsheet.getActiveSheet();
      
      // Get the target sheet
      let targetSheet = spreadsheet.getSheetByName(targetSheetName);
      
      // Create the target sheet if it doesn't exist
      if (!targetSheet) {
        targetSheet = spreadsheet.insertSheet(targetSheetName);
        console.log(`Created new sheet: ${targetSheetName}`);
      }
      
      // Get all data from source sheet
      const sourceData = sourceSheet.getDataRange().getValues();
      
      if (sourceData.length === 0) {
        throw new Error('No data found in source sheet');
      }
      
      // Filter relevant entries based on your criteria
      const relevantEntries = filterRelevantEntries(sourceData, targetSheetName);
      
      if (relevantEntries.length === 0) {
        return {
          success: true,
          message: 'No relevant entries found to push',
          count: 0
        };
      }
      
      // Get the last row in target sheet to append new data
      const lastRow = targetSheet.getLastRow();
      const startRow = lastRow + 1;
      
      // Clear existing data if this is a fresh push (optional)
      // targetSheet.clear();
      // startRow = 1;
      
      // Push the relevant entries to the target sheet
      if (relevantEntries.length > 0) {
        const targetRange = targetSheet.getRange(startRow, 1, relevantEntries.length, relevantEntries[0].length);
        targetRange.setValues(relevantEntries);
      }
      
      console.log(`Pushed ${relevantEntries.length} entries to ${targetSheetName}`);
      
      return {
        success: true,
        message: `Successfully pushed ${relevantEntries.length} entries to ${targetSheetName}`,
        count: relevantEntries.length
      };
      
    } catch (error) {
      console.error('Error pushing relevant entries:', error);
      throw new Error(`Failed to push entries: ${error.message}`);
    }
  }
  
  /**
   * Filters data based on relevance criteria for the target sheet
   * @param {Array} data - The source data array
   * @param {string} targetSheetName - Name of the target sheet
   * @returns {Array} Filtered array of relevant entries
   */
  function filterRelevantEntries(data, targetSheetName) {
    try {
      // Skip header row
      const headerRow = data[0];
      const dataRows = data.slice(1);
      
      let filteredData = [];
      
      // Add header row to filtered data
      filteredData.push(headerRow);
      
      // Filter based on target sheet type
      if (targetSheetName === 'analogy-buffer2') {
        // Filter for analogy-related entries
        filteredData = filteredData.concat(dataRows.filter(row => {
          // Customize this logic based on your data structure
          // Example: Check if any column contains analogy-related keywords
          return row.some(cell => {
            if (typeof cell === 'string') {
              const cellLower = cell.toLowerCase();
              return cellLower.includes('analogy') || 
                     cellLower.includes('comparison') || 
                     cellLower.includes('similar') ||
                     cellLower.includes('like') ||
                     cellLower.includes('metaphor');
            }
            return false;
          });
        }));
        
      } else if (targetSheetName === 'humane-buffer2') {
        // Filter for humane-related entries
        filteredData = filteredData.concat(dataRows.filter(row => {
          // Customize this logic based on your data structure
          // Example: Check if any column contains humane-related keywords
          return row.some(cell => {
            if (typeof cell === 'string') {
              const cellLower = cell.toLowerCase();
              return cellLower.includes('humane') || 
                     cellLower.includes('human') || 
                     cellLower.includes('compassion') ||
                     cellLower.includes('empathy') ||
                     cellLower.includes('kindness') ||
                     cellLower.includes('ethical');
            }
            return false;
          });
        }));
      }
      
      // Remove header if no data rows were found
      if (filteredData.length === 1) {
        return [];
      }
      
      return filteredData;
      
    } catch (error) {
      console.error('Error filtering relevant entries:', error);
      throw new Error(`Failed to filter entries: ${error.message}`);
    }
  }
  
  /**
   * Gets the current document count from the designated cell
   * @returns {number} Current document count
   */
  function getCurrentDocumentCount() {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const targetSheet = spreadsheet.getSheetByName('Dashboard') || spreadsheet.getActiveSheet();
      
      const countValue = targetSheet.getRange('A1').getValue();
      return typeof countValue === 'number' ? countValue : 0;
      
    } catch (error) {
      console.error('Error getting current document count:', error);
      return 0;
    }
  }
  
  /**
   * Utility function to get sheet statistics
   * @param {string} sheetName - Name of the sheet to analyze
   * @returns {Object} Statistics about the sheet
   */
  function getSheetStats(sheetName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = spreadsheet.getSheetByName(sheetName);
      
      if (!sheet) {
        return {
          exists: false,
          rowCount: 0,
          columnCount: 0,
          lastUpdated: null
        };
      }
      
      const range = sheet.getDataRange();
      
      return {
        exists: true,
        rowCount: range.getNumRows(),
        columnCount: range.getNumColumns(),
        lastUpdated: new Date().toISOString()
      };
      
    } catch (error) {
      console.error('Error getting sheet stats:', error);
      return {
        exists: false,
        error: error.message
      };
    }
  }
  
  /**
   * Clears all data from the specified buffer sheet
   * @param {string} sheetName - Name of the sheet to clear
   * @returns {Object} Result object
   */
  function clearBufferSheet(sheetName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = spreadsheet.getSheetByName(sheetName);
      
      if (!sheet) {
        throw new Error(`Sheet ${sheetName} not found`);
      }
      
      sheet.clear();
      
      return {
        success: true,
        message: `Sheet ${sheetName} cleared successfully`
      };
      
    } catch (error) {
      console.error('Error clearing buffer sheet:', error);
      throw new Error(`Failed to clear sheet: ${error.message}`);
    }
  }
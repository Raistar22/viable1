/**
 * Processes all Gmail threads across all clients.
 * Delegates to function in attachment_downloader.gs
 */
function processAllGmail() {
    if (typeof runAttachmentDownloader === 'function') {
      runAttachmentDownloader(); // function from attachment_downloader.gs
    } else {
      throw new Error("Attachment downloader function not defined.");
    }
  }
  
  /**
   * Gets dynamic dashboard data from relevant scripts
   */
  function getDashboardData() {
    const analogyStats = getClientStats("analogy");
    const humaneStats = getClientStats("humane");
  
    return {
      systemStatus: "Healthy",
      totalClients: 2,
      activeClients: 2,
      clients: {
        analogy: analogyStats,
        humane: humaneStats
      }
    };
  }
  
  /**
   * Fetches stats for a specific client.
   * Label format should match Gmail label used in attachment_downloader.gs
   */
  function getClientStats(clientName) {
    const labelName = `client-${clientName}`;
    let pendingCount = 0;
    let processedCount = 0;
    let lastUpdated = "Not Available";
  
    try {
      const label = GmailApp.getUserLabelByName(labelName);
      if (label) {
        const threads = label.getThreads();
        pendingCount = threads.length;
  
        // Check processed count from Sheet
        const sheet = getSheetByName(`${clientName}-buffer`);
        if (sheet) {
          const data = sheet.getDataRange().getValues();
          processedCount = data.length > 1 ? data.length - 1 : 0; // subtract header
          lastUpdated = sheet.getRange(1, sheet.getLastColumn()).getNote() || "Not Available";
        }
      }
    } catch (e) {
      Logger.log(`Error for ${clientName}: ${e.message}`);
    }
  
    return {
      pending: pendingCount,
      processed: processedCount,
      lastUpdated: lastUpdated
    };
  }
  
  /**
   * Utility to get a sheet by name
   */
  function getSheetByName(name) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return ss.getSheetByName(name);
  }
  
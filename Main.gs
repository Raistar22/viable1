/**
 * Runs automatically when the spreadsheet is opened.
 * Creates a custom menu in the Google Sheets UI to launch the tools.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // Create a single main menu for all your custom tools
  // You can customize 'My Sheet Tools' to any name you prefer.
  ui.createMenu('Viable Tools')
      .addItem('Viable Gmail Attachment Downloader', 'showGmailAttachmentDownloaderSidebar')
      .addItem('Viable Drive Renamer', 'showDriveRenamerSidebar')
      .addItem('Viable Routing Manager','showRerouteIndexSidebar' ) // Renamed for clarity
      .addToUi();
}

/**
 * Displays the sidebar for the Drive File Renamer tool.
 * Loads the 'DriveRenamer.html' file.
 */
function showDriveRenamerSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('DriveRenamer')
    .setTitle('Drive File Renamer')
    .setWidth(750); // Set desired width for the sidebar
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Displays the sidebar for the Gmail Attachment Downloader tool.
 * Loads the 'index.html' file (assuming this is your downloader's HTML).
 */
function showGmailAttachmentDownloaderSidebar() { // Renamed from showSidebar for clarity
  const html = HtmlService.createHtmlOutputFromFile('index') // Ensure 'index.html' is the correct file name for your downloader's UI
      .setTitle('Gmail Attachment Downloader')
      .setWidth(300); // Set desired width for the sidebar
  SpreadsheetApp.getUi().showSidebar(html);
}
/**
 * Displays the sidebar for the Viable Routing Manager tool.
 * Loads the 'RerouteIndex.html' file (assuming this is your downloader's HTML).
 */

function showRerouteIndexSidebar(){
   const html = HtmlService.createHtmlOutputFromFile('RerouteIndex') // Ensure 'index.html' is the correct file name for your downloader's UI
      .setTitle('Viable Routing Manager')
      .setWidth(300); // Set desired width for the sidebar
  SpreadsheetApp.getUi().showSidebar(html);
}
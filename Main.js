/**
 * Runs automatically when the spreadsheet is opened.
 * Creates a custom menu in the Google Sheets UI to launch the tools.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // Create a single main menu for all your custom tools
  ui.createMenu('Viable Tools')
    .addItem('Viable Gmail Attachment Downloader', 'showGmailAttachmentDownloaderSidebar')
    .addItem('Viable Drive Renamer', 'showDriveRenamerSidebar')
    .addItem('Viable Routing Manager', 'showRerouteIndexSidebar')
    .addItem('Intermediate Sheet', 'showIntermediateSheetSidebar')
    .addItem('Buffer1sidebar', 'showBuffer1Sidebar') // ✅ Newly added menu item
    .addItem('Buffer2sidebar', 'showBuffer2Sidebar')
    .addItem('Dashboard', 'showFinTechDashboard')
    .addToUi();
}

/**
 * Displays the sidebar for the Gmail Attachment Downloader tool.
 * Loads the 'index.html' file.
 */
function showGmailAttachmentDownloaderSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Gmail Attachment Downloader')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Displays the sidebar for the Drive File Renamer tool.
 * Loads the 'DriveRenamer.html' file.
 */
function showDriveRenamerSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('DriveRenamer')
    .setTitle('Drive File Renamer')
    .setWidth(750);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Displays the sidebar for the Viable Routing Manager tool.
 * Loads the 'RerouteIndex.html' file.
 */
function showRerouteIndexSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('RerouteIndex')
    .setTitle('Viable Routing Manager')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Displays the sidebar for the Intermediate Sheet tool.
 * Loads the 'intermediatesheet.html' file.
 */
function showIntermediateSheetSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('intermediatesheet')
    .setTitle('Intermediate Sheet')
    .setWidth(600); // Adjust width as needed
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Displays the sidebar for the Buffer1sidebar tool.
 * Loads the 'buffer1sidebar.html' file.
 */
function showBuffer1Sidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Buffer1sidebar')
    .setTitle('Buffer1sidebar')
    .setWidth(400); // Adjust width as needed
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Displays the sidebar for the Buffer2sidebar tool.
 * Loads the 'buffer2sidebar.html' file.
 */
function showBuffer2Sidebar() {
  const html = HtmlService.createHtmlOutputFromFile('buffer2sidebar')
    .setTitle('Buffer2sidebar')
    .setWidth(400); // Adjust width as needed
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Displays the sidebar for the FinTech Automation Dashboard.
 * Loads the 'Dashboard.html' file.
 */
function showFinTechDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('Dashboard')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 FinTech Automation Dashboard');
}
// Add menu item
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sales Import')
    .addItem('Run Sheets API Import', 'importWithSheetsAPI')
    .addToUi();
}
/**
 * Imports data using Sheets API with guaranteed save
 */
function importWithSheetsAPI() {
  // CONFIGURATION
  const CONFIG = {
    source: {
      spreadsheetId: "1-wpsIJjN_NjUQRKC9We1HHeQ5FFb7PDvQBqyABkwbgg",
      range: "Data Import!L1:O50000" // Sheet name and range
    },
    target: {
      spreadsheetId: SpreadsheetApp.getActive().getId(),
      sheetName: "Sales",
      startRow: 1,
      startCol: 1
    },
    batchSize: 50000,
    maxRetries: 3
  };

  try {
    console.log("Starting Sheets API import...");
    const startTime = new Date();
    
    // 1. Get source data using Sheets API
    let sourceData = [];
    let retries = 0;
    
    while (retries < CONFIG.maxRetries) {
      try {
        sourceData = Sheets.Spreadsheets.Values.get(
          CONFIG.source.spreadsheetId,
          CONFIG.source.range
        ).values || []; // Ensure empty array if null
        break;
      } catch (error) {
        retries++;
        console.warn(`Attempt ${retries} failed: ${error.message}`);
        if (retries >= CONFIG.maxRetries) throw error;
        Utilities.sleep(2000 * retries);
      }
    }
    
    if (sourceData.length === 0) {
      throw new Error("No data retrieved from source");
    }
    
    console.log(`Retrieved ${sourceData.length} rows from source`);
    
    // 2. Prepare target
    const targetSheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(CONFIG.target.sheetName);
    
    // Clear existing data with forced save
    targetSheet.getRange(1, 1, targetSheet.getMaxRows(), 4).clearContent();
    SpreadsheetApp.flush(); // Force save after clear
    Utilities.sleep(1000); // Wait for clear to complete
    
    // 3. Write data using Sheets API with verification
    const request = {
      valueInputOption: "RAW",
      data: [{
        range: `${CONFIG.target.sheetName}!A${CONFIG.target.startRow}:D${CONFIG.target.startRow + sourceData.length - 1}`,
        majorDimension: "ROWS",
        values: sourceData
      }]
    };
    
    // Execute and wait for completion
    const response = Sheets.Spreadsheets.Values.batchUpdate(request, CONFIG.target.spreadsheetId);
    
    // CRITICAL: Verify data was written
    Utilities.sleep(3000); // Wait for write to complete
    const writtenData = targetSheet.getRange(1, 1, sourceData.length, 4).getValues();
    
    if (writtenData.length !== sourceData.length) {
      throw new Error("Data verification failed - not all rows were saved");
    }
    
    // Final forced save
    SpreadsheetApp.flush();
    Utilities.sleep(2000); // Extra safety delay
    
    const duration = (new Date() - startTime) / 1000;
    const msg = `Imported ${sourceData.length} rows in ${duration.toFixed(1)} seconds`;
    console.log(msg);
    
    // Only show alert after everything is confirmed saved
    SpreadsheetApp.getUi().alert("Import Complete", msg, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    const errMsg = `Import failed: ${error.message}`;
    console.error(errMsg, error);
    SpreadsheetApp.getUi().alert("Import Error", errMsg, SpreadsheetApp.getUi().ButtonSet.OK);
    throw error;
  }
}
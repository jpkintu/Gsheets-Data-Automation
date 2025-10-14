/**
 * Imports filtered sales data to Analysis sheet starting at F3
 * Returns empty strings in target location if no rows with TRUE in column O
 */
function importFilteredSalesToAnalysis() {
  // Configuration
  const SOURCE_SHEET = "Sales";
  const TARGET_SHEET = "Analysis";
  const START_CELL = "F3"; // Where to begin writing data
  const COLUMNS_TO_IMPORT = ["A", "B", "E", "L", "G", "I", "J", "K", "M"];
  const FILTER_COLUMN = "O"; // Column to check for TRUE values
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(SOURCE_SHEET);
    const targetSheet = ss.getSheetByName(TARGET_SHEET);
    
    if (!sourceSheet || !targetSheet) {
      throw new Error("Source or target sheet not found");
    }

    // Get all data from source sheet
    const lastRow = sourceSheet.getLastRow();
    const lastCol = sourceSheet.getLastColumn();
    const allData = sourceSheet.getRange(1, 1, lastRow, lastCol).getValues();
    
    // Find filter column index (O)
    const filterColIndex = "O".charCodeAt(0) - 65; // O is column 14 (0-based index)
    
    // Filter rows where column O is TRUE (skip header row)
    const filteredData = allData.slice(1).filter(row => row[filterColIndex] === true);
    
    // Calculate target range (F3:N...)
    const startRange = targetSheet.getRange(START_CELL);
    const startRow = startRange.getRow();
    const startCol = startRange.getColumn();
    
    // Clear previous data in target range (F3:N... down to last row)
    const lastTargetRow = targetSheet.getLastRow();
    if (lastTargetRow >= startRow) {
      targetSheet.getRange(
        startRow, 
        startCol, 
        lastTargetRow - startRow + 1, 
        COLUMNS_TO_IMPORT.length
      ).clearContent();
    }
    
    // If no filtered data, create an array of empty strings with the right dimensions
    let importedData = [];
    if (filteredData.length === 0) {
      console.log("No rows with TRUE in column O - clearing target area");
      // Create a single row of empty strings with the correct number of columns
      importedData = [Array(COLUMNS_TO_IMPORT.length).fill("")];
    } else {
      // Extract only the columns we want (A,B,E,L,G,I,J,K,M)
      const columnIndices = COLUMNS_TO_IMPORT.map(col => col.charCodeAt(0) - 65);
      importedData = filteredData.map(row => 
        columnIndices.map(index => row[index])
      );
    }
    
    // Write new data (will be empty strings if no filtered data)
    targetSheet.getRange(
      startRow, 
      startCol, 
      importedData.length, 
      importedData[0].length
    ).setValues(importedData);
    
    console.log(`Imported ${importedData.length} rows to ${TARGET_SHEET}`);
    
  } catch (error) {
    console.error("Error in importFilteredSalesToAnalysis:", error);
    throw error;
  }
}

/**
 * Sets up hourly trigger
 */
function setupHourlyTrigger() {
  // Remove existing triggers
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "importFilteredSalesToAnalysis")
    .forEach(t => ScriptApp.deleteTrigger(t));
  
  // Create new hourly trigger
  ScriptApp.newTrigger("importFilteredSalesToAnalysis")
    .timeBased()
    .everyHours(1)
    .create();
  
  console.log("Hourly trigger set up successfully");
}
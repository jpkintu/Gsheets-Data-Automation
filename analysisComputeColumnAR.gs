function computeColumnAR() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Analysis");
  
  // Get all necessary data in one batch
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange("A3:AR" + lastRow);
  const data = range.getValues();
  
  // Column indices (0-based)
  const colA = 0;    // Column A
  const colAI = 34;  // Column AI (index 34)
  const colAJ = 35;  // Column AJ
  const colAK = 36;  // Column AK
  const colAQ = 42;  // Column AQ
  const colAR = 43;  // Column AR (target column)
  
  // Prepare results array
  const results = [];
  
  // Process each row
  for (let i = 0; i < data.length; i++) {
    const a = data[i][colAI]; // AI value
    const b = data[i][colAJ]; // AJ value
    const p = data[i][colAQ]; // AQ value
    const j = data[i][colAK]; // AK value
    
    // If AI is empty, return empty string
    if (a === "" || a === null || a === undefined) {
      results.push([""]);
      continue;
    }
    
    // If p <= 0, return j
    if (p <= 0) {
      results.push([j]);
      continue;
    }
    
    // Otherwise find matching row
    let matchFound = false;
    let matchValue = "No Match";
    
    // Search through all rows for match
    for (let k = 0; k < data.length; k++) {
      if (data[k][colAI] === a && 
          data[k][colAJ] === b && 
          data[k][colAQ] <= 0 && 
          data[k][colA] > j) {  // Column A > j
        matchFound = true;
        matchValue = data[k][colAK]; // Return AK value from matching row
        break;
      }
    }
    
    results.push([matchValue]);
  }
  
  // Write results to column AR starting at row 3
  if (results.length > 0) {
    sheet.getRange(3, colAR + 1, results.length, 1).setValues(results);
    console.log(`Successfully wrote ${results.length} values to column AR`);
  }
  
  // Format column AR
  sheet.getRange(3, colAR + 1, results.length, 1)
    .setNumberFormat("@"); // Text format to match formula behavior
}
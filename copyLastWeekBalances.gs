function copyLastWeekBalances() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const balancesSheet = ss.getSheetByName("Balances Query");
  
  if (!balancesSheet) {
    throw new Error("Balances Query sheet not found");
  }

  // Get week headers (J1:CA1)
  const weekHeaders = balancesSheet.getRange("J1:CA1").getValues()[0];
  
  // Find the last non-empty week header
  let lastWeekCol = -1;
  for (let i = weekHeaders.length - 1; i >= 0; i--) {
    if (weekHeaders[i] && weekHeaders[i] !== "") {
      lastWeekCol = i + 10; // Column J is 10 (1-based)
      break;
    }
  }

  if (lastWeekCol === -1) {
    console.log("No week headers found");
    return;
  }

  // Get the last week's data (entire column starting at row 2)
  const lastWeekData = balancesSheet.getRange(2, lastWeekCol, balancesSheet.getLastRow() - 1, 1).getValues();
  
  // Write to column G (G2:G)
  balancesSheet.getRange("G2:G" + (lastWeekData.length + 1)).setValues(lastWeekData);
  
  // Format as numbers with 2 decimals
  balancesSheet.getRange("G2:G" + (lastWeekData.length + 1)).setNumberFormat("#,###");
  
  console.log(`Copied last week (${weekHeaders[lastWeekCol-10]}) balances to column G`);
}

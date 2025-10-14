function computeWeeklyBankings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const weeklyBankingSheet = ss.getSheetByName("Weekly Banking Query");
  const merchantStmtsSheet = ss.getSheetByName("Merchant Stmts");
  
  if (!weeklyBankingSheet || !merchantStmtsSheet) {
    throw new Error("Required sheets not found");
  }

  // Get agent data from Weekly Banking Query
  const agentIds = weeklyBankingSheet.getRange("A2:A").getValues().flat();
  const weekHeaders = weeklyBankingSheet.getRange("C1:LM1").getValues()[0];
  
  // Get pre-summed banking data (A=agentId, B=week, C=amount)
  const lastRow = merchantStmtsSheet.getLastRow();
  const summedBankings = merchantStmtsSheet.getRange("A2:C" + lastRow).getValues();
  
  // Create a map for quick lookup {agentId: {week: total}}
  const agentWeekTotals = {};
  
  // Process all pre-summed records
  summedBankings.forEach(row => {
    const agentId = row[0]; // Column A
    const week = row[1];    // Column B
    const amount = row[2];  // Column C
    
    if (!agentWeekTotals[agentId]) {
      agentWeekTotals[agentId] = {};
    }
    agentWeekTotals[agentId][week] = amount;
  });
  
  // Prepare results matrix
  const results = [];
  const lastAgentRow = agentIds.filter(String).length;
  
  for (let i = 0; i < lastAgentRow; i++) {
    const agentId = agentIds[i];
    const rowData = [];
    
    weekHeaders.forEach(week => {
      if (!week || week === "") {
        rowData.push("");
      } else if (!agentId) {
        rowData.push("");
      } else {
        const total = agentWeekTotals[agentId]?.[week] || "";
        rowData.push(total);
      }
    });
    
    results.push(rowData);
  }
  
  // Write results to weekly banking sheet (C2:LM...)
  if (results.length > 0 && results[0].length > 0) {
    weeklyBankingSheet.getRange(2, 3, results.length, results[0].length)
      .setValues(results)
      .setNumberFormat("#,###"); // Format as numbers
  }
  
  console.log("Weekly banking computation completed using pre-summed data");
}

// Set up time-based trigger to run weekly
function setupWeeklyTrigger() {
  // Delete existing triggers to avoid duplicates
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'computeWeeklyBankings')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Create new weekly trigger (runs every Monday at 6 AM)
  ScriptApp.newTrigger('computeWeeklyBankings')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(6)
    .create();
  
  console.log("Weekly banking trigger set up");
}
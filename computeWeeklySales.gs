function computeWeeklySales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const weeklySheet = ss.getSheetByName("Weekly Sales Query");
  const salesSheet = ss.getSheetByName("Sales");
  
  if (!weeklySheet || !salesSheet) {
    throw new Error("Required sheets not found");
  }

  // Get agent data from Weekly Sales Query
  const agentIds = weeklySheet.getRange("A2:A").getValues().flat();
  const agentNames = weeklySheet.getRange("B2:B").getValues().flat();
  const weekHeaders = weeklySheet.getRange("C1:LM1").getValues()[0];
  
  // Get sales data
  const salesData = salesSheet.getRange("E2:N" + salesSheet.getLastRow()).getValues();
  
  // Create a map for quick lookup {agentId: {week: total}}
  const agentWeekTotals = {};
  
  // Process all sales records
  salesData.forEach(row => {
    const amount = row[0]; // Column E (index 0)
    const agentId = row[8]; // Column M (index 8)
    const week = row[9]; // Column N (index 9)
    
    if (!agentWeekTotals[agentId]) {
      agentWeekTotals[agentId] = {};
    }
    
    if (!agentWeekTotals[agentId][week]) {
      agentWeekTotals[agentId][week] = 0;
    }
    
    agentWeekTotals[agentId][week] += amount;
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
  
  // Write results to weekly sheet (C2:LM...)
  if (results.length > 0 && results[0].length > 0) {
    weeklySheet.getRange(2, 3, results.length, results[0].length)
      .setValues(results)
      .setNumberFormat("#,###"); // Format as numbers with 2 decimals
  }
  
  console.log("Weekly sales computation completed");
}

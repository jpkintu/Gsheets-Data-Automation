function generateFilteredTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Get source data
  const merchantSheet = ss.getSheetByName('Merchant Raw');
  const merchantData = merchantSheet.getRange('A2:I').getValues();
  
  // 2. Get date range from Dashboard
  const dashboardSheet = ss.getSheetByName('Dashboard');
  const startDate = new Date(dashboardSheet.getRange('B1').getValue());
  const endDate = new Date(dashboardSheet.getRange('D1').getValue());
  
  // 3. Filter and transform data
  const filteredData = merchantData.filter(row => {
    const rowDate = new Date(row[7]); // H column (index 7)
    return rowDate && rowDate >= startDate && rowDate <= endDate;
  }).map(row => [
    row[1],  // B column - Agent ID
    row[6],  // G column - Name
    row[7],  // H column - Date
    row[5]   // F column - Amount
  ]);
  
  // 4. Write to Analysis sheet
  const analysisSheet = ss.getSheetByName('Analysis');
  
  // Clear only existing data in Q3:T (not entire columns)
  const lastRowWithData = analysisSheet.getRange('Q:T').getValues()
    .findIndex(row => row.join('').trim() === '');
  const clearRange = lastRowWithData === -1 ? 
    analysisSheet.getRange('Q3:T' + analysisSheet.getLastRow()) :
    analysisSheet.getRange('Q3:T' + (lastRowWithData + 2));
  clearRange.clearContent();
  
  // Write new data if any exists
  if (filteredData.length > 0) {
    const outputRange = analysisSheet.getRange('Q3:T' + (3 + filteredData.length - 1));
    outputRange.setValues(filteredData);
    
  
  }
  
  console.log(`Successfully wrote ${filteredData.length} rows to Analysis sheet`);
}

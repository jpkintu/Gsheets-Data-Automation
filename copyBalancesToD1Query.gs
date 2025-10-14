function copyBalancesToD1Query() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const balancesSheet = ss.getSheetByName("Balances Query");
  const d1QuerySheet = ss.getSheetByName("D1 Query");

  if (!balancesSheet || !d1QuerySheet) {
    throw new Error("Required sheets not found");
  }

  // 1. Copy columns E, F, I from Balances Query to A, B, C in D1 Query
  const agentIds = balancesSheet.getRange("E1:E" + balancesSheet.getLastRow()).getValues();
  const agentNames = balancesSheet.getRange("F1:F" + balancesSheet.getLastRow()).getValues();
  const agentCategories = balancesSheet.getRange("I1:I" + balancesSheet.getLastRow()).getValues();

  // Combine the columns and write to D1 Query
  const basicInfo = agentIds.map((id, i) => [id[0], agentNames[i][0], agentCategories[i][0]]);
  d1QuerySheet.getRange("A1:C" + (basicInfo.length + 0)).setValues(basicInfo);

  // 2. Find the last 10 non-empty week headers in Balances Query (J1:CA1)
  const weekHeaders = balancesSheet.getRange("J1:CA1").getValues()[0];
  
  // Get indices of all non-empty week headers
  const nonEmptyWeeks = weekHeaders
    .map((header, index) => ({ header, index }))
    .filter(item => item.header && item.header !== "");

  // Take the last 10 non-empty weeks
  const last10Weeks = nonEmptyWeeks.slice(-10);

  if (last10Weeks.length === 0) {
    console.log("No week headers found");
    return;
  }

  // 3. Copy the corresponding week data from Balances Query to D1 Query (columns D-M)
  const weekData = [];
  last10Weeks.forEach(week => {
    const columnData = balancesSheet.getRange(2, week.index + 10, balancesSheet.getLastRow() - 1, 1)
      .getValues()
      .map(row => row[0]);
    weekData.push(columnData);
  });

  // Transpose the data to get it in the right format (rows x 10 columns)
  const transposedData = weekData[0].map((_, i) => 
    weekData.map(week => week[i])
  );

  // Write to D1 Query (D2:M...)
  if (transposedData.length > 0) {
    d1QuerySheet.getRange("D2:M" + (transposedData.length + 1))
      .setValues(transposedData)
     .setNumberFormat("#,###;(#,###)");
  }

  // Update headers in D1 Query (D1:M1)
  const newHeaders = last10Weeks.map(week => week.header);
  d1QuerySheet.getRange("D1:M1").setValues([newHeaders]);

  console.log(`Copied ${basicInfo.length} agents and ${last10Weeks.length} weeks of data to D1 Query`);
}

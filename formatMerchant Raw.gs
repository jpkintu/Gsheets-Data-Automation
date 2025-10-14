function formatAgentNames() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Merchant Raw"); 
  const range = sheet.getRange("B2:B" + sheet.getLastRow()); // Column B
  const values = range.getValues();

  // Format each value (capitalize first letter, lowercase the rest)
  const formattedValues = values.map(row => {
    if (row[0] && typeof row[0] === 'string') {
      return [row[0].charAt(0).toUpperCase() + row[0].slice(1).toLowerCase()];
    }
    return [""]; // Skip if empty
  });

  // Apply changes back to the sheet
  range.setValues(formattedValues);
  console.log("Formatted Column B (e.g., 's114' → 'S114')");
}
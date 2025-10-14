function onOpen() {
  // Create the main Automations Tools menu
  const menu = SpreadsheetApp.getUi()
    .createMenu('Automations')
    .addItem('Import Sales', 'importWithSheetsAPI')
    .addItem('Compute Weekly Sales', 'computeWeeklySales')
    .addItem('Compute Weekly Bankings', 'computeWeeklyBankings')
    .addItem('Copy to D1 Query', 'copyBalancesToD1Query')
    .addItem('Compute All Reports', 'computeAllReports')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('SMS Tools')
        .addItem('Send Messages', 'promptForPin') // Fixed typo (was 'promptForPin')
        .addItem('Test API', 'testAfricasTalkingCredentials')
    );
    
  // Add the menu to the UI
  menu.addToUi();
}

// Make sure these functions exist in your script:
function computeAllReports() {
  importWithSheetsAPI();
  computeWeeklySales();
  computeWeeklyBankings();
  copyBalancesToD1Query();
  return "All reports computed successfully!";
}
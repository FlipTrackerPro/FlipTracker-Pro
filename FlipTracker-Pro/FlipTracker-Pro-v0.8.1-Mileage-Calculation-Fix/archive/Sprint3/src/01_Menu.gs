function onOpen() {
  SpreadsheetApp.getUi().createMenu('FlipTracker Pro')
    .addItem('Install / Repair Sprint 3','initializeFlipTrackerProSprint3')
    .addSeparator()
    .addItem('Add Inventory Item','showAddItemForm')
    .addItem('Edit Selected Inventory Item','showEditSelectedItemForm')
    .addItem('Find Inventory Item','findInventoryItem')
    .addItem('Show Slow Inventory','showSlowInventory')
    .addSeparator()
    .addItem('Record Sale','showRecordSaleForm')
    .addItem('Record Expense','showRecordExpenseForm')
    .addItem('Record Mileage','showRecordMileageForm')
    .addItem('Add Packaging Supply','showPackagingForm')
    .addSeparator()
    .addItem('Refresh Dashboard','refreshDashboardSprint3')
    .addItem('Go to Dashboard','goToDashboardSprint3')
    .addToUi();
}

function initializeFlipTrackerProSprint3() {
  buildAdminSprint3_();
  buildSettingsSprint3_();
  buildInventorySprint3_();
  buildSalesSprint3_();
  buildExpensesSprint3_();
  buildMileageSprint3_();
  buildPackagingSprint3_();
  buildDashboardSprint3_();
  goToDashboardSprint3();
  SpreadsheetApp.getActive().toast('Sprint 3 transaction engine is ready.','FlipTracker Pro',6);
}

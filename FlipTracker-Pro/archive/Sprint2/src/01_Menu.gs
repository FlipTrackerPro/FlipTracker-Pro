function onOpen() {
  SpreadsheetApp.getUi().createMenu('FlipTracker Pro')
    .addItem('Install / Repair Sprint 2','initializeFlipTrackerProSprint2')
    .addSeparator()
    .addItem('Add Inventory Item','showAddItemForm')
    .addItem('Edit Selected Item','showEditSelectedItemForm')
    .addItem('Find Item','findInventoryItem')
    .addItem('Show Slow Inventory','showSlowInventory')
    .addSeparator()
    .addItem('Refresh Dashboard','refreshDashboardSprint2')
    .addItem('Go to Inventory','goToInventorySprint2')
    .addToUi();
}

function initializeFlipTrackerProSprint2() {
  buildAdminSprint2_();
  buildInventorySprint2_();
  buildDashboardSprint2_();
  SpreadsheetApp.getActive().toast('Sprint 2 inventory engine is ready.','FlipTracker Pro',5);
}

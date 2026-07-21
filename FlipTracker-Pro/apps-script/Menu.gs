function onOpen() {
  SpreadsheetApp.getUi().createMenu('FlipTracker Pro')
    .addItem('Install / Repair v0.4.3','installFlipTrackerPro')
    .addItem('Upgrade Existing Workbook','upgradeFlipTrackerPro')
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
    .addItem('Go to Packaging Inventory','goToPackagingInventory3_')
    .addItem('Refresh Packaging Dropdowns','refreshPackagingDropdowns3_')
    .addSeparator()
    .addItem('Refresh Dashboard','refreshDashboardSprint3')
    .addItem('Open CRA Tax Centre','goToTaxCentreV04')
    .addItem('Refresh CRA Tax Centre','refreshTaxCentreV04')
    .addItem('Build Accountant Export','buildAccountantExportV04_')
    .addItem('Run Calculation Audit','runFlipTrackerCalculationAudit')
    .addItem('Go to Dashboard','goToDashboardSprint3')
    .addToUi();
}

function initializeFlipTrackerProSprint3() { installFlipTrackerPro(); }

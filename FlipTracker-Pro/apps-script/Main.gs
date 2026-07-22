function installFlipTrackerPro() {
  buildAdminSprint3_(); buildSettingsSprint3_(); buildInventorySprint3_();
  buildSalesSprint3_(); buildExpensesSprint3_(); buildMileageSprint3_();
  buildPackagingSprint3_(); buildDashboardSprint3_(); buildTaxCentreV04_();
  const p=PropertiesService.getDocumentProperties();
  buildInventorySprint3_(); // Always repair validation and calculated columns.
  buildInventorySprint3_();
  repairInventoryCalculations3_();
  p.setProperty('FTP_SCHEMA_VERSION','5.5'); p.setProperty('FTP_APP_VERSION',FTP3.VERSION);
  goToDashboardSprint3();
  SpreadsheetApp.getActive().toast('FlipTracker Pro v0.5.5 is ready.','FlipTracker Pro',6);
}

function upgradeFlipTrackerPro() {
  const p=PropertiesService.getDocumentProperties();
  const current=Number(p.getProperty('FTP_SCHEMA_VERSION')||0);
  if(current<1) migrateToSchema1_();
  if(current<2) migrateToSchema2_();
  if(current<3) migrateToSchema3_();
  if(current<4) migrateToSchema4_();
  if(current<4.1) migrateToSchema41_();
  if(current<4.2) migrateToSchema42_();
  if(current<4.3) migrateToSchema43_();
  if(current<4.4) migrateToSchema44_();
  if(current<4.5) migrateToSchema45_();
  if(current<4.6) migrateToSchema46_();
  if(current<4.7) migrateToSchema47_();
  if(current<4.8) migrateToSchema48_();
  if(current<4.9) migrateToSchema49_();
  if(current<5.0) migrateToSchema50_();
  if(current<5.1) migrateToSchema51_();
  if(current<5.2) migrateToSchema52_();
  if(current<5.3) migrateToSchema53_();
  if(current<5.4) migrateToSchema54_();
  if(current<5.5) migrateToSchema55_();
  buildInventorySprint3_();
  repairInventoryCalculations3_();
  p.setProperty('FTP_SCHEMA_VERSION','5.5'); p.setProperty('FTP_APP_VERSION',FTP3.VERSION);
  SpreadsheetApp.getActive().toast('FlipTracker Pro upgraded to schema 5.5.','FlipTracker Pro',6);
}
function migrateToSchema1_(){buildAdminSprint3_();buildSettingsSprint3_();buildInventorySprint3_();}
function migrateToSchema2_(){buildInventorySprint3_();}
function migrateToSchema3_(){buildSalesSprint3_();buildExpensesSprint3_();buildMileageSprint3_();buildPackagingSprint3_();buildDashboardSprint3_();}
function migrateToSchema4_(){buildSettingsSprint3_();buildTaxCentreV04_();buildDashboardSprint3_();}
function migrateToSchema41_(){buildTaxCentreV04_();buildDashboardSprint3_();}
function migrateToSchema42_(){buildPackagingSprint3_();buildSalesSprint3_();buildDashboardSprint3_();}
function migrateToSchema43_(){buildPackagingSprint3_();buildSalesSprint3_();refreshPackagingDropdowns3_();}
function migrateToSchema44_(){buildAdminSprint3_();buildInventorySprint3_();buildSalesSprint3_();}
function migrateToSchema45_(){buildInventorySprint3_();buildSalesSprint3_();}
function migrateToSchema46_(){buildSalesSprint3_();}
function migrateToSchema47_(){buildAdminSprint3_();buildSalesSprint3_();}
function migrateToSchema48_(){
  buildInventorySprint3_();
  buildSalesSprint3_();
  populateSalesDescriptions48_();
}
function populateSalesDescriptions48_(){
  const sales=sheet3_(FTP3.SHEETS.SALES);
  if(sales.getLastRow()<2)return;
  const inventory=sheet3_(FTP3.SHEETS.INVENTORY);
  const map={};
  if(inventory.getLastRow()>1){
    const headers=inventory.getRange(1,1,1,inventory.getLastColumn()).getDisplayValues()[0];
    const itemIdIndex=headerIndex3_(headers,'Item ID'), descriptionIndex=headerIndex3_(headers,'Description');
    inventory.getRange(2,1,inventory.getLastRow()-1,inventory.getLastColumn()).getDisplayValues()
      .forEach(r=>{if(r[itemIdIndex])map[String(r[itemIdIndex])]=String(r[descriptionIndex]||'');});
  }
  const ids=sales.getRange(2,2,sales.getLastRow()-1,1).getDisplayValues();
  const descriptions=ids.map(r=>[map[String(r[0])]||'']);
  sales.getRange(2,3,descriptions.length,1).setValues(descriptions);
}
function migrateToSchema49_(){buildInventorySprint3_();buildSalesSprint3_();refreshPackagingDropdowns3_();}
function migrateToSchema50_(){buildInventorySprint3_();buildSalesSprint3_();buildDashboardSprint3_();buildTaxCentreV04_();refreshPackagingDropdowns3_();}
function migrateToSchema51_(){buildInventorySprint3_();buildSalesSprint3_();refreshPackagingDropdowns3_();}
function migrateToSchema52_(){buildInventorySprint3_();buildSalesSprint3_();buildDashboardSprint3_();refreshPackagingDropdowns3_();}
function migrateToSchema53_(){buildInventorySprint3_();repairInventoryCalculations3_();buildDashboardSprint3_();}
function migrateToSchema54_(){buildInventorySprint3_();repairInventoryCalculations3_();buildSalesSprint3_();buildDashboardSprint3_();buildTaxCentreV04_();}
function migrateToSchema55_(){buildInventorySprint3_();repairInventoryCalculations3_();buildSalesSprint3_();}
function getFlipTrackerVersion(){const p=PropertiesService.getDocumentProperties();return{appVersion:p.getProperty('FTP_APP_VERSION')||FTP3.VERSION,schemaVersion:p.getProperty('FTP_SCHEMA_VERSION')||'unversioned'};}

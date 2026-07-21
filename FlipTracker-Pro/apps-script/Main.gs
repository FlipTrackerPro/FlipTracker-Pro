function installFlipTrackerPro() {
  buildAdminSprint3_(); buildSettingsSprint3_(); buildInventorySprint3_();
  buildSalesSprint3_(); buildExpensesSprint3_(); buildMileageSprint3_();
  buildPackagingSprint3_(); buildDashboardSprint3_(); buildTaxCentreV04_();
  const p=PropertiesService.getDocumentProperties();
  p.setProperty('FTP_SCHEMA_VERSION','4.2'); p.setProperty('FTP_APP_VERSION',FTP3.VERSION);
  goToDashboardSprint3();
  SpreadsheetApp.getActive().toast('FlipTracker Pro v0.4 is ready.','FlipTracker Pro',6);
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
  p.setProperty('FTP_SCHEMA_VERSION','4.2'); p.setProperty('FTP_APP_VERSION',FTP3.VERSION);
  SpreadsheetApp.getActive().toast('FlipTracker Pro upgraded to schema 4.2.','FlipTracker Pro',6);
}
function migrateToSchema1_(){buildAdminSprint3_();buildSettingsSprint3_();buildInventorySprint3_();}
function migrateToSchema2_(){buildInventorySprint3_();}
function migrateToSchema3_(){buildSalesSprint3_();buildExpensesSprint3_();buildMileageSprint3_();buildPackagingSprint3_();buildDashboardSprint3_();}
function migrateToSchema4_(){buildSettingsSprint3_();buildTaxCentreV04_();buildDashboardSprint3_();}
function migrateToSchema41_(){buildTaxCentreV04_();buildDashboardSprint3_();}
function migrateToSchema42_(){buildPackagingSprint3_();buildSalesSprint3_();buildDashboardSprint3_();}
function getFlipTrackerVersion(){const p=PropertiesService.getDocumentProperties();return{appVersion:p.getProperty('FTP_APP_VERSION')||FTP3.VERSION,schemaVersion:p.getProperty('FTP_SCHEMA_VERSION')||'unversioned'};}

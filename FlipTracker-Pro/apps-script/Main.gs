function installFlipTrackerPro() {
  initializeFlipTrackerProSprint3();
}

function upgradeFlipTrackerPro() {
  const props = PropertiesService.getDocumentProperties();
  const current = Number(props.getProperty('FTP_SCHEMA_VERSION') || 0);
  if (current < 1) migrateToSchema1_();
  if (current < 2) migrateToSchema2_();
  if (current < 3) migrateToSchema3_();
  props.setProperty('FTP_SCHEMA_VERSION', '3');
  props.setProperty('FTP_APP_VERSION', FTP3.VERSION);
  SpreadsheetApp.getActive().toast('FlipTracker Pro upgraded to schema 3.','FlipTracker Pro',6);
}

function migrateToSchema1_() {
  buildAdminSprint3_(); buildSettingsSprint3_(); buildInventorySprint3_();
}
function migrateToSchema2_() { buildInventorySprint3_(); }
function migrateToSchema3_() {
  buildSalesSprint3_(); buildExpensesSprint3_(); buildMileageSprint3_();
  buildPackagingSprint3_(); buildDashboardSprint3_();
}
function getFlipTrackerVersion() {
  const p=PropertiesService.getDocumentProperties();
  return {appVersion:p.getProperty('FTP_APP_VERSION')||FTP3.VERSION,schemaVersion:p.getProperty('FTP_SCHEMA_VERSION')||'unversioned'};
}

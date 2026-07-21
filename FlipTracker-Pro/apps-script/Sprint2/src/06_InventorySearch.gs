function findInventoryItem() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    'Find Inventory Item',
    'Enter Item ID, SKU, barcode, or part of the title:',
    ui.ButtonSet.OK_CANCEL
  );

  if (res.getSelectedButton() !== ui.Button.OK) return;

  const q = res.getResponseText().trim().toLowerCase();
  if (!q) return;

  const s = sheet_(FTP2.SHEETS.INVENTORY);
  const last = s.getLastRow();

  if (last < 2) {
    ui.alert('No inventory items found.');
    return;
  }

  const data = s.getRange(2,1,last-1,5).getDisplayValues();
  const index = data.findIndex(r =>
    [r[0],r[2],r[3],r[4]].some(v => String(v).toLowerCase().includes(q))
  );

  if (index < 0) {
    ui.alert('No matching item found.');
    return;
  }

  SpreadsheetApp.getActive().setActiveSheet(s);
  s.setActiveRange(s.getRange(index+2,1,1,FTP2.HEADERS.length));
}

function showSlowInventory() {
  const s = sheet_(FTP2.SHEETS.INVENTORY);
  SpreadsheetApp.getActive().setActiveSheet(s);

  if (!s.getFilter()) {
    s.getRange(1,1,FTP2.ROWS+1,FTP2.HEADERS.length).createFilter();
  }

  const criteria = SpreadsheetApp.newFilterCriteria()
    .whenNumberGreaterThanOrEqualTo(90).build();

  s.getFilter().setColumnFilterCriteria(20,criteria);
}

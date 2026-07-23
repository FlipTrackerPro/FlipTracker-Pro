function buildInventorySprint2_() {
  const s = sheet_(FTP2.SHEETS.INVENTORY);
  const h = FTP2.HEADERS;

  if (s.getMaxColumns() < h.length) {
    s.insertColumnsAfter(s.getMaxColumns(),h.length-s.getMaxColumns());
  }
  if (s.getMaxRows() < FTP2.ROWS + 1) {
    s.insertRowsAfter(s.getMaxRows(),FTP2.ROWS+1-s.getMaxRows());
  }

  s.getRange(1,1,1,h.length).setValues([h]);
  header2_(s.getRange(1,1,1,h.length));
  s.setFrozenRows(1);
  s.setFrozenColumns(3);
  s.setRowHeight(1,42);

  const validations = [
    [6,'FTP2_Categories'],[7,'FTP2_PurchaseLocations'],
    [8,'FTP2_StorageLocations'],[9,'FTP2_Conditions'],
    [17,'FTP2_Marketplaces'],[19,'FTP2_Statuses']
  ];

  validations.forEach(([col,name]) => {
    const range = SpreadsheetApp.getActive().getRange(name);
    s.getRange(2,col,FTP2.ROWS,1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(range,true)
        .setAllowInvalid(false)
        .build()
    );
  });

  s.getRange(2,2,FTP2.ROWS,1).setNumberFormat('yyyy-mm-dd');
  s.getRange(2,18,FTP2.ROWS,1).setNumberFormat('yyyy-mm-dd');
  s.getRange(2,11,FTP2.ROWS,6).setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
  s.getRange(2,21,FTP2.ROWS,1).setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
  s.getRange(2,22,FTP2.ROWS,1).setNumberFormat('0.0%;[Red]-0.0%');
  s.getRange(2,26,FTP2.ROWS,2).setNumberFormat('yyyy-mm-dd hh:mm');

  if (s.getFilter()) s.getFilter().remove();
  s.getRange(1,1,FTP2.ROWS+1,h.length).createFilter();

  const data = s.getRange(2,1,FTP2.ROWS,h.length);
  s.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$S2="Sold"')
      .setBackground(FTP2.COLORS.LIGHT_GREEN).setRanges([data]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$S2="Listed"')
      .setBackground(FTP2.COLORS.LIGHT_BLUE).setRanges([data]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($T2>=30,$T2<60,$C2<>"",$S2<>"Sold")')
      .setBackground(FTP2.COLORS.GOLD).setRanges([data]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($T2>=60,$T2<90,$C2<>"",$S2<>"Sold")')
      .setBackground('#FCE5CD').setRanges([data]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($T2>=90,$C2<>"",$S2<>"Sold",$S2<>"Archived")')
      .setBackground(FTP2.COLORS.LIGHT_RED).setRanges([data]).build()
  ]);

  const widths = [120,105,240,110,120,150,170,160,110,80,110,95,130,110,120,110,150,105,120,110,125,110,180,180,260,145,145];
  widths.forEach((w,i) => s.setColumnWidth(i+1,w));
}

function nextItemId_() {
  const s = sheet_(FTP2.SHEETS.INVENTORY);
  const rows = Math.max(s.getLastRow()-1,1);
  const ids = s.getRange(2,1,rows,1).getDisplayValues().flat();

  const max = ids.reduce((m,id) => {
    const n = parseInt(String(id).replace(/\D/g,''),10);
    return isNaN(n) ? m : Math.max(m,n);
  },0);

  return 'ITM-' + String(max+1).padStart(5,'0');
}

function saveInventoryItem(form) {
  if (!form.title || !form.purchaseDate) {
    throw new Error('Purchase date and title are required.');
  }

  const s = sheet_(FTP2.SHEETS.INVENTORY);
  const editing = Number(form.row) > 1;
  const row = editing ? Number(form.row) : Math.max(s.getLastRow()+1,2);
  const existing = editing
    ? s.getRange(row,1,1,FTP2.HEADERS.length).getValues()[0]
    : [];

  const purchasePrice = num_(form.purchasePrice);
  const tax = num_(form.taxPaid);
  const shipping = num_(form.acquisitionShipping);
  const expected = num_(form.expectedSalePrice);
  const total = purchasePrice + tax + shipping;
  const profit = expected ? expected-total : '';
  const roi = total && profit !== '' ? profit/total : '';
  const purchaseDate = new Date(form.purchaseDate+'T12:00:00');
  const days = Math.max(0,Math.floor((new Date()-purchaseDate)/86400000));
  const now = new Date();

  const values = [
    editing ? existing[0] : nextItemId_(),
    purchaseDate,form.title,form.sku||'',form.barcode||'',form.category||'',
    form.purchaseLocation||'',form.storageLocation||'',form.condition||'',
    num_(form.quantity)||1,purchasePrice,tax,shipping,total,expected,
    num_(form.listedPrice),form.marketplace||'',existing[17]||'',
    form.status||'Purchased',days,profit,roi,form.receiptLink||'',
    form.photoLink||'',form.notes||'',editing?(existing[25]||now):now,now
  ];

  s.getRange(row,1,1,values.length).setValues([values]);
  s.getRange(row,2).setNumberFormat('yyyy-mm-dd');
  s.getRange(row,11,1,6).setNumberFormat('$#,##0.00');
  s.getRange(row,21).setNumberFormat('$#,##0.00');
  s.getRange(row,22).setNumberFormat('0.0%');
  refreshDashboardSprint2();
}

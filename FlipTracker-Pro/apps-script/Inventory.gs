function buildInventorySprint3_() {
  const s = sheet3_(FTP3.SHEETS.INVENTORY);
  const h = FTP3.INVENTORY_HEADERS;
  ensureSize3_(s,FTP3.ROWS+1,h.length);
  s.getRange(1,1,1,h.length).setValues([h]);
  header3_(s.getRange(1,1,1,h.length));
  s.setFrozenRows(1);
  s.setFrozenColumns(3);
  s.setRowHeight(1,42);

  [[6,'FTP3_Categories'],[7,'FTP3_PurchaseLocations'],
   [8,'FTP3_StorageLocations'],[9,'FTP3_Conditions'],
   [17,'FTP3_Marketplaces'],[19,'FTP3_Statuses']]
    .forEach(v => setValidation3_(s,v[0],v[1],FTP3.ROWS));

  s.getRange(2,2,FTP3.ROWS,1).setNumberFormat('yyyy-mm-dd');
  s.getRange(2,18,FTP3.ROWS,1).setNumberFormat('yyyy-mm-dd');
  s.getRange(2,11,FTP3.ROWS,6).setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
  s.getRange(2,21,FTP3.ROWS,1).setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
  s.getRange(2,22,FTP3.ROWS,1).setNumberFormat('0.0%;[Red]-0.0%');
  s.getRange(2,26,FTP3.ROWS,2).setNumberFormat('yyyy-mm-dd hh:mm');

  if (s.getFilter()) s.getFilter().remove();
  s.getRange(1,1,FTP3.ROWS+1,h.length).createFilter();

  const data = s.getRange(2,1,FTP3.ROWS,h.length);
  s.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$S2="Sold"')
      .setBackground(FTP3.COLORS.LIGHT_GREEN).setRanges([data]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$S2="Listed"')
      .setBackground(FTP3.COLORS.LIGHT_BLUE).setRanges([data]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($T2>=30,$T2<60,$C2<>"",$S2<>"Sold")')
      .setBackground(FTP3.COLORS.GOLD).setRanges([data]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($T2>=60,$T2<90,$C2<>"",$S2<>"Sold")')
      .setBackground(FTP3.COLORS.ORANGE).setRanges([data]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($T2>=90,$C2<>"",$S2<>"Sold",$S2<>"Archived")')
      .setBackground(FTP3.COLORS.LIGHT_RED).setRanges([data]).build()
  ]);
}

function saveInventoryItem(form) {
  if (!form.title || !form.purchaseDate) throw new Error('Purchase date and title are required.');
  const s = sheet3_(FTP3.SHEETS.INVENTORY);
  const editing = Number(form.row) > 1;
  const row = editing ? Number(form.row) : Math.max(s.getLastRow()+1,2);
  const old = editing ? s.getRange(row,1,1,FTP3.INVENTORY_HEADERS.length).getValues()[0] : [];
  const purchasePrice = num3_(form.purchasePrice);
  const tax = num3_(form.taxPaid);
  const shipping = num3_(form.acquisitionShipping);
  const expected = num3_(form.expectedSalePrice);
  const quantity = Math.max(1,Math.floor(num3_(form.quantity) || 1));
  const total = purchasePrice * quantity + tax + shipping;
  const profit = expected ? expected-total : '';
  const roi = total && profit !== '' ? profit/total : '';
  const purchaseDate = date3_(form.purchaseDate);
  const days = Math.max(0,Math.floor((new Date()-purchaseDate)/86400000));
  const now = new Date();

  const values = [
    editing ? old[0] : nextId3_(FTP3.SHEETS.INVENTORY,1,'ITM'),
    purchaseDate,form.title,form.sku||'',form.barcode||'',form.category||'',
    form.purchaseLocation||'',form.storageLocation||'',form.condition||'',
    quantity,purchasePrice,tax,shipping,total,expected,
    num3_(form.listedPrice),form.marketplace||'',old[17]||'',
    form.status||'Purchased',days,profit,roi,form.receiptLink||'',
    form.photoLink||'',form.notes||'',editing?(old[25]||now):now,now
  ];
  s.getRange(row,1,1,values.length).setValues([values]);
  refreshDashboardSprint3();
}

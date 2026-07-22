function buildInventorySprint3_() {
  migrateSheetHeaders3_(FTP3.SHEETS.INVENTORY,FTP3.INVENTORY_HEADERS,{'Description':['Title']});
  const s=sheet3_(FTP3.SHEETS.INVENTORY), h=FTP3.INVENTORY_HEADERS, c=headerMap3_(s);
  ensureSize3_(s,FTP3.ROWS+1,h.length);
  s.getRange(1,1,1,h.length).setValues([h]);header3_(s.getRange(1,1,1,h.length));
  s.setFrozenRows(1);s.setFrozenColumns(2);s.setRowHeight(1,42);

  // Remove stale validation rules left behind by earlier column layouts.
  s.getRange(2,1,FTP3.ROWS,h.length).clearDataValidations();

  [
    ['Category','FTP3_Categories'],['Purchase Location','FTP3_PurchaseLocations'],
    ['Storage Location','FTP3_StorageLocations'],['Condition','FTP3_Conditions'],
    ['Marketplace','FTP3_Marketplaces'],['Status','FTP3_Statuses']
  ].forEach(v=>setValidation3_(s,c[v[0]],v[1],FTP3.ROWS));
  const quantityRule=SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(1).setAllowInvalid(false).build();
  s.getRange(2,c['Quantity'],FTP3.ROWS,1).setDataValidation(quantityRule);

  s.getRange(2,c['Purchase Date'],FTP3.ROWS,1).setNumberFormat('yyyy-mm-dd');
  s.getRange(2,c['Listing Date'],FTP3.ROWS,1).setNumberFormat('yyyy-mm-dd');
  ['Purchase Price','Tax Paid','Acquisition Shipping','Total Cost','Listed Price','Expected Sale Price','Projected Profit']
    .forEach(name=>s.getRange(2,c[name],FTP3.ROWS,1).setNumberFormat('$#,##0.00;[Red]-$#,##0.00'));
  s.getRange(2,c['Projected ROI %'],FTP3.ROWS,1).setNumberFormat('0.0%;[Red]-0.0%');
  s.getRange(2,c['Days in Inventory'],FTP3.ROWS,1).setNumberFormat('0');

  // Calculated fields are maintained by a header-based recalculation routine.
  repairInventoryCalculations3_(s);
  ['Created At','Updated At'].forEach(name=>s.getRange(2,c[name],FTP3.ROWS,1).setNumberFormat('yyyy-mm-dd hh:mm'));

  if(s.getFilter())s.getFilter().remove();
  s.getRange(1,1,FTP3.ROWS+1,h.length).createFilter();

  const data=s.getRange(2,1,FTP3.ROWS,h.length);
  const statusLetter=columnLetter3_(c['Status']);
  const daysLetter=columnLetter3_(c['Days in Inventory']);
  const descriptionLetter=columnLetter3_(c['Description']);
  s.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$'+statusLetter+'2="Sale Pending"').setBackground(FTP3.COLORS.GOLD).setRanges([data]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$'+statusLetter+'2="Sold"').setBackground(FTP3.COLORS.LIGHT_GREEN).setRanges([data]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$'+statusLetter+'2="Listed"').setBackground(FTP3.COLORS.LIGHT_BLUE).setRanges([data]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($'+daysLetter+'2>=30,$'+daysLetter+'2<60,$'+descriptionLetter+'2<>"",$'+statusLetter+'2<>"Sold")').setBackground(FTP3.COLORS.GOLD).setRanges([data]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($'+daysLetter+'2>=60,$'+daysLetter+'2<90,$'+descriptionLetter+'2<>"",$'+statusLetter+'2<>"Sold")').setBackground(FTP3.COLORS.ORANGE).setRanges([data]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($'+daysLetter+'2>=90,$'+descriptionLetter+'2<>"",$'+statusLetter+'2<>"Sold",$'+statusLetter+'2<>"Archived")').setBackground(FTP3.COLORS.LIGHT_RED).setRanges([data]).build()
  ]);

  s.hideColumns(c['Created At'],2);
}

function saveInventoryItem(form) {
  if(!form.description||!form.purchaseDate)throw new Error('Purchase date and description are required.');
  const s=sheet3_(FTP3.SHEETS.INVENTORY), headers=FTP3.INVENTORY_HEADERS;
  const editing=Number(form.row)>1, row=editing?Number(form.row):Math.max(s.getLastRow()+1,2);
  const oldValues=editing?s.getRange(row,1,1,headers.length).getValues()[0]:[];
  const old=editing?rowRecord3_(headers,oldValues):{};
  const purchasePrice=num3_(form.purchasePrice), tax=num3_(form.taxPaid), shipping=num3_(form.acquisitionShipping);
  const quantity=Math.max(1,Math.floor(num3_(form.quantity)||1));
  const total=purchasePrice*quantity+tax+shipping;
  const listed=num3_(form.listedPrice), expected=num3_(form.expectedSalePrice);
  const projectionPrice=listed>0?listed:expected;
  const profit=projectionPrice>0?projectionPrice-total:'';
  const roi=total&&profit!==''?profit/total:'';
  const purchaseDate=date3_(form.purchaseDate);
  const days=Math.max(0,Math.floor((new Date()-purchaseDate)/86400000));
  const status=String(form.status||'Purchased');
  const listingDate=status==='Listed'?(old['Listing Date']||new Date()):(old['Listing Date']||'');
  const now=new Date();

  const record={
    'Item ID':editing?old['Item ID']:nextId3_(FTP3.SHEETS.INVENTORY,1,'ITM'),
    'Description':form.description,'Category':form.category||'','Condition':form.condition||'',
    'SKU':form.sku||'','Barcode':form.barcode||'','Quantity':quantity,
    'Purchase Date':purchaseDate,'Purchase Location':form.purchaseLocation||'',
    'Purchase Price':purchasePrice,'Tax Paid':tax,'Acquisition Shipping':shipping,
    'Total Cost':total,'Status':status,'Marketplace':form.marketplace||'',
    'Listing Date':listingDate,'Storage Location':form.storageLocation||'',
    'Listed Price':listed,'Expected Sale Price':expected,'Days in Inventory':days,
    'Projected Profit':profit,'Projected ROI %':roi,
    'Created At':editing?(old['Created At']||now):now,'Updated At':now,
    'Notes':form.notes||'','Photo Link':form.photoLink||'','Receipt Link':form.receiptLink||''
  };
  s.getRange(row,1,1,headers.length).setValues([headers.map(h=>record[h])]);
  recalculateInventoryRow3_(s,row);
  SpreadsheetApp.flush();
  refreshDashboardSprint3();
}

function recalculateInventoryRow3_(sheet,row) {
  const s=sheet||sheet3_(FTP3.SHEETS.INVENTORY);
  if(row<2)return;
  const c=headerMap3_(s);
  const description=String(s.getRange(row,c['Description']).getDisplayValue()||'').trim();
  if(!description){
    ['Total Cost','Days in Inventory','Projected Profit','Projected ROI %'].forEach(h=>s.getRange(row,c[h]).clearContent());
    return;
  }
  const purchasePrice=num3_(s.getRange(row,c['Purchase Price']).getValue());
  const quantity=Math.max(1,Math.floor(num3_(s.getRange(row,c['Quantity']).getValue())||1));
  const tax=num3_(s.getRange(row,c['Tax Paid']).getValue());
  const shipping=num3_(s.getRange(row,c['Acquisition Shipping']).getValue());
  const total=purchasePrice*quantity+tax+shipping;
  const purchaseDate=s.getRange(row,c['Purchase Date']).getValue();
  const listed=num3_(s.getRange(row,c['Listed Price']).getValue());
  const expected=num3_(s.getRange(row,c['Expected Sale Price']).getValue());
  const saleValue=listed>0?listed:expected;
  const profit=saleValue>0?saleValue-total:'';
  const roi=total>0&&profit!==''?profit/total:'';
  let days='';
  if(purchaseDate instanceof Date&&!isNaN(purchaseDate))days=Math.max(0,Math.floor((new Date()-purchaseDate)/86400000));
  s.getRange(row,c['Quantity']).setValue(quantity);
  s.getRange(row,c['Total Cost']).setValue(total).setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
  s.getRange(row,c['Days in Inventory']).setValue(days).setNumberFormat('0');
  s.getRange(row,c['Projected Profit']).setValue(profit).setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
  s.getRange(row,c['Projected ROI %']).setValue(roi).setNumberFormat('0.0%;[Red]-0.0%');
}

function repairInventoryCalculations3_(sheet) {
  const s=sheet||sheet3_(FTP3.SHEETS.INVENTORY);
  if(s.getLastRow()<2)return;
  for(let row=2;row<=s.getLastRow();row++)recalculateInventoryRow3_(s,row);
}

function repairInventoryCalculations() {
  const s=sheet3_(FTP3.SHEETS.INVENTORY);
  repairInventoryCalculations3_(s);
  SpreadsheetApp.flush();
  SpreadsheetApp.getActive().toast('Inventory Total Cost and projected values recalculated.','FlipTracker Pro',6);
}

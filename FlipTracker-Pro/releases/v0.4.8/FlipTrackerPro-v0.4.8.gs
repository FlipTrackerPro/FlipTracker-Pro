const FTP3 = {
  APP_NAME: 'FlipTracker Pro',
  VERSION: '0.4.8',
  ROWS: 1000,
  SHEETS: {
    DASHBOARD:'Dashboard', INVENTORY:'Inventory', SALES:'Sales',
    EXPENSES:'Expenses', MILEAGE:'Mileage', PACKAGING:'Packaging',
    ADMIN:'Admin', SETTINGS:'Settings', TAX_CENTRE:'Tax Centre', ACCOUNTANT_EXPORT:'Accountant Export', PACKAGING_LISTS:'Packaging Lists'
  },
  COLORS: {
    NAVY:'#1F4E78', BLUE:'#5B9BD5', LIGHT_BLUE:'#D9EAF7',
    GREEN:'#70AD47', LIGHT_GREEN:'#E2F0D9', GOLD:'#FFF2CC',
    ORANGE:'#FCE5CD', LIGHT_RED:'#FCE4D6', GRAY:'#F3F6F9',
    BORDER:'#B7C9D6', WHITE:'#FFFFFF', TEXT:'#1F2937'
  },
  INVENTORY_HEADERS: [
    'Item ID','Purchase Date','Description','SKU','Barcode','Category',
    'Purchase Location','Storage Location','Condition','Quantity',
    'Purchase Price','Tax Paid','Acquisition Shipping','Total Cost',
    'Expected Sale Price','Listed Price','Marketplace','Listing Date',
    'Status','Days in Inventory','Projected Profit','Projected ROI %',
    'Receipt Link','Photo Link','Notes','Created At','Updated At'
  ],
  SALES_HEADERS: [
    'Sale ID','Item ID','Description','Sale Date','Marketplace','Sale Price',
    'Shipping Charged','Shipping Actual','Packaging Cost',
    'Marketplace Fees','Payment Fees','Promotion Expense',
    'GST/HST Collected','Item Cost','Gross Revenue','Total Selling Costs',
    'Net Proceeds','Realized Profit','Realized ROI %','Days to Sell',
    'Buyer','Tracking Number','Notes','Created At',
    'Box Used','Box Qty','Bubble Wrap Used','Bubble Wrap Qty',
    'Mailer Used','Mailer Qty','Tape Used','Tape Qty',
    'Other Packaging Used','Other Packaging Qty','Packaging Verified'
  ],
  EXPENSE_HEADERS: [
    'Expense ID','Date','Category','Vendor','Description','Subtotal',
    'GST/HST Paid','Total','Business Use %','Deductible Amount',
    'Payment Method','Receipt Link','Notes','Created At'
  ],
  MILEAGE_HEADERS: [
    'Trip ID','Date','Start','Destination','Business Purpose',
    'Odometer Start','Odometer End','Total Kilometres',
    'Business Kilometres','CRA Rate','Claim Amount','Notes','Created At'
  ],
  PACKAGING_HEADERS: [
    'Packaging ID','Category','Item Name','Size','Unit of Measure',
    'Units Purchased','Purchase Cost','Cost Per Unit','Quantity On Hand',
    'Reorder Level','Supplier','SKU / Barcode','Product Link','Active',
    'Notes','Updated At'
  ]
};


function sheet3_(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureSize3_(sheet, rows, cols) {
  if (sheet.getMaxRows() < rows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), rows - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < cols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), cols - sheet.getMaxColumns());
  }
}

function header3_(range) {
  range.setBackground(FTP3.COLORS.NAVY)
    .setFontColor(FTP3.COLORS.WHITE).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrap(true);
}

function borders3_(range) {
  range.setBorder(true,true,true,true,true,true,FTP3.COLORS.BORDER,
    SpreadsheetApp.BorderStyle.SOLID);
}

function num3_(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function date3_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') return value;
  return new Date(String(value) + 'T12:00:00');
}

function nextId3_(sheetName, column, prefix) {
  const s = sheet3_(sheetName);
  const last = Math.max(s.getLastRow(), 2);
  const ids = s.getRange(2,column,last-1,1).getDisplayValues().flat();
  const max = ids.reduce((m,id) => {
    const n = parseInt(String(id).replace(/\D/g,''),10);
    return isNaN(n) ? m : Math.max(m,n);
  },0);
  return prefix + '-' + String(max + 1).padStart(5,'0');
}

function setValidation3_(sheet, column, namedRange, rows) {
  const source = SpreadsheetApp.getActive().getRange(namedRange);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(source,true).setAllowInvalid(false).build();
  sheet.getRange(2,column,rows,1).setDataValidation(rule);
}

function selectedInventoryRow3_() {
  const s = SpreadsheetApp.getActiveSheet();
  if (s.getName() !== FTP3.SHEETS.INVENTORY || s.getActiveCell().getRow() < 2) {
    throw new Error('Select an inventory item row first.');
  }
  return s.getActiveCell().getRow();
}

function inventoryItemById3_(itemId) {
  const s = sheet3_(FTP3.SHEETS.INVENTORY);
  if (s.getLastRow() < 2) return null;
  const ids = s.getRange(2,1,s.getLastRow()-1,1).getDisplayValues().flat();
  const index = ids.findIndex(id => id === itemId);
  if (index < 0) return null;
  const row = index + 2;
  return {row:row, values:s.getRange(row,1,1,FTP3.INVENTORY_HEADERS.length).getValues()[0]};
}

function activeInventoryChoices3_() {
  const s = sheet3_(FTP3.SHEETS.INVENTORY);
  if (s.getLastRow() < 2) return [];
  return s.getRange(2,1,s.getLastRow()-1,19).getDisplayValues()
    .filter(r => r[0] && r[2] && !['Sold','Archived'].includes(r[18]))
    .map(r => ({id:r[0], description:r[2] || ''}));
}


function headerMap3_(sheet) {
  if (sheet.getLastRow() < 1) return {};
  const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0];
  return headers.reduce((m,h,i)=>{if(h)m[String(h).trim()]=i+1;return m;},{});
}

function migrateSheetHeaders3_(sheetName, newHeaders, aliases) {
  const s=sheet3_(sheetName);
  ensureSize3_(s,Math.max(s.getMaxRows(),2),newHeaders.length);
  if(s.getLastRow()<1){s.getRange(1,1,1,newHeaders.length).setValues([newHeaders]);return;}
  const oldHeaders=s.getRange(1,1,1,s.getLastColumn()).getDisplayValues()[0];
  const oldData=s.getLastRow()>1?s.getRange(2,1,s.getLastRow()-1,s.getLastColumn()).getValues():[];
  const lookup={}; oldHeaders.forEach((h,i)=>{if(h)lookup[String(h).trim()]=i;});
  aliases=aliases||{};
  const migrated=oldData.map(row=>newHeaders.map(h=>{
    const candidates=[h].concat(aliases[h]||[]);
    const found=candidates.find(c=>Object.prototype.hasOwnProperty.call(lookup,c));
    return found===undefined?'':row[lookup[found]];
  }));
  s.clear(); ensureSize3_(s,Math.max(300,migrated.length+1),newHeaders.length);
  s.getRange(1,1,1,newHeaders.length).setValues([newHeaders]);
  if(migrated.length)s.getRange(2,1,migrated.length,newHeaders.length).setValues(migrated);
}


function buildAdminSprint3_() {
  const s = sheet3_(FTP3.SHEETS.ADMIN);
  const lists = {
    Categories:['Electronics','Tools','Collectibles','Clothing','Furniture','Automotive','Household','Other'],
    PurchaseLocations:['Garage Sale','Thrift Store','Value Village','Facebook Marketplace','Auction','Retail Clearance','Other'],
    Marketplaces:['eBay','Facebook Marketplace','Kijiji','Poshmark','Etsy','Local Sale','Other'],
    Statuses:['Purchased','Needs Cleaning','Needs Testing','Ready to List','Listed','Sale Pending','Sold','Shipped','Archived'],
    Conditions:['New','Open Box','Like New','Good','Fair','For Parts'],
    StorageLocations:['Garage','Basement','Shelf A','Shelf B','Bin 1','Bin 2','Other'],
    ExpenseCategories:['Fuel','Packaging','Shipping Supplies','Advertising','Storage','Phone','Internet','Office','Software','Bank Fees','Equipment','Professional Fees','Other'],
    PaymentMethods:['Cash','Credit Card','Debit','PayPal','Bank Transfer','Other'],
    PackagingTypes:['Box','Bubble Wrap','Mailer','Tape','Label','Packing Paper','Other']
  };
  const existingLists = {};
  if (s.getLastColumn() > 0 && s.getLastRow() > 1) {
    const oldHeaders = s.getRange(1,1,1,s.getLastColumn()).getDisplayValues()[0];
    oldHeaders.forEach((name,index) => {
      if (!name) return;
      existingLists[name] = s.getRange(2,index+1,s.getLastRow()-1,1).getDisplayValues().flat().filter(Boolean);
    });
  }
  s.clear();
  let col = 1;
  Object.keys(lists).forEach(name => {
    s.getRange(1,col).setValue(name);
    const merged = lists[name].concat(existingLists[name] || []).filter((v,i,a) => a.indexOf(v) === i);
    const values = merged.map(v => [v]);
    s.getRange(2,col,values.length,1).setValues(values);
    const rangeName = 'FTP3_' + name;
    SpreadsheetApp.getActive().getNamedRanges()
      .filter(n => n.getName() === rangeName).forEach(n => n.remove());
    SpreadsheetApp.getActive().setNamedRange(rangeName,s.getRange(2,col,50,1));
    s.setColumnWidth(col,180);
    col++;
  });
  header3_(s.getRange(1,1,1,col-1));
  borders3_(s.getRange(1,1,50,col-1));
  s.setFrozenRows(1);
}

function buildSettingsSprint3_() {
  const s = sheet3_(FTP3.SHEETS.SETTINGS);
  const existing = {};
  if (s.getLastRow() >= 2) {
    s.getRange(2,1,s.getLastRow()-1,2).getValues().forEach(r => {
      if (r[0] !== '') existing[String(r[0])] = r[1];
    });
  }
  s.clear();
  const defaults = [
    ['Business Name','','Optional'],
    ['Business Number','','Optional'],
    ['GST/HST Registered','No','Yes or No'],
    ['GST/HST Number','','Optional'],
    ['Province','Ontario',''],
    ['Currency','CAD',''],
    ['Fiscal Year Start','January 1',''],
    ['Default HST Rate',0.13,'Editable'],
    ['CRA Mileage Rate',0,'Optional estimate; confirm tax treatment'],
    ['Tax Year',new Date().getFullYear(),'Tax Centre reporting year'],
    ['Low Packaging Stock Alert','Yes','Yes or No'],
    ['Opening Inventory Value',0,'Cost value at start of tax year'],
    ['Inventory Valuation Method','Lower of cost and FMV','Use a consistent CRA-accepted method'],
    ['GST/HST Reporting Method','Regular','Regular or Quick; Tax Centre estimates Regular only'],
    ['Include Mileage Estimate in Expenses','No','Professional review recommended']
  ];
  const values = [['Setting','Value','Notes']].concat(defaults.map(r => [r[0], Object.prototype.hasOwnProperty.call(existing,r[0]) ? existing[r[0]] : r[1], r[2]]));
  s.getRange(1,1,values.length,3).setValues(values);
  header3_(s.getRange(1,1,1,3));
  borders3_(s.getRange(1,1,values.length,3));
  s.setFrozenRows(1);
  s.setColumnWidths(1,3,220);
  s.getRange('B9').setNumberFormat('0.00%');
  s.getRange('B10').setNumberFormat('$#,##0.00');
  s.getRange('B13').setNumberFormat('$#,##0.00');
  const yesNo = SpreadsheetApp.newDataValidation().requireValueInList(['Yes','No'],true).setAllowInvalid(false).build();
  s.getRange('B4').setDataValidation(yesNo);
  s.getRange('B12').setDataValidation(yesNo);
  s.getRange('B16').setDataValidation(yesNo);
  s.getRange('B14').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Lower of cost and FMV','Fair market value of entire inventory','Cost'],true).setAllowInvalid(false).build());
  s.getRange('B15').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Regular','Quick'],true).setAllowInvalid(false).build());
}


function buildInventorySprint3_() {
  migrateSheetHeaders3_(FTP3.SHEETS.INVENTORY,FTP3.INVENTORY_HEADERS,{'Description':['Title']});
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
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$S2="Sale Pending"')
      .setBackground(FTP3.COLORS.GOLD).setRanges([data]).build(),
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
  if (!form.description || !form.purchaseDate) throw new Error('Purchase date and description are required.');
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
    purchaseDate,form.description,form.sku||'',form.barcode||'',form.category||'',
    form.purchaseLocation||'',form.storageLocation||'',form.condition||'',
    quantity,purchasePrice,tax,shipping,total,expected,
    num3_(form.listedPrice),form.marketplace||'',old[17]||'',
    form.status||'Purchased',days,profit,roi,form.receiptLink||'',
    form.photoLink||'',form.notes||'',editing?(old[25]||now):now,now
  ];
  s.getRange(row,1,1,values.length).setValues([values]);
  refreshDashboardSprint3();
}


function showAddItemForm() {
  showInventoryForm3_(null,'Add Inventory Item');
}

function showEditSelectedItemForm() {
  const row = selectedInventoryRow3_();
  const s = sheet3_(FTP3.SHEETS.INVENTORY);
  showInventoryForm3_({row:row,values:s.getRange(row,1,1,FTP3.INVENTORY_HEADERS.length).getValues()[0]},
    'Edit Inventory Item');
}

function showInventoryForm3_(existing,title) {
  const html = HtmlService.createHtmlOutput(inventoryFormHtml3_(existing))
    .setWidth(540).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html,title);
}

function inventoryFormHtml3_(existing) {
  const data = JSON.stringify(existing || null).replace(/</g,'\\u003c');
  return `<!doctype html><html><head><base target="_top"><style>
  body{font-family:Arial;padding:14px;color:#1F2937}label{display:block;font-weight:700;margin-top:8px}
  input,textarea{width:100%;box-sizing:border-box;padding:8px;margin-top:3px;border:1px solid #B7C9D6;border-radius:4px}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}.actions{display:flex;gap:8px;margin-top:16px}
  button{padding:10px 16px;border:0;border-radius:4px;background:#1F4E78;color:#fff;font-weight:700}
  .secondary{background:#6B7280}</style></head><body>
  <form id="f"><input type="hidden" name="row"><div class="grid">
  <div><label>Purchase date</label><input type="date" name="purchaseDate" required></div>
  <div><label>Description</label><input name="description" required></div>
  <div><label>SKU</label><input name="sku"></div><div><label>Barcode</label><input name="barcode"></div>
  <div><label>Category</label><input name="category"></div><div><label>Purchase location</label><input name="purchaseLocation"></div>
  <div><label>Storage location</label><input name="storageLocation"></div><div><label>Condition</label><input name="condition"></div>
  <div><label>Quantity</label><input type="number" min="1" name="quantity" value="1"></div>
  <div><label>Status</label><input name="status" value="Purchased"></div>
  <div><label>Purchase price</label><input type="number" step="0.01" min="0" name="purchasePrice"></div>
  <div><label>Tax paid</label><input type="number" step="0.01" min="0" name="taxPaid"></div>
  <div><label>Acquisition shipping</label><input type="number" step="0.01" min="0" name="acquisitionShipping"></div>
  <div><label>Expected sale price</label><input type="number" step="0.01" min="0" name="expectedSalePrice"></div>
  <div><label>Listed price</label><input type="number" step="0.01" min="0" name="listedPrice"></div>
  <div><label>Marketplace</label><input name="marketplace"></div></div>
  <label>Receipt link</label><input type="url" name="receiptLink"><label>Photo link</label><input type="url" name="photoLink">
  <label>Notes</label><textarea name="notes" rows="3"></textarea>
  <div class="actions"><button type="submit">Save Item</button><button type="button" class="secondary" onclick="google.script.host.close()">Cancel</button></div>
  </form><script>
  const old=${data};
  if(old){const v=old.values;const m={row:old.row,purchaseDate:v[1] instanceof String?v[1]:'',description:v[2],sku:v[3],
  barcode:v[4],category:v[5],purchaseLocation:v[6],storageLocation:v[7],condition:v[8],quantity:v[9],
  purchasePrice:v[10],taxPaid:v[11],acquisitionShipping:v[12],expectedSalePrice:v[14],listedPrice:v[15],
  marketplace:v[16],status:v[18],receiptLink:v[22],photoLink:v[23],notes:v[24]};
  Object.keys(m).forEach(k=>{const e=document.querySelector('[name="'+k+'"]');if(e)e.value=m[k]??'';});}
  document.getElementById('f').addEventListener('submit',e=>{e.preventDefault();
  google.script.run.withSuccessHandler(()=>google.script.host.close()).withFailureHandler(x=>alert(x.message))
  .saveInventoryItem(Object.fromEntries(new FormData(e.target).entries()));});</script></body></html>`;
}

function findInventoryItem() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Find Inventory Item','Enter Item ID, SKU, barcode, or description:',ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const q = res.getResponseText().trim().toLowerCase();
  if (!q) return;
  const s = sheet3_(FTP3.SHEETS.INVENTORY);
  if (s.getLastRow() < 2) return ui.alert('No inventory items found.');
  const rows = s.getRange(2,1,s.getLastRow()-1,5).getDisplayValues();
  const i = rows.findIndex(r => [r[0],r[2],r[3],r[4]].some(v => String(v).toLowerCase().includes(q)));
  if (i < 0) return ui.alert('No matching item found.');
  SpreadsheetApp.getActive().setActiveSheet(s);
  s.setActiveRange(s.getRange(i+2,1,1,FTP3.INVENTORY_HEADERS.length));
}

function showSlowInventory() {
  const s = sheet3_(FTP3.SHEETS.INVENTORY);
  SpreadsheetApp.getActive().setActiveSheet(s);
  if (!s.getFilter()) s.getRange(1,1,FTP3.ROWS+1,FTP3.INVENTORY_HEADERS.length).createFilter();
  s.getFilter().setColumnFilterCriteria(20,
    SpreadsheetApp.newFilterCriteria().whenNumberGreaterThanOrEqualTo(90).build());
}


function buildPackagingSprint3_() {
  const aliases={
    'Category':['Type'], 'Item Name':['Description'], 'Unit of Measure':[],
    'Units Purchased':[], 'Purchase Cost':[], 'Cost Per Unit':[],
    'Quantity On Hand':[], 'Reorder Level':[], 'Supplier':[],
    'SKU / Barcode':['Supplier SKU'], 'Product Link':[], 'Notes':[], 'Updated At':[]
  };
  migrateSheetHeaders3_(FTP3.SHEETS.PACKAGING,FTP3.PACKAGING_HEADERS,aliases);
  const s=sheet3_(FTP3.SHEETS.PACKAGING);
  ensureSize3_(s,300,FTP3.PACKAGING_HEADERS.length);
  header3_(s.getRange(1,1,1,FTP3.PACKAGING_HEADERS.length));
  s.setFrozenRows(1);
  setValidation3_(s,2,'FTP3_PackagingTypes',299);
  const unitRule=SpreadsheetApp.newDataValidation().requireValueInList(['Each','Roll','Foot','Sheet','Metre','Bundle','Other'],true).setAllowInvalid(true).build();
  s.getRange(2,5,299,1).setDataValidation(unitRule);
  const yesNo=SpreadsheetApp.newDataValidation().requireValueInList(['Yes','No'],true).setAllowInvalid(false).build();
  s.getRange(2,14,299,1).setDataValidation(yesNo);
  s.getRange(2,7,299,2).setNumberFormat('$#,##0.000');
  s.getRange(2,6,299,1).setNumberFormat('0.000');
  s.getRange(2,9,299,2).setNumberFormat('0.000');
  s.getRange(2,16,299,1).setNumberFormat('yyyy-mm-dd hh:mm');
  if(s.getLastRow()>1){
    const active=s.getRange(2,14,s.getLastRow()-1,1); active.getValues().forEach((r,i)=>{if(!r[0])active.getCell(i+1,1).setValue('Yes');});
    const uom=s.getRange(2,5,s.getLastRow()-1,1); uom.getValues().forEach((r,i)=>{if(!r[0])uom.getCell(i+1,1).setValue('Each');});
  }
  const data=s.getRange(2,1,299,FTP3.PACKAGING_HEADERS.length);
  s.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($I2<=$J2,$A2<>"",$N2="Yes")')
      .setBackground(FTP3.COLORS.LIGHT_RED).setRanges([data]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($I2>0,$I2<=$J2*1.5,$A2<>"",$N2="Yes")')
      .setBackground(FTP3.COLORS.GOLD).setRanges([data]).build()
  ]);
  if(s.getFilter())s.getFilter().remove();
  s.getRange(1,1,Math.max(2,s.getLastRow()),FTP3.PACKAGING_HEADERS.length).createFilter();
  s.autoResizeColumns(1,FTP3.PACKAGING_HEADERS.length);
  [3,13,15].forEach(c=>s.setColumnWidth(c,c===15?240:180));
  if(SpreadsheetApp.getActive().getSheetByName(FTP3.SHEETS.SALES))refreshPackagingDropdowns3_();
}

function showPackagingForm() {
  const html=HtmlService.createHtmlOutput(`<!doctype html><html><head><base target="_top"><style>
  body{font-family:Arial;padding:14px;color:#1F2937}label{display:block;font-weight:700;margin-top:8px}
  input,select,textarea{width:100%;box-sizing:border-box;padding:8px;margin-top:3px;border:1px solid #B7C9D6;border-radius:4px}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}button{margin-top:16px;padding:10px 16px;border:0;border-radius:4px;background:#1F4E78;color:#fff;font-weight:700}
  </style></head><body><form id="f"><div class="grid">
  <div><label>Category</label><input name="category" placeholder="Box, Bubble Wrap, Tape..." required></div>
  <div><label>Item name</label><input name="itemName" placeholder="12x10x8 Shipping Box" required></div>
  <div><label>Size</label><input name="size"></div>
  <div><label>Unit of measure</label><select name="unitOfMeasure"><option>Each</option><option>Roll</option><option>Foot</option><option>Sheet</option><option>Metre</option><option>Bundle</option><option>Other</option></select></div>
  <div><label>Units purchased</label><input type="number" step="0.001" min="0.001" name="unitsPurchased" required></div>
  <div><label>Purchase cost</label><input type="number" step="0.01" min="0" name="purchaseCost" required></div>
  <div><label>Quantity on hand</label><input type="number" step="0.001" min="0" name="quantityOnHand"></div>
  <div><label>Reorder level</label><input type="number" step="0.001" min="0" name="reorderLevel"></div>
  <div><label>Supplier</label><input name="supplier"></div><div><label>SKU / Barcode</label><input name="supplierSku"></div>
  <div><label>Product link</label><input type="url" name="productLink"></div><div><label>Active</label><select name="active"><option>Yes</option><option>No</option></select></div>
  </div><label>Notes</label><textarea name="notes" rows="3"></textarea>
  <button type="submit">Add Supply</button></form><script>
  document.getElementById('f').addEventListener('submit',e=>{e.preventDefault();
  google.script.run.withSuccessHandler(()=>google.script.host.close()).withFailureHandler(x=>alert(x.message))
  .savePackaging3_(Object.fromEntries(new FormData(e.target).entries()));});</script></body></html>`).setWidth(560).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html,'Add Packaging Supply');
}

function savePackaging3_(form) {
  const units=num3_(form.unitsPurchased); const purchaseCost=num3_(form.purchaseCost);
  if(units<=0)throw new Error('Units purchased must be greater than zero.');
  if(purchaseCost<0)throw new Error('Purchase cost cannot be negative.');
  const values=[nextId3_(FTP3.SHEETS.PACKAGING,1,'PKG'),form.category||'',form.itemName||'',form.size||'',form.unitOfMeasure||'Each',units,purchaseCost,purchaseCost/units,
    form.quantityOnHand===''||form.quantityOnHand==null?units:num3_(form.quantityOnHand),num3_(form.reorderLevel),form.supplier||'',form.supplierSku||'',form.productLink||'',form.active||'Yes',form.notes||'',new Date()];
  const s=sheet3_(FTP3.SHEETS.PACKAGING);s.getRange(Math.max(s.getLastRow()+1,2),1,1,values.length).setValues([values]);refreshPackagingDropdowns3_();refreshDashboardSprint3();
}

function packagingChoices3_() {
  const s=sheet3_(FTP3.SHEETS.PACKAGING); if(s.getLastRow()<2)return [];
  return s.getRange(2,1,s.getLastRow()-1,FTP3.PACKAGING_HEADERS.length).getValues().filter(r=>r[0]&&String(r[13]||'Yes')!=='No').map(r=>({
    id:String(r[0]),category:String(r[1]||''),label:String(r[0])+' — '+String(r[2]||r[1]||'Packaging')+(r[3]?' ('+r[3]+')':''),
    unit:String(r[4]||'Each'),cost:num3_(r[7]),available:num3_(r[8])
  }));
}

function packagingItemById3_(id) {
  if(!id)return null; id=packagingIdFromSelection3_(id); const s=sheet3_(FTP3.SHEETS.PACKAGING);if(s.getLastRow()<2)return null;
  const ids=s.getRange(2,1,s.getLastRow()-1,1).getDisplayValues().flat();const index=ids.findIndex(x=>x===id);
  if(index<0)return null;const row=index+2;return{row:row,values:s.getRange(row,1,1,FTP3.PACKAGING_HEADERS.length).getValues()[0]};
}

function calculatePackagingUsage3_(form) {
  const specs=[['boxId','boxQty','Box'],['bubbleId','bubbleQty','Bubble Wrap'],['mailerId','mailerQty','Mailer'],['tapeId','tapeQty','Tape'],['otherPackagingId','otherPackagingQty','Other']];
  const usage=[];let total=0;
  specs.forEach(([idKey,qtyKey,label])=>{const selection=String(form[idKey]||'').trim();const id=packagingIdFromSelection3_(selection);const qty=num3_(form[qtyKey]);if(!id&&qty>0)throw new Error(label+' quantity was entered without selecting an item.');if(!id)return;if(qty<=0)throw new Error(label+' quantity must be greater than zero.');const item=packagingItemById3_(id);if(!item)throw new Error(label+' packaging item was not found: '+id);const available=num3_(item.values[8]);if(qty>available+0.000001)throw new Error(label+' quantity exceeds stock. Available: '+available+' '+String(item.values[4]||'units'));const cost=num3_(item.values[7])*qty;usage.push({id:id,qty:qty,row:item.row,cost:cost});total+=cost;});
  return{usage:usage,total:Math.round(total*100)/100};
}

function deductPackagingUsage3_(usage) {
  const s=sheet3_(FTP3.SHEETS.PACKAGING);usage.forEach(u=>{const current=num3_(s.getRange(u.row,9).getValue());if(u.qty>current+0.000001)throw new Error('Packaging stock changed before the sale was saved. Reopen the sale form.');s.getRange(u.row,9).setValue(current-u.qty);s.getRange(u.row,16).setValue(new Date());});
}

function goToPackagingInventory3_(){SpreadsheetApp.getActive().setActiveSheet(sheet3_(FTP3.SHEETS.PACKAGING));}


function packagingIdFromSelection3_(selection) {
  const text=String(selection||'').trim();
  if(!text)return '';
  const match=text.match(/^(PKG-\d+)/i);
  return match?match[1].toUpperCase():text;
}

function packagingListLabel3_(row) {
  return String(row[0])+' — '+String(row[2]||row[1]||'Packaging')+(row[3]?' ('+row[3]+')':'');
}

function packagingCategoryKey3_(category) {
  const c=String(category||'').toLowerCase();
  if(c.indexOf('box')>=0)return 'Box';
  if(c.indexOf('bubble')>=0)return 'Bubble Wrap';
  if(c.indexOf('mailer')>=0||c.indexOf('envelope')>=0)return 'Mailer';
  if(c.indexOf('tape')>=0)return 'Tape';
  return 'Other';
}

function refreshPackagingDropdowns3_() {
  const ss=SpreadsheetApp.getActive();
  const packaging=ss.getSheetByName(FTP3.SHEETS.PACKAGING);
  const sales=ss.getSheetByName(FTP3.SHEETS.SALES);
  if(!packaging||!sales)return;
  let lists=ss.getSheetByName(FTP3.SHEETS.PACKAGING_LISTS);
  if(!lists)lists=ss.insertSheet(FTP3.SHEETS.PACKAGING_LISTS);
  lists.clear();ensureSize3_(lists,Math.max(300,lists.getMaxRows()),5);
  const headers=['Boxes','Bubble Wrap','Mailers','Tape','Other Packaging'];
  lists.getRange(1,1,1,5).setValues([headers]);header3_(lists.getRange(1,1,1,5));
  const buckets={'Box':[],'Bubble Wrap':[],'Mailer':[],'Tape':[],'Other':[]};
  if(packaging.getLastRow()>1){
    packaging.getRange(2,1,packaging.getLastRow()-1,FTP3.PACKAGING_HEADERS.length).getValues().forEach(r=>{
      if(!r[0]||String(r[13]||'Yes').toLowerCase()==='no')return;
      buckets[packagingCategoryKey3_(r[1])].push(packagingListLabel3_(r));
    });
  }
  const ordered=[buckets.Box,buckets['Bubble Wrap'],buckets.Mailer,buckets.Tape,buckets.Other];
  ordered.forEach((values,i)=>{values.sort();if(values.length)lists.getRange(2,i+1,values.length,1).setValues(values.map(v=>[v]));});
  const salesColumns=[24,26,28,30,32];
  salesColumns.forEach((col,i)=>{
    const count=ordered[i].length;
    const target=sales.getRange(2,col,FTP3.ROWS,1);
    target.clearDataValidations();
    if(count){
      const source=lists.getRange(2,i+1,count,1);
      const rule=SpreadsheetApp.newDataValidation().requireValueInRange(source,true).setAllowInvalid(false).setHelpText('Choose an active packaging item from the Packaging sheet.').build();
      target.setDataValidation(rule);
    }
  });
  lists.hideSheet();
  ss.toast('Packaging dropdowns refreshed.','FlipTracker Pro',4);
}

function onEdit(e) {
  if(!e||!e.range)return;
  const sheet=e.range.getSheet();
  if(sheet.getName()===FTP3.SHEETS.PACKAGING && e.range.getRow()>1){
    const relevant=[1,2,3,4,14];
    if(relevant.indexOf(e.range.getColumn())>=0)refreshPackagingDropdowns3_();
    return;
  }
  if(sheet.getName()===FTP3.SHEETS.INVENTORY && e.range.getRow()>1 && e.range.getColumn()===19 && String(e.value)==='Sold'){
    const row=e.range.getRow();
    const itemId=sheet.getRange(row,1).getDisplayValue();
    if(!itemId)return;
    if(saleExistsForInventory3_(itemId)){
      SpreadsheetApp.getActive().toast('A Sales record already exists for '+itemId+'.','FlipTracker Pro',6);
      return;
    }
    const previous=e.oldValue||'Listed';
    PropertiesService.getDocumentProperties().setProperty('FTP_PREV_STATUS_'+itemId,previous);
    e.range.setValue('Sale Pending');
    e.range.setNote('Sale details are incomplete. Keep this row selected, then use FlipTracker Pro → Complete Selected Pending Sale.');
    sheet.setActiveRange(sheet.getRange(row,1,1,FTP3.INVENTORY_HEADERS.length));
    SpreadsheetApp.getActive().toast(
      'Sale Pending created for '+itemId+'. Use FlipTracker Pro → Complete Selected Pending Sale.',
      'Complete the sale',
      10
    );
  }
}


function buildSalesSprint3_() {
  migrateSheetHeaders3_(FTP3.SHEETS.SALES,FTP3.SALES_HEADERS,{});
  const s=sheet3_(FTP3.SHEETS.SALES);ensureSize3_(s,FTP3.ROWS+1,FTP3.SALES_HEADERS.length);
  s.getRange(1,1,1,FTP3.SALES_HEADERS.length).setValues([FTP3.SALES_HEADERS]);header3_(s.getRange(1,1,1,FTP3.SALES_HEADERS.length));s.setFrozenRows(1);
  setValidation3_(s,5,'FTP3_Marketplaces',FTP3.ROWS);s.getRange(2,4,FTP3.ROWS,1).setNumberFormat('yyyy-mm-dd');s.getRange(2,6,FTP3.ROWS,13).setNumberFormat('$#,##0.00;[Red]-$#,##0.00');s.getRange(2,19,FTP3.ROWS,1).setNumberFormat('0.0%;[Red]-0.0%');s.getRange(2,24,FTP3.ROWS,1).setNumberFormat('yyyy-mm-dd hh:mm');
  s.getRange(2,26,FTP3.ROWS,1).setNumberFormat('0.000');s.getRange(2,28,FTP3.ROWS,1).setNumberFormat('0.000');s.getRange(2,30,FTP3.ROWS,1).setNumberFormat('0.000');s.getRange(2,32,FTP3.ROWS,1).setNumberFormat('0.000');s.getRange(2,34,FTP3.ROWS,1).setNumberFormat('0.000');
  if(s.getFilter())s.getFilter().remove();s.getRange(1,1,FTP3.ROWS+1,FTP3.SALES_HEADERS.length).createFilter();borders3_(s.getRange(1,1,Math.min(FTP3.ROWS+1,200),FTP3.SALES_HEADERS.length));
  if(SpreadsheetApp.getActive().getSheetByName(FTP3.SHEETS.PACKAGING))refreshPackagingDropdowns3_();
}

function showRecordSaleForm() {showRecordSaleFormForItem3_('');}

function completeSelectedPendingSale3_() {
  const ss=SpreadsheetApp.getActive();
  const range=ss.getActiveRange();
  const sheet=range && range.getSheet();
  if(!range || !sheet || sheet.getName()!==FTP3.SHEETS.INVENTORY || range.getRow()<2){
    SpreadsheetApp.getUi().alert('Select the Sale Pending item row on the Inventory sheet first.');
    return;
  }
  const row=range.getRow();
  const itemId=sheet.getRange(row,1).getDisplayValue();
  const status=sheet.getRange(row,19).getDisplayValue();
  if(!itemId){
    SpreadsheetApp.getUi().alert('The selected Inventory row does not have an Item ID.');
    return;
  }
  if(saleExistsForInventory3_(itemId)){
    sheet.getRange(row,19).setValue('Sold').clearNote();
    SpreadsheetApp.getUi().alert('A Sales record already exists for '+itemId+'. The Inventory status was set to Sold.');
    return;
  }
  if(status!=='Sale Pending'){
    const response=SpreadsheetApp.getUi().alert(
      'Start sale?',
      'The selected item is currently "'+status+'". Change it to Sale Pending and open the sale form?',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    if(response!==SpreadsheetApp.getUi().Button.YES)return;
    PropertiesService.getDocumentProperties().setProperty('FTP_PREV_STATUS_'+itemId,status||'Listed');
    sheet.getRange(row,19).setValue('Sale Pending');
  }
  showRecordSaleFormForItem3_(itemId);
}
function showRecordSaleFormForItem3_(itemId) {
  const items=activeInventoryChoices3_();
  if(itemId && !items.some(x=>x.id===itemId)){
    const item=inventoryItemById3_(itemId);
    if(item)items.unshift({id:itemId,description:String(item.values[2]||'')});
  }
  if(!items.length){SpreadsheetApp.getUi().alert('There are no active inventory items available to sell.');return;}
  const html=HtmlService.createHtmlOutput(saleFormHtml3_(items,packagingChoices3_(),marketplaceChoices3_(),itemId)).setWidth(680).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html,itemId?'Complete Sale':'Record Sale');
}
function saleFormHtml3_(items,packages,marketplaces,selectedItemId) {
  const esc=v=>String(v==null?'':v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
  const itemOptions=items.map(x=>`<option value="${esc(x.id)}" data-description="${esc(x.description||'')}" ${x.id===selectedItemId?'selected':''}>${esc(x.id)}</option>`).join('');
  const marketplaceOptions='<option value="">Select marketplace</option>'+marketplaces.map(x=>`<option value="${esc(x)}">${esc(x)}</option>`).join('');
  const makeOptions=(category)=>'<option value="">None</option>'+packages.filter(x=>category==='Other'||x.category.toLowerCase().indexOf(category.toLowerCase())>=0).map(x=>`<option value="${esc(x.id)}" data-cost="${Number(x.cost)||0}" data-stock="${Number(x.available)||0}" data-unit="${esc(x.unit)}">${esc(x.label)} — ${Number(x.available)||0} ${esc(x.unit)} @ $${(Number(x.cost)||0).toFixed(3)}</option>`).join('');
  return `<!doctype html><html><head><base target="_top"><style>
  body{font-family:Arial;padding:14px;color:#1F2937}label{display:block;font-weight:700;margin-top:8px}input,select,textarea{width:100%;box-sizing:border-box;padding:8px;margin-top:3px;border:1px solid #B7C9D6;border-radius:4px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}.section{margin-top:14px;padding:10px;background:#F3F6F9;border-radius:6px}.summary{font-weight:700;margin-top:10px}.actions{display:flex;align-items:center;gap:8px;margin-top:16px}.actions button{padding:10px 16px;border:0;border-radius:4px;background:#1F4E78;color:#fff;font-weight:700;cursor:pointer}.actions button:disabled{opacity:.55;cursor:not-allowed}.actions .cancel{background:#6B7280}.status{margin-top:10px;min-height:20px;font-weight:700}.error{color:#B91C1C}.success{color:#166534}</style></head><body><form id="f" novalidate>
  <div class="grid"><div><label>Item ID</label><select name="itemId" required>${itemOptions}</select></div><div><label>Description</label><input name="description" readonly></div></div><div class="grid"><div><label>Sale date</label><input type="date" name="saleDate" required></div><div><label>Marketplace</label><select name="marketplace" required>${marketplaceOptions}</select></div><div><label>Sale price</label><input type="number" step="0.01" min="0" name="salePrice" required></div><div><label>Shipping charged</label><input type="number" step="0.01" min="0" name="shippingCharged"></div><div><label>Shipping actual</label><input type="number" step="0.01" min="0" name="shippingActual"></div><div><label>Marketplace fees</label><input type="number" step="0.01" min="0" name="marketplaceFees"></div><div><label>Payment fees</label><input type="number" step="0.01" min="0" name="paymentFees"></div><div><label>Promotion expense</label><input type="number" step="0.01" min="0" name="promotionExpense"></div><div><label>GST/HST collected</label><input type="number" step="0.01" min="0" name="taxCollected"></div></div>
  <div class="section"><b>Packaging used</b><div class="grid"><div><label>Box used</label><select name="boxId" class="pkg">${makeOptions('Box')}</select></div><div><label>Box quantity</label><input name="boxQty" class="qty" type="number" min="0" step="0.001" value="0"></div><div><label>Bubble wrap used</label><select name="bubbleId" class="pkg">${makeOptions('Bubble')}</select></div><div><label>Bubble wrap quantity</label><input name="bubbleQty" class="qty" type="number" min="0" step="0.001" value="0"></div><div><label>Mailer used</label><select name="mailerId" class="pkg">${makeOptions('Mailer')}</select></div><div><label>Mailer quantity</label><input name="mailerQty" class="qty" type="number" min="0" step="0.001" value="0"></div><div><label>Tape used</label><select name="tapeId" class="pkg">${makeOptions('Tape')}</select></div><div><label>Tape quantity</label><input name="tapeQty" class="qty" type="number" min="0" step="0.001" value="0"></div><div><label>Other packaging</label><select name="otherPackagingId" class="pkg">${makeOptions('Other')}</select></div><div><label>Other quantity</label><input name="otherPackagingQty" class="qty" type="number" min="0" step="0.001" value="0"></div></div><div class="summary">Calculated packaging cost: <span id="pkgCost">$0.00</span></div></div>
  <div class="grid"><div><label>Buyer</label><input name="buyer"></div><div><label>Tracking number</label><input name="trackingNumber"></div></div><label>Notes</label><textarea name="notes" rows="3"></textarea>
  <div class="actions"><button id="saveBtn" type="button" onclick="submitSale()">Accept Sale</button><button id="cancelBtn" class="cancel" type="button" onclick="cancelSale()">Cancel</button></div><div id="status" class="status"></div></form><script>
  const form=document.getElementById('f'),saveBtn=document.getElementById('saveBtn'),cancelBtn=document.getElementById('cancelBtn'),statusBox=document.getElementById('status');
  document.querySelector('[name="saleDate"]').value=new Date().toISOString().slice(0,10);
  function showStatus(message,isError){statusBox.textContent=message||'';statusBox.className='status '+(isError?'error':'success');}
  function setBusy(busy){saveBtn.disabled=busy;cancelBtn.disabled=busy;saveBtn.textContent=busy?'Saving…':'Accept Sale';}
  function value(name){const el=form.elements[name];return el?String(el.value||'').trim():'';}
  function formDataObject(){const data={};Array.from(form.elements).forEach(el=>{if(el.name)data[el.name]=el.value;});return data;}
  function validate(){
    if(!value('itemId'))return 'Select an inventory item.';
    if(!value('saleDate'))return 'Enter the sale date.';
    if(!value('marketplace'))return 'Select the marketplace.';
    if(value('salePrice')==='')return 'Enter the sale price.';
    if(Number(value('salePrice'))<0)return 'Sale price cannot be negative.';
    const pairs=[['boxId','boxQty','box'],['bubbleId','bubbleQty','bubble wrap'],['mailerId','mailerQty','mailer'],['tapeId','tapeQty','tape'],['otherPackagingId','otherPackagingQty','other packaging']];
    for(const [id,qty,label] of pairs){const selected=value(id),amount=Number(value(qty)||0);if(amount<0)return 'Packaging quantities cannot be negative.';if(amount>0&&!selected)return 'Select the '+label+' item used, or set its quantity to zero.';}
    return '';
  }
  function submitSale(){
    showStatus('',false);const problem=validate();if(problem){showStatus(problem,true);return;}
    setBusy(true);showStatus('Saving sale…',false);
    let finished=false;
    const watchdog=setTimeout(()=>{
      if(finished)return;
      setBusy(false);
      showStatus('The save is taking too long. Check the Sales sheet before trying again because the sale may already have been recorded.',true);
    },20000);
    google.script.run
      .withSuccessHandler(result=>{finished=true;clearTimeout(watchdog);showStatus((result&&result.message)||'Sale saved successfully.',false);setTimeout(()=>google.script.host.close(),500);})
      .withFailureHandler(error=>{finished=true;clearTimeout(watchdog);setBusy(false);const message=error&&error.message?error.message:String(error||'The sale could not be saved.');showStatus(message,true);})
      .saveSale3_(formDataObject());
  }
  function cancelSale(){setBusy(true);const id=value('itemId');google.script.run.withSuccessHandler(()=>google.script.host.close()).withFailureHandler(error=>{setBusy(false);showStatus(error&&error.message?error.message:String(error),true);}).cancelPendingSale3_(id);}
  function syncDescription(){
    const sel=form.elements.itemId;
    const opt=sel.options[sel.selectedIndex];
    form.elements.description.value=opt?String(opt.dataset.description||''):'';
  }
  form.elements.itemId.addEventListener('change',syncDescription);
  function calc(){let total=0;[['boxId','boxQty'],['bubbleId','bubbleQty'],['mailerId','mailerQty'],['tapeId','tapeQty'],['otherPackagingId','otherPackagingQty']].forEach(([s,q])=>{const sel=form.elements[s],opt=sel.options[sel.selectedIndex];total+=(Number(opt&&opt.dataset.cost)||0)*(Number(form.elements[q].value)||0);});document.getElementById('pkgCost').textContent='$'+total.toFixed(2);}
  document.querySelectorAll('.pkg,.qty').forEach(x=>{x.addEventListener('input',calc);x.addEventListener('change',calc);});
  form.addEventListener('submit',e=>{e.preventDefault();submitSale();});
  syncDescription();
  calc();
  </script></body></html>`;
}

function marketplaceChoices3_() {
  const ss=SpreadsheetApp.getActive();
  const named=ss.getRangeByName('FTP3_Marketplaces');
  if(!named)return ['eBay','Facebook Marketplace','Kijiji','Poshmark','Etsy','Local Sale','Other'];
  const values=named.getDisplayValues().flat().map(v=>String(v).trim()).filter(Boolean);
  return values.filter((v,i,a)=>a.indexOf(v)===i);
}

function saveSale3_(form) {
  const lock=LockService.getDocumentLock();if(!lock.tryLock(5000))throw new Error('FlipTracker Pro is busy with another update. Please wait a few seconds and try again.');
  try{
    const item=inventoryItemById3_(form.itemId);if(!item)throw new Error('The selected inventory item was not found.');const inv=item.values;if(saleExistsForInventory3_(form.itemId))throw new Error('A completed Sales record already exists for this inventory item.');
    const pkg=calculatePackagingUsage3_(form);const salePrice=num3_(form.salePrice),shippingCharged=num3_(form.shippingCharged),shippingActual=num3_(form.shippingActual),marketFees=num3_(form.marketplaceFees),paymentFees=num3_(form.paymentFees),promotion=num3_(form.promotionExpense),taxCollected=num3_(form.taxCollected),itemCost=num3_(inv[13]);
    if(!form.saleDate)throw new Error('Sale date is required.');if(salePrice<0)throw new Error('Sale price cannot be negative.');
    const grossRevenue=salePrice+shippingCharged,sellingCosts=shippingActual+pkg.total+marketFees+paymentFees+promotion,netProceeds=grossRevenue-sellingCosts,realizedProfit=netProceeds-itemCost,roi=itemCost?realizedProfit/itemCost:'';
    const saleDate=date3_(form.saleDate),purchaseDate=inv[1] instanceof Date?inv[1]:date3_(inv[1]),days=purchaseDate?Math.max(0,Math.floor((saleDate-purchaseDate)/86400000)):'';
    const description=String(inv[2]||form.description||'');
    const values=[nextId3_(FTP3.SHEETS.SALES,1,'SAL'),form.itemId,description,saleDate,form.marketplace||'',salePrice,shippingCharged,shippingActual,pkg.total,marketFees,paymentFees,promotion,taxCollected,itemCost,grossRevenue,sellingCosts,netProceeds,realizedProfit,roi,days,form.buyer||'',form.trackingNumber||'',form.notes||'',new Date(),packagingIdFromSelection3_(form.boxId),num3_(form.boxQty),packagingIdFromSelection3_(form.bubbleId),num3_(form.bubbleQty),packagingIdFromSelection3_(form.mailerId),num3_(form.mailerQty),packagingIdFromSelection3_(form.tapeId),num3_(form.tapeQty),packagingIdFromSelection3_(form.otherPackagingId),num3_(form.otherPackagingQty),'Yes'];
    const sales=sheet3_(FTP3.SHEETS.SALES),row=Math.max(sales.getLastRow()+1,2);sales.getRange(row,1,1,values.length).setValues([values]);
    try{deductPackagingUsage3_(pkg.usage);}catch(err){sales.deleteRow(row);throw err;}
    const inventory=sheet3_(FTP3.SHEETS.INVENTORY);inventory.getRange(item.row,19).setValue('Sold').clearNote();inventory.getRange(item.row,27).setValue(new Date());PropertiesService.getDocumentProperties().deleteProperty('FTP_PREV_STATUS_'+form.itemId);SpreadsheetApp.flush();
    return {ok:true,message:'Sale '+values[0]+' saved successfully.'};
  }finally{lock.releaseLock();}
}

function saleExistsForInventory3_(itemId){
  const s=sheet3_(FTP3.SHEETS.SALES);
  if(s.getLastRow()<2)return false;
  return s.getRange(2,2,s.getLastRow()-1,1).getDisplayValues().flat().some(v=>v===itemId);
}
function cancelPendingSale3_(itemId){
  const item=inventoryItemById3_(itemId);
  if(item && String(item.values[18])==='Sale Pending'){
    const previous=PropertiesService.getDocumentProperties().getProperty('FTP_PREV_STATUS_'+itemId)||'Listed';
    sheet3_(FTP3.SHEETS.INVENTORY).getRange(item.row,19).setValue(previous).clearNote();
    PropertiesService.getDocumentProperties().deleteProperty('FTP_PREV_STATUS_'+itemId);
  }
}


function buildExpensesSprint3_() {
  const s = sheet3_(FTP3.SHEETS.EXPENSES);
  ensureSize3_(s,FTP3.ROWS+1,FTP3.EXPENSE_HEADERS.length);
  s.getRange(1,1,1,FTP3.EXPENSE_HEADERS.length).setValues([FTP3.EXPENSE_HEADERS]);
  header3_(s.getRange(1,1,1,FTP3.EXPENSE_HEADERS.length));
  s.setFrozenRows(1);
  setValidation3_(s,3,'FTP3_ExpenseCategories',FTP3.ROWS);
  setValidation3_(s,11,'FTP3_PaymentMethods',FTP3.ROWS);
  s.getRange(2,2,FTP3.ROWS,1).setNumberFormat('yyyy-mm-dd');
  s.getRange(2,6,FTP3.ROWS,4).setNumberFormat('$#,##0.00');
  s.getRange(2,9,FTP3.ROWS,1).setNumberFormat('0.0%');
  s.getRange(2,10,FTP3.ROWS,1).setNumberFormat('$#,##0.00');
  s.getRange(2,14,FTP3.ROWS,1).setNumberFormat('yyyy-mm-dd hh:mm');
  if (s.getFilter()) s.getFilter().remove();
  s.getRange(1,1,FTP3.ROWS+1,FTP3.EXPENSE_HEADERS.length).createFilter();
}

function showRecordExpenseForm() {
  const html = HtmlService.createHtmlOutput(`<!doctype html><html><head><base target="_top"><style>
  body{font-family:Arial;padding:14px;color:#1F2937}label{display:block;font-weight:700;margin-top:8px}
  input,textarea{width:100%;box-sizing:border-box;padding:8px;margin-top:3px;border:1px solid #B7C9D6;border-radius:4px}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}button{margin-top:16px;padding:10px 16px;border:0;border-radius:4px;background:#1F4E78;color:#fff;font-weight:700}
  </style></head><body><form id="f"><div class="grid">
  <div><label>Date</label><input type="date" name="date" required></div><div><label>Category</label><input name="category" required></div>
  <div><label>Vendor</label><input name="vendor"></div><div><label>Payment method</label><input name="paymentMethod"></div>
  <div><label>Subtotal</label><input type="number" step="0.01" min="0" name="subtotal"></div>
  <div><label>GST/HST paid</label><input type="number" step="0.01" min="0" name="taxPaid"></div>
  <div><label>Business use %</label><input type="number" step="0.1" min="0" max="100" name="businessUse" value="100"></div>
  <div><label>Receipt link</label><input type="url" name="receiptLink"></div></div>
  <label>Description</label><input name="description" required><label>Notes</label><textarea name="notes" rows="3"></textarea>
  <button type="submit">Record Expense</button></form><script>
  document.querySelector('[name="date"]').value=new Date().toISOString().slice(0,10);
  document.getElementById('f').addEventListener('submit',e=>{e.preventDefault();
  google.script.run.withSuccessHandler(()=>google.script.host.close()).withFailureHandler(x=>alert(x.message))
  .saveExpense3_(Object.fromEntries(new FormData(e.target).entries()));});</script></body></html>`)
  .setWidth(520).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html,'Record Expense');
}

function saveExpense3_(form) {
  if (!form.date || !form.category || !form.description) throw new Error('Date, category, and description are required.');
  const subtotal = num3_(form.subtotal);
  const tax = num3_(form.taxPaid);
  const total = subtotal + tax;
  const percent = Math.min(100,Math.max(0,num3_(form.businessUse))) / 100;
  const deductible = total * percent;
  const values = [
    nextId3_(FTP3.SHEETS.EXPENSES,1,'EXP'),date3_(form.date),form.category||'',
    form.vendor||'',form.description||'',subtotal,tax,total,percent,deductible,
    form.paymentMethod||'',form.receiptLink||'',form.notes||'',new Date()
  ];
  const s = sheet3_(FTP3.SHEETS.EXPENSES);
  s.getRange(Math.max(s.getLastRow()+1,2),1,1,values.length).setValues([values]);
  refreshDashboardSprint3();
}


function buildMileageSprint3_() {
  const s = sheet3_(FTP3.SHEETS.MILEAGE);
  ensureSize3_(s,FTP3.ROWS+1,FTP3.MILEAGE_HEADERS.length);
  s.getRange(1,1,1,FTP3.MILEAGE_HEADERS.length).setValues([FTP3.MILEAGE_HEADERS]);
  header3_(s.getRange(1,1,1,FTP3.MILEAGE_HEADERS.length));
  s.setFrozenRows(1);
  s.getRange(2,2,FTP3.ROWS,1).setNumberFormat('yyyy-mm-dd');
  s.getRange(2,10,FTP3.ROWS,2).setNumberFormat('$#,##0.00');
  s.getRange(2,13,FTP3.ROWS,1).setNumberFormat('yyyy-mm-dd hh:mm');
  if (s.getFilter()) s.getFilter().remove();
  s.getRange(1,1,FTP3.ROWS+1,FTP3.MILEAGE_HEADERS.length).createFilter();
}

function showRecordMileageForm() {
  const html = HtmlService.createHtmlOutput(`<!doctype html><html><head><base target="_top"><style>
  body{font-family:Arial;padding:14px;color:#1F2937}label{display:block;font-weight:700;margin-top:8px}
  input,textarea{width:100%;box-sizing:border-box;padding:8px;margin-top:3px;border:1px solid #B7C9D6;border-radius:4px}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}button{margin-top:16px;padding:10px 16px;border:0;border-radius:4px;background:#1F4E78;color:#fff;font-weight:700}
  </style></head><body><form id="f"><div class="grid">
  <div><label>Date</label><input type="date" name="date" required></div><div><label>CRA rate</label><input type="number" step="0.001" min="0" name="craRate"></div>
  <div><label>Start</label><input name="start"></div><div><label>Destination</label><input name="destination"></div>
  <div><label>Odometer start</label><input type="number" step="0.1" min="0" name="odoStart"></div>
  <div><label>Odometer end</label><input type="number" step="0.1" min="0" name="odoEnd"></div>
  <div><label>Business kilometres</label><input type="number" step="0.1" min="0" name="businessKm"></div></div>
  <label>Business purpose</label><input name="purpose" required><label>Notes</label><textarea name="notes" rows="3"></textarea>
  <button type="submit">Record Mileage</button></form><script>
  document.querySelector('[name="date"]').value=new Date().toISOString().slice(0,10);
  document.getElementById('f').addEventListener('submit',e=>{e.preventDefault();
  google.script.run.withSuccessHandler(()=>google.script.host.close()).withFailureHandler(x=>alert(x.message))
  .saveMileage3_(Object.fromEntries(new FormData(e.target).entries()));});</script></body></html>`)
  .setWidth(520).setHeight(590);
  SpreadsheetApp.getUi().showModalDialog(html,'Record Mileage');
}

function saveMileage3_(form) {
  if (!form.date || !form.purpose) throw new Error('Date and business purpose are required.');
  const start = num3_(form.odoStart);
  const end = num3_(form.odoEnd);
  const total = Math.max(0,end-start);
  const requestedBusiness = form.businessKm === '' || form.businessKm == null ? total : num3_(form.businessKm);
  const business = Math.min(total,Math.max(0,requestedBusiness));
  const configuredRate = num3_(getSettingV04_('CRA Mileage Rate',0));
  const rate = form.craRate === '' || form.craRate == null ? configuredRate : num3_(form.craRate);
  const claim = business * rate;
  const values = [
    nextId3_(FTP3.SHEETS.MILEAGE,1,'TRP'),date3_(form.date),form.start||'',
    form.destination||'',form.purpose||'',start,end,total,business,rate,claim,
    form.notes||'',new Date()
  ];
  const s = sheet3_(FTP3.SHEETS.MILEAGE);
  s.getRange(Math.max(s.getLastRow()+1,2),1,1,values.length).setValues([values]);
  refreshDashboardSprint3();
}


function buildDashboardSprint3_() {
  const s = sheet3_(FTP3.SHEETS.DASHBOARD);
  s.clear();
  s.setHiddenGridlines(true);
  s.getRange('A1:H2').merge().setValue('FlipTracker Pro — v0.4')
    .setBackground(FTP3.COLORS.NAVY).setFontColor(FTP3.COLORS.WHITE)
    .setFontSize(22).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  const cards = [
    ['A4:B6','Current Inventory Cost','SUMIFS(Inventory!N2:N,Inventory!S2:S,"<>Sold",Inventory!S2:S,"<>Archived")','$#,##0.00'],
    ['C4:D6','Gross Sales','SUM(Sales!O2:O)','$#,##0.00'],
    ['E4:F6','Realized Profit','SUM(Sales!R2:R)','$#,##0.00'],
    ['G4:H6','Average ROI','IFERROR(AVERAGE(FILTER(Sales!S2:S,Sales!A2:A<>"")),0)','0.0%'],
    ['A8:B10','Items Sold','COUNTA(Sales!A2:A)','0'],
    ['C8:D10','Business Expenses','SUM(Expenses!J2:J)','$#,##0.00'],
    ['E8:F10','Mileage Claims','SUM(Mileage!K2:K)','$#,##0.00'],
    ['G8:H10','90+ Day Items','COUNTIFS(Inventory!T2:T,">=90",Inventory!S2:S,"<>Sold",Inventory!C2:C,"<>")','0']
  ];

  cards.forEach(([a1,title,formula,fmt]) => {
    const r = s.getRange(a1);
    r.merge();
    r.getCell(1,1).setFormula(`="${title}"&CHAR(10)&TEXT(${formula},"${fmt}")`);
    r.setBackground(FTP3.COLORS.LIGHT_BLUE).setFontColor(FTP3.COLORS.TEXT)
      .setFontWeight('bold').setFontSize(14).setWrap(true)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,FTP3.COLORS.BORDER,
        SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });

  s.getRange('A13:H13').merge().setValue('v0.4 Transaction and Tax Workflow');
  header3_(s.getRange('A13:H13'));
  s.getRange('A14:H18').merge().setValue(
    'Record a sale from the FlipTracker Pro menu. Sprint 3 automatically retrieves ' +
    'the item cost, calculates gross revenue, total selling costs, net proceeds, ' +
    'realized profit, ROI, and days to sell, then marks the inventory item Sold. ' +
    'Expenses, mileage, packaging supplies, and tax-year estimates are tracked in dedicated sheets. Open the CRA Tax Centre from the menu for COGS, inventory, GST/HST, and accountant summaries.'
  ).setBackground(FTP3.COLORS.GRAY).setWrap(true).setVerticalAlignment('middle');

  s.getRange('A20:D20').merge().setValue('Profit after operating expenses');
  s.getRange('E20:H20').merge().setFormula('=SUM(Sales!R2:R)-SUM(Expenses!J2:J)-IF(IFERROR(VLOOKUP("Include Mileage Estimate in Expenses",Settings!A:B,2,FALSE),"No")="Yes",SUM(Mileage!K2:K),0)');
  header3_(s.getRange('A20:D20'));
  s.getRange('E20:H20').setBackground(FTP3.COLORS.LIGHT_GREEN)
    .setFontWeight('bold').setFontSize(16).setNumberFormat('$#,##0.00')
    .setHorizontalAlignment('center');

  s.getRange('A23:H23').merge().setValue('Packaging Inventory');header3_(s.getRange('A23:H23'));
  const pkgCards=[
    ['A24:B26','Packaging Inventory Value','SUMPRODUCT(Packaging!H2:H,Packaging!I2:I)','$#,##0.00'],
    ['C24:D26','Low-Stock Supplies','COUNTIFS(Packaging!A2:A,"<>",Packaging!N2:N,"Yes",Packaging!I2:I,"<="&Packaging!J2:J)','0'],
    ['E24:F26','Packaging Cost This Year','SUMIFS(Sales!I2:I,Sales!D2:D,">="&DATE(YEAR(TODAY()),1,1),Sales!D2:D,"<"&DATE(YEAR(TODAY())+1,1,1))','$#,##0.00'],
    ['G24:H26','Average Packaging / Sale','IFERROR(AVERAGE(FILTER(Sales!I2:I,Sales!A2:A<>"")),0)','$#,##0.00']
  ];
  pkgCards.forEach(([a1,title,formula,fmt])=>{const r=s.getRange(a1);r.merge();r.getCell(1,1).setFormula(`="${title}"&CHAR(10)&TEXT(${formula},"${fmt}")`);r.setBackground(FTP3.COLORS.GOLD).setFontColor(FTP3.COLORS.TEXT).setFontWeight('bold').setFontSize(13).setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle').setBorder(true,true,true,true,false,false,FTP3.COLORS.BORDER,SpreadsheetApp.BorderStyle.SOLID_MEDIUM);});

  s.setColumnWidths(1,8,120);
  s.setFrozenRows(2);
}

function refreshDashboardSprint3() {
  buildDashboardSprint3_();
}

function goToDashboardSprint3() {
  SpreadsheetApp.getActive().setActiveSheet(sheet3_(FTP3.SHEETS.DASHBOARD));
}


function buildTaxCentreV04_() {
  const s = sheet3_(FTP3.SHEETS.TAX_CENTRE);
  s.clear();
  s.setHiddenGridlines(true);
  ensureSize3_(s,80,10);
  s.getRange('A1:J2').merge().setValue('FlipTracker Pro — CRA Tax Centre')
    .setBackground(FTP3.COLORS.NAVY).setFontColor(FTP3.COLORS.WHITE)
    .setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');

  const year = Math.floor(num3_(getSettingV04_('Tax Year',new Date().getFullYear())));
  const calc = calculateTaxCentreV041_(year);
  s.getRange('A4:B4').setValues([['Selected Tax Year',year]]);
  s.getRange('A4').setFontWeight('bold');
  s.getRange('B4').setNumberFormat('0');
  s.getRange('D4:J4').merge().setValue('Bookkeeping estimates only — review with a qualified tax professional before filing.')
    .setBackground(FTP3.COLORS.GOLD).setFontWeight('bold').setWrap(true);

  const labels = [
    ['A6:D6','Income and Gross Profit'],['A7:C7','Gross sales and shipping revenue'],
    ['A8:C8','GST/HST collected'],['A9:C9','Revenue excluding GST/HST'],
    ['A10:C10','Opening inventory'],['A11:C11','Inventory purchases during year'],
    ['A12:C12','Ending inventory at cost'],['A13:C13','Estimated cost of goods sold'],
    ['A14:C14','Estimated gross profit'],['F6:J6','Operating Expenses'],
    ['F7:I7','Deductible expenses entered'],['F8:I8','Eligible GST/HST ITCs — operating expenses'],
    ['F9:I9','Income-tax expense deduction'],['F10:I10','Mileage estimate'],
    ['F11:I11','Total operating deductions estimate'],['F12:I12','Estimated net business income'],
    ['A17:D17','GST/HST Regular-Method Estimate'],['A18:C18','GST/HST collected'],
    ['A19:C19','Potential ITCs — inventory purchases'],['A20:C20','Potential ITCs — operating expenses'],
    ['A21:C21','Estimated net GST/HST'],['F17:J17','Year-end Inventory'],
    ['F18:I18','Active item count at year-end'],['F19:I19','Ending inventory at recorded cost'],
    ['F20:I20','Valuation method']
  ];
  labels.forEach(x=>{const r=s.getRange(x[0]);r.merge();r.setValue(x[1]);});
  ['A6:D6','F6:J6','A17:D17','F17:J17'].forEach(a1=>header3_(s.getRange(a1)));

  s.getRange('D7:D14').setValues([[calc.grossRevenue],[calc.taxCollected],[calc.revenueExTax],
    [calc.openingInventory],[calc.purchases],[calc.endingInventory],[calc.cogs],[calc.grossProfit]]);
  s.getRange('J7:J12').setValues([[calc.expenseEntered],[calc.expenseItc],[calc.incomeTaxExpense],
    [calc.mileageDeduction],[calc.operatingDeductions],[calc.netBusinessIncome]]);
  s.getRange('D18:D21').setValues([[calc.taxCollected],[calc.inventoryItc],[calc.expenseItc],[calc.netTax]]);
  s.getRange('J18:J20').setValues([[calc.activeItemCount],[calc.endingInventory],[calc.valuationMethod]]);

  s.getRange('A24:J24').merge().setValue('Expense Category Summary'); header3_(s.getRange('A24:J24'));
  s.getRange('A25:C25').setValues([['Category','Income-tax deduction','Eligible GST/HST ITC']]); header3_(s.getRange('A25:C25'));
  const summary = calc.expenseSummary.length ? calc.expenseSummary : [['No expenses',0,0]];
  s.getRange(26,1,summary.length,3).setValues(summary);

  s.getRange('F25:J25').setValues([['Metric','Amount','CRA Reference','','']]); header3_(s.getRange('F25:J25'));
  const refs=[
    ['Inventory and COGS','','https://www.canada.ca/en/revenue-agency/services/tax/businesses/small-businesses-self-employed-income/business-income-tax-reporting/inventory-cost-goods-sold.html','',''],
    ['T2125','','https://www.canada.ca/en/revenue-agency/services/forms-publications/forms/t2125.html','',''],
    ['GST/HST records','','https://www.canada.ca/en/revenue-agency/services/tax/businesses/topics/keeping-records/gst-hst-payroll-records.html','',''],
    ['Input tax credits','','https://www.canada.ca/en/revenue-agency/services/tax/businesses/topics/gst-hst-businesses/calculate-prepare-report/input-tax-credit.html','',''],
    ['Motor vehicle records','','https://www.canada.ca/en/revenue-agency/services/tax/businesses/topics/sole-proprietorships-partnerships/report-business-income-expenses/completing-form-t2125/calculating-motor-vehicle-expenses.html','','']
  ];
  s.getRange(26,6,refs.length,5).setValues(refs);
  s.getRange('D7:D21').setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
  s.getRange('J7:J19').setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
  s.getRange('J18').setNumberFormat('0'); s.getRange('J20').setNumberFormat('@');
  s.getRange('B26:C60').setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
  borders3_(s.getRange('A4:J31'));
  s.setColumnWidths(1,10,120);s.setColumnWidth(1,190);s.setColumnWidth(6,190);s.setColumnWidths(8,3,150);
  s.setFrozenRows(2);
}

function calculateTaxCentreV041_(year) {
  const start = new Date(year,0,1), end = new Date(year,11,31,23,59,59,999);
  const registered = String(getSettingV04_('GST/HST Registered','No')).toLowerCase()==='yes';
  const method = String(getSettingV04_('GST/HST Reporting Method','Regular'));
  const regular = registered && method.toLowerCase()==='regular';
  const includeMileage = String(getSettingV04_('Include Mileage Estimate in Expenses','No')).toLowerCase()==='yes';
  const openingInventory = num3_(getSettingV04_('Opening Inventory Value',0));
  const valuationMethod = String(getSettingV04_('Inventory Valuation Method','Lower of cost and FMV'));

  const salesSheet=sheet3_(FTP3.SHEETS.SALES), invSheet=sheet3_(FTP3.SHEETS.INVENTORY);
  const expSheet=sheet3_(FTP3.SHEETS.EXPENSES), milSheet=sheet3_(FTP3.SHEETS.MILEAGE);
  const sales=salesSheet.getLastRow()>1?salesSheet.getRange(2,1,salesSheet.getLastRow()-1,FTP3.SALES_HEADERS.length).getValues():[];
  const inventory=invSheet.getLastRow()>1?invSheet.getRange(2,1,invSheet.getLastRow()-1,FTP3.INVENTORY_HEADERS.length).getValues():[];
  const expenses=expSheet.getLastRow()>1?expSheet.getRange(2,1,expSheet.getLastRow()-1,FTP3.EXPENSE_HEADERS.length).getValues():[];
  const mileage=milSheet.getLastRow()>1?milSheet.getRange(2,1,milSheet.getLastRow()-1,FTP3.MILEAGE_HEADERS.length).getValues():[];
  const inYear=d=>d instanceof Date && !isNaN(d) && d>=start && d<=end;

  const yearSales=sales.filter(r=>r[0]&&inYear(r[3]));
  const grossRevenue=yearSales.reduce((n,r)=>n+num3_(r[14]),0);
  const taxCollected=yearSales.reduce((n,r)=>n+num3_(r[12]),0);
  const revenueExTax=grossRevenue; // Sales Gross Revenue deliberately excludes separately entered GST/HST.
  const inventoryById={}; inventory.filter(r=>r[0]).forEach(r=>inventoryById[String(r[0])]=r);
  const specificCogs=yearSales.reduce((n,r)=>{
    const source=inventoryById[String(r[1])];
    const recoverableTax=regular&&source?num3_(source[11]):0;
    return n+Math.max(0,num3_(r[13])-recoverableTax);
  },0);
  const yearPurchases=inventory.filter(r=>r[0]&&inYear(r[1]));
  const purchases=yearPurchases.reduce((n,r)=>n+Math.max(0,num3_(r[13])-(regular?num3_(r[11]):0)),0);
  const inventoryItc=regular?yearPurchases.reduce((n,r)=>n+num3_(r[11]),0):0;

  const soldByEnd={};
  sales.filter(r=>r[0]&&r[1]&&r[3] instanceof Date&&!isNaN(r[3])&&r[3]<=end)
    .forEach(r=>{soldByEnd[String(r[1])]=true;});
  const activeAtEnd=inventory.filter(r=>r[0]&&r[1] instanceof Date&&!isNaN(r[1])&&r[1]<=end&&!soldByEnd[String(r[0])]);
  const endingInventory=activeAtEnd.reduce((n,r)=>n+Math.max(0,num3_(r[13])-(regular?num3_(r[11]):0)),0);
  const cogs=specificCogs || Math.max(0,openingInventory+purchases-endingInventory);
  const grossProfit=revenueExTax-cogs;

  let expenseEntered=0, expenseItc=0, incomeTaxExpense=0;
  const byCategory={};
  expenses.filter(r=>r[0]&&inYear(r[1])).forEach(r=>{
    const businessPct=Math.min(1,Math.max(0,num3_(r[8])));
    const entered=num3_(r[9]);
    const itc=regular?num3_(r[6])*businessPct:0;
    const deduction=Math.max(0,entered-itc);
    expenseEntered+=entered;expenseItc+=itc;incomeTaxExpense+=deduction;
    const cat=String(r[2]||'Uncategorized');
    if(!byCategory[cat])byCategory[cat]=[0,0];
    byCategory[cat][0]+=deduction;byCategory[cat][1]+=itc;
  });
  const mileageDeduction=includeMileage?mileage.filter(r=>r[0]&&inYear(r[1])).reduce((n,r)=>n+num3_(r[10]),0):0;
  const operatingDeductions=incomeTaxExpense+mileageDeduction;
  const netBusinessIncome=grossProfit-operatingDeductions;
  const netTax=regular?taxCollected-inventoryItc-expenseItc:0;
  return {year,grossRevenue,taxCollected,revenueExTax,openingInventory,purchases,endingInventory,
    cogs,grossProfit,expenseEntered,expenseItc,incomeTaxExpense,mileageDeduction,
    operatingDeductions,netBusinessIncome,inventoryItc,netTax,activeItemCount:activeAtEnd.length,
    valuationMethod,expenseSummary:Object.keys(byCategory).sort().map(k=>[k,byCategory[k][0],byCategory[k][1]])};
}

function getSettingV04_(name,fallback) {
  const s=sheet3_(FTP3.SHEETS.SETTINGS);
  if(s.getLastRow()<2)return fallback;
  const vals=s.getRange(2,1,s.getLastRow()-1,2).getValues();
  const row=vals.find(r=>String(r[0])===name);
  return row&&row[1]!==''?row[1]:fallback;
}
function refreshTaxCentreV04(){buildTaxCentreV04_();}
function goToTaxCentreV04(){SpreadsheetApp.getActive().setActiveSheet(sheet3_(FTP3.SHEETS.TAX_CENTRE));}

function buildAccountantExportV04_() {
  buildTaxCentreV04_();
  const s=sheet3_(FTP3.SHEETS.ACCOUNTANT_EXPORT);s.clear();s.setHiddenGridlines(true);
  const year=getSettingV04_('Tax Year',new Date().getFullYear());
  const values=[
    ['FlipTracker Pro Accountant Export','',''],['Tax Year',year,''],
    ['Business Name',getSettingV04_('Business Name',''),''],['Business Number',getSettingV04_('Business Number',''),''],
    ['GST/HST Number',getSettingV04_('GST/HST Number',''),''],['','',''],['Summary','Amount','Source'],
    ['Gross revenue including shipping',"='Tax Centre'!D7",'Sales'],['GST/HST collected',"='Tax Centre'!D8",'Sales'],
    ['Revenue excluding GST/HST',"='Tax Centre'!D9",'Calculated'],['Opening inventory',"='Tax Centre'!D10",'Settings'],
    ['Purchases during year',"='Tax Centre'!D11",'Inventory'],['Ending inventory',"='Tax Centre'!D12",'Inventory'],
    ['Estimated COGS',"='Tax Centre'!D13",'Sales / Inventory'],['Gross profit',"='Tax Centre'!D14",'Calculated'],
    ['Operating deductions',"='Tax Centre'!J11",'Expenses / Mileage'],['Estimated net business income',"='Tax Centre'!J12",'Calculated'],
    ['Estimated net GST/HST',"='Tax Centre'!D21",'Regular-method estimate']
  ];
  s.getRange(1,1,values.length,3).setValues(values);
  values.forEach((r,i)=>{if(typeof r[1]==='string'&&r[1].startsWith('='))s.getRange(i+1,2).setFormula(r[1]);});
  s.getRange('A1:C1').merge();header3_(s.getRange('A1:C1'));header3_(s.getRange('A7:C7'));
  s.getRange('B8:B18').setNumberFormat('$#,##0.00;[Red]-$#,##0.00');borders3_(s.getRange('A1:C18'));
  s.setColumnWidth(1,260);s.setColumnWidth(2,150);s.setColumnWidth(3,180);s.setFrozenRows(7);
  SpreadsheetApp.getActive().setActiveSheet(s);
  SpreadsheetApp.getActive().toast('Accountant Export refreshed for '+year+'.','FlipTracker Pro',6);
}


function runFlipTrackerCalculationAudit() {
  const issues=[];
  const add=(severity,area,row,message)=>issues.push([severity,area,row||'',message]);
  const inv=sheet3_(FTP3.SHEETS.INVENTORY), sales=sheet3_(FTP3.SHEETS.SALES);
  const exp=sheet3_(FTP3.SHEETS.EXPENSES), mil=sheet3_(FTP3.SHEETS.MILEAGE), pkg=sheet3_(FTP3.SHEETS.PACKAGING);
  const invRows=inv.getLastRow()>1?inv.getRange(2,1,inv.getLastRow()-1,FTP3.INVENTORY_HEADERS.length).getValues():[];
  invRows.forEach((r,i)=>{
    if(!r[0])return; const expected=num3_(r[10])*Math.max(1,num3_(r[9]))+num3_(r[11])+num3_(r[12]);
    if(Math.abs(num3_(r[13])-expected)>0.01)add('HIGH','Inventory',i+2,'Total Cost does not equal Purchase Price × Quantity + Tax + Acquisition Shipping.');
  });
  const saleRows=sales.getLastRow()>1?sales.getRange(2,1,sales.getLastRow()-1,FTP3.SALES_HEADERS.length).getValues():[];
  saleRows.forEach((r,i)=>{if(!r[0])return;
    const gross=num3_(r[5])+num3_(r[6]); const costs=num3_(r[7])+num3_(r[8])+num3_(r[9])+num3_(r[10])+num3_(r[11]);
    const net=gross-costs, profit=net-num3_(r[13]);
    if(Math.abs(num3_(r[14])-gross)>0.01)add('CRITICAL','Sales',i+2,'Gross Revenue is inconsistent.');
    if(Math.abs(num3_(r[15])-costs)>0.01)add('CRITICAL','Sales',i+2,'Total Selling Costs incorrectly includes or omits a cost.');
    if(Math.abs(num3_(r[16])-net)>0.01)add('CRITICAL','Sales',i+2,'Net Proceeds is inconsistent.');
    if(Math.abs(num3_(r[17])-profit)>0.01)add('CRITICAL','Sales',i+2,'Realized Profit is inconsistent.');
  });
  const expRows=exp.getLastRow()>1?exp.getRange(2,1,exp.getLastRow()-1,FTP3.EXPENSE_HEADERS.length).getValues():[];
  expRows.forEach((r,i)=>{if(!r[0])return;const total=num3_(r[5])+num3_(r[6]);const ded=total*Math.min(1,Math.max(0,num3_(r[8])));
    if(Math.abs(num3_(r[7])-total)>0.01)add('HIGH','Expenses',i+2,'Total does not equal Subtotal + GST/HST Paid.');
    if(Math.abs(num3_(r[9])-ded)>0.01)add('HIGH','Expenses',i+2,'Deductible Amount does not equal Total × Business Use %.');
  });
  const milRows=mil.getLastRow()>1?mil.getRange(2,1,mil.getLastRow()-1,FTP3.MILEAGE_HEADERS.length).getValues():[];
  milRows.forEach((r,i)=>{if(!r[0])return;const total=Math.max(0,num3_(r[6])-num3_(r[5]));
    if(Math.abs(num3_(r[7])-total)>0.01)add('HIGH','Mileage',i+2,'Total Kilometres is inconsistent with odometers.');
    if(num3_(r[8])>total+0.001)add('HIGH','Mileage',i+2,'Business Kilometres exceeds Total Kilometres.');
    if(Math.abs(num3_(r[10])-num3_(r[8])*num3_(r[9]))>0.01)add('HIGH','Mileage',i+2,'Claim Amount does not equal Business Kilometres × CRA Rate.');
  });
  const pkgRows=pkg.getLastRow()>1?pkg.getRange(2,1,pkg.getLastRow()-1,FTP3.PACKAGING_HEADERS.length).getValues():[];
  pkgRows.forEach((r,i)=>{if(!r[0])return;const cpu=num3_(r[5])?num3_(r[6])/num3_(r[5]):0;
    if(Math.abs(num3_(r[7])-cpu)>0.001)add('MEDIUM','Packaging',i+2,'Cost Per Unit is inconsistent.');
    if(num3_(r[8])<-0.000001)add('CRITICAL','Packaging',i+2,'Quantity On Hand is negative.');
  });
  saleRows.forEach((r,i)=>{if(!r[0])return;const refs=[[r[24],r[25]],[r[26],r[27]],[r[28],r[29]],[r[30],r[31]],[r[32],r[33]]];let calculated=0;refs.forEach(([id,qty])=>{if(!id)return;const item=packagingItemById3_(String(id));if(!item)add('HIGH','Sales',i+2,'A packaging item referenced by this sale no longer exists: '+id);else calculated+=num3_(item.values[7])*num3_(qty);});if(String(r[34]||'')==='Yes'&&Math.abs(num3_(r[8])-calculated)>0.011)add('HIGH','Sales',i+2,'Stored Packaging Cost does not match current referenced unit costs. Historical costs may have changed.');});
  const q=sheet3_('Calculation Audit');q.clear();q.getRange(1,1,1,4).setValues([['Severity','Area','Row','Finding']]);header3_(q.getRange(1,1,1,4));
  if(issues.length)q.getRange(2,1,issues.length,4).setValues(issues);else q.getRange(2,1).setValue('PASS — no row-level calculation inconsistencies detected.');
  q.setFrozenRows(1);q.setColumnWidth(1,100);q.setColumnWidth(2,120);q.setColumnWidth(3,70);q.setColumnWidth(4,520);q.getRange('A:D').setWrap(true);
  SpreadsheetApp.getActive().setActiveSheet(q);SpreadsheetApp.getActive().toast(issues.length+' calculation issue(s) found.','FlipTracker Pro Audit',8);
  return issues;
}


function onOpen() {
  SpreadsheetApp.getUi().createMenu('FlipTracker Pro')
    .addItem('Install / Repair v0.4.8','installFlipTrackerPro')
    .addItem('Upgrade Existing Workbook','upgradeFlipTrackerPro')
    .addSeparator()
    .addItem('Add Inventory Item','showAddItemForm')
    .addItem('Edit Selected Inventory Item','showEditSelectedItemForm')
    .addItem('Find Inventory Item','findInventoryItem')
    .addItem('Show Slow Inventory','showSlowInventory')
    .addSeparator()
    .addItem('Complete Selected Pending Sale','completeSelectedPendingSale3_')
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


function installFlipTrackerPro() {
  buildAdminSprint3_(); buildSettingsSprint3_(); buildInventorySprint3_();
  buildSalesSprint3_(); buildExpensesSprint3_(); buildMileageSprint3_();
  buildPackagingSprint3_(); buildDashboardSprint3_(); buildTaxCentreV04_();
  const p=PropertiesService.getDocumentProperties();
  p.setProperty('FTP_SCHEMA_VERSION','4.8'); p.setProperty('FTP_APP_VERSION',FTP3.VERSION);
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
  if(current<4.3) migrateToSchema43_();
  if(current<4.4) migrateToSchema44_();
  if(current<4.5) migrateToSchema45_();
  if(current<4.6) migrateToSchema46_();
  if(current<4.7) migrateToSchema47_();
  if(current<4.8) migrateToSchema48_();
  p.setProperty('FTP_SCHEMA_VERSION','4.8'); p.setProperty('FTP_APP_VERSION',FTP3.VERSION);
  SpreadsheetApp.getActive().toast('FlipTracker Pro upgraded to schema 4.4.','FlipTracker Pro',6);
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
    inventory.getRange(2,1,inventory.getLastRow()-1,3).getDisplayValues()
      .forEach(r=>{if(r[0])map[String(r[0])]=String(r[2]||'');});
  }
  const ids=sales.getRange(2,2,sales.getLastRow()-1,1).getDisplayValues();
  const descriptions=ids.map(r=>[map[String(r[0])]||'']);
  sales.getRange(2,3,descriptions.length,1).setValues(descriptions);
}
function getFlipTrackerVersion(){const p=PropertiesService.getDocumentProperties();return{appVersion:p.getProperty('FTP_APP_VERSION')||FTP3.VERSION,schemaVersion:p.getProperty('FTP_SCHEMA_VERSION')||'unversioned'};}


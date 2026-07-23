// FlipTracker Pro v0.4.0 — CRA Tax Centre

// ===== Config.gs =====
const FTP3 = {
  APP_NAME: 'FlipTracker Pro',
  VERSION: '0.4.0 CRA Tax Centre',
  ROWS: 1000,
  SHEETS: {
    DASHBOARD:'Dashboard', INVENTORY:'Inventory', SALES:'Sales',
    EXPENSES:'Expenses', MILEAGE:'Mileage', PACKAGING:'Packaging',
    ADMIN:'Admin', SETTINGS:'Settings', TAX_CENTRE:'Tax Centre', ACCOUNTANT_EXPORT:'Accountant Export'
  },
  COLORS: {
    NAVY:'#1F4E78', BLUE:'#5B9BD5', LIGHT_BLUE:'#D9EAF7',
    GREEN:'#70AD47', LIGHT_GREEN:'#E2F0D9', GOLD:'#FFF2CC',
    ORANGE:'#FCE5CD', LIGHT_RED:'#FCE4D6', GRAY:'#F3F6F9',
    BORDER:'#B7C9D6', WHITE:'#FFFFFF', TEXT:'#1F2937'
  },
  INVENTORY_HEADERS: [
    'Item ID','Purchase Date','Title','SKU','Barcode','Category',
    'Purchase Location','Storage Location','Condition','Quantity',
    'Purchase Price','Tax Paid','Acquisition Shipping','Total Cost',
    'Expected Sale Price','Listed Price','Marketplace','Listing Date',
    'Status','Days in Inventory','Projected Profit','Projected ROI %',
    'Receipt Link','Photo Link','Notes','Created At','Updated At'
  ],
  SALES_HEADERS: [
    'Sale ID','Item ID','Sale Date','Marketplace','Sale Price',
    'Shipping Charged','Shipping Actual','Packaging Cost',
    'Marketplace Fees','Payment Fees','Promotion Expense',
    'GST/HST Collected','Item Cost','Gross Revenue','Total Selling Costs',
    'Net Proceeds','Realized Profit','Realized ROI %','Days to Sell',
    'Buyer','Tracking Number','Notes','Created At'
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
    'Packaging ID','Type','Description','Size','Units Purchased',
    'Purchase Cost','Cost Per Unit','Quantity On Hand','Reorder Level',
    'Supplier','Product Link','Notes','Updated At'
  ]
};

// ===== Utilities.gs =====
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
    .map(r => ({id:r[0], label:r[0] + ' — ' + r[2]}));
}

// ===== AdminSettings.gs =====
function buildAdminSprint3_() {
  const s = sheet3_(FTP3.SHEETS.ADMIN);
  s.clear();
  const lists = {
    Categories:['Electronics','Tools','Collectibles','Clothing','Furniture','Automotive','Household','Other'],
    PurchaseLocations:['Garage Sale','Thrift Store','Value Village','Facebook Marketplace','Auction','Retail Clearance','Other'],
    Marketplaces:['eBay','Facebook Marketplace','Kijiji','Poshmark','Etsy','Local Sale','Other'],
    Statuses:['Purchased','Needs Cleaning','Needs Testing','Ready to List','Listed','Sold','Shipped','Archived'],
    Conditions:['New','Open Box','Like New','Good','Fair','For Parts'],
    StorageLocations:['Garage','Basement','Shelf A','Shelf B','Bin 1','Bin 2','Other'],
    ExpenseCategories:['Fuel','Packaging','Shipping Supplies','Advertising','Storage','Phone','Internet','Office','Software','Bank Fees','Equipment','Professional Fees','Other'],
    PaymentMethods:['Cash','Credit Card','Debit','PayPal','Bank Transfer','Other'],
    PackagingTypes:['Box','Bubble Wrap','Mailer','Tape','Label','Packing Paper','Other']
  };
  let col = 1;
  Object.keys(lists).forEach(name => {
    s.getRange(1,col).setValue(name);
    const values = lists[name].map(v => [v]);
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

// ===== Inventory.gs =====
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
  const total = purchasePrice + tax + shipping;
  const profit = expected ? expected-total : '';
  const roi = total && profit !== '' ? profit/total : '';
  const purchaseDate = date3_(form.purchaseDate);
  const days = Math.max(0,Math.floor((new Date()-purchaseDate)/86400000));
  const now = new Date();

  const values = [
    editing ? old[0] : nextId3_(FTP3.SHEETS.INVENTORY,1,'ITM'),
    purchaseDate,form.title,form.sku||'',form.barcode||'',form.category||'',
    form.purchaseLocation||'',form.storageLocation||'',form.condition||'',
    num3_(form.quantity)||1,purchasePrice,tax,shipping,total,expected,
    num3_(form.listedPrice),form.marketplace||'',old[17]||'',
    form.status||'Purchased',days,profit,roi,form.receiptLink||'',
    form.photoLink||'',form.notes||'',editing?(old[25]||now):now,now
  ];
  s.getRange(row,1,1,values.length).setValues([values]);
  refreshDashboardSprint3();
}

// ===== InventoryForms.gs =====
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
  <div><label>Title</label><input name="title" required></div>
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
  if(old){const v=old.values;const m={row:old.row,purchaseDate:v[1] instanceof String?v[1]:'',title:v[2],sku:v[3],
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
  const res = ui.prompt('Find Inventory Item','Enter Item ID, SKU, barcode, or title:',ui.ButtonSet.OK_CANCEL);
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

// ===== Sales.gs =====
function buildSalesSprint3_() {
  const s = sheet3_(FTP3.SHEETS.SALES);
  ensureSize3_(s,FTP3.ROWS+1,FTP3.SALES_HEADERS.length);
  s.getRange(1,1,1,FTP3.SALES_HEADERS.length).setValues([FTP3.SALES_HEADERS]);
  header3_(s.getRange(1,1,1,FTP3.SALES_HEADERS.length));
  s.setFrozenRows(1);
  setValidation3_(s,4,'FTP3_Marketplaces',FTP3.ROWS);
  s.getRange(2,3,FTP3.ROWS,1).setNumberFormat('yyyy-mm-dd');
  s.getRange(2,5,FTP3.ROWS,13).setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
  s.getRange(2,18,FTP3.ROWS,1).setNumberFormat('0.0%;[Red]-0.0%');
  s.getRange(2,23,FTP3.ROWS,1).setNumberFormat('yyyy-mm-dd hh:mm');
  if (s.getFilter()) s.getFilter().remove();
  s.getRange(1,1,FTP3.ROWS+1,FTP3.SALES_HEADERS.length).createFilter();
  borders3_(s.getRange(1,1,Math.min(FTP3.ROWS+1,200),FTP3.SALES_HEADERS.length));
}

function showRecordSaleForm() {
  const items = activeInventoryChoices3_();
  if (!items.length) {
    SpreadsheetApp.getUi().alert('There are no active inventory items available to sell.');
    return;
  }
  const html = HtmlService.createHtmlOutput(saleFormHtml3_(items)).setWidth(540).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html,'Record Sale');
}

function saleFormHtml3_(items) {
  const options = items.map(x => `<option value="${x.id}">${x.label}</option>`).join('');
  return `<!doctype html><html><head><base target="_top"><style>
  body{font-family:Arial;padding:14px;color:#1F2937}label{display:block;font-weight:700;margin-top:8px}
  input,select,textarea{width:100%;box-sizing:border-box;padding:8px;margin-top:3px;border:1px solid #B7C9D6;border-radius:4px}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}button{margin-top:16px;padding:10px 16px;border:0;border-radius:4px;background:#1F4E78;color:#fff;font-weight:700}
  </style></head><body><form id="f">
  <label>Inventory item</label><select name="itemId" required>${options}</select><div class="grid">
  <div><label>Sale date</label><input type="date" name="saleDate" required></div>
  <div><label>Marketplace</label><input name="marketplace" required></div>
  <div><label>Sale price</label><input type="number" step="0.01" min="0" name="salePrice" required></div>
  <div><label>Shipping charged</label><input type="number" step="0.01" min="0" name="shippingCharged"></div>
  <div><label>Shipping actual</label><input type="number" step="0.01" min="0" name="shippingActual"></div>
  <div><label>Packaging cost</label><input type="number" step="0.01" min="0" name="packagingCost"></div>
  <div><label>Marketplace fees</label><input type="number" step="0.01" min="0" name="marketplaceFees"></div>
  <div><label>Payment fees</label><input type="number" step="0.01" min="0" name="paymentFees"></div>
  <div><label>Promotion expense</label><input type="number" step="0.01" min="0" name="promotionExpense"></div>
  <div><label>GST/HST collected</label><input type="number" step="0.01" min="0" name="taxCollected"></div>
  <div><label>Buyer</label><input name="buyer"></div><div><label>Tracking number</label><input name="trackingNumber"></div>
  </div><label>Notes</label><textarea name="notes" rows="3"></textarea>
  <button type="submit">Record Sale</button></form><script>
  document.querySelector('[name="saleDate"]').value=new Date().toISOString().slice(0,10);
  document.getElementById('f').addEventListener('submit',e=>{e.preventDefault();
  google.script.run.withSuccessHandler(()=>google.script.host.close()).withFailureHandler(x=>alert(x.message))
  .saveSale3_(Object.fromEntries(new FormData(e.target).entries()));});</script></body></html>`;
}

function saveSale3_(form) {
  const item = inventoryItemById3_(form.itemId);
  if (!item) throw new Error('The selected inventory item was not found.');
  const inv = item.values;
  const salePrice = num3_(form.salePrice);
  const shippingCharged = num3_(form.shippingCharged);
  const shippingActual = num3_(form.shippingActual);
  const packaging = num3_(form.packagingCost);
  const marketFees = num3_(form.marketplaceFees);
  const paymentFees = num3_(form.paymentFees);
  const promotion = num3_(form.promotionExpense);
  const taxCollected = num3_(form.taxCollected);
  const itemCost = num3_(inv[13]);
  const grossRevenue = salePrice + shippingCharged;
  const sellingCosts = shippingActual + packaging + marketFees + paymentFees + promotion + taxCollected;
  const netProceeds = grossRevenue - sellingCosts;
  const realizedProfit = netProceeds - itemCost;
  const roi = itemCost ? realizedProfit / itemCost : '';
  const saleDate = date3_(form.saleDate);
  const purchaseDate = inv[1] instanceof Date ? inv[1] : date3_(inv[1]);
  const days = purchaseDate ? Math.max(0,Math.floor((saleDate-purchaseDate)/86400000)) : '';

  const values = [
    nextId3_(FTP3.SHEETS.SALES,1,'SAL'),form.itemId,saleDate,form.marketplace||'',
    salePrice,shippingCharged,shippingActual,packaging,marketFees,paymentFees,
    promotion,taxCollected,itemCost,grossRevenue,sellingCosts,netProceeds,
    realizedProfit,roi,days,form.buyer||'',form.trackingNumber||'',form.notes||'',new Date()
  ];
  const sales = sheet3_(FTP3.SHEETS.SALES);
  const row = Math.max(sales.getLastRow()+1,2);
  sales.getRange(row,1,1,values.length).setValues([values]);

  const inventory = sheet3_(FTP3.SHEETS.INVENTORY);
  inventory.getRange(item.row,19).setValue('Sold');
  inventory.getRange(item.row,27).setValue(new Date());
  refreshDashboardSprint3();
}

// ===== Expenses.gs =====
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

// ===== Mileage.gs =====
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
  const start = num3_(form.odoStart);
  const end = num3_(form.odoEnd);
  const total = Math.max(0,end-start);
  const business = num3_(form.businessKm) || total;
  const rate = num3_(form.craRate);
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

// ===== Packaging.gs =====
function buildPackagingSprint3_() {
  const s = sheet3_(FTP3.SHEETS.PACKAGING);
  ensureSize3_(s,300,FTP3.PACKAGING_HEADERS.length);
  s.getRange(1,1,1,FTP3.PACKAGING_HEADERS.length).setValues([FTP3.PACKAGING_HEADERS]);
  header3_(s.getRange(1,1,1,FTP3.PACKAGING_HEADERS.length));
  s.setFrozenRows(1);
  setValidation3_(s,2,'FTP3_PackagingTypes',299);
  s.getRange(2,6,299,2).setNumberFormat('$#,##0.000');
  s.getRange(2,13,299,1).setNumberFormat('yyyy-mm-dd hh:mm');
  const data = s.getRange(2,1,299,FTP3.PACKAGING_HEADERS.length);
  s.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($H2<=$I2,$A2<>"")')
      .setBackground(FTP3.COLORS.LIGHT_RED).setRanges([data]).build()
  ]);
}

function showPackagingForm() {
  const html = HtmlService.createHtmlOutput(`<!doctype html><html><head><base target="_top"><style>
  body{font-family:Arial;padding:14px;color:#1F2937}label{display:block;font-weight:700;margin-top:8px}
  input,textarea{width:100%;box-sizing:border-box;padding:8px;margin-top:3px;border:1px solid #B7C9D6;border-radius:4px}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}button{margin-top:16px;padding:10px 16px;border:0;border-radius:4px;background:#1F4E78;color:#fff;font-weight:700}
  </style></head><body><form id="f"><div class="grid">
  <div><label>Type</label><input name="type" required></div><div><label>Size</label><input name="size"></div>
  <div><label>Units purchased</label><input type="number" step="1" min="1" name="unitsPurchased" required></div>
  <div><label>Purchase cost</label><input type="number" step="0.01" min="0" name="purchaseCost" required></div>
  <div><label>Quantity on hand</label><input type="number" step="1" min="0" name="quantityOnHand"></div>
  <div><label>Reorder level</label><input type="number" step="1" min="0" name="reorderLevel"></div>
  <div><label>Supplier</label><input name="supplier"></div><div><label>Product link</label><input type="url" name="productLink"></div>
  </div><label>Description</label><input name="description"><label>Notes</label><textarea name="notes" rows="3"></textarea>
  <button type="submit">Add Supply</button></form><script>
  document.getElementById('f').addEventListener('submit',e=>{e.preventDefault();
  google.script.run.withSuccessHandler(()=>google.script.host.close()).withFailureHandler(x=>alert(x.message))
  .savePackaging3_(Object.fromEntries(new FormData(e.target).entries()));});</script></body></html>`)
  .setWidth(520).setHeight(610);
  SpreadsheetApp.getUi().showModalDialog(html,'Add Packaging Supply');
}

function savePackaging3_(form) {
  const units = Math.max(1,num3_(form.unitsPurchased));
  const purchaseCost = num3_(form.purchaseCost);
  const values = [
    nextId3_(FTP3.SHEETS.PACKAGING,1,'PKG'),form.type||'',form.description||'',
    form.size||'',units,purchaseCost,purchaseCost/units,
    num3_(form.quantityOnHand),num3_(form.reorderLevel),form.supplier||'',
    form.productLink||'',form.notes||'',new Date()
  ];
  const s = sheet3_(FTP3.SHEETS.PACKAGING);
  s.getRange(Math.max(s.getLastRow()+1,2),1,1,values.length).setValues([values]);
}

// ===== Dashboard.gs =====
function buildDashboardSprint3_() {
  const s = sheet3_(FTP3.SHEETS.DASHBOARD);
  s.clear();
  s.setHiddenGridlines(true);
  s.getRange('A1:H2').merge().setValue('FlipTracker Pro — v0.4')
    .setBackground(FTP3.COLORS.NAVY).setFontColor(FTP3.COLORS.WHITE)
    .setFontSize(22).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  const cards = [
    ['A4:B6','Inventory Cost','SUM(Inventory!N2:N)','$#,##0.00'],
    ['C4:D6','Gross Sales','SUM(Sales!N2:N)','$#,##0.00'],
    ['E4:F6','Realized Profit','SUM(Sales!Q2:Q)','$#,##0.00'],
    ['G4:H6','Average ROI','IFERROR(AVERAGE(FILTER(Sales!R2:R,Sales!A2:A<>"")),0)','0.0%'],
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
  s.getRange('E20:H20').merge().setFormula('=SUM(Sales!Q2:Q)-SUM(Expenses!J2:J)-SUM(Mileage!K2:K)');
  header3_(s.getRange('A20:D20'));
  s.getRange('E20:H20').setBackground(FTP3.COLORS.LIGHT_GREEN)
    .setFontWeight('bold').setFontSize(16).setNumberFormat('$#,##0.00')
    .setHorizontalAlignment('center');

  s.setColumnWidths(1,8,120);
  s.setFrozenRows(2);
}

function refreshDashboardSprint3() {
  buildDashboardSprint3_();
}

function goToDashboardSprint3() {
  SpreadsheetApp.getActive().setActiveSheet(sheet3_(FTP3.SHEETS.DASHBOARD));
}

// ===== TaxCentre.gs =====
function buildTaxCentreV04_() {
  const s = sheet3_(FTP3.SHEETS.TAX_CENTRE);
  s.clear();
  s.setHiddenGridlines(true);
  ensureSize3_(s,80,10);
  s.getRange('A1:J2').merge().setValue('FlipTracker Pro — CRA Tax Centre')
    .setBackground(FTP3.COLORS.NAVY).setFontColor(FTP3.COLORS.WHITE)
    .setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');

  s.getRange('A4:B4').setValues([['Selected Tax Year', getSettingV04_('Tax Year', new Date().getFullYear())]]);
  s.getRange('A4').setFontWeight('bold');
  s.getRange('B4').setNumberFormat('0');
  s.getRange('D4:J4').merge().setValue('Bookkeeping estimates only — review with a qualified tax professional before filing.')
    .setBackground(FTP3.COLORS.GOLD).setFontWeight('bold').setWrap(true);

  const labels = [
    ['A6:D6','Income and Gross Profit'],
    ['A7:C7','Gross sales and shipping revenue'],
    ['A8:C8','GST/HST collected'],
    ['A9:C9','Revenue excluding GST/HST'],
    ['A10:C10','Opening inventory'],
    ['A11:C11','Inventory purchases during year'],
    ['A12:C12','Ending inventory at cost'],
    ['A13:C13','Estimated cost of goods sold'],
    ['A14:C14','Estimated gross profit'],
    ['F6:J6','Operating Expenses'],
    ['F7:I7','Deductible expenses entered'],
    ['F8:I8','GST/HST paid on business-use portion'],
    ['F9:I9','Expense amount excluding eligible GST/HST'],
    ['F10:I10','Mileage estimate'],
    ['F11:I11','Total operating deductions estimate'],
    ['F12:I12','Estimated net business income'],
    ['A17:D17','GST/HST Regular-Method Estimate'],
    ['A18:C18','GST/HST collected'],
    ['A19:C19','Potential ITCs — inventory purchases'],
    ['A20:C20','Potential ITCs — operating expenses'],
    ['A21:C21','Estimated net GST/HST'],
    ['F17:J17','Year-end Inventory'],
    ['F18:I18','Active item count'],
    ['F19:I19','Ending inventory at recorded cost'],
    ['F20:I20','Valuation method'],
  ];
  labels.forEach(x => { const r=s.getRange(x[0]); r.merge(); r.setValue(x[1]); });
  ['A6:D6','F6:J6','A17:D17','F17:J17'].forEach(a => header3_(s.getRange(a)));

  const y='Settings!B11';
  const salesYear='YEAR(Sales!C2:C)='+y;
  const expYear='YEAR(Expenses!B2:B)='+y;
  const invPurchaseYear='YEAR(Inventory!B2:B)='+y;
  const invUnsoldByYearEnd='(Inventory!S2:S<>"Sold")+(Inventory!A2:A<>"")*(IFERROR(YEAR(Inventory!B2:B),0)<='+y+')';
  s.getRange('D7').setFormula(`=IFERROR(SUM(FILTER(Sales!N2:N,${salesYear})),0)`);
  s.getRange('D8').setFormula(`=IFERROR(SUM(FILTER(Sales!L2:L,${salesYear})),0)`);
  s.getRange('D9').setFormula('=D7-D8');
  s.getRange('D10').setFormula("=IFERROR(VLOOKUP(\"Opening Inventory Value\",Settings!A:B,2,FALSE),0)");
  s.getRange('D11').setFormula(`=IFERROR(SUM(FILTER(Inventory!N2:N,${invPurchaseYear})),0)`);
  s.getRange('D12').setFormula(`=IFERROR(SUM(FILTER(Inventory!N2:N,Inventory!A2:A<>\"\",Inventory!B2:B<=DATE(${y},12,31),((Inventory!S2:S<>\"Sold\")+(Inventory!A2:A<>\"\")*(IFERROR(YEAR(Inventory!B2:B),0)=${y})*(Inventory!S2:S=\"Sold\")*(IFERROR(YEAR(Inventory!B2:B),0)>${y}))>0)),0)`);
  // Item cost recorded on sales provides the most reliable specific-item COGS estimate.
  s.getRange('D13').setFormula(`=IFERROR(SUM(FILTER(Sales!M2:M,${salesYear})),D10+D11-D12)`);
  s.getRange('D14').setFormula('=D9-D13');

  s.getRange('J7').setFormula(`=IFERROR(SUM(FILTER(Expenses!J2:J,${expYear})),0)`);
  s.getRange('J8').setFormula(`=IFERROR(SUM(FILTER(Expenses!G2:G*Expenses!I2:I,${expYear})),0)`);
  s.getRange('J9').setFormula('=J7-J8');
  s.getRange('J10').setFormula(`=IF(IFERROR(VLOOKUP(\"Include Mileage Estimate in Expenses\",Settings!A:B,2,FALSE),\"No\")=\"Yes\",IFERROR(SUM(FILTER(Mileage!K2:K,YEAR(Mileage!B2:B)=${y})),0),0)`);
  s.getRange('J11').setFormula('=J9+J10');
  s.getRange('J12').setFormula('=D14-J11');

  s.getRange('D18').setFormula('=D8');
  s.getRange('D19').setFormula(`=IFERROR(SUM(FILTER(Inventory!L2:L,${invPurchaseYear})),0)`);
  s.getRange('D20').setFormula('=J8');
  s.getRange('D21').setFormula('=D18-D19-D20');

  s.getRange('J18').setFormula(`=IFERROR(COUNTA(FILTER(Inventory!A2:A,Inventory!A2:A<>\"\",Inventory!B2:B<=DATE(${y},12,31),Inventory!S2:S<>\"Sold\")),0)`);
  s.getRange('J19').setFormula('=D12');
  s.getRange('J20').setFormula('=IFERROR(VLOOKUP("Inventory Valuation Method",Settings!A:B,2,FALSE),"")');

  s.getRange('A24:J24').merge().setValue('Expense Category Summary'); header3_(s.getRange('A24:J24'));
  s.getRange('A25:C25').setValues([['Category','Deductible Amount','GST/HST Paid (business portion)']]); header3_(s.getRange('A25:C25'));
  s.getRange('A26').setFormula(`=IFERROR(QUERY(FILTER({Expenses!C2:C,Expenses!J2:J,Expenses!G2:G*Expenses!I2:I},YEAR(Expenses!B2:B)=${y}),\"select Col1,sum(Col2),sum(Col3) where Col1 is not null group by Col1 label sum(Col2) '',sum(Col3) ''\",0),{\"No expenses\",0,0})`);

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
  s.getRange('B26:C60').setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
  borders3_(s.getRange('A4:J31'));
  s.setColumnWidths(1,10,120); s.setColumnWidth(1,190); s.setColumnWidth(6,190); s.setColumnWidths(8,3,150);
  s.setFrozenRows(2);
}

function getSettingV04_(name, fallback) {
  const s=sheet3_(FTP3.SHEETS.SETTINGS);
  if (s.getLastRow()<2) return fallback;
  const vals=s.getRange(2,1,s.getLastRow()-1,2).getValues();
  const row=vals.find(r=>String(r[0])===name);
  return row && row[1]!=='' ? row[1] : fallback;
}

function refreshTaxCentreV04() { buildTaxCentreV04_(); }
function goToTaxCentreV04() { SpreadsheetApp.getActive().setActiveSheet(sheet3_(FTP3.SHEETS.TAX_CENTRE)); }

function buildAccountantExportV04_() {
  const s=sheet3_(FTP3.SHEETS.ACCOUNTANT_EXPORT);
  s.clear(); s.setHiddenGridlines(true);
  const year=getSettingV04_('Tax Year',new Date().getFullYear());
  const values=[
    ['FlipTracker Pro Accountant Export','',''],
    ['Tax Year',year,''],
    ['Business Name',getSettingV04_('Business Name',''),''],
    ['Business Number',getSettingV04_('Business Number',''),''],
    ['GST/HST Number',getSettingV04_('GST/HST Number',''),''],
    ['','',''],
    ['Summary','Amount','Source'],
    ['Gross revenue including shipping',"='Tax Centre'!D7",'Sales'],
    ['GST/HST collected',"='Tax Centre'!D8",'Sales'],
    ['Revenue excluding GST/HST',"='Tax Centre'!D9",'Calculated'],
    ['Opening inventory',"='Tax Centre'!D10",'Settings'],
    ['Purchases during year',"='Tax Centre'!D11",'Inventory'],
    ['Ending inventory',"='Tax Centre'!D12",'Inventory'],
    ['Estimated COGS',"='Tax Centre'!D13",'Sales / Inventory'],
    ['Gross profit',"='Tax Centre'!D14",'Calculated'],
    ['Operating deductions',"='Tax Centre'!J11",'Expenses / Mileage'],
    ['Estimated net business income',"='Tax Centre'!J12",'Calculated'],
    ['Estimated net GST/HST',"='Tax Centre'!D21",'Regular-method estimate']
  ];
  s.getRange(1,1,values.length,3).setValues(values);
  values.forEach((r,i)=>{ if(typeof r[1]==='string' && r[1].startsWith('=')) s.getRange(i+1,2).setFormula(r[1]); });
  s.getRange('A1:C1').merge(); header3_(s.getRange('A1:C1')); header3_(s.getRange('A7:C7'));
  s.getRange('B8:B18').setNumberFormat('$#,##0.00;[Red]-$#,##0.00'); borders3_(s.getRange('A1:C18'));
  s.setColumnWidth(1,260); s.setColumnWidth(2,150); s.setColumnWidth(3,180); s.setFrozenRows(7);
  SpreadsheetApp.getActive().setActiveSheet(s);
  SpreadsheetApp.getActive().toast('Accountant Export refreshed for '+year+'.','FlipTracker Pro',6);
}

// ===== Menu.gs =====
function onOpen() {
  SpreadsheetApp.getUi().createMenu('FlipTracker Pro')
    .addItem('Install / Repair v0.4','installFlipTrackerPro')
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
    .addSeparator()
    .addItem('Refresh Dashboard','refreshDashboardSprint3')
    .addItem('Open CRA Tax Centre','goToTaxCentreV04')
    .addItem('Refresh CRA Tax Centre','refreshTaxCentreV04')
    .addItem('Build Accountant Export','buildAccountantExportV04_')
    .addItem('Go to Dashboard','goToDashboardSprint3')
    .addToUi();
}

function initializeFlipTrackerProSprint3() { installFlipTrackerPro(); }

// ===== Main.gs =====
function installFlipTrackerPro() {
  buildAdminSprint3_(); buildSettingsSprint3_(); buildInventorySprint3_();
  buildSalesSprint3_(); buildExpensesSprint3_(); buildMileageSprint3_();
  buildPackagingSprint3_(); buildDashboardSprint3_(); buildTaxCentreV04_();
  const p=PropertiesService.getDocumentProperties();
  p.setProperty('FTP_SCHEMA_VERSION','4'); p.setProperty('FTP_APP_VERSION',FTP3.VERSION);
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
  p.setProperty('FTP_SCHEMA_VERSION','4'); p.setProperty('FTP_APP_VERSION',FTP3.VERSION);
  SpreadsheetApp.getActive().toast('FlipTracker Pro upgraded to schema 4.','FlipTracker Pro',6);
}
function migrateToSchema1_(){buildAdminSprint3_();buildSettingsSprint3_();buildInventorySprint3_();}
function migrateToSchema2_(){buildInventorySprint3_();}
function migrateToSchema3_(){buildSalesSprint3_();buildExpensesSprint3_();buildMileageSprint3_();buildPackagingSprint3_();buildDashboardSprint3_();}
function migrateToSchema4_(){buildSettingsSprint3_();buildTaxCentreV04_();buildDashboardSprint3_();}
function getFlipTrackerVersion(){const p=PropertiesService.getDocumentProperties();return{appVersion:p.getProperty('FTP_APP_VERSION')||FTP3.VERSION,schemaVersion:p.getProperty('FTP_SCHEMA_VERSION')||'unversioned'};}


const FTP2 = {
  APP_NAME: 'FlipTracker Pro',
  VERSION: '1.0 Sprint 2',
  ROWS: 1000,
  SHEETS: { DASHBOARD:'Dashboard', INVENTORY:'Inventory', ADMIN:'Admin' },
  COLORS: {
    NAVY:'#1F4E78', LIGHT_BLUE:'#D9EAF7', LIGHT_GREEN:'#E2F0D9',
    LIGHT_RED:'#FCE4D6', GOLD:'#FFF2CC', BORDER:'#B7C9D6',
    WHITE:'#FFFFFF', TEXT:'#1F2937'
  },
  HEADERS: [
    'Item ID','Purchase Date','Title','SKU','Barcode','Category',
    'Purchase Location','Storage Location','Condition','Quantity',
    'Purchase Price','Tax Paid','Acquisition Shipping','Total Cost',
    'Expected Sale Price','Listed Price','Marketplace','Listing Date',
    'Status','Days in Inventory','Projected Profit','Projected ROI %',
    'Receipt Link','Photo Link','Notes','Created At','Updated At'
  ]
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('FlipTracker Pro')
    .addItem('Install / Repair Sprint 2','initializeFlipTrackerProSprint2')
    .addSeparator()
    .addItem('Add Inventory Item','showAddItemForm')
    .addItem('Edit Selected Item','showEditSelectedItemForm')
    .addItem('Find Item','findInventoryItem')
    .addItem('Show Slow Inventory','showSlowInventory')
    .addSeparator()
    .addItem('Refresh Dashboard','refreshDashboardSprint2')
    .addItem('Go to Inventory','goToInventorySprint2')
    .addToUi();
}

function initializeFlipTrackerProSprint2() {
  buildAdminSprint2_();
  buildInventorySprint2_();
  buildDashboardSprint2_();
  SpreadsheetApp.getActive().toast('Sprint 2 inventory engine is ready.','FlipTracker Pro',5);
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function header2_(range) {
  range.setBackground(FTP2.COLORS.NAVY)
    .setFontColor(FTP2.COLORS.WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
}

function buildAdminSprint2_() {
  const s = sheet_(FTP2.SHEETS.ADMIN);
  const lists = {
    Categories:['Electronics','Tools','Collectibles','Clothing','Furniture','Automotive','Household','Other'],
    PurchaseLocations:['Garage Sale','Thrift Store','Value Village','Facebook Marketplace','Auction','Retail Clearance','Other'],
    Marketplaces:['eBay','Facebook Marketplace','Kijiji','Poshmark','Etsy','Local Sale','Other'],
    Statuses:['Purchased','Needs Cleaning','Needs Testing','Ready to List','Listed','Sold','Shipped','Archived'],
    Conditions:['New','Open Box','Like New','Good','Fair','For Parts'],
    StorageLocations:['Garage','Basement','Shelf A','Shelf B','Bin 1','Bin 2','Other']
  };

  let c = 1;
  Object.keys(lists).forEach(name => {
    s.getRange(1,c).setValue(name);
    const values = lists[name].map(v => [v]);
    s.getRange(2,c,values.length,1).setValues(values);

    const named = 'FTP2_' + name;
    SpreadsheetApp.getActive().getNamedRanges()
      .filter(n => n.getName() === named)
      .forEach(n => n.remove());

    SpreadsheetApp.getActive().setNamedRange(named,s.getRange(2,c,50,1));
    s.setColumnWidth(c,180);
    c++;
  });

  header2_(s.getRange(1,1,1,c-1));
  s.setFrozenRows(1);
}

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

function showAddItemForm() {
  const template = HtmlService.createTemplate(itemFormHtml_());
  template.itemJson = 'null';
  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(520).setHeight(690),
    'Add Inventory Item'
  );
}

function showEditSelectedItemForm() {
  const s = SpreadsheetApp.getActiveSheet();

  if (s.getName() !== FTP2.SHEETS.INVENTORY || s.getActiveCell().getRow() < 2) {
    SpreadsheetApp.getUi().alert('Select an inventory row first.');
    return;
  }

  const row = s.getActiveCell().getRow();
  const item = s.getRange(row,1,1,FTP2.HEADERS.length).getValues()[0];
  const template = HtmlService.createTemplate(itemFormHtml_());
  template.itemJson = JSON.stringify({row:row,values:item});

  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(520).setHeight(690),
    'Edit Inventory Item'
  );
}

function itemFormHtml_() {
  return `<!doctype html><html><head><base target="_top"><style>
  body{font-family:Arial,sans-serif;padding:14px;color:#1F2937}
  label{display:block;font-weight:700;margin-top:9px}
  input,textarea{width:100%;box-sizing:border-box;padding:8px;margin-top:3px;border:1px solid #B7C9D6;border-radius:4px}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}
  .actions{margin-top:16px;display:flex;gap:8px}
  button{padding:10px 16px;border:0;border-radius:4px;background:#1F4E78;color:white;font-weight:700;cursor:pointer}
  .secondary{background:#6B7280}</style></head><body>
  <form id="f"><input type="hidden" name="row" id="row"><div class="grid">
  <div><label>Purchase date</label><input type="date" name="purchaseDate" required></div>
  <div><label>Title</label><input name="title" required></div>
  <div><label>SKU</label><input name="sku"></div>
  <div><label>Barcode</label><input name="barcode"></div>
  <div><label>Category</label><input name="category" required></div>
  <div><label>Purchase location</label><input name="purchaseLocation"></div>
  <div><label>Storage location</label><input name="storageLocation"></div>
  <div><label>Condition</label><input name="condition"></div>
  <div><label>Quantity</label><input type="number" min="1" name="quantity" value="1"></div>
  <div><label>Status</label><input name="status" value="Purchased"></div>
  <div><label>Purchase price</label><input type="number" step="0.01" min="0" name="purchasePrice"></div>
  <div><label>Tax paid</label><input type="number" step="0.01" min="0" name="taxPaid"></div>
  <div><label>Acquisition shipping</label><input type="number" step="0.01" min="0" name="acquisitionShipping"></div>
  <div><label>Expected sale price</label><input type="number" step="0.01" min="0" name="expectedSalePrice"></div>
  <div><label>Listed price</label><input type="number" step="0.01" min="0" name="listedPrice"></div>
  <div><label>Marketplace</label><input name="marketplace"></div>
  </div>
  <label>Receipt link</label><input type="url" name="receiptLink">
  <label>Photo link</label><input type="url" name="photoLink">
  <label>Notes</label><textarea name="notes" rows="3"></textarea>
  <div class="actions"><button type="submit">Save Item</button>
  <button type="button" class="secondary" onclick="google.script.host.close()">Cancel</button></div></form>
  <script>
  const existing = JSON.parse(<?!= JSON.stringify(itemJson) ?>);
  if(existing){
    document.getElementById('row').value=existing.row;
    const v=existing.values;
    const map={purchaseDate:v[1],title:v[2],sku:v[3],barcode:v[4],category:v[5],
      purchaseLocation:v[6],storageLocation:v[7],condition:v[8],quantity:v[9],
      purchasePrice:v[10],taxPaid:v[11],acquisitionShipping:v[12],
      expectedSalePrice:v[14],listedPrice:v[15],marketplace:v[16],status:v[18],
      receiptLink:v[22],photoLink:v[23],notes:v[24]};
    Object.keys(map).forEach(k=>{
      const e=document.querySelector('[name="'+k+'"]');
      if(e)e.value=map[k]??'';
    });
  }
  document.getElementById('f').addEventListener('submit',e=>{
    e.preventDefault();
    const d=Object.fromEntries(new FormData(e.target).entries());
    google.script.run.withSuccessHandler(()=>google.script.host.close())
      .withFailureHandler(err=>alert(err.message)).saveInventoryItem(d);
  });
  </script></body></html>`;
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

function num_(v) {
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

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

function buildDashboardSprint2_() {
  const s = sheet_(FTP2.SHEETS.DASHBOARD);
  s.clear();
  s.setHiddenGridlines(true);

  s.getRange('A1:H2').merge().setValue('FlipTracker Pro — Sprint 2')
    .setBackground(FTP2.COLORS.NAVY).setFontColor(FTP2.COLORS.WHITE)
    .setFontSize(22).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  const cards = [
    ['A4:B6','Inventory Cost','=SUM(Inventory!N2:N)','$#,##0.00'],
    ['C4:D6','Potential Value','=SUM(Inventory!O2:O)','$#,##0.00'],
    ['E4:F6','Potential Profit','=SUM(Inventory!U2:U)','$#,##0.00'],
    ['G4:H6','Average ROI','=IFERROR(AVERAGE(FILTER(Inventory!V2:V,Inventory!C2:C<>"")),0)','0.0%'],
    ['A8:B10','Items Tracked','=COUNTA(Inventory!C2:C)','0'],
    ['C8:D10','Items Listed','=COUNTIF(Inventory!S2:S,"Listed")','0'],
    ['E8:F10','Ready to List','=COUNTIF(Inventory!S2:S,"Ready to List")','0'],
    ['G8:H10','90+ Days','=COUNTIFS(Inventory!T2:T,">=90",Inventory!S2:S,"<>Sold",Inventory!C2:C,"<>")','0']
  ];

  cards.forEach(([a,title,formula,format]) => {
    const r = s.getRange(a);
    r.merge();
    r.getCell(1,1).setFormula(
      `="${title}"&CHAR(10)&TEXT(${formula.substring(1)},"${format}")`
    );
    r.setBackground(FTP2.COLORS.LIGHT_BLUE)
      .setFontWeight('bold').setFontSize(14).setWrap(true)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,FTP2.COLORS.BORDER,
        SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });

  s.getRange('A13:H13').merge().setValue('Inventory Workflow')
    .setBackground(FTP2.COLORS.NAVY).setFontColor(FTP2.COLORS.WHITE)
    .setFontWeight('bold').setHorizontalAlignment('center');

  s.getRange('A14:H18').merge().setValue(
    'Use FlipTracker Pro → Add Inventory Item to create stable inventory records. ' +
    'Select a row and choose Edit Selected Item to update it. Find Item searches ' +
    'Item ID, title, SKU, and barcode. Show Slow Inventory filters items held for ' +
    '90 days or longer.'
  ).setBackground('#F3F6F9').setWrap(true).setVerticalAlignment('middle');

  s.setColumnWidths(1,8,120);
}

function refreshDashboardSprint2() {
  buildDashboardSprint2_();
  SpreadsheetApp.getActive().setActiveSheet(sheet_(FTP2.SHEETS.DASHBOARD));
}

function goToInventorySprint2() {
  SpreadsheetApp.getActive().setActiveSheet(sheet_(FTP2.SHEETS.INVENTORY));
}


const FTP = {
  APP_NAME: 'FlipTracker Pro',
  VERSION: '1.0 Sprint 1',
  ROWS: 1000,
  COLORS: {
    NAVY:'#1F4E78', BLUE:'#5B9BD5', LIGHT_BLUE:'#D9EAF7',
    GREEN:'#70AD47', LIGHT_GREEN:'#E2F0D9', GOLD:'#FFC000',
    LIGHT_RED:'#FCE4D6', GRAY:'#F3F6F9', BORDER:'#B7C9D6',
    WHITE:'#FFFFFF', TEXT:'#1F2937'
  },
  SHEETS: ['Dashboard','Inventory','Sales','Expenses','Mileage',
           'CRA Tax Centre','Reports','Admin','Settings'],
  INVENTORY_HEADERS: [
    'Item ID','Purchase Date','Title','SKU','Barcode','Category',
    'Purchase Location','Storage Location','Condition','Quantity',
    'Purchase Price','Tax Paid','Acquisition Shipping','Total Cost',
    'Expected Sale Price','Listed Price','Marketplace','Listing Date',
    'Status','Days in Inventory','Projected Profit','Projected ROI %',
    'Receipt Link','Photo Link','Notes'
  ]
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('FlipTracker Pro')
    .addItem('Build / Repair Workbook','initializeFlipTrackerPro')
    .addSeparator()
    .addItem('Go to Dashboard','goToDashboard')
    .addItem('Go to Inventory','goToInventory')
    .addItem('Refresh Dashboard','refreshDashboard')
    .addSeparator()
    .addItem('About','showAbout')
    .addToUi();
}

function initializeFlipTrackerPro() {
  const ss = SpreadsheetApp.getActive();
  FTP.SHEETS.forEach(name => getOrCreateSheet_(name));
  buildAdmin_();
  buildSettings_();
  buildInventory_();
  buildPlaceholders_();
  buildDashboard_();
  ss.setActiveSheet(ss.getSheetByName('Dashboard'));
  ss.toast('Sprint 1 is ready.','FlipTracker Pro',5);
}

function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureSize_(sheet, rows, cols) {
  if (sheet.getMaxRows() < rows)
    sheet.insertRowsAfter(sheet.getMaxRows(), rows-sheet.getMaxRows());
  if (sheet.getMaxColumns() < cols)
    sheet.insertColumnsAfter(sheet.getMaxColumns(), cols-sheet.getMaxColumns());
}

function header_(range) {
  range.setBackground(FTP.COLORS.NAVY)
    .setFontColor(FTP.COLORS.WHITE).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrap(true);
}

function borders_(range) {
  range.setBorder(true,true,true,true,true,true,FTP.COLORS.BORDER,
                  SpreadsheetApp.BorderStyle.SOLID);
}

function buildAdmin_() {
  const s=getOrCreateSheet_('Admin'); s.clear(); s.setFrozenRows(1);
  const lists={
    Categories:['Electronics','Tools','Collectibles','Clothing','Furniture','Automotive','Household','Other'],
    PurchaseLocations:['Garage Sale','Thrift Store','Value Village','Facebook Marketplace','Auction','Retail Clearance','Other'],
    Marketplaces:['eBay','Facebook Marketplace','Kijiji','Poshmark','Etsy','Local Sale','Other'],
    Statuses:['Purchased','Needs Cleaning','Needs Testing','Ready to List','Listed','Sold','Shipped','Archived'],
    Conditions:['New','Open Box','Like New','Good','Fair','For Parts'],
    StorageLocations:['Garage','Basement','Shelf A','Shelf B','Bin 1','Bin 2','Other'],
    ExpenseCategories:['Fuel','Packaging','Shipping Supplies','Advertising','Storage','Phone','Internet','Office','Software','Bank Fees','Equipment','Other']
  };
  let col=1;
  Object.keys(lists).forEach(key=>{
    s.getRange(1,col).setValue(key);
    const vals=lists[key].map(v=>[v]);
    s.getRange(2,col,vals.length,1).setValues(vals);
    SpreadsheetApp.getActive().setNamedRange('FTP_'+key,s.getRange(2,col,50,1));
    s.setColumnWidth(col,180); col++;
  });
  header_(s.getRange(1,1,1,col-1));
  borders_(s.getRange(1,1,50,col-1));
}

function buildSettings_() {
  const s=getOrCreateSheet_('Settings'); s.clear();
  const data=[
    ['Setting','Value','Notes'],
    ['Business Name','','Optional'],
    ['Business Number','','Optional'],
    ['GST/HST Registered','No','Yes or No'],
    ['GST/HST Number','','Optional'],
    ['Province','Ontario',''],
    ['Currency','CAD',''],
    ['Fiscal Year Start','January 1',''],
    ['Default HST Rate',0.13,'Editable'],
    ['CRA Mileage Rate',0,'Enter rate for applicable tax year'],
    ['Tax Year',new Date().getFullYear(),'']
  ];
  s.getRange(1,1,data.length,3).setValues(data);
  header_(s.getRange(1,1,1,3)); borders_(s.getRange(1,1,data.length,3));
  s.setFrozenRows(1); s.setColumnWidths(1,3,190);
  s.getRange('B9').setNumberFormat('0.00%');
  s.getRange('B10').setNumberFormat('$#,##0.00');
  s.getRange('B4').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes','No'],true).setAllowInvalid(false).build()
  );
}

function buildInventory_() {
  const s=getOrCreateSheet_('Inventory'), h=FTP.INVENTORY_HEADERS;
  ensureSize_(s,FTP.ROWS+1,h.length);
  const oldRows=Math.max(s.getLastRow()-1,0);
  const old=oldRows ? s.getRange(2,1,oldRows,h.length).getValues() : [];
  s.clear(); s.setFrozenRows(1); s.setFrozenColumns(3);
  s.getRange(1,1,1,h.length).setValues([h]);
  if(old.length) s.getRange(2,1,old.length,h.length).setValues(old);

  const a=[],n=[],t=[],u=[],v=[];
  for(let r=2;r<=FTP.ROWS+1;r++){
    a.push([`=IF(C${r}="","","ITM-"&TEXT(ROW()-1,"00000"))`]);
    n.push([`=IF(C${r}="","",SUM(K${r}:M${r}))`]);
    t.push([`=IF(B${r}="","",MAX(0,TODAY()-B${r}))`]);
    u.push([`=IF(OR(O${r}="",N${r}=""),"",O${r}-N${r})`]);
    v.push([`=IFERROR(U${r}/N${r},"")`]);
  }
  s.getRange(2,1,FTP.ROWS,1).setFormulas(a);
  s.getRange(2,14,FTP.ROWS,1).setFormulas(n);
  s.getRange(2,20,FTP.ROWS,1).setFormulas(t);
  s.getRange(2,21,FTP.ROWS,1).setFormulas(u);
  s.getRange(2,22,FTP.ROWS,1).setFormulas(v);

  const validations=[
    [6,'FTP_Categories'],[7,'FTP_PurchaseLocations'],
    [8,'FTP_StorageLocations'],[9,'FTP_Conditions'],
    [17,'FTP_Marketplaces'],[19,'FTP_Statuses']
  ];
  validations.forEach(([c,name])=>{
    const rng=SpreadsheetApp.getActive().getRange(name);
    const rule=SpreadsheetApp.newDataValidation()
      .requireValueInRange(rng,true).setAllowInvalid(false).build();
    s.getRange(2,c,FTP.ROWS,1).setDataValidation(rule);
  });

  s.getRange(2,2,FTP.ROWS,1).setNumberFormat('yyyy-mm-dd');
  s.getRange(2,18,FTP.ROWS,1).setNumberFormat('yyyy-mm-dd');
  s.getRange(2,11,FTP.ROWS,6).setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
  s.getRange(2,21,FTP.ROWS,1).setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
  s.getRange(2,22,FTP.ROWS,1).setNumberFormat('0.0%;[Red]-0.0%');

  header_(s.getRange(1,1,1,h.length));
  borders_(s.getRange(1,1,FTP.ROWS+1,h.length));
  s.setRowHeight(1,42);

  const widths=[110,105,240,110,110,150,170,160,110,80,110,95,130,110,120,110,150,105,120,110,125,110,180,180,260];
  widths.forEach((w,i)=>s.setColumnWidth(i+1,w));

  if(s.getFilter()) s.getFilter().remove();
  s.getRange(1,1,FTP.ROWS+1,h.length).createFilter();

  const data=s.getRange(2,1,FTP.ROWS,h.length);
  s.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$S2="Sold"').setBackground(FTP.COLORS.LIGHT_GREEN)
      .setRanges([data]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$S2="Listed"').setBackground(FTP.COLORS.LIGHT_BLUE)
      .setRanges([data]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($T2>90,$S2<>"Sold",$S2<>"Archived",$C2<>"")')
      .setBackground(FTP.COLORS.LIGHT_RED).setRanges([data]).build()
  ]);
}

function buildPlaceholders_() {
  const defs={
    Sales:['Sale ID','Item ID','Sale Date','Marketplace','Sale Price','Shipping Charged','Shipping Paid','Marketplace Fees','Payment Fees','GST/HST Collected','Net Profit','Days to Sell','Buyer','Notes'],
    Expenses:['Expense ID','Date','Category','Vendor','Description','Subtotal','GST/HST Paid','Total','Business Use %','Deductible Amount','Receipt Link','Notes'],
    Mileage:['Trip ID','Date','Start','Destination','Business Purpose','Total Kilometres','Business Kilometres','CRA Rate','Claim Amount','Notes']
  };
  Object.keys(defs).forEach(name=>{
    const s=getOrCreateSheet_(name), h=defs[name]; s.clear();
    ensureSize_(s,100,h.length);
    s.getRange(1,1,1,h.length).setValues([h]);
    header_(s.getRange(1,1,1,h.length));
    borders_(s.getRange(1,1,100,h.length));
    s.setFrozenRows(1);
  });
  const cra=getOrCreateSheet_('CRA Tax Centre'); cra.clear();
  cra.getRange('A1:B2').setValues([
    ['CRA Tax Centre','Status'],
    ['Profit & Loss, inventory valuation, mileage and GST/HST reporting','Scheduled for Sprint 4']
  ]);
  header_(cra.getRange('A1:B1')); borders_(cra.getRange('A1:B2'));
  const rep=getOrCreateSheet_('Reports'); rep.clear();
  rep.getRange('A1:B2').setValues([
    ['Reports','Status'],['Business and accountant-ready reports','Scheduled for later sprint']
  ]);
  header_(rep.getRange('A1:B1')); borders_(rep.getRange('A1:B2'));
}

function buildDashboard_() {
  const s=getOrCreateSheet_('Dashboard'); s.clear(); s.setHiddenGridlines(true);
  s.getRange('A1:H2').merge().setValue('FlipTracker Pro Dashboard')
    .setBackground(FTP.COLORS.NAVY).setFontColor(FTP.COLORS.WHITE)
    .setFontSize(22).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  const cards=[
    ['A4:B6','Inventory Cost','SUM(Inventory!N2:N)','$#,##0.00'],
    ['C4:D6','Potential Value','SUM(Inventory!O2:O)','$#,##0.00'],
    ['E4:F6','Potential Profit','SUM(Inventory!U2:U)','$#,##0.00'],
    ['G4:H6','Average ROI','IFERROR(AVERAGE(FILTER(Inventory!V2:V,Inventory!C2:C<>"")),0)','0.0%'],
    ['A8:B10','Items Tracked','COUNTA(Inventory!C2:C)','0'],
    ['C8:D10','Items Listed','COUNTIF(Inventory!S2:S,"Listed")','0'],
    ['E8:F10','Items Sold','COUNTIF(Inventory!S2:S,"Sold")','0'],
    ['G8:H10','Over 90 Days','COUNTIFS(Inventory!T2:T,">90",Inventory!S2:S,"<>Sold",Inventory!C2:C,"<>")','0']
  ];
  cards.forEach(([a1,title,formula,fmt])=>{
    const r=s.getRange(a1); r.merge();
    r.getCell(1,1).setFormula(`="${title}"&CHAR(10)&TEXT(${formula},"${fmt}")`);
    r.setBackground(FTP.COLORS.LIGHT_BLUE).setFontColor(FTP.COLORS.TEXT)
      .setFontWeight('bold').setFontSize(14).setWrap(true)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,FTP.COLORS.BORDER,
                 SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });
  s.getRange('A13:H13').merge().setValue('Sprint 1 Foundation Installed');
  header_(s.getRange('A13:H13'));
  s.getRange('A14:H18').merge().setValue(
    'Begin entering items on the Inventory sheet. Automatic IDs, total cost, projected profit, ROI, days in inventory, drop-downs, filters, borders and status colouring are active.'
  ).setBackground(FTP.COLORS.GRAY).setWrap(true).setVerticalAlignment('middle');
  s.setColumnWidths(1,8,120); s.setFrozenRows(2);
}

function refreshDashboard(){ buildDashboard_(); goToDashboard(); }
function goToDashboard(){ SpreadsheetApp.getActive().setActiveSheet(getOrCreateSheet_('Dashboard')); }
function goToInventory(){ SpreadsheetApp.getActive().setActiveSheet(getOrCreateSheet_('Inventory')); }
function showAbout(){
  SpreadsheetApp.getUi().alert('FlipTracker Pro',FTP.VERSION+
    '\n\nSprint 1: workbook foundation, Dashboard, Inventory, Admin, Settings, formulas, formatting and drop-downs.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

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

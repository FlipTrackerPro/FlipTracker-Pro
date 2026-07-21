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

  const yearSales=sales.filter(r=>r[0]&&inYear(r[2]));
  const grossRevenue=yearSales.reduce((n,r)=>n+num3_(r[13]),0);
  const taxCollected=yearSales.reduce((n,r)=>n+num3_(r[11]),0);
  const revenueExTax=grossRevenue; // Sales Gross Revenue deliberately excludes separately entered GST/HST.
  const inventoryById={}; inventory.filter(r=>r[0]).forEach(r=>inventoryById[String(r[0])]=r);
  const specificCogs=yearSales.reduce((n,r)=>{
    const source=inventoryById[String(r[1])];
    const recoverableTax=regular&&source?num3_(source[11]):0;
    return n+Math.max(0,num3_(r[12])-recoverableTax);
  },0);
  const yearPurchases=inventory.filter(r=>r[0]&&inYear(r[1]));
  const purchases=yearPurchases.reduce((n,r)=>n+Math.max(0,num3_(r[13])-(regular?num3_(r[11]):0)),0);
  const inventoryItc=regular?yearPurchases.reduce((n,r)=>n+num3_(r[11]),0):0;

  const soldByEnd={};
  sales.filter(r=>r[0]&&r[1]&&r[2] instanceof Date&&!isNaN(r[2])&&r[2]<=end)
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

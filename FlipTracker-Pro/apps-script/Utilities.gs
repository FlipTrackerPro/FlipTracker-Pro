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


function roundMoney3_(value) {
  const n=num3_(value);
  return Math.round((n+Number.EPSILON)*100)/100;
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
  const s=sheet3_(FTP3.SHEETS.INVENTORY);
  if(s.getLastRow()<2)return [];
  const headers=s.getRange(1,1,1,s.getLastColumn()).getDisplayValues()[0].map(String);
  const rows=s.getRange(2,1,s.getLastRow()-1,s.getLastColumn()).getDisplayValues();
  const id=headerIndex3_(headers,'Item ID'), description=headerIndex3_(headers,'Description'), status=headerIndex3_(headers,'Status');
  return rows.filter(r=>r[id]&&r[description]&&!['Sold','Archived'].includes(r[status]))
    .map(r=>({id:r[id],description:r[description]||''}));
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

function headerIndex3_(headers,name) {
  const i=headers.indexOf(name);
  if(i<0)throw new Error('Required column not found: '+name);
  return i;
}

function rowRecord3_(headers,row) {
  return headers.reduce((o,h,i)=>{o[h]=row[i];return o;},{});
}

function columnLetter3_(column) {
  let n=Number(column), result='';
  while(n>0){n--;result=String.fromCharCode(65+(n%26))+result;n=Math.floor(n/26);}
  return result;
}

function sheetColumnLetter3_(sheetName,headerName) {
  const s=sheet3_(sheetName);
  const map=headerMap3_(s);
  if(!map[headerName])throw new Error(sheetName+' is missing required column: '+headerName);
  return columnLetter3_(map[headerName]);
}


function applyFlipTrackerNumberFormats3_() {
  const currency='$#,##0.00;[Red]-$#,##0.00';
  const integer='0';
  const percent='0%;[Red]-0%';

  const specs=[
    {
      sheet:FTP3.SHEETS.INVENTORY,
      currency:['Purchase Price','Tax Paid','Acquisition Shipping','Total Cost','Listed Price','Expected Sale Price','Projected Profit'],
      integer:['Quantity','Days in Inventory'],
      percent:['Projected ROI %']
    },
    {
      sheet:FTP3.SHEETS.SALES,
      currency:['Sale Price','Shipping Charged','Shipping Actual','Packaging Cost','Marketplace Fees','Payment Fees','Promotion Expense','GST/HST Collected','Item Cost','Gross Revenue','Total Selling Costs','Net Proceeds','Realized Profit'],
      integer:['Days to Sell','Box Qty','Bubble Wrap Qty','Mailer Qty','Tape Qty','Other Packaging Qty'],
      percent:['Realized ROI %']
    },
    {
      sheet:FTP3.SHEETS.EXPENSES,
      currency:['Subtotal','GST/HST Paid','Total','Deductible Amount'],
      integer:[],
      percent:['Business Use %']
    },
    {
      sheet:FTP3.SHEETS.MILEAGE,
      currency:['CRA Rate','Claim Amount'],
      integer:['Start Odometer','End Odometer','Google Maps Distance','Total Kilometres','Business Kilometres'],
      percent:[]
    },
    {
      sheet:FTP3.SHEETS.PACKAGING,
      currency:['Purchase Cost','Cost Per Unit'],
      integer:['Units Purchased','Quantity On Hand','Reorder Level'],
      percent:[]
    }
  ];

  specs.forEach(spec=>{
    const s=SpreadsheetApp.getActive().getSheetByName(spec.sheet);
    if(!s)return;
    const map=headerMap3_(s);
    spec.currency.forEach(h=>{if(map[h])s.getRange(2,map[h],Math.max(1,s.getMaxRows()-1),1).setNumberFormat(currency);});
    spec.integer.forEach(h=>{if(map[h])s.getRange(2,map[h],Math.max(1,s.getMaxRows()-1),1).setNumberFormat(integer);});
    spec.percent.forEach(h=>{if(map[h])s.getRange(2,map[h],Math.max(1,s.getMaxRows()-1),1).setNumberFormat(percent);});
  });

  const admin=SpreadsheetApp.getActive().getSheetByName(FTP3.SHEETS.ADMIN);
  if(admin){
    admin.getRange('B9').setNumberFormat('0%');
    admin.getRange('B10').setNumberFormat(currency);
    admin.getRange('B13').setNumberFormat(currency);
  }
}

function wholeNumberRule3_(minimum) {
  return SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied('=OR(INDIRECT(ADDRESS(ROW(),COLUMN()))="",AND(ISNUMBER(INDIRECT(ADDRESS(ROW(),COLUMN()))),INDIRECT(ADDRESS(ROW(),COLUMN()))=INT(INDIRECT(ADDRESS(ROW(),COLUMN()))),INDIRECT(ADDRESS(ROW(),COLUMN()))>='+Number(minimum||0)+'))')
    .setAllowInvalid(false)
    .setHelpText('Enter a whole number with no decimal places.')
    .build();
}

function repairWholeNumberValues3_() {
  const specs=[
    [FTP3.SHEETS.INVENTORY,['Quantity','Days in Inventory']],
    [FTP3.SHEETS.SALES,['Days to Sell','Box Qty','Bubble Wrap Qty','Mailer Qty','Tape Qty','Other Packaging Qty']],
    [FTP3.SHEETS.MILEAGE,['Start Odometer','End Odometer','Google Maps Distance','Total Kilometres','Business Kilometres']],
    [FTP3.SHEETS.PACKAGING,['Units Purchased','Quantity On Hand','Reorder Level']]
  ];

  specs.forEach(([sheetName,headers])=>{
    const s=SpreadsheetApp.getActive().getSheetByName(sheetName);
    if(!s||s.getLastRow()<2)return;
    const map=headerMap3_(s);
    headers.forEach(header=>{
      if(!map[header])return;
      const range=s.getRange(2,map[header],s.getLastRow()-1,1);
      const values=range.getValues().map(([v])=>{
        if(v===''||v===null||v instanceof Date)return [v];
        const n=Number(v);
        return [Number.isFinite(n)?Math.round(n):v];
      });
      range.setValues(values).setNumberFormat('0');
    });
  });
}

function standardizeFlipTrackerNumbers() {
  repairWholeNumberValues3_();
  repairPackagingCostPerUnit3_();
  applyFlipTrackerNumberFormats3_();
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    'Number Formatting Complete',
    'All non-currency numeric fields now display as whole numbers. Dollar fields display two decimal places. Packaging Cost Per Unit has been recalculated.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

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

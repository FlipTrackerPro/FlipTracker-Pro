function inventoryHealth71_(days,status,description) {
  if(!String(description||'').trim())return '';
  const normalized=String(status||'').trim().toLowerCase();
  if(normalized==='sold')return 'Sold';
  if(normalized==='archived')return 'Archived';
  const age=Math.max(0,Math.floor(num3_(days)));
  if(age>=120)return 'Dead Stock';
  if(age>=90)return 'Stale';
  if(age>=30)return 'Aging';
  return 'Fresh';
}

function recalculateInventoryHealthRow71_(sheet,row) {
  const s=sheet||sheet3_(FTP3.SHEETS.INVENTORY);
  if(row<2)return;
  const c=headerMap3_(s);
  if(!c['Inventory Health'])return;
  const description=s.getRange(row,c['Description']).getDisplayValue();
  const status=s.getRange(row,c['Status']).getDisplayValue();
  const days=s.getRange(row,c['Days in Inventory']).getValue();
  s.getRange(row,c['Inventory Health']).setValue(inventoryHealth71_(days,status,description));
}

function repairInventoryHealth71_(sheet) {
  const s=sheet||sheet3_(FTP3.SHEETS.INVENTORY);
  if(s.getLastRow()<2)return 0;
  const c=headerMap3_(s);
  if(!c['Inventory Health'])return 0;
  const count=s.getLastRow()-1;
  const descriptions=s.getRange(2,c['Description'],count,1).getDisplayValues();
  const statuses=s.getRange(2,c['Status'],count,1).getDisplayValues();
  const days=s.getRange(2,c['Days in Inventory'],count,1).getValues();
  const values=descriptions.map((r,i)=>[inventoryHealth71_(days[i][0],statuses[i][0],r[0])]);
  s.getRange(2,c['Inventory Health'],count,1).setValues(values);
  return values.filter(r=>r[0]).length;
}

function repairInventoryHealth() {
  const count=repairInventoryHealth71_();
  SpreadsheetApp.flush();
  SpreadsheetApp.getActive().toast(count+' inventory health records refreshed.','FlipTracker Pro',6);
}

function applyInventoryHealthFormatting71_(sheet) {
  const s=sheet||sheet3_(FTP3.SHEETS.INVENTORY);
  const c=headerMap3_(s);
  if(!c['Inventory Health'])return;
  const range=s.getRange(2,c['Inventory Health'],FTP3.ROWS,1);
  s.setConditionalFormatRules(s.getConditionalFormatRules().filter(rule=>
    !rule.getRanges().some(r=>r.getColumn()===c['Inventory Health'])
  ).concat([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Fresh').setBackground(FTP3.COLORS.LIGHT_GREEN).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Aging').setBackground(FTP3.COLORS.GOLD).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Stale').setBackground(FTP3.COLORS.ORANGE).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Dead Stock').setBackground(FTP3.COLORS.LIGHT_RED).setRanges([range]).build()
  ]));
}

function showInventoryHealthSummary71_() {
  const s=sheet3_(FTP3.SHEETS.INVENTORY),c=headerMap3_(s);
  if(s.getLastRow()<2){SpreadsheetApp.getUi().alert('Inventory Health','No inventory records found.',SpreadsheetApp.getUi().ButtonSet.OK);return;}
  const values=s.getRange(2,c['Inventory Health'],s.getLastRow()-1,1).getDisplayValues().flat();
  const counts={'Fresh':0,'Aging':0,'Stale':0,'Dead Stock':0};
  values.forEach(v=>{if(Object.prototype.hasOwnProperty.call(counts,v))counts[v]++;});
  SpreadsheetApp.getUi().alert('Inventory Health','Fresh: '+counts.Fresh+'\nAging: '+counts.Aging+'\nStale: '+counts.Stale+'\nDead Stock: '+counts['Dead Stock'],SpreadsheetApp.getUi().ButtonSet.OK);
}

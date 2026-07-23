function buildMileageLists70_() {
  const s=sheet3_(FTP3.SHEETS.MILEAGE_LISTS);
  ensureSize3_(s,200,10);
  const headers=['Start Locations','End Locations','Business Purposes','Vehicles','','Favorite Trip','Start Location','End Location','Business Purpose','Vehicle'];
  s.getRange(1,1,1,headers.length).setValues([headers]);
  header3_(s.getRange(1,1,1,4)); header3_(s.getRange(1,6,1,5));
  const defaults={
    1:['Home','Storage Unit','Garage','Office'],
    2:['Canada Post','UPS Store','FedEx','Value Village','Salvation Army','Goodwill','Walmart','Costco','Storage Unit','Customer Pickup','Estate Sale','Auction','Flea Market'],
    3:['Inventory Sourcing','Ship Orders','Customer Pickup','Banking','Office Supplies','Packaging Supplies','Equipment Repair','Photography','Bookkeeping','Inventory Organization'],
    4:['Honda CR-V','Truck','Rental','Other']
  };
  Object.keys(defaults).forEach(k=>{
    const col=Number(k), existing=s.getRange(2,col,Math.max(1,s.getLastRow()-1),1).getDisplayValues().flat().filter(Boolean);
    if(!existing.length)s.getRange(2,col,defaults[col].length,1).setValues(defaults[col].map(v=>[v]));
  });
  const favoriteExisting=s.getRange(2,6,Math.max(1,s.getLastRow()-1),1).getDisplayValues().flat().filter(Boolean);
  if(!favoriteExisting.length){
    s.getRange(2,6,3,5).setValues([
      ['Post Office','Home','Canada Post','Ship Orders','Honda CR-V'],
      ['Thrift Run','Home','Value Village','Inventory Sourcing','Honda CR-V'],
      ['Storage','Home','Storage Unit','Inventory Organization','Honda CR-V']
    ]);
  }
  s.setFrozenRows(1); s.setColumnWidths(1,4,175); s.setColumnWidth(5,25); s.setColumnWidths(6,5,175);
  borders3_(s.getRange(1,1,Math.max(20,s.getLastRow()),4)); borders3_(s.getRange(1,6,Math.max(20,s.getLastRow()),5));
  refreshMileageDropdowns70_();
}

function mileageListRange70_(column){
  const s=sheet3_(FTP3.SHEETS.MILEAGE_LISTS);
  const last=Math.max(2,s.getLastRow());
  return s.getRange(2,column,last-1,1);
}

function refreshMileageDropdowns70_(){
  const m=sheet3_(FTP3.SHEETS.MILEAGE); if(m.getLastRow()<1)return;
  const map=headerMap3_(m), rows=Math.max(1,m.getMaxRows()-1);
  const specs={'Start Location':1,'End Location':2,'Business Purpose':3,'Vehicle':4};
  Object.keys(specs).forEach(h=>{
    if(!map[h])return;
    const rule=SpreadsheetApp.newDataValidation().requireValueInRange(mileageListRange70_(specs[h]),true).setAllowInvalid(true).build();
    m.getRange(2,map[h],rows,1).setDataValidation(rule);
  });
}

function buildMileageSprint3_() {
  buildMileageLists70_();
  const s=sheet3_(FTP3.SHEETS.MILEAGE);
  migrateMileageSchema70_(s);
  ensureSize3_(s,FTP3.ROWS+1,FTP3.MILEAGE_HEADERS.length);
  s.getRange(1,1,1,FTP3.MILEAGE_HEADERS.length).setValues([FTP3.MILEAGE_HEADERS]);
  header3_(s.getRange(1,1,1,FTP3.MILEAGE_HEADERS.length)); s.setFrozenRows(1);
  const map=headerMap3_(s);
  s.getRange(2,map['Date'],FTP3.ROWS,1).setNumberFormat('yyyy-mm-dd');
  ['Start Odometer','End Odometer','Total Kilometres','Business Kilometres'].forEach(h=>s.getRange(2,map[h],FTP3.ROWS,1).setNumberFormat('0'));
  applyMileageTotalFormulas81_(s);
  ['CRA Rate','Claim Amount'].forEach(h=>s.getRange(2,map[h],FTP3.ROWS,1).setNumberFormat('$#,##0.00'));
  s.getRange(2,map['Created At'],FTP3.ROWS,1).setNumberFormat('yyyy-mm-dd hh:mm');
  const yesNo=SpreadsheetApp.newDataValidation().requireValueInList(['No','Yes'],true).setAllowInvalid(false).build();
  s.getRange(2,map['Round Trip'],FTP3.ROWS,1).setDataValidation(yesNo);
  refreshMileageDropdowns70_();
  if(s.getFilter())s.getFilter().remove();
  s.getRange(1,1,FTP3.ROWS+1,FTP3.MILEAGE_HEADERS.length).createFilter();
  s.autoResizeColumns(1,FTP3.MILEAGE_HEADERS.length);
  ['Start Location','End Location','Business Purpose','Vehicle','Route Link','Notes'].forEach(h=>s.setColumnWidth(map[h],180));
}

function migrateMileageSchema70_(s) {
  if(s.getLastRow()<1)return;
  const oldHeaders=s.getRange(1,1,1,s.getLastColumn()).getDisplayValues()[0];
  if(!oldHeaders.some(Boolean)||oldHeaders.join('|')===FTP3.MILEAGE_HEADERS.join('|'))return;
  const aliases={'Start Location':['Start'],'End Location':['Destination','End'],'Start Odometer':['Odometer Start','Mileage Start'],'End Odometer':['Odometer End','Mileage End']};
  const oldMap={}; oldHeaders.forEach((h,i)=>{if(h)oldMap[String(h).trim()]=i;});
  const oldRows=s.getLastRow()>1?s.getRange(2,1,s.getLastRow()-1,s.getLastColumn()).getValues():[];
  const newRows=oldRows.map(row=>FTP3.MILEAGE_HEADERS.map(header=>{
    const candidates=[header].concat(aliases[header]||[]);
    for(const name of candidates)if(Object.prototype.hasOwnProperty.call(oldMap,name))return row[oldMap[name]];
    if(header==='Round Trip')return 'No';
    if(header==='Route Link'){
      const start=oldMap['Start Location']!=null?row[oldMap['Start Location']]:(oldMap['Start']!=null?row[oldMap['Start']]:'');
      const end=oldMap['End Location']!=null?row[oldMap['End Location']]:(oldMap['Destination']!=null?row[oldMap['Destination']]:'');
      return mileageRouteFormula70_(start,end);
    }
    return '';
  }));
  s.clear(); ensureSize3_(s,Math.max(FTP3.ROWS+1,newRows.length+1),FTP3.MILEAGE_HEADERS.length);
  s.getRange(1,1,1,FTP3.MILEAGE_HEADERS.length).setValues([FTP3.MILEAGE_HEADERS]);
  if(newRows.length)s.getRange(2,1,newRows.length,FTP3.MILEAGE_HEADERS.length).setValues(newRows);
}

function getMileageFormData70_(){
  const s=sheet3_(FTP3.SHEETS.MILEAGE_LISTS), last=Math.max(2,s.getLastRow());
  const values=s.getRange(2,1,last-1,10).getDisplayValues();
  const clean=col=>values.map(r=>r[col]).filter(Boolean);
  return {starts:clean(0),ends:clean(1),purposes:clean(2),vehicles:clean(3),favorites:values.filter(r=>r[5]).map(r=>({name:r[5],start:r[6],end:r[7],purpose:r[8],vehicle:r[9]}))};
}

function showRecordMileageForm() {
  const data=JSON.stringify(getMileageFormData70_()).replace(/</g,'\\u003c');
  const html=HtmlService.createHtmlOutput(`<!doctype html><html><head><base target="_top"><style>
  body{font-family:Arial;padding:14px;color:#1F2937}label{display:block;font-weight:700;margin-top:8px}input,select,textarea{width:100%;box-sizing:border-box;padding:8px;margin-top:3px;border:1px solid #B7C9D6;border-radius:4px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}button{margin-top:16px;padding:10px 16px;border:0;border-radius:4px;background:#1F4E78;color:#fff;font-weight:700}.hint{font-size:12px;color:#667085;margin-top:3px}</style></head><body><form id="f">
  <label>Favorite trip</label><select id="favorite"><option value="">Choose a saved trip (optional)</option></select><div class="grid">
  <div><label>Date</label><input type="date" name="date" required></div><div><label>CRA rate per km</label><input type="number" step="0.01" min="0" name="craRate"></div>
  <div><label>Start location</label><select name="startLocation"></select></div><div><label>End location</label><select name="endLocation"></select></div>
  <div><label>Business purpose</label><select name="purpose" required></select></div><div><label>Vehicle</label><select name="vehicle"></select></div>
  <div><label>Round trip</label><select name="roundTrip"><option>No</option><option>Yes</option></select></div><div></div>
  <div><label>Start odometer</label><input type="number" step="1" min="0" name="odoStart"></div><div><label>End odometer</label><input type="number" step="1" min="0" name="odoEnd"></div>
  <div><label>Business kilometres</label><input type="number" step="1" min="0" name="businessKm"><div class="hint">Blank uses full trip distance.</div></div></div>
  <label>Notes</label><textarea name="notes" rows="3"></textarea><button type="submit">Record Mileage</button></form><script>
  const D=${data}; const fill=(name,items)=>{const e=document.querySelector('[name="'+name+'"]');e.innerHTML='<option value="">Select...</option>'+items.map(x=>'<option>'+String(x).replace(/&/g,'&amp;').replace(/</g,'&lt;')+'</option>').join('');};
  fill('startLocation',D.starts);fill('endLocation',D.ends);fill('purpose',D.purposes);fill('vehicle',D.vehicles);
  const fav=document.getElementById('favorite');D.favorites.forEach((x,i)=>fav.add(new Option(x.name,i)));fav.onchange=()=>{if(fav.value==='')return;const x=D.favorites[Number(fav.value)];['startLocation','endLocation','purpose','vehicle'].forEach(k=>document.querySelector('[name="'+k+'"]').value=x[k]||'');};
  document.querySelector('[name="date"]').value=new Date().toISOString().slice(0,10);
  document.getElementById('f').addEventListener('submit',e=>{e.preventDefault();google.script.run.withSuccessHandler(()=>google.script.host.close()).withFailureHandler(x=>alert(x.message)).saveMileage3(Object.fromEntries(new FormData(e.target).entries()));});</script></body></html>`).setWidth(620).setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html,'Record Mileage');
}

function mileageRouteFormula70_(start,end){
  if(!start||!end)return '';
  const url='https://www.google.com/maps/dir/?api=1&origin='+encodeURIComponent(start)+'&destination='+encodeURIComponent(end);
  return '=HYPERLINK("'+url.replace(/"/g,'""')+'","Open Route")';
}
function saveMileage3(form){return saveMileage3_(form);}
function saveMileage3_(form) {
  if(!form.date||!form.purpose)throw new Error('Date and business purpose are required.');
  const odoStart=Math.round(num3_(form.odoStart)), odoEnd=Math.round(num3_(form.odoEnd));
  const total=Math.max(0,odoEnd-odoStart);
  const roundTrip=String(form.roundTrip||'No')==='Yes'?'Yes':'No';
  const requestedBusiness=form.businessKm===''||form.businessKm==null?total:Math.round(num3_(form.businessKm));
  const business=Math.min(total,Math.max(0,requestedBusiness));
  const configuredRate=num3_(getSettingV04_('CRA Mileage Rate',0));
  const rate=form.craRate===''||form.craRate==null?configuredRate:num3_(form.craRate), claim=roundMoney3_(business*rate);
  const start=form.startLocation||'', end=form.endLocation||'';
  const record={'Trip ID':nextId3_(FTP3.SHEETS.MILEAGE,1,'TRP'),'Date':date3_(form.date),'Start Location':start,'End Location':end,'Business Purpose':form.purpose||'','Round Trip':roundTrip,'Vehicle':form.vehicle||'','Start Odometer':odoStart,'End Odometer':odoEnd,'Total Kilometres':total,'Business Kilometres':business,'CRA Rate':rate,'Claim Amount':claim,'Route Link':mileageRouteFormula70_(start,end),'Notes':form.notes||'','Created At':new Date()};
  const s=sheet3_(FTP3.SHEETS.MILEAGE), row=Math.max(s.getLastRow()+1,2);
  s.getRange(row,1,1,FTP3.MILEAGE_HEADERS.length).setValues([FTP3.MILEAGE_HEADERS.map(h=>record[h])]);
  setMileageTotalFormula81_(s,row);
  refreshDashboardSprint3();
}

function setMileageTotalFormula81_(sheet,row) {
  const map=headerMap3_(sheet);
  if(!map['Start Odometer']||!map['End Odometer']||!map['Total Kilometres'])return;
  const start=sheet.getRange(row,map['Start Odometer']).getA1Notation();
  const end=sheet.getRange(row,map['End Odometer']).getA1Notation();
  sheet.getRange(row,map['Total Kilometres']).setFormula('=IF(OR('+start+'="",'+end+'=""),"",MAX(0,'+end+'-'+start+'))').setNumberFormat('0');
}

function applyMileageTotalFormulas81_(sheet) {
  const map=headerMap3_(sheet);
  if(!map['Start Odometer']||!map['End Odometer']||!map['Total Kilometres'])return;
  const rows=Math.max(1,sheet.getMaxRows()-1);
  const startCol=columnLetter3_(map['Start Odometer']);
  const endCol=columnLetter3_(map['End Odometer']);
  const formulas=Array.from({length:rows},(_,i)=>{const r=i+2;return ['=IF(OR('+startCol+r+'="",'+endCol+r+'=""),"",MAX(0,'+endCol+r+'-'+startCol+r+'))'];});
  sheet.getRange(2,map['Total Kilometres'],rows,1).setFormulas(formulas).setNumberFormat('0');
}

function goToMileageLists70_(){SpreadsheetApp.getActive().setActiveSheet(sheet3_(FTP3.SHEETS.MILEAGE_LISTS));}

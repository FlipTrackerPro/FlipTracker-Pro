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
  <div><label>Date</label><input type="date" name="date" required></div><div><label>CRA rate</label><input type="number" step="0.01" min="0" name="craRate"></div>
  <div><label>Start</label><input name="start"></div><div><label>Destination</label><input name="destination"></div>
  <div><label>Odometer start</label><input type="number" step="1" min="0" name="odoStart"></div>
  <div><label>Odometer end</label><input type="number" step="1" min="0" name="odoEnd"></div>
  <div><label>Business kilometres</label><input type="number" step="1" min="0" name="businessKm"></div></div>
  <label>Business purpose</label><input name="purpose" required><label>Notes</label><textarea name="notes" rows="3"></textarea>
  <button type="submit">Record Mileage</button></form><script>
  document.querySelector('[name="date"]').value=new Date().toISOString().slice(0,10);
  document.getElementById('f').addEventListener('submit',e=>{e.preventDefault();
  google.script.run.withSuccessHandler(()=>google.script.host.close()).withFailureHandler(x=>alert(x.message))
  .saveMileage3(Object.fromEntries(new FormData(e.target).entries()));});</script></body></html>`)
  .setWidth(520).setHeight(590);
  SpreadsheetApp.getUi().showModalDialog(html,'Record Mileage');
}

function saveMileage3(form) { return saveMileage3_(form); }

function saveMileage3_(form) {
  if (!form.date || !form.purpose) throw new Error('Date and business purpose are required.');
  const start = Math.round(num3_(form.odoStart));
  const end = Math.round(num3_(form.odoEnd));
  const total = Math.max(0,end-start);
  const requestedBusiness = form.businessKm === '' || form.businessKm == null ? total : Math.round(num3_(form.businessKm));
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

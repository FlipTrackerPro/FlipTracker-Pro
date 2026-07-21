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

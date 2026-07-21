function buildPackagingSprint3_() {
  const aliases={
    'Category':['Type'], 'Item Name':['Description'], 'Unit of Measure':[],
    'Units Purchased':[], 'Purchase Cost':[], 'Cost Per Unit':[],
    'Quantity On Hand':[], 'Reorder Level':[], 'Supplier':[],
    'Product Link':[], 'Notes':[], 'Updated At':[]
  };
  migrateSheetHeaders3_(FTP3.SHEETS.PACKAGING,FTP3.PACKAGING_HEADERS,aliases);
  const s=sheet3_(FTP3.SHEETS.PACKAGING);
  ensureSize3_(s,300,FTP3.PACKAGING_HEADERS.length);
  header3_(s.getRange(1,1,1,FTP3.PACKAGING_HEADERS.length));
  s.setFrozenRows(1);
  setValidation3_(s,2,'FTP3_PackagingTypes',299);
  const unitRule=SpreadsheetApp.newDataValidation().requireValueInList(['Each','Roll','Foot','Sheet','Metre','Bundle','Other'],true).setAllowInvalid(true).build();
  s.getRange(2,5,299,1).setDataValidation(unitRule);
  const yesNo=SpreadsheetApp.newDataValidation().requireValueInList(['Yes','No'],true).setAllowInvalid(false).build();
  s.getRange(2,14,299,1).setDataValidation(yesNo);
  s.getRange(2,7,299,2).setNumberFormat('$#,##0.000');
  s.getRange(2,6,299,1).setNumberFormat('0.000');
  s.getRange(2,9,299,2).setNumberFormat('0.000');
  s.getRange(2,16,299,1).setNumberFormat('yyyy-mm-dd hh:mm');
  if(s.getLastRow()>1){
    const active=s.getRange(2,14,s.getLastRow()-1,1); active.getValues().forEach((r,i)=>{if(!r[0])active.getCell(i+1,1).setValue('Yes');});
    const uom=s.getRange(2,5,s.getLastRow()-1,1); uom.getValues().forEach((r,i)=>{if(!r[0])uom.getCell(i+1,1).setValue('Each');});
  }
  const data=s.getRange(2,1,299,FTP3.PACKAGING_HEADERS.length);
  s.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($I2<=$J2,$A2<>"",$N2="Yes")')
      .setBackground(FTP3.COLORS.LIGHT_RED).setRanges([data]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($I2>0,$I2<=$J2*1.5,$A2<>"",$N2="Yes")')
      .setBackground(FTP3.COLORS.GOLD).setRanges([data]).build()
  ]);
  if(s.getFilter())s.getFilter().remove();
  s.getRange(1,1,Math.max(2,s.getLastRow()),FTP3.PACKAGING_HEADERS.length).createFilter();
  s.autoResizeColumns(1,FTP3.PACKAGING_HEADERS.length);
  [3,13,15].forEach(c=>s.setColumnWidth(c,c===15?240:180));
}

function showPackagingForm() {
  const html=HtmlService.createHtmlOutput(`<!doctype html><html><head><base target="_top"><style>
  body{font-family:Arial;padding:14px;color:#1F2937}label{display:block;font-weight:700;margin-top:8px}
  input,select,textarea{width:100%;box-sizing:border-box;padding:8px;margin-top:3px;border:1px solid #B7C9D6;border-radius:4px}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}button{margin-top:16px;padding:10px 16px;border:0;border-radius:4px;background:#1F4E78;color:#fff;font-weight:700}
  </style></head><body><form id="f"><div class="grid">
  <div><label>Category</label><input name="category" placeholder="Box, Bubble Wrap, Tape..." required></div>
  <div><label>Item name</label><input name="itemName" placeholder="12x10x8 Shipping Box" required></div>
  <div><label>Size</label><input name="size"></div>
  <div><label>Unit of measure</label><select name="unitOfMeasure"><option>Each</option><option>Roll</option><option>Foot</option><option>Sheet</option><option>Metre</option><option>Bundle</option><option>Other</option></select></div>
  <div><label>Units purchased</label><input type="number" step="0.001" min="0.001" name="unitsPurchased" required></div>
  <div><label>Purchase cost</label><input type="number" step="0.01" min="0" name="purchaseCost" required></div>
  <div><label>Quantity on hand</label><input type="number" step="0.001" min="0" name="quantityOnHand"></div>
  <div><label>Reorder level</label><input type="number" step="0.001" min="0" name="reorderLevel"></div>
  <div><label>Supplier</label><input name="supplier"></div><div><label>Supplier SKU</label><input name="supplierSku"></div>
  <div><label>Product link</label><input type="url" name="productLink"></div><div><label>Active</label><select name="active"><option>Yes</option><option>No</option></select></div>
  </div><label>Notes</label><textarea name="notes" rows="3"></textarea>
  <button type="submit">Add Supply</button></form><script>
  document.getElementById('f').addEventListener('submit',e=>{e.preventDefault();
  google.script.run.withSuccessHandler(()=>google.script.host.close()).withFailureHandler(x=>alert(x.message))
  .savePackaging3_(Object.fromEntries(new FormData(e.target).entries()));});</script></body></html>`).setWidth(560).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html,'Add Packaging Supply');
}

function savePackaging3_(form) {
  const units=num3_(form.unitsPurchased); const purchaseCost=num3_(form.purchaseCost);
  if(units<=0)throw new Error('Units purchased must be greater than zero.');
  if(purchaseCost<0)throw new Error('Purchase cost cannot be negative.');
  const values=[nextId3_(FTP3.SHEETS.PACKAGING,1,'PKG'),form.category||'',form.itemName||'',form.size||'',form.unitOfMeasure||'Each',units,purchaseCost,purchaseCost/units,
    form.quantityOnHand===''||form.quantityOnHand==null?units:num3_(form.quantityOnHand),num3_(form.reorderLevel),form.supplier||'',form.supplierSku||'',form.productLink||'',form.active||'Yes',form.notes||'',new Date()];
  const s=sheet3_(FTP3.SHEETS.PACKAGING);s.getRange(Math.max(s.getLastRow()+1,2),1,1,values.length).setValues([values]);refreshDashboardSprint3();
}

function packagingChoices3_() {
  const s=sheet3_(FTP3.SHEETS.PACKAGING); if(s.getLastRow()<2)return [];
  return s.getRange(2,1,s.getLastRow()-1,FTP3.PACKAGING_HEADERS.length).getValues().filter(r=>r[0]&&String(r[13]||'Yes')!=='No').map(r=>({
    id:String(r[0]),category:String(r[1]||''),label:String(r[0])+' — '+String(r[2]||r[1]||'Packaging')+(r[3]?' ('+r[3]+')':''),
    unit:String(r[4]||'Each'),cost:num3_(r[7]),available:num3_(r[8])
  }));
}

function packagingItemById3_(id) {
  if(!id)return null; const s=sheet3_(FTP3.SHEETS.PACKAGING);if(s.getLastRow()<2)return null;
  const ids=s.getRange(2,1,s.getLastRow()-1,1).getDisplayValues().flat();const index=ids.findIndex(x=>x===id);
  if(index<0)return null;const row=index+2;return{row:row,values:s.getRange(row,1,1,FTP3.PACKAGING_HEADERS.length).getValues()[0]};
}

function calculatePackagingUsage3_(form) {
  const specs=[['boxId','boxQty','Box'],['bubbleId','bubbleQty','Bubble Wrap'],['mailerId','mailerQty','Mailer'],['tapeId','tapeQty','Tape'],['otherPackagingId','otherPackagingQty','Other']];
  const usage=[];let total=0;
  specs.forEach(([idKey,qtyKey,label])=>{const id=String(form[idKey]||'').trim();const qty=num3_(form[qtyKey]);if(!id&&qty>0)throw new Error(label+' quantity was entered without selecting an item.');if(!id)return;if(qty<=0)throw new Error(label+' quantity must be greater than zero.');const item=packagingItemById3_(id);if(!item)throw new Error(label+' packaging item was not found: '+id);const available=num3_(item.values[8]);if(qty>available+0.000001)throw new Error(label+' quantity exceeds stock. Available: '+available+' '+String(item.values[4]||'units'));const cost=num3_(item.values[7])*qty;usage.push({id:id,qty:qty,row:item.row,cost:cost});total+=cost;});
  return{usage:usage,total:Math.round(total*100)/100};
}

function deductPackagingUsage3_(usage) {
  const s=sheet3_(FTP3.SHEETS.PACKAGING);usage.forEach(u=>{const current=num3_(s.getRange(u.row,9).getValue());if(u.qty>current+0.000001)throw new Error('Packaging stock changed before the sale was saved. Reopen the sale form.');s.getRange(u.row,9).setValue(current-u.qty);s.getRange(u.row,16).setValue(new Date());});
}

function goToPackagingInventory3_(){SpreadsheetApp.getActive().setActiveSheet(sheet3_(FTP3.SHEETS.PACKAGING));}

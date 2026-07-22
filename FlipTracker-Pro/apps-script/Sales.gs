function buildSalesSprint3_() {
  migrateSheetHeaders3_(FTP3.SHEETS.SALES,FTP3.SALES_HEADERS,{});
  const s=sheet3_(FTP3.SHEETS.SALES);ensureSize3_(s,FTP3.ROWS+1,FTP3.SALES_HEADERS.length);
  s.getRange(1,1,1,FTP3.SALES_HEADERS.length).setValues([FTP3.SALES_HEADERS]);header3_(s.getRange(1,1,1,FTP3.SALES_HEADERS.length));s.setFrozenRows(1);
  setValidation3_(s,4,'FTP3_Marketplaces',FTP3.ROWS);s.getRange(2,3,FTP3.ROWS,1).setNumberFormat('yyyy-mm-dd');s.getRange(2,5,FTP3.ROWS,13).setNumberFormat('$#,##0.00;[Red]-$#,##0.00');s.getRange(2,18,FTP3.ROWS,1).setNumberFormat('0.0%;[Red]-0.0%');s.getRange(2,23,FTP3.ROWS,1).setNumberFormat('yyyy-mm-dd hh:mm');
  s.getRange(2,25,FTP3.ROWS,1).setNumberFormat('0.000');s.getRange(2,27,FTP3.ROWS,1).setNumberFormat('0.000');s.getRange(2,29,FTP3.ROWS,1).setNumberFormat('0.000');s.getRange(2,31,FTP3.ROWS,1).setNumberFormat('0.000');s.getRange(2,33,FTP3.ROWS,1).setNumberFormat('0.000');
  if(s.getFilter())s.getFilter().remove();s.getRange(1,1,FTP3.ROWS+1,FTP3.SALES_HEADERS.length).createFilter();borders3_(s.getRange(1,1,Math.min(FTP3.ROWS+1,200),FTP3.SALES_HEADERS.length));
  if(SpreadsheetApp.getActive().getSheetByName(FTP3.SHEETS.PACKAGING))refreshPackagingDropdowns3_();
}

function showRecordSaleForm() {showRecordSaleFormForItem3_('');}

function completeSelectedPendingSale3_() {
  const ss=SpreadsheetApp.getActive();
  const range=ss.getActiveRange();
  const sheet=range && range.getSheet();
  if(!range || !sheet || sheet.getName()!==FTP3.SHEETS.INVENTORY || range.getRow()<2){
    SpreadsheetApp.getUi().alert('Select the Sale Pending item row on the Inventory sheet first.');
    return;
  }
  const row=range.getRow();
  const itemId=sheet.getRange(row,1).getDisplayValue();
  const status=sheet.getRange(row,19).getDisplayValue();
  if(!itemId){
    SpreadsheetApp.getUi().alert('The selected Inventory row does not have an Item ID.');
    return;
  }
  if(saleExistsForInventory3_(itemId)){
    sheet.getRange(row,19).setValue('Sold').clearNote();
    SpreadsheetApp.getUi().alert('A Sales record already exists for '+itemId+'. The Inventory status was set to Sold.');
    return;
  }
  if(status!=='Sale Pending'){
    const response=SpreadsheetApp.getUi().alert(
      'Start sale?',
      'The selected item is currently "'+status+'". Change it to Sale Pending and open the sale form?',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    if(response!==SpreadsheetApp.getUi().Button.YES)return;
    PropertiesService.getDocumentProperties().setProperty('FTP_PREV_STATUS_'+itemId,status||'Listed');
    sheet.getRange(row,19).setValue('Sale Pending');
  }
  showRecordSaleFormForItem3_(itemId);
}
function showRecordSaleFormForItem3_(itemId) {
  const items=activeInventoryChoices3_();
  if(itemId && !items.some(x=>x.id===itemId)){
    const item=inventoryItemById3_(itemId);
    if(item)items.unshift({id:itemId,label:itemId+' — '+item.values[2]});
  }
  if(!items.length){SpreadsheetApp.getUi().alert('There are no active inventory items available to sell.');return;}
  const html=HtmlService.createHtmlOutput(saleFormHtml3_(items,packagingChoices3_(),itemId)).setWidth(680).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html,itemId?'Complete Sale':'Record Sale');
}
function saleFormHtml3_(items,packages,selectedItemId) {
  const itemOptions=items.map(x=>`<option value="${x.id}" ${x.id===selectedItemId?'selected':''}>${x.label}</option>`).join('');
  const makeOptions=(category)=>'<option value="">None</option>'+packages.filter(x=>category==='Other'||x.category.toLowerCase().indexOf(category.toLowerCase())>=0).map(x=>`<option value="${x.id}" data-cost="${x.cost}" data-stock="${x.available}" data-unit="${x.unit}">${x.label} — ${x.available} ${x.unit} @ $${x.cost.toFixed(3)}</option>`).join('');
  return `<!doctype html><html><head><base target="_top"><style>body{font-family:Arial;padding:14px;color:#1F2937}label{display:block;font-weight:700;margin-top:8px}input,select,textarea{width:100%;box-sizing:border-box;padding:8px;margin-top:3px;border:1px solid #B7C9D6;border-radius:4px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}.section{margin-top:14px;padding:10px;background:#F3F6F9;border-radius:6px}.summary{font-weight:700;margin-top:10px}button{margin-top:16px;padding:10px 16px;border:0;border-radius:4px;background:#1F4E78;color:#fff;font-weight:700}</style></head><body><form id="f">
  <label>Inventory item</label><select name="itemId" required>${itemOptions}</select><div class="grid"><div><label>Sale date</label><input type="date" name="saleDate" required></div><div><label>Marketplace</label><input name="marketplace" required></div><div><label>Sale price</label><input type="number" step="0.01" min="0" name="salePrice" required></div><div><label>Shipping charged</label><input type="number" step="0.01" min="0" name="shippingCharged"></div><div><label>Shipping actual</label><input type="number" step="0.01" min="0" name="shippingActual"></div><div><label>Marketplace fees</label><input type="number" step="0.01" min="0" name="marketplaceFees"></div><div><label>Payment fees</label><input type="number" step="0.01" min="0" name="paymentFees"></div><div><label>Promotion expense</label><input type="number" step="0.01" min="0" name="promotionExpense"></div><div><label>GST/HST collected</label><input type="number" step="0.01" min="0" name="taxCollected"></div></div>
  <div class="section"><b>Packaging used</b><div class="grid"><div><label>Box used</label><select name="boxId" class="pkg">${makeOptions('Box')}</select></div><div><label>Box quantity</label><input name="boxQty" class="qty" type="number" min="0" step="0.001" value="0"></div><div><label>Bubble wrap used</label><select name="bubbleId" class="pkg">${makeOptions('Bubble')}</select></div><div><label>Bubble wrap quantity</label><input name="bubbleQty" class="qty" type="number" min="0" step="0.001" value="0"></div><div><label>Mailer used</label><select name="mailerId" class="pkg">${makeOptions('Mailer')}</select></div><div><label>Mailer quantity</label><input name="mailerQty" class="qty" type="number" min="0" step="0.001" value="0"></div><div><label>Tape used</label><select name="tapeId" class="pkg">${makeOptions('Tape')}</select></div><div><label>Tape quantity</label><input name="tapeQty" class="qty" type="number" min="0" step="0.001" value="0"></div><div><label>Other packaging</label><select name="otherPackagingId" class="pkg">${makeOptions('Other')}</select></div><div><label>Other quantity</label><input name="otherPackagingQty" class="qty" type="number" min="0" step="0.001" value="0"></div></div><div class="summary">Calculated packaging cost: <span id="pkgCost">$0.00</span></div></div>
  <div class="grid"><div><label>Buyer</label><input name="buyer"></div><div><label>Tracking number</label><input name="trackingNumber"></div></div><label>Notes</label><textarea name="notes" rows="3"></textarea><button type="submit">Record Sale</button><button type="button" style="background:#6B7280;margin-left:8px" onclick="cancelSale()">Cancel</button></form><script>
  document.querySelector('[name="saleDate"]').value=new Date().toISOString().slice(0,10);function cancelSale(){const id=document.querySelector('[name="itemId"]').value;google.script.run.withSuccessHandler(()=>google.script.host.close()).cancelPendingSale3_(id);}function calc(){let total=0;[['boxId','boxQty'],['bubbleId','bubbleQty'],['mailerId','mailerQty'],['tapeId','tapeQty'],['otherPackagingId','otherPackagingQty']].forEach(([s,q])=>{const sel=document.querySelector('[name="'+s+'"]');const opt=sel.options[sel.selectedIndex];total+=(Number(opt.dataset.cost)||0)*(Number(document.querySelector('[name="'+q+'"]').value)||0);});document.getElementById('pkgCost').textContent='$'+total.toFixed(2);}document.querySelectorAll('.pkg,.qty').forEach(x=>x.addEventListener('input',calc));
  document.getElementById('f').addEventListener('submit',e=>{e.preventDefault();google.script.run.withSuccessHandler(()=>google.script.host.close()).withFailureHandler(x=>alert(x.message)).saveSale3_(Object.fromEntries(new FormData(e.target).entries()));});</script></body></html>`;
}

function saveSale3_(form) {
  const lock=LockService.getDocumentLock();lock.waitLock(30000);
  try{
    const item=inventoryItemById3_(form.itemId);if(!item)throw new Error('The selected inventory item was not found.');const inv=item.values;if(saleExistsForInventory3_(form.itemId))throw new Error('A completed Sales record already exists for this inventory item.');
    const pkg=calculatePackagingUsage3_(form);const salePrice=num3_(form.salePrice),shippingCharged=num3_(form.shippingCharged),shippingActual=num3_(form.shippingActual),marketFees=num3_(form.marketplaceFees),paymentFees=num3_(form.paymentFees),promotion=num3_(form.promotionExpense),taxCollected=num3_(form.taxCollected),itemCost=num3_(inv[13]);
    if(!form.saleDate)throw new Error('Sale date is required.');if(salePrice<0)throw new Error('Sale price cannot be negative.');
    const grossRevenue=salePrice+shippingCharged,sellingCosts=shippingActual+pkg.total+marketFees+paymentFees+promotion,netProceeds=grossRevenue-sellingCosts,realizedProfit=netProceeds-itemCost,roi=itemCost?realizedProfit/itemCost:'';
    const saleDate=date3_(form.saleDate),purchaseDate=inv[1] instanceof Date?inv[1]:date3_(inv[1]),days=purchaseDate?Math.max(0,Math.floor((saleDate-purchaseDate)/86400000)):'';
    const values=[nextId3_(FTP3.SHEETS.SALES,1,'SAL'),form.itemId,saleDate,form.marketplace||'',salePrice,shippingCharged,shippingActual,pkg.total,marketFees,paymentFees,promotion,taxCollected,itemCost,grossRevenue,sellingCosts,netProceeds,realizedProfit,roi,days,form.buyer||'',form.trackingNumber||'',form.notes||'',new Date(),packagingIdFromSelection3_(form.boxId),num3_(form.boxQty),packagingIdFromSelection3_(form.bubbleId),num3_(form.bubbleQty),packagingIdFromSelection3_(form.mailerId),num3_(form.mailerQty),packagingIdFromSelection3_(form.tapeId),num3_(form.tapeQty),packagingIdFromSelection3_(form.otherPackagingId),num3_(form.otherPackagingQty),'Yes'];
    const sales=sheet3_(FTP3.SHEETS.SALES),row=Math.max(sales.getLastRow()+1,2);sales.getRange(row,1,1,values.length).setValues([values]);
    try{deductPackagingUsage3_(pkg.usage);}catch(err){sales.deleteRow(row);throw err;}
    const inventory=sheet3_(FTP3.SHEETS.INVENTORY);inventory.getRange(item.row,19).setValue('Sold').clearNote();inventory.getRange(item.row,27).setValue(new Date());PropertiesService.getDocumentProperties().deleteProperty('FTP_PREV_STATUS_'+form.itemId);refreshDashboardSprint3();
  }finally{lock.releaseLock();}
}

function saleExistsForInventory3_(itemId){
  const s=sheet3_(FTP3.SHEETS.SALES);
  if(s.getLastRow()<2)return false;
  return s.getRange(2,2,s.getLastRow()-1,1).getDisplayValues().flat().some(v=>v===itemId);
}
function cancelPendingSale3_(itemId){
  const item=inventoryItemById3_(itemId);
  if(item && String(item.values[18])==='Sale Pending'){
    const previous=PropertiesService.getDocumentProperties().getProperty('FTP_PREV_STATUS_'+itemId)||'Listed';
    sheet3_(FTP3.SHEETS.INVENTORY).getRange(item.row,19).setValue(previous).clearNote();
    PropertiesService.getDocumentProperties().deleteProperty('FTP_PREV_STATUS_'+itemId);
  }
}

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
  const html=HtmlService.createHtmlOutput(saleFormHtml3_(items,packagingChoices3_(),marketplaceChoices3_(),itemId)).setWidth(680).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html,itemId?'Complete Sale':'Record Sale');
}
function saleFormHtml3_(items,packages,marketplaces,selectedItemId) {
  const esc=v=>String(v==null?'':v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
  const itemOptions=items.map(x=>`<option value="${esc(x.id)}" ${x.id===selectedItemId?'selected':''}>${esc(x.label)}</option>`).join('');
  const marketplaceOptions='<option value="">Select marketplace</option>'+marketplaces.map(x=>`<option value="${esc(x)}">${esc(x)}</option>`).join('');
  const makeOptions=(category)=>'<option value="">None</option>'+packages.filter(x=>category==='Other'||x.category.toLowerCase().indexOf(category.toLowerCase())>=0).map(x=>`<option value="${esc(x.id)}" data-cost="${Number(x.cost)||0}" data-stock="${Number(x.available)||0}" data-unit="${esc(x.unit)}">${esc(x.label)} — ${Number(x.available)||0} ${esc(x.unit)} @ $${(Number(x.cost)||0).toFixed(3)}</option>`).join('');
  return `<!doctype html><html><head><base target="_top"><style>
  body{font-family:Arial;padding:14px;color:#1F2937}label{display:block;font-weight:700;margin-top:8px}input,select,textarea{width:100%;box-sizing:border-box;padding:8px;margin-top:3px;border:1px solid #B7C9D6;border-radius:4px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}.section{margin-top:14px;padding:10px;background:#F3F6F9;border-radius:6px}.summary{font-weight:700;margin-top:10px}.actions{display:flex;align-items:center;gap:8px;margin-top:16px}.actions button{padding:10px 16px;border:0;border-radius:4px;background:#1F4E78;color:#fff;font-weight:700;cursor:pointer}.actions button:disabled{opacity:.55;cursor:not-allowed}.actions .cancel{background:#6B7280}.status{margin-top:10px;min-height:20px;font-weight:700}.error{color:#B91C1C}.success{color:#166534}</style></head><body><form id="f" novalidate>
  <label>Inventory item</label><select name="itemId" required>${itemOptions}</select><div class="grid"><div><label>Sale date</label><input type="date" name="saleDate" required></div><div><label>Marketplace</label><select name="marketplace" required>${marketplaceOptions}</select></div><div><label>Sale price</label><input type="number" step="0.01" min="0" name="salePrice" required></div><div><label>Shipping charged</label><input type="number" step="0.01" min="0" name="shippingCharged"></div><div><label>Shipping actual</label><input type="number" step="0.01" min="0" name="shippingActual"></div><div><label>Marketplace fees</label><input type="number" step="0.01" min="0" name="marketplaceFees"></div><div><label>Payment fees</label><input type="number" step="0.01" min="0" name="paymentFees"></div><div><label>Promotion expense</label><input type="number" step="0.01" min="0" name="promotionExpense"></div><div><label>GST/HST collected</label><input type="number" step="0.01" min="0" name="taxCollected"></div></div>
  <div class="section"><b>Packaging used</b><div class="grid"><div><label>Box used</label><select name="boxId" class="pkg">${makeOptions('Box')}</select></div><div><label>Box quantity</label><input name="boxQty" class="qty" type="number" min="0" step="0.001" value="0"></div><div><label>Bubble wrap used</label><select name="bubbleId" class="pkg">${makeOptions('Bubble')}</select></div><div><label>Bubble wrap quantity</label><input name="bubbleQty" class="qty" type="number" min="0" step="0.001" value="0"></div><div><label>Mailer used</label><select name="mailerId" class="pkg">${makeOptions('Mailer')}</select></div><div><label>Mailer quantity</label><input name="mailerQty" class="qty" type="number" min="0" step="0.001" value="0"></div><div><label>Tape used</label><select name="tapeId" class="pkg">${makeOptions('Tape')}</select></div><div><label>Tape quantity</label><input name="tapeQty" class="qty" type="number" min="0" step="0.001" value="0"></div><div><label>Other packaging</label><select name="otherPackagingId" class="pkg">${makeOptions('Other')}</select></div><div><label>Other quantity</label><input name="otherPackagingQty" class="qty" type="number" min="0" step="0.001" value="0"></div></div><div class="summary">Calculated packaging cost: <span id="pkgCost">$0.00</span></div></div>
  <div class="grid"><div><label>Buyer</label><input name="buyer"></div><div><label>Tracking number</label><input name="trackingNumber"></div></div><label>Notes</label><textarea name="notes" rows="3"></textarea>
  <div class="actions"><button id="saveBtn" type="button" onclick="submitSale()">Accept Sale</button><button id="cancelBtn" class="cancel" type="button" onclick="cancelSale()">Cancel</button></div><div id="status" class="status"></div></form><script>
  const form=document.getElementById('f'),saveBtn=document.getElementById('saveBtn'),cancelBtn=document.getElementById('cancelBtn'),statusBox=document.getElementById('status');
  document.querySelector('[name="saleDate"]').value=new Date().toISOString().slice(0,10);
  function showStatus(message,isError){statusBox.textContent=message||'';statusBox.className='status '+(isError?'error':'success');}
  function setBusy(busy){saveBtn.disabled=busy;cancelBtn.disabled=busy;saveBtn.textContent=busy?'Saving…':'Accept Sale';}
  function value(name){const el=form.elements[name];return el?String(el.value||'').trim():'';}
  function formDataObject(){const data={};Array.from(form.elements).forEach(el=>{if(el.name)data[el.name]=el.value;});return data;}
  function validate(){
    if(!value('itemId'))return 'Select an inventory item.';
    if(!value('saleDate'))return 'Enter the sale date.';
    if(!value('marketplace'))return 'Select the marketplace.';
    if(value('salePrice')==='')return 'Enter the sale price.';
    if(Number(value('salePrice'))<0)return 'Sale price cannot be negative.';
    const pairs=[['boxId','boxQty','box'],['bubbleId','bubbleQty','bubble wrap'],['mailerId','mailerQty','mailer'],['tapeId','tapeQty','tape'],['otherPackagingId','otherPackagingQty','other packaging']];
    for(const [id,qty,label] of pairs){const selected=value(id),amount=Number(value(qty)||0);if(amount<0)return 'Packaging quantities cannot be negative.';if(amount>0&&!selected)return 'Select the '+label+' item used, or set its quantity to zero.';}
    return '';
  }
  function submitSale(){
    showStatus('',false);const problem=validate();if(problem){showStatus(problem,true);return;}
    setBusy(true);showStatus('Saving sale…',false);
    let finished=false;
    const watchdog=setTimeout(()=>{
      if(finished)return;
      setBusy(false);
      showStatus('The save is taking too long. Check the Sales sheet before trying again because the sale may already have been recorded.',true);
    },20000);
    google.script.run
      .withSuccessHandler(result=>{finished=true;clearTimeout(watchdog);showStatus((result&&result.message)||'Sale saved successfully.',false);setTimeout(()=>google.script.host.close(),500);})
      .withFailureHandler(error=>{finished=true;clearTimeout(watchdog);setBusy(false);const message=error&&error.message?error.message:String(error||'The sale could not be saved.');showStatus(message,true);})
      .saveSale3_(formDataObject());
  }
  function cancelSale(){setBusy(true);const id=value('itemId');google.script.run.withSuccessHandler(()=>google.script.host.close()).withFailureHandler(error=>{setBusy(false);showStatus(error&&error.message?error.message:String(error),true);}).cancelPendingSale3_(id);}
  function calc(){let total=0;[['boxId','boxQty'],['bubbleId','bubbleQty'],['mailerId','mailerQty'],['tapeId','tapeQty'],['otherPackagingId','otherPackagingQty']].forEach(([s,q])=>{const sel=form.elements[s],opt=sel.options[sel.selectedIndex];total+=(Number(opt&&opt.dataset.cost)||0)*(Number(form.elements[q].value)||0);});document.getElementById('pkgCost').textContent='$'+total.toFixed(2);}
  document.querySelectorAll('.pkg,.qty').forEach(x=>{x.addEventListener('input',calc);x.addEventListener('change',calc);});
  form.addEventListener('submit',e=>{e.preventDefault();submitSale();});
  calc();
  </script></body></html>`;
}

function marketplaceChoices3_() {
  const ss=SpreadsheetApp.getActive();
  const named=ss.getRangeByName('FTP3_Marketplaces');
  if(!named)return ['eBay','Facebook Marketplace','Kijiji','Poshmark','Etsy','Local Sale','Other'];
  const values=named.getDisplayValues().flat().map(v=>String(v).trim()).filter(Boolean);
  return values.filter((v,i,a)=>a.indexOf(v)===i);
}

function saveSale3_(form) {
  const lock=LockService.getDocumentLock();if(!lock.tryLock(5000))throw new Error('FlipTracker Pro is busy with another update. Please wait a few seconds and try again.');
  try{
    const item=inventoryItemById3_(form.itemId);if(!item)throw new Error('The selected inventory item was not found.');const inv=item.values;if(saleExistsForInventory3_(form.itemId))throw new Error('A completed Sales record already exists for this inventory item.');
    const pkg=calculatePackagingUsage3_(form);const salePrice=num3_(form.salePrice),shippingCharged=num3_(form.shippingCharged),shippingActual=num3_(form.shippingActual),marketFees=num3_(form.marketplaceFees),paymentFees=num3_(form.paymentFees),promotion=num3_(form.promotionExpense),taxCollected=num3_(form.taxCollected),itemCost=num3_(inv[13]);
    if(!form.saleDate)throw new Error('Sale date is required.');if(salePrice<0)throw new Error('Sale price cannot be negative.');
    const grossRevenue=salePrice+shippingCharged,sellingCosts=shippingActual+pkg.total+marketFees+paymentFees+promotion,netProceeds=grossRevenue-sellingCosts,realizedProfit=netProceeds-itemCost,roi=itemCost?realizedProfit/itemCost:'';
    const saleDate=date3_(form.saleDate),purchaseDate=inv[1] instanceof Date?inv[1]:date3_(inv[1]),days=purchaseDate?Math.max(0,Math.floor((saleDate-purchaseDate)/86400000)):'';
    const values=[nextId3_(FTP3.SHEETS.SALES,1,'SAL'),form.itemId,saleDate,form.marketplace||'',salePrice,shippingCharged,shippingActual,pkg.total,marketFees,paymentFees,promotion,taxCollected,itemCost,grossRevenue,sellingCosts,netProceeds,realizedProfit,roi,days,form.buyer||'',form.trackingNumber||'',form.notes||'',new Date(),packagingIdFromSelection3_(form.boxId),num3_(form.boxQty),packagingIdFromSelection3_(form.bubbleId),num3_(form.bubbleQty),packagingIdFromSelection3_(form.mailerId),num3_(form.mailerQty),packagingIdFromSelection3_(form.tapeId),num3_(form.tapeQty),packagingIdFromSelection3_(form.otherPackagingId),num3_(form.otherPackagingQty),'Yes'];
    const sales=sheet3_(FTP3.SHEETS.SALES),row=Math.max(sales.getLastRow()+1,2);sales.getRange(row,1,1,values.length).setValues([values]);
    try{deductPackagingUsage3_(pkg.usage);}catch(err){sales.deleteRow(row);throw err;}
    const inventory=sheet3_(FTP3.SHEETS.INVENTORY);inventory.getRange(item.row,19).setValue('Sold').clearNote();inventory.getRange(item.row,27).setValue(new Date());PropertiesService.getDocumentProperties().deleteProperty('FTP_PREV_STATUS_'+form.itemId);SpreadsheetApp.flush();
    return {ok:true,message:'Sale '+values[0]+' saved successfully.'};
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

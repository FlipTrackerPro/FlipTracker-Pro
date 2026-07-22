function buildSalesSprint3_() {
  migrateSheetHeaders3_(FTP3.SHEETS.SALES,FTP3.SALES_HEADERS,{});
  const s=sheet3_(FTP3.SHEETS.SALES);ensureSize3_(s,FTP3.ROWS+1,FTP3.SALES_HEADERS.length);
  s.getRange(1,1,1,FTP3.SALES_HEADERS.length).setValues([FTP3.SALES_HEADERS]);header3_(s.getRange(1,1,1,FTP3.SALES_HEADERS.length));s.setFrozenRows(1);
  setValidation3_(s,5,'FTP3_Marketplaces',FTP3.ROWS);s.getRange(2,4,FTP3.ROWS,1).setNumberFormat('yyyy-mm-dd');s.getRange(2,6,FTP3.ROWS,13).setNumberFormat('$#,##0.00;[Red]-$#,##0.00');s.getRange(2,19,FTP3.ROWS,1).setNumberFormat('0.0%;[Red]-0.0%');s.getRange(2,24,FTP3.ROWS,1).setNumberFormat('yyyy-mm-dd hh:mm');
  s.getRange(2,26,FTP3.ROWS,1).setNumberFormat('0.000');s.getRange(2,28,FTP3.ROWS,1).setNumberFormat('0.000');s.getRange(2,30,FTP3.ROWS,1).setNumberFormat('0.000');s.getRange(2,32,FTP3.ROWS,1).setNumberFormat('0.000');s.getRange(2,34,FTP3.ROWS,1).setNumberFormat('0.000');
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
  const invMap=headerMap3_(sheet);const status=sheet.getRange(row,invMap['Status']).getDisplayValue();
  if(!itemId){
    SpreadsheetApp.getUi().alert('The selected Inventory row does not have an Item ID.');
    return;
  }
  if(saleExistsForInventory3_(itemId)){
    sheet.getRange(row,invMap['Status']).setValue('Sold').clearNote();
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
    sheet.getRange(row,invMap['Status']).setValue('Sale Pending');
  }
  showRecordSaleFormForItem3_(itemId);
}
function showRecordSaleFormForItem3_(itemId) {
  const items=activeInventoryChoices3_();
  if(itemId && !items.some(x=>x.id===itemId)){
    const item=inventoryItemById3_(itemId);
    if(item)items.unshift({id:itemId,description:String(rowRecord3_(FTP3.INVENTORY_HEADERS,item.values)['Description']||'')});
  }
  if(!items.length){SpreadsheetApp.getUi().alert('There are no active inventory items available to sell.');return;}
  const html=HtmlService.createHtmlOutput(saleFormHtml3_(items,packagingChoices3_(),marketplaceChoices3_(),itemId)).setWidth(680).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html,itemId?'Complete Sale':'Record Sale');
}
function saleFormHtml3_(items,packages,marketplaces,selectedItemId) {
  const esc=v=>String(v==null?'':v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
  const itemOptions=items.map(x=>`<option value="${esc(x.id)}" data-description="${esc(x.description||'')}" ${x.id===selectedItemId?'selected':''}>${esc(x.id)}</option>`).join('');
  const marketplaceOptions='<option value="">Select marketplace</option>'+marketplaces.map(x=>`<option value="${esc(x)}">${esc(x)}</option>`).join('');
  const makeOptions=(category)=>'<option value="">None</option>'+packages.filter(x=>category==='Other'||x.category.toLowerCase().indexOf(category.toLowerCase())>=0).map(x=>`<option value="${esc(x.id)}" data-cost="${Number(x.cost)||0}" data-stock="${Number(x.available)||0}" data-unit="${esc(x.unit)}" data-label="${esc(x.label)}">${esc(x.id)}</option>`).join('');
  return `<!doctype html><html><head><base target="_top"><style>
  body{font-family:Arial;padding:14px;color:#1F2937}label{display:block;font-weight:700;margin-top:8px}input,select,textarea{width:100%;box-sizing:border-box;padding:8px;margin-top:3px;border:1px solid #B7C9D6;border-radius:4px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}.section{margin-top:14px;padding:10px;background:#F3F6F9;border-radius:6px}.summary{font-weight:700;margin-top:10px}.packageDetails{font-size:12px;color:#4B5563;margin-top:4px;min-height:16px}.actions{display:flex;align-items:center;gap:8px;margin-top:16px}.actions button{padding:10px 16px;border:0;border-radius:4px;background:#1F4E78;color:#fff;font-weight:700;cursor:pointer}.actions button:disabled{opacity:.55;cursor:not-allowed}.actions .cancel{background:#6B7280}.status{margin-top:10px;min-height:20px;font-weight:700}.error{color:#B91C1C}.success{color:#166534}</style></head><body><form id="f" novalidate>
  <div class="grid"><div><label>Item ID</label><select name="itemId" required>${itemOptions}</select></div><div><label>Description</label><input name="description" readonly></div></div><div class="grid"><div><label>Sale date</label><input type="date" name="saleDate" required></div><div><label>Marketplace</label><select name="marketplace" required>${marketplaceOptions}</select></div><div><label>Sale price</label><input type="number" step="0.01" min="0" name="salePrice" required></div><div><label>Shipping charged</label><input type="number" step="0.01" min="0" name="shippingCharged"></div><div><label>Shipping actual</label><input type="number" step="0.01" min="0" name="shippingActual"></div><div><label>Marketplace fees</label><input type="number" step="0.01" min="0" name="marketplaceFees"></div><div><label>Payment fees</label><input type="number" step="0.01" min="0" name="paymentFees"></div><div><label>Promotion expense</label><input type="number" step="0.01" min="0" name="promotionExpense"></div><div><label>GST/HST collected</label><input type="number" step="0.01" min="0" name="taxCollected"></div></div>
  <div class="section"><b>Packaging used</b><div class="grid"><div><label>Box used (Packaging ID)</label><select name="boxId" class="pkg">${makeOptions('Box')}</select><div id="boxDetails" class="packageDetails"></div></div><div><label>Box quantity</label><input name="boxQty" class="qty" type="number" min="0" step="0.001" value="0"></div><div><label>Bubble wrap used</label><select name="bubbleId" class="pkg">${makeOptions('Bubble')}</select></div><div><label>Bubble wrap quantity</label><input name="bubbleQty" class="qty" type="number" min="0" step="0.001" value="0"></div><div><label>Mailer used</label><select name="mailerId" class="pkg">${makeOptions('Mailer')}</select></div><div><label>Mailer quantity</label><input name="mailerQty" class="qty" type="number" min="0" step="0.001" value="0"></div><div><label>Tape used</label><select name="tapeId" class="pkg">${makeOptions('Tape')}</select></div><div><label>Tape quantity</label><input name="tapeQty" class="qty" type="number" min="0" step="0.001" value="0"></div><div><label>Other packaging</label><select name="otherPackagingId" class="pkg">${makeOptions('Other')}</select></div><div><label>Other quantity</label><input name="otherPackagingQty" class="qty" type="number" min="0" step="0.001" value="0"></div></div><div class="summary">Calculated packaging cost: <span id="pkgCost">$0.00</span></div></div>
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
      .saveSale3(formDataObject());
  }
  function cancelSale(){setBusy(true);const id=value('itemId');google.script.run.withSuccessHandler(()=>google.script.host.close()).withFailureHandler(error=>{setBusy(false);showStatus(error&&error.message?error.message:String(error),true);}).cancelPendingSale3(id);}
  function syncDescription(){
    const sel=form.elements.itemId;
    const opt=sel.options[sel.selectedIndex];
    form.elements.description.value=opt?String(opt.dataset.description||''):'';
  }
  form.elements.itemId.addEventListener('change',syncDescription);
  function calc(){let total=0;[['boxId','boxQty'],['bubbleId','bubbleQty'],['mailerId','mailerQty'],['tapeId','tapeQty'],['otherPackagingId','otherPackagingQty']].forEach(([s,q])=>{const sel=form.elements[s],opt=sel.options[sel.selectedIndex];total+=(Number(opt&&opt.dataset.cost)||0)*(Number(form.elements[q].value)||0);});document.getElementById('pkgCost').textContent='$'+total.toFixed(2);const box=form.elements.boxId.options[form.elements.boxId.selectedIndex];document.getElementById('boxDetails').textContent=box&&box.value?(box.dataset.label+' | $'+(Number(box.dataset.cost)||0).toFixed(3)+' per '+(box.dataset.unit||'unit')+' | '+(Number(box.dataset.stock)||0)+' available'):'';}
  document.querySelectorAll('.pkg,.qty').forEach(x=>{x.addEventListener('input',calc);x.addEventListener('change',calc);});
  form.addEventListener('submit',e=>{e.preventDefault();submitSale();});
  syncDescription();
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

function saveSale3(form) { return saveSale3_(form); }
function cancelPendingSale3(itemId) { return cancelPendingSale3_(itemId); }

function saveSale3_(form) {
  const itemId=String(form&&form.itemId||'').trim();
  if(!itemId)throw new Error('Item ID is missing from the Complete Sale form.');
  const lock=LockService.getDocumentLock();
  if(!lock.tryLock(5000))throw new Error('FlipTracker Pro is busy. Wait a few seconds and try again.');
  try{
    const item=inventoryItemById3_(itemId);
    if(!item)throw new Error('Inventory Item ID '+itemId+' was not found.');
    if(saleExistsForInventory3_(itemId))throw new Error('A Sales record already exists for '+itemId+'.');
    const inv=item.values;
    const saleDate=date3_(form.saleDate);
    if(!saleDate || isNaN(saleDate.getTime()))throw new Error('Enter a valid sale date.');
    const salePrice=num3_(form.salePrice);
    if(salePrice<0)throw new Error('Sale price cannot be negative.');
    const shippingCharged=num3_(form.shippingCharged);
    const shippingActual=num3_(form.shippingActual);
    const marketFees=num3_(form.marketplaceFees);
    const paymentFees=num3_(form.paymentFees);
    const promotion=num3_(form.promotionExpense);
    const taxCollected=num3_(form.taxCollected);
    const invRecord=rowRecord3_(FTP3.INVENTORY_HEADERS,inv);const itemCost=num3_(invRecord['Total Cost']);
    const pkg=calculatePackagingUsage3_(form);
    const grossRevenue=salePrice+shippingCharged;
    const sellingCosts=shippingActual+pkg.total+marketFees+paymentFees+promotion;
    const netProceeds=grossRevenue-sellingCosts;
    const realizedProfit=netProceeds-itemCost;
    const roi=itemCost?realizedProfit/itemCost:'';
    const purchaseValue=invRecord['Purchase Date'];const purchaseDate=purchaseValue instanceof Date?purchaseValue:date3_(purchaseValue);
    const days=purchaseDate&&!isNaN(purchaseDate.getTime())?Math.max(0,Math.floor((saleDate-purchaseDate)/86400000)):'';
    const saleId=nextId3_(FTP3.SHEETS.SALES,1,'SAL');
    const record={
      'Sale ID':saleId,'Item ID':itemId,'Description':String(invRecord['Description']||''),
      'Sale Date':saleDate,'Marketplace':String(form.marketplace||''),'Sale Price':salePrice,
      'Shipping Charged':shippingCharged,'Shipping Actual':shippingActual,'Packaging Cost':pkg.total,
      'Marketplace Fees':marketFees,'Payment Fees':paymentFees,'Promotion Expense':promotion,
      'GST/HST Collected':taxCollected,'Item Cost':itemCost,'Gross Revenue':grossRevenue,
      'Total Selling Costs':sellingCosts,'Net Proceeds':netProceeds,'Realized Profit':realizedProfit,
      'Realized ROI %':roi,'Days to Sell':days,'Buyer':String(form.buyer||''),
      'Tracking Number':String(form.trackingNumber||''),'Notes':String(form.notes||''),'Created At':new Date(),
      'Box Used':packagingIdFromSelection3_(form.boxId),'Box Qty':num3_(form.boxQty),
      'Bubble Wrap Used':packagingIdFromSelection3_(form.bubbleId),'Bubble Wrap Qty':num3_(form.bubbleQty),
      'Mailer Used':packagingIdFromSelection3_(form.mailerId),'Mailer Qty':num3_(form.mailerQty),
      'Tape Used':packagingIdFromSelection3_(form.tapeId),'Tape Qty':num3_(form.tapeQty),
      'Other Packaging Used':packagingIdFromSelection3_(form.otherPackagingId),'Other Packaging Qty':num3_(form.otherPackagingQty),
      'Packaging Verified':'Yes'
    };
    const sales=sheet3_(FTP3.SHEETS.SALES);
    const headers=sales.getRange(1,1,1,sales.getLastColumn()).getDisplayValues()[0].map(h=>String(h).trim());
    const missing=FTP3.SALES_HEADERS.filter(h=>headers.indexOf(h)<0);
    if(missing.length)throw new Error('Sales sheet is missing required columns: '+missing.join(', ')+'. Run upgradeFlipTrackerPro() and try again.');
    const row=Math.max(sales.getLastRow()+1,2);
    const values=headers.map(h=>Object.prototype.hasOwnProperty.call(record,h)?record[h]:'');
    sales.getRange(row,1,1,headers.length).setValues([values]);
    SpreadsheetApp.flush();
    try{
      deductPackagingUsage3_(pkg.usage);
      const inventory=sheet3_(FTP3.SHEETS.INVENTORY);
      const invHeaders=headerMap3_(inventory);
      inventory.getRange(item.row,invHeaders['Status']||19).setValue('Sold').clearNote();
      inventory.getRange(item.row,invHeaders['Updated At']||27).setValue(new Date());
      PropertiesService.getDocumentProperties().deleteProperty('FTP_PREV_STATUS_'+itemId);
      SpreadsheetApp.flush();
    }catch(postWriteError){
      sales.deleteRow(row);
      throw new Error('The sale was not completed and was rolled back: '+postWriteError.message);
    }
    return {ok:true,saleId:saleId,row:row,message:'Sale '+saleId+' recorded on the Sales sheet.'};
  }catch(err){
    console.error('saveSale3_ failed: '+(err&&err.stack?err.stack:err));
    throw err;
  }finally{
    lock.releaseLock();
  }
}

function saleExistsForInventory3_(itemId){
  const s=sheet3_(FTP3.SHEETS.SALES);
  if(s.getLastRow()<2)return false;
  return s.getRange(2,2,s.getLastRow()-1,1).getDisplayValues().flat().some(v=>v===itemId);
}
function cancelPendingSale3_(itemId){
  const item=inventoryItemById3_(itemId);
  if(item && String(rowRecord3_(FTP3.INVENTORY_HEADERS,item.values)['Status'])==='Sale Pending'){
    const previous=PropertiesService.getDocumentProperties().getProperty('FTP_PREV_STATUS_'+itemId)||'Listed';
    const invSheet=sheet3_(FTP3.SHEETS.INVENTORY);const invMap=headerMap3_(invSheet);invSheet.getRange(item.row,invMap['Status']).setValue(previous).clearNote();
    PropertiesService.getDocumentProperties().deleteProperty('FTP_PREV_STATUS_'+itemId);
  }
}

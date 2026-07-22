function showAddItemForm(){showInventoryForm3_(null,'Add Inventory Item');}
function showEditSelectedItemForm(){
  const row=selectedInventoryRow3_(),s=sheet3_(FTP3.SHEETS.INVENTORY);
  showInventoryForm3_({row:row,values:s.getRange(row,1,1,FTP3.INVENTORY_HEADERS.length).getValues()[0]},'Edit Inventory Item');
}
function showInventoryForm3_(existing,title){
  const html=HtmlService.createHtmlOutput(inventoryFormHtml3_(existing)).setWidth(560).setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html,title);
}
function inventoryFormHtml3_(existing){
  const payload=existing?{row:existing.row,record:rowRecord3_(FTP3.INVENTORY_HEADERS,existing.values)}:null;
  const data=JSON.stringify(payload).replace(/</g,'\\u003c');
  return `<!doctype html><html><head><base target="_top"><style>
  body{font-family:Arial;padding:14px;color:#1F2937}label{display:block;font-weight:700;margin-top:8px}
  input,textarea{width:100%;box-sizing:border-box;padding:8px;margin-top:3px;border:1px solid #B7C9D6;border-radius:4px}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}.actions{display:flex;gap:8px;margin-top:16px}
  button{padding:10px 16px;border:0;border-radius:4px;background:#1F4E78;color:#fff;font-weight:700}.secondary{background:#6B7280}</style></head><body>
  <form id="f"><input type="hidden" name="row"><div class="grid">
  <div><label>Description</label><input name="description" required></div>
  <div><label>Category</label><input name="category"></div>
  <div><label>Condition</label><input name="condition"></div>
  <div><label>Quantity</label><input type="number" min="1" name="quantity" value="1"></div>
  <div><label>SKU</label><input name="sku"></div><div><label>Barcode</label><input name="barcode"></div>
  <div><label>Purchase date</label><input type="date" name="purchaseDate" required></div>
  <div><label>Purchase location</label><input name="purchaseLocation"></div>
  <div><label>Purchase price</label><input type="number" step="0.01" min="0" name="purchasePrice"></div>
  <div><label>Tax paid</label><input type="number" step="0.01" min="0" name="taxPaid"></div>
  <div><label>Acquisition shipping</label><input type="number" step="0.01" min="0" name="acquisitionShipping"></div>
  <div><label>Storage location</label><input name="storageLocation"></div>
  <div><label>Status</label><input name="status" value="Purchased"></div>
  <div><label>Marketplace</label><input name="marketplace"></div>
  <div><label>Listed price</label><input type="number" step="0.01" min="0" name="listedPrice"></div>
  <div><label>Expected sale price</label><input type="number" step="0.01" min="0" name="expectedSalePrice"></div></div>
  <label>Receipt link</label><input type="url" name="receiptLink"><label>Photo link</label><input type="url" name="photoLink">
  <label>Notes</label><textarea name="notes" rows="3"></textarea>
  <div class="actions"><button type="submit">Save Item</button><button type="button" class="secondary" onclick="google.script.host.close()">Cancel</button></div>
  </form><script>
  const old=${data};
  const iso=d=>{if(!d)return'';const x=new Date(d);return isNaN(x)?'':x.toISOString().slice(0,10);};
  if(old){const r=old.record,m={row:old.row,description:r['Description'],category:r['Category'],condition:r['Condition'],quantity:r['Quantity'],
  sku:r['SKU'],barcode:r['Barcode'],purchaseDate:iso(r['Purchase Date']),purchaseLocation:r['Purchase Location'],
  purchasePrice:r['Purchase Price'],taxPaid:r['Tax Paid'],acquisitionShipping:r['Acquisition Shipping'],
  storageLocation:r['Storage Location'],status:r['Status'],marketplace:r['Marketplace'],listedPrice:r['Listed Price'],
  expectedSalePrice:r['Expected Sale Price'],receiptLink:r['Receipt Link'],photoLink:r['Photo Link'],notes:r['Notes']};
  Object.keys(m).forEach(k=>{const e=document.querySelector('[name="'+k+'"]');if(e)e.value=m[k]??'';});}
  document.getElementById('f').addEventListener('submit',e=>{e.preventDefault();google.script.run
  .withSuccessHandler(()=>google.script.host.close()).withFailureHandler(x=>alert(x.message))
  .saveInventoryItem(Object.fromEntries(new FormData(e.target).entries()));});</script></body></html>`;
}
function findInventoryItem(){
  const ui=SpreadsheetApp.getUi(),res=ui.prompt('Find Inventory Item','Enter Item ID, SKU, barcode, or description:',ui.ButtonSet.OK_CANCEL);
  if(res.getSelectedButton()!==ui.Button.OK)return;const q=res.getResponseText().trim().toLowerCase();if(!q)return;
  const s=sheet3_(FTP3.SHEETS.INVENTORY);if(s.getLastRow()<2)return ui.alert('No inventory items found.');
  const headers=s.getRange(1,1,1,s.getLastColumn()).getDisplayValues()[0], rows=s.getRange(2,1,s.getLastRow()-1,s.getLastColumn()).getDisplayValues();
  const cols=['Item ID','Description','SKU','Barcode'].map(h=>headerIndex3_(headers,h));
  const i=rows.findIndex(r=>cols.some(c=>String(r[c]).toLowerCase().includes(q)));
  if(i<0)return ui.alert('No matching item found.');SpreadsheetApp.getActive().setActiveSheet(s);s.setActiveRange(s.getRange(i+2,1,1,FTP3.INVENTORY_HEADERS.length));
}
function showSlowInventory(){
  const s=sheet3_(FTP3.SHEETS.INVENTORY),map=headerMap3_(s);SpreadsheetApp.getActive().setActiveSheet(s);
  if(!s.getFilter())s.getRange(1,1,FTP3.ROWS+1,FTP3.INVENTORY_HEADERS.length).createFilter();
  s.getFilter().setColumnFilterCriteria(map['Days in Inventory'],SpreadsheetApp.newFilterCriteria().whenNumberGreaterThanOrEqualTo(90).build());
}

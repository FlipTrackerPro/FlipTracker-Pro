function showAddItemForm() {
  const template = HtmlService.createTemplate(itemFormHtml_());
  template.itemJson = 'null';
  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(520).setHeight(690),
    'Add Inventory Item'
  );
}

function showEditSelectedItemForm() {
  const s = SpreadsheetApp.getActiveSheet();

  if (s.getName() !== FTP2.SHEETS.INVENTORY || s.getActiveCell().getRow() < 2) {
    SpreadsheetApp.getUi().alert('Select an inventory row first.');
    return;
  }

  const row = s.getActiveCell().getRow();
  const item = s.getRange(row,1,1,FTP2.HEADERS.length).getValues()[0];
  const template = HtmlService.createTemplate(itemFormHtml_());
  template.itemJson = JSON.stringify({row:row,values:item});

  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(520).setHeight(690),
    'Edit Inventory Item'
  );
}

function itemFormHtml_() {
  return `<!doctype html><html><head><base target="_top"><style>
  body{font-family:Arial,sans-serif;padding:14px;color:#1F2937}
  label{display:block;font-weight:700;margin-top:9px}
  input,textarea{width:100%;box-sizing:border-box;padding:8px;margin-top:3px;border:1px solid #B7C9D6;border-radius:4px}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}
  .actions{margin-top:16px;display:flex;gap:8px}
  button{padding:10px 16px;border:0;border-radius:4px;background:#1F4E78;color:white;font-weight:700;cursor:pointer}
  .secondary{background:#6B7280}</style></head><body>
  <form id="f"><input type="hidden" name="row" id="row"><div class="grid">
  <div><label>Purchase date</label><input type="date" name="purchaseDate" required></div>
  <div><label>Title</label><input name="title" required></div>
  <div><label>SKU</label><input name="sku"></div>
  <div><label>Barcode</label><input name="barcode"></div>
  <div><label>Category</label><input name="category" required></div>
  <div><label>Purchase location</label><input name="purchaseLocation"></div>
  <div><label>Storage location</label><input name="storageLocation"></div>
  <div><label>Condition</label><input name="condition"></div>
  <div><label>Quantity</label><input type="number" min="1" name="quantity" value="1"></div>
  <div><label>Status</label><input name="status" value="Purchased"></div>
  <div><label>Purchase price</label><input type="number" step="0.01" min="0" name="purchasePrice"></div>
  <div><label>Tax paid</label><input type="number" step="0.01" min="0" name="taxPaid"></div>
  <div><label>Acquisition shipping</label><input type="number" step="0.01" min="0" name="acquisitionShipping"></div>
  <div><label>Expected sale price</label><input type="number" step="0.01" min="0" name="expectedSalePrice"></div>
  <div><label>Listed price</label><input type="number" step="0.01" min="0" name="listedPrice"></div>
  <div><label>Marketplace</label><input name="marketplace"></div>
  </div>
  <label>Receipt link</label><input type="url" name="receiptLink">
  <label>Photo link</label><input type="url" name="photoLink">
  <label>Notes</label><textarea name="notes" rows="3"></textarea>
  <div class="actions"><button type="submit">Save Item</button>
  <button type="button" class="secondary" onclick="google.script.host.close()">Cancel</button></div></form>
  <script>
  const existing = JSON.parse(<?!= JSON.stringify(itemJson) ?>);
  if(existing){
    document.getElementById('row').value=existing.row;
    const v=existing.values;
    const map={purchaseDate:v[1],title:v[2],sku:v[3],barcode:v[4],category:v[5],
      purchaseLocation:v[6],storageLocation:v[7],condition:v[8],quantity:v[9],
      purchasePrice:v[10],taxPaid:v[11],acquisitionShipping:v[12],
      expectedSalePrice:v[14],listedPrice:v[15],marketplace:v[16],status:v[18],
      receiptLink:v[22],photoLink:v[23],notes:v[24]};
    Object.keys(map).forEach(k=>{
      const e=document.querySelector('[name="'+k+'"]');
      if(e)e.value=map[k]??'';
    });
  }
  document.getElementById('f').addEventListener('submit',e=>{
    e.preventDefault();
    const d=Object.fromEntries(new FormData(e.target).entries());
    google.script.run.withSuccessHandler(()=>google.script.host.close())
      .withFailureHandler(err=>alert(err.message)).saveInventoryItem(d);
  });
  </script></body></html>`;
}

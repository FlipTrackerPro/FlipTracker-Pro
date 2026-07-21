function buildSalesSprint3_() {
  const s = sheet3_(FTP3.SHEETS.SALES);
  ensureSize3_(s,FTP3.ROWS+1,FTP3.SALES_HEADERS.length);
  s.getRange(1,1,1,FTP3.SALES_HEADERS.length).setValues([FTP3.SALES_HEADERS]);
  header3_(s.getRange(1,1,1,FTP3.SALES_HEADERS.length));
  s.setFrozenRows(1);
  setValidation3_(s,4,'FTP3_Marketplaces',FTP3.ROWS);
  s.getRange(2,3,FTP3.ROWS,1).setNumberFormat('yyyy-mm-dd');
  s.getRange(2,5,FTP3.ROWS,13).setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
  s.getRange(2,18,FTP3.ROWS,1).setNumberFormat('0.0%;[Red]-0.0%');
  s.getRange(2,23,FTP3.ROWS,1).setNumberFormat('yyyy-mm-dd hh:mm');
  if (s.getFilter()) s.getFilter().remove();
  s.getRange(1,1,FTP3.ROWS+1,FTP3.SALES_HEADERS.length).createFilter();
  borders3_(s.getRange(1,1,Math.min(FTP3.ROWS+1,200),FTP3.SALES_HEADERS.length));
}

function showRecordSaleForm() {
  const items = activeInventoryChoices3_();
  if (!items.length) {
    SpreadsheetApp.getUi().alert('There are no active inventory items available to sell.');
    return;
  }
  const html = HtmlService.createHtmlOutput(saleFormHtml3_(items)).setWidth(540).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html,'Record Sale');
}

function saleFormHtml3_(items) {
  const options = items.map(x => `<option value="${x.id}">${x.label}</option>`).join('');
  return `<!doctype html><html><head><base target="_top"><style>
  body{font-family:Arial;padding:14px;color:#1F2937}label{display:block;font-weight:700;margin-top:8px}
  input,select,textarea{width:100%;box-sizing:border-box;padding:8px;margin-top:3px;border:1px solid #B7C9D6;border-radius:4px}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}button{margin-top:16px;padding:10px 16px;border:0;border-radius:4px;background:#1F4E78;color:#fff;font-weight:700}
  </style></head><body><form id="f">
  <label>Inventory item</label><select name="itemId" required>${options}</select><div class="grid">
  <div><label>Sale date</label><input type="date" name="saleDate" required></div>
  <div><label>Marketplace</label><input name="marketplace" required></div>
  <div><label>Sale price</label><input type="number" step="0.01" min="0" name="salePrice" required></div>
  <div><label>Shipping charged</label><input type="number" step="0.01" min="0" name="shippingCharged"></div>
  <div><label>Shipping actual</label><input type="number" step="0.01" min="0" name="shippingActual"></div>
  <div><label>Packaging cost</label><input type="number" step="0.01" min="0" name="packagingCost"></div>
  <div><label>Marketplace fees</label><input type="number" step="0.01" min="0" name="marketplaceFees"></div>
  <div><label>Payment fees</label><input type="number" step="0.01" min="0" name="paymentFees"></div>
  <div><label>Promotion expense</label><input type="number" step="0.01" min="0" name="promotionExpense"></div>
  <div><label>GST/HST collected</label><input type="number" step="0.01" min="0" name="taxCollected"></div>
  <div><label>Buyer</label><input name="buyer"></div><div><label>Tracking number</label><input name="trackingNumber"></div>
  </div><label>Notes</label><textarea name="notes" rows="3"></textarea>
  <button type="submit">Record Sale</button></form><script>
  document.querySelector('[name="saleDate"]').value=new Date().toISOString().slice(0,10);
  document.getElementById('f').addEventListener('submit',e=>{e.preventDefault();
  google.script.run.withSuccessHandler(()=>google.script.host.close()).withFailureHandler(x=>alert(x.message))
  .saveSale3_(Object.fromEntries(new FormData(e.target).entries()));});</script></body></html>`;
}

function saveSale3_(form) {
  const item = inventoryItemById3_(form.itemId);
  if (!item) throw new Error('The selected inventory item was not found.');
  const inv = item.values;
  const salePrice = num3_(form.salePrice);
  const shippingCharged = num3_(form.shippingCharged);
  const shippingActual = num3_(form.shippingActual);
  const packaging = num3_(form.packagingCost);
  const marketFees = num3_(form.marketplaceFees);
  const paymentFees = num3_(form.paymentFees);
  const promotion = num3_(form.promotionExpense);
  const taxCollected = num3_(form.taxCollected);
  const itemCost = num3_(inv[13]);
  const grossRevenue = salePrice + shippingCharged;
  const sellingCosts = shippingActual + packaging + marketFees + paymentFees + promotion + taxCollected;
  const netProceeds = grossRevenue - sellingCosts;
  const realizedProfit = netProceeds - itemCost;
  const roi = itemCost ? realizedProfit / itemCost : '';
  const saleDate = date3_(form.saleDate);
  const purchaseDate = inv[1] instanceof Date ? inv[1] : date3_(inv[1]);
  const days = purchaseDate ? Math.max(0,Math.floor((saleDate-purchaseDate)/86400000)) : '';

  const values = [
    nextId3_(FTP3.SHEETS.SALES,1,'SAL'),form.itemId,saleDate,form.marketplace||'',
    salePrice,shippingCharged,shippingActual,packaging,marketFees,paymentFees,
    promotion,taxCollected,itemCost,grossRevenue,sellingCosts,netProceeds,
    realizedProfit,roi,days,form.buyer||'',form.trackingNumber||'',form.notes||'',new Date()
  ];
  const sales = sheet3_(FTP3.SHEETS.SALES);
  const row = Math.max(sales.getLastRow()+1,2);
  sales.getRange(row,1,1,values.length).setValues([values]);

  const inventory = sheet3_(FTP3.SHEETS.INVENTORY);
  inventory.getRange(item.row,19).setValue('Sold');
  inventory.getRange(item.row,27).setValue(new Date());
  refreshDashboardSprint3();
}

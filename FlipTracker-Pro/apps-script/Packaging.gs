function buildPackagingSprint3_() {
  const aliases={
    'Category':['Type'], 'Item Name':['Description'], 'Unit of Measure':[],
    'Units Purchased':[], 'Purchase Cost':[], 'Cost Per Unit':[],
    'Quantity On Hand':[], 'Reorder Level':[], 'Supplier':[],
    'SKU / Barcode':['Supplier SKU'], 'Product Link':[], 'Notes':[], 'Updated At':[]
  };
  migrateSheetHeaders3_(FTP3.SHEETS.PACKAGING,FTP3.PACKAGING_HEADERS,aliases);
  const s=sheet3_(FTP3.SHEETS.PACKAGING);
  ensureSize3_(s,300,FTP3.PACKAGING_HEADERS.length);
  s.getRange(1,1,1,FTP3.PACKAGING_HEADERS.length).setValues([FTP3.PACKAGING_HEADERS]);
  header3_(s.getRange(1,1,1,FTP3.PACKAGING_HEADERS.length));
  s.setFrozenRows(1);
  const map=headerMap3_(s);

  setValidation3_(s,map['Category'],'FTP3_PackagingTypes',299);
  const unitRule=SpreadsheetApp.newDataValidation().requireValueInList(['Each','Roll','Foot','Sheet','Metre','Bundle','Other'],true).setAllowInvalid(true).build();
  s.getRange(2,map['Unit of Measure'],299,1).setDataValidation(unitRule);
  const yesNo=SpreadsheetApp.newDataValidation().requireValueInList(['Yes','No'],true).setAllowInvalid(false).build();
  s.getRange(2,map['Active'],299,1).setDataValidation(yesNo);

  s.getRange(2,map['Units Purchased'],299,1).clearDataValidations().setDataValidation(wholeNumberRule3_(1)).setNumberFormat('0');
  ['Quantity On Hand','Reorder Level'].forEach(name=>{
    s.getRange(2,map[name],299,1).clearDataValidations().setDataValidation(wholeNumberRule3_(0)).setNumberFormat('0');
  });
  ['Purchase Cost','Cost Per Unit'].forEach(name=>s.getRange(2,map[name],299,1).setNumberFormat('$#,##0.00;[Red]-$#,##0.00'));
  s.getRange(2,map['Updated At'],299,1).setNumberFormat('yyyy-mm-dd hh:mm');

  repairPackagingCostPerUnit3_();

  if(s.getLastRow()>1){
    const active=s.getRange(2,map['Active'],s.getLastRow()-1,1);
    active.getValues().forEach((r,i)=>{if(!r[0])active.getCell(i+1,1).setValue('Yes');});
    const uom=s.getRange(2,map['Unit of Measure'],s.getLastRow()-1,1);
    uom.getValues().forEach((r,i)=>{if(!r[0])uom.getCell(i+1,1).setValue('Each');});
  }

  const data=s.getRange(2,1,299,FTP3.PACKAGING_HEADERS.length);
  const qtyLetter=columnLetter3_(map['Quantity On Hand']);
  const reorderLetter=columnLetter3_(map['Reorder Level']);
  const idLetter=columnLetter3_(map['Packaging ID']);
  const activeLetter=columnLetter3_(map['Active']);
  s.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($'+qtyLetter+'2<=$'+reorderLetter+'2,$'+idLetter+'2<>"",$'+activeLetter+'2="Yes")')
      .setBackground(FTP3.COLORS.LIGHT_RED).setRanges([data]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($'+qtyLetter+'2>0,$'+qtyLetter+'2<=$'+reorderLetter+'2*1.5,$'+idLetter+'2<>"",$'+activeLetter+'2="Yes")')
      .setBackground(FTP3.COLORS.GOLD).setRanges([data]).build()
  ]);

  if(s.getFilter())s.getFilter().remove();
  s.getRange(1,1,Math.max(2,s.getLastRow()),FTP3.PACKAGING_HEADERS.length).createFilter();
  s.autoResizeColumns(1,FTP3.PACKAGING_HEADERS.length);
  [map['Item Name'],map['Notes'],map['Product Link']].forEach(c=>s.setColumnWidth(c,c===map['Product Link']?240:180));
  if(SpreadsheetApp.getActive().getSheetByName(FTP3.SHEETS.SALES))refreshPackagingDropdowns3_();
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
  <div><label>Units purchased</label><input type="number" step="1" min="1" name="unitsPurchased" required></div>
  <div><label>Purchase cost</label><input type="number" step="0.01" min="0" name="purchaseCost" required></div>
  <div><label>Quantity on hand</label><input type="number" step="1" min="0" name="quantityOnHand"></div>
  <div><label>Reorder level</label><input type="number" step="1" min="0" name="reorderLevel"></div>
  <div><label>Supplier</label><input name="supplier"></div><div><label>SKU / Barcode</label><input name="supplierSku"></div>
  <div><label>Product link</label><input type="url" name="productLink"></div><div><label>Active</label><select name="active"><option>Yes</option><option>No</option></select></div>
  </div><label>Notes</label><textarea name="notes" rows="3"></textarea>
  <button type="submit">Add Supply</button></form><script>
  document.getElementById('f').addEventListener('submit',e=>{e.preventDefault();
  google.script.run.withSuccessHandler(()=>google.script.host.close()).withFailureHandler(x=>alert(x.message))
  .savePackaging3(Object.fromEntries(new FormData(e.target).entries()));});</script></body></html>`).setWidth(560).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html,'Add Packaging Supply');
}

function savePackaging3(form) { return savePackaging3_(form); }

function savePackaging3_(form) {
  const units=num3_(form.unitsPurchased);
  const purchaseCost=num3_(form.purchaseCost);
  const quantity=form.quantityOnHand===''||form.quantityOnHand==null?units:num3_(form.quantityOnHand);
  const reorder=num3_(form.reorderLevel);
  if(units<=0||!Number.isInteger(units))throw new Error('Units purchased must be a whole number greater than zero.');
  if(!Number.isInteger(quantity)||quantity<0)throw new Error('Quantity on hand must be a whole number.');
  if(!Number.isInteger(reorder)||reorder<0)throw new Error('Reorder level must be a whole number.');
  if(purchaseCost<0)throw new Error('Purchase cost cannot be negative.');

  const s=sheet3_(FTP3.SHEETS.PACKAGING);
  const map=headerMap3_(s);
  const row=Math.max(s.getLastRow()+1,2);
  const record={
    'Packaging ID':nextId3_(FTP3.SHEETS.PACKAGING,map['Packaging ID'],'PKG'),
    'Category':form.category||'','Item Name':form.itemName||'','Size':form.size||'',
    'Unit of Measure':form.unitOfMeasure||'Each','Units Purchased':units,
    'Purchase Cost':purchaseCost,'Cost Per Unit':'','Quantity On Hand':quantity,
    'Reorder Level':reorder,'Supplier':form.supplier||'','SKU / Barcode':form.supplierSku||'',
    'Product Link':form.productLink||'','Active':form.active||'Yes','Notes':form.notes||'',
    'Updated At':new Date()
  };
  s.getRange(row,1,1,FTP3.PACKAGING_HEADERS.length).setValues([FTP3.PACKAGING_HEADERS.map(h=>record[h])]);
  setPackagingCostFormula3_(s,row,map);
  SpreadsheetApp.flush();
  refreshPackagingDropdowns3_();
  refreshDashboardSprint3();
}


function setPackagingCostFormula3_(sheet,row,map) {
  map=map||headerMap3_(sheet);
  const unitsLetter=columnLetter3_(map['Units Purchased']);
  const purchaseLetter=columnLetter3_(map['Purchase Cost']);
  sheet.getRange(row,map['Cost Per Unit'])
    .setFormula('=IF(OR('+unitsLetter+row+'="",'+unitsLetter+row+'=0,'+purchaseLetter+row+'=""),"",ROUND('+purchaseLetter+row+'/'+unitsLetter+row+',2))')
    .setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
}

function repairPackagingCostPerUnit3_() {
  const s=SpreadsheetApp.getActive().getSheetByName(FTP3.SHEETS.PACKAGING);
  if(!s)return;
  const map=headerMap3_(s);
  const rows=Math.max(1,s.getMaxRows()-1);
  const unitsLetter=columnLetter3_(map['Units Purchased']);
  const purchaseLetter=columnLetter3_(map['Purchase Cost']);
  const formulas=[];
  for(let row=2;row<=rows+1;row++){
    formulas.push(['=IF(OR('+unitsLetter+row+'="",'+unitsLetter+row+'=0,'+purchaseLetter+row+'=""),"",ROUND('+purchaseLetter+row+'/'+unitsLetter+row+',2))']);
  }
  s.getRange(2,map['Cost Per Unit'],rows,1)
    .setFormulas(formulas)
    .setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
}

function repairPackagingCostPerUnit() {
  repairPackagingCostPerUnit3_();
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Packaging Cost Repair','Cost Per Unit now calculates automatically as Purchase Cost ÷ Units Purchased.',SpreadsheetApp.getUi().ButtonSet.OK);
}

function packagingChoices3_() {
  const s=sheet3_(FTP3.SHEETS.PACKAGING); if(s.getLastRow()<2)return [];
  return s.getRange(2,1,s.getLastRow()-1,FTP3.PACKAGING_HEADERS.length).getValues().filter(r=>r[0]&&String(r[13]||'Yes')!=='No').map(r=>({
    id:String(r[0]),category:String(r[1]||''),label:String(r[0])+' — '+String(r[2]||r[1]||'Packaging')+(r[3]?' ('+r[3]+')':''),
    unit:String(r[4]||'Each'),cost:num3_(r[7]),available:num3_(r[8])
  }));
}

function packagingItemById3_(id) {
  const normalized=packagingIdFromSelection3_(id);
  if(!normalized)return null;
  const s=sheet3_(FTP3.SHEETS.PACKAGING);
  if(s.getLastRow()<2)return null;
  const map=headerMap3_(s);
  const idColumn=map['Packaging ID'];
  const ids=s.getRange(2,idColumn,s.getLastRow()-1,1).getDisplayValues().flat()
    .map(v=>packagingIdFromSelection3_(v));
  const index=ids.findIndex(v=>v===normalized);
  if(index<0)return null;
  const row=index+2;
  return {row:row,values:s.getRange(row,1,1,FTP3.PACKAGING_HEADERS.length).getValues()[0]};
}

function calculatePackagingUsage3_(form) {
  const specs=[['boxId','boxQty','Box'],['bubbleId','bubbleQty','Bubble Wrap'],['mailerId','mailerQty','Mailer'],['tapeId','tapeQty','Tape'],['otherPackagingId','otherPackagingQty','Other']];
  const usage=[];let total=0;
  specs.forEach(([idKey,qtyKey,expectedCategory])=>{
    const id=packagingIdFromSelection3_(String(form[idKey]||'').trim());
    const qty=num3_(form[qtyKey]);
    if(!id&&qty>0)throw new Error(expectedCategory+' quantity was entered without selecting a Packaging ID.');
    if(!id)return;
    if(qty<=0)throw new Error(expectedCategory+' quantity must be greater than zero.');
    const item=packagingItemById3_(id);
    if(!item)throw new Error(expectedCategory+' Packaging ID was not found: '+id);
    const record=rowRecord3_(FTP3.PACKAGING_HEADERS,item.values);
    const actualCategory=packagingCategoryKey3_(record['Category']);
    if(expectedCategory!=='Other'&&actualCategory!==expectedCategory)throw new Error(id+' is categorized as '+actualCategory+', not '+expectedCategory+'.');
    if(expectedCategory==='Other'&&actualCategory!=='Other')throw new Error(id+' belongs in the '+actualCategory+' field, not Other Packaging.');
    const available=num3_(record['Quantity On Hand']);
    if(qty>available+0.000001)throw new Error(expectedCategory+' quantity exceeds stock. Available: '+available+' '+String(record['Unit of Measure']||'units'));
    const cost=num3_(record['Cost Per Unit'])*qty;
    usage.push({id:id,qty:qty,row:item.row,cost:cost});total+=cost;
  });
  return{usage:usage,total:Math.round(total*100)/100};
}

function deductPackagingUsage3_(usage) {
  const s=sheet3_(FTP3.SHEETS.PACKAGING),map=headerMap3_(s);usage.forEach(u=>{const current=num3_(s.getRange(u.row,map['Quantity On Hand']).getValue());if(u.qty>current)throw new Error('Packaging stock changed before the sale was saved. Reopen the sale form.');s.getRange(u.row,map['Quantity On Hand']).setValue(Math.round(current-u.qty)).setNumberFormat('0');s.getRange(u.row,map['Updated At']).setValue(new Date());});
}

function goToPackagingInventory3_(){SpreadsheetApp.getActive().setActiveSheet(sheet3_(FTP3.SHEETS.PACKAGING));}


function packagingIdFromSelection3_(selection) {
  const text=String(selection||'').trim();
  if(!text)return '';
  const match=text.match(/\bPKG[\s_-]*0*(\d+)\b/i);
  if(match)return 'PKG-'+String(Number(match[1])).padStart(5,'0');
  return text.toUpperCase();
}

function packagingListLabel3_(row) {
  return String(row[0])+' — '+String(row[2]||row[1]||'Packaging')+(row[3]?' ('+row[3]+')':'');
}

function packagingCategoryKey3_(category) {
  const c=String(category||'').toLowerCase();
  if(c.indexOf('box')>=0)return 'Box';
  if(c.indexOf('bubble')>=0)return 'Bubble Wrap';
  if(c.indexOf('mailer')>=0||c.indexOf('envelope')>=0)return 'Mailer';
  if(c.indexOf('tape')>=0)return 'Tape';
  return 'Other';
}

function refreshPackagingDropdowns3_() {
  const ss=SpreadsheetApp.getActive();
  const packaging=ss.getSheetByName(FTP3.SHEETS.PACKAGING);
  const sales=ss.getSheetByName(FTP3.SHEETS.SALES);
  if(!packaging||!sales)return;
  let lists=ss.getSheetByName(FTP3.SHEETS.PACKAGING_LISTS);
  if(!lists)lists=ss.insertSheet(FTP3.SHEETS.PACKAGING_LISTS);
  lists.clear();ensureSize3_(lists,Math.max(300,lists.getMaxRows()),5);
  const headers=['Boxes','Bubble Wrap','Mailers','Tape','Other Packaging'];
  lists.getRange(1,1,1,5).setValues([headers]);header3_(lists.getRange(1,1,1,5));
  const buckets={'Box':[],'Bubble Wrap':[],'Mailer':[],'Tape':[],'Other':[]};
  if(packaging.getLastRow()>1){
    packaging.getRange(2,1,packaging.getLastRow()-1,FTP3.PACKAGING_HEADERS.length).getValues().forEach(r=>{
      if(!r[0]||String(r[13]||'Yes').toLowerCase()==='no')return;
      buckets[packagingCategoryKey3_(r[1])].push(packagingIdFromSelection3_(r[0]));
    });
  }
  const ordered=[buckets.Box,buckets['Bubble Wrap'],buckets.Mailer,buckets.Tape,buckets.Other];
  ordered.forEach((values,i)=>{values.sort();if(values.length)lists.getRange(2,i+1,values.length,1).setValues(values.map(v=>[v]));});
  const salesMap=headerMap3_(sales);
  const salesColumns=[
    salesMap['Box Used'],salesMap['Bubble Wrap Used'],salesMap['Mailer Used'],
    salesMap['Tape Used'],salesMap['Other Packaging Used']
  ];
  salesColumns.forEach((col,i)=>{
    const count=ordered[i].length;
    const target=sales.getRange(2,col,FTP3.ROWS,1);
    target.clearDataValidations();
    if(count){
      const source=lists.getRange(2,i+1,count,1);
      const rule=SpreadsheetApp.newDataValidation().requireValueInRange(source,true).setAllowInvalid(false).setHelpText('Choose an active packaging item from the Packaging sheet.').build();
      target.setDataValidation(rule);
    }
  });
  lists.hideSheet();
  ss.toast('Packaging dropdowns refreshed.','FlipTracker Pro',4);
}

function onEdit(e) {
  if(!e||!e.range)return;
  if(handleExpenseEdit3_(e))return;
  const sheet=e.range.getSheet();
  if(sheet.getName()===FTP3.SHEETS.PACKAGING && e.range.getRow()>1){
    const map=headerMap3_(sheet),row=e.range.getRow(),col=e.range.getColumn();
    const wholeHeaders=['Units Purchased','Quantity On Hand','Reorder Level'];
    if(wholeHeaders.map(h=>map[h]).indexOf(col)>=0){
      const value=e.range.getValue();
      if(value!==''&&value!==null){
        const n=Number(value);
        if(!Number.isFinite(n)||n<0||Math.floor(n)!==n){
          e.range.setValue(e.oldValue||'');
          SpreadsheetApp.getActive().toast('Enter a whole number with no decimal places.','FlipTracker Pro',5);
          return;
        }
      }
    }
    if(col===map['Units Purchased']||col===map['Purchase Cost'])setPackagingCostFormula3_(sheet,row,map);
    if([map['Packaging ID'],map['Category'],map['Item Name'],map['Size'],map['Active']].indexOf(col)>=0)refreshPackagingDropdowns3_();
    return;
  }
  if(sheet.getName()!==FTP3.SHEETS.INVENTORY||e.range.getRow()<2)return;
  const map=headerMap3_(sheet), row=e.range.getRow();
  const recalculationHeaders=['Description','Quantity','Purchase Date','Purchase Price','Tax Paid','Acquisition Shipping','Listed Price','Expected Sale Price'];
  const recalculationColumns=recalculationHeaders.map(h=>map[h]);
  if(recalculationColumns.indexOf(e.range.getColumn())>=0)recalculateInventoryRow3_(sheet,row);
  if(e.range.getColumn()===map['Status']){
    const value=String(e.value||'');
    if(value==='Listed'&&!sheet.getRange(row,map['Listing Date']).getValue()){
      sheet.getRange(row,map['Listing Date']).setValue(new Date()).setNumberFormat('yyyy-mm-dd');
    }
    if(value==='Sold'){
      const itemId=sheet.getRange(row,map['Item ID']).getDisplayValue();
      if(!itemId)return;
      if(saleExistsForInventory3_(itemId)){
        SpreadsheetApp.getActive().toast('A Sales record already exists for '+itemId+'.','FlipTracker Pro',6);return;
      }
      const previous=e.oldValue||'Listed';
      PropertiesService.getDocumentProperties().setProperty('FTP_PREV_STATUS_'+itemId,previous);
      e.range.setValue('Sale Pending');
      e.range.setNote('Sale details are incomplete. Keep this row selected, then use FlipTracker Pro → Complete Selected Pending Sale.');
      sheet.setActiveRange(sheet.getRange(row,1,1,FTP3.INVENTORY_HEADERS.length));
      SpreadsheetApp.getActive().toast('Sale Pending created for '+itemId+'. Use FlipTracker Pro → Complete Selected Pending Sale.','Complete the sale',10);
    }
  }
}

function repairSalesPackagingIds3_() {
  const sales=sheet3_(FTP3.SHEETS.SALES);
  const map=headerMap3_(sales);
  const fields=['Box Used','Bubble Wrap Used','Mailer Used','Tape Used','Other Packaging Used'];
  if(sales.getLastRow()<2){
    refreshPackagingDropdowns3_();
    return {repaired:0,missing:[]};
  }

  let repaired=0;
  const missing=[];
  fields.forEach(name=>{
    const col=map[name];
    const range=sales.getRange(2,col,sales.getLastRow()-1,1);
    const values=range.getDisplayValues();
    let changed=false;
    values.forEach((row,i)=>{
      const original=String(row[0]||'').trim();
      if(!original)return;
      const normalized=packagingIdFromSelection3_(original);
      const item=packagingItemById3_(normalized);
      if(item){
        if(original!==normalized){
          values[i][0]=normalized;
          repaired++;
          changed=true;
        }
      }else{
        missing.push({row:i+2,column:name,value:original});
      }
    });
    if(changed)range.setValues(values);
  });

  refreshPackagingDropdowns3_();
  SpreadsheetApp.flush();
  return {repaired:repaired,missing:missing};
}

function repairSalesPackagingIds() {
  const result=repairSalesPackagingIds3_();
  const message=result.missing.length
    ? result.repaired+' packaging reference(s) normalized. '+result.missing.length+' reference(s) still do not match a Packaging ID; see Calculation Audit.'
    : result.repaired+' packaging reference(s) normalized. All referenced Packaging IDs now exist.';
  SpreadsheetApp.getUi().alert('Packaging ID Repair',message,SpreadsheetApp.getUi().ButtonSet.OK);
  return result;
}

function repairPackagingCostAudit3_() {
  repairPackagingCostPerUnit3_();
  SpreadsheetApp.flush();
}

function repairPackagingCostAudit() {
  repairPackagingCostAudit3_();
  SpreadsheetApp.getUi().alert(
    'Packaging Cost Audit Repair',
    'Cost Per Unit has been recalculated to two decimal places. The Calculation Audit now compares against the same rounded currency value.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

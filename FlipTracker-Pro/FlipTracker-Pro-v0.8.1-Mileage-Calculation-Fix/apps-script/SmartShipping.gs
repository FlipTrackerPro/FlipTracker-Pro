/** FlipTracker Pro v0.8.0 Smart Shipping */
function buildShippingSettings80_() {
  const s=sheet3_(FTP3.SHEETS.SHIPPING_SETTINGS);
  const headers=['Carrier','Service','Base Charge','Per Kilogram','Active','Notes'];
  ensureSize3_(s,100,headers.length);
  s.getRange(1,1,1,headers.length).setValues([headers]);header3_(s.getRange(1,1,1,headers.length));s.setFrozenRows(1);
  if(s.getLastRow()<2){
    s.getRange(2,1,6,headers.length).setValues([
      ['Canada Post','Expedited Parcel',12.00,1.50,'Yes','Editable planning estimate only'],
      ['Canada Post','Regular Parcel',11.00,1.35,'Yes','Editable planning estimate only'],
      ['UPS','Standard',14.00,1.75,'Yes','Editable planning estimate only'],
      ['FedEx','Ground',15.00,1.85,'Yes','Editable planning estimate only'],
      ['Purolator','Ground',15.00,1.80,'Yes','Editable planning estimate only'],
      ['Other','Custom',0,0,'Yes','Enter your own rates']
    ]);
  }
  s.getRange(2,3,99,2).setNumberFormat('$#,##0.00');
  s.getRange(2,5,99,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes','No'],true).setAllowInvalid(false).build());
  if(s.getFilter())s.getFilter().remove();s.getRange(1,1,Math.max(2,s.getLastRow()),headers.length).createFilter();
  s.autoResizeColumns(1,headers.length);s.setColumnWidth(6,240);
}

function shippingRates80_(){
  const s=sheet3_(FTP3.SHEETS.SHIPPING_SETTINGS);if(s.getLastRow()<2)return[];
  return s.getRange(2,1,s.getLastRow()-1,6).getValues().filter(r=>r[0]&&String(r[4]||'Yes').toLowerCase()!=='no').map(r=>({carrier:String(r[0]),service:String(r[1]),base:num3_(r[2]),perKg:num3_(r[3])}));
}

function parsePackageSize80_(text){
  const n=String(text||'').match(/(\d+(?:\.\d+)?)\s*[x×]\s*(\d+(?:\.\d+)?)\s*[x×]\s*(\d+(?:\.\d+)?)/i);
  return n?{l:Number(n[1]),w:Number(n[2]),h:Number(n[3])}:null;
}

function recommendShipping80_(input){
  const dims=[num3_(input.length),num3_(input.width),num3_(input.height)].sort((a,b)=>b-a),weight=Math.max(0,num3_(input.weight));
  if(dims.some(x=>x<=0))throw new Error('Enter length, width, and height greater than zero.');
  const choices=packagingChoices3_().filter(x=>packagingCategoryKey3_(x.category)==='Box'&&x.available>0).map(x=>{
    const d=parsePackageSize80_(x.label);if(!d)return null;const box=[d.l,d.w,d.h].sort((a,b)=>b-a);
    const fits=box.every((v,i)=>v>=dims[i]);return fits?Object.assign({},x,{volume:box[0]*box[1]*box[2],dims:box}):null;
  }).filter(Boolean).sort((a,b)=>a.volume-b.volume||a.cost-b.cost);
  const box=choices[0]||null;
  const rates=shippingRates80_().map(r=>Object.assign({},r,{estimate:roundMoney3_(r.base+r.perKg*weight)})).sort((a,b)=>a.estimate-b.estimate);
  const rate=input.carrier?rates.find(r=>r.carrier===input.carrier&&(!input.service||r.service===input.service)):rates[0];
  const material=box?roundMoney3_(box.cost):0;
  return {box:box?{id:box.id,label:box.label,cost:material,available:box.available}:null,rate:rate||null,estimatedShipping:rate?rate.estimate:0,packagingCost:material,totalFulfilment:roundMoney3_((rate?rate.estimate:0)+material)};
}

function showSmartShippingAssistant(){
  buildShippingSettings80_();
  const sales=sheet3_(FTP3.SHEETS.SALES),map=headerMap3_(sales),active=SpreadsheetApp.getActiveRange();
  let saleId='';if(active&&active.getSheet().getName()===FTP3.SHEETS.SALES&&active.getRow()>1)saleId=sales.getRange(active.getRow(),map['Sale ID']).getDisplayValue();
  const rates=shippingRates80_();
  const opts=rates.map(r=>'<option value="'+escapeHtml80_(r.carrier+'|'+r.service)+'">'+escapeHtml80_(r.carrier+' — '+r.service)+'</option>').join('');
  const html=HtmlService.createHtmlOutput(`<!doctype html><html><head><base target="_top"><style>body{font-family:Arial;padding:14px;color:#1F2937}label{display:block;font-weight:700;margin-top:8px}input,select{width:100%;box-sizing:border-box;padding:8px;margin-top:3px;border:1px solid #B7C9D6;border-radius:4px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}.result{margin-top:14px;padding:12px;background:#F3F6F9;border-radius:6px;white-space:pre-line}.actions{display:flex;gap:8px;margin-top:14px}button{padding:10px 14px;border:0;border-radius:4px;background:#1F4E78;color:white;font-weight:700}</style></head><body>
  <label>Sale ID</label><input id="saleId" value="${escapeHtml80_(saleId)}" placeholder="Select a Sales row first or enter Sale ID">
  <div class="grid"><div><label>Length (cm)</label><input id="length" type="number" min="0" step="0.1"></div><div><label>Width (cm)</label><input id="width" type="number" min="0" step="0.1"></div><div><label>Height (cm)</label><input id="height" type="number" min="0" step="0.1"></div><div><label>Weight (kg)</label><input id="weight" type="number" min="0" step="0.01"></div></div>
  <label>Preferred carrier/service</label><select id="rate"><option value="">Lowest configured estimate</option>${opts}</select>
  <div class="actions"><button onclick="calculate()">Recommend</button><button onclick="save()">Save to Sale</button></div><div id="result" class="result">Enter package details, then choose Recommend.</div>
  <script>let recommendation=null;function data(){const r=document.getElementById('rate').value.split('|');return{length:length.value,width:width.value,height:height.value,weight:weight.value,carrier:r[0]||'',service:r[1]||''};}function render(x){recommendation=x;result.textContent=(x.box?'Recommended box: '+x.box.id+' — '+x.box.label+'\nPackaging cost: $'+x.packagingCost.toFixed(2):'No fitting box with readable dimensions was found.')+(x.rate?'\nShipping estimate: $'+x.estimatedShipping.toFixed(2)+' ('+x.rate.carrier+' — '+x.rate.service+')':'\nNo active shipping rate found.')+'\nEstimated fulfilment total: $'+x.totalFulfilment.toFixed(2);}function calculate(){google.script.run.withSuccessHandler(render).withFailureHandler(e=>alert(e.message)).recommendShipping80_(data());}function save(){if(!recommendation){calculate();return;}google.script.run.withSuccessHandler(m=>{alert(m);google.script.host.close();}).withFailureHandler(e=>alert(e.message)).saveSmartShippingToSale80_(saleId.value,data(),recommendation);}</script></body></html>`).setWidth(580).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html,'Smart Shipping Assistant');
}

function escapeHtml80_(v){return String(v==null?'':v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');}

function saveSmartShippingToSale80_(saleId,input,result){
  saleId=String(saleId||'').trim();if(!saleId)throw new Error('Enter a Sale ID or select a Sales row.');
  const s=sheet3_(FTP3.SHEETS.SALES),map=headerMap3_(s);if(s.getLastRow()<2)throw new Error('No sales records exist.');
  const ids=s.getRange(2,map['Sale ID'],s.getLastRow()-1,1).getDisplayValues().flat();const idx=ids.indexOf(saleId);if(idx<0)throw new Error('Sale ID not found: '+saleId);const row=idx+2;
  const rate=result.rate||{};const values={
    'Package Length (cm)':num3_(input.length),'Package Width (cm)':num3_(input.width),'Package Height (cm)':num3_(input.height),'Package Weight (kg)':num3_(input.weight),
    'Shipping Carrier':rate.carrier||input.carrier||'','Shipping Service':rate.service||input.service||'','Suggested Packaging':result.box?result.box.id:'','Estimated Shipping':num3_(result.estimatedShipping)
  };
  Object.keys(values).forEach(h=>s.getRange(row,map[h]).setValue(values[h]));
  const actual=num3_(s.getRange(row,map['Shipping Actual']).getValue()),charged=num3_(s.getRange(row,map['Shipping Charged']).getValue());
  s.getRange(row,map['Shipping Variance']).setValue(roundMoney3_(charged-(actual||num3_(result.estimatedShipping))));
  if(result.box && !s.getRange(row,map['Box Used']).getValue()){s.getRange(row,map['Box Used']).setValue(result.box.id);s.getRange(row,map['Box Qty']).setValue(1);}
  SpreadsheetApp.flush();refreshDashboardSprint3();return 'Smart shipping details saved to '+saleId+'.';
}

function refreshShippingVariance80_(){
  const s=sheet3_(FTP3.SHEETS.SALES),m=headerMap3_(s);if(s.getLastRow()<2)return;
  for(let r=2;r<=s.getLastRow();r++){if(!s.getRange(r,m['Sale ID']).getValue())continue;const charged=num3_(s.getRange(r,m['Shipping Charged']).getValue()),actual=num3_(s.getRange(r,m['Shipping Actual']).getValue()),estimate=num3_(s.getRange(r,m['Estimated Shipping']).getValue());s.getRange(r,m['Shipping Variance']).setValue(roundMoney3_(charged-(actual||estimate)));}
}

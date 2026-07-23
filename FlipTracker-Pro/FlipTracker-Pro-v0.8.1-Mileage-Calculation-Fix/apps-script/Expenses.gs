function buildExpensesSprint3_() {
  const s = sheet3_(FTP3.SHEETS.EXPENSES);
  ensureSize3_(s,FTP3.ROWS+1,FTP3.EXPENSE_HEADERS.length);
  s.getRange(1,1,1,FTP3.EXPENSE_HEADERS.length).setValues([FTP3.EXPENSE_HEADERS]);
  header3_(s.getRange(1,1,1,FTP3.EXPENSE_HEADERS.length));
  s.setFrozenRows(1);
  setValidation3_(s,3,'FTP3_ExpenseCategories',FTP3.ROWS);
  setValidation3_(s,11,'FTP3_PaymentMethods',FTP3.ROWS);
  s.getRange(2,2,FTP3.ROWS,1).setNumberFormat('yyyy-mm-dd');
  s.getRange(2,6,FTP3.ROWS,4).setNumberFormat('$#,##0.00');
  s.getRange(2,9,FTP3.ROWS,1).setNumberFormat('0%');
  s.getRange(2,10,FTP3.ROWS,1).setNumberFormat('$#,##0.00');
  s.getRange(2,14,FTP3.ROWS,1).setNumberFormat('yyyy-mm-dd hh:mm');
  if (s.getFilter()) s.getFilter().remove();
  s.getRange(1,1,FTP3.ROWS+1,FTP3.EXPENSE_HEADERS.length).createFilter();
}

function showRecordExpenseForm() {
  const html = HtmlService.createHtmlOutput(`<!doctype html><html><head><base target="_top"><style>
  body{font-family:Arial;padding:14px;color:#1F2937}label{display:block;font-weight:700;margin-top:8px}
  input,textarea{width:100%;box-sizing:border-box;padding:8px;margin-top:3px;border:1px solid #B7C9D6;border-radius:4px}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}button{margin-top:16px;padding:10px 16px;border:0;border-radius:4px;background:#1F4E78;color:#fff;font-weight:700}
  </style></head><body><form id="f"><div class="grid">
  <div><label>Date</label><input type="date" name="expenseDate" required></div><div><label>Category</label><input name="category" required></div>
  <div><label>Vendor</label><input name="vendor"></div><div><label>Payment method</label><input name="paymentMethod"></div>
  <div><label>Subtotal</label><input type="number" step="0.01" min="0" name="subtotal"></div>
  <div><label>GST/HST paid</label><input type="number" step="0.01" min="0" name="taxPaid"></div>
  <div><label>Business use %</label><input type="number" step="1" min="0" max="100" name="businessUse" value="100"></div>
  <div><label>Receipt link</label><input type="url" name="receiptLink"></div></div>
  <label>Description</label><input name="description" required><label>Notes</label><textarea name="notes" rows="3"></textarea>
  <button type="submit">Record Expense</button></form><script>
  document.querySelector('[name="expenseDate"]').value=new Date().toISOString().slice(0,10);
  document.getElementById('f').addEventListener('submit',e=>{e.preventDefault();
  google.script.run.withSuccessHandler(()=>google.script.host.close()).withFailureHandler(x=>alert(x.message))
  .saveExpense3(Object.fromEntries(new FormData(e.target).entries()));});</script></body></html>`)
  .setWidth(520).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html,'Record Expense');
}

function saveExpense3(form) { return saveExpense3_(form); }

function saveExpense3_(form) {
  const subtotal=num3_(form.subtotal);
  const tax=num3_(form.taxPaid);
  const businessUse=Math.round(Math.min(100,Math.max(0,num3_(form.businessUse))));
  const percent=businessUse/100;

  if(subtotal<0||tax<0)throw new Error('Subtotal and GST/HST Paid cannot be negative.');

  const s=sheet3_(FTP3.SHEETS.EXPENSES);
  const map=headerMap3_(s);
  const row=Math.max(s.getLastRow()+1,2);
  const record={
    'Expense ID':nextId3_(FTP3.SHEETS.EXPENSES,map['Expense ID'],'EXP'),
    'Expense Date':form.expenseDate?new Date(form.expenseDate):new Date(),
    'Vendor':form.vendor||'','Category':form.category||'',
    'Description':form.description||'','Subtotal':subtotal,
    'GST/HST Paid':tax,'Total':'','Payment Method':form.paymentMethod||'',
    'Receipt Link':form.receiptLink||'','Business Use %':percent,
    'Deductible Amount':'','Notes':form.notes||'','Created At':new Date()
  };
  s.getRange(row,1,1,FTP3.EXPENSE_HEADERS.length).setValues([FTP3.EXPENSE_HEADERS.map(h=>record[h])]);
  setExpenseCalculationFormulas3_(s,row,map);
  SpreadsheetApp.flush();
  refreshDashboardSprint3();
}


function setExpenseCalculationFormulas3_(sheet,row,map) {
  map=map||headerMap3_(sheet);
  const subtotal=columnLetter3_(map['Subtotal']);
  const tax=columnLetter3_(map['GST/HST Paid']);
  const total=columnLetter3_(map['Total']);
  const business=columnLetter3_(map['Business Use %']);

  sheet.getRange(row,map['Total'])
    .setFormula('=IF('+subtotal+row+'="","",ROUND('+subtotal+row+'+IF('+tax+row+'="",0,'+tax+row+'),2))')
    .setNumberFormat('$#,##0.00;[Red]-$#,##0.00');

  sheet.getRange(row,map['Deductible Amount'])
    .setFormula('=IF('+total+row+'="","",ROUND('+total+row+'*IF('+business+row+'="",1,IF('+business+row+'>1,'+business+row+'/100,'+business+row+')),2))')
    .setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
}

function repairExpenseCalculations3_() {
  normalizeExpenseBusinessUse3_();
  const s=SpreadsheetApp.getActive().getSheetByName(FTP3.SHEETS.EXPENSES);
  if(!s)return;
  const map=headerMap3_(s);
  const rows=Math.max(1,s.getMaxRows()-1);

  const subtotal=columnLetter3_(map['Subtotal']);
  const tax=columnLetter3_(map['GST/HST Paid']);
  const total=columnLetter3_(map['Total']);
  const business=columnLetter3_(map['Business Use %']);

  const totalFormulas=[];
  const deductibleFormulas=[];
  for(let row=2;row<=rows+1;row++){
    totalFormulas.push(['=IF('+subtotal+row+'="","",ROUND('+subtotal+row+'+IF('+tax+row+'="",0,'+tax+row+'),2))']);
    deductibleFormulas.push(['=IF('+total+row+'="","",ROUND('+total+row+'*IF('+business+row+'="",1,IF('+business+row+'>1,'+business+row+'/100,'+business+row+')),2))']);
  }

  s.getRange(2,map['Total'],rows,1)
    .setFormulas(totalFormulas)
    .setNumberFormat('$#,##0.00;[Red]-$#,##0.00');

  s.getRange(2,map['Deductible Amount'],rows,1)
    .setFormulas(deductibleFormulas)
    .setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
}

function repairExpenseCalculations() {
  repairExpenseCalculations3_();
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    'Expense Calculations Repaired',
    'Total now calculates as Subtotal + GST/HST Paid. Deductible Amount also recalculates automatically.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


function handleExpenseEdit3_(e) {
  const sheet=e.range.getSheet();
  if(sheet.getName()!==FTP3.SHEETS.EXPENSES||e.range.getRow()<2)return false;
  const map=headerMap3_(sheet);
  const col=e.range.getColumn();
  if([map['Subtotal'],map['GST/HST Paid'],map['Business Use %']].indexOf(col)>=0){
    if(col===map['Business Use %']){
      const value=e.range.getValue();
      if(value!==''&&value!==null){
        const n=Number(value);
        if(!Number.isFinite(n)||n<0||n>100){
          e.range.setValue(e.oldValue||'');
          SpreadsheetApp.getActive().toast('Business Use must be from 0 to 100.','FlipTracker Pro',5);
          return true;
        }
        e.range.setValue(n>1?n/100:n).setNumberFormat('0%');
      }
    }
    setExpenseCalculationFormulas3_(sheet,e.range.getRow(),map);
    return true;
  }
  return false;
}


function normalizeExpenseBusinessUse3_() {
  const s=SpreadsheetApp.getActive().getSheetByName(FTP3.SHEETS.EXPENSES);
  if(!s||s.getLastRow()<2)return 0;
  const map=headerMap3_(s);
  const range=s.getRange(2,map['Business Use %'],s.getLastRow()-1,1);
  const values=range.getValues();
  let changed=0;
  values.forEach(row=>{
    if(row[0]===''||row[0]===null)return;
    const n=Number(row[0]);
    if(!Number.isFinite(n))return;
    const normalized=Math.max(0,Math.min(1,n>1?n/100:n));
    if(Math.abs(n-normalized)>0.0000001){row[0]=normalized;changed++;}
  });
  if(changed)range.setValues(values);
  range.setNumberFormat('0%');
  return changed;
}

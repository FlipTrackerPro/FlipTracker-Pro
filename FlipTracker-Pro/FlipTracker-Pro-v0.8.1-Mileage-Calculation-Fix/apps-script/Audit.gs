function runFlipTrackerCalculationAudit() {
  const issues=[],add=(severity,area,row,message)=>issues.push([severity,area,row||'',message]);
  const inv=sheet3_(FTP3.SHEETS.INVENTORY),sales=sheet3_(FTP3.SHEETS.SALES),exp=sheet3_(FTP3.SHEETS.EXPENSES),mil=sheet3_(FTP3.SHEETS.MILEAGE),pkg=sheet3_(FTP3.SHEETS.PACKAGING);
  const rows=(s,headers)=>s.getLastRow()>1?s.getRange(2,1,s.getLastRow()-1,headers.length).getValues():[];

  rows(inv,FTP3.INVENTORY_HEADERS).forEach((r,i)=>{
    const x=rowRecord3_(FTP3.INVENTORY_HEADERS,r);if(!x['Item ID'])return;
    const expected=num3_(x['Purchase Price'])*Math.max(1,num3_(x['Quantity']))+num3_(x['Tax Paid'])+num3_(x['Acquisition Shipping']);
    if(Math.abs(num3_(x['Total Cost'])-expected)>0.01)add('HIGH','Inventory',i+2,'Total Cost does not equal Purchase Price × Quantity + Tax Paid + Acquisition Shipping.');
    const basis=num3_(x['Listed Price'])>0?num3_(x['Listed Price']):num3_(x['Expected Sale Price']);
    const projected=basis>0?basis-expected:'';
    if(projected!==''&&Math.abs(num3_(x['Projected Profit'])-projected)>0.01)add('MEDIUM','Inventory',i+2,'Projected Profit does not use Listed Price when available, otherwise Expected Sale Price, less Total Cost.');
    const roi=expected&&projected!==''?projected/expected:'';
    if(roi!==''&&Math.abs(num3_(x['Projected ROI %'])-roi)>0.0001)add('MEDIUM','Inventory',i+2,'Projected ROI % does not equal Projected Profit ÷ Total Cost.');
  });

  const saleRows=rows(sales,FTP3.SALES_HEADERS);
  saleRows.forEach((r,i)=>{
    const x=rowRecord3_(FTP3.SALES_HEADERS,r);if(!x['Sale ID'])return;
    const gross=num3_(x['Sale Price'])+num3_(x['Shipping Charged']);
    const costs=num3_(x['Shipping Actual'])+num3_(x['Packaging Cost'])+num3_(x['Marketplace Fees'])+num3_(x['Payment Fees'])+num3_(x['Promotion Expense']);
    const net=gross-costs,profit=net-num3_(x['Item Cost']);
    if(Math.abs(num3_(x['Gross Revenue'])-gross)>0.01)add('CRITICAL','Sales',i+2,'Gross Revenue is inconsistent.');
    if(Math.abs(num3_(x['Total Selling Costs'])-costs)>0.01)add('CRITICAL','Sales',i+2,'Total Selling Costs incorrectly includes or omits a cost.');
    if(Math.abs(num3_(x['Net Proceeds'])-net)>0.01)add('CRITICAL','Sales',i+2,'Net Proceeds is inconsistent.');
    if(Math.abs(num3_(x['Realized Profit'])-profit)>0.01)add('CRITICAL','Sales',i+2,'Realized Profit is inconsistent.');
    const refs=[['Box Used','Box Qty'],['Bubble Wrap Used','Bubble Wrap Qty'],['Mailer Used','Mailer Qty'],['Tape Used','Tape Qty'],['Other Packaging Used','Other Packaging Qty']];
    let calculated=0;refs.forEach(([idName,qtyName])=>{
      const raw=x[idName],qty=x[qtyName];if(!raw)return;
      const id=packagingIdFromSelection3_(raw);
      const item=packagingItemById3_(id);
      if(!item){
        add('HIGH','Sales',i+2,'Packaging ID does not exist on Packaging sheet: '+raw);
      }else{
        const p=rowRecord3_(FTP3.PACKAGING_HEADERS,item.values);
        if(String(p['Active']||'Yes').toLowerCase()==='no')add('MEDIUM','Sales',i+2,'Referenced packaging item is inactive: '+id);
        calculated+=num3_(p['Cost Per Unit'])*num3_(qty);
      }
    });
    if(String(x['Packaging Verified']||'')==='Yes'&&Math.abs(num3_(x['Packaging Cost'])-calculated)>0.011)add('HIGH','Sales',i+2,'Stored Packaging Cost does not match referenced Packaging IDs and quantities.');
  });

  rows(exp,FTP3.EXPENSE_HEADERS).forEach((r,i)=>{const x=rowRecord3_(FTP3.EXPENSE_HEADERS,r);if(!x['Expense ID'])return;const total=num3_(x['Subtotal'])+num3_(x['GST/HST Paid']);const rawUse=num3_(x['Business Use %']);const normalizedUse=rawUse>1?rawUse/100:rawUse;const ded=total*Math.min(1,Math.max(0,normalizedUse));if(Math.abs(num3_(x['Total'])-total)>0.01)add('HIGH','Expenses',i+2,'Total does not equal Subtotal + GST/HST Paid.');if(Math.abs(num3_(x['Deductible Amount'])-ded)>0.01)add('HIGH','Expenses',i+2,'Deductible Amount does not equal Total × Business Use %.');});
  rows(mil,FTP3.MILEAGE_HEADERS).forEach((r,i)=>{const x=rowRecord3_(FTP3.MILEAGE_HEADERS,r);if(!x['Trip ID'])return;const total=Math.max(0,num3_(x['End Odometer'])-num3_(x['Start Odometer']));if(Math.abs(num3_(x['Total Kilometres'])-total)>0.01)add('HIGH','Mileage',i+2,'Total Kilometres is inconsistent with odometers.');if(num3_(x['Business Kilometres'])>total+0.001)add('HIGH','Mileage',i+2,'Business Kilometres exceeds Total Kilometres.');if(Math.abs(num3_(x['Claim Amount'])-num3_(x['Business Kilometres'])*num3_(x['CRA Rate']))>0.01)add('HIGH','Mileage',i+2,'Claim Amount is inconsistent.');});
  rows(pkg,FTP3.PACKAGING_HEADERS).forEach((r,i)=>{
    const x=rowRecord3_(FTP3.PACKAGING_HEADERS,r);
    if(!x['Packaging ID'])return;
    const units=num3_(x['Units Purchased']);
    const expectedCpu=units?roundMoney3_(num3_(x['Purchase Cost'])/units):0;
    const actualCpu=roundMoney3_(num3_(x['Cost Per Unit']));
    if(Math.abs(actualCpu-expectedCpu)>0.009){
      add('MEDIUM','Packaging',i+2,'Cost Per Unit is inconsistent. Expected $'+expectedCpu.toFixed(2)+'.');
    }
    if(num3_(x['Quantity On Hand'])<0)add('CRITICAL','Packaging',i+2,'Quantity On Hand is negative.');
  });

  const q=sheet3_('Calculation Audit');q.clear();q.getRange(1,1,1,4).setValues([['Severity','Area','Row','Finding']]);header3_(q.getRange(1,1,1,4));
  if(issues.length)q.getRange(2,1,issues.length,4).setValues(issues);else q.getRange(2,1).setValue('PASS — no row-level calculation inconsistencies detected.');
  q.setFrozenRows(1);q.setColumnWidth(1,100);q.setColumnWidth(2,120);q.setColumnWidth(3,70);q.setColumnWidth(4,520);q.getRange('A:D').setWrap(true);
  SpreadsheetApp.getActive().setActiveSheet(q);SpreadsheetApp.getActive().toast(issues.length+' calculation issue(s) found.','FlipTracker Pro Audit',8);return issues;
}

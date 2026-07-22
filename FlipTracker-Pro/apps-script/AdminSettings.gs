function buildAdminSprint3_() {
  const s = sheet3_(FTP3.SHEETS.ADMIN);
  const lists = {
    Categories:['Electronics','Tools','Collectibles','Clothing','Furniture','Automotive','Household','Other'],
    PurchaseLocations:['Garage Sale','Thrift Store','Value Village','Facebook Marketplace','Auction','Retail Clearance','Other'],
    Marketplaces:['eBay','Facebook Marketplace','Kijiji','Poshmark','Etsy','Local Sale','Other'],
    Statuses:['Purchased','Needs Cleaning','Needs Testing','Ready to List','Listed','Sale Pending','Sold','Shipped','Archived'],
    Conditions:['New','Open Box','Like New','Good','Fair','For Parts'],
    StorageLocations:['Garage','Basement','Shelf A','Shelf B','Bin 1','Bin 2','Other'],
    ExpenseCategories:['Fuel','Packaging','Shipping Supplies','Advertising','Storage','Phone','Internet','Office','Software','Bank Fees','Equipment','Professional Fees','Other'],
    PaymentMethods:['Cash','Credit Card','Debit','PayPal','Bank Transfer','Other'],
    PackagingTypes:['Box','Bubble Wrap','Mailer','Tape','Label','Packing Paper','Other']
  };
  const existingLists = {};
  if (s.getLastColumn() > 0 && s.getLastRow() > 1) {
    const oldHeaders = s.getRange(1,1,1,s.getLastColumn()).getDisplayValues()[0];
    oldHeaders.forEach((name,index) => {
      if (!name) return;
      existingLists[name] = s.getRange(2,index+1,s.getLastRow()-1,1).getDisplayValues().flat().filter(Boolean);
    });
  }
  s.clear();
  let col = 1;
  Object.keys(lists).forEach(name => {
    s.getRange(1,col).setValue(name);
    const merged = lists[name].concat(existingLists[name] || []).filter((v,i,a) => a.indexOf(v) === i);
    const values = merged.map(v => [v]);
    s.getRange(2,col,values.length,1).setValues(values);
    const rangeName = 'FTP3_' + name;
    SpreadsheetApp.getActive().getNamedRanges()
      .filter(n => n.getName() === rangeName).forEach(n => n.remove());
    SpreadsheetApp.getActive().setNamedRange(rangeName,s.getRange(2,col,50,1));
    s.setColumnWidth(col,180);
    col++;
  });
  header3_(s.getRange(1,1,1,col-1));
  borders3_(s.getRange(1,1,50,col-1));
  s.setFrozenRows(1);
}

function buildSettingsSprint3_() {
  const s = sheet3_(FTP3.SHEETS.SETTINGS);
  const existing = {};
  if (s.getLastRow() >= 2) {
    s.getRange(2,1,s.getLastRow()-1,2).getValues().forEach(r => {
      if (r[0] !== '') existing[String(r[0])] = r[1];
    });
  }
  s.clear();
  const defaults = [
    ['Business Name','','Optional'],
    ['Business Number','','Optional'],
    ['GST/HST Registered','No','Yes or No'],
    ['GST/HST Number','','Optional'],
    ['Province','Ontario',''],
    ['Currency','CAD',''],
    ['Fiscal Year Start','January 1',''],
    ['Default HST Rate',0.13,'Editable'],
    ['CRA Mileage Rate',0,'Optional estimate; confirm tax treatment'],
    ['Tax Year',new Date().getFullYear(),'Tax Centre reporting year'],
    ['Low Packaging Stock Alert','Yes','Yes or No'],
    ['Opening Inventory Value',0,'Cost value at start of tax year'],
    ['Inventory Valuation Method','Lower of cost and FMV','Use a consistent CRA-accepted method'],
    ['GST/HST Reporting Method','Regular','Regular or Quick; Tax Centre estimates Regular only'],
    ['Include Mileage Estimate in Expenses','No','Professional review recommended']
  ];
  const values = [['Setting','Value','Notes']].concat(defaults.map(r => [r[0], Object.prototype.hasOwnProperty.call(existing,r[0]) ? existing[r[0]] : r[1], r[2]]));
  s.getRange(1,1,values.length,3).setValues(values);
  header3_(s.getRange(1,1,1,3));
  borders3_(s.getRange(1,1,values.length,3));
  s.setFrozenRows(1);
  s.setColumnWidths(1,3,220);
  s.getRange('B9').setNumberFormat('0.00%');
  s.getRange('B10').setNumberFormat('$#,##0.00');
  s.getRange('B13').setNumberFormat('$#,##0.00');
  const yesNo = SpreadsheetApp.newDataValidation().requireValueInList(['Yes','No'],true).setAllowInvalid(false).build();
  s.getRange('B4').setDataValidation(yesNo);
  s.getRange('B12').setDataValidation(yesNo);
  s.getRange('B16').setDataValidation(yesNo);
  s.getRange('B14').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Lower of cost and FMV','Fair market value of entire inventory','Cost'],true).setAllowInvalid(false).build());
  s.getRange('B15').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Regular','Quick'],true).setAllowInvalid(false).build());
}

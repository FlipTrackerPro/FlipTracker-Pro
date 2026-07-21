function buildAdminSprint3_() {
  const s = sheet3_(FTP3.SHEETS.ADMIN);
  s.clear();
  const lists = {
    Categories:['Electronics','Tools','Collectibles','Clothing','Furniture','Automotive','Household','Other'],
    PurchaseLocations:['Garage Sale','Thrift Store','Value Village','Facebook Marketplace','Auction','Retail Clearance','Other'],
    Marketplaces:['eBay','Facebook Marketplace','Kijiji','Poshmark','Etsy','Local Sale','Other'],
    Statuses:['Purchased','Needs Cleaning','Needs Testing','Ready to List','Listed','Sold','Shipped','Archived'],
    Conditions:['New','Open Box','Like New','Good','Fair','For Parts'],
    StorageLocations:['Garage','Basement','Shelf A','Shelf B','Bin 1','Bin 2','Other'],
    ExpenseCategories:['Fuel','Packaging','Shipping Supplies','Advertising','Storage','Phone','Internet','Office','Software','Bank Fees','Equipment','Professional Fees','Other'],
    PaymentMethods:['Cash','Credit Card','Debit','PayPal','Bank Transfer','Other'],
    PackagingTypes:['Box','Bubble Wrap','Mailer','Tape','Label','Packing Paper','Other']
  };
  let col = 1;
  Object.keys(lists).forEach(name => {
    s.getRange(1,col).setValue(name);
    const values = lists[name].map(v => [v]);
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
  s.clear();
  const values = [
    ['Setting','Value','Notes'],
    ['Business Name','','Optional'],
    ['Business Number','','Optional'],
    ['GST/HST Registered','No','Yes or No'],
    ['GST/HST Number','','Optional'],
    ['Province','Ontario',''],
    ['Currency','CAD',''],
    ['Fiscal Year Start','January 1',''],
    ['Default HST Rate',0.13,'Editable'],
    ['CRA Mileage Rate',0,'Enter the applicable tax-year rate'],
    ['Tax Year',new Date().getFullYear(),''],
    ['Low Packaging Stock Alert','Yes','Yes or No']
  ];
  s.getRange(1,1,values.length,3).setValues(values);
  header3_(s.getRange(1,1,1,3));
  borders3_(s.getRange(1,1,values.length,3));
  s.setFrozenRows(1);
  s.setColumnWidths(1,3,200);
  s.getRange('B9').setNumberFormat('0.00%');
  s.getRange('B10').setNumberFormat('$#,##0.00');
  const yesNo = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes','No'],true).setAllowInvalid(false).build();
  s.getRange('B4').setDataValidation(yesNo);
  s.getRange('B12').setDataValidation(yesNo);
}

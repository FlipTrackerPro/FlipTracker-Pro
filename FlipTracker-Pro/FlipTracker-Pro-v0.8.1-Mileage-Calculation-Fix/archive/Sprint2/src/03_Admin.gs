function buildAdminSprint2_() {
  const s = sheet_(FTP2.SHEETS.ADMIN);
  const lists = {
    Categories:['Electronics','Tools','Collectibles','Clothing','Furniture','Automotive','Household','Other'],
    PurchaseLocations:['Garage Sale','Thrift Store','Value Village','Facebook Marketplace','Auction','Retail Clearance','Other'],
    Marketplaces:['eBay','Facebook Marketplace','Kijiji','Poshmark','Etsy','Local Sale','Other'],
    Statuses:['Purchased','Needs Cleaning','Needs Testing','Ready to List','Listed','Sold','Shipped','Archived'],
    Conditions:['New','Open Box','Like New','Good','Fair','For Parts'],
    StorageLocations:['Garage','Basement','Shelf A','Shelf B','Bin 1','Bin 2','Other']
  };

  let c = 1;
  Object.keys(lists).forEach(name => {
    s.getRange(1,c).setValue(name);
    const values = lists[name].map(v => [v]);
    s.getRange(2,c,values.length,1).setValues(values);

    const named = 'FTP2_' + name;
    SpreadsheetApp.getActive().getNamedRanges()
      .filter(n => n.getName() === named)
      .forEach(n => n.remove());

    SpreadsheetApp.getActive().setNamedRange(named,s.getRange(2,c,50,1));
    s.setColumnWidth(c,180);
    c++;
  });

  header2_(s.getRange(1,1,1,c-1));
  s.setFrozenRows(1);
}

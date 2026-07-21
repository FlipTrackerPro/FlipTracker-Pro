function buildDashboardSprint2_() {
  const s = sheet_(FTP2.SHEETS.DASHBOARD);
  s.clear();
  s.setHiddenGridlines(true);

  s.getRange('A1:H2').merge().setValue('FlipTracker Pro — Sprint 2')
    .setBackground(FTP2.COLORS.NAVY).setFontColor(FTP2.COLORS.WHITE)
    .setFontSize(22).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  const cards = [
    ['A4:B6','Inventory Cost','=SUM(Inventory!N2:N)','$#,##0.00'],
    ['C4:D6','Potential Value','=SUM(Inventory!O2:O)','$#,##0.00'],
    ['E4:F6','Potential Profit','=SUM(Inventory!U2:U)','$#,##0.00'],
    ['G4:H6','Average ROI','=IFERROR(AVERAGE(FILTER(Inventory!V2:V,Inventory!C2:C<>"")),0)','0.0%'],
    ['A8:B10','Items Tracked','=COUNTA(Inventory!C2:C)','0'],
    ['C8:D10','Items Listed','=COUNTIF(Inventory!S2:S,"Listed")','0'],
    ['E8:F10','Ready to List','=COUNTIF(Inventory!S2:S,"Ready to List")','0'],
    ['G8:H10','90+ Days','=COUNTIFS(Inventory!T2:T,">=90",Inventory!S2:S,"<>Sold",Inventory!C2:C,"<>")','0']
  ];

  cards.forEach(([a,title,formula,format]) => {
    const r = s.getRange(a);
    r.merge();
    r.getCell(1,1).setFormula(
      `="${title}"&CHAR(10)&TEXT(${formula.substring(1)},"${format}")`
    );
    r.setBackground(FTP2.COLORS.LIGHT_BLUE)
      .setFontWeight('bold').setFontSize(14).setWrap(true)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,FTP2.COLORS.BORDER,
        SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });

  s.getRange('A13:H13').merge().setValue('Inventory Workflow')
    .setBackground(FTP2.COLORS.NAVY).setFontColor(FTP2.COLORS.WHITE)
    .setFontWeight('bold').setHorizontalAlignment('center');

  s.getRange('A14:H18').merge().setValue(
    'Use FlipTracker Pro → Add Inventory Item to create stable inventory records. ' +
    'Select a row and choose Edit Selected Item to update it. Find Item searches ' +
    'Item ID, title, SKU, and barcode. Show Slow Inventory filters items held for ' +
    '90 days or longer.'
  ).setBackground('#F3F6F9').setWrap(true).setVerticalAlignment('middle');

  s.setColumnWidths(1,8,120);
}

function refreshDashboardSprint2() {
  buildDashboardSprint2_();
  SpreadsheetApp.getActive().setActiveSheet(sheet_(FTP2.SHEETS.DASHBOARD));
}

function goToInventorySprint2() {
  SpreadsheetApp.getActive().setActiveSheet(sheet_(FTP2.SHEETS.INVENTORY));
}

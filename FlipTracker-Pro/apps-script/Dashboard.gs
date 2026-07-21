function buildDashboardSprint3_() {
  const s = sheet3_(FTP3.SHEETS.DASHBOARD);
  s.clear();
  s.setHiddenGridlines(true);
  s.getRange('A1:H2').merge().setValue('FlipTracker Pro — Sprint 3')
    .setBackground(FTP3.COLORS.NAVY).setFontColor(FTP3.COLORS.WHITE)
    .setFontSize(22).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  const cards = [
    ['A4:B6','Inventory Cost','SUM(Inventory!N2:N)','$#,##0.00'],
    ['C4:D6','Gross Sales','SUM(Sales!N2:N)','$#,##0.00'],
    ['E4:F6','Realized Profit','SUM(Sales!Q2:Q)','$#,##0.00'],
    ['G4:H6','Average ROI','IFERROR(AVERAGE(FILTER(Sales!R2:R,Sales!A2:A<>"")),0)','0.0%'],
    ['A8:B10','Items Sold','COUNTA(Sales!A2:A)','0'],
    ['C8:D10','Business Expenses','SUM(Expenses!J2:J)','$#,##0.00'],
    ['E8:F10','Mileage Claims','SUM(Mileage!K2:K)','$#,##0.00'],
    ['G8:H10','90+ Day Items','COUNTIFS(Inventory!T2:T,">=90",Inventory!S2:S,"<>Sold",Inventory!C2:C,"<>")','0']
  ];

  cards.forEach(([a1,title,formula,fmt]) => {
    const r = s.getRange(a1);
    r.merge();
    r.getCell(1,1).setFormula(`="${title}"&CHAR(10)&TEXT(${formula},"${fmt}")`);
    r.setBackground(FTP3.COLORS.LIGHT_BLUE).setFontColor(FTP3.COLORS.TEXT)
      .setFontWeight('bold').setFontSize(14).setWrap(true)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,FTP3.COLORS.BORDER,
        SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });

  s.getRange('A13:H13').merge().setValue('Sprint 3 Transaction Workflow');
  header3_(s.getRange('A13:H13'));
  s.getRange('A14:H18').merge().setValue(
    'Record a sale from the FlipTracker Pro menu. Sprint 3 automatically retrieves ' +
    'the item cost, calculates gross revenue, total selling costs, net proceeds, ' +
    'realized profit, ROI, and days to sell, then marks the inventory item Sold. ' +
    'Expenses, mileage, and packaging supplies are now tracked in dedicated sheets.'
  ).setBackground(FTP3.COLORS.GRAY).setWrap(true).setVerticalAlignment('middle');

  s.getRange('A20:D20').merge().setValue('Profit after operating expenses');
  s.getRange('E20:H20').merge().setFormula('=SUM(Sales!Q2:Q)-SUM(Expenses!J2:J)-SUM(Mileage!K2:K)');
  header3_(s.getRange('A20:D20'));
  s.getRange('E20:H20').setBackground(FTP3.COLORS.LIGHT_GREEN)
    .setFontWeight('bold').setFontSize(16).setNumberFormat('$#,##0.00')
    .setHorizontalAlignment('center');

  s.setColumnWidths(1,8,120);
  s.setFrozenRows(2);
}

function refreshDashboardSprint3() {
  buildDashboardSprint3_();
}

function goToDashboardSprint3() {
  SpreadsheetApp.getActive().setActiveSheet(sheet3_(FTP3.SHEETS.DASHBOARD));
}

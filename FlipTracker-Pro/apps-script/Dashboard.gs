function buildDashboardSprint3_() {
  const s = sheet3_(FTP3.SHEETS.DASHBOARD);
  s.clear();
  s.setHiddenGridlines(true);
  s.getRange('A1:H2').merge().setValue('FlipTracker Pro — v0.4')
    .setBackground(FTP3.COLORS.NAVY).setFontColor(FTP3.COLORS.WHITE)
    .setFontSize(22).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  const cards = [
    ['A4:B6','Current Inventory Cost','SUMIFS(Inventory!N2:N,Inventory!S2:S,"<>Sold",Inventory!S2:S,"<>Archived")','$#,##0.00'],
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

  s.getRange('A13:H13').merge().setValue('v0.4 Transaction and Tax Workflow');
  header3_(s.getRange('A13:H13'));
  s.getRange('A14:H18').merge().setValue(
    'Record a sale from the FlipTracker Pro menu. Sprint 3 automatically retrieves ' +
    'the item cost, calculates gross revenue, total selling costs, net proceeds, ' +
    'realized profit, ROI, and days to sell, then marks the inventory item Sold. ' +
    'Expenses, mileage, packaging supplies, and tax-year estimates are tracked in dedicated sheets. Open the CRA Tax Centre from the menu for COGS, inventory, GST/HST, and accountant summaries.'
  ).setBackground(FTP3.COLORS.GRAY).setWrap(true).setVerticalAlignment('middle');

  s.getRange('A20:D20').merge().setValue('Profit after operating expenses');
  s.getRange('E20:H20').merge().setFormula('=SUM(Sales!Q2:Q)-SUM(Expenses!J2:J)-IF(IFERROR(VLOOKUP("Include Mileage Estimate in Expenses",Settings!A:B,2,FALSE),"No")="Yes",SUM(Mileage!K2:K),0)');
  header3_(s.getRange('A20:D20'));
  s.getRange('E20:H20').setBackground(FTP3.COLORS.LIGHT_GREEN)
    .setFontWeight('bold').setFontSize(16).setNumberFormat('$#,##0.00')
    .setHorizontalAlignment('center');

  s.getRange('A23:H23').merge().setValue('Packaging Inventory');header3_(s.getRange('A23:H23'));
  const pkgCards=[
    ['A24:B26','Packaging Inventory Value','SUMPRODUCT(Packaging!H2:H,Packaging!I2:I)','$#,##0.00'],
    ['C24:D26','Low-Stock Supplies','COUNTIFS(Packaging!A2:A,"<>",Packaging!N2:N,"Yes",Packaging!I2:I,"<="&Packaging!J2:J)','0'],
    ['E24:F26','Packaging Cost This Year','SUMIFS(Sales!H2:H,Sales!C2:C,">="&DATE(YEAR(TODAY()),1,1),Sales!C2:C,"<"&DATE(YEAR(TODAY())+1,1,1))','$#,##0.00'],
    ['G24:H26','Average Packaging / Sale','IFERROR(AVERAGE(FILTER(Sales!H2:H,Sales!A2:A<>"")),0)','$#,##0.00']
  ];
  pkgCards.forEach(([a1,title,formula,fmt])=>{const r=s.getRange(a1);r.merge();r.getCell(1,1).setFormula(`="${title}"&CHAR(10)&TEXT(${formula},"${fmt}")`);r.setBackground(FTP3.COLORS.GOLD).setFontColor(FTP3.COLORS.TEXT).setFontWeight('bold').setFontSize(13).setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle').setBorder(true,true,true,true,false,false,FTP3.COLORS.BORDER,SpreadsheetApp.BorderStyle.SOLID_MEDIUM);});

  s.setColumnWidths(1,8,120);
  s.setFrozenRows(2);
}

function refreshDashboardSprint3() {
  buildDashboardSprint3_();
}

function goToDashboardSprint3() {
  SpreadsheetApp.getActive().setActiveSheet(sheet3_(FTP3.SHEETS.DASHBOARD));
}

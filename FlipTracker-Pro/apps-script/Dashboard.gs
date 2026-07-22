function buildDashboardSprint3_() {
  const s=sheet3_(FTP3.SHEETS.DASHBOARD);s.clear();s.setHiddenGridlines(true);
  s.getRange('A1:H2').merge().setValue('FlipTracker Pro — v0.5').setBackground(FTP3.COLORS.NAVY).setFontColor(FTP3.COLORS.WHITE).setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');

  const iTotal=sheetColumnLetter3_(FTP3.SHEETS.INVENTORY,'Total Cost'),iStatus=sheetColumnLetter3_(FTP3.SHEETS.INVENTORY,'Status'),iDays=sheetColumnLetter3_(FTP3.SHEETS.INVENTORY,'Days in Inventory'),iDesc=sheetColumnLetter3_(FTP3.SHEETS.INVENTORY,'Description');
  const sGross=sheetColumnLetter3_(FTP3.SHEETS.SALES,'Gross Revenue'),sProfit=sheetColumnLetter3_(FTP3.SHEETS.SALES,'Realized Profit'),sRoi=sheetColumnLetter3_(FTP3.SHEETS.SALES,'Realized ROI %'),sId=sheetColumnLetter3_(FTP3.SHEETS.SALES,'Sale ID');
  const sPkg=sheetColumnLetter3_(FTP3.SHEETS.SALES,'Packaging Cost'),sDate=sheetColumnLetter3_(FTP3.SHEETS.SALES,'Sale Date');
  const cards=[
    ['A4:B6','Current Inventory Cost',`SUMIFS(Inventory!${iTotal}2:${iTotal},Inventory!${iStatus}2:${iStatus},"<>Sold",Inventory!${iStatus}2:${iStatus},"<>Archived")`,'$#,##0.00'],
    ['C4:D6','Gross Sales',`SUM(Sales!${sGross}2:${sGross})`,'$#,##0.00'],
    ['E4:F6','Realized Profit',`SUM(Sales!${sProfit}2:${sProfit})`,'$#,##0.00'],
    ['G4:H6','Average ROI',`IFERROR(AVERAGE(FILTER(Sales!${sRoi}2:${sRoi},Sales!${sId}2:${sId}<>"")),0)`,'0.0%'],
    ['A8:B10','Items Sold',`COUNTA(Sales!${sId}2:${sId})`,'0'],
    ['C8:D10','Business Expenses','SUM(Expenses!J2:J)','$#,##0.00'],
    ['E8:F10','Mileage Claims','SUM(Mileage!K2:K)','$#,##0.00'],
    ['G8:H10','90+ Day Items',`COUNTIFS(Inventory!${iDays}2:${iDays},">=90",Inventory!${iStatus}2:${iStatus},"<>Sold",Inventory!${iDesc}2:${iDesc},"<>")`,'0']
  ];
  cards.forEach(([a1,title,formula,fmt])=>{const r=s.getRange(a1);r.merge();r.getCell(1,1).setFormula(`="${title}"&CHAR(10)&TEXT(${formula},"${fmt}")`);r.setBackground(FTP3.COLORS.LIGHT_BLUE).setFontColor(FTP3.COLORS.TEXT).setFontWeight('bold').setFontSize(14).setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle').setBorder(true,true,true,true,false,false,FTP3.COLORS.BORDER,SpreadsheetApp.BorderStyle.SOLID_MEDIUM);});
  s.getRange('A13:H13').merge().setValue('Transaction and Tax Workflow');header3_(s.getRange('A13:H13'));
  s.getRange('A14:H18').merge().setValue('Inventory follows the item lifecycle from identification and purchase through listing and sale. Total Cost includes Purchase Price × Quantity, Tax Paid, and Acquisition Shipping. Sales and tax reporting use header-based mappings to prevent column-order errors.').setBackground(FTP3.COLORS.GRAY).setWrap(true).setVerticalAlignment('middle');
  s.getRange('A20:D20').merge().setValue('Profit after operating expenses');
  s.getRange('E20:H20').merge().setFormula(`=SUM(Sales!${sProfit}2:${sProfit})-SUM(Expenses!J2:J)-IF(IFERROR(VLOOKUP("Include Mileage Estimate in Expenses",Settings!A:B,2,FALSE),"No")="Yes",SUM(Mileage!K2:K),0)`);
  header3_(s.getRange('A20:D20'));s.getRange('E20:H20').setBackground(FTP3.COLORS.LIGHT_GREEN).setFontWeight('bold').setFontSize(16).setNumberFormat('$#,##0.00').setHorizontalAlignment('center');
  s.getRange('A23:H23').merge().setValue('Packaging Inventory');header3_(s.getRange('A23:H23'));
  const pkgCards=[
    ['A24:B26','Packaging Inventory Value','SUMPRODUCT(Packaging!H2:H,Packaging!I2:I)','$#,##0.00'],
    ['C24:D26','Low-Stock Supplies','COUNTIFS(Packaging!A2:A,"<>",Packaging!N2:N,"Yes",Packaging!I2:I,"<="&Packaging!J2:J)','0'],
    ['E24:F26','Packaging Cost This Year',`SUMIFS(Sales!${sPkg}2:${sPkg},Sales!${sDate}2:${sDate},">="&DATE(YEAR(TODAY()),1,1),Sales!${sDate}2:${sDate},"<"&DATE(YEAR(TODAY())+1,1,1))`,'$#,##0.00'],
    ['G24:H26','Average Packaging / Sale',`IFERROR(AVERAGE(FILTER(Sales!${sPkg}2:${sPkg},Sales!${sId}2:${sId}<>"")),0)`,'$#,##0.00']
  ];
  pkgCards.forEach(([a1,title,formula,fmt])=>{const r=s.getRange(a1);r.merge();r.getCell(1,1).setFormula(`="${title}"&CHAR(10)&TEXT(${formula},"${fmt}")`);r.setBackground(FTP3.COLORS.GOLD).setFontColor(FTP3.COLORS.TEXT).setFontWeight('bold').setFontSize(13).setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle').setBorder(true,true,true,true,false,false,FTP3.COLORS.BORDER,SpreadsheetApp.BorderStyle.SOLID_MEDIUM);});
  s.setColumnWidths(1,8,120);s.setFrozenRows(2);
}
function refreshDashboardSprint3(){buildDashboardSprint3_();}
function goToDashboardSprint3(){SpreadsheetApp.getActive().setActiveSheet(sheet3_(FTP3.SHEETS.DASHBOARD));}

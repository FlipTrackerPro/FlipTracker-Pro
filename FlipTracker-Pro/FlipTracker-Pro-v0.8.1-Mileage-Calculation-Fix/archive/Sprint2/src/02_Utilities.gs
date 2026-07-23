function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function header2_(range) {
  range.setBackground(FTP2.COLORS.NAVY)
    .setFontColor(FTP2.COLORS.WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
}

function num_(v) {
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

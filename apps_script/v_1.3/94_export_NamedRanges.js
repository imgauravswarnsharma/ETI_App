function export_NamedRanges() {
  const ss = SpreadsheetApp.getActive();
  const out = ss.insertSheet('_Named_Ranges');

  out.appendRow(['Name','Range']);

  ss.getNamedRanges().forEach(nr => {
    out.appendRow([nr.getName(), nr.getRange().getA1Notation()]);
  });
}

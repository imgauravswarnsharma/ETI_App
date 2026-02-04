function export_ConditionalFormatting() {
  const ss = SpreadsheetApp.getActive();
  const out = ss.insertSheet('_Conditional_Formatting');

  out.appendRow(['Sheet','Ranges','Rule_Type']);

  ss.getSheets().forEach(sh => {
    sh.getConditionalFormatRules().forEach(rule => {
      out.appendRow([
        sh.getName(),
        rule.getRanges().map(r => r.getA1Notation()).join(', '),
        rule.getBooleanCondition()?.getCriteriaType() || 'GRADIENT'
      ]);
    });
  });
}

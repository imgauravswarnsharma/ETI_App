function export_DataValidations() {
  const ss = SpreadsheetApp.getActive();
  const out = ss.insertSheet('_Data_Validations');

  out.appendRow(['Sheet','Range','Rule']);

  ss.getSheets().forEach(sh => {
    const range = sh.getDataRange();
    const validations = range.getDataValidations();

    validations.forEach((row, r) => {
      row.forEach((rule, c) => {
        if (rule) {
          out.appendRow([
            sh.getName(),
            sh.getRange(r+1,c+1).getA1Notation(),
            rule.getCriteriaType()
          ]);
        }
      });
    });
  });
}

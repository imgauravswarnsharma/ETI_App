/**
 * Config_Conditional_Formatting — Manifest Extract v1 (API-safe)
 * Read-only, safe
 */

function extractConditionalFormattingConfig() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let out = ss.getSheetByName('Config_Conditional_Formatting');
  if (!out) out = ss.insertSheet('Config_Conditional_Formatting');
  out.clear();

  const headers = [
    'Sheet_Name',
    'Applies_To_Range',
    'Rule_Type',
    'Formula_or_Criteria',
    'Notes'
  ];

  out.getRange(1, 1, 1, headers.length)
     .setValues([headers])
     .setFontWeight('bold');

  let writeRow = 4;
  let lastSheet = null;

  ss.getSheets().forEach(sh => {

    const sheetName = sh.getName();
    const rules = sh.getConditionalFormatRules();
    if (!rules || rules.length === 0) return;

    if (lastSheet !== null && sheetName !== lastSheet) {
      writeRow += 2;
    }

    rules.forEach(rule => {

      const ranges = rule.getRanges()
        .map(r => r.getA1Notation())
        .join(', ');

      const booleanCond = rule.getBooleanCondition();
      const gradientCond = rule.getGradientCondition();

      let ruleType = '';
      let criteria = '';
      let notes = '';

      if (booleanCond) {
        ruleType = booleanCond.getCriteriaType();
        const vals = booleanCond.getCriteriaValues();
        criteria = vals && vals.length ? vals.join(' | ') : '';
      }

      if (gradientCond) {
        ruleType = 'GRADIENT';
        notes = [
          `MinColor=${gradientCond.getMinColor()}`,
          gradientCond.getMidColor() ? `MidColor=${gradientCond.getMidColor()}` : '',
          `MaxColor=${gradientCond.getMaxColor()}`
        ].filter(Boolean).join(' ; ');
      }

      out.getRange(writeRow, 1, 1, headers.length).setValues([[
        sheetName,
        ranges,
        ruleType,
        criteria,
        notes
      ]]);

      writeRow++;
    });

    lastSheet = sheetName;
  });
}

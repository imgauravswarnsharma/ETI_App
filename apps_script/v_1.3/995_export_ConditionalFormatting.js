/**
 * Script Name: extractConditionalFormattingConfig
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1
 * Current Status: ACTIVE (Performance Optimized)
 *
 * Purpose:
 * - Extract conditional formatting configuration as read-only manifest
 *
 * Preconditions:
 * - Active spreadsheet
 *
 * Algorithm (Optimized – Logic Unchanged):
 * 1. Read conditional rules per sheet
 * 2. Build output in memory
 * 3. Single batch write
 */

function extractConditionalFormattingConfig() {

  const dataSS = SpreadsheetApp.getActiveSpreadsheet();
  const metaSS = getMetadataSpreadsheet_();

  let out = metaSS.getSheetByName('Config_Conditional_Formatting');
  if (!out) out = metaSS.insertSheet('Config_Conditional_Formatting');
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

  let output = [];
  let lastSheet = null;

  dataSS.getSheets().forEach(sh => {

    const sheetName = sh.getName();
    const rules = sh.getConditionalFormatRules();
    if (!rules || rules.length === 0) return;

    if (lastSheet !== null && sheetName !== lastSheet) {
      output.push(new Array(headers.length).fill(''));
      output.push(new Array(headers.length).fill(''));
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

        const midPoint = gradientCond.getMidpoint();

        notes = [
          `MinColor=${gradientCond.getMinColor()}`,
          midPoint && midPoint.getColor() ? `MidColor=${midPoint.getColor()}` : '',
          `MaxColor=${gradientCond.getMaxColor()}`
        ].filter(Boolean).join(' ; ');
      }

      output.push([
        sheetName,
        ranges,
        ruleType,
        criteria,
        notes
      ]);
    });

    lastSheet = sheetName;
  });

  if (output.length > 0) {
    out.getRange(4, 1, output.length, headers.length)
       .setValues(output);
  }
}
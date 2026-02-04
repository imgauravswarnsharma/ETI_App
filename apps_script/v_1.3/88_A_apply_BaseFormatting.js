function applyBaseFormatting_(sheet) {

  /* ===============================
     1. Freeze first two rows
     =============================== */
  sheet.setFrozenRows(2);

  /* ===============================
     2. Bold header row (Row 1)
     =============================== */
  const headerRange = sheet.getRange(
    1,
    1,
    1,
    sheet.getMaxColumns()
  );
  headerRange.setFontWeight('bold');

  /* ===============================
     3. Global alignment
     =============================== */
  const fullRange = sheet.getRange(
    1,
    1,
    sheet.getMaxRows(),
    sheet.getMaxColumns()
  );

  fullRange
    .setHorizontalAlignment('left')
    .setVerticalAlignment('top');
}

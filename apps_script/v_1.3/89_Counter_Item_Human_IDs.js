function generateItemID_H() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const itemsSh = ss.getSheetByName('Lookup_Items');
  const ctrlSh  = ss.getSheetByName('Counter_Control');
  if (!itemsSh || !ctrlSh) throw new Error('Required sheet not found');

  // Read control table
  const ctrlData = ctrlSh.getDataRange().getValues();
  const ctrlHeader = ctrlData[0];

  const colCtrl = n => ctrlHeader.indexOf(n);

  const IDX_CTRL = {
    entity: colCtrl('Entity_Key'),
    total: colCtrl('Total_Counter')
  };

  for (const [k, v] of Object.entries(IDX_CTRL)) {
    if (v === -1) throw new Error(`Missing control column: ${k}`);
  }

  // Locate ITEM row
  let ctrlRow = -1;
  for (let i = 1; i < ctrlData.length; i++) {
    if (ctrlData[i][IDX_CTRL.entity] === 'ITEM') {
      ctrlRow = i;
      break;
    }
  }
  if (ctrlRow === -1) throw new Error('ITEM row not found in Counter_Control');

  let counter = Number(ctrlData[ctrlRow][IDX_CTRL.total]) || 0;

  // Read items
  const data = itemsSh.getDataRange().getValues();
  if (data.length < 2) return;

  const header = data[0];
  const col = n => header.indexOf(n);

  const IDX_ITEMS = {
    itemIdM: col('Item_ID_M'),
    itemIdH: col('Item_ID_H')
  };

  for (const [k, v] of Object.entries(IDX_ITEMS)) {
    if (v === -1) throw new Error(`Missing item column: ${k}`);
  }

  let wrote = false;

  for (let i = 1; i < data.length; i++) {
    const r = data[i];

    if (r[IDX_ITEMS.itemIdM] && !r[IDX_ITEMS.itemIdH]) {
      counter += 1;
      const humanId = 'ITEM-' + String(counter).padStart(6, '0');
      itemsSh.getRange(i + 1, IDX_ITEMS.itemIdH + 1).setValue(humanId);
      wrote = true;
    }
  }

  // Persist counter only if changes were made
  if (wrote) {
    ctrlSh.getRange(ctrlRow + 1, IDX_CTRL.total + 1).setValue(counter);
  }
}

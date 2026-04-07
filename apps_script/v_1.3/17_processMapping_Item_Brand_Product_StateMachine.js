/**
 * Script Name: processMapping_Item_Brand_Product_StateMachine
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Reconcile Item–Brand–Product mappings with current entity states
 * - Update snapshot status columns
 * - Derive mapping flags deterministically
 * - Repair drift caused by entity lifecycle changes
 *
 * Preconditions:
 * - Sheets must exist:
 *     Mapping_Item_Brand_Product
 *     Lookup_Items
 *     Lookup_Brands
 *     Lookup_Products
 *     Automation_Control
 */

function processMapping_Item_Brand_Product_StateMachine() {

  const EXECUTION_ID = Utilities.getUuid();
  const SCRIPT_NAME = 'processMapping_Item_Brand_Product_StateMachine';

  const MAP_SHEET   = 'Mapping_Item_Brand_Product';
  const ITEM_SHEET  = 'Lookup_Items';
  const BRAND_SHEET = 'Lookup_Brands';
  const PROD_SHEET  = 'Lookup_Products';
  const CTRL_SHEET  = 'Automation_Control';

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  /* ======================================
     CONTROL SWITCH
  ====================================== */

  const ctrl = ss.getSheetByName(CTRL_SHEET);
  if (!ctrl) throw new Error('Automation_Control sheet missing');

  const runIntegrity = ctrl.getRange("K2").getValue();

  if (runIntegrity !== true) {
    console.log(`[${SCRIPT_NAME}] skipped via control switch`);
    return;
  }

  console.log(`[${SCRIPT_NAME}] START`);

  /* ======================================
     LOAD SHEETS
  ====================================== */

  const mapSh   = ss.getSheetByName(MAP_SHEET);
  const itemSh  = ss.getSheetByName(ITEM_SHEET);
  const brandSh = ss.getSheetByName(BRAND_SHEET);
  const prodSh  = ss.getSheetByName(PROD_SHEET);

  if (!mapSh || !itemSh || !brandSh || !prodSh) {
    throw new Error('Required sheet missing');
  }

  const mapData = mapSh.getDataRange().getValues();
  const mapHdr  = mapData[0];

  const col = n => mapHdr.indexOf(n);

  const IDX = {

    itemCanon: col('Item_Name_Canonical'),
    brandCanon: col('Brand_Name_Canonical'),
    productCanon: col('Product_Name_Canonical'),

    itemStatus: col('Item_Status_Snapshot'),
    brandStatus: col('Brand_Status_Snapshot'),
    productStatus: col('Product_Status_Snapshot'),

    mapActive: col('Is_Mapping_Active'),
    analytics: col('Is_Analytics_Enabled'),
    archived: col('Is_Archived'),

    notes: col('Notes'),

    itemId: col('Item_ID_Machine'),
    brandId: col('Brand_ID_Machine'),
    productId: col('Product_ID_Machine')
  };

  for (const [k,v] of Object.entries(IDX)) {
    if (v === -1) throw new Error(`Missing column: ${k}`);
  }

  /* ======================================
     BUILD ENTITY STATE MAPS
  ====================================== */

  function buildStateMap(sheet, idColName) {

    const data = sheet.getDataRange().getValues();
    const hdr  = data[0];
    const c = n => hdr.indexOf(n);

    const IDX_STATE = {
      id: c(idColName),
      approved: c('Is_Approved'),
      active: c('Is_Active'),
      archived: c('Is_Archived')
    };

    const stateMap = {};

    for (let i = 1; i < data.length; i++) {

      const r = data[i];
      const id = r[IDX_STATE.id];

      if (!id) continue;

      stateMap[id] = {

        approved: r[IDX_STATE.approved],
        active: r[IDX_STATE.active],
        archived: r[IDX_STATE.archived]
      };
    }

    return stateMap;
  }

  const itemState  = buildStateMap(itemSh,  'Item_ID_Machine');
  const brandState = buildStateMap(brandSh, 'Brand_ID_Machine');
  const prodState  = buildStateMap(prodSh,  'Product_ID_Machine');

  /* ======================================
     PROCESS MAPPINGS
  ====================================== */

  let repaired = 0;
  let valid = 0;

  function resolveStatus(state) {

    if (!state) return 'Unknown';

    if (state.archived) return 'Archived';
    if (state.active)   return 'Active';
    if (state.approved) return 'Approved (Hidden Dropdown)';

    return 'Rejected';
  }

  for (let i = 1; i < mapData.length; i++) {

    const row = mapData[i];

    const itemId    = row[IDX.itemId];
    const brandId   = row[IDX.brandId];
    const productId = row[IDX.productId];

    const itemStatus   = resolveStatus(itemState[itemId]);
    const brandStatus  = resolveStatus(brandState[brandId]);
    const productStatus= resolveStatus(prodState[productId]);

    const prevActive = row[IDX.mapActive];

    const newActive =
      !(itemStatus === 'Archived' ||
        brandStatus === 'Archived' ||
        productStatus === 'Archived');

    row[IDX.itemStatus]   = itemStatus;
    row[IDX.brandStatus]  = brandStatus;
    row[IDX.productStatus]= productStatus;

    row[IDX.mapActive] = newActive;

    row[IDX.analytics] = true;

    if (prevActive !== newActive) {

      repaired++;

      row[IDX.notes] =
        `Mapping state updated due to entity lifecycle change`;
    }
    else {

      valid++;
    }
  }

  /* ======================================
     WRITE BACK
  ====================================== */

  mapSh
    .getRange(2,1,mapData.length-1,mapHdr.length)
    .setValues(mapData.slice(1));

  console.log(
    `[${SCRIPT_NAME}] VALID=${valid} REPAIRED=${repaired}`
  );

}
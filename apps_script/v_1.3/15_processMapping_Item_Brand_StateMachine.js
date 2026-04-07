/**
 * Script Name: processMapping_Item_Brand_StateMachine
 * Script Language: Google Apps Script (JavaScript)
 * Version Introduced: v1.3
 * Current Status: ACTIVE
 *
 * Purpose:
 * - Reconcile mapping rows against current entity governance state
 * - Update status snapshots
 * - Derive mapping flags deterministically
 * - Repair drift caused by entity lifecycle changes
 *
 * Preconditions:
 * - Sheet exists: Mapping_Item_Brand
 * - Sheet exists: Lookup_Items
 * - Sheet exists: Lookup_Brands
 * - Sheet exists: Automation_Control
 */

function processMapping_Item_Brand_StateMachine() {

  const EXECUTION_ID = Utilities.getUuid();
  const SCRIPT_NAME = 'processMapping_Item_Brand_StateMachine';

  const MAP_SHEET = 'Mapping_Item_Brand';
  const ITEM_SHEET = 'Lookup_Items';
  const BRAND_SHEET = 'Lookup_Brands';
  const CTRL_SHEET = 'Automation_Control';

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const ctrl = ss.getSheetByName(CTRL_SHEET);
  if (!ctrl) throw new Error('Automation_Control sheet missing');

  const runIntegrity = ctrl.getRange("K2").getValue();

  if (runIntegrity !== true) {
    console.log(`[${SCRIPT_NAME}] skipped via control switch`);
    return;
  }

  const mapSh = ss.getSheetByName(MAP_SHEET);
  const itemSh = ss.getSheetByName(ITEM_SHEET);
  const brandSh = ss.getSheetByName(BRAND_SHEET);

  if (!mapSh || !itemSh || !brandSh) {
    throw new Error('Required sheet missing');
  }

  const mapData = mapSh.getDataRange().getValues();
  const mapHdr = mapData[0];

  const col = n => mapHdr.indexOf(n);

  const IDX = {

    itemCanon: col('Item_Name_Canonical'),
    brandCanon: col('Brand_Name_Canonical'),

    itemStatus: col('Item_Status_Snapshot'),
    brandStatus: col('Brand_Status_Snapshot'),

    mapActive: col('Is_Mapping_Active'),
    analytics: col('Is_Analytics_Enabled'),
    archived: col('Is_Archived'),

    notes: col('Notes'),

    itemId: col('Item_ID_Machine'),
    brandId: col('Brand_ID_Machine')
  };

  for (const [k,v] of Object.entries(IDX)) {
    if (v === -1) throw new Error(`Missing column: ${k}`);
  }

  /* =========================
     BUILD ITEM STATE MAP
     ========================= */

  const itemData = itemSh.getDataRange().getValues();
  const itemHdr = itemData[0];

  const ic = n => itemHdr.indexOf(n);

  const IDX_ITEM = {
    id: ic('Item_ID_Machine'),
    approved: ic('Is_Approved'),
    active: ic('Is_Active'),
    archived: ic('Is_Archived')
  };

  const itemState = {};

  for (let i = 1; i < itemData.length; i++) {

    const r = itemData[i];

    const id = r[IDX_ITEM.id];

    if (!id) continue;

    itemState[id] = {

      approved: r[IDX_ITEM.approved],
      active: r[IDX_ITEM.active],
      archived: r[IDX_ITEM.archived]
    };
  }

  /* =========================
     BUILD BRAND STATE MAP
     ========================= */

  const brandData = brandSh.getDataRange().getValues();
  const brandHdr = brandData[0];

  const bc = n => brandHdr.indexOf(n);

  const IDX_BRAND = {
    id: bc('Brand_ID_Machine'),
    approved: bc('Is_Approved'),
    active: bc('Is_Active'),
    archived: bc('Is_Archived')
  };

  const brandState = {};

  for (let i = 1; i < brandData.length; i++) {

    const r = brandData[i];

    const id = r[IDX_BRAND.id];

    if (!id) continue;

    brandState[id] = {

      approved: r[IDX_BRAND.approved],
      active: r[IDX_BRAND.active],
      archived: r[IDX_BRAND.archived]
    };
  }

  /* =========================
     PROCESS MAPPINGS
     ========================= */

  let repaired = 0;
  let valid = 0;

  for (let i = 1; i < mapData.length; i++) {

    const row = mapData[i];

    const itemId = row[IDX.itemId];
    const brandId = row[IDX.brandId];

    const item = itemState[itemId];
    const brand = brandState[brandId];

    let itemStatus = 'Unknown';
    let brandStatus = 'Unknown';

    if (item) {

      if (item.archived) itemStatus = 'Archived';
      else if (item.active) itemStatus = 'Active';
      else if (item.approved) itemStatus = 'Approved (Hidden Dropdown)';
    }

    if (brand) {

      if (brand.archived) brandStatus = 'Archived';
      else if (brand.active) brandStatus = 'Active';
      else if (brand.approved) brandStatus = 'Approved (Hidden Dropdown)';
    }

    const prevActive = row[IDX.mapActive];

    const newActive =
      !(itemStatus === 'Archived' || brandStatus === 'Archived');

    row[IDX.itemStatus] = itemStatus;
    row[IDX.brandStatus] = brandStatus;

    row[IDX.mapActive] = newActive;

    row[IDX.analytics] = true;

    if (prevActive !== newActive) {

      repaired++;

      row[IDX.notes] =
        `Mapping state updated due to entity status change`;
    } else {

      valid++;
    }
  }

  mapSh
    .getRange(2,1,mapData.length-1,mapHdr.length)
    .setValues(mapData.slice(1));

  console.log(
    `[${SCRIPT_NAME}] VALID=${valid} REPAIRED=${repaired}`
  );

}
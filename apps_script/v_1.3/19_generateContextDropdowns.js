function generateContextDropdowns() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const mapSheet = ss.getSheetByName("Mapping_Item_Brand_Product");
  const lookupBrands = ss.getSheetByName("Lookup_Brands");
  const lookupProducts = ss.getSheetByName("Lookup_Products");
  const stagingBrands = ss.getSheetByName("Staging_Lookup_Brands");
  const stagingProducts = ss.getSheetByName("Staging_Lookup_Products");
  const ctxSheet = ss.getSheetByName("Context_Dropdowns");

  const now = new Date();

  // Reset table
  ctxSheet.clearContents();

  const header = [
    "Context_Type",
    "Item_ID_Machine",
    "Brand_ID_Machine",
    "Product_ID_Machine",
    "Display_Value",
    "Priority",
    "Source",
    "Created_At"
  ];

  ctxSheet.getRange(1,1,1,header.length).setValues([header]);

  const rows = [];
  const seen = new Set();

  // -------------------------
  // Mapping: Item-Brand
  // -------------------------

  const mapData = mapSheet.getDataRange().getValues();
  const mapHeader = mapData[0];

  const idxItem = mapHeader.indexOf("Item_ID_Machine");
  const idxBrand = mapHeader.indexOf("Brand_ID_Machine");
  const idxProduct = mapHeader.indexOf("Product_ID_Machine");
  const idxBrandName = mapHeader.indexOf("Brand_Name_Canonical");
  const idxProductName = mapHeader.indexOf("Product_Name_Canonical");
  const idxActive = mapHeader.indexOf("Is_Mapping_Active");

  for (let i = 1; i < mapData.length; i++) {

    const r = mapData[i];

    if (idxActive < 0 || r[idxActive] !== true) continue;

    const item = r[idxItem];
    const brand = r[idxBrand];
    const product = r[idxProduct];

    const brandName = r[idxBrandName];
    const productName = r[idxProductName];

    const keyBrand = "IB_" + item + "_" + brand;

    if (!seen.has(keyBrand)) {

      rows.push([
        "ITEM_BRAND",
        item,
        brand,
        "",
        brandName,
        1,
        "Mapping_Item_Brand_Product",
        now
      ]);

      seen.add(keyBrand);

    }

    if (product && productName) {

      const keyProduct = "IBP_" + item + "_" + brand + "_" + product;

      if (!seen.has(keyProduct)) {

        rows.push([
          "ITEM_BRAND_PRODUCT",
          item,
          brand,
          product,
          productName,
          1,
          "Mapping_Item_Brand_Product",
          now
        ]);

        seen.add(keyProduct);

      }

    }

  }

  // -------------------------
  // Lookup Brands
  // -------------------------

  const lbData = lookupBrands.getDataRange().getValues();
  const lbHeader = lbData[0];

  const lbBrandID = lbHeader.indexOf("Brand_ID_Machine");
  const lbBrandName = lbHeader.indexOf("Brand_Name");
  const lbActive = lbHeader.indexOf("Is_Active");

  for (let i = 1; i < lbData.length; i++) {

    const r = lbData[i];

    if (lbActive < 0 || r[lbActive] !== true) continue;

    rows.push([
      "ITEM_BRAND",
      "",
      r[lbBrandID],
      "",
      r[lbBrandName],
      2,
      "Lookup_Brands",
      now
    ]);

  }

  // -------------------------
  // Lookup Products
  // -------------------------

  const lpData = lookupProducts.getDataRange().getValues();
  const lpHeader = lpData[0];

  const lpProductID = lpHeader.indexOf("Product_ID_Machine");
  const lpProductName = lpHeader.indexOf("Product_Name");
  const lpActive = lpHeader.indexOf("Is_Active");

  for (let i = 1; i < lpData.length; i++) {

    const r = lpData[i];

    if (lpActive < 0 || r[lpActive] !== true) continue;

    rows.push([
      "ITEM_BRAND_PRODUCT",
      "",
      "",
      r[lpProductID],
      r[lpProductName],
      2,
      "Lookup_Products",
      now
    ]);

  }

  // -------------------------
  // Staging Brands
  // -------------------------

  const sbData = stagingBrands.getDataRange().getValues();
  const sbHeader = sbData[0];

  const sbBrandID = sbHeader.indexOf("Staging_Brand_ID_Machine");
  const sbBrandName = sbHeader.indexOf("Brand_Name_Entered");
  const sbActive = sbHeader.indexOf("Is_Active");
  const sbPromoted = sbHeader.indexOf("Is_Lookup_Promoted");

  for (let i = 1; i < sbData.length; i++) {

    const r = sbData[i];

    if (sbActive < 0 || r[sbActive] !== true) continue;
    if (sbPromoted >= 0 && r[sbPromoted] === true) continue;

    rows.push([
      "ITEM_BRAND",
      "",
      r[sbBrandID],
      "",
      r[sbBrandName],
      3,
      "Staging_Brands",
      now
    ]);

  }

  // -------------------------
  // Staging Products
  // -------------------------

  const spData = stagingProducts.getDataRange().getValues();
  const spHeader = spData[0];

  const spProductID = spHeader.indexOf("Staging_Product_ID_Machine");
  const spProductName = spHeader.indexOf("Product_Name_Entered");
  const spActive = spHeader.indexOf("Is_Active");
  const spPromoted = spHeader.indexOf("Is_Lookup_Promoted");

  for (let i = 1; i < spData.length; i++) {

    const r = spData[i];

    if (spActive < 0 || r[spActive] !== true) continue;
    if (spPromoted >= 0 && r[spPromoted] === true) continue;

    rows.push([
      "ITEM_BRAND_PRODUCT",
      "",
      "",
      r[spProductID],
      r[spProductName],
      3,
      "Staging_Products",
      now
    ]);

  }

  // -------------------------
  // SORT ROWS (Priority → Display_Value)
  // -------------------------

  rows.sort((a, b) => {

    const priorityDiff = a[5] - b[5];

    if (priorityDiff !== 0) return priorityDiff;

    return String(a[4]).localeCompare(String(b[4]));

  });

  if (rows.length > 0) {

    ctxSheet.getRange(2,1,rows.length,rows[0].length).setValues(rows);

  }

}

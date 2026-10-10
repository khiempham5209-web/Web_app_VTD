// FULL-TABS V1: install with 00/10/11 helpers; follow README, not legacy installation comments below.
/*************************************************
SYNC HANG LOI -> HANG LOI CHUYEN SANG GHTK

Paste file nay vao cung Apps Script project hien tai.

Luồng chạy:
- DS SKU co SKU moi -> refresh dropdown "Mã vật tư" ben tab "Hàng lỗi".
- User chon "Mã vật tư" ben tab "Hàng lỗi" -> tu fill Barcode + Tên sản phẩm.
- User dien Ngày nhận, Số lượng, Tình trạng GHTK nhận, GHTK bàn giao
  -> sync sang file dich tab "Hàng lỗi (chuyển sang GHTK)".
- Khi user cap nhat GHTK bàn giao sau, target se update theo du lieu hien tai.
- Trigger chi xu ly 2 tab "Hàng lỗi" va "DS SKU"; tab khac se bo qua.

Sau khi paste code:
1. Chạy installSyncHangLoiGHTKTrigger() một lần và cấp quyền.
2. Chạy refreshHangLoiMaterialDropdown() nếu muốn cập nhật dropdown ngay.
3. Chạy syncAllHangLoiToGHTK() nếu muốn đẩy dữ liệu đang có sẵn.
*************************************************/

const DEFECT_SYNC_CONFIG = {
  sourceSpreadsheetId: "1V39AVE2JfEtMPYqnG1-J75fKPDXU37fCnJ9Na3UXOco",
  targetSpreadsheetId: "1ZLqeo_djwMJofqBCD3Ilk6lIkUmVsEfK7T5Oj9D6wmI",
  sourceSheetName: "Hàng lỗi",
  skuSheetName: "DS SKU",
  targetSheetName: "Hàng lỗi (chuyển sang GHTK)",
  sourceHeaderRow: 1,
  sourceDataStartRow: 2,
  skuHeaderRow: 1,
  skuDataStartRow: 2,
  targetHeaderRow: 1,
  targetDataStartRow: 2,
  targetColumnCount: 8
};

function installSyncHangLoiGHTKTrigger() { return pnInstallTriggers(); /* superseded */

  const fn = "onEditSyncHangLoiGHTK";
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === fn)
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger(fn)
    .forSpreadsheet(DEFECT_SYNC_CONFIG.sourceSpreadsheetId)
    .onEdit()
    .create();

  refreshHangLoiMaterialDropdown();
  defectSyncFillProductInfoAll_();
  return "Installed trigger: " + fn;
}

function onEditSyncHangLoiGHTK(e) {
  try {
    if (!e || !e.range) return;
    return defectSyncWithLock_(function() {
      const sh = e.range.getSheet();
      const sheetName = sh.getName();

      if (sheetName === DEFECT_SYNC_CONFIG.skuSheetName) {
        if (e.range.getRow() < DEFECT_SYNC_CONFIG.skuHeaderRow) return;
        refreshHangLoiMaterialDropdown();
        defectSyncFillProductInfoAll_();
        return syncAllHangLoiToGHTK();
      }

      if (sheetName === DEFECT_SYNC_CONFIG.sourceSheetName) {
        if (e.range.getRow() < DEFECT_SYNC_CONFIG.sourceDataStartRow) return;
        defectSyncFillProductInfoForEditedRows_(e.range);
        return syncAllHangLoiToGHTK();
      }
    });
  } catch (err) {
    console.error(err && err.stack || err);
  }
}

function refreshHangLoiMaterialDropdown() {
  const source = defectSyncSourceSheet_();
  const sku = defectSyncSkuSheet_();
  const sourceCol = defectSyncHeaderMap_(defectSyncHeaders_(source, DEFECT_SYNC_CONFIG.sourceHeaderRow));
  const skuCol = defectSyncHeaderMap_(defectSyncHeaders_(sku, DEFECT_SYNC_CONFIG.skuHeaderRow));
  const sourceMaterialCol = defectSyncFirstCol_(sourceCol, ["ma vat tu", "mã vật tư"]);
  const skuMaterialCol = defectSyncFirstCol_(skuCol, ["ma vat tu", "mã vật tư"]);
  if (!sourceMaterialCol) throw new Error("Khong thay cot Ma vat tu trong " + DEFECT_SYNC_CONFIG.sourceSheetName);
  if (!skuMaterialCol) throw new Error("Khong thay cot Ma vat tu trong " + DEFECT_SYNC_CONFIG.skuSheetName);

  const lastSkuRow = Math.max(sku.getLastRow(), DEFECT_SYNC_CONFIG.skuDataStartRow);
  const skuRange = sku.getRange(DEFECT_SYNC_CONFIG.skuDataStartRow, skuMaterialCol, lastSkuRow - DEFECT_SYNC_CONFIG.skuDataStartRow + 1, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(skuRange, true)
    .setAllowInvalid(false)
    .build();

  const maxRows = Math.max(source.getMaxRows() - DEFECT_SYNC_CONFIG.sourceDataStartRow + 1, 1);
  source.getRange(DEFECT_SYNC_CONFIG.sourceDataStartRow, sourceMaterialCol, maxRows, 1).setDataValidation(rule);
  return {ok: true};
}

function syncAllHangLoiToGHTK() {
  const source = defectSyncSourceSheet_();
  const target = defectSyncTargetSheet_();
  defectSyncEnsureTargetHeaders_(target);

  const sourceCol = defectSyncHeaderMap_(defectSyncHeaders_(source, DEFECT_SYNC_CONFIG.sourceHeaderRow));
  const lastRow = source.getLastRow();
  const output = [];

  if (lastRow >= DEFECT_SYNC_CONFIG.sourceDataStartRow) {
    const rows = source.getRange(
      DEFECT_SYNC_CONFIG.sourceDataStartRow,
      1,
      lastRow - DEFECT_SYNC_CONFIG.sourceDataStartRow + 1,
      source.getLastColumn()
    ).getDisplayValues();

    rows.forEach(row => {
      const materialCode = defectSyncClean_(defectSyncPick_(row, sourceCol, ["ma vat tu", "mã vật tư"]));
      const receivedDate = defectSyncClean_(defectSyncPick_(row, sourceCol, ["ngay nhan", "ngày nhận"]));
      const quantity = defectSyncClean_(defectSyncPick_(row, sourceCol, ["so luong", "số lượng"]));
      const statusReceived = defectSyncClean_(defectSyncPick_(row, sourceCol, ["tinh trang ghtk nhan", "tình trạng ghtk nhận"]));
      const ghtkHandover = defectSyncClean_(defectSyncPick_(row, sourceCol, ["ghtk ban giao", "ghtk bàn giao"]));
      const note = defectSyncClean_(defectSyncPick_(row, sourceCol, ["ghi chu", "ghi chú", "note"]));

      if (!materialCode && !receivedDate && !quantity && !statusReceived && !ghtkHandover && !note) return;
      if (!materialCode) return;

      output.push([
        receivedDate,
        materialCode,
        defectSyncPick_(row, sourceCol, ["barcode"]),
        defectSyncPick_(row, sourceCol, ["ten san pham", "tên sản phẩm"]),
        quantity,
        statusReceived,
        ghtkHandover,
        note
      ]);
    });
  }

  pnRequireReady_();
  const pnOldLast = target.getLastRow();
  if (output.length) {
    target.getRange(DEFECT_SYNC_CONFIG.targetDataStartRow, 1, output.length, DEFECT_SYNC_CONFIG.targetColumnCount).setValues(output);
  }
  if(pnOldLast > output.length + 1) target.getRange(output.length + 2, 1, pnOldLast - output.length - 1, DEFECT_SYNC_CONFIG.targetColumnCount).clearContent();
  defectSyncFormatTarget_(target, output.length);
  return {ok: true, count: output.length};
}

function defectSyncFillProductInfoForEditedRows_(range) {
  const source = defectSyncSourceSheet_();
  const sourceCol = defectSyncHeaderMap_(defectSyncHeaders_(source, DEFECT_SYNC_CONFIG.sourceHeaderRow));
  const materialCol = defectSyncFirstCol_(sourceCol, ["ma vat tu", "mã vật tư"]);
  if (!materialCol) throw new Error("Khong thay cot Ma vat tu trong " + DEFECT_SYNC_CONFIG.sourceSheetName);

  const editFirstCol = range.getColumn();
  const editLastCol = editFirstCol + range.getNumColumns() - 1;
  if (materialCol < editFirstCol || materialCol > editLastCol) return;

  const startRow = Math.max(range.getRow(), DEFECT_SYNC_CONFIG.sourceDataStartRow);
  const endRow = range.getRow() + range.getNumRows() - 1;
  if (endRow < DEFECT_SYNC_CONFIG.sourceDataStartRow) return;
  defectSyncFillProductInfoRows_(startRow, endRow - startRow + 1);
}

function defectSyncFillProductInfoAll_() {
  const source = defectSyncSourceSheet_();
  const lastRow = source.getLastRow();
  if (lastRow < DEFECT_SYNC_CONFIG.sourceDataStartRow) return {ok: true, updated: 0};
  return defectSyncFillProductInfoRows_(DEFECT_SYNC_CONFIG.sourceDataStartRow, lastRow - DEFECT_SYNC_CONFIG.sourceDataStartRow + 1);
}

function defectSyncFillProductInfoRows_(startRow, numRows) {
  if (!numRows) return {ok: true, updated: 0};

  const source = defectSyncSourceSheet_();
  const sourceCol = defectSyncHeaderMap_(defectSyncHeaders_(source, DEFECT_SYNC_CONFIG.sourceHeaderRow));
  const materialCol = defectSyncFirstCol_(sourceCol, ["ma vat tu", "mã vật tư"]);
  const barcodeCol = defectSyncFirstCol_(sourceCol, ["barcode"]);
  const productNameCol = defectSyncFirstCol_(sourceCol, ["ten san pham", "tên sản phẩm"]);
  if (!materialCol) throw new Error("Khong thay cot Ma vat tu trong " + DEFECT_SYNC_CONFIG.sourceSheetName);
  if (!barcodeCol) throw new Error("Khong thay cot Barcode trong " + DEFECT_SYNC_CONFIG.sourceSheetName);
  if (!productNameCol) throw new Error("Khong thay cot Ten san pham trong " + DEFECT_SYNC_CONFIG.sourceSheetName);

  const skuMap = defectSyncSkuMap_();
  const materialValues = source.getRange(startRow, materialCol, numRows, 1).getDisplayValues();
  const barcodeValues = source.getRange(startRow, barcodeCol, numRows, 1).getDisplayValues();
  const productNameValues = source.getRange(startRow, productNameCol, numRows, 1).getDisplayValues();
  let updated = 0;

  materialValues.forEach((row, index) => {
    const materialCode = defectSyncClean_(row[0]);
    const sku = skuMap[defectSyncNorm_(materialCode)] || {};
    const nextBarcode = sku.barcode || "";
    const nextProductName = sku.productName || "";

    if (barcodeValues[index][0] !== nextBarcode) {
      barcodeValues[index][0] = nextBarcode;
      updated++;
    }
    if (productNameValues[index][0] !== nextProductName) {
      productNameValues[index][0] = nextProductName;
      updated++;
    }
  });

  if (updated) {
    source.getRange(startRow, barcodeCol, numRows, 1).setValues(barcodeValues);
    source.getRange(startRow, productNameCol, numRows, 1).setValues(productNameValues);
  }
  return {ok: true, updated};
}

function defectSyncSkuMap_() {
  const sku = defectSyncSkuSheet_();
  const col = defectSyncHeaderMap_(defectSyncHeaders_(sku, DEFECT_SYNC_CONFIG.skuHeaderRow));
  const materialCol = defectSyncFirstCol_(col, ["ma vat tu", "mã vật tư"]);
  const barcodeCol = defectSyncFirstCol_(col, ["barcode"]);
  const productNameCol = defectSyncFirstCol_(col, ["ten vat tu", "tên vật tư", "ten san pham", "tên sản phẩm"]);
  if (!materialCol) throw new Error("Khong thay cot Ma vat tu trong " + DEFECT_SYNC_CONFIG.skuSheetName);
  if (!barcodeCol) throw new Error("Khong thay cot Barcode trong " + DEFECT_SYNC_CONFIG.skuSheetName);
  if (!productNameCol) throw new Error("Khong thay cot Ten vat tu trong " + DEFECT_SYNC_CONFIG.skuSheetName);

  const lastRow = sku.getLastRow();
  if (lastRow < DEFECT_SYNC_CONFIG.skuDataStartRow) return {};

  const values = sku.getRange(
    DEFECT_SYNC_CONFIG.skuDataStartRow,
    1,
    lastRow - DEFECT_SYNC_CONFIG.skuDataStartRow + 1,
    sku.getLastColumn()
  ).getDisplayValues();
  const out = {};
  values.forEach(row => {
    const materialCode = defectSyncClean_(row[materialCol - 1]);
    if (!materialCode) return;
    out[defectSyncNorm_(materialCode)] = {
      barcode: defectSyncClean_(row[barcodeCol - 1]),
      productName: defectSyncClean_(row[productNameCol - 1])
    };
  });
  return out;
}

function defectSyncEnsureTargetHeaders_(target) {
  const headers = ["Ngày nhận", "Mã vật tư", "Barcode", "Tên sản phẩm", "Số lượng", "Tình trạng GHTK nhận", "GHTK bàn giao", "Ghi chú"];
  target.getRange(DEFECT_SYNC_CONFIG.targetHeaderRow, 1, 1, headers.length).setValues([headers]);
}

function defectSyncClearTargetData_(target) {
  const lastRow = target.getLastRow();
  if (lastRow >= DEFECT_SYNC_CONFIG.targetDataStartRow) {
    target.getRange(
      DEFECT_SYNC_CONFIG.targetDataStartRow,
      1,
      lastRow - DEFECT_SYNC_CONFIG.targetDataStartRow + 1,
      DEFECT_SYNC_CONFIG.targetColumnCount
    ).clearContent();
  }
}

function defectSyncFormatTarget_(target, dataCount) {
  const rows = Math.max(Number(dataCount || 0) + 1, 1);
  target.getRange(DEFECT_SYNC_CONFIG.targetHeaderRow, 1, rows, DEFECT_SYNC_CONFIG.targetColumnCount)
    .setFontFamily("Times New Roman")
    .setFontSize(12)
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true);
  target.getRange(DEFECT_SYNC_CONFIG.targetHeaderRow, 1, 1, DEFECT_SYNC_CONFIG.targetColumnCount)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  if (dataCount) {
    target.getRange(DEFECT_SYNC_CONFIG.targetDataStartRow, 1, dataCount, DEFECT_SYNC_CONFIG.targetColumnCount).setFontWeight("normal");
  }
}

function defectSyncSourceSheet_() {
  const ss = pnOpenById_(DEFECT_SYNC_CONFIG.sourceSpreadsheetId);
  const sh = ss.getSheetByName(DEFECT_SYNC_CONFIG.sourceSheetName);
  if (!sh) throw new Error("Khong thay tab " + DEFECT_SYNC_CONFIG.sourceSheetName);
  return sh;
}

function defectSyncSkuSheet_() {
  const ss = pnOpenById_(DEFECT_SYNC_CONFIG.sourceSpreadsheetId);
  const sh = ss.getSheetByName(DEFECT_SYNC_CONFIG.skuSheetName);
  if (!sh) throw new Error("Khong thay tab " + DEFECT_SYNC_CONFIG.skuSheetName);
  return sh;
}

function defectSyncTargetSheet_() {
  const ss = pnOpenById_(DEFECT_SYNC_CONFIG.targetSpreadsheetId);
  const sh = ss.getSheetByName(DEFECT_SYNC_CONFIG.targetSheetName);
  if (!sh) throw new Error("Khong thay tab " + DEFECT_SYNC_CONFIG.targetSheetName);
  return sh;
}

function defectSyncHeaders_(sh, headerRow) {
  return sh.getRange(Number(headerRow || 1), 1, 1, sh.getLastColumn()).getDisplayValues()[0];
}

function defectSyncHeaderMap_(headers) {
  const out = {};
  headers.forEach((h, i) => out[defectSyncNorm_(h)] = i + 1);
  return out;
}

function defectSyncFirstCol_(col, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const c = col[defectSyncNorm_(aliases[i])];
    if (c) return c;
  }
  return 0;
}

function defectSyncPick_(row, col, aliases) {
  const c = defectSyncFirstCol_(col, aliases);
  return c ? row[c - 1] : "";
}

function defectSyncClean_(value) {
  return String(value == null ? "" : value).trim();
}

function defectSyncWithLock_(fn) {
  return pnWithLock_(fn);
}

function defectSyncNorm_(value) {
  return defectSyncClean_(value).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d").replace(/\s+/g, " ");
}

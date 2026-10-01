// FULL-TABS V1: install with 00/10/11 helpers; follow README, not legacy installation comments below.
/*************************************************
SYNC TT NHAP <-> DS SKU  (DS SKU la NGUON CHUAN: chi them ma moi / dien o trong, khong ghi de)

Thay the ban cu. Khac biet chinh:
- KHONG BAO GIO ghi de Barcode / Tên / Quy cách DA CO trong "DS SKU".
- Ma moi o "TT Nhập" chua co trong "DS SKU" -> tu them sang DS SKU (addNewSkuToDSSKU = true).
  Ma da co nhung o Barcode/Ten/Quy cach dang TRONG -> dien vao tu TT Nhập.
- Nhap/sua o "TT Nhập" -> Barcode VA Tên bánh tu lay theo "DS SKU" (key: Mã Vật Tư = Ordered item).
- Bo ham skuSyncEnsureTargetHeaders_ (ban cu ghi lai dong tieu de DS SKU moi lan chay).
- Sua Barcode / Tên Vật Tư o "DS SKU" -> tu cap nhat moi dong cung ma o "TT Nhập".
- "Ngày" trong o TT Nhập van tu dien ngay hien tai khi nhap tay.

Luu y: onEdit chi chay khi NGUOI sua tay. Du lieu do script khac ghi vao
(vd. OCR) phai tu dien - code OCR da tu lam viec nay.

Sau khi paste code:
1. Chay installSyncTTNhapDSSKUTrigger() mot lan va cap quyen.
2. Muon chuan hoa lai toan bo TT Nhập theo DS SKU: chay syncAllDSSKUToTTNhap().
*************************************************/

const SKU_SYNC_CONFIG = {
  spreadsheetId: "1V39AVE2JfEtMPYqnG1-J75fKPDXU37fCnJ9Na3UXOco",
  sourceSheetName: "TT Nhập",
  targetSheetName: "DS SKU",
  sourceHeaderRow: 2,
  sourceDataStartRow: 3,
  targetHeaderRow: 1,
  targetDataStartRow: 2,
  addNewSkuToDSSKU: true
};

const SKU_ALIAS = {
  date: ["ngay", "ngày"],
  material: ["ordered item", "ma vat tu", "mã vật tư"],
  barcode: ["barcode"],
  name: ["ten banh", "tên bánh", "ten vat tu", "tên vật tư"],
  quyCach: ["quy cach/thung", "quy cách/thùng", "quy cach", "quy cách"]
};

function installSyncTTNhapDSSKUTrigger() { return pnInstallTriggers(); /* superseded */

  const fn = "onEditSyncTTNhapDSSKU";
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === fn)
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger(fn)
    .forSpreadsheet(SKU_SYNC_CONFIG.spreadsheetId)
    .onEdit()
    .create();

  return "Installed trigger: " + fn;
}

function onEditSyncTTNhapDSSKU(e) {
  try {
    if (!e || !e.range) return;
    return skuSyncWithLock_(function() {
      const sh = e.range.getSheet();
      const sheetName = sh.getName();

      // Sua o TT Nhập -> dien Ngay, them ma moi sang DS SKU, lay Barcode/Ten tu DS SKU cho cac dong vua sua
      if (sheetName === SKU_SYNC_CONFIG.sourceSheetName) {
        if (e.range.getRow() + e.range.getNumRows() - 1 < SKU_SYNC_CONFIG.sourceDataStartRow) return;
        const sourceCol = skuSyncHeaderMap_(skuSyncHeaders_(sh, SKU_SYNC_CONFIG.sourceHeaderRow));
        const watchedColumns = [
          skuSyncFirstCol_(sourceCol, SKU_ALIAS.material),
          skuSyncFirstCol_(sourceCol, SKU_ALIAS.barcode),
          skuSyncFirstCol_(sourceCol, SKU_ALIAS.name),
          skuSyncFirstCol_(sourceCol, SKU_ALIAS.quyCach)
        ].filter(Boolean);
        if (!skuSyncRangeIntersectsAnyColumn_(e.range, watchedColumns)) {
          return {ok: true, skipped: true, reason: "edit column not watched"};
        }
        const startRow = Math.max(e.range.getRow(), SKU_SYNC_CONFIG.sourceDataStartRow);
        const endRow = e.range.getRow() + e.range.getNumRows() - 1;
        return syncTTNhapRows_(startRow, endRow - startRow + 1);
      }

      // Sua Barcode/Ten o DS SKU -> cap nhat moi dong cung ma o TT Nhập
      if (sheetName === SKU_SYNC_CONFIG.targetSheetName) {
        if (e.range.getRow() + e.range.getNumRows() - 1 < SKU_SYNC_CONFIG.targetDataStartRow) return;
        const col = skuSyncHeaderMap_(skuSyncHeaders_(sh, SKU_SYNC_CONFIG.targetHeaderRow));
        const watched = [
          skuSyncFirstCol_(col, SKU_ALIAS.material),
          skuSyncFirstCol_(col, SKU_ALIAS.barcode),
          skuSyncFirstCol_(col, SKU_ALIAS.name)
        ].filter(Boolean);
        if (!skuSyncRangeIntersectsAnyColumn_(e.range, watched)) return;
        const startRow = Math.max(e.range.getRow(), SKU_SYNC_CONFIG.targetDataStartRow);
        const endRow = e.range.getRow() + e.range.getNumRows() - 1;
        const map = skuSyncMasterMap_(startRow, endRow - startRow + 1);
        return skuSyncApplyMasterToTTNhapRows_(map, SKU_SYNC_CONFIG.sourceDataStartRow, skuSyncSourceRowCount_());
      }
    });
  } catch (err) {
    console.error(err && err.stack || err);
  }
}

// Chuan hoa lai TOAN BO TT Nhập theo DS SKU (Barcode + Ten)
function syncAllDSSKUToTTNhap() {
  return skuSyncWithLock_(function() {
    return skuSyncApplyMasterToTTNhapRows_(skuSyncMasterMap_(), SKU_SYNC_CONFIG.sourceDataStartRow, skuSyncSourceRowCount_());
  });
}

// Giu ten cu de khong loi neu co cho khac dang goi
function syncAllDSSKUBarcodeToTTNhap() { return syncAllDSSKUToTTNhap(); }
function syncAllTTNhapToDSSKU() {
  return skuSyncWithLock_(function() {
    return syncTTNhapRows_(SKU_SYNC_CONFIG.sourceDataStartRow, skuSyncSourceRowCount_());
  });
}

function syncTTNhapRows_(startRow, numRows) {
  if (!numRows) return {ok: true, added: 0, updated: 0};
  const source = skuSyncSourceSheet_();
  const sourceCol = skuSyncHeaderMap_(skuSyncHeaders_(source, SKU_SYNC_CONFIG.sourceHeaderRow));
  skuSyncFillMissingSourceDates_(source, sourceCol, startRow, numRows);

  const dsResult = SKU_SYNC_CONFIG.addNewSkuToDSSKU
    ? skuSyncAddNewToDSSKU_(source, sourceCol, startRow, numRows) : {added: 0, filled: 0};

  const result = skuSyncApplyMasterToTTNhapRows_(skuSyncMasterMap_(), startRow, numRows);
  result.added = dsResult.added;
  result.filledDSSKU = dsResult.filled;
  return result;
}

// DS SKU: ma -> {barcode, name}. Chi DOC.
function skuSyncMasterMap_(startRow, numRows) {
  const target = skuSyncTargetSheet_();
  const col = skuSyncHeaderMap_(skuSyncHeaders_(target, SKU_SYNC_CONFIG.targetHeaderRow));
  const materialCol = skuSyncFirstCol_(col, SKU_ALIAS.material);
  const barcodeCol = skuSyncFirstCol_(col, SKU_ALIAS.barcode);
  const nameCol = skuSyncFirstCol_(col, SKU_ALIAS.name);
  if (!materialCol) throw new Error("Khong thay cot Ma Vat Tu trong " + SKU_SYNC_CONFIG.targetSheetName);

  startRow = startRow || SKU_SYNC_CONFIG.targetDataStartRow;
  if (numRows == null) numRows = target.getLastRow() - startRow + 1;
  if (numRows <= 0) return {};
  const values = target.getRange(startRow, 1, numRows, target.getLastColumn()).getDisplayValues();
  const out = {};
  values.forEach(row => {
    const code = skuSyncClean_(row[materialCol - 1]);
    if (!code || skuSyncIsHeaderLikeMaterial_(code)) return;
    out[skuSyncNorm_(code)] = {
      barcode: barcodeCol ? skuSyncClean_(row[barcodeCol - 1]) : "",
      name: nameCol ? skuSyncClean_(row[nameCol - 1]) : ""
    };
  });
  return out;
}

// Ghi Barcode + Ten tu DS SKU vao TT Nhập (chi ghi o cot Barcode/Ten cua TT Nhập).
function skuSyncApplyMasterToTTNhapRows_(master, startRow, numRows) {
  if (!numRows || numRows <= 0) return {ok: true, updated: 0, missing: []};
  const source = skuSyncSourceSheet_();
  const col = skuSyncHeaderMap_(skuSyncHeaders_(source, SKU_SYNC_CONFIG.sourceHeaderRow));
  const materialCol = skuSyncFirstCol_(col, SKU_ALIAS.material);
  const barcodeCol = skuSyncFirstCol_(col, SKU_ALIAS.barcode);
  const nameCol = skuSyncFirstCol_(col, SKU_ALIAS.name);
  if (!materialCol) throw new Error("Khong thay cot Ordered item trong " + SKU_SYNC_CONFIG.sourceSheetName);

  const materials = source.getRange(startRow, materialCol, numRows, 1).getDisplayValues();
  const barcodes = barcodeCol ? source.getRange(startRow, barcodeCol, numRows, 1).getDisplayValues() : null;
  const names = nameCol ? source.getRange(startRow, nameCol, numRows, 1).getDisplayValues() : null;
  let updated = 0;
  const missing = [];

  materials.forEach((row, i) => {
    const code = skuSyncClean_(row[0]);
    if (!code || skuSyncIsHeaderLikeMaterial_(code)) return;
    const info = master[skuSyncNorm_(code)];
    if (!info) { missing.push(code); return; }
    if (barcodes && info.barcode && barcodes[i][0] !== info.barcode) { barcodes[i][0] = info.barcode; updated++; }
    if (names && info.name && names[i][0] !== info.name) { names[i][0] = info.name; updated++; }
  });

  if (updated) {
    if (barcodes) source.getRange(startRow, barcodeCol, numRows, 1).setValues(barcodes);
    if (names) source.getRange(startRow, nameCol, numRows, 1).setValues(names);
  }
  return {ok: true, updated: updated, missing: missing};
}

/*
  Dua ma tu TT Nhập sang DS SKU:
  - Ma CHUA CO trong DS SKU -> them dong moi (Ma, Barcode, Ten, Quy cach lay tu TT Nhập).
  - Ma DA CO -> chi dien vao o dang TRONG cua DS SKU (vd. vua them ma, sau moi go ten).
    O da co gia tri thi KHONG BAO GIO ghi de -> ten/barcode chuan khong bi doi.
*/
function skuSyncAddNewToDSSKU_(source, sourceCol, startRow, numRows) {
  const target = skuSyncTargetSheet_();
  const tCol = skuSyncHeaderMap_(skuSyncHeaders_(target, SKU_SYNC_CONFIG.targetHeaderRow));
  const cMat = skuSyncFirstCol_(tCol, SKU_ALIAS.material);
  if (!cMat) throw new Error("Khong thay cot Ma Vat Tu trong " + SKU_SYNC_CONFIG.targetSheetName);
  const fields = [
    ["barcode", skuSyncFirstCol_(tCol, SKU_ALIAS.barcode)],
    ["name", skuSyncFirstCol_(tCol, SKU_ALIAS.name)],
    ["quyCach", skuSyncFirstCol_(tCol, SKU_ALIAS.quyCach)]
  ].filter(f => f[1]);

  const tStart = SKU_SYNC_CONFIG.targetDataStartRow;
  const tCount = Math.max(target.getLastRow() - tStart + 1, 0);
  const lastCol = Math.max(target.getLastColumn(), 5);
  const tValues = tCount ? target.getRange(tStart, 1, tCount, lastCol).getDisplayValues() : [];
  const index = {};
  tValues.forEach((r, i) => {
    const code = skuSyncClean_(r[cMat - 1]);
    if (code && !skuSyncIsHeaderLikeMaterial_(code)) index[skuSyncNorm_(code)] = {values: r, row: tStart + i};
  });

  const rows = source.getRange(startRow, 1, numRows, source.getLastColumn()).getDisplayValues();
  const appendRows = [];
  const cellUpdates = [];
  rows.forEach(row => {
    const code = skuSyncClean_(skuSyncPick_(row, sourceCol, SKU_ALIAS.material));
    if (!code || skuSyncIsHeaderLikeMaterial_(code)) return;
    const key = skuSyncNorm_(code);
    const vals = {
      barcode: skuSyncClean_(skuSyncPick_(row, sourceCol, SKU_ALIAS.barcode)),
      name: skuSyncClean_(skuSyncPick_(row, sourceCol, SKU_ALIAS.name)),
      quyCach: skuSyncClean_(skuSyncPick_(row, sourceCol, SKU_ALIAS.quyCach))
    };
    let entry = index[key];
    if (!entry) {
      const out = new Array(lastCol).fill("");
      out[cMat - 1] = code;
      appendRows.push(out);
      entry = index[key] = {values: out, row: 0};
    }
    fields.forEach(f => {
      const c = f[1];
      if (!vals[f[0]] || skuSyncClean_(entry.values[c - 1])) return;
      entry.values[c - 1] = vals[f[0]];
      if (entry.row) cellUpdates.push([entry.row, c, vals[f[0]]]);
    });
  });

  cellUpdates.forEach(u => target.getRange(u[0], u[1]).setValue(u[2]));
  if (appendRows.length) {
    const start = Math.max(target.getLastRow() + 1, tStart);
    target.getRange(start, 1, appendRows.length, lastCol).setValues(appendRows);
    skuSyncFormatTargetRows_(target, start, appendRows.length);
    skuSyncRenumberSTT_(target);
  }
  return {added: appendRows.length, filled: cellUpdates.length};
}

function skuSyncFillMissingSourceDates_(source, sourceCol, startRow, numRows) {
  const dateCol = skuSyncFirstCol_(sourceCol, SKU_ALIAS.date);
  const materialCol = skuSyncFirstCol_(sourceCol, SKU_ALIAS.material);
  if (!dateCol || !materialCol || !numRows) return {ok: true, updated: 0};

  const dateRange = source.getRange(startRow, dateCol, numRows, 1);
  const dateValues = dateRange.getDisplayValues();
  const materialValues = source.getRange(startRow, materialCol, numRows, 1).getDisplayValues();
  const today = skuSyncTodayText_();
  let updated = 0;

  materialValues.forEach((row, index) => {
    const materialCode = skuSyncClean_(row[0]);
    if (!materialCode || skuSyncIsHeaderLikeMaterial_(materialCode) || skuSyncClean_(dateValues[index][0])) return;
    dateValues[index][0] = today;
    updated++;
  });

  if (updated) dateRange.setValues(dateValues);
  return {ok: true, updated};
}

function skuSyncSourceRowCount_() {
  return Math.max(skuSyncSourceSheet_().getLastRow() - SKU_SYNC_CONFIG.sourceDataStartRow + 1, 0);
}

function skuSyncRenumberSTT_(target) {
  const lastRow = target.getLastRow();
  if (lastRow < SKU_SYNC_CONFIG.targetDataStartRow) return;
  const col = skuSyncHeaderMap_(skuSyncHeaders_(target, SKU_SYNC_CONFIG.targetHeaderRow));
  const sttCol = skuSyncFirstCol_(col, ["stt"]);
  const materialCol = skuSyncFirstCol_(col, SKU_ALIAS.material);
  if (!sttCol || !materialCol) return;

  const materials = target.getRange(SKU_SYNC_CONFIG.targetDataStartRow, materialCol, lastRow - SKU_SYNC_CONFIG.targetDataStartRow + 1, 1).getDisplayValues();
  let count = 0;
  const stt = materials.map(row => {
    if (!skuSyncClean_(row[0])) return [""];
    count++;
    return [count];
  });
  target.getRange(SKU_SYNC_CONFIG.targetDataStartRow, sttCol, stt.length, 1).setValues(stt);
}

function skuSyncFormatTargetRows_(target, startRow, numRows) {
  target.getRange(startRow, 1, numRows, 5)
    .setFontFamily("Times New Roman")
    .setFontSize(12)
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true);
}

function skuSyncWithLock_(fn) {
  return pnWithLock_(fn);
}

function skuSyncRangeIntersectsAnyColumn_(range, columns) {
  const firstCol = range.getColumn();
  const lastCol = firstCol + range.getNumColumns() - 1;
  return columns.some(col => col >= firstCol && col <= lastCol);
}

function skuSyncSourceSheet_() {
  const ss = SpreadsheetApp.openById(SKU_SYNC_CONFIG.spreadsheetId);
  const sh = ss.getSheetByName(SKU_SYNC_CONFIG.sourceSheetName);
  if (!sh) throw new Error("Khong thay tab " + SKU_SYNC_CONFIG.sourceSheetName);
  return sh;
}

function skuSyncTargetSheet_() {
  const ss = SpreadsheetApp.openById(SKU_SYNC_CONFIG.spreadsheetId);
  const sh = ss.getSheetByName(SKU_SYNC_CONFIG.targetSheetName);
  if (!sh) throw new Error("Khong thay tab " + SKU_SYNC_CONFIG.targetSheetName);
  return sh;
}

function skuSyncHeaders_(sh, headerRow) {
  return sh.getRange(Number(headerRow || 1), 1, 1, sh.getLastColumn()).getDisplayValues()[0];
}

function skuSyncHeaderMap_(headers) {
  const out = {};
  headers.forEach((h, i) => out[skuSyncNorm_(h)] = i + 1);
  return out;
}

function skuSyncFirstCol_(col, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const c = col[skuSyncNorm_(aliases[i])];
    if (c) return c;
  }
  return 0;
}

function skuSyncPick_(row, col, aliases) {
  const c = skuSyncFirstCol_(col, aliases);
  return c ? row[c - 1] : "";
}

function skuSyncSetOutputByAliases_(output, col, aliases, value) {
  const c = skuSyncFirstCol_(col, aliases);
  if (c) output[c - 1] = value;
}

function skuSyncClean_(value) {
  return String(value == null ? "" : value).trim();
}

function skuSyncTodayText_() {
  return Utilities.formatDate(new Date(), "Asia/Saigon", "dd/MM/yyyy");
}

function skuSyncIsHeaderLikeMaterial_(value) {
  const key = skuSyncNorm_(value);
  return key === "ordered item" || key === "ma vat tu" || key === "mã vật tư";
}

function skuSyncNorm_(value) {
  return skuSyncClean_(value).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/đ/g, "d").replace(/\s+/g, " ");
}

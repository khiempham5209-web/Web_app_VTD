// FULL-TABS V1: install with 00/10/11 helpers; follow README, not legacy installation comments below.
/*************************************************
SYNC SU VU -> SU VU BANH XEP

Nguon: spreadsheet [FFPL]_Pham Nguyen, tab "Sự vụ" gid 529840243
Dich: spreadsheet VHFF_Pham Nguyen, tab "Sự vụ bánh xẹp" gid 318091884

Sau khi paste code:
1. Chay installSyncSuVuBanhXepTrigger() mot lan va cap quyen.
2. Chay syncAllSuVuBanhXep() hoac syncAllSuvuBanhXep() de sync tay.
*************************************************/

const SUVU_BANH_XEP_SYNC_CONFIG = {
  sourceSpreadsheetId: "1V39AVE2JfEtMPYqnG1-J75fKPDXU37fCnJ9Na3UXOco",
  sourceSheetId: 529840243,
  orderSheetId: 872196600,
  orderSheetName: "File đơn_full",
  targetSpreadsheetId: "1ZLqeo_djwMJofqBCD3Ilk6lIkUmVsEfK7T5Oj9D6wmI",
  targetSheetId: 318091884,
  headerRow: 1,
  dataStartRow: 2,
  outputHeaders: [
    "Mã đơn",
    "Tên khách hàng",
    "Kho đích",
    "QL Lastmile",
    "Sản phẩm bị xẹp",
    "Số lượng",
    "Đã hoàn FF?",
    "Ghi chú",
    "Nguyên nhân",
    "Hình ảnh"
  ]
};

function installSyncSuVuBanhXepTrigger() { return pnInstallTriggers(); /* superseded */

  const handlers = ["onEditSyncSuVuBanhXep", "onEditSyncSuvuBanhXep"];
  ScriptApp.getProjectTriggers()
    .filter(t => handlers.indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("onEditSyncSuVuBanhXep")
    .forSpreadsheet(SUVU_BANH_XEP_SYNC_CONFIG.sourceSpreadsheetId)
    .onEdit()
    .create();

  return "Installed trigger: onEditSyncSuVuBanhXep";
}

function installSyncSuvuBanhXepTrigger() { return pnInstallTriggers(); /* superseded */

  return installSyncSuVuBanhXepTrigger();
}

function onEditSyncSuVuBanhXep(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getSheetId() !== SUVU_BANH_XEP_SYNC_CONFIG.sourceSheetId) return;

    const editLastRow = e.range.getRow() + e.range.getNumRows() - 1;
    if (
      e.range.getRow() < SUVU_BANH_XEP_SYNC_CONFIG.dataStartRow &&
      editLastRow < SUVU_BANH_XEP_SYNC_CONFIG.dataStartRow
    ) {
      return;
    }

    return suVuBanhXepSyncWithLock_(syncAllSuVuBanhXep_);
  } catch (err) {
    console.error(err && err.stack || err);
  }
}

function onEditSyncSuvuBanhXep(e) {
  return onEditSyncSuVuBanhXep(e);
}

function syncAllSuVuBanhXep() {
  const result = suVuBanhXepSyncWithLock_(syncAllSuVuBanhXep_);
  console.log(JSON.stringify(result));
  return result;
}

function syncAllSuvuBanhXep() {
  return syncAllSuVuBanhXep();
}

function syncAllSuVuBanhXep_() {
  const source = suVuBanhXepSourceSheet_();
  const target = suVuBanhXepTargetSheet_();
  const sourceLastRow = source.getLastRow();
  const sourceLastCol = source.getLastColumn();
  const targetWidth = SUVU_BANH_XEP_SYNC_CONFIG.outputHeaders.length;

  if (sourceLastCol < 1) return {ok: true, count: 0};

  suVuBanhXepEnsureSize_(
    target,
    Math.max(sourceLastRow, SUVU_BANH_XEP_SYNC_CONFIG.dataStartRow),
    targetWidth
  );

  const sourceHeaders = source
    .getRange(SUVU_BANH_XEP_SYNC_CONFIG.headerRow, 1, 1, sourceLastCol)
    .getDisplayValues()[0];

  const sourceMap = suVuBanhXepHeaderMap_(sourceHeaders);
  const orderInfoMap = suVuBanhXepOrderInfoMap_();
  const outputHeaders = SUVU_BANH_XEP_SYNC_CONFIG.outputHeaders;
  const outputRows = [];

  if (sourceLastRow >= SUVU_BANH_XEP_SYNC_CONFIG.dataStartRow) {
    const sourceRows = source
      .getRange(
        SUVU_BANH_XEP_SYNC_CONFIG.dataStartRow,
        1,
        sourceLastRow - SUVU_BANH_XEP_SYNC_CONFIG.dataStartRow + 1,
        sourceLastCol
      )
      .getDisplayValues();

    sourceRows.forEach(row => {
      if (!row.some(v => suVuBanhXepClean_(v))) return;
      outputRows.push(suVuBanhXepBuildOutputRow_(row, sourceMap, orderInfoMap));
    });
  }

  pnRequireReady_();
  const pnOldLast = target.getLastRow();
  target
    .getRange(SUVU_BANH_XEP_SYNC_CONFIG.headerRow, 1, 1, outputHeaders.length)
    .setValues([outputHeaders]);

  if (outputRows.length) {
    target
      .getRange(SUVU_BANH_XEP_SYNC_CONFIG.dataStartRow, 1, outputRows.length, outputHeaders.length)
      .setValues(outputRows);
  }

  if(pnOldLast > outputRows.length + 1) target.getRange(outputRows.length + 2, 1, pnOldLast - outputRows.length - 1, targetWidth).clearContent();
  suVuBanhXepFormatTarget_(target, outputRows.length, outputHeaders.length);

  return {
    ok: true,
    count: outputRows.length,
    sourceSheetId: SUVU_BANH_XEP_SYNC_CONFIG.sourceSheetId,
    targetSheetId: SUVU_BANH_XEP_SYNC_CONFIG.targetSheetId
  };
}

function suVuBanhXepBuildOutputRow_(sourceRow, sourceMap, orderInfoMap) {
  const maDon = suVuBanhXepPick_(sourceRow, sourceMap, [
    "ma don",
    "mã đơn",
    "mã đơn",
    "mã đơn ghtk",
    "mã đơn ghtk"
  ]);

  const orderInfo = orderInfoMap[suVuBanhXepNorm_(maDon)] || {};

  return [
    maDon,
    suVuBanhXepPick_(sourceRow, sourceMap, ["ten khach hang", "tên khách hàng", "khach hang", "khách hàng"]),
    suVuBanhXepPick_(sourceRow, sourceMap, ["kho dich", "kho đích"]) || orderInfo.khoDich || "",
    orderInfo.qlLastmile || "",
    suVuBanhXepPick_(sourceRow, sourceMap, ["san pham bi xep", "sản phẩm bị xẹp", "sản phẩm bị xẹp"]),
    suVuBanhXepPick_(sourceRow, sourceMap, ["so luong", "số lượng", "số lượng"]),
    suVuBanhXepPick_(sourceRow, sourceMap, ["check hoan", "check hoàn", "da hoan ff", "đã hoàn ff", "đã hoàn ff?"]),
    suVuBanhXepPick_(sourceRow, sourceMap, ["ghi chu", "ghi chú", "note"], 8),
    suVuBanhXepPick_(sourceRow, sourceMap, ["nguyen nhan", "nguyên nhân"], 9),
    suVuBanhXepPick_(sourceRow, sourceMap, ["hinh anh", "hình ảnh", "link anh", "link ảnh", "link hinh anh", "link hình ảnh", "ảnh"])
  ];
}

function suVuBanhXepOrderInfoMap_() {
  const sh = suVuBanhXepOrderSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return {};

  const lastCol = sh.getLastColumn();
  const headers = pnHeaders_(sh);
  const col = suVuBanhXepHeaderMap_(headers);
  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  const out = {};

  values.forEach(row => {
    const maDon = suVuBanhXepPick_(row, col, ["ma don", "mã đơn", "mã đơn", "ma don ghtk", "mã đơn ghtk", "mã đơn ghtk"]);
    if (!maDon) return;

    out[suVuBanhXepNorm_(maDon)] = {
      khoDich: suVuBanhXepPick_(row, col, ["kho dich", "kho đích"]),
      qlLastmile: suVuBanhXepPick_(row, col, ["ql lastmile", "quan ly", "quản lý"])
    };
  });

  return out;
}

function suVuBanhXepClearTarget_(target, width) {
  const rows = Math.max(target.getLastRow(), 1);
  target.getRange(1, 1, rows, width).clearContent();
}

function suVuBanhXepFormatTarget_(target, dataCount, width) {
  const rows = Math.max(Number(dataCount || 0) + 1, 1);

  target
    .getRange(1, 1, rows, width)
    .setFontFamily("Times New Roman")
    .setFontSize(12)
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true);

  target
    .getRange(1, 1, 1, width)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
}

function suVuBanhXepEnsureSize_(sheet, minRows, minCols) {
  const maxRows = sheet.getMaxRows();
  if (maxRows < minRows) sheet.insertRowsAfter(maxRows, minRows - maxRows);

  const maxCols = sheet.getMaxColumns();
  if (maxCols < minCols) sheet.insertColumnsAfter(maxCols, minCols - maxCols);
}

function suVuBanhXepSourceSheet_() {
  return suVuBanhXepSheetById_(
    SUVU_BANH_XEP_SYNC_CONFIG.sourceSpreadsheetId,
    SUVU_BANH_XEP_SYNC_CONFIG.sourceSheetId,
    "source"
  );
}

function suVuBanhXepTargetSheet_() {
  return suVuBanhXepSheetById_(
    SUVU_BANH_XEP_SYNC_CONFIG.targetSpreadsheetId,
    SUVU_BANH_XEP_SYNC_CONFIG.targetSheetId,
    "target"
  );
}

function suVuBanhXepOrderSheet_() {
  return suVuBanhXepSheetByIdOrName_(
    SUVU_BANH_XEP_SYNC_CONFIG.sourceSpreadsheetId,
    SUVU_BANH_XEP_SYNC_CONFIG.orderSheetId,
    SUVU_BANH_XEP_SYNC_CONFIG.orderSheetName,
    "order"
  );
}

function suVuBanhXepSheetById_(spreadsheetId, sheetId, label) {
  const ss = pnOpenById_(spreadsheetId);
  const sheets = ss.getSheets();

  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === sheetId) return sheets[i];
  }

  throw new Error("Khong thay tab " + label + " gid=" + sheetId);
}

function suVuBanhXepSheetByIdOrName_(spreadsheetId, sheetId, sheetName, label) {
  const ss = pnOpenById_(spreadsheetId);
  const sheets = ss.getSheets();

  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === sheetId) return sheets[i];
  }

  const byName = ss.getSheetByName(sheetName);
  if (byName) return byName;

  throw new Error("Khong thay tab " + label + " gid=" + sheetId + " name=" + sheetName);
}

function suVuBanhXepHeaderMap_(headers) {
  const out = {};

  headers.forEach((header, index) => {
    const key = suVuBanhXepNorm_(header);
    if (key && !out[key]) out[key] = index + 1;
  });

  return out;
}

function suVuBanhXepPick_(row, col, aliases, fallbackIndex) {
  for (let i = 0; i < aliases.length; i++) {
    const c = col[suVuBanhXepNorm_(aliases[i])];
    if (c) return suVuBanhXepClean_(row[c - 1]);
  }

  if (fallbackIndex && fallbackIndex > 0) {
    return suVuBanhXepClean_(row[fallbackIndex - 1]);
  }

  return "";
}

function suVuBanhXepSyncWithLock_(fn) { return pnWithLock_(fn); }

function suVuBanhXepClean_(value) {
  return String(value == null ? "" : value).trim();
}

function suVuBanhXepNorm_(value) {
  return suVuBanhXepClean_(value)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/\s+/g, " ");
}
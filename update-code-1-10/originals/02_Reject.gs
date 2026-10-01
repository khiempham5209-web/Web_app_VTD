/*************************************************
SYNC CHUNG TU KHONG DAT YC

Paste file nay vao cung Apps Script project voi Code API.gs.

Luồng chạy:
- User tick "Đã tạo sv" = TRUE ở tab nguồn "Chứng từ_FF" thì onEdit sync.
- Nếu app/API lưu dữ liệu trực tiếp, Code API.gs gọi thẳng
  syncOneChungTuKhongDatYCBySourceRow_(rowNumber) sau khi lưu.
- Script append/update sang file đích tab "Chứng từ không đạt YC".
- Key update: "Mã đơn GHTK".
- Lookup "Tỉnh đích", "Kho đích", "Vị trí giao", "Quản lý" từ tab nguồn
  "File đơn" theo "Mã đơn".
- Hai cột "Ecom FF -> Kho đích" và "Ecom Kho đích -> FF" ở file đích
  để user tự điền, script không ghi đè.

Sau khi paste code:
1. Chạy installSyncChungTuKhongDatYCTrigger() một lần và cấp quyền.
2. Nếu muốn đẩy các dòng đang TRUE sẵn, chạy syncAllChungTuKhongDatYC().
*************************************************/

const REJECT_SYNC_CONFIG = {
  sourceSpreadsheetId: "1V39AVE2JfEtMPYqnG1-J75fKPDXU37fCnJ9Na3UXOco",
  sourceSheetName: "Chứng từ_FF",
  orderSheetName: "File đơn",
  targetSpreadsheetId: "1ZLqeo_djwMJofqBCD3Ilk6lIkUmVsEfK7T5Oj9D6wmI",
  targetSheetName: "Chứng từ không đạt YC"
};

function installSyncChungTuKhongDatYCTrigger() {
  const fn = "onEditSyncChungTuKhongDatYC";
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === fn)
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger(fn)
    .forSpreadsheet(REJECT_SYNC_CONFIG.sourceSpreadsheetId)
    .onEdit()
    .create();

  return "Installed trigger: " + fn;
}

function onEditSyncChungTuKhongDatYC(e) {
  try {
    if (!e || !e.range) return;
    return rejectSyncWithLock_(function() {
      const sh = e.range.getSheet();
      if (sh.getName() !== REJECT_SYNC_CONFIG.sourceSheetName) return;
      if (e.range.getRow() < 2) return;

      const col = rejectSyncHeaderMap_(rejectSyncHeaders_(sh));
      const createCol = rejectSyncFirstCol_(col, ["da tao sv", "đã tạo sv"]);
      if (!createCol) throw new Error("Khong thay cot Da tao sv.");

      const startRow = Math.max(e.range.getRow(), 2);
      const endRow = e.range.getRow() + e.range.getNumRows() - 1;
      const checked = sh.getRange(startRow, createCol, endRow - startRow + 1, 1).getValues();
      const rows = [];
      checked.forEach((row, i) => {
        if (rejectSyncIsTrue_(row[0])) rows.push(startRow + i);
      });
      if (!rows.length) return {ok: true, count: 0};
      return syncChungTuKhongDatYCRows_(rows);
    });
  } catch (err) {
    console.error(err && err.stack || err);
  }
}

function syncAllChungTuKhongDatYC() {
  return rejectSyncWithLock_(function() {
    const sh = rejectSyncSourceSheet_();
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return {ok: true, count: 0};

    const col = rejectSyncHeaderMap_(rejectSyncHeaders_(sh));
    const createCol = rejectSyncFirstCol_(col, ["da tao sv", "đã tạo sv"]);
    if (!createCol) throw new Error("Khong thay cot Da tao sv.");

    const rows = [];
    const checked = sh.getRange(2, createCol, lastRow - 1, 1).getValues();
    checked.forEach((row, i) => {
      if (rejectSyncIsTrue_(row[0])) rows.push(i + 2);
    });
    if (!rows.length) return {ok: true, count: 0};
    return syncChungTuKhongDatYCRows_(rows);
  });
}

function syncChungTuKhongDatYCRows_(rowNumbers) {
  const sources = rejectSyncSourceRecordsByRows_(rowNumbers);
  const orderMap = rejectSyncOrderInfoMap_();
  const target = rejectSyncTargetSheet_();
  rejectSyncEnsureTargetHeaders_(target);
  const targetCol = rejectSyncHeaderMap_(rejectSyncHeaders_(target));
  const targetRowByMaDon = rejectSyncTargetRowMapByMaDon_(target, targetCol);
  const outputWidth = target.getLastColumn();
  const appendRows = [];
  let updated = 0;

  sources.forEach(source => {
    if (!source.maDonGhtk || !rejectSyncIsTrue_(source.daTaoSv)) return;
    const order = orderMap[rejectSyncNorm_(source.maDonGhtk)] || {};
    const targetRow = targetRowByMaDon[rejectSyncNorm_(source.maDonGhtk)] || 0;
    const output = new Array(outputWidth).fill("");
    if (targetRow) rejectSyncPreserveManualTargetValues_(target, targetCol, targetRow, output);

    rejectSyncBuildOutput_(output, targetCol, source, order);
    if (targetRow) {
      target.getRange(targetRow, 1, 1, output.length).setValues([output]);
      updated++;
    } else {
      appendRows.push(output);
      targetRowByMaDon[rejectSyncNorm_(source.maDonGhtk)] = true;
    }
  });

  if (appendRows.length) {
    const startRow = Math.max(target.getLastRow() + 1, 2);
    target.getRange(startRow, 1, appendRows.length, outputWidth).setValues(appendRows);
  }
  return {ok: true, count: updated + appendRows.length, updated, added: appendRows.length};
}

function rejectSyncBuildOutput_(output, targetCol, source, order) {
  rejectSyncSetOutputByAliases_(output, targetCol, ["ngay tao don", "ngay len don"], source.ngayLenDon);
  rejectSyncSetOutputByAliases_(output, targetCol, ["ma don ghtk", "ma don"], source.maDonGhtk);
  rejectSyncSetOutputByAliases_(output, targetCol, ["tinh dich"], order.tinhDich || "");
  rejectSyncSetOutputByAliases_(output, targetCol, ["kho dich"], order.khoDich || "");
  rejectSyncSetOutputByAliases_(output, targetCol, ["ql lastmile", "quan ly"], order.quanLy || "");
  rejectSyncSetOutputByAliases_(output, targetCol, ["vi tri giao"], order.viTriGiao || "");
  rejectSyncSetOutputByAliases_(output, targetCol, ["ten khach hang", "khach hang"], source.customer);
  rejectSyncSetOutputByAliases_(output, targetCol, ["dia chi"], source.address);
  rejectSyncSetOutputByAliases_(output, targetCol, ["so don hang", "od"], source.orderNo);
  rejectSyncSetOutputByAliases_(output, targetCol, ["ma po", "po"], source.po);
  rejectSyncSetOutputByAliases_(output, targetCol, ["tinh trang chung tu"], source.ghiChuChungTu);
  rejectSyncSetOutputByAliases_(output, targetCol, ["ff nhan chung tu"], rejectSyncMapXacThucHoaDon_(source.xacThuc));
  rejectSyncSetOutputByAliases_(output, targetCol, ["ff note check"], rejectSyncMapTrangThai_(source.trangThai));
}

function rejectSyncSourceRecordsByRows_(rowNumbers) {
  const sh = rejectSyncSourceSheet_();
  const col = rejectSyncHeaderMap_(rejectSyncHeaders_(sh));
  const minRow = Math.min.apply(null, rowNumbers);
  const maxRow = Math.max.apply(null, rowNumbers);
  const values = sh.getRange(minRow, 1, maxRow - minRow + 1, sh.getLastColumn()).getDisplayValues();
  const rowSet = {};
  rowNumbers.forEach(row => rowSet[row] = true);
  const out = [];
  values.forEach((row, index) => {
    const rowNumber = minRow + index;
    if (!rowSet[rowNumber]) return;
    out.push(rejectSyncRecordFromRowValues_(row, col));
  });
  return out;
}

function rejectSyncTargetRowMapByMaDon_(sh, col) {
  const out = {};
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return out;
  const maDonCol = rejectSyncFirstCol_(col, ["ma don ghtk", "ma don"]);
  if (!maDonCol) return out;
  const values = sh.getRange(2, maDonCol, lastRow - 1, 1).getDisplayValues();
  values.forEach((row, i) => {
    const maDon = rejectSyncClean_(row[0]);
    if (maDon) out[rejectSyncNorm_(maDon)] = i + 2;
  });
  return out;
}

function rejectSyncWithLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/*
// Old row-by-row versions kept disabled after batch rewrite.
function onEditSyncChungTuKhongDatYC_OLD(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== REJECT_SYNC_CONFIG.sourceSheetName) return;
    if (e.range.getRow() < 2) return;

    const col = rejectSyncHeaderMap_(rejectSyncHeaders_(sh));
    const createCol = rejectSyncFirstCol_(col, ["da tao sv", "đã tạo sv"]);
    if (!createCol) throw new Error("Khong thay cot Da tao sv.");

    for (let r = e.range.getRow(); r < e.range.getRow() + e.range.getNumRows(); r++) {
      if (r >= 2 && rejectSyncIsTrue_(sh.getRange(r, createCol).getValue())) {
        syncOneChungTuKhongDatYCBySourceRow_(r);
      }
    }
  } catch (err) {
    console.error(err && err.stack || err);
  }
}

function syncAllChungTuKhongDatYC_OLD() {
  const sh = rejectSyncSourceSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return {ok: true, count: 0};

  const col = rejectSyncHeaderMap_(rejectSyncHeaders_(sh));
  const createCol = rejectSyncFirstCol_(col, ["da tao sv", "đã tạo sv"]);
  if (!createCol) throw new Error("Khong thay cot Da tao sv.");

  let count = 0;
  const checked = sh.getRange(2, createCol, lastRow - 1, 1).getValues();
  checked.forEach((row, i) => {
    if (rejectSyncIsTrue_(row[0])) {
      syncOneChungTuKhongDatYCBySourceRow_(i + 2);
      count++;
    }
  });
  return {ok: true, count};
}
*/

function syncOneChungTuKhongDatYCBySourceRow_(rowNumber) {
  const source = rejectSyncSourceRecordByRow_(rowNumber);
  if (!source.maDonGhtk || !rejectSyncIsTrue_(source.daTaoSv)) return null;

  const orderMap = rejectSyncOrderInfoMap_();
  const order = orderMap[rejectSyncNorm_(source.maDonGhtk)] || {};
  const target = rejectSyncTargetSheet_();
  rejectSyncEnsureTargetHeaders_(target);
  const targetRow = rejectSyncFindTargetRowByMaDon_(target, source.maDonGhtk) || Math.max(target.getLastRow() + 1, 2);
  const targetCol = rejectSyncHeaderMap_(rejectSyncHeaders_(target));
  const output = new Array(target.getLastColumn()).fill("");
  rejectSyncPreserveManualTargetValues_(target, targetCol, targetRow, output);

  rejectSyncSetOutputByAliases_(output, targetCol, ["ngay tao don", "ngay len don"], source.ngayLenDon);
  rejectSyncSetOutputByAliases_(output, targetCol, ["ma don ghtk", "ma don"], source.maDonGhtk);
  rejectSyncSetOutputByAliases_(output, targetCol, ["tinh dich"], order.tinhDich || "");
  rejectSyncSetOutputByAliases_(output, targetCol, ["kho dich"], order.khoDich || "");
  rejectSyncSetOutputByAliases_(output, targetCol, ["ql lastmile", "quan ly"], order.quanLy || "");
  rejectSyncSetOutputByAliases_(output, targetCol, ["vi tri giao"], order.viTriGiao || "");
  rejectSyncSetOutputByAliases_(output, targetCol, ["ten khach hang", "khach hang"], source.customer);
  rejectSyncSetOutputByAliases_(output, targetCol, ["dia chi"], source.address);
  rejectSyncSetOutputByAliases_(output, targetCol, ["so don hang", "od"], source.orderNo);
  rejectSyncSetOutputByAliases_(output, targetCol, ["ma po", "po"], source.po);
  rejectSyncSetOutputByAliases_(output, targetCol, ["tinh trang chung tu"], source.ghiChuChungTu);
  rejectSyncSetOutputByAliases_(output, targetCol, ["ff nhan chung tu"], rejectSyncMapXacThucHoaDon_(source.xacThuc));
  rejectSyncSetOutputByAliases_(output, targetCol, ["ff note check"], rejectSyncMapTrangThai_(source.trangThai));

  target.getRange(targetRow, 1, 1, output.length).setValues([output]);
  return {ok: true, targetRow, maDonGhtk: source.maDonGhtk};
}

function rejectSyncSourceRecordByRow_(rowNumber) {
  const sh = rejectSyncSourceSheet_();
  const col = rejectSyncHeaderMap_(rejectSyncHeaders_(sh));
  const row = sh.getRange(rowNumber, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  return rejectSyncRecordFromRowValues_(row, col);
}

function rejectSyncRecordFromRowValues_(row, col) {
  return {
    ngayLenDon: rejectSyncClean_(rejectSyncPick_(row, col, ["ngay len don"])),
    maDonGhtk: rejectSyncClean_(rejectSyncPick_(row, col, ["ma don ghtk", "ma don"])),
    customer: rejectSyncClean_(rejectSyncPick_(row, col, ["khach hang"])),
    po: rejectSyncClean_(rejectSyncPick_(row, col, ["ma po", "po"])),
    orderNo: rejectSyncClean_(rejectSyncPick_(row, col, ["so don hang", "od"])),
    address: rejectSyncClean_(rejectSyncPick_(row, col, ["dia chi nhan hang"])),
    xacThuc: rejectSyncClean_(rejectSyncPick_(row, col, ["xac thuc hoa don"])),
    daTaoSv: rejectSyncPick_(row, col, ["da tao sv", "đã tạo sv"]),
    ghiChuChungTu: rejectSyncClean_(rejectSyncPick_(row, col, ["ghi chu chung tu"])),
    trangThai: rejectSyncClean_(rejectSyncPick_(row, col, ["trang thai"]))
  };
}

function rejectSyncOrderInfoMap_() {
  const ss = SpreadsheetApp.openById(REJECT_SYNC_CONFIG.sourceSpreadsheetId);
  const sh = ss.getSheetByName(REJECT_SYNC_CONFIG.orderSheetName);
  if (!sh) throw new Error("Khong thay tab " + REJECT_SYNC_CONFIG.orderSheetName);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return {};

  const col = rejectSyncHeaderMap_(rejectSyncHeaders_(sh));
  const values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getDisplayValues();
  const out = {};
  values.forEach(row => {
    const maDon = rejectSyncClean_(rejectSyncPick_(row, col, ["ma don", "ma don ghtk"]));
    if (!maDon) return;
    out[rejectSyncNorm_(maDon)] = {
      tinhDich: rejectSyncClean_(rejectSyncPick_(row, col, ["tinh thanh dia chi giao"])),
      khoDich: rejectSyncClean_(rejectSyncPick_(row, col, ["kho dich"])),
      viTriGiao: rejectSyncClean_(rejectSyncPick_(row, col, ["vi tri giao"])),
      quanLy: rejectSyncClean_(rejectSyncPick_(row, col, ["quan ly", "ql lastmile"]))
    };
  });
  return out;
}

function rejectSyncSourceSheet_() {
  const ss = SpreadsheetApp.openById(REJECT_SYNC_CONFIG.sourceSpreadsheetId);
  const sh = ss.getSheetByName(REJECT_SYNC_CONFIG.sourceSheetName);
  if (!sh) throw new Error("Khong thay tab " + REJECT_SYNC_CONFIG.sourceSheetName);
  return sh;
}

function rejectSyncTargetSheet_() {
  const ss = SpreadsheetApp.openById(REJECT_SYNC_CONFIG.targetSpreadsheetId);
  const sh = ss.getSheetByName(REJECT_SYNC_CONFIG.targetSheetName);
  if (!sh) throw new Error("Khong thay tab " + REJECT_SYNC_CONFIG.targetSheetName);
  return sh;
}

function rejectSyncEnsureTargetHeaders_(sh) {
  const headers = [
    "Ngày tạo đơn", "Mã đơn GHTK", "Tỉnh đích", "Kho đích", "QL Lastmile",
    "Vị trí giao", "Kho hiện tại", "Tên khách hàng", "Địa chỉ", "Số đơn hàng",
    "Mã PO", "Tình trạng chứng từ", "FF nhận chứng từ", "Ecom FF -> Kho đích",
    "Ecom Kho đích -> FF", "FF note check"
  ];
  const current = sh.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
  const emptyOrMismatch = current.every(v => !rejectSyncClean_(v)) || rejectSyncNorm_(current[0]) !== rejectSyncNorm_(headers[0]);
  if (emptyOrMismatch) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function rejectSyncPreserveManualTargetValues_(target, targetCol, targetRow, output) {
  if (targetRow > target.getLastRow()) return;
  const manualAliases = [
    ["ecom ff -> kho dich"],
    ["ecom kho dich -> ff"]
  ];
  const current = target.getRange(targetRow, 1, 1, target.getLastColumn()).getDisplayValues()[0];
  manualAliases.forEach(aliases => {
    const col = rejectSyncFirstCol_(targetCol, aliases);
    if (col) output[col - 1] = current[col - 1];
  });
}

function rejectSyncFindTargetRowByMaDon_(sh, maDonGhtk) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  const col = rejectSyncHeaderMap_(rejectSyncHeaders_(sh));
  const maDonCol = rejectSyncFirstCol_(col, ["ma don ghtk", "ma don"]);
  if (!maDonCol) return 0;
  const q = rejectSyncNorm_(maDonGhtk);
  const values = sh.getRange(2, maDonCol, lastRow - 1, 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (rejectSyncNorm_(values[i][0]) === q) return i + 2;
  }
  return 0;
}

function rejectSyncMapXacThucHoaDon_(value) {
  const key = rejectSyncNorm_(value);
  if (key === "da nhan chung tu") return "Đã nhận";
  if (key === "chua nhan chung tu") return "Chưa nhận";
  return rejectSyncClean_(value);
}

function rejectSyncMapTrangThai_(value) {
  const key = rejectSyncNorm_(value);
  if (key === "sai/thieu chung tu") return "Chưa đạt yêu cầu";
  if (key === "binh thuong") return "Đạt yêu cầu";
  return rejectSyncClean_(value);
}

function rejectSyncHeaders_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
}

function rejectSyncHeaderMap_(headers) {
  const out = {};
  headers.forEach((h, i) => out[rejectSyncNorm_(h)] = i + 1);
  return out;
}

function rejectSyncFirstCol_(col, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const c = col[rejectSyncNorm_(aliases[i])];
    if (c) return c;
  }
  return 0;
}

function rejectSyncPick_(row, col, aliases) {
  const c = rejectSyncFirstCol_(col, aliases);
  return c ? row[c - 1] : "";
}

function rejectSyncSetOutputByAliases_(output, col, aliases, value) {
  const c = rejectSyncFirstCol_(col, aliases);
  if (c) output[c - 1] = value;
}

function rejectSyncIsTrue_(value) {
  if (value === true) return true;
  const key = rejectSyncNorm_(value);
  return key === "true" || key === "1" || key === "yes" || key === "x";
}

function rejectSyncClean_(value) {
  return String(value == null ? "" : value).trim();
}

function rejectSyncNorm_(value) {
  return rejectSyncClean_(value).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d").replace(/\s+/g, " ");
}

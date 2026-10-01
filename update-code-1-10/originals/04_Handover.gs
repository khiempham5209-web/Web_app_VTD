/*************************************************
SYNC BAN GIAO CHUNG TU

Paste file nay vao Apps Script project rieng/cung project hien tai deu duoc.

Luồng chạy:
- Tab nguồn: "Chứng từ_FF".
- Tab đích: "Bàn giao chứng từ".
- A1 của tab đích là dropdown ngày, tự quét từ cột "Ngày bàn giao CT"
  bên tab nguồn.
- Khi A1 thay đổi, script lọc toàn bộ dòng nguồn có "Ngày bàn giao CT"
  trùng ngày A1 rồi ghi sang A3:H.
- Khi mở file, script refresh lại báo cáo theo A1 hiện tại để dòng chữ ký
  lấy ngày hôm nay.
- G1 = tổng số dòng lọc được.

Sau khi paste code:
1. Chạy installSyncBanGiaoChungTuTrigger() một lần và cấp quyền.
2. Chạy refreshBanGiaoChungTuDropdown() nếu muốn cập nhật dropdown ngay.
*************************************************/

const HANDOVER_SYNC_CONFIG = {
  spreadsheetId: "1V39AVE2JfEtMPYqnG1-J75fKPDXU37fCnJ9Na3UXOco",
  sourceSheetName: "Chứng từ_FF",
  targetSheetName: "Bàn giao chứng từ",
  timezone: "Asia/Saigon"
};

function installSyncBanGiaoChungTuTrigger() {
  const editFn = "onEditSyncBanGiaoChungTu";
  const openFn = "onOpenSyncBanGiaoChungTu";
  ScriptApp.getProjectTriggers()
    .filter(t => [editFn, openFn].indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger(editFn)
    .forSpreadsheet(HANDOVER_SYNC_CONFIG.spreadsheetId)
    .onEdit()
    .create();

  ScriptApp.newTrigger(openFn)
    .forSpreadsheet(HANDOVER_SYNC_CONFIG.spreadsheetId)
    .onOpen()
    .create();

  refreshBanGiaoChungTuDropdown();
  return "Installed triggers: " + editFn + ", " + openFn;
}

function onOpenSyncBanGiaoChungTu() {
  try {
    refreshBanGiaoChungTuDropdown();
    return syncBanGiaoChungTuBySelectedDate();
  } catch (err) {
    console.error(err && err.stack || err);
  }
}

function onEditSyncBanGiaoChungTu(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();

    if (sh.getName() === HANDOVER_SYNC_CONFIG.targetSheetName && e.range.getA1Notation() === "A1") {
      refreshBanGiaoChungTuDropdown();
      return syncBanGiaoChungTuBySelectedDate();
    }

    if (sh.getName() === HANDOVER_SYNC_CONFIG.sourceSheetName) {
      const sourceCol = handoverSyncHeaderMap_(handoverSyncHeaders_(sh));
      const handoverDateCol = handoverSyncFirstCol_(sourceCol, ["ngay ban giao ct"]);
      const editedFirstCol = e.range.getColumn();
      const editedLastCol = editedFirstCol + e.range.getNumColumns() - 1;
      if (handoverDateCol >= editedFirstCol && handoverDateCol <= editedLastCol) {
        refreshBanGiaoChungTuDropdown();
      }
    }
  } catch (err) {
    console.error(err && err.stack || err);
  }
}

function refreshBanGiaoChungTuDropdown() {
  const target = handoverSyncTargetSheet_();
  const dates = handoverSyncUniqueHandoverDates_();
  const cell = target.getRange("A1");

  if (!dates.length) {
    cell.clearDataValidations();
    return {ok: true, count: 0};
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(dates, true)
    .setAllowInvalid(false)
    .build();
  cell.setDataValidation(rule);
  return {ok: true, count: dates.length};
}

function syncBanGiaoChungTuBySelectedDate() {
  const source = handoverSyncSourceSheet_();
  const target = handoverSyncTargetSheet_();
  handoverSyncEnsureTargetHeaders_(target);

  const selectedDate = handoverSyncDateKey_(target.getRange("A1").getDisplayValue());
  handoverSyncClearOutput_(target);
  if (!selectedDate) {
    target.getRange("G1").setValue(0);
    return {ok: true, count: 0};
  }

  const sourceCol = handoverSyncHeaderMap_(handoverSyncHeaders_(source));
  const lastRow = source.getLastRow();
  if (lastRow < 2) {
    target.getRange("G1").setValue(0);
    handoverSyncApplySummary_(target, [], sourceCol);
    return {ok: true, count: 0};
  }

  const values = source.getRange(2, 1, lastRow - 1, source.getLastColumn()).getDisplayValues();
  const filteredRows = values.filter(row => handoverSyncDateKey_(handoverSyncPick_(row, sourceCol, ["ngay ban giao ct"])) === selectedDate);
  const rows = filteredRows.map(row => [
      handoverSyncPick_(row, sourceCol, ["ngay len don"]),
      handoverSyncPick_(row, sourceCol, ["khach hang"]),
      handoverSyncPick_(row, sourceCol, ["dia chi nhan hang"]),
      handoverSyncPick_(row, sourceCol, ["ma po", "po"]),
      handoverSyncPick_(row, sourceCol, ["so don hang", "od"]),
      handoverSyncPick_(row, sourceCol, ["ma don ghtk", "ma don"]),
      handoverSyncPick_(row, sourceCol, ["ngay ban giao ct"]),
      handoverSyncPick_(row, sourceCol, ["note"])
    ]);

  if (rows.length) {
    target.getRange(3, 1, rows.length, 8)
      .setValues(rows)
      .setFontWeight("normal")
      .setHorizontalAlignment("left")
      .setVerticalAlignment("middle");
  }
  target.getRange("G1").setValue(rows.length);
  handoverSyncApplyTableFormat_(target, rows.length);
  handoverSyncApplyBorders_(target, rows.length);
  handoverSyncApplySummary_(target, filteredRows, sourceCol);
  handoverSyncApplySignature_(target, rows.length);
  return {ok: true, count: rows.length};
}

function handoverSyncUniqueHandoverDates_() {
  const source = handoverSyncSourceSheet_();
  const lastRow = source.getLastRow();
  if (lastRow < 2) return [];

  const col = handoverSyncHeaderMap_(handoverSyncHeaders_(source));
  const dateCol = handoverSyncFirstCol_(col, ["ngay ban giao ct"]);
  if (!dateCol) throw new Error("Khong thay cot Ngay ban giao CT trong " + HANDOVER_SYNC_CONFIG.sourceSheetName);

  const values = source.getRange(2, dateCol, lastRow - 1, 1).getDisplayValues();
  const seen = {};
  values.forEach(row => {
    const text = handoverSyncDateText_(row[0]);
    const key = handoverSyncDateKey_(text);
    if (key) seen[key] = text;
  });
  return Object.keys(seen).sort().map(key => seen[key]);
}

function handoverSyncEnsureTargetHeaders_(target) {
  const headers = [
    "Ngày lên đơn", "Tên khách hàng", "Địa chỉ nhận", "Mã PO",
    "Số đơn hàng", "Mã đơn GHTK", "Ngày bàn giao CT", "Ghi chú"
  ];
  target.getRange(2, 1, 1, headers.length).setValues([headers]);
}

function handoverSyncClearOutput_(target) {
  const lastRow = target.getLastRow();
  if (lastRow >= 3) {
    const oldRange = target.getRange(3, 1, lastRow - 2, 8);
    oldRange.breakApart();
    oldRange.clearContent();
    oldRange.clearFormat();
    oldRange.clearDataValidations();
    oldRange.setBorder(false, false, false, false, false, false);
  }
  target.getRange(2, 1, 1, 8).setBorder(false, false, false, false, false, false);
  handoverSyncClearSummary_(target);
}

function handoverSyncApplyBorders_(target, dataCount) {
  const rowCount = Math.max(Number(dataCount || 0) + 1, 1);
  target.getRange(2, 1, rowCount, 8).setBorder(true, true, true, true, true, true);
}

function handoverSyncApplyTableFormat_(target, dataCount) {
  const rowCount = Math.max(Number(dataCount || 0) + 1, 1);
  const tableRange = target.getRange(2, 1, rowCount, 8);
  tableRange
    .setFontFamily("Times New Roman")
    .setFontSize(13)
    .setVerticalAlignment("middle");

  target.getRange(2, 1, 1, 8)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setWrap(true);

  if (dataCount) {
    target.getRange(3, 1, dataCount, 8)
      .setFontWeight("normal")
      .setHorizontalAlignment("left")
      .setWrap(false);
    target.getRange(3, 3, dataCount, 1).setWrap(true);
    target.autoResizeRows(3, dataCount);
  }
}

function handoverSyncApplySummary_(target, filteredRows, sourceCol) {
  handoverSyncClearSummary_(target);
  const summary = handoverSyncBuildPivotSummary_(filteredRows, sourceCol);
  const startRow = 1;
  const startCol = 10; // J
  const colCount = Math.max(summary.headers.length, 2);
  const rowCount = Math.max(summary.rows.length + 2, 2);
  handoverSyncEnsureColumns_(target, startCol + colCount - 1);

  target.getRange(startRow + 1, startCol, 1, colCount).setValues([summary.headers]);
  if (summary.rows.length) {
    target.getRange(startRow + 2, startCol, summary.rows.length, colCount).setValues(summary.rows);
  }

  handoverSyncFormatSummaryTable_(target, startRow, startCol, rowCount, colCount);
}

function handoverSyncClearSummary_(target) {
  const maxRows = target.getMaxRows();
  const maxCols = Math.max(target.getMaxColumns() - 9, 1);
  const range = target.getRange(2, 10, maxRows - 1, maxCols);
  range.breakApart();
  range.clearContent();
  range.clearFormat();
  range.setBorder(false, false, false, false, false, false);
}

function handoverSyncBuildPivotSummary_(rows, sourceCol) {
  const storeSeen = {};
  const dateSeen = {};
  const counts = {};

  rows.forEach(row => {
    const store = handoverSyncClean_(handoverSyncPick_(row, sourceCol, ["loai sieu thi"]));
    const dateText = handoverSyncDateText_(handoverSyncPick_(row, sourceCol, ["ngay len don"]));
    const dateKey = handoverSyncDateKey_(dateText);
    if (!store || !dateKey) return;

    storeSeen[store] = true;
    dateSeen[dateKey] = dateText;
    counts[dateKey] = counts[dateKey] || {};
    counts[dateKey][store] = (counts[dateKey][store] || 0) + 1;
  });

  const stores = Object.keys(storeSeen).sort();
  const dateKeys = Object.keys(dateSeen).sort();
  const headers = ["Loại siêu thị"].concat(stores).concat(["Tổng"]);
  const storeTotals = {};
  let grandTotal = 0;
  const outputRows = dateKeys.map(dateKey => {
    let total = 0;
    const row = [dateSeen[dateKey]];
    stores.forEach(store => {
      const value = counts[dateKey][store] || "";
      if (value) total += value;
      if (value) storeTotals[store] = (storeTotals[store] || 0) + value;
      row.push(value);
    });
    row.push(total);
    grandTotal += total;
    return row;
  });

  if (outputRows.length) {
    const totalRow = ["Tổng"];
    stores.forEach(store => totalRow.push(storeTotals[store] || ""));
    totalRow.push(grandTotal);
    outputRows.push(totalRow);
  }

  return {headers, rows: outputRows};
}

function handoverSyncFormatSummaryTable_(target, startRow, startCol, rowCount, colCount) {
  target.getRange(startRow, startCol, rowCount, colCount)
    .setFontFamily("Times New Roman")
    .setFontSize(13)
    .setVerticalAlignment("middle");

  target.getRange(startRow, startCol, 1, colCount)
    .setFontWeight("bold")
    .setHorizontalAlignment("left");

  target.getRange(startRow + 1, startCol, rowCount - 1, colCount)
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true);

  target.getRange(startRow + 1, startCol, 1, colCount).setFontWeight("bold");

  if (rowCount > 2) {
    target.getRange(startRow + rowCount - 1, startCol, 1, colCount).setFontWeight("bold");
  }
  target.getRange(startRow + 1, startCol + colCount - 1, rowCount - 1, 1).setFontWeight("bold");
}

function handoverSyncEnsureColumns_(target, neededLastCol) {
  const maxCols = target.getMaxColumns();
  if (neededLastCol > maxCols) {
    target.insertColumnsAfter(maxCols, neededLastCol - maxCols);
  }
}

function handoverSyncCountRows_(rows, getKey) {
  const counts = {};
  rows.forEach(row => {
    const key = handoverSyncClean_(getKey(row));
    if (!key) return;
    counts[key] = (counts[key] || 0) + 1;
  });
  return Object.keys(counts).sort().map(key => [key, counts[key]]);
}

function handoverSyncApplySignature_(target, dataCount) {
  if (!dataCount) return;

  const signatureRow = Number(dataCount) + 4;
  const dateRange = target.getRange(signatureRow, 1, 1, 3);
  const leftTop = target.getRange(signatureRow + 1, 1, 1, 3);
  const leftBottom = target.getRange(signatureRow + 2, 1, 1, 3);
  const rightTop = target.getRange(signatureRow + 1, 6, 1, 3);
  const rightBottom = target.getRange(signatureRow + 2, 6, 1, 3);

  [dateRange, leftTop, leftBottom, rightTop, rightBottom].forEach(range => range.merge());
  dateRange.setValue(handoverSyncTodayHanoiText_());
  leftTop.setValue("Nhân sự bàn giao");
  leftBottom.setValue("(Ký, ghi rõ họ tên)");
  rightTop.setValue("Người nhận");
  rightBottom.setValue("(Ký, ghi rõ họ tên)");

  target.getRange(signatureRow, 1, 3, 8)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Times New Roman")
    .setFontSize(15)
    .setFontWeight("bold");
  dateRange.setHorizontalAlignment("center");
}

function handoverSyncSourceSheet_() {
  const ss = SpreadsheetApp.openById(HANDOVER_SYNC_CONFIG.spreadsheetId);
  const sh = ss.getSheetByName(HANDOVER_SYNC_CONFIG.sourceSheetName);
  if (!sh) throw new Error("Khong thay tab " + HANDOVER_SYNC_CONFIG.sourceSheetName);
  return sh;
}

function handoverSyncTargetSheet_() {
  const ss = SpreadsheetApp.openById(HANDOVER_SYNC_CONFIG.spreadsheetId);
  const sh = ss.getSheetByName(HANDOVER_SYNC_CONFIG.targetSheetName);
  if (!sh) throw new Error("Khong thay tab " + HANDOVER_SYNC_CONFIG.targetSheetName);
  return sh;
}

function handoverSyncHeaders_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
}

function handoverSyncHeaderMap_(headers) {
  const out = {};
  headers.forEach((h, i) => out[handoverSyncNorm_(h)] = i + 1);
  return out;
}

function handoverSyncFirstCol_(col, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const c = col[handoverSyncNorm_(aliases[i])];
    if (c) return c;
  }
  return 0;
}

function handoverSyncPick_(row, col, aliases) {
  const c = handoverSyncFirstCol_(col, aliases);
  return c ? row[c - 1] : "";
}

function handoverSyncDateText_(value) {
  const key = handoverSyncDateKey_(value);
  if (!key) return "";
  const parts = key.split("-");
  return parts[2] + "/" + parts[1] + "/" + parts[0];
}

function handoverSyncDateKey_(value) {
  const text = handoverSyncClean_(value);
  if (!text) return "";

  let m = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return m[3] + "-" + String(Number(m[2])).padStart(2, "0") + "-" + String(Number(m[1])).padStart(2, "0");

  m = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + "-" + String(Number(m[2])).padStart(2, "0") + "-" + String(Number(m[3])).padStart(2, "0");

  const d = new Date(text);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, HANDOVER_SYNC_CONFIG.timezone, "yyyy-MM-dd");
  }
  return "";
}

function handoverSyncTodayHanoiText_() {
  const now = new Date();
  const day = Utilities.formatDate(now, HANDOVER_SYNC_CONFIG.timezone, "d");
  const month = Utilities.formatDate(now, HANDOVER_SYNC_CONFIG.timezone, "M");
  const year = Utilities.formatDate(now, HANDOVER_SYNC_CONFIG.timezone, "yyyy");
  return "Hà Nội, Ngày " + day + " tháng " + month + " năm " + year;
}

function handoverSyncClean_(value) {
  return String(value == null ? "" : value).trim();
}

function handoverSyncNorm_(value) {
  return handoverSyncClean_(value).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d").replace(/\s+/g, " ");
}

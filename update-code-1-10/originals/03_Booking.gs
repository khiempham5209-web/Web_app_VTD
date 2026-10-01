/*************************************************
SYNC BOOKING -> CHUNG TU_FF

Paste file nay vao Apps Script project rieng.

Luồng chạy:
- Khi tab "Booking" được thêm/sửa dòng, script kiểm tra "Row Labels".
- Nếu "Row Labels" chưa tồn tại trong cột "Số đơn hàng" của tab "Chứng từ_FF",
  script append dòng mới sang "Chứng từ_FF".
- Nếu paste dữ liệu Booking từ cột B trở đi và cột "Ngày" đang trống,
  script tự điền ngày hiện tại vào cột "Ngày".
- "Ngày lên đơn" = cột "Ngày" của tab "Booking".
- Nếu sửa tay cột "Ngày" sau đó, script chỉ update lại "Ngày lên đơn"
  bên "Chứng từ_FF", không append dòng mới.
- "Mã đơn GHTK" = lookup "Row Labels" qua tab "File đơn":
  Row Labels -> File đơn!Mã đơn hàng KH, lấy File đơn!Mã đơn.
- Khi tab "File đơn" được up/sửa sau, script chạy lại phần điền "Mã đơn GHTK"
  cho các dòng "Chứng từ_FF" đang còn trống.
- Nếu thông tin đã có trong "Chứng từ_FF" thì bỏ qua, không add trùng.

Sau khi paste code:
1. Chạy installSyncBookingToChungTuFFTrigger() một lần và cấp quyền.
2. Nếu muốn đẩy các dòng Booking hiện có, chạy syncAllBookingToChungTuFF().
3. Nếu chỉ muốn điền lại Mã đơn GHTK cho dòng Chứng từ_FF có sẵn,
   chạy fillMissingGhtkCodesInChungTuFF().
*************************************************/

const BOOKING_SYNC_CONFIG = {
  spreadsheetId: "1V39AVE2JfEtMPYqnG1-J75fKPDXU37fCnJ9Na3UXOco",
  bookingSheetName: "Booking",
  orderSheetName: "File đơn",
  targetSheetName: "Chứng từ_FF",
  timezone: "Asia/Saigon"
};

function installSyncBookingToChungTuFFTrigger() {
  const editFn = "onEditSyncBookingToChungTuFF";
  const changeFn = "onChangeSyncBookingToChungTuFF";
  ScriptApp.getProjectTriggers()
    .filter(t => [editFn, changeFn].indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger(editFn)
    .forSpreadsheet(BOOKING_SYNC_CONFIG.spreadsheetId)
    .onEdit()
    .create();

  ScriptApp.newTrigger(changeFn)
    .forSpreadsheet(BOOKING_SYNC_CONFIG.spreadsheetId)
    .onChange()
    .create();

  return "Installed triggers: " + editFn + ", " + changeFn;
}

function onEditSyncBookingToChungTuFF(e) {
  try {
    if (!e || !e.range) return;
    return bookingSyncWithLock_(function() {
      const sh = e.range.getSheet();
      if (sh.getName() === BOOKING_SYNC_CONFIG.orderSheetName) {
        return fillMissingGhtkCodesInChungTuFF();
      }
      if (sh.getName() === BOOKING_SYNC_CONFIG.targetSheetName) {
        return bookingSyncRepairControlColumnsForEditedTargetRange_(e.range);
      }
      if (sh.getName() !== BOOKING_SYNC_CONFIG.bookingSheetName) return;
      if (e.range.getRow() < 2) return;

      const bookingCol = bookingSyncHeaderMap_(bookingSyncHeaders_(sh));
      const dateCol = bookingSyncFirstCol_(bookingCol, ["ngay", "ngày", "ngay len don", "ngày lên đơn"]);
      const watchedColumns = [
        bookingSyncFirstCol_(bookingCol, ["row labels", "so don hang"]),
        bookingSyncFirstCol_(bookingCol, ["customer name", "khach hang"]),
        bookingSyncFirstCol_(bookingCol, ["cust po number", "ma po"]),
        bookingSyncFirstCol_(bookingCol, ["ship to address1", "dia chi nhan hang"]),
        bookingSyncFirstCol_(bookingCol, ["chuoi", "chuỗi", "loai sieu thi"])
      ].filter(Boolean);
      const editedDateColumn = dateCol && bookingSyncRangeIntersectsAnyColumn_(e.range, [dateCol]);
      if (!bookingSyncRangeIntersectsAnyColumn_(e.range, watchedColumns)) {
        if (editedDateColumn) {
          const startRow = e.range.getRow();
          const endRow = startRow + e.range.getNumRows() - 1;
          return bookingSyncUpdateTargetDatesFromBookingRows_(startRow, endRow);
        }
        return {ok: true, skipped: true, reason: "edit column not watched"};
      }

      const startRow = e.range.getRow();
      const endRow = startRow + e.range.getNumRows() - 1;
      bookingSyncFillMissingBookingDates_(sh, bookingCol, startRow, endRow);
      return syncBookingRowsToChungTuFF_(startRow, endRow);
    });
  } catch (err) {
    console.error(err && err.stack || err);
  }
}

function onChangeSyncBookingToChungTuFF(e) {
  try {
    const changeType = e && e.changeType ? String(e.changeType) : "";
    if (["EDIT", "INSERT_ROW", "INSERT_COLUMN", "OTHER"].indexOf(changeType) < 0) return;
    return fillMissingGhtkCodesInChungTuFF();
  } catch (err) {
    console.error(err && err.stack || err);
  }
}

function syncAllBookingToChungTuFF() {
  return bookingSyncWithLock_(function() {
    const booking = bookingSyncBookingSheet_();
    const lastRow = booking.getLastRow();
    if (lastRow < 2) return {ok: true, added: 0, skipped: 0};
    const result = syncBookingRowsToChungTuFF_(2, lastRow);
    const repaired = bookingSyncRepairControlColumns_();
    result.repairedControlRows = repaired.updatedRows;
    return result;
  });
}

function repairMissingControlColumnsInChungTuFF() {
  return bookingSyncWithLock_(bookingSyncRepairControlColumns_);
}

function fillMissingGhtkCodesInChungTuFF() {
  const target = bookingSyncTargetSheet_();
  const lastRow = target.getLastRow();
  if (lastRow < 2) return {ok: true, updated: 0, missing: 0};

  const col = bookingSyncHeaderMap_(bookingSyncHeaders_(target));
  const ghtkCol = bookingSyncFirstCol_(col, ["ma don ghtk", "ma don"]);
  const orderNoCol = bookingSyncFirstCol_(col, ["so don hang", "row labels", "od"]);
  if (!ghtkCol) throw new Error("Khong thay cot Ma don GHTK trong " + BOOKING_SYNC_CONFIG.targetSheetName);
  if (!orderNoCol) throw new Error("Khong thay cot So don hang trong " + BOOKING_SYNC_CONFIG.targetSheetName);

  const orderMap = bookingSyncGhtkMap_();
  const values = target.getRange(2, 1, lastRow - 1, target.getLastColumn()).getDisplayValues();
  const ghtkValues = target.getRange(2, ghtkCol, lastRow - 1, 1).getDisplayValues();
  let updated = 0;
  let missing = 0;

  values.forEach((row, i) => {
    const currentGhtk = bookingSyncClean_(row[ghtkCol - 1]);
    const orderNo = bookingSyncClean_(row[orderNoCol - 1]);
    if (currentGhtk || !orderNo) return;

    const maDonGhtk = orderMap[bookingSyncNorm_(orderNo)] || "";
    if (!maDonGhtk) {
      missing++;
      return;
    }

    ghtkValues[i][0] = maDonGhtk;
    updated++;
  });

  if (updated) target.getRange(2, ghtkCol, ghtkValues.length, 1).setValues(ghtkValues);
  return {ok: true, updated, missing};
}

function syncOneBookingRowToChungTuFF_(rowNumber) {
  const result = syncBookingRowsToChungTuFF_(rowNumber, rowNumber);
  return {
    ok: true,
    added: Boolean(result && result.added),
    reason: result && result.added ? "" : "missing Row Labels or already exists"
  };
}

function syncBookingRowsToChungTuFF_(startRow, endRow) {
  const booking = bookingSyncBookingSheet_();
  const target = bookingSyncTargetSheet_();
  bookingSyncEnsureTargetHeaders_(target);

  const bookingCol = bookingSyncHeaderMap_(bookingSyncHeaders_(booking));
  bookingSyncFillMissingBookingDates_(booking, bookingCol, startRow, endRow);

  const rowCount = endRow - startRow + 1;
  const bookingValues = booking.getRange(startRow, 1, rowCount, booking.getLastColumn()).getDisplayValues();
  const targetCol = bookingSyncHeaderMap_(bookingSyncHeaders_(target));
  const orderMap = bookingSyncGhtkMap_();
  const existingOrders = bookingSyncExistingTargetOrders_(target);
  const outputWidth = target.getLastColumn();
  const output = [];
  let skipped = 0;

  bookingValues.forEach(row => {
    const record = bookingSyncRecordFromRow_(row, bookingCol);
    const orderKey = bookingSyncNorm_(record.orderNo);
    if (!record.orderNo || existingOrders[orderKey]) {
      skipped++;
      return;
    }
    existingOrders[orderKey] = true;
    record.maDonGhtk = orderMap[orderKey] || "";

    const outputRow = new Array(outputWidth).fill("");
    bookingSyncSetOutputByAliases_(outputRow, targetCol, ["ngay len don"], record.orderDate || bookingSyncTodayText_());
    bookingSyncSetOutputByAliases_(outputRow, targetCol, ["ma don ghtk", "ma don"], record.maDonGhtk);
    bookingSyncSetOutputByAliases_(outputRow, targetCol, ["khach hang"], record.customer);
    bookingSyncSetOutputByAliases_(outputRow, targetCol, ["ma po", "po"], record.po);
    bookingSyncSetOutputByAliases_(outputRow, targetCol, ["so don hang", "od"], record.orderNo);
    bookingSyncSetOutputByAliases_(outputRow, targetCol, ["dia chi nhan hang"], record.address);
    bookingSyncSetOutputByAliases_(outputRow, targetCol, ["loai sieu thi"], record.storeChain);
    output.push(outputRow);
  });

  if (!output.length) return {ok: true, added: 0, skipped};

  const nextRow = Math.max(target.getLastRow() + 1, 2);
  target.getRange(nextRow, 1, output.length, outputWidth).setValues(output);
  bookingSyncApplyControlColumnsBatch_(target, targetCol, nextRow, output.length);
  return {ok: true, added: output.length, skipped};
}

function bookingSyncRecordFromRow_(row, col) {
  return {
    orderDate: bookingSyncClean_(bookingSyncPick_(row, col, ["ngay", "ngày", "ngay len don", "ngày lên đơn"])),
    orderNo: bookingSyncClean_(bookingSyncPick_(row, col, ["row labels", "so don hang"])),
    customer: bookingSyncClean_(bookingSyncPick_(row, col, ["customer name", "khach hang"])),
    po: bookingSyncClean_(bookingSyncPick_(row, col, ["cust po number", "ma po"])),
    address: bookingSyncClean_(bookingSyncPick_(row, col, ["ship to address1", "dia chi nhan hang"])),
    storeChain: bookingSyncNormalizeStoreChain_(bookingSyncPick_(row, col, ["chuoi", "chuỗi", "loai sieu thi"]))
  };
}

function bookingSyncTargetOrderExists_(target, orderNo) {
  const lastRow = target.getLastRow();
  if (lastRow < 2) return false;
  const col = bookingSyncHeaderMap_(bookingSyncHeaders_(target));
  const orderCol = bookingSyncFirstCol_(col, ["so don hang", "row labels", "od"]);
  if (!orderCol) throw new Error("Khong thay cot So don hang trong " + BOOKING_SYNC_CONFIG.targetSheetName);

  const q = bookingSyncNorm_(orderNo);
  const values = target.getRange(2, orderCol, lastRow - 1, 1).getDisplayValues();
  return values.some(row => bookingSyncNorm_(row[0]) === q);
}

function bookingSyncExistingTargetOrders_(target) {
  const lastRow = target.getLastRow();
  const out = {};
  if (lastRow < 2) return out;
  const col = bookingSyncHeaderMap_(bookingSyncHeaders_(target));
  const orderCol = bookingSyncFirstCol_(col, ["so don hang", "row labels", "od"]);
  if (!orderCol) throw new Error("Khong thay cot So don hang trong " + BOOKING_SYNC_CONFIG.targetSheetName);

  const values = target.getRange(2, orderCol, lastRow - 1, 1).getDisplayValues();
  values.forEach(row => {
    const orderNo = bookingSyncClean_(row[0]);
    if (orderNo) out[bookingSyncNorm_(orderNo)] = true;
  });
  return out;
}

function bookingSyncGhtkCodeByOrderNo_(orderNo) {
  const orderMap = bookingSyncGhtkMap_();
  return orderMap[bookingSyncNorm_(orderNo)] || "";
}

function bookingSyncGhtkMap_() {
  const sh = bookingSyncOrderSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return {};

  const col = bookingSyncHeaderMap_(bookingSyncHeaders_(sh));
  const maDonCol = bookingSyncFirstCol_(col, ["ma don", "ma don ghtk"]);
  const orderNoCol = bookingSyncFirstCol_(col, ["ma don hang kh", "so don hang", "row labels"]);
  if (!maDonCol) throw new Error("Khong thay cot Ma don trong " + BOOKING_SYNC_CONFIG.orderSheetName);
  if (!orderNoCol) throw new Error("Khong thay cot Ma don hang KH trong " + BOOKING_SYNC_CONFIG.orderSheetName);

  const values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getDisplayValues();
  const out = {};
  values.forEach(row => {
    const orderNo = bookingSyncClean_(row[orderNoCol - 1]);
    const maDon = bookingSyncClean_(row[maDonCol - 1]);
    if (orderNo && maDon) out[bookingSyncNorm_(orderNo)] = maDon;
  });
  return out;
}

function bookingSyncBookingSheet_() {
  const ss = SpreadsheetApp.openById(BOOKING_SYNC_CONFIG.spreadsheetId);
  const sh = ss.getSheetByName(BOOKING_SYNC_CONFIG.bookingSheetName);
  if (!sh) throw new Error("Khong thay tab " + BOOKING_SYNC_CONFIG.bookingSheetName);
  return sh;
}

function bookingSyncOrderSheet_() {
  const ss = SpreadsheetApp.openById(BOOKING_SYNC_CONFIG.spreadsheetId);
  const sh = ss.getSheetByName(BOOKING_SYNC_CONFIG.orderSheetName);
  if (!sh) throw new Error("Khong thay tab " + BOOKING_SYNC_CONFIG.orderSheetName);
  return sh;
}

function bookingSyncTargetSheet_() {
  const ss = SpreadsheetApp.openById(BOOKING_SYNC_CONFIG.spreadsheetId);
  const sh = ss.getSheetByName(BOOKING_SYNC_CONFIG.targetSheetName);
  if (!sh) throw new Error("Khong thay tab " + BOOKING_SYNC_CONFIG.targetSheetName);
  return sh;
}

function bookingSyncEnsureTargetHeaders_(sh) {
  const headers = [
    "Ngày lên đơn", "Mã đơn GHTK", "Khách Hàng", "Mã PO", "Số đơn hàng",
    "Địa chỉ nhận hàng", "Xác thực hóa đơn", "Ngày bàn giao CT", "Đã tạo sv",
    "Ghi chú chứng từ", "Note", "Link ảnh", "Loại siêu thị", "Trạng thái",
    "Thời gian", "User thao tác"
  ];
  const current = sh.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
  const emptyOrMismatch = current.every(v => !bookingSyncClean_(v)) || bookingSyncNorm_(current[0]) !== bookingSyncNorm_(headers[0]);
  if (emptyOrMismatch) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function bookingSyncApplyControlColumns_(target, targetCol, rowNumber) {
  bookingSyncApplyControlColumnsBatch_(target, targetCol, rowNumber, 1);
}

function bookingSyncApplyControlColumnsBatch_(target, targetCol, startRow, rowCount) {
  const templateRow = bookingSyncFindControlTemplateRow_(target, targetCol) || bookingSyncFindTemplateRow_(target, startRow);
  const columns = [
    {aliases: ["xac thuc hoa don"], value: "Chưa nhận chứng từ", fillBlankOnly: false},
    {aliases: ["da tao sv", "đã tạo sv"], value: false, fillBlankOnly: false},
    {aliases: ["loai sieu thi"], value: "", fillBlankOnly: true},
    {aliases: ["trang thai"], value: "", fillBlankOnly: true}
  ];

  columns.forEach(item => {
    const col = bookingSyncFirstCol_(targetCol, item.aliases);
    if (!col) return;
    const range = target.getRange(startRow, col, rowCount, 1);
    if (templateRow) {
      target.getRange(templateRow, col).copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      target.getRange(templateRow, col).copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    }
    if (!item.fillBlankOnly) {
      range.setValues(Array.from({length: rowCount}, () => [item.value]));
      return;
    }

    const values = range.getValues();
    let changed = false;
    values.forEach(row => {
      if (row[0] !== "" && row[0] != null) return;
      row[0] = item.value;
      changed = true;
    });
    if (changed) range.setValues(values);
  });
}

function bookingSyncRepairControlColumns_() {
  const target = bookingSyncTargetSheet_();
  const lastRow = target.getLastRow();
  if (lastRow < 2) return {ok: true, updatedRows: 0};

  const targetCol = bookingSyncHeaderMap_(bookingSyncHeaders_(target));
  const orderCol = bookingSyncFirstCol_(targetCol, ["so don hang", "row labels", "od"]);
  const ghtkCol = bookingSyncFirstCol_(targetCol, ["ma don ghtk", "ma don"]);
  const customerCol = bookingSyncFirstCol_(targetCol, ["khach hang"]);
  if (!orderCol && !ghtkCol) throw new Error("Khong thay cot So don hang/Ma don GHTK trong " + BOOKING_SYNC_CONFIG.targetSheetName);

  const templateRow = bookingSyncFindControlTemplateRow_(target, targetCol) || bookingSyncFindTemplateRow_(target, lastRow + 1);
  const rowCount = lastRow - 1;
  const targetValues = target.getRange(2, 1, rowCount, target.getLastColumn()).getDisplayValues();
  const columns = [
    {aliases: ["xac thuc hoa don"], blankValue: "Chưa nhận chứng từ"},
    {aliases: ["da tao sv", "đã tạo sv"], blankValue: false},
    {aliases: ["loai sieu thi"], blankValueFn: row => bookingSyncClassifyStoreFromCustomer_(customerCol ? row[customerCol - 1] : "")},
    {aliases: ["trang thai"], blankValue: ""}
  ];

  let updatedRows = 0;
  const touchedRows = {};

  columns.forEach(item => {
    const col = bookingSyncFirstCol_(targetCol, item.aliases);
    if (!col) return;

    const range = target.getRange(2, col, rowCount, 1);
    if (templateRow) {
      target.getRange(templateRow, col).copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      target.getRange(templateRow, col).copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    }

    const values = range.getValues();
    let changed = false;
    values.forEach((row, index) => {
      if (!bookingSyncTargetRowHasKey_(targetValues[index], orderCol, ghtkCol)) return;
      if (row[0] !== "" && row[0] != null) return;
      const nextValue = item.blankValueFn ? item.blankValueFn(targetValues[index]) : item.blankValue;
      if (nextValue === "" || nextValue == null) return;
      row[0] = nextValue;
      changed = true;
      touchedRows[index + 2] = true;
    });

    if (changed) range.setValues(values);
  });

  updatedRows = Object.keys(touchedRows).length;
  return {ok: true, updatedRows};
}

function bookingSyncRepairControlColumnsForEditedTargetRange_(range) {
  const target = range.getSheet();
  if (range.getRow() < 2) return {ok: true, skipped: true, reason: "header row"};

  const targetCol = bookingSyncHeaderMap_(bookingSyncHeaders_(target));
  const watchedColumns = [
    bookingSyncFirstCol_(targetCol, ["ma don ghtk", "ma don"]),
    bookingSyncFirstCol_(targetCol, ["so don hang", "row labels", "od"]),
    bookingSyncFirstCol_(targetCol, ["khach hang"]),
    bookingSyncFirstCol_(targetCol, ["dia chi nhan hang"])
  ].filter(Boolean);

  if (!bookingSyncRangeIntersectsAnyColumn_(range, watchedColumns)) {
    return {ok: true, skipped: true, reason: "target edit column not watched"};
  }

  const startRow = range.getRow();
  const rowCount = range.getNumRows();
  return bookingSyncRepairControlColumnsInRange_(target, targetCol, startRow, rowCount);
}

function bookingSyncRepairControlColumnsInRange_(target, targetCol, startRow, rowCount) {
  const safeStartRow = Math.max(startRow, 2);
  const safeRowCount = Math.max(0, rowCount - (safeStartRow - startRow));
  if (!safeRowCount) return {ok: true, updatedRows: 0};

  const orderCol = bookingSyncFirstCol_(targetCol, ["so don hang", "row labels", "od"]);
  const ghtkCol = bookingSyncFirstCol_(targetCol, ["ma don ghtk", "ma don"]);
  const customerCol = bookingSyncFirstCol_(targetCol, ["khach hang"]);
  if (!orderCol && !ghtkCol) throw new Error("Khong thay cot So don hang/Ma don GHTK trong " + BOOKING_SYNC_CONFIG.targetSheetName);

  const templateRow = bookingSyncFindControlTemplateRow_(target, targetCol) || bookingSyncFindTemplateRow_(target, safeStartRow);
  const targetValues = target.getRange(safeStartRow, 1, safeRowCount, target.getLastColumn()).getDisplayValues();
  const columns = [
    {aliases: ["xac thuc hoa don"], blankValue: "Chưa nhận chứng từ"},
    {aliases: ["da tao sv", "đã tạo sv"], blankValue: false},
    {aliases: ["loai sieu thi"], blankValueFn: row => bookingSyncClassifyStoreFromCustomer_(customerCol ? row[customerCol - 1] : "")},
    {aliases: ["trang thai"], blankValue: ""}
  ];
  const touchedRows = {};

  columns.forEach(item => {
    const col = bookingSyncFirstCol_(targetCol, item.aliases);
    if (!col) return;

    const targetRows = [];
    targetValues.forEach((row, index) => {
      if (bookingSyncTargetRowHasKey_(row, orderCol, ghtkCol)) targetRows.push(safeStartRow + index);
    });
    if (!targetRows.length) return;

    const minRow = Math.min.apply(null, targetRows);
    const maxRow = Math.max.apply(null, targetRows);
    const range = target.getRange(minRow, col, maxRow - minRow + 1, 1);
    if (templateRow) {
      target.getRange(templateRow, col).copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      target.getRange(templateRow, col).copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    }

    const values = range.getValues();
    let changed = false;
    targetRows.forEach(rowNumber => {
      const index = rowNumber - minRow;
      if (values[index][0] !== "" && values[index][0] != null) return;
      const sourceRow = targetValues[rowNumber - safeStartRow];
      const nextValue = item.blankValueFn ? item.blankValueFn(sourceRow) : item.blankValue;
      if (nextValue === "" || nextValue == null) return;
      values[index][0] = nextValue;
      changed = true;
      touchedRows[rowNumber] = true;
    });
    if (changed) range.setValues(values);
  });

  return {ok: true, updatedRows: Object.keys(touchedRows).length};
}

function bookingSyncTargetRowHasKey_(row, orderCol, ghtkCol) {
  const orderNo = orderCol ? bookingSyncClean_(row[orderCol - 1]) : "";
  const ghtk = ghtkCol ? bookingSyncClean_(row[ghtkCol - 1]) : "";
  return Boolean(orderNo || ghtk);
}

function bookingSyncFindControlTemplateRow_(target, targetCol) {
  const controlCols = [
    bookingSyncFirstCol_(targetCol, ["xac thuc hoa don"]),
    bookingSyncFirstCol_(targetCol, ["da tao sv", "đã tạo sv"]),
    bookingSyncFirstCol_(targetCol, ["loai sieu thi"]),
    bookingSyncFirstCol_(targetCol, ["trang thai"])
  ].filter(Boolean);
  if (!controlCols.length) return 0;

  const lastRow = target.getLastRow();
  for (let row = 2; row <= lastRow; row++) {
    for (let i = 0; i < controlCols.length; i++) {
      if (target.getRange(row, controlCols[i]).getDataValidation()) return row;
    }
  }
  return 0;
}

function bookingSyncFindTemplateRow_(target, newRowNumber) {
  const lastTemplateRow = Math.min(newRowNumber - 1, target.getLastRow());
  for (let row = lastTemplateRow; row >= 2; row--) {
    const values = target.getRange(row, 1, 1, target.getLastColumn()).getDisplayValues()[0];
    if (values.some(v => bookingSyncClean_(v))) return row;
  }
  return 0;
}

function bookingSyncSetOutputByAliases_(output, col, aliases, value) {
  const c = bookingSyncFirstCol_(col, aliases);
  if (c) output[c - 1] = value;
}

function bookingSyncFillMissingBookingDates_(booking, bookingCol, startRow, endRow) {
  const dateCol = bookingSyncFirstCol_(bookingCol, ["ngay", "ngày", "ngay len don", "ngày lên đơn"]);
  const orderCol = bookingSyncFirstCol_(bookingCol, ["row labels", "so don hang"]);
  if (!dateCol || !orderCol || endRow < startRow) return {ok: true, updated: 0};

  const rowCount = endRow - startRow + 1;
  const dateRange = booking.getRange(startRow, dateCol, rowCount, 1);
  const dateValues = dateRange.getDisplayValues();
  const orderValues = booking.getRange(startRow, orderCol, rowCount, 1).getDisplayValues();
  const today = bookingSyncTodayText_();
  let updated = 0;

  for (let i = 0; i < rowCount; i++) {
    if (!bookingSyncClean_(dateValues[i][0]) && bookingSyncClean_(orderValues[i][0])) {
      dateValues[i][0] = today;
      updated++;
    }
  }

  if (updated) dateRange.setValues(dateValues);
  return {ok: true, updated};
}

function bookingSyncUpdateTargetDatesFromBookingRows_(startRow, endRow) {
  const booking = bookingSyncBookingSheet_();
  const target = bookingSyncTargetSheet_();
  const bookingCol = bookingSyncHeaderMap_(bookingSyncHeaders_(booking));
  const targetCol = bookingSyncHeaderMap_(bookingSyncHeaders_(target));
  const bookingDateCol = bookingSyncFirstCol_(bookingCol, ["ngay", "ngày", "ngay len don", "ngày lên đơn"]);
  const bookingOrderCol = bookingSyncFirstCol_(bookingCol, ["row labels", "so don hang"]);
  const targetDateCol = bookingSyncFirstCol_(targetCol, ["ngay len don"]);
  const targetOrderCol = bookingSyncFirstCol_(targetCol, ["so don hang", "row labels", "od"]);
  if (!bookingDateCol || !bookingOrderCol || !targetDateCol || !targetOrderCol) {
    return {ok: false, reason: "missing date/order column"};
  }

  const rowCount = endRow - startRow + 1;
  const bookingValues = booking.getRange(startRow, 1, rowCount, booking.getLastColumn()).getDisplayValues();
  const dateByOrder = {};
  bookingValues.forEach(row => {
    const orderNo = bookingSyncClean_(row[bookingOrderCol - 1]);
    const orderDate = bookingSyncClean_(row[bookingDateCol - 1]);
    if (orderNo && orderDate) dateByOrder[bookingSyncNorm_(orderNo)] = orderDate;
  });

  const lastTargetRow = target.getLastRow();
  if (lastTargetRow < 2) return {ok: true, updated: 0};
  const targetValues = target.getRange(2, 1, lastTargetRow - 1, target.getLastColumn()).getDisplayValues();
  const targetDateValues = target.getRange(2, targetDateCol, lastTargetRow - 1, 1).getDisplayValues();
  let updated = 0;
  targetValues.forEach((row, index) => {
    const orderNo = bookingSyncClean_(row[targetOrderCol - 1]);
    const nextDate = dateByOrder[bookingSyncNorm_(orderNo)];
    if (!nextDate || bookingSyncClean_(row[targetDateCol - 1]) === nextDate) return;
    targetDateValues[index][0] = nextDate;
    updated++;
  });
  if (updated) target.getRange(2, targetDateCol, targetDateValues.length, 1).setValues(targetDateValues);
  return {ok: true, updated};
}

function bookingSyncWithLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function bookingSyncRangeIntersectsAnyColumn_(range, columns) {
  const firstCol = range.getColumn();
  const lastCol = firstCol + range.getNumColumns() - 1;
  return columns.some(col => col >= firstCol && col <= lastCol);
}

function bookingSyncHeaders_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
}

function bookingSyncHeaderMap_(headers) {
  const out = {};
  headers.forEach((h, i) => out[bookingSyncNorm_(h)] = i + 1);
  return out;
}

function bookingSyncFirstCol_(col, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const c = col[bookingSyncNorm_(aliases[i])];
    if (c) return c;
  }
  return 0;
}

function bookingSyncPick_(row, col, aliases) {
  const c = bookingSyncFirstCol_(col, aliases);
  return c ? row[c - 1] : "";
}

function bookingSyncNormalizeStoreChain_(value) {
  const text = bookingSyncClean_(value);
  const key = bookingSyncNorm_(text);
  if (key === "vinmart" || key === "vinmart+" || key === "winmart" || key === "winmart+") return "Winmart";
  return text;
}

function bookingSyncClassifyStoreFromCustomer_(customer) {
  const text = bookingSyncClean_(customer);
  const key = bookingSyncNorm_(text);
  if (!key) return "";

  if (key.indexOf("vinmart") >= 0 || key.indexOf("winmart") >= 0 || key.indexOf("wincommerce") >= 0 || key.indexOf("wincom") >= 0) return "Winmart";
  if (key.indexOf("circlek") >= 0 || key.indexOf("circle k") >= 0) return "Circle K";
  if (key.indexOf("7-eleven") >= 0 || key.indexOf("7 eleven") >= 0 || key.indexOf("7eleven") >= 0) return "7-ELEVEN";
  if (key.indexOf("family mart") >= 0 || key.indexOf("familymart") >= 0) return "FAMILY MART (MB)";
  if (key.indexOf("coopfood") >= 0 || key.indexOf("coop food") >= 0 || key.indexOf("co.opfood") >= 0) return "Coopfood";
  if (key.indexOf("tmart") >= 0 || key.indexOf("t mart") >= 0) return "TMart";
  if (key.indexOf("unit mart") >= 0 || key.indexOf("unitmart") >= 0) return "UNIT Mart";
  if (key.indexOf("vital") >= 0) return "Vital";
  if (key.indexOf("brg") >= 0) return "BRG";
  if (key.indexOf("aeon") >= 0 || key.indexOf("maxvalu") >= 0 || key.indexOf("max valu") >= 0) return "AEON";
  if (key.indexOf("an nam") >= 0 || key.indexOf("annam") >= 0) return "AN NAM GOURMET";

  return "";
}

function bookingSyncTodayText_() {
  return Utilities.formatDate(new Date(), BOOKING_SYNC_CONFIG.timezone, "dd/MM/yyyy");
}

function bookingSyncClean_(value) {
  return String(value == null ? "" : value).trim();
}

function bookingSyncNorm_(value) {
  return bookingSyncClean_(value).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d").replace(/\s+/g, " ");
}

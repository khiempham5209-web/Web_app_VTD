/*************************************************
SYNC KHU VUC BOOKING / FILE DON -> BOOKING

Paste file nay vao cung Apps Script project hien tai.
Code nay tach rieng, khong sua luong Booking -> Chung tu_FF cu.

Luong chay:
- Khi tab "Booking" duoc edit/paste, code lay Row Labels + Quan/dia chi
  de phan loai Khu vuc = Noi thanh / Huyen.
- Khi tab "File don" duoc edit/paste, code uu tien lay cot "Khu vuc"
  tu File don theo "Ma don hang KH".
- Ghi vao tab "Booking", cot "Khu vuc", map theo "Row Labels".
- Tat ca deu lay theo ten header, khong phu thuoc vi tri cot.
- Neu tab "Booking" chua co cot "Khu vuc", code tu chen sau cot "Note".

Sau khi paste code:
1. Chay installSyncKhuVucBookingToChungTuFFTrigger() mot lan va cap quyen.
2. Chay syncAllKhuVucBookingToChungTuFF() de fill lai du lieu dang co san.
*************************************************/

const KHUVUC_BOOKING_SYNC_CONFIG = {
  spreadsheetId: "1V39AVE2JfEtMPYqnG1-J75fKPDXU37fCnJ9Na3UXOco",
  bookingSheetName: "Booking",
  orderSheetName: "File đơn",
  bookingAreaHeader: "Khu vực",
  bookingAreaInsertAfterAliases: ["note", "cbm", "thanh tien", "thành tiền"],
  dataStartRow: 2
};

function installSyncKhuVucBookingToChungTuFFTrigger() {
  const editFn = "onEditSyncKhuVucBookingToChungTuFF";
  const changeFn = "onChangeSyncKhuVucBookingToChungTuFF";
  ScriptApp.getProjectTriggers()
    .filter(t => [editFn, changeFn].indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger(editFn)
    .forSpreadsheet(KHUVUC_BOOKING_SYNC_CONFIG.spreadsheetId)
    .onEdit()
    .create();

  ScriptApp.newTrigger(changeFn)
    .forSpreadsheet(KHUVUC_BOOKING_SYNC_CONFIG.spreadsheetId)
    .onChange()
    .create();

  return "Installed triggers: " + editFn + ", " + changeFn;
}

function onEditSyncKhuVucBookingToChungTuFF(e) {
  try {
    if (!e || !e.range) return;
    const sheetName = e.range.getSheet().getName();
    if (
      sheetName !== KHUVUC_BOOKING_SYNC_CONFIG.bookingSheetName &&
      sheetName !== KHUVUC_BOOKING_SYNC_CONFIG.orderSheetName
    ) {
      return;
    }

    return khuVucBookingSyncWithLock_(syncAllKhuVucBookingToChungTuFF_);
  } catch (err) {
    console.error(err && err.stack || err);
  }
}

function onChangeSyncKhuVucBookingToChungTuFF(e) {
  try {
    const changeType = e && e.changeType ? String(e.changeType) : "";
    if (["EDIT", "INSERT_ROW", "INSERT_COLUMN", "OTHER"].indexOf(changeType) < 0) return;
    return khuVucBookingSyncWithLock_(syncAllKhuVucBookingToChungTuFF_);
  } catch (err) {
    console.error(err && err.stack || err);
  }
}

function syncAllKhuVucBookingToChungTuFF() {
  const result = khuVucBookingSyncWithLock_(syncAllKhuVucBookingToChungTuFF_);
  console.log(JSON.stringify(result));
  return result;
}

function syncAllKhuVucBookingToChungTuFF_() {
  const booking = khuVucBookingSheet_();
  const bookingCol = khuVucBookingEnsureBookingAreaColumn_(booking);
  const orderNoCol = khuVucBookingFirstCol_(bookingCol, ["row labels", "so don hang", "od"]);
  const districtCol = khuVucBookingFirstCol_(bookingCol, ["quan", "quận", "huyen", "huyện"]);
  const addressCol = khuVucBookingFirstCol_(bookingCol, ["ship to address1", "dia chi nhan hang", "dia chi"]);
  const areaCol = khuVucBookingFirstCol_(bookingCol, ["khu vuc"]);

  if (!orderNoCol) throw new Error("Khong thay cot Row Labels trong " + KHUVUC_BOOKING_SYNC_CONFIG.bookingSheetName);
  if (!areaCol) throw new Error("Khong tao/thay cot Khu vuc trong " + KHUVUC_BOOKING_SYNC_CONFIG.bookingSheetName);

  const lastRow = booking.getLastRow();
  if (lastRow < KHUVUC_BOOKING_SYNC_CONFIG.dataStartRow) {
    return {ok: true, updated: 0, missing: 0, totalBooking: 0};
  }

  const fileDonAreaMap = khuVucFileDonAreaMap_();
  const rowCount = lastRow - KHUVUC_BOOKING_SYNC_CONFIG.dataStartRow + 1;
  const values = booking
    .getRange(KHUVUC_BOOKING_SYNC_CONFIG.dataStartRow, 1, rowCount, booking.getLastColumn())
    .getDisplayValues();
  const areaValues = booking
    .getRange(KHUVUC_BOOKING_SYNC_CONFIG.dataStartRow, areaCol, rowCount, 1)
    .getDisplayValues();

  let updated = 0;
  let missing = 0;

  values.forEach((row, index) => {
    const orderNo = khuVucBookingClean_(row[orderNoCol - 1]);
    if (!orderNo) return;

    const key = khuVucBookingNorm_(orderNo);
    const district = districtCol ? row[districtCol - 1] : "";
    const address = addressCol ? row[addressCol - 1] : "";
    const nextArea = fileDonAreaMap[key] || khuVucBookingClassifyArea_(district, address);

    if (!nextArea) {
      missing++;
      return;
    }

    if (khuVucBookingClean_(areaValues[index][0]) === nextArea) return;
    areaValues[index][0] = nextArea;
    updated++;
  });

  if (updated) {
    booking
      .getRange(KHUVUC_BOOKING_SYNC_CONFIG.dataStartRow, areaCol, areaValues.length, 1)
      .setValues(areaValues);
  }

  return {
    ok: true,
    updated: updated,
    missing: missing,
    totalBooking: values.length,
    fileDonMapped: Object.keys(fileDonAreaMap).length
  };
}

function khuVucFileDonAreaMap_() {
  const sh = khuVucBookingOrderSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < KHUVUC_BOOKING_SYNC_CONFIG.dataStartRow) return {};

  const col = khuVucBookingHeaderMap_(khuVucBookingHeaders_(sh));
  const orderNoCol = khuVucBookingFirstCol_(col, ["ma don hang kh", "so don hang", "row labels", "od"]);
  const areaCol = khuVucBookingFirstCol_(col, ["khu vuc"]);
  if (!orderNoCol || !areaCol) return {};

  const values = sh
    .getRange(KHUVUC_BOOKING_SYNC_CONFIG.dataStartRow, 1, lastRow - KHUVUC_BOOKING_SYNC_CONFIG.dataStartRow + 1, sh.getLastColumn())
    .getDisplayValues();
  const out = {};

  values.forEach(row => {
    const orderNo = khuVucBookingClean_(row[orderNoCol - 1]);
    const area = khuVucBookingNormalizeArea_(row[areaCol - 1]);
    if (orderNo && area) out[khuVucBookingNorm_(orderNo)] = area;
  });

  return out;
}

function khuVucBookingClassifyArea_(districtValue, addressValue) {
  const districtText = khuVucBookingClean_(districtValue);
  const districtKey = khuVucBookingDistrictKey_(districtText);
  if (KHUVUC_BOOKING_NOI_THANH_[districtKey]) return "Nội thành";
  if (KHUVUC_BOOKING_HUYEN_[districtKey]) return "Huyện";

  const addressKey = khuVucBookingDistrictKey_(addressValue);
  if (KHUVUC_BOOKING_NOI_THANH_[addressKey]) return "Nội thành";
  if (KHUVUC_BOOKING_HUYEN_[addressKey]) return "Huyện";

  const textKey = khuVucBookingNorm_(districtText + " " + khuVucBookingClean_(addressValue));
  if (/\bquan\b/.test(textKey)) return "Nội thành";
  if (/\bhuyen\b/.test(textKey) || /\bthi xa\b/.test(textKey)) return "Huyện";
  return "";
}

function khuVucBookingNormalizeArea_(value) {
  const key = khuVucBookingNorm_(value);
  if (!key) return "";
  if (key.indexOf("noi thanh") >= 0 || key === "nt") return "Nội thành";
  if (key.indexOf("huyen") >= 0 || key.indexOf("ngoai thanh") >= 0 || key === "h") return "Huyện";
  return khuVucBookingClean_(value);
}

function khuVucBookingDistrictKey_(value) {
  let key = khuVucBookingNorm_(value)
    .replace(/\b(quan|huyen|thi xa|tx|tp|thanh pho|phuong|xa|thi tran)\b/g, " ")
    .replace(/\bha noi\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const all = Object.keys(KHUVUC_BOOKING_NOI_THANH_).concat(Object.keys(KHUVUC_BOOKING_HUYEN_));
  for (let i = 0; i < all.length; i++) {
    if (key === all[i] || key.indexOf(all[i]) >= 0) return all[i];
  }
  return key;
}

const KHUVUC_BOOKING_NOI_THANH_ = khuVucBookingListToMap_([
  "Ba Đình",
  "Hoàn Kiếm",
  "Tây Hồ",
  "Long Biên",
  "Cầu Giấy",
  "Đống Đa",
  "Hai Bà Trưng",
  "Hoàng Mai",
  "Thanh Xuân",
  "Nam Từ Liêm",
  "Bắc Từ Liêm",
  "Hà Đông",
  "Hoài Đức"
]);

const KHUVUC_BOOKING_HUYEN_ = khuVucBookingListToMap_([
  "Sơn Tây",
  "Ba Vì",
  "Phúc Thọ",
  "Đan Phượng",
  "Quốc Oai",
  "Thạch Thất",
  "Chương Mỹ",
  "Thanh Oai",
  "Thường Tín",
  "Phú Xuyên",
  "Ứng Hòa",
  "Mỹ Đức",
  "Sóc Sơn",
  "Đông Anh",
  "Gia Lâm",
  "Mê Linh",
  "Thanh Trì"
]);

function khuVucBookingEnsureBookingAreaColumn_(booking) {
  let headers = khuVucBookingHeaders_(booking);
  let col = khuVucBookingHeaderMap_(headers);
  if (khuVucBookingFirstCol_(col, ["khu vuc"])) return col;

  const insertAfterCol = khuVucBookingFirstCol_(col, KHUVUC_BOOKING_SYNC_CONFIG.bookingAreaInsertAfterAliases) || booking.getLastColumn();
  booking.insertColumnAfter(insertAfterCol);
  booking.getRange(1, insertAfterCol + 1).setValue(KHUVUC_BOOKING_SYNC_CONFIG.bookingAreaHeader);

  const lastRow = Math.max(booking.getLastRow(), 1);
  if (lastRow >= 1) {
    booking.getRange(1, insertAfterCol, lastRow, 1)
      .copyTo(booking.getRange(1, insertAfterCol + 1, lastRow, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }

  headers = khuVucBookingHeaders_(booking);
  return khuVucBookingHeaderMap_(headers);
}

function khuVucBookingSheet_() {
  return khuVucBookingSheetByName_(KHUVUC_BOOKING_SYNC_CONFIG.bookingSheetName);
}

function khuVucBookingOrderSheet_() {
  return khuVucBookingSheetByName_(KHUVUC_BOOKING_SYNC_CONFIG.orderSheetName);
}

function khuVucBookingSheetByName_(sheetName) {
  const ss = SpreadsheetApp.openById(KHUVUC_BOOKING_SYNC_CONFIG.spreadsheetId);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error("Khong thay tab " + sheetName);
  return sh;
}

function khuVucBookingHeaders_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
}

function khuVucBookingHeaderMap_(headers) {
  const out = {};
  headers.forEach((header, index) => {
    const key = khuVucBookingNorm_(header);
    if (key && !out[key]) out[key] = index + 1;
  });
  return out;
}

function khuVucBookingFirstCol_(col, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const c = col[khuVucBookingNorm_(aliases[i])];
    if (c) return c;
  }
  return 0;
}

function khuVucBookingSyncWithLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    const result = {ok: false, skipped: true, reason: "Dang co sync khac chay, thu lai sau vai giay"};
    console.warn(JSON.stringify(result));
    return result;
  }

  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function khuVucBookingListToMap_(items) {
  const out = {};
  items.forEach(item => out[khuVucBookingNorm_(item)] = true);
  return out;
}

function khuVucBookingClean_(value) {
  return String(value == null ? "" : value).trim();
}

function khuVucBookingNorm_(value) {
  return khuVucBookingClean_(value)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/\s+/g, " ");
}

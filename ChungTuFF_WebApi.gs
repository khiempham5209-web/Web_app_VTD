/*************************************************
CHUNG TU FF WEB API

Tao Google Apps Script project rieng, paste file nay vao Code.gs,
sua CONFIG.rootFolderId neu muon chi dinh san folder Drive, roi Deploy
as Web app. Frontend goi URL /exec cua SCP rieng nay.
*************************************************/

const CONFIG = {
  spreadsheetId: "1V39AVE2JfEtMPYqnG1-J75fKPDXU37fCnJ9Na3UXOco",
  sheetName: "Chứng từ_FF",
  rootFolderId: "",
  rootFolderName: "Chung_tu_FF_Images",
  timezone: "Asia/Saigon"
};

function doGet() {
  return json_({ok: true, service: "ChungTuFF_WebApi", sheetName: CONFIG.sheetName});
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const action = String(body.action || "").trim();
    const params = body.params || {};
    const map = {init: apiInit_, lookup: apiLookup_, page: apiPage_, today: apiToday_, save: apiSave_};
    if (!map[action]) return json_({ok: false, message: "Action khong hop le: " + action});
    return json_(map[action](params));
  } catch (err) {
    return json_({ok: false, message: String(err && err.stack || err)});
  }
}

function apiInit_() {
  const sh = sheet_();
  return ok_({
    sheetName: sh.getName(),
    lastRow: sh.getLastRow(),
    lastColumn: sh.getLastColumn(),
    headers: headers_(sh),
    dataVersion: PropertiesService.getScriptProperties().getProperty("PN_DOCOPS_DATA_VERSION") || ""
  });
}

function apiLookup_(params) {
  const query = clean_(params.query || params.maDon || params.po || params.orderNo);
  if (!query) return fail_("Thieu ma don can tra.");
  const found = findRecord_(query);
  if (!found) return ok_({found: false, record: null, message: "Khong tim thay ma don trong Chung tu_FF."});
  return ok_({found: true, record: found.record});
}

function apiPage_(params) {
  const sh = sheet_();
  const lastRow = sh.getLastRow();
  const lastCol = Math.min(sh.getLastColumn(), 23);
  const startRow = Math.max(2, Number(params.startRow || params.cursor || 2));
  const pageSize = Math.max(1, Math.min(Number(params.pageSize || 500), 2000));
  if (startRow > lastRow) return ok_({records: [], count: 0, done: true, nextRow: startRow, lastRow});
  const take = Math.min(pageSize, lastRow - startRow + 1);
  const values = sh.getRange(startRow, 1, take, lastCol).getDisplayValues();
  const col = headerMap_(headers_(sh));
  const records = values.map((row, i) => recordFromRow_(row, col, startRow + i)).filter(r => r.maDon || r.po || r.orderNo);
  return ok_({records, count: records.length, done: startRow + take > lastRow, nextRow: startRow + take, lastRow});
}

function apiToday_(params) {
  params = params || {};
  const today = clean_(params.date) || Utilities.formatDate(new Date(), CONFIG.timezone, "dd/MM/yyyy");
  const todayKey = dateKey_(today);
  const rows = [];
  let startRow = 2;
  while (true) {
    const page = apiPage_({startRow: startRow, pageSize: 2000});
    Array.prototype.push.apply(rows, page.records || []);
    if (page.done || !page.nextRow || Number(page.nextRow) <= startRow) break;
    startRow = Number(page.nextRow);
  }
  const records = rows.filter(r => dateKey_(r.thoiGian || r.time) === todayKey && (r.user || r.thoiGian || r.time)).map(r => {
    r.date = dateOnly_(r.thoiGian || r.time) || today;
    r.time = timeOnly_(r.thoiGian || r.time) || r.time || "";
    r.syncStatus = "Done";
    return r;
  });
  return ok_({
    date: today,
    records,
    count: records.length,
    dataVersion: PropertiesService.getScriptProperties().getProperty("PN_DOCOPS_DATA_VERSION") || ""
  });
}

function apiSave_(params) {
  const query = clean_(params.maDon || params.query || params.po || params.orderNo);
  const rowNumber = Number(params.rowNumber || 0);
  if (!query && !rowNumber) return fail_("Thieu ma don can luu.");

  const sh = sheet_();
  const headers = headers_(sh);
  const col = headerMap_(headers);
  const found = rowNumber >= 2 ? {rowNumber, record: getRecordByRow_(sh, col, rowNumber)} : findRecord_(query);
  if (!found || !found.rowNumber) return fail_("Khong tim thay dong can cap nhat trong Chung tu_FF.");

  const maDonForFolder = clean_(query || found.record.maDon || found.record.po || found.record.orderNo);
  const files = Array.isArray(params.files) ? params.files : [];
  const upload = files.length ? uploadFiles_(maDonForFolder, files) : {linkAnh: "", files: [], folderUrl: ""};
  const now = new Date();
  const timeText = Utilities.formatDate(now, CONFIG.timezone, "dd/MM/yyyy HH:mm:ss");
  const user = clean_(params.userEmail || params.user || Session.getActiveUser().getEmail() || "");

  setCellByAliases_(sh, found.rowNumber, col, ["xac thuc hoa don"], clean_(params.xacThuc || params.xacThucHoaDon));
  setCellByAliases_(sh, found.rowNumber, col, ["ghi chu chung tu"], String(params.note || params.ghiChu || ""));
  if (upload.linkAnh) setCellByAliases_(sh, found.rowNumber, col, ["link anh"], upload.linkAnh);
  setCellByAliases_(sh, found.rowNumber, col, ["trang thai"], clean_(params.status || params.trangThai));
  setCellByAliases_(sh, found.rowNumber, col, ["thoi gian"], timeText);
  setCellByAliases_(sh, found.rowNumber, col, ["user thao tac"], user);

  SpreadsheetApp.flush();
  const dataVersion = String(Date.now());
  PropertiesService.getScriptProperties().setProperty("PN_DOCOPS_DATA_VERSION", dataVersion);
  return ok_({
    message: "Da luu thao tac chung tu.",
    rowNumber: found.rowNumber,
    linkAnh: upload.linkAnh,
    fileCount: upload.files.length,
    folderUrl: upload.folderUrl,
    time: timeText,
    user,
    dataVersion
  });
}

function sheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const sh = ss.getSheetByName(CONFIG.sheetName);
  if (!sh) throw new Error("Khong thay tab " + CONFIG.sheetName);
  return sh;
}

function headers_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
}

function headerMap_(headers) {
  const out = {};
  headers.forEach((h, i) => out[norm_(h)] = i + 1);
  return out;
}

function findRecord_(query) {
  const sh = sheet_();
  const lastRow = sh.getLastRow();
  const lastCol = Math.min(sh.getLastColumn(), 23);
  if (lastRow <= 1) return null;
  const col = headerMap_(headers_(sh));
  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  const q = norm_(query);
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const candidates = [
      pick_(row, col, ["ma don ghtk", "ma don"]),
      pick_(row, col, ["ma po", "po"]),
      pick_(row, col, ["so don hang", "od"])
    ];
    if (candidates.some(x => norm_(x) === q)) return {rowNumber: i + 2, record: recordFromRow_(row, col, i + 2)};
  }
  return null;
}

function getRecordByRow_(sh, col, rowNumber) {
  const row = sh.getRange(rowNumber, 1, 1, Math.min(sh.getLastColumn(), 23)).getDisplayValues()[0];
  return recordFromRow_(row, col, rowNumber);
}

function recordFromRow_(row, col, rowNumber) {
  const maDon = clean_(pick_(row, col, ["ma don ghtk", "ma don"]));
  const po = clean_(pick_(row, col, ["ma po", "po"]));
  const orderNo = clean_(pick_(row, col, ["so don hang", "od"]));
  const thoiGian = clean_(pick_(row, col, ["thoi gian"]));
  return {
    rowNumber,
    ngayLenDon: clean_(pick_(row, col, ["ngay len don"])),
    maDon: maDon || po || orderNo,
    maDonGhtk: maDon,
    customer: clean_(pick_(row, col, ["khach hang"])),
    po,
    orderNo,
    address: clean_(pick_(row, col, ["dia chi nhan hang"])),
    xacThuc: clean_(pick_(row, col, ["xac thuc hoa don"])),
    note: clean_(pick_(row, col, ["ghi chu chung tu"])),
    linkAnh: clean_(pick_(row, col, ["link anh"])),
    storeType: clean_(pick_(row, col, ["loai sieu thi"])),
    status: clean_(pick_(row, col, ["trang thai"])),
    thoiGian,
    time: thoiGian,
    user: clean_(pick_(row, col, ["user thao tac"]))
  };
}

function uploadFiles_(maDon, files) {
  const root = rootFolder_();
  const useFolder = files.length > 1;
  const folder = useFolder ? getOrCreateSubFolder_(root, safeName_(maDon || "chung-tu")) : root;
  const uploaded = [];
  files.forEach((file, index) => {
    const mimeType = clean_(file.mimeType || file.type || "image/jpeg");
    const base64 = String(file.base64 || file.data || "").replace(/^data:[^,]+,/, "");
    if (!base64) throw new Error("Anh " + (index + 1) + " thieu du lieu base64.");
    const fileName = pnImageFileName_(maDon, file.fileName || file.name, index);
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, fileName);
    const driveFile = folder.createFile(blob);
    driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    uploaded.push({
      id: driveFile.getId(),
      name: driveFile.getName(),
      url: driveFile.getUrl(),
      directUrl: "https://drive.google.com/uc?id=" + driveFile.getId()
    });
  });
  if (useFolder) folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {
    linkAnh: useFolder ? folder.getUrl() : uploaded[0].directUrl,
    folderUrl: useFolder ? folder.getUrl() : "",
    files: uploaded
  };
}

function pnImageFileName_(maDon, name, index) {
  let fileName = safeName_(name || "");
  const extMatch = fileName.match(/(\.[a-z0-9]{2,5})$/i);
  const ext = extMatch ? extMatch[1] : ".jpg";
  if (!fileName || /^VTD_/i.test(fileName) || !/^PN_/i.test(fileName)) {
    fileName = "PN_" + safeName_(maDon || "chung-tu") + "_" + (index + 1) + ext;
  }
  return fileName;
}

function rootFolder_() {
  if (CONFIG.rootFolderId) return DriveApp.getFolderById(CONFIG.rootFolderId);
  const folders = DriveApp.getFoldersByName(CONFIG.rootFolderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(CONFIG.rootFolderName);
}

function getOrCreateSubFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}

function setCellByAliases_(sh, rowNumber, col, aliases, value) {
  const c = firstCol_(col, aliases);
  if (!c) return;
  sh.getRange(rowNumber, c).setValue(value);
}

function firstCol_(col, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const c = col[norm_(aliases[i])];
    if (c) return c;
  }
  return 0;
}

function pick_(row, col, aliases) {
  const c = firstCol_(col, aliases);
  return c ? row[c - 1] : "";
}

function clean_(value) {
  return String(value == null ? "" : value).trim();
}

function norm_(value) {
  return clean_(value).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d").replace(/\s+/g, " ");
}

function dateKey_(value) {
  const text = clean_(value);
  const m = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (!m) return text;
  return m[3] + "-" + String(Number(m[2])).padStart(2, "0") + "-" + String(Number(m[1])).padStart(2, "0");
}

function dateOnly_(value) {
  const m = clean_(value).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? `${String(Number(m[1])).padStart(2, "0")}/${String(Number(m[2])).padStart(2, "0")}/${m[3]}` : "";
}

function timeOnly_(value) {
  const m = clean_(value).match(/\b(\d{1,2}:\d{2}(?::\d{2})?)\b/);
  return m ? m[1] : "";
}

function safeName_(value) {
  return clean_(value).replace(/[\\/:*?"<>|#%{}~&]/g, "_").slice(0, 120) || "file";
}

function ok_(data) {
  data = data || {};
  data.ok = true;
  return data;
}

function fail_(message) {
  return {ok: false, message};
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

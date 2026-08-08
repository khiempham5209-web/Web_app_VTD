/*************************************************
CHUNG TU FF WEB API

Tao Google Apps Script project rieng, paste file nay vao Code.gs,
sua CONFIG.rootFolderId neu muon chi dinh san folder Drive, roi Deploy
as Web app. Frontend goi URL /exec cua SCP rieng nay.
*************************************************/

const CONFIG = {
  spreadsheetId: "1V39AVE2JfEtMPYqnG1-J75fKPDXU37fCnJ9Na3UXOco",
  sheetName: "Chứng từ_FF",
  skuSheetName: "DS SKU",
  productReturnSheetName: "Hoàn sản phẩm",
  skuSchemaVersion: 3,
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
    const map = {init: apiInit_, lookup: apiLookup_, page: apiPage_, today: apiToday_, skuInit: apiSkuInit_, skuPage: apiSkuPage_, save: apiSave_};
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

function apiSkuInit_() {
  const sh = skuSheet_();
  return ok_({
    sheetName: sh.getName(),
    lastRow: sh.getLastRow(),
    lastColumn: sh.getLastColumn(),
    headers: headers_(sh),
    schemaVersion: CONFIG.skuSchemaVersion,
    dataVersion: skuDataVersion_(sh)
  });
}

function apiSkuPage_(params) {
  const sh = skuSheet_();
  const lastRow = sh.getLastRow();
  const lastCol = Math.min(sh.getLastColumn(), 10);
  const startRow = Math.max(2, Number(params.startRow || params.cursor || 2));
  const pageSize = Math.max(1, Math.min(Number(params.pageSize || 500), 2000));
  if (startRow > lastRow) return ok_({records: [], count: 0, done: true, nextRow: startRow, lastRow});
  const take = Math.min(pageSize, lastRow - startRow + 1);
  const values = sh.getRange(startRow, 1, take, lastCol).getDisplayValues();
  const col = headerMap_(headers_(sh));
  const records = values.map((row, index) => {
    const dimensions = parseSkuDimensions_(pick_(row, col, ["kich thuoc", "kich thuoc cm", "kich thuoc (cm)"]));
    return {
      rowNumber: startRow + index,
      material: clean_(pick_(row, col, ["ma vat tu", "material", "ma sp"])),
      barcode: clean_(pick_(row, col, ["barcode", "ma vach"])),
      name: clean_(pick_(row, col, ["ten vat tu", "ten san pham", "ten sku"])),
      quyCach: clean_(pick_(row, col, ["quy cach", "don vi"])),
      dimensions: clean_(pick_(row, col, ["kich thuoc", "kich thuoc cm", "kich thuoc (cm)"])),
      lengthCm: dimensions.lengthCm,
      widthCm: dimensions.widthCm,
      heightCm: dimensions.heightCm,
      weightGram: parseSkuWeightGram_(pick_(row, col, ["khoi luong", "khoi luong g", "khoi luong (g)", "trong luong"]))
    };
  }).filter(row => row.material || row.barcode || row.name);
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
  const productIndex = productReturnIndexForDate_(today);
  const records = rows.filter(r => dateKey_(r.thoiGian || r.time) === todayKey && (r.user || r.thoiGian || r.time)).map(r => {
    const maDonKey = norm_(r.maDonGhtk || r.maDon);
    const orderNoKey = norm_(r.orderNo);
    const products = (maDonKey && productIndex.byMaDon[maDonKey]) || (orderNoKey && productIndex.byOrderNo[orderNoKey]) || [];
    r.date = dateOnly_(r.thoiGian || r.time) || today;
    r.time = timeOnly_(r.thoiGian || r.time) || r.time || "";
    r.syncStatus = "Done";
    r.products = products;
    r.productCount = products.length;
    r.returnType = products.length ? "Sản phẩm + chứng từ" : "Chứng từ";
    return r;
  });
  return ok_({
    date: today,
    records,
    count: records.length,
    dataVersion: PropertiesService.getScriptProperties().getProperty("PN_DOCOPS_DATA_VERSION") || ""
  });
}

function productReturnIndexForDate_(dateText) {
  const out = {byMaDon: {}, byOrderNo: {}};
  const sh = productReturnSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return out;
  const headers = headers_(sh);
  const col = headerMap_(headers);
  const rows = sh.getRange(2, 1, lastRow - 1, headers.length).getDisplayValues();
  const targetDateKey = dateKey_(dateText);
  rows.forEach((row, index) => {
    const date = clean_(pick_(row, col, ["ngay hoan tra", "ngay"]));
    if (dateKey_(date) !== targetDateKey) return;
    const maDon = clean_(pick_(row, col, ["ma don ghtk", "ma don"]));
    const orderNo = clean_(pick_(row, col, ["so don hang", "od"]));
    const dimensions = clean_(pick_(row, col, ["kich thuoc", "kich thuoc cm", "kich thuoc (cm)"]));
    const parsedDimensions = parseSkuDimensions_(dimensions);
    const product = {
      rowNumber: index + 2,
      maDon,
      orderNo,
      date,
      time: clean_(pick_(row, col, ["thoi gian", "gio"])),
      customer: clean_(pick_(row, col, ["ten khach hang", "khach hang"])),
      po: clean_(pick_(row, col, ["ma po", "po"])),
      material: clean_(pick_(row, col, ["ma vat tu", "material"])),
      barcode: clean_(pick_(row, col, ["barcode", "ma vach"])),
      name: clean_(pick_(row, col, ["ten san pham", "ten vat tu"])),
      quantity: clean_(pick_(row, col, ["so luong"])),
      status: clean_(pick_(row, col, ["tinh trang ff", "tinh trang"])),
      note: clean_(pick_(row, col, ["ghi chu"])),
      dimensions,
      lengthCm: parsedDimensions.lengthCm,
      widthCm: parsedDimensions.widthCm,
      heightCm: parsedDimensions.heightCm,
      weight: clean_(pick_(row, col, ["khoi luong", "khoi luong g", "khoi luong (g)", "trong luong"])),
      cbm: clean_(pick_(row, col, ["cbm"])),
      imageLink: clean_(pick_(row, col, ["hinh anh", "link anh"])),
      user: clean_(pick_(row, col, ["user thao tac", "user"]))
    };
    const maDonKey = norm_(maDon);
    const orderNoKey = norm_(orderNo);
    if (maDonKey) (out.byMaDon[maDonKey] || (out.byMaDon[maDonKey] = [])).push(product);
    if (orderNoKey) (out.byOrderNo[orderNoKey] || (out.byOrderNo[orderNoKey] = [])).push(product);
  });
  return out;
}

function apiSave_(params) {
  const query = clean_(params.maDon || params.query || params.po || params.orderNo);
  const rowNumber = Number(params.rowNumber || 0);
  if (!query && !rowNumber) return fail_("Thieu ma don can luu.");

  const clientId = clean_(params.clientId);
  const saved = readSavedRequest_(clientId);
  if (saved) return saved;

  const returnType = clean_(params.returnType || params.loaiHoan || "Chung tu");
  const normalizedType = norm_(returnType);
  const hasProducts = normalizedType === "san pham + chung tu" || normalizedType === "product-docs";
  if (normalizedType !== "chung tu" && normalizedType !== "docs" && !hasProducts) return fail_("Loai hoan khong hop le.");

  const sh = sheet_();
  const headers = headers_(sh);
  const col = headerMap_(headers);
  const found = rowNumber >= 2 ? {rowNumber, record: getRecordByRow_(sh, col, rowNumber)} : findRecord_(query);
  if (!found || !found.rowNumber) return fail_("Khong tim thay dong can cap nhat trong Chung tu_FF.");

  const maDonForFolder = clean_(query || found.record.maDon || found.record.po || found.record.orderNo);
  const files = Array.isArray(params.files) ? params.files : [];
  if (!files.length) return fail_("Chung tu can it nhat 1 anh.");
  const items = hasProducts ? normalizeProductItems_(params.items) : [];
  if (hasProducts && !items.length) return fail_("San pham + chung tu can it nhat 1 SKU.");
  for (let i = 0; i < items.length; i++) {
    const invalid = validateProductItem_(items[i], i + 1);
    if (invalid) return fail_(invalid);
  }

  const upload = uploadFiles_(maDonForFolder, files);
  const productUploads = items.map(item => uploadProductFiles_(found.record.orderNo || params.orderNo || maDonForFolder, item));
  const now = new Date();
  const dateText = Utilities.formatDate(now, CONFIG.timezone, "dd/MM/yyyy");
  const clockText = Utilities.formatDate(now, CONFIG.timezone, "HH:mm:ss");
  const timeText = dateText + " " + clockText;
  const user = clean_(params.userEmail || params.user || Session.getActiveUser().getEmail() || "");

  setCellByAliases_(sh, found.rowNumber, col, ["xac thuc hoa don"], clean_(params.xacThuc || params.xacThucHoaDon));
  setCellByAliases_(sh, found.rowNumber, col, ["ghi chu chung tu"], String(params.note || params.ghiChu || ""));
  if (upload.linkAnh) setCellByAliases_(sh, found.rowNumber, col, ["link anh"], upload.linkAnh);
  setCellByAliases_(sh, found.rowNumber, col, ["trang thai"], clean_(params.status || params.trangThai));
  setCellByAliases_(sh, found.rowNumber, col, ["thoi gian"], timeText);
  setCellByAliases_(sh, found.rowNumber, col, ["user thao tac"], user);

  const productRows = [];
  if (hasProducts) {
    const productSheet = productReturnSheet_();
    const productHeaders = headers_(productSheet);
    const productCol = headerMap_(productHeaders);
    items.forEach((item, index) => {
      const values = new Array(productHeaders.length).fill("");
      setArrayByAliases_(values, productCol, ["ngay hoan tra"], dateText);
      setArrayByAliases_(values, productCol, ["thoi gian"], clockText);
      setArrayByAliases_(values, productCol, ["ma don", "ma don ghtk"], found.record.maDonGhtk || found.record.maDon || query);
      setArrayByAliases_(values, productCol, ["ten khach hang", "khach hang"], found.record.customer || "");
      setArrayByAliases_(values, productCol, ["so don hang", "od"], found.record.orderNo || "");
      setArrayByAliases_(values, productCol, ["ma po", "po"], found.record.po || "");
      setArrayByAliases_(values, productCol, ["ma vat tu", "material"], item.material);
      setArrayByAliases_(values, productCol, ["barcode", "ma vach"], item.barcode);
      setArrayByAliases_(values, productCol, ["ten san pham", "ten vat tu"], item.name);
      setArrayByAliases_(values, productCol, ["so luong"], item.quantity);
      setArrayByAliases_(values, productCol, ["tinh trang ff", "tinh trang"], item.status);
      setArrayByAliases_(values, productCol, ["ghi chu"], item.note);
      setArrayByAliases_(values, productCol, ["kich thuoc", "kich thuoc cm", "kich thuoc (cm)"], formatProductDimensions_(item));
      setArrayByAliases_(values, productCol, ["khoi luong", "khoi luong g", "khoi luong (g)", "trong luong"], formatProductWeight_(item.weightGram));
      setArrayByAliases_(values, productCol, ["loai sieu thi"], found.record.storeType || "");
      setArrayByAliases_(values, productCol, ["hinh anh", "link anh"], productUploads[index].linkAnh);
      setArrayByAliases_(values, productCol, ["user thao tac", "user"], user);
      setArrayByAliases_(values, productCol, ["cbm"], calculateProductCbm_(item));
      productRows.push(values);
    });
    if (productRows.length) {
      const startRow = Math.max(2, productSheet.getLastRow() + 1);
      productSheet.getRange(startRow, 1, productRows.length, productHeaders.length).setValues(productRows);
    }
  }

  let skuMeasurementsUpdated = 0;
  if (hasProducts) {
    try {
      skuMeasurementsUpdated = fillBlankSkuMeasurements_(items);
    } catch (skuUpdateError) {
      console.error(skuUpdateError && skuUpdateError.stack || skuUpdateError);
    }
  }

  SpreadsheetApp.flush();
  callRejectSyncLocal_(found.rowNumber);
  const dataVersion = String(Date.now());
  PropertiesService.getScriptProperties().setProperty("PN_DOCOPS_DATA_VERSION", dataVersion);
  const response = ok_({
    message: "Da luu thao tac chung tu.",
    rowNumber: found.rowNumber,
    linkAnh: upload.linkAnh,
    fileCount: upload.files.length,
    folderUrl: upload.folderUrl,
    time: timeText,
    user,
    productCount: productRows.length,
    productLinks: productUploads.map(row => row.linkAnh),
    skuMeasurementsUpdated,
    dataVersion
  });
  rememberSavedRequest_(clientId, response);
  return response;
}

function normalizeProductItems_(items) {
  return (Array.isArray(items) ? items : []).map(item => ({
    clientItemId: clean_(item.clientItemId),
    material: clean_(item.material || item.maVatTu),
    barcode: clean_(item.barcode),
    name: clean_(item.name || item.tenSanPham),
    quantity: Number(item.quantity || item.soLuong || 0),
    status: clean_(item.status || item.tinhTrang),
    note: clean_(item.note || item.ghiChu),
    lengthCm: Number(item.lengthCm || item.dai || 0),
    widthCm: Number(item.widthCm || item.rong || 0),
    heightCm: Number(item.heightCm || item.cao || 0),
    weightGram: Number(item.weightGram || item.khoiLuong || 0),
    files: Array.isArray(item.files) ? item.files : (Array.isArray(item.images) ? item.images : [])
  }));
}

function validateProductItem_(item, number) {
  const label = "SKU " + number + ": ";
  if (!item.name) return label + "thieu ten san pham.";
  if (!item.barcode) return label + "thieu Barcode.";
  if (!item.material) return label + "thieu ma vat tu.";
  if (!item.status) return label + "thieu tinh trang.";
  if (!(item.quantity > 0)) return label + "thieu so luong hop le.";
  if (!item.note) return label + "thieu ghi chu.";
  if (!(item.lengthCm > 0) || !(item.widthCm > 0) || !(item.heightCm > 0)) return label + "thieu kich thuoc Dai, Rong hoac Cao.";
  if (!(item.weightGram > 0)) return label + "thieu khoi luong theo gram.";
  if (!item.files.length) return label + "can it nhat 1 anh.";
  return "";
}

function calculateProductCbm_(item) {
  const cbm = Number(item.lengthCm) * Number(item.widthCm) * Number(item.heightCm) * Number(item.quantity) / 1000000;
  return Number(cbm.toFixed(6));
}

function formatProductDimensions_(item) {
  return cleanNumberText_(item.lengthCm) + " x " + cleanNumberText_(item.widthCm) + " x " + cleanNumberText_(item.heightCm) + " cm";
}

function formatProductWeight_(weightGram) {
  return cleanNumberText_(weightGram) + " g";
}

function parseSkuDimensions_(value) {
  const text = clean_(value).replace(/,/g, ".").replace(/cm/gi, "");
  const parts = text.split(/[xX×*]/).map(part => Number(String(part).trim())).filter(number => Number.isFinite(number));
  return {
    lengthCm: parts.length > 0 && parts[0] > 0 ? parts[0] : 0,
    widthCm: parts.length > 1 && parts[1] > 0 ? parts[1] : 0,
    heightCm: parts.length > 2 && parts[2] > 0 ? parts[2] : 0
  };
}

function parseSkuWeightGram_(value) {
  const text = clean_(value).toLowerCase().replace(/,/g, ".");
  const match = text.match(/-?\d+(?:\.\d+)?/);
  if (!match) return 0;
  const number = Number(match[0]);
  if (!(number > 0)) return 0;
  return /kg\b/.test(text) ? number * 1000 : number;
}

function fillBlankSkuMeasurements_(items) {
  if (!items || !items.length) return 0;
  const sh = skuSheet_();
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return 0;

  const headers = headers_(sh);
  const col = headerMap_(headers);
  const materialCol = findColumnByAliases_(col, ["ma vat tu", "material", "ma sp"]);
  const barcodeCol = findColumnByAliases_(col, ["barcode", "ma vach"]);
  const dimensionsCol = findColumnByAliases_(col, ["kich thuoc", "kich thuoc cm", "kich thuoc (cm)"]);
  const weightCol = findColumnByAliases_(col, ["khoi luong", "khoi luong g", "khoi luong (g)", "trong luong"]);
  if ((!materialCol && !barcodeCol) || (!dimensionsCol && !weightCol)) return 0;

  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  const byMaterial = {};
  const byBarcode = {};
  values.forEach((row, index) => {
    const sheetRow = index + 2;
    const material = materialCol ? clean_(row[materialCol - 1]) : "";
    const barcode = barcodeCol ? clean_(row[barcodeCol - 1]) : "";
    if (material && !byMaterial[material]) byMaterial[material] = sheetRow;
    if (barcode && !byBarcode[barcode]) byBarcode[barcode] = sheetRow;
  });

  let updated = 0;
  items.forEach(item => {
    const sheetRow = byMaterial[clean_(item.material)] || byBarcode[clean_(item.barcode)] || 0;
    if (!sheetRow) return;
    const current = values[sheetRow - 2];
    if (dimensionsCol && !clean_(current[dimensionsCol - 1])) {
      sh.getRange(sheetRow, dimensionsCol).setValue(
        cleanNumberText_(item.lengthCm) + "x" + cleanNumberText_(item.widthCm) + "x" + cleanNumberText_(item.heightCm)
      );
      current[dimensionsCol - 1] = "updated";
      updated++;
    }
    if (weightCol && !clean_(current[weightCol - 1])) {
      sh.getRange(sheetRow, weightCol).setValue(Number(item.weightGram));
      current[weightCol - 1] = "updated";
      updated++;
    }
  });
  if (updated > 0) PropertiesService.getScriptProperties().setProperty("PN_DOCOPS_SKU_DATA_VERSION", String(Date.now()));
  return updated;
}

function findColumnByAliases_(columnMap, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const column = Number(columnMap[norm_(aliases[i])] || 0);
    if (column) return column;
  }
  return 0;
}

function cleanNumberText_(value) {
  const number = Number(value || 0);
  return Number.isInteger(number) ? String(number) : String(Number(number.toFixed(2)));
}

function skuDataVersion_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = Math.min(sh.getLastColumn(), 10);
  const values = lastRow > 0 && lastCol > 0 ? sh.getRange(1, 1, lastRow, lastCol).getDisplayValues() : [];
  const text = JSON.stringify(values);
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text, Utilities.Charset.UTF_8);
  return digest.map(value => (value + 256).toString(16).slice(-2)).join("");
}

function callRejectSyncLocal_(rowNumber) {
  try {
    if (typeof syncOneChungTuKhongDatYCBySourceRow_ === "function") {
      syncOneChungTuKhongDatYCBySourceRow_(rowNumber);
    }
  } catch (err) {
    console.error(err && err.stack || err);
  }
}

function sheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const sh = ss.getSheetByName(CONFIG.sheetName);
  if (!sh) throw new Error("Khong thay tab " + CONFIG.sheetName);
  return sh;
}

function skuSheet_() {
  const sh = SpreadsheetApp.openById(CONFIG.spreadsheetId).getSheetByName(CONFIG.skuSheetName);
  if (!sh) throw new Error("Khong thay tab " + CONFIG.skuSheetName);
  return sh;
}

function productReturnSheet_() {
  const sh = SpreadsheetApp.openById(CONFIG.spreadsheetId).getSheetByName(CONFIG.productReturnSheetName);
  if (!sh) throw new Error("Khong thay tab " + CONFIG.productReturnSheetName);
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

function uploadProductFiles_(orderNo, item) {
  const files = item.files || [];
  const root = rootFolder_();
  const useFolder = files.length > 1;
  const folder = useFolder ? getOrCreateSubFolder_(root, safeName_(orderNo || "hoan-san-pham")) : root;
  const uploaded = [];
  files.forEach((file, index) => {
    const mimeType = clean_(file.mimeType || file.type || "image/jpeg");
    const base64 = String(file.base64 || file.data || "").replace(/^data:[^,]+,/, "");
    if (!base64) throw new Error("Anh san pham " + (index + 1) + " thieu du lieu base64.");
    const extMatch = safeName_(file.fileName || file.name || "").match(/(\.[a-z0-9]{2,5})$/i);
    const ext = extMatch ? extMatch[1] : ".jpg";
    const fileName = "PN_" + safeName_(orderNo || "don") + "_" + safeName_(item.material || item.barcode || "sku") + "_" + (index + 1) + ext;
    const driveFile = folder.createFile(Utilities.newBlob(Utilities.base64Decode(base64), mimeType, fileName));
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

function setArrayByAliases_(row, col, aliases, value) {
  const c = firstCol_(col, aliases);
  if (c) row[c - 1] = value;
}

function readSavedRequest_(clientId) {
  if (!clientId) return null;
  const raw = PropertiesService.getScriptProperties().getProperty("PN_SAVE_" + safePropertyKey_(clientId));
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (err) { return null; }
}

function rememberSavedRequest_(clientId, response) {
  if (!clientId) return;
  const props = PropertiesService.getScriptProperties();
  props.setProperty("PN_SAVE_" + safePropertyKey_(clientId), JSON.stringify(response));
  const savedKeys = Object.keys(props.getProperties()).filter(key => key.indexOf("PN_SAVE_") === 0);
  if (savedKeys.length > 300) props.deleteProperty(savedKeys[0]);
}

function safePropertyKey_(value) {
  return clean_(value).replace(/[^a-z0-9_-]/gi, "_").slice(0, 80);
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

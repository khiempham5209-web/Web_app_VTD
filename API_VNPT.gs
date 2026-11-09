/**
 * API_VNPT.gs — API nhập hoàn VNPT cho App Chứng Từ (dựa trên API TH).
 * Cùng project với VanHanh_VNPT.gs + CanhBaoLoi_VNPT.gs, deploy Web app riêng (URL riêng với TH).
 * Giữ nguyên payload/callback như TH. Nhận cả tên action cũ (saveReturnTh…) và tên mới (saveReturnVnpt…).
 * Sau khi lưu/sửa/xóa: gọi vnptRefreshHoan() để cập nhật cột HOÀN ở Tổng hợp.
 */
const RETURN_VNPT_CONFIG = {
  spreadsheetId: '1EOyNJyjUoXXXmOH-OJCD_Vsyauhmm2Uj_Ujnqaj6e-s',
  inputSheet: 'Hoàn trả',
  skuSheet: 'Thông tin sku',
  imageFolderName: 'HangHoanVNPT_Images',
  imageRootFolderId: '',
  logSheet: 'Log sửa VNPT',
  alertEmail: 'khiempham5209@gmail.com',
  timezone: 'Asia/Ho_Chi_Minh'
};


function returnVnpt_testAuth() {
  const root = RETURN_VNPT_CONFIG.imageRootFolderId
    ? DriveApp.getFolderById(RETURN_VNPT_CONFIG.imageRootFolderId)
    : DriveApp.getRootFolder();
  const folder = returnVnpt_getOrCreateFolder_(root, RETURN_VNPT_CONFIG.imageFolderName);
  const file = folder.createFile(Utilities.newBlob('ok', 'text/plain', 'return-vnpt-auth-test.txt'));
  file.setTrashed(true);
  Logger.log('Drive OK: ' + folder.getName());
  return 'Drive OK: ' + folder.getName();
}


function doGet(e) {
  return returnVnpt_json_({
    ok: true,
    service: 'NhapHoanVNPT',
    actions: ['listSku', 'skuSync', 'keySync', 'createSku', 'saveReturnVnpt', 'updateReturnVnpt', 'deleteReturnVnpt', 'listToday']
  });
}


function doPost(e) {
  try {
    const payload = returnVnpt_parseBody_(e);
    const action = String(payload.action || '').trim();
    const params = payload.params || {};
    if (action === 'listSku') return returnVnpt_json_(returnVnpt_listSku_());
    if (action === 'skuSync') return returnVnpt_json_(returnVnpt_skuSync_(params));
    if (action === 'keySync') return returnVnpt_json_(returnVnpt_keySync_(params));
    if (action === 'listToday') return returnVnpt_json_(returnVnpt_listToday_(params));
    const writers = {
      createSku: () => returnVnpt_createSku_(params),
      saveReturnVnpt: () => returnVnpt_save_(params),
      updateReturnVnpt: () => returnVnpt_update_(params, 'update'),
      deleteReturnVnpt: () => returnVnpt_update_(params, 'delete')
    };
    // Tên action cũ của TH: cho phép app dùng lại đúng client TH, chỉ đổi URL.
    writers.saveReturnTh = writers.saveReturnVnpt;
    writers.updateReturnTh = writers.updateReturnVnpt;
    writers.deleteReturnTh = writers.deleteReturnVnpt;
    if (writers[action]) return returnVnpt_json_(returnVnpt_withLock_(writers[action]));
    return returnVnpt_json_({ok: false, message: 'Action khong hop le: ' + action});
  } catch (err) {
    return returnVnpt_json_({ok: false, message: String(err && err.message ? err.message : err)});
  }
}




function returnVnpt_save_(params) {
  const maDon = returnVnpt_cleanMaDon_(
    params.maDon ||
    params.ma_don ||
    params.orderCode ||
    params.order_code ||
    params['Mã đơn'] ||
    params['Ma don'] ||
    params['ma don'] ||
    ''
  ).trim();
  const items = returnVnpt_itemsFromParams_(params);
  const userEmail = String(params.userEmail || params.email || '').trim();
  const files = Array.isArray(params.files) ? params.files : [];


  if (!maDon) return {ok: false, message: 'Thieu Ma don'};
  if (!items.length) return {ok: false, message: 'Thieu SKU'};
  if (items.length === 1 && !items[0].files.length && files.length) items[0].files = files;
  for (let i = 0; i < items.length; i++) {
    if (!items[i].tenSanPham) return {ok: false, message: 'SKU ' + (i + 1) + ' thieu Ten san pham'};
    if (!items[i].soLuong) return {ok: false, message: 'SKU ' + (i + 1) + ' thieu So luong'};
    if (!items[i].tinhTrang) return {ok: false, message: 'SKU ' + (i + 1) + ' thieu Tinh trang'};
    if (!items[i].files.length) return {ok: false, message: 'SKU ' + (i + 1) + ' thieu hinh anh'};
    const iccidError = returnVnpt_iccidError_(items[i], 'SKU ' + (i + 1));
    if (iccidError) return {ok: false, message: iccidError};
  }


  const ss = SpreadsheetApp.openById(RETURN_VNPT_CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(RETURN_VNPT_CONFIG.inputSheet);
  if (!sheet) return {ok: false, message: 'Khong tim thay tab: ' + RETURN_VNPT_CONFIG.inputSheet};


  // Kiểm tra header trước khi tải ảnh hoặc ghi dữ liệu.
  const headers = sheet.getRange(3,1,1,sheet.getLastColumn()).getDisplayValues()[0];
  const inputColumns = returnVnpt_inputColumns_(headers);
  const now = new Date();
  const dateText = Utilities.formatDate(now, RETURN_VNPT_CONFIG.timezone, 'dd/MM/yyyy');
  const timeText = Utilities.formatDate(now, RETURN_VNPT_CONFIG.timezone, 'HH:mm:ss');
  const clientBatchId = String(params.clientBatchId || params.idempotencyKey || params.clientId || '').trim();
  const batchId = returnVnpt_safeId_(clientBatchId) || ('VNPT-' + Utilities.formatDate(now, RETURN_VNPT_CONFIG.timezone, 'yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().slice(0, 8));
  const rowIds = items.map((item, index) => returnVnpt_safeId_(item.clientItemId) || (items.length === 1 ? batchId : batchId + '-' + (index + 1)));
  const existingIds = returnVnpt_existingIds_(sheet, rowIds);
  const pendingItems = items.map((item, index) => ({item, index, rowId: rowIds[index]})).filter(entry => !existingIds[entry.rowId]);
  if (!pendingItems.length) {
    return {ok: true, id: batchId, duplicate: true, row: 0, rows: 0};
  }
  const imageLinksByIndex = {};
  pendingItems.forEach(entry => {
    imageLinksByIndex[entry.index] = returnVnpt_uploadImages_(maDon, entry.rowId, entry.item.files, now, entry.index);
  });


  const rows = pendingItems.map(entry => {
    const item = entry.item;
    const index = entry.index;
    const rowId = entry.rowId;
    const rowData = {
      id: rowId,
      dateText,
      timeText,
      maDon,
      tenSanPham: item.tenSanPham,
      soLuong: item.soLuong,
      iccid: item.iccid,
      tinhTrang: item.tinhTrang,
      loaiHoan: item.loaiHoan,
      daTaoSv: '',
      ghiChu: item.ghiChu,
      imageLink: imageLinksByIndex[index] || '',
      userEmail
    };
    const row = headers.map(header => returnVnpt_valueForHeader_(header, rowData));
    returnVnpt_applyFixedInputLayout_(row, rowData, headers);
    return row;
  });


  const targetRow = returnVnpt_findFirstInputRow_(sheet, rows.length, headers);
  returnVnpt_writeInputRows_(sheet,targetRow,rows,inputColumns);
  returnVnpt_refreshSummaryIfAvailable_();


  return {ok: true, id: batchId, imageLinks: pendingItems.map(entry => imageLinksByIndex[entry.index] || ''), imageLink: imageLinksByIndex[pendingItems[0].index] || '', row: targetRow, rows: rows.length};
}


function returnVnpt_itemsFromParams_(params) {
  let items = [];
  if (Array.isArray(params.items)) {
    items = params.items;
  } else if (typeof params.items === 'string' && params.items.trim()) {
    try {
      const parsed = JSON.parse(params.items);
      if (Array.isArray(parsed)) items = parsed;
    } catch (err) {
      items = [];
    }
  }
  if (!items.length) {
    items = [{
      tenSanPham: params.tenSanPham,
      soLuong: params.soLuong,
      tinhTrang: params.tinhTrang,
      ghiChu: params.ghiChu
    }];
  }
  return items.map(item => ({
    clientItemId: String((item && (item.clientItemId || item.idempotencyKey || item.id)) || '').trim(),
    tenSanPham: String((item && (item.tenSanPham || item.product || item.name)) || '').trim(),
    soLuong: String((item && (item.soLuong || item.qty || item.quantity)) || '').trim(),
    tinhTrang: String((item && (item.tinhTrang || item.status)) || '').trim(),
    loaiHoan: returnVnpt_loaiHoan_((item && (item.loaiHoan || item['Loại hoàn'])) || params.loaiHoan || params['Loại hoàn'] || ''),
    ghiChu: String((item && (item.ghiChu || item.note)) || '').trim(),
    iccid: returnVnpt_iccidText_(item && (item.iccid || item.iccids || item.ICCID)),
    files: returnVnpt_itemFiles_(item)
  })).filter(item => item.tenSanPham || item.soLuong || item.tinhTrang || item.ghiChu || item.iccid);
}

/** ICCID: nhận mảng hoặc chuỗi (phẩy/xuống dòng/khoảng trắng), bỏ trùng, ghi chung 1 ô cách nhau ", ". */
function returnVnpt_iccidList_(value) {
  const raw = Array.isArray(value) ? value : String(value || '').split(/[\s,;]+/);
  const seen = {}, out = [];
  raw.forEach(v => {
    const code = String(v || '').trim().replace(/^'+/, '').toUpperCase();
    if (code && !seen[code]) { seen[code] = true; out.push(code); }
  });
  return out;
}

function returnVnpt_iccidText_(value) {
  return returnVnpt_iccidList_(value).join(', ');
}

/** ICCID không bắt buộc; có quét thì số mã phải bằng Số lượng. */
function returnVnpt_iccidError_(item, label) {
  const codes = returnVnpt_iccidList_(item.iccid);
  if (!codes.length) return '';
  const qty = Number(String(item.soLuong || '').replace(',', '.'));
  if (codes.length !== qty) return label + ': số ICCID (' + codes.length + ') phải bằng Số lượng (' + item.soLuong + ')';
  return '';
}


function returnVnpt_itemFiles_(item) {
  if (!item) return [];
  if (Array.isArray(item.files) && item.files.length) return item.files;
  if (Array.isArray(item.images) && item.images.length) return item.images;
  if (item.image) return [item.image];
  return [];
}


function returnVnpt_valueForHeader_(header, data) {
  const key = returnVnpt_norm_(header);
  if (!key) return '';
  if (key === 'id') return data.id;
  if (key === 'iccid') return data.iccid || '';
  if (key === 'ngay thang' || key === 'ngay') return data.dateText;
  if (key === 'thoi gian' || key === 'gio') return data.timeText;
  if (key === 'ma don' || key === 'ma don ghtk' || key.indexOf('ma don') >= 0) return data.maDon;
  if (key === 'ten san pham' || key === 'ten san pham hoan ve' || key.indexOf('ten san pham') >= 0) return data.tenSanPham;
  if (key === 'so luong' || key.indexOf('so luong') >= 0) return data.soLuong;
  if (key === 'tinh trang' || key.indexOf('tinh trang') >= 0) return data.tinhTrang;
  if (key === 'loai hoan') return data.loaiHoan || 'Hoàn trả';
  if (key === 'da tao sv' || key.indexOf('tao sv') >= 0) return data.daTaoSv;
  if (key === 'ghi chu' || key.indexOf('ghi chu') >= 0) return data.ghiChu;
  if (key === 'hinh anh' || key === 'anh' || key.indexOf('hinh anh') >= 0) return data.imageLink;
  if (key === 'user thao tac' || key === 'user email' || key === 'user_email' || key.indexOf('user') >= 0) return data.userEmail;
  return '';
}


function returnVnpt_applyFixedInputLayout_(row,data,headers) {
  // Giữ tên helper cũ, nhưng không còn ghi đè theo số cột cố định.
  if (!headers) throw new Error('Thiếu header khi chuẩn bị dữ liệu hoàn.');
  const columns=returnVnpt_inputColumns_(headers);
  Object.keys(columns).forEach(field=>{row[columns[field]]=data[field];});
}

function returnVnpt_refreshSummaryIfAvailable_() {
  try {
    // Chỉ cập nhật luồng HOÀN. Lỗi vận hành đã được báo email trong vnguard_, không làm hỏng lượt lưu.
    if (typeof vnptRefreshHoan === 'function') vnptRefreshHoan();
  } catch (err) {
    console.log('vnptRefreshHoan skipped: ' + (err && err.message ? err.message : err));
  }
}


function returnVnpt_findFirstInputRow_(sheet,rowCount,headers) {
  const startRow=4,needed=Math.max(1,Number(rowCount)||1);
  headers=headers || sheet.getRange(3,1,1,sheet.getLastColumn()).getDisplayValues()[0];
  const columns=returnVnpt_inputColumns_(headers),width=headers.length;
  const count=Math.max(0,sheet.getLastRow()-startRow+1);
  const values=count ? sheet.getRange(startRow,1,count,width).getDisplayValues() : [];
  const formulas=count ? sheet.getRange(startRow,1,count,width).getFormulas() : [];
  const fields=Object.keys(columns).filter(k=>k!=='loaiHoan' && k!=='daTaoSv');
  let consecutive=0;
  for (let i=0;i<values.length+needed;i++) {
    const row=values[i]||[],f=formulas[i]||[];
    const occupied=fields.some(k=>String(row[columns[k]]||'').trim()) ||
      Object.keys(columns).some(k=>!!f[columns[k]]);
    consecutive=occupied ? 0 : consecutive+1;
    if (consecutive===needed) return startRow+i-needed+1;
  }
  throw new Error('Không tìm được vùng trống đủ số dòng SKU.');
}

function returnVnpt_existingIds_(sheet,ids) {
  const wanted=new Set((ids||[]).map(v=>String(v||'').trim()).filter(Boolean)),found={};
  const lastRow=sheet.getLastRow();
  if (!wanted.size || lastRow<4) return found;
  const headers=sheet.getRange(3,1,1,sheet.getLastColumn()).getDisplayValues()[0];
  const idCol=returnVnpt_inputColumns_(headers).id;
  sheet.getRange(4,idCol+1,lastRow-3,1).getDisplayValues().flat().forEach(value=>{
    const key=String(value||'').trim();if(wanted.has(key))found[key]=true;
  });
  return found;
}

function returnVnpt_safeId_(value) {
  return String(value || '').trim().replace(/[^\w.-]/g, '_').slice(0, 120);
}


function returnVnpt_uploadImages_(maDon, id, files, now, skuIndex) {
  const parent = RETURN_VNPT_CONFIG.imageRootFolderId
    ? DriveApp.getFolderById(RETURN_VNPT_CONFIG.imageRootFolderId)
    : DriveApp.getRootFolder();
  const root = returnVnpt_getOrCreateFolder_(parent, RETURN_VNPT_CONFIG.imageFolderName);
  const safeOrder = returnVnpt_safeName_((maDon || id) + (skuIndex != null ? '_SKU' + (Number(skuIndex) + 1) : ''));
  if (files.length === 1) {
    const file = returnVnpt_createImageFile_(root, files[0], safeOrder, now, 0);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/uc?id=' + file.getId();
  }


  const folder = returnVnpt_getOrCreateFolder_(root, safeOrder);
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  files.forEach((filePayload, index) => returnVnpt_createImageFile_(folder, filePayload, safeOrder, now, index));
  return 'https://drive.google.com/drive/folders/' + folder.getId();
}


function returnVnpt_createImageFile_(folder, filePayload, safeOrder, now, index) {
  const mimeType = String(filePayload.mimeType || filePayload.type || 'image/jpeg');
  const base64 = returnVnpt_extractBase64_(filePayload);
  if (!base64) throw new Error('Anh khong co du lieu base64');
  const bytes = Utilities.base64Decode(base64);
  const ext = mimeType.indexOf('png') >= 0 ? 'png' : 'jpg';
  const stamp = Utilities.formatDate(now, RETURN_VNPT_CONFIG.timezone, 'yyyyMMdd_HHmmss');
  const name = safeOrder + '_' + stamp + '_' + index + '.' + ext;
  const blob = Utilities.newBlob(bytes, mimeType, name);
  return folder.createFile(blob);
}


function returnVnpt_extractBase64_(filePayload) {
  const raw = String(filePayload.base64 || filePayload.data || '');
  const match = raw.match(/^data:[^;]+;base64,(.*)$/);
  return match ? match[1] : raw;
}


function returnVnpt_getOrCreateFolder_(parent, name) {
  const iter = parent.getFoldersByName(name);
  if (iter.hasNext()) return iter.next();
  return parent.createFolder(name);
}


function returnVnpt_findHeader_(headers, names) {
  const normalized = headers.map(returnVnpt_norm_);
  for (let i = 0; i < names.length; i++) {
    const idx = normalized.indexOf(returnVnpt_norm_(names[i]));
    if (idx >= 0) return idx;
  }
  return -1;
}


function returnVnpt_norm_(value) {
  return String(value || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
}


function returnVnpt_safeName_(value) {
  return String(value || 'no-code').replace(/[\\/:*?"<>|#%{}~&]/g, '_').slice(0, 80);
}


// Structured SKU source used by the shared searchable picker in web/APK.
function returnVnpt_listSku_() {
  const ss = SpreadsheetApp.openById(RETURN_VNPT_CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(RETURN_VNPT_CONFIG.skuSheet);
  if (!sheet) return {ok: false, message: 'Khong tim thay tab SKU: ' + RETURN_VNPT_CONFIG.skuSheet};
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 4 || lastCol < 1) return {ok: true, items: []};
  const headers = sheet.getRange(3, 1, 1, lastCol).getDisplayValues()[0];
  const subSkuCol = returnVnpt_findHeader_(headers, ['SUB SKU', 'Sub SKU', 'SKU', 'Ma SKU', 'Mã SKU']);
  const nameCol = returnVnpt_findHeader_(headers, ['Tên sản phẩm', 'Tên sản phẩm hoàn về', 'Ten san pham', 'Ten SKU', 'Tên SKU', 'Product', 'Name']);
  const unitCol = returnVnpt_findHeader_(headers, ['Don vi', 'Đơn vị', 'Unit', 'DVT']);
  if (nameCol < 0) return {ok: false, message: 'Tab SKU thieu cot Ten san pham'};
  const values = sheet.getRange(4, 1, lastRow - 3, lastCol).getDisplayValues();
  const seen = {};
  const items = [];
  values.forEach(row => {
    const name = String(row[nameCol] || '').trim();
    if (!name) return;
    const subSku = subSkuCol >= 0 ? String(row[subSkuCol] || '').trim() : '';
    const unit = unitCol >= 0 ? String(row[unitCol] || '').trim() : '';
    const key = (subSku || name).toLowerCase();
    if (seen[key]) return;
    seen[key] = true;
    items.push({subSku: subSku, name: name, unit: unit});
  });
  return {ok: true, items: items};
}


function returnVnpt_createSku_(params) {
  params = params || {};
  const subSku = String(params.subSku || params.sub_sku || params.sku || '').trim();
  const name = String(params.name || params.tenSku || params.tenSanPham || '').trim();
  const unit = String(params.unit || params.donVi || '').trim();
  if (!subSku) return {ok: false, message: 'Thiếu SUB SKU'};
  if (!name) return {ok: false, message: 'Thiếu Tên SKU'};
  if (!unit) return {ok: false, message: 'Thiếu Đơn vị'};
  const ss = SpreadsheetApp.openById(RETURN_VNPT_CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(RETURN_VNPT_CONFIG.skuSheet);
  if (!sheet) return {ok: false, message: 'Không tìm thấy tab SKU: ' + RETURN_VNPT_CONFIG.skuSheet};
  const lastCol = Math.max(3, sheet.getLastColumn());
  const headers = sheet.getRange(3, 1, 1, lastCol).getDisplayValues()[0];
  const subSkuCol = returnVnpt_findHeader_(headers, ['SUB SKU', 'Sub SKU', 'SKU', 'Ma SKU', 'Mã SKU']);
  const nameCol = returnVnpt_findHeader_(headers, ['Ten SKU', 'Tên SKU', 'Tên sản phẩm', 'Ten san pham', 'Product', 'Name']);
  const unitCol = returnVnpt_findHeader_(headers, ['Don vi', 'Đơn vị', 'Unit', 'DVT']);
  if (subSkuCol < 0 || nameCol < 0 || unitCol < 0) return {ok:false, message:'Tab Thông tin sku cần đủ cột SUB SKU, Tên SKU, Đơn vị.'};
  const lastRow = sheet.getLastRow();
  if (lastRow >= 4) {
    const values = sheet.getRange(4, 1, lastRow - 3, lastCol).getDisplayValues();
    for (let i = 0; i < values.length; i++) {
      const existingSub = String(values[i][subSkuCol] || '').trim();
      const existingName = String(values[i][nameCol] || '').trim();
      if (existingSub.toLowerCase() === subSku.toLowerCase()) return {ok:false, message:'SUB SKU đã tồn tại: ' + subSku};
      if (existingName.toLowerCase() === name.toLowerCase()) return {ok:false, message:'Tên SKU đã tồn tại: ' + name};
    }
  }
  const row = new Array(lastCol).fill('');
  row[subSkuCol] = subSku;
  row[nameCol] = name;
  row[unitCol] = unit;
  const targetRow = Math.max(4, sheet.getLastRow() + 1);
  // SUB SKU VNPT có mã 19 số: ghi dạng văn bản để Sheets không đổi thành số và làm tròn.
  sheet.getRange(targetRow, subSkuCol + 1).setNumberFormat('@');
  sheet.getRange(targetRow, 1, 1, lastCol).setValues([row]);
  return {ok:true, item:{subSku:subSku, name:name, unit:unit}, row:targetRow};
}


function returnVnpt_listToday_(params) {
  params = params || {};
  const ss = SpreadsheetApp.openById(RETURN_VNPT_CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(RETURN_VNPT_CONFIG.inputSheet);
  if (!sheet) return {ok: false, message: 'Khong tim thay tab: ' + RETURN_VNPT_CONFIG.inputSheet};
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 4 || lastCol < 1) return {ok: true, items: [], total: 0};


  const headers = sheet.getRange(3, 1, 1, lastCol).getDisplayValues()[0];
  const dateCol = returnVnpt_findHeader_(headers, ['Ngay thang', 'Ngày tháng', 'Ngay', 'Ngày']);
  const timeCol = returnVnpt_findHeader_(headers, ['Thoi gian', 'Thời gian', 'Gio', 'Giờ']);
  const maDonCol = returnVnpt_findHeader_(headers, ['Ma don', 'Mã đơn', 'Ma don GHTK', 'Mã đơn GHTK']);
  const nameCol = returnVnpt_findHeader_(headers, ['Ten san pham', 'Tên sản phẩm', 'Ten SKU', 'Tên SKU']);
  const qtyCol = returnVnpt_findHeader_(headers, ['So luong', 'Số lượng', 'SL']);
  const statusCol = returnVnpt_findHeader_(headers, ['Tinh trang', 'Tình trạng']);
  const noteCol = returnVnpt_findHeader_(headers, ['Ghi chu', 'Ghi chú', 'Note']);
  const imageCol = returnVnpt_findHeader_(headers, ['Hinh anh', 'Hình ảnh', 'Image']);
  const userCol = returnVnpt_findHeader_(headers, ['User thao tac', 'User thao tác', 'User_email', 'Email']);
  const idCol = returnVnpt_findHeader_(headers, ['ID']);
  const typeCol = returnVnpt_findHeader_(headers, ['Loại hoàn', 'Loai hoan']);
  const iccidCol = returnVnpt_findHeader_(headers, ['ICCID']);
  const requestedDate = returnVnpt_normalizeDateText_(params.date || Utilities.formatDate(new Date(), RETURN_VNPT_CONFIG.timezone, 'dd/MM/yyyy'));
  const values = sheet.getRange(4, 1, lastRow - 3, lastCol).getDisplayValues();
  const groups = {};


  values.forEach((row, index) => {
    const dateText = returnVnpt_normalizeDateText_(returnVnpt_cell_(row, dateCol));
    if (dateText !== requestedDate) return;
    const maDon = returnVnpt_cell_(row, maDonCol);
    if (!maDon) return;
    const timeText = returnVnpt_cell_(row, timeCol);
    const userEmail = returnVnpt_cell_(row, userCol);
    const key = [dateText, timeText, maDon, userEmail].join('|');
    if (!groups[key]) {
      groups[key] = {
        id: returnVnpt_cell_(row, idCol) || key,
        date: dateText,
        time: timeText,
        maDon: maDon,
        user: userEmail,
        status: 'done',
        qty: 0,
        products: '',
        items: []
      };
    }
    const item = {
      id: returnVnpt_cell_(row, idCol),
      tenSanPham: returnVnpt_cell_(row, nameCol),
      soLuong: returnVnpt_cell_(row, qtyCol),
      iccid: returnVnpt_cell_(row, iccidCol),
      tinhTrang: returnVnpt_cell_(row, statusCol),
      loaiHoan: returnVnpt_cell_(row, typeCol) || 'Hoàn trả',
      ghiChu: returnVnpt_cell_(row, noteCol),
      imageLink: returnVnpt_cell_(row, imageCol),
      row: index + 4
    };
    groups[key].items.push(item);
    groups[key].qty += Number(item.soLuong || 0) || 0;
  });


  const items = Object.keys(groups).map(key => {
    const group = groups[key];
    group.products = group.items.map(item => item.tenSanPham).filter(Boolean).join('; ');
    return group;
  }).sort((a, b) => String(b.time || '').localeCompare(String(a.time || '')));
  return {ok: true, items: items, total: items.length};
}


function returnVnpt_cell_(row, col) {
  return col >= 0 ? String(row[col] || '').trim() : '';
}


function returnVnpt_normalizeDateText_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, RETURN_VNPT_CONFIG.timezone, 'dd/MM/yyyy');
  const text = String(value || '').trim();
  const m = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return ('0' + Number(m[1])).slice(-2) + '/' + ('0' + Number(m[2])).slice(-2) + '/' + m[3];
  return text;
}


function returnVnpt_parseBody_(e) {
  const text = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
  return JSON.parse(text || '{}');
}


function returnVnpt_json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
/** App cũ không gửi loại hoàn: mặc định Hoàn trả. Cho phép app mới gửi Ecom. */
function returnVnpt_loaiHoan_(value) {
  const key=returnVnpt_norm_(value).replace(/\s+/g,' ');
  if (!key || key==='hoan tra') return 'Hoàn trả';
  if (key==='ecom') return 'Ecom';
  throw new Error('Loại hoàn không hợp lệ: '+String(value));
}
function returnVnpt_inputColumns_(headers) {
  const fields={
    id:['ID'],
    dateText:['Ngày tháng','Ngày'],
    timeText:['Thời gian','Giờ'],
    maDon:['Mã đơn','Mã đơn GHTK','Ma don'],
    tenSanPham:['Tên sản phẩm','Tên sản phẩm hoàn về','Ten san pham'],
    soLuong:['Số lượng','SL'],
    iccid:['ICCID'],
    tinhTrang:['Tình trạng','Tinh trang'],
    loaiHoan:['Loại hoàn','Loai hoan'],
    daTaoSv:['Đã tạo SV','Da tao SV'],
    ghiChu:['Ghi chú','Note'],
    imageLink:['Hình ảnh','Ảnh','Image'],
    userEmail:['User thao tác','User email','User_email','Email']
  };
  const result={},used=new Set();
  Object.keys(fields).forEach(field=>{
    const aliases=fields[field].map(returnVnpt_norm_);
    const matches=headers.map((h,i)=>aliases.includes(returnVnpt_norm_(h))?i:-1).filter(i=>i>=0);
    if (matches.length!==1) throw new Error((matches.length?'Trùng':'Thiếu')+' header "'+fields[field][0]+'" tại Hoàn trả.');
    if (used.has(matches[0])) throw new Error('Header trùng chức năng tại Hoàn trả.');
    result[field]=matches[0];used.add(matches[0]);
  });
  return result;
}
function returnVnpt_writeInputRows_(sheet,startRow,rows,columns) {
  const indices=Object.values(columns).sort((a,b)=>a-b),blocks=[];
  const lastRow=startRow+rows.length-1;
  if (lastRow>sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(),lastRow-sheet.getMaxRows());
  for (let i=0;i<indices.length;) {
    let j=i+1;
    while (j<indices.length && indices[j]===indices[j-1]+1) j++;
    const left=indices[i],width=j-i,range=sheet.getRange(startRow,left+1,rows.length,width);
    if (range.getFormulas().some(row=>row.some(Boolean)))
      throw new Error('Không ghi đè công thức trong vùng nhập hoàn.');
    blocks.push({range:range,values:rows.map(row=>row.slice(left,left+width))});
    i=j;
  }
  // ICCID 19-20 số: ghi dạng văn bản để Sheets không đổi thành số và làm tròn.
  if (columns.iccid !== undefined) sheet.getRange(startRow,columns.iccid+1,rows.length,1).setNumberFormat('@');
  // Chỉ ghi 13 cột nhập liệu (có ICCID); giữ bảng thống kê/cột khác ở bên phải hoặc xen giữa.
  blocks.forEach(b=>b.range.setValues(b.values));
}

/* ===== 5.9.27: mã đơn từ link GHTK, đồng bộ SKU/dữ liệu theo mã, sửa/xóa đơn ===== */

/** Quét QR GHTK ra link https://i.ghtk.vn/1598188396 thì chỉ lấy 10 số cuối làm mã đơn. */
function returnVnpt_cleanMaDon_(value) {
  const text = String(value || '').trim();
  if (!/ghtk\.vn/i.test(text)) return text;
  const digits = (text.split(/[?#]/)[0].replace(/\/+$/, '').split('/').pop() || '').replace(/\D/g, '');
  return digits.length >= 10 ? digits.slice(-10) : text;
}

function returnVnpt_withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return {ok: false, message: 'Hệ thống Nhập VNPT đang bận, app sẽ tự thử lại.'};
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function returnVnpt_hash_(value) {
  const text = JSON.stringify(value);
  let a = 2166136261, b = 5381;
  for (let i = 0; i < text.length; i++) {
    a = Math.imul(a ^ text.charCodeAt(i), 16777619);
    b = Math.imul(b, 33) ^ text.charCodeAt(i);
  }
  return (a >>> 0).toString(16) + '-' + (b >>> 0).toString(16);
}

function returnVnpt_syncResult_(records, versions) {
  versions = versions || {};
  return {
    ok: true,
    protocol: 1,
    total: records.length,
    entries: records.map(r => ({key: r.syncKey, version: r.contentVersion, rowNumber: r.rowNumber})),
    records: records.filter(r => versions[r.syncKey] !== r.contentVersion)
  };
}

/** Giống skuSync của PN: trả đủ danh sách mã + phiên bản, chỉ gửi nội dung SKU mới/đổi. Khóa = SUB SKU. */
function returnVnpt_skuSync_(params) {
  const ss = SpreadsheetApp.openById(RETURN_VNPT_CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(RETURN_VNPT_CONFIG.skuSheet);
  if (!sheet) return {ok: false, message: 'Không tìm thấy tab SKU: ' + RETURN_VNPT_CONFIG.skuSheet};
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 4 || lastCol < 1) return returnVnpt_syncResult_([], (params || {}).versions);
  const headers = sheet.getRange(3, 1, 1, lastCol).getDisplayValues()[0];
  const subSkuCol = returnVnpt_findHeader_(headers, ['SUB SKU', 'Sub SKU', 'SKU', 'Ma SKU', 'Mã SKU']);
  const nameCol = returnVnpt_findHeader_(headers, ['Tên sản phẩm', 'Tên sản phẩm hoàn về', 'Ten san pham', 'Ten SKU', 'Tên SKU', 'Product', 'Name']);
  const unitCol = returnVnpt_findHeader_(headers, ['Don vi', 'Đơn vị', 'Unit', 'DVT']);
  if (subSkuCol < 0 || nameCol < 0) return {ok: false, message: 'Tab Thông tin sku cần có cột SUB SKU và Tên SKU.'};
  const values = sheet.getRange(4, 1, lastRow - 3, lastCol).getDisplayValues();
  const records = [], seen = Object.create(null);
  for (let i = 0; i < values.length; i++) {
    const rowNumber = i + 4;
    const subSku = String(values[i][subSkuCol] || '').trim();
    const name = String(values[i][nameCol] || '').trim();
    const unit = unitCol >= 0 ? String(values[i][unitCol] || '').trim() : '';
    if (!subSku && !name) continue;
    if (!subSku) return {ok: false, message: 'Thông tin SKU dòng ' + rowNumber + ' thiếu SUB SKU (' + name + ').'};
    if (!name) return {ok: false, message: 'Thông tin SKU dòng ' + rowNumber + ' thiếu Tên SKU (' + subSku + ').'};
    const key = 'SUBSKU:' + subSku.toLowerCase();
    const content = {subSku, name, unit};
    const signature = JSON.stringify(content);
    if (seen[key]) {
      if (seen[key] === signature) continue;
      return {ok: false, message: 'Thông tin SKU trùng SUB SKU khác nội dung: ' + subSku + ' (dòng ' + rowNumber + ').'};
    }
    seen[key] = signature;
    records.push(Object.assign({syncKey: key, rowNumber, contentVersion: returnVnpt_hash_(content)}, content));
  }
  return returnVnpt_syncResult_(records, (params || {}).versions);
}

/** Dữ liệu tab Hoàn trả cho màn Tìm kiếm VNPT. Khóa = cột ID (dòng nhập tay không có ID thì theo số dòng). */
function returnVnpt_keySync_(params) {
  const ss = SpreadsheetApp.openById(RETURN_VNPT_CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(RETURN_VNPT_CONFIG.inputSheet);
  if (!sheet) return {ok: false, message: 'Không tìm thấy tab: ' + RETURN_VNPT_CONFIG.inputSheet};
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 4 || lastCol < 1) return returnVnpt_syncResult_([], (params || {}).versions);
  const headers = sheet.getRange(3, 1, 1, lastCol).getDisplayValues()[0];
  const c = returnVnpt_inputColumns_(headers);
  const values = sheet.getRange(4, 1, lastRow - 3, lastCol).getDisplayValues();
  const records = [], seen = Object.create(null);
  values.forEach((row, i) => {
    const rowNumber = i + 4;
    const cell = field => String(row[c[field]] || '').trim();
    const content = {
      id: cell('id'),
      date: returnVnpt_normalizeDateText_(cell('dateText')),
      time: cell('timeText'),
      maDon: cell('maDon'),
      tenSanPham: cell('tenSanPham'),
      soLuong: cell('soLuong'),
      iccid: cell('iccid'),
      tinhTrang: cell('tinhTrang'),
      loaiHoan: cell('loaiHoan') || 'Hoàn trả',
      daTaoSv: cell('daTaoSv'),
      ghiChu: cell('ghiChu'),
      imageLink: cell('imageLink'),
      user: cell('userEmail')
    };
    if (!content.maDon && !content.tenSanPham) return;
    let key = content.id ? 'ID:' + content.id.toLowerCase() : 'ROW:' + rowNumber;
    if (seen[key]) key += '#' + rowNumber;
    seen[key] = true;
    records.push(Object.assign({syncKey: key, rowNumber, contentVersion: returnVnpt_hash_(content)}, content));
  });
  return returnVnpt_syncResult_(records, (params || {}).versions);
}

/** Sửa / xóa đơn trong ngày theo cột ID. Không xóa dòng: chỉ xóa nội dung 13 cột nhập liệu để giữ công thức/bảng bên cạnh. */
function returnVnpt_update_(params, mode) {
  params = params || {};
  const editor = String(params.userEmail || params.email || '').trim();
  const editId = returnVnpt_safeId_(params.editId || params.clientId || '');
  const groupIds = (Array.isArray(params.groupIds) ? params.groupIds : []).map(returnVnpt_safeId_).filter(Boolean);
  if (!groupIds.length) return {ok: false, message: 'Thiếu ID đơn cần sửa.'};

  const ss = SpreadsheetApp.openById(RETURN_VNPT_CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(RETURN_VNPT_CONFIG.inputSheet);
  if (!sheet) return {ok: false, message: 'Không tìm thấy tab: ' + RETURN_VNPT_CONFIG.inputSheet};
  if (editId && returnVnpt_logHasEdit_(ss, editId)) return {ok: true, id: editId, duplicate: true};

  const width = sheet.getLastColumn();
  const headers = sheet.getRange(3, 1, 1, width).getDisplayValues()[0];
  const columns = returnVnpt_inputColumns_(headers);
  const lastRow = sheet.getLastRow();
  const rowById = {};
  if (lastRow >= 4) {
    sheet.getRange(4, columns.id + 1, lastRow - 3, 1).getDisplayValues().forEach((v, i) => {
      const key = String(v[0] || '').trim();
      if (key && !rowById[key]) rowById[key] = i + 4;
    });
  }
  const existingIds = groupIds.filter(id => rowById[id]);
  if (!existingIds.length) return {ok: false, message: 'Không tìm thấy đơn trên sheet (có thể đã bị xóa).'};

  const now = new Date();
  const today = Utilities.formatDate(now, RETURN_VNPT_CONFIG.timezone, 'dd/MM/yyyy');
  const before = {};
  existingIds.forEach(id => {
    const row = sheet.getRange(rowById[id], 1, 1, width).getDisplayValues()[0];
    const data = {};
    Object.keys(columns).forEach(field => { data[field] = String(row[columns[field]] || '').trim(); });
    before[id] = data;
  });
  for (let i = 0; i < existingIds.length; i++) {
    if (returnVnpt_normalizeDateText_(before[existingIds[i]].dateText) !== today) return {ok: false, message: 'Chỉ được sửa/xóa đơn nhập trong ngày.'};
  }
  const base = before[existingIds[0]];
  const blankRow = () => headers.map(() => '');
  const clearRow = id => returnVnpt_writeInputRows_(sheet, rowById[id], [blankRow()], columns);

  if (mode === 'delete' || params.deleteOrder === true) {
    existingIds.forEach(clearRow);
    returnVnpt_refreshSummaryIfAvailable_();
    const mailError = returnVnpt_logEdit_(ss, now, editor, 'Xóa đơn', base.maDon, editId, existingIds, before, null);
    return {ok: true, id: editId, deleted: existingIds.length, mailError};
  }

  const maDon = returnVnpt_cleanMaDon_(params.maDon || base.maDon);
  if (!maDon) return {ok: false, message: 'Thiếu Mã đơn'};
  const loaiHoan = returnVnpt_loaiHoan_(params.loaiHoan || base.loaiHoan || '');
  const items = returnVnpt_itemsFromParams_(params);
  const deleteIds = (Array.isArray(params.deleteIds) ? params.deleteIds : []).map(returnVnpt_safeId_).filter(id => existingIds.indexOf(id) >= 0);
  if (!items.length) return {ok: false, message: 'Đơn phải còn ít nhất 1 SKU. Muốn bỏ cả đơn hãy dùng Xóa đơn.'};
  const plans = items.map((item, index) => {
    const id = returnVnpt_safeId_(item.clientItemId) || (editId + '-' + (index + 1));
    return {item, index, id, exists: !!rowById[id]};
  });
  for (let i = 0; i < plans.length; i++) {
    const p = plans[i], label = 'SKU ' + (i + 1);
    if (!p.item.tenSanPham) return {ok: false, message: label + ' thiếu Tên sản phẩm'};
    if (!p.item.soLuong) return {ok: false, message: label + ' thiếu Số lượng'};
    if (!p.item.tinhTrang) return {ok: false, message: label + ' thiếu Tình trạng'};
    if (!p.exists && !p.item.files.length) return {ok: false, message: label + ' (SKU mới) thiếu hình ảnh'};
    const iccidError = returnVnpt_iccidError_(p.item, label);
    if (iccidError) return {ok: false, message: iccidError};
  }

  const rowFor = data => {
    const row = blankRow();
    returnVnpt_applyFixedInputLayout_(row, data, headers);
    return row;
  };
  const after = {};
  plans.forEach(p => {
    const old = p.exists ? (before[p.id] || {}) : {};
    const imageLink = p.item.files.length ? returnVnpt_uploadImages_(maDon, p.id, p.item.files, now, p.index) : (old.imageLink || '');
    const data = {
      id: p.id,
      dateText: old.dateText || base.dateText,
      timeText: old.timeText || base.timeText,
      maDon,
      tenSanPham: p.item.tenSanPham,
      soLuong: p.item.soLuong,
      iccid: p.item.iccid,
      tinhTrang: p.item.tinhTrang,
      loaiHoan,
      daTaoSv: old.daTaoSv || '',
      ghiChu: p.item.ghiChu,
      imageLink,
      userEmail: old.userEmail || base.userEmail
    };
    after[p.id] = data;
    p.row = rowFor(data);
  });
  // SKU cũ không gửi lên và không yêu cầu xóa: giữ nguyên, chỉ đồng bộ Mã đơn/Loại hoàn để không bị tách đơn.
  existingIds.forEach(id => {
    if (after[id] || deleteIds.indexOf(id) >= 0) return;
    const data = Object.assign({}, before[id], {maDon, loaiHoan});
    after[id] = data;
    returnVnpt_writeInputRows_(sheet, rowById[id], [rowFor(data)], columns);
  });
  plans.filter(p => p.exists).forEach(p => returnVnpt_writeInputRows_(sheet, rowById[p.id], [p.row], columns));
  deleteIds.filter(id => !after[id]).forEach(clearRow);
  const added = plans.filter(p => !p.exists);
  if (added.length) {
    const targetRow = returnVnpt_findFirstInputRow_(sheet, added.length, headers);
    returnVnpt_writeInputRows_(sheet, targetRow, added.map(p => p.row), columns);
  }
  returnVnpt_refreshSummaryIfAvailable_();
  const mailError = returnVnpt_logEdit_(ss, now, editor, 'Sửa đơn', maDon, editId, existingIds.concat(added.map(p => p.id)), before, after, deleteIds);
  return {ok: true, id: editId, updated: plans.length - added.length, added: added.length, deleted: deleteIds.length, mailError};
}

function returnVnpt_logSheet_(ss) {
  let sheet = ss.getSheetByName(RETURN_VNPT_CONFIG.logSheet);
  if (!sheet) sheet = ss.insertSheet(RETURN_VNPT_CONFIG.logSheet);
  if (sheet.getLastRow() < 1) sheet.appendRow(['Thời gian', 'User thao tác', 'Thao tác', 'Mã đơn', 'Edit ID', 'ID dòng', 'Trước khi sửa', 'Sau khi sửa', 'SKU bị xóa']);
  return sheet;
}

function returnVnpt_logHasEdit_(ss, editId) {
  const sheet = ss.getSheetByName(RETURN_VNPT_CONFIG.logSheet);
  if (!sheet || sheet.getLastRow() < 2) return false;
  return sheet.getRange(2, 5, sheet.getLastRow() - 1, 1).getDisplayValues().some(v => String(v[0] || '').trim() === editId);
}

function returnVnpt_summaryText_(map) {
  if (!map) return '';
  return Object.keys(map).map(id => {
    const d = map[id] || {};
    return id + ': ' + [d.maDon, d.loaiHoan, d.tenSanPham, 'SL ' + d.soLuong, d.iccid ? 'ICCID ' + d.iccid : '', d.tinhTrang, d.ghiChu].filter(x => x !== undefined && x !== '').join(' | ');
  }).join('\n');
}

/** Ghi log sửa/xóa vào tab Log sửa VNPT và gửi mail báo admin. Lỗi mail chỉ trả về cho app biết, không làm hỏng lượt lưu. */
function returnVnpt_logEdit_(ss, now, editor, action, maDon, editId, ids, before, after, deleteIds) {
  const stamp = Utilities.formatDate(now, RETURN_VNPT_CONFIG.timezone, 'dd/MM/yyyy HH:mm:ss');
  const beforeText = returnVnpt_summaryText_(before);
  const afterText = returnVnpt_summaryText_(after);
  const deletedText = (deleteIds || []).join(', ');
  returnVnpt_logSheet_(ss).appendRow([stamp, editor, action, maDon, editId, ids.join(', '), beforeText, afterText, deletedText]);
  if (!RETURN_VNPT_CONFIG.alertEmail) return '';
  try {
    MailApp.sendEmail({
      to: RETURN_VNPT_CONFIG.alertEmail,
      subject: '[Nhập VNPT] ' + action + ' ' + maDon + ' - ' + (editor || 'không rõ user'),
      body: [
        'Thời gian: ' + stamp,
        'User thao tác: ' + editor,
        'Thao tác: ' + action,
        'Mã đơn: ' + maDon,
        'Edit ID: ' + editId,
        '',
        'Trước khi sửa:', beforeText || '-',
        '',
        'Sau khi sửa:', afterText || '-',
        deletedText ? '\nSKU bị xóa: ' + deletedText : ''
      ].join('\n')
    });
    return '';
  } catch (err) {
    return String(err && err.message ? err.message : err);
  }
}

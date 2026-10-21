/** API tương thích Loại hoàn — app cũ mặc định Hoàn trả; giữ action/payload/callback. */
const RETURN_TH_CONFIG = {
  spreadsheetId: '1TVck35ibO3zJ6lkKpHNyXckgLyjvUxEoLbMa-6kt-Pk',
  inputSheet: 'Hàng hoàn',
  skuSheet: 'Thông tin SKU',
  imageFolderName: 'HangHoanTH_Images',
  imageRootFolderId: '',
  logSheet: 'Log sửa TH',
  alertEmail: 'khiempham5209@gmail.com',
  timezone: 'Asia/Ho_Chi_Minh'
};


function returnTh_testAuth() {
  const root = RETURN_TH_CONFIG.imageRootFolderId
    ? DriveApp.getFolderById(RETURN_TH_CONFIG.imageRootFolderId)
    : DriveApp.getRootFolder();
  const folder = returnTh_getOrCreateFolder_(root, RETURN_TH_CONFIG.imageFolderName);
  const file = folder.createFile(Utilities.newBlob('ok', 'text/plain', 'return-th-auth-test.txt'));
  file.setTrashed(true);
  Logger.log('Drive OK: ' + folder.getName());
  return 'Drive OK: ' + folder.getName();
}


function doGet(e) {
  return returnTh_json_({
    ok: true,
    service: 'NhapHoanTH',
    actions: ['listSku', 'skuSync', 'keySync', 'createSku', 'saveReturnTh', 'updateReturnTh', 'deleteReturnTh', 'listToday']
  });
}


function doPost(e) {
  try {
    const payload = returnTh_parseBody_(e);
    const action = String(payload.action || '').trim();
    const params = payload.params || {};
    if (action === 'listSku') return returnTh_json_(returnTh_listSku_());
    if (action === 'skuSync') return returnTh_json_(returnTh_skuSync_(params));
    if (action === 'keySync') return returnTh_json_(returnTh_keySync_(params));
    if (action === 'listToday') return returnTh_json_(returnTh_listToday_(params));
    const writers = {
      createSku: () => returnTh_createSku_(params),
      saveReturnTh: () => returnTh_save_(params),
      updateReturnTh: () => returnTh_update_(params, 'update'),
      deleteReturnTh: () => returnTh_update_(params, 'delete')
    };
    if (writers[action]) return returnTh_json_(returnTh_withLock_(writers[action]));
    return returnTh_json_({ok: false, message: 'Action khong hop le: ' + action});
  } catch (err) {
    return returnTh_json_({ok: false, message: String(err && err.message ? err.message : err)});
  }
}




function returnTh_save_(params) {
  const maDon = returnTh_cleanMaDon_(
    params.maDon ||
    params.ma_don ||
    params.orderCode ||
    params.order_code ||
    params['Mã đơn'] ||
    params['Ma don'] ||
    params['ma don'] ||
    ''
  ).trim();
  const items = returnTh_itemsFromParams_(params);
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
  }


  const ss = SpreadsheetApp.openById(RETURN_TH_CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(RETURN_TH_CONFIG.inputSheet);
  if (!sheet) return {ok: false, message: 'Khong tim thay tab: ' + RETURN_TH_CONFIG.inputSheet};


  // Kiểm tra header trước khi tải ảnh hoặc ghi dữ liệu.
  const headers = sheet.getRange(3,1,1,sheet.getLastColumn()).getDisplayValues()[0];
  const inputColumns = returnTh_inputColumns_(headers);
  const now = new Date();
  const dateText = Utilities.formatDate(now, RETURN_TH_CONFIG.timezone, 'dd/MM/yyyy');
  const timeText = Utilities.formatDate(now, RETURN_TH_CONFIG.timezone, 'HH:mm:ss');
  const clientBatchId = String(params.clientBatchId || params.idempotencyKey || params.clientId || '').trim();
  const batchId = returnTh_safeId_(clientBatchId) || ('TH-' + Utilities.formatDate(now, RETURN_TH_CONFIG.timezone, 'yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().slice(0, 8));
  const rowIds = items.map((item, index) => returnTh_safeId_(item.clientItemId) || (items.length === 1 ? batchId : batchId + '-' + (index + 1)));
  const existingIds = returnTh_existingIds_(sheet, rowIds);
  const pendingItems = items.map((item, index) => ({item, index, rowId: rowIds[index]})).filter(entry => !existingIds[entry.rowId]);
  if (!pendingItems.length) {
    return {ok: true, id: batchId, duplicate: true, row: 0, rows: 0};
  }
  const imageLinksByIndex = {};
  pendingItems.forEach(entry => {
    imageLinksByIndex[entry.index] = returnTh_uploadImages_(maDon, entry.rowId, entry.item.files, now, entry.index);
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
      tinhTrang: item.tinhTrang,
      loaiHoan: item.loaiHoan,
      daTaoSv: '',
      ghiChu: item.ghiChu,
      imageLink: imageLinksByIndex[index] || '',
      userEmail
    };
    const row = headers.map(header => returnTh_valueForHeader_(header, rowData));
    returnTh_applyFixedInputLayout_(row, rowData, headers);
    return row;
  });


  const targetRow = returnTh_findFirstInputRow_(sheet, rows.length, headers);
  returnTh_writeInputRows_(sheet,targetRow,rows,inputColumns);
  returnTh_refreshSummaryIfAvailable_();


  return {ok: true, id: batchId, imageLinks: pendingItems.map(entry => imageLinksByIndex[entry.index] || ''), imageLink: imageLinksByIndex[pendingItems[0].index] || '', row: targetRow, rows: rows.length};
}


function returnTh_itemsFromParams_(params) {
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
    loaiHoan: returnTh_loaiHoan_((item && (item.loaiHoan || item['Loại hoàn'])) || params.loaiHoan || params['Loại hoàn'] || ''),
    ghiChu: String((item && (item.ghiChu || item.note)) || '').trim(),
    files: returnTh_itemFiles_(item)
  })).filter(item => item.tenSanPham || item.soLuong || item.tinhTrang || item.ghiChu);
}


function returnTh_itemFiles_(item) {
  if (!item) return [];
  if (Array.isArray(item.files) && item.files.length) return item.files;
  if (Array.isArray(item.images) && item.images.length) return item.images;
  if (item.image) return [item.image];
  return [];
}


function returnTh_valueForHeader_(header, data) {
  const key = returnTh_norm_(header);
  if (!key) return '';
  if (key === 'id') return data.id;
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


function returnTh_applyFixedInputLayout_(row,data,headers) {
  // Giữ tên helper cũ, nhưng không còn ghi đè theo số cột cố định.
  if (!headers) throw new Error('Thiếu header khi chuẩn bị dữ liệu hoàn.');
  const columns=returnTh_inputColumns_(headers);
  Object.keys(columns).forEach(field=>{row[columns[field]]=data[field];});
}

function returnTh_refreshSummaryIfAvailable_() {
  try {
    if (typeof refreshTongHopAll === 'function') refreshTongHopAll();
  } catch (err) {
    console.log('refreshTongHopAll skipped: ' + (err && err.message ? err.message : err));
  }
}


function returnTh_findFirstInputRow_(sheet,rowCount,headers) {
  const startRow=4,needed=Math.max(1,Number(rowCount)||1);
  headers=headers || sheet.getRange(3,1,1,sheet.getLastColumn()).getDisplayValues()[0];
  const columns=returnTh_inputColumns_(headers),width=headers.length;
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

function returnTh_existingIds_(sheet,ids) {
  const wanted=new Set((ids||[]).map(v=>String(v||'').trim()).filter(Boolean)),found={};
  const lastRow=sheet.getLastRow();
  if (!wanted.size || lastRow<4) return found;
  const headers=sheet.getRange(3,1,1,sheet.getLastColumn()).getDisplayValues()[0];
  const idCol=returnTh_inputColumns_(headers).id;
  sheet.getRange(4,idCol+1,lastRow-3,1).getDisplayValues().flat().forEach(value=>{
    const key=String(value||'').trim();if(wanted.has(key))found[key]=true;
  });
  return found;
}

function returnTh_safeId_(value) {
  return String(value || '').trim().replace(/[^\w.-]/g, '_').slice(0, 120);
}


function returnTh_uploadImages_(maDon, id, files, now, skuIndex) {
  const parent = RETURN_TH_CONFIG.imageRootFolderId
    ? DriveApp.getFolderById(RETURN_TH_CONFIG.imageRootFolderId)
    : DriveApp.getRootFolder();
  const root = returnTh_getOrCreateFolder_(parent, RETURN_TH_CONFIG.imageFolderName);
  const safeOrder = returnTh_safeName_((maDon || id) + (skuIndex != null ? '_SKU' + (Number(skuIndex) + 1) : ''));
  if (files.length === 1) {
    const file = returnTh_createImageFile_(root, files[0], safeOrder, now, 0);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/uc?id=' + file.getId();
  }


  const folder = returnTh_getOrCreateFolder_(root, safeOrder);
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  files.forEach((filePayload, index) => returnTh_createImageFile_(folder, filePayload, safeOrder, now, index));
  return 'https://drive.google.com/drive/folders/' + folder.getId();
}


function returnTh_createImageFile_(folder, filePayload, safeOrder, now, index) {
  const mimeType = String(filePayload.mimeType || filePayload.type || 'image/jpeg');
  const base64 = returnTh_extractBase64_(filePayload);
  if (!base64) throw new Error('Anh khong co du lieu base64');
  const bytes = Utilities.base64Decode(base64);
  const ext = mimeType.indexOf('png') >= 0 ? 'png' : 'jpg';
  const stamp = Utilities.formatDate(now, RETURN_TH_CONFIG.timezone, 'yyyyMMdd_HHmmss');
  const name = safeOrder + '_' + stamp + '_' + index + '.' + ext;
  const blob = Utilities.newBlob(bytes, mimeType, name);
  return folder.createFile(blob);
}


function returnTh_extractBase64_(filePayload) {
  const raw = String(filePayload.base64 || filePayload.data || '');
  const match = raw.match(/^data:[^;]+;base64,(.*)$/);
  return match ? match[1] : raw;
}


function returnTh_getOrCreateFolder_(parent, name) {
  const iter = parent.getFoldersByName(name);
  if (iter.hasNext()) return iter.next();
  return parent.createFolder(name);
}


function returnTh_findHeader_(headers, names) {
  const normalized = headers.map(returnTh_norm_);
  for (let i = 0; i < names.length; i++) {
    const idx = normalized.indexOf(returnTh_norm_(names[i]));
    if (idx >= 0) return idx;
  }
  return -1;
}


function returnTh_norm_(value) {
  return String(value || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
}


function returnTh_safeName_(value) {
  return String(value || 'no-code').replace(/[\\/:*?"<>|#%{}~&]/g, '_').slice(0, 80);
}


// Structured SKU source used by the shared searchable picker in web/APK.
function returnTh_listSku_() {
  const ss = SpreadsheetApp.openById(RETURN_TH_CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(RETURN_TH_CONFIG.skuSheet);
  if (!sheet) return {ok: false, message: 'Khong tim thay tab SKU: ' + RETURN_TH_CONFIG.skuSheet};
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 4 || lastCol < 1) return {ok: true, items: []};
  const headers = sheet.getRange(3, 1, 1, lastCol).getDisplayValues()[0];
  const subSkuCol = returnTh_findHeader_(headers, ['SUB SKU', 'Sub SKU', 'SKU', 'Ma SKU', 'Mã SKU']);
  const nameCol = returnTh_findHeader_(headers, ['Tên sản phẩm', 'Tên sản phẩm hoàn về', 'Ten san pham', 'Ten SKU', 'Tên SKU', 'Product', 'Name']);
  const unitCol = returnTh_findHeader_(headers, ['Don vi', 'Đơn vị', 'Unit', 'DVT']);
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


function returnTh_createSku_(params) {
  params = params || {};
  const subSku = String(params.subSku || params.sub_sku || params.sku || '').trim();
  const name = String(params.name || params.tenSku || params.tenSanPham || '').trim();
  const unit = String(params.unit || params.donVi || '').trim();
  if (!subSku) return {ok: false, message: 'Thiếu SUB SKU'};
  if (!name) return {ok: false, message: 'Thiếu Tên SKU'};
  if (!unit) return {ok: false, message: 'Thiếu Đơn vị'};
  const ss = SpreadsheetApp.openById(RETURN_TH_CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(RETURN_TH_CONFIG.skuSheet);
  if (!sheet) return {ok: false, message: 'Không tìm thấy tab SKU: ' + RETURN_TH_CONFIG.skuSheet};
  const lastCol = Math.max(3, sheet.getLastColumn());
  const headers = sheet.getRange(3, 1, 1, lastCol).getDisplayValues()[0];
  const subSkuCol = returnTh_findHeader_(headers, ['SUB SKU', 'Sub SKU', 'SKU', 'Ma SKU', 'Mã SKU']);
  const nameCol = returnTh_findHeader_(headers, ['Ten SKU', 'Tên SKU', 'Tên sản phẩm', 'Ten san pham', 'Product', 'Name']);
  const unitCol = returnTh_findHeader_(headers, ['Don vi', 'Đơn vị', 'Unit', 'DVT']);
  if (subSkuCol < 0 || nameCol < 0 || unitCol < 0) return {ok:false, message:'Tab Thông tin SKU cần đủ cột SUB SKU, Tên SKU, Đơn vị.'};
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
  sheet.getRange(targetRow, 1, 1, lastCol).setValues([row]);
  return {ok:true, item:{subSku:subSku, name:name, unit:unit}, row:targetRow};
}


function returnTh_listToday_(params) {
  params = params || {};
  const ss = SpreadsheetApp.openById(RETURN_TH_CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(RETURN_TH_CONFIG.inputSheet);
  if (!sheet) return {ok: false, message: 'Khong tim thay tab: ' + RETURN_TH_CONFIG.inputSheet};
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 4 || lastCol < 1) return {ok: true, items: [], total: 0};


  const headers = sheet.getRange(3, 1, 1, lastCol).getDisplayValues()[0];
  const dateCol = returnTh_findHeader_(headers, ['Ngay thang', 'Ngày tháng', 'Ngay', 'Ngày']);
  const timeCol = returnTh_findHeader_(headers, ['Thoi gian', 'Thời gian', 'Gio', 'Giờ']);
  const maDonCol = returnTh_findHeader_(headers, ['Ma don', 'Mã đơn', 'Ma don GHTK', 'Mã đơn GHTK']);
  const nameCol = returnTh_findHeader_(headers, ['Ten san pham', 'Tên sản phẩm', 'Ten SKU', 'Tên SKU']);
  const qtyCol = returnTh_findHeader_(headers, ['So luong', 'Số lượng', 'SL']);
  const statusCol = returnTh_findHeader_(headers, ['Tinh trang', 'Tình trạng']);
  const noteCol = returnTh_findHeader_(headers, ['Ghi chu', 'Ghi chú', 'Note']);
  const imageCol = returnTh_findHeader_(headers, ['Hinh anh', 'Hình ảnh', 'Image']);
  const userCol = returnTh_findHeader_(headers, ['User thao tac', 'User thao tác', 'User_email', 'Email']);
  const idCol = returnTh_findHeader_(headers, ['ID']);
  const typeCol = returnTh_findHeader_(headers, ['Loại hoàn', 'Loai hoan']);
  const requestedDate = returnTh_normalizeDateText_(params.date || Utilities.formatDate(new Date(), RETURN_TH_CONFIG.timezone, 'dd/MM/yyyy'));
  const values = sheet.getRange(4, 1, lastRow - 3, lastCol).getDisplayValues();
  const groups = {};


  values.forEach((row, index) => {
    const dateText = returnTh_normalizeDateText_(returnTh_cell_(row, dateCol));
    if (dateText !== requestedDate) return;
    const maDon = returnTh_cell_(row, maDonCol);
    if (!maDon) return;
    const timeText = returnTh_cell_(row, timeCol);
    const userEmail = returnTh_cell_(row, userCol);
    const key = [dateText, timeText, maDon, userEmail].join('|');
    if (!groups[key]) {
      groups[key] = {
        id: returnTh_cell_(row, idCol) || key,
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
      id: returnTh_cell_(row, idCol),
      tenSanPham: returnTh_cell_(row, nameCol),
      soLuong: returnTh_cell_(row, qtyCol),
      tinhTrang: returnTh_cell_(row, statusCol),
      loaiHoan: returnTh_cell_(row, typeCol) || 'Hoàn trả',
      ghiChu: returnTh_cell_(row, noteCol),
      imageLink: returnTh_cell_(row, imageCol),
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


function returnTh_cell_(row, col) {
  return col >= 0 ? String(row[col] || '').trim() : '';
}


function returnTh_normalizeDateText_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, RETURN_TH_CONFIG.timezone, 'dd/MM/yyyy');
  const text = String(value || '').trim();
  const m = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return ('0' + Number(m[1])).slice(-2) + '/' + ('0' + Number(m[2])).slice(-2) + '/' + m[3];
  return text;
}


function returnTh_parseBody_(e) {
  const text = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
  return JSON.parse(text || '{}');
}


function returnTh_json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
/** App cũ không gửi loại hoàn: mặc định Hoàn trả. Cho phép app mới gửi Ecom. */
function returnTh_loaiHoan_(value) {
  const key=returnTh_norm_(value).replace(/\s+/g,' ');
  if (!key || key==='hoan tra') return 'Hoàn trả';
  if (key==='ecom') return 'Ecom';
  throw new Error('Loại hoàn không hợp lệ: '+String(value));
}
function returnTh_inputColumns_(headers) {
  const fields={
    id:['ID'],
    dateText:['Ngày tháng','Ngày'],
    timeText:['Thời gian','Giờ'],
    maDon:['Mã đơn','Mã đơn GHTK','Ma don'],
    tenSanPham:['Tên sản phẩm','Tên sản phẩm hoàn về','Ten san pham'],
    soLuong:['Số lượng','SL'],
    tinhTrang:['Tình trạng','Tinh trang'],
    loaiHoan:['Loại hoàn','Loai hoan'],
    daTaoSv:['Đã tạo SV','Da tao SV'],
    ghiChu:['Ghi chú','Note'],
    imageLink:['Hình ảnh','Ảnh','Image'],
    userEmail:['User thao tác','User email','User_email','Email']
  };
  const result={},used=new Set();
  Object.keys(fields).forEach(field=>{
    const aliases=fields[field].map(returnTh_norm_);
    const matches=headers.map((h,i)=>aliases.includes(returnTh_norm_(h))?i:-1).filter(i=>i>=0);
    if (matches.length!==1) throw new Error((matches.length?'Trùng':'Thiếu')+' header "'+fields[field][0]+'" tại Hàng hoàn.');
    if (used.has(matches[0])) throw new Error('Header trùng chức năng tại Hàng hoàn.');
    result[field]=matches[0];used.add(matches[0]);
  });
  return result;
}
function returnTh_writeInputRows_(sheet,startRow,rows,columns) {
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
  // Chỉ ghi 12 cột nhập liệu; giữ bảng thống kê/cột khác ở bên phải hoặc xen giữa.
  blocks.forEach(b=>b.range.setValues(b.values));
}

/* ===== 5.9.27: mã đơn từ link GHTK, đồng bộ SKU/dữ liệu theo mã, sửa/xóa đơn ===== */

/** Quét QR GHTK ra link https://i.ghtk.vn/1598188396 thì chỉ lấy 10 số cuối làm mã đơn. */
function returnTh_cleanMaDon_(value) {
  const text = String(value || '').trim();
  if (!/ghtk\.vn/i.test(text)) return text;
  const digits = (text.split(/[?#]/)[0].replace(/\/+$/, '').split('/').pop() || '').replace(/\D/g, '');
  return digits.length >= 10 ? digits.slice(-10) : text;
}

function returnTh_withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return {ok: false, message: 'Hệ thống Nhập TH đang bận, app sẽ tự thử lại.'};
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function returnTh_hash_(value) {
  const text = JSON.stringify(value);
  let a = 2166136261, b = 5381;
  for (let i = 0; i < text.length; i++) {
    a = Math.imul(a ^ text.charCodeAt(i), 16777619);
    b = Math.imul(b, 33) ^ text.charCodeAt(i);
  }
  return (a >>> 0).toString(16) + '-' + (b >>> 0).toString(16);
}

function returnTh_syncResult_(records, versions) {
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
function returnTh_skuSync_(params) {
  const ss = SpreadsheetApp.openById(RETURN_TH_CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(RETURN_TH_CONFIG.skuSheet);
  if (!sheet) return {ok: false, message: 'Không tìm thấy tab SKU: ' + RETURN_TH_CONFIG.skuSheet};
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 4 || lastCol < 1) return returnTh_syncResult_([], (params || {}).versions);
  const headers = sheet.getRange(3, 1, 1, lastCol).getDisplayValues()[0];
  const subSkuCol = returnTh_findHeader_(headers, ['SUB SKU', 'Sub SKU', 'SKU', 'Ma SKU', 'Mã SKU']);
  const nameCol = returnTh_findHeader_(headers, ['Tên sản phẩm', 'Tên sản phẩm hoàn về', 'Ten san pham', 'Ten SKU', 'Tên SKU', 'Product', 'Name']);
  const unitCol = returnTh_findHeader_(headers, ['Don vi', 'Đơn vị', 'Unit', 'DVT']);
  if (subSkuCol < 0 || nameCol < 0) return {ok: false, message: 'Tab Thông tin SKU cần có cột SUB SKU và Tên SKU.'};
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
    records.push(Object.assign({syncKey: key, rowNumber, contentVersion: returnTh_hash_(content)}, content));
  }
  return returnTh_syncResult_(records, (params || {}).versions);
}

/** Dữ liệu tab Hàng hoàn cho màn Tìm kiếm TH. Khóa = cột ID (dòng nhập tay không có ID thì theo số dòng). */
function returnTh_keySync_(params) {
  const ss = SpreadsheetApp.openById(RETURN_TH_CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(RETURN_TH_CONFIG.inputSheet);
  if (!sheet) return {ok: false, message: 'Không tìm thấy tab: ' + RETURN_TH_CONFIG.inputSheet};
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 4 || lastCol < 1) return returnTh_syncResult_([], (params || {}).versions);
  const headers = sheet.getRange(3, 1, 1, lastCol).getDisplayValues()[0];
  const c = returnTh_inputColumns_(headers);
  const values = sheet.getRange(4, 1, lastRow - 3, lastCol).getDisplayValues();
  const records = [], seen = Object.create(null);
  values.forEach((row, i) => {
    const rowNumber = i + 4;
    const cell = field => String(row[c[field]] || '').trim();
    const content = {
      id: cell('id'),
      date: returnTh_normalizeDateText_(cell('dateText')),
      time: cell('timeText'),
      maDon: cell('maDon'),
      tenSanPham: cell('tenSanPham'),
      soLuong: cell('soLuong'),
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
    records.push(Object.assign({syncKey: key, rowNumber, contentVersion: returnTh_hash_(content)}, content));
  });
  return returnTh_syncResult_(records, (params || {}).versions);
}

/** Sửa / xóa đơn trong ngày theo cột ID. Không xóa dòng: chỉ xóa nội dung 12 cột nhập liệu để giữ công thức/bảng bên cạnh. */
function returnTh_update_(params, mode) {
  params = params || {};
  const editor = String(params.userEmail || params.email || '').trim();
  const editId = returnTh_safeId_(params.editId || params.clientId || '');
  const groupIds = (Array.isArray(params.groupIds) ? params.groupIds : []).map(returnTh_safeId_).filter(Boolean);
  if (!groupIds.length) return {ok: false, message: 'Thiếu ID đơn cần sửa.'};

  const ss = SpreadsheetApp.openById(RETURN_TH_CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(RETURN_TH_CONFIG.inputSheet);
  if (!sheet) return {ok: false, message: 'Không tìm thấy tab: ' + RETURN_TH_CONFIG.inputSheet};
  if (editId && returnTh_logHasEdit_(ss, editId)) return {ok: true, id: editId, duplicate: true};

  const width = sheet.getLastColumn();
  const headers = sheet.getRange(3, 1, 1, width).getDisplayValues()[0];
  const columns = returnTh_inputColumns_(headers);
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
  const today = Utilities.formatDate(now, RETURN_TH_CONFIG.timezone, 'dd/MM/yyyy');
  const before = {};
  existingIds.forEach(id => {
    const row = sheet.getRange(rowById[id], 1, 1, width).getDisplayValues()[0];
    const data = {};
    Object.keys(columns).forEach(field => { data[field] = String(row[columns[field]] || '').trim(); });
    before[id] = data;
  });
  for (let i = 0; i < existingIds.length; i++) {
    if (returnTh_normalizeDateText_(before[existingIds[i]].dateText) !== today) return {ok: false, message: 'Chỉ được sửa/xóa đơn nhập trong ngày.'};
  }
  const base = before[existingIds[0]];
  const blankRow = () => headers.map(() => '');
  const clearRow = id => returnTh_writeInputRows_(sheet, rowById[id], [blankRow()], columns);

  if (mode === 'delete' || params.deleteOrder === true) {
    existingIds.forEach(clearRow);
    returnTh_refreshSummaryIfAvailable_();
    const mailError = returnTh_logEdit_(ss, now, editor, 'Xóa đơn', base.maDon, editId, existingIds, before, null);
    return {ok: true, id: editId, deleted: existingIds.length, mailError};
  }

  const maDon = returnTh_cleanMaDon_(params.maDon || base.maDon);
  if (!maDon) return {ok: false, message: 'Thiếu Mã đơn'};
  const loaiHoan = returnTh_loaiHoan_(params.loaiHoan || base.loaiHoan || '');
  const items = returnTh_itemsFromParams_(params);
  const deleteIds = (Array.isArray(params.deleteIds) ? params.deleteIds : []).map(returnTh_safeId_).filter(id => existingIds.indexOf(id) >= 0);
  if (!items.length) return {ok: false, message: 'Đơn phải còn ít nhất 1 SKU. Muốn bỏ cả đơn hãy dùng Xóa đơn.'};
  const plans = items.map((item, index) => {
    const id = returnTh_safeId_(item.clientItemId) || (editId + '-' + (index + 1));
    return {item, index, id, exists: !!rowById[id]};
  });
  for (let i = 0; i < plans.length; i++) {
    const p = plans[i], label = 'SKU ' + (i + 1);
    if (!p.item.tenSanPham) return {ok: false, message: label + ' thiếu Tên sản phẩm'};
    if (!p.item.soLuong) return {ok: false, message: label + ' thiếu Số lượng'};
    if (!p.item.tinhTrang) return {ok: false, message: label + ' thiếu Tình trạng'};
    if (!p.exists && !p.item.files.length) return {ok: false, message: label + ' (SKU mới) thiếu hình ảnh'};
  }

  const rowFor = data => {
    const row = blankRow();
    returnTh_applyFixedInputLayout_(row, data, headers);
    return row;
  };
  const after = {};
  plans.forEach(p => {
    const old = p.exists ? (before[p.id] || {}) : {};
    const imageLink = p.item.files.length ? returnTh_uploadImages_(maDon, p.id, p.item.files, now, p.index) : (old.imageLink || '');
    const data = {
      id: p.id,
      dateText: old.dateText || base.dateText,
      timeText: old.timeText || base.timeText,
      maDon,
      tenSanPham: p.item.tenSanPham,
      soLuong: p.item.soLuong,
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
    returnTh_writeInputRows_(sheet, rowById[id], [rowFor(data)], columns);
  });
  plans.filter(p => p.exists).forEach(p => returnTh_writeInputRows_(sheet, rowById[p.id], [p.row], columns));
  deleteIds.filter(id => !after[id]).forEach(clearRow);
  const added = plans.filter(p => !p.exists);
  if (added.length) {
    const targetRow = returnTh_findFirstInputRow_(sheet, added.length, headers);
    returnTh_writeInputRows_(sheet, targetRow, added.map(p => p.row), columns);
  }
  returnTh_refreshSummaryIfAvailable_();
  const mailError = returnTh_logEdit_(ss, now, editor, 'Sửa đơn', maDon, editId, existingIds.concat(added.map(p => p.id)), before, after, deleteIds);
  return {ok: true, id: editId, updated: plans.length - added.length, added: added.length, deleted: deleteIds.length, mailError};
}

function returnTh_logSheet_(ss) {
  let sheet = ss.getSheetByName(RETURN_TH_CONFIG.logSheet);
  if (!sheet) sheet = ss.insertSheet(RETURN_TH_CONFIG.logSheet);
  if (sheet.getLastRow() < 1) sheet.appendRow(['Thời gian', 'User thao tác', 'Thao tác', 'Mã đơn', 'Edit ID', 'ID dòng', 'Trước khi sửa', 'Sau khi sửa', 'SKU bị xóa']);
  return sheet;
}

function returnTh_logHasEdit_(ss, editId) {
  const sheet = ss.getSheetByName(RETURN_TH_CONFIG.logSheet);
  if (!sheet || sheet.getLastRow() < 2) return false;
  return sheet.getRange(2, 5, sheet.getLastRow() - 1, 1).getDisplayValues().some(v => String(v[0] || '').trim() === editId);
}

function returnTh_summaryText_(map) {
  if (!map) return '';
  return Object.keys(map).map(id => {
    const d = map[id] || {};
    return id + ': ' + [d.maDon, d.loaiHoan, d.tenSanPham, 'SL ' + d.soLuong, d.tinhTrang, d.ghiChu].filter(x => x !== undefined && x !== '').join(' | ');
  }).join('\n');
}

/** Ghi log sửa/xóa vào tab Log sửa TH và gửi mail báo admin. Lỗi mail chỉ trả về cho app biết, không làm hỏng lượt lưu. */
function returnTh_logEdit_(ss, now, editor, action, maDon, editId, ids, before, after, deleteIds) {
  const stamp = Utilities.formatDate(now, RETURN_TH_CONFIG.timezone, 'dd/MM/yyyy HH:mm:ss');
  const beforeText = returnTh_summaryText_(before);
  const afterText = returnTh_summaryText_(after);
  const deletedText = (deleteIds || []).join(', ');
  returnTh_logSheet_(ss).appendRow([stamp, editor, action, maDon, editId, ids.join(', '), beforeText, afterText, deletedText]);
  if (!RETURN_TH_CONFIG.alertEmail) return '';
  try {
    MailApp.sendEmail({
      to: RETURN_TH_CONFIG.alertEmail,
      subject: '[Nhập TH] ' + action + ' ' + maDon + ' - ' + (editor || 'không rõ user'),
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

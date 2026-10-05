// FULL-TABS V1: install with 00/10/11 helpers; follow README, not legacy installation comments below.
const SOURCE_SPREADSHEET_ID = '1V39AVE2JfEtMPYqnG1-J75fKPDXU37fCnJ9Na3UXOco';
const TARGET_SPREADSHEET_ID = '1ZLqeo_djwMJofqBCD3Ilk6lIkUmVsEfK7T5Oj9D6wmI';

const DOC_SOURCE_SHEET = 'Chứng từ_full';
const DOC_TARGET_SHEET = 'Đã bàn giao CT';
const PRODUCT_SOURCE_SHEET = 'Hoàn sản phẩm';
const PRODUCT_TARGET_SHEET = 'Bàn giao SP_Hàng hoàn';

function syncAllBanGiao() {
  syncBanGiaoChungTu();
  syncBanGiaoSanPham();
}

/* Cột đích ghi theo TÊN tiêu đề trên tab đích: muốn đổi thứ tự cột thì kéo cột ngay trên Sheet đích,
 * dữ liệu tự đi theo. columns: [tên cột đích, tên cột nguồn]. Cột đích chưa có thì được thêm vào cuối. */
function syncBanGiaoChungTu() {
  syncSheetByDate_({
    sourceSheetName: DOC_SOURCE_SHEET,
    targetSheetName: DOC_TARGET_SHEET,
    columns: [
      ['Ngày lên đơn', 'Ngày lên đơn'],
      ['Mã đơn GHTK', 'Mã đơn GHTK'],
      ['Khách Hàng', 'Khách Hàng'],
      ['Mã PO', 'Mã PO'],
      ['Số đơn hàng', 'Số đơn hàng'],
      ['Địa chỉ nhận hàng', 'Địa chỉ nhận hàng'],
      ['Xác thực hóa đơn', 'Xác thực hóa đơn'],
      ['Ngày bàn giao CT', 'Ngày bàn giao CT'],
      ['Note', 'Note'],
      ['Link ảnh', 'Link ảnh'],
      ['Loại siêu thị', 'Loại siêu thị'],
    ],
    handoverHeader: 'Ngày bàn giao CT',
    dateHeaders: ['Ngày lên đơn', 'Ngày bàn giao CT'],
  });
}

function syncBanGiaoSanPham() {
  syncSheetWhenHasProduct_({
    sourceSheetName: PRODUCT_SOURCE_SHEET,
    targetSheetName: PRODUCT_TARGET_SHEET,
    // [tên cột đích, tên cột nguồn]. LOOKUP = lấy từ Booking_full theo Số đơn hàng.
    columns: [
      ['Ngày nhận trả', 'Ngày hoàn trả'],
      ['Mã đơn', 'Mã đơn'],
      ['Tên khách hàng', 'Tên khách hàng'],
      ['Số đơn hàng', 'Số đơn hàng'],
      ['Mã PO', 'Mã PO'],
      ['Ngày booking', 'LOOKUP'],
      ['Mã vật tư', 'Mã vật tư'],
      ['Barcode', 'Barcode'],
      ['Tên sản phẩm', 'Tên sản phẩm'],
      ['Số lượng', 'Số lượng'],
      ['Phân loại', 'Phân loại'],
      ['Tình trạng FF', 'Tình trạng FF'],
      ['Ghi chú', 'Ghi chú'],
      ['Hạn sử dụng', 'Hạn sử dụng'],
      ['Kích thước', 'Kích thước'],
      ['Khối lượng', 'Khối lượng'],
      ['Loại siêu thị', 'Loại siêu thị'],
      ['Hình ảnh', 'Hình ảnh'],
      ['CBM', 'CBM'],
      ['Ngày bàn giao', 'Ngày bàn giao'],
    ],
    // Co san pham la day sang dich. Ngay ban giao co the trong va update sau.
    requiredAnyHeaders: ['Mã vật tư', 'Tên sản phẩm', 'Mã đơn'],
    handoverHeader: 'Ngày bàn giao',
    lookupByOrderNumber: {
      sheetName: 'Booking_full',
      // Booking dung cot Row Labels lam khoa, gia tri cua cot nay la So don hang.
      sourceKeyHeaders: ['Row Labels', 'Số đơn hàng', 'Mã đơn hàng KH', 'Mã đơn hàng', 'OD'],
      lookupValueHeaders: ['Ngày'],
    },
    dateHeaders: ['Ngày nhận trả', 'Ngày bàn giao', 'Ngày booking'],
    timeHeaders: [],
  });
}

function pnLegacyBanGiaoOnEdit(e) {
  handleBanGiaoEdit(e);
}

function handleBanGiaoEdit(e) {
  if (!e || !e.source || e.source.getId() !== SOURCE_SPREADSHEET_ID) return;

  return banGiaoWithLock_(function() {
    const range = e.range;
    const sheet = range.getSheet();
    const sheetName = sheet.getName();
    const startCol = range.getColumn();
    const endCol = startCol + range.getNumColumns() - 1;

    if (sheetName === DOC_SOURCE_SHEET) {
      const handoverColumn = findHeaderColumn_(sheet, 'Ngày bàn giao CT');
      if (handoverColumn && startCol <= handoverColumn && endCol >= handoverColumn) {
        if (rangeContainsValidDate_(range, handoverColumn)) syncBanGiaoChungTu();
      }
      return;
    }

    if (sheetName === PRODUCT_SOURCE_SHEET) {
      // San pham: bat cu update nao tren tab Hoan san pham deu rebuild target.
      // Như vậy app/user nhập dữ liệu trước vẫn sang target, ngày bàn giao điền sau sẽ update tiếp.
      syncBanGiaoSanPham();
    }
  });
}

function setupBanGiaoTrigger() { return pnInstallTriggers(); /* superseded */

  const ss = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'handleBanGiaoEdit')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger('handleBanGiaoEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
}

function syncSheetByDate_(config) {
  const sourceSheet = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID).getSheetByName(config.sourceSheetName);
  const targetSheet = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID).getSheetByName(config.targetSheetName);
  if (!sourceSheet) throw new Error(`Khong tim thay tab nguon: ${config.sourceSheetName}`);
  if (!targetSheet) throw new Error(`Khong tim thay tab dich: ${config.targetSheetName}`);

  const lastRow = sourceSheet.getLastRow();
  const sourceHeaderMap = getHeaderMap_(sourceSheet);
  const sourceColumns = config.columns.map(([, src]) => requireHeaderColumn_(sourceHeaderMap, src, config.sourceSheetName));
  const handoverColumn = requireHeaderColumn_(sourceHeaderMap, config.handoverHeader, config.sourceSheetName);
  const maxSourceCol = Math.max(...sourceColumns, handoverColumn);

  const output = [];
  if (lastRow >= 2) {
    const rows = sourceSheet.getRange(2, 1, lastRow - 1, maxSourceCol).getValues();
    rows.forEach((row, index) => {
      const handoverValue = row[handoverColumn - 1];
      if (!isRealDate_(handoverValue)) return;
      output.push({sortDate: handoverValue, sourceIndex: index, values: sourceColumns.map(col => row[col - 1])});
    });
  }

  output.sort((a, b) => {
    const dateDiff = a.sortDate.getTime() - b.sortDate.getTime();
    return dateDiff || a.sourceIndex - b.sourceIndex;
  });

  pnRequireReady_();
  writeTargetByHeaders_(targetSheet, config, output.map(item => item.values));
}

function syncSheetWhenHasProduct_(config) {
  const sourceSheet = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID).getSheetByName(config.sourceSheetName);
  const targetSheet = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID).getSheetByName(config.targetSheetName);
  if (!sourceSheet) throw new Error(`Khong tim thay tab nguon: ${config.sourceSheetName}`);
  if (!targetSheet) throw new Error(`Khong tim thay tab dich: ${config.targetSheetName}`);

  const lastRow = sourceSheet.getLastRow();
  const sourceHeaderMap = getHeaderMap_(sourceSheet);
  // Cột LOOKUP không đọc từ nguồn (0), lấy từ Booking_full theo Số đơn hàng.
  const sourceColumns = config.columns.map(([, src]) => src === 'LOOKUP' ? 0 : requireHeaderColumn_(sourceHeaderMap, src, config.sourceSheetName));
  const requiredColumns = (config.requiredAnyHeaders || []).map(header => requireHeaderColumn_(sourceHeaderMap, header, config.sourceSheetName));
  const handoverColumn = requireHeaderColumn_(sourceHeaderMap, config.handoverHeader, config.sourceSheetName);
  const lookupMap = config.lookupByOrderNumber ? buildLookupByOrderNumber_(config.lookupByOrderNumber) : null;
  const sourceOrderColumn = lookupMap ? requireHeaderColumn_(sourceHeaderMap, 'Số đơn hàng', config.sourceSheetName) : 0;
  const maxSourceCol = Math.max(...sourceColumns, ...requiredColumns, handoverColumn, sourceOrderColumn);

  const output = [];
  if (lastRow >= 2) {
    const rows = sourceSheet.getRange(2, 1, lastRow - 1, maxSourceCol).getValues();
    rows.forEach((row, index) => {
      const hasProductInfo = requiredColumns.some(col => hasValue_(row[col - 1]));
      if (!hasProductInfo) return;
      const handoverValue = row[handoverColumn - 1];
      const lookup = lookupMap ? (lookupMap[normalizeLookupKey_(row[sourceOrderColumn - 1])] || '') : '';
      output.push({
        sortDate: isRealDate_(handoverValue) ? handoverValue : null,
        sourceIndex: index,
        values: sourceColumns.map(col => col ? row[col - 1] : lookup),
      });
    });
  }

  output.sort((a, b) => {
    if (a.sortDate && b.sortDate) {
      const dateDiff = a.sortDate.getTime() - b.sortDate.getTime();
      if (dateDiff) return dateDiff;
    }
    if (a.sortDate && !b.sortDate) return -1;
    if (!a.sortDate && b.sortDate) return 1;
    return a.sourceIndex - b.sourceIndex;
  });

  pnRequireReady_();
  writeTargetByHeaders_(targetSheet, config, output.map(item => item.values));
}

/* Ghi dữ liệu sang tab đích theo TÊN cột ở dòng 1 của tab đích.
 * - Cột đích chưa có tiêu đề -> thêm vào cuối (tab mới hoàn toàn thì tạo theo thứ tự trong config).
 * - Chỉ ghi/xóa các cột có trong config; cột khác trên tab đích không bị đụng.
 * - Định dạng ngày/giờ theo tên cột. */
function writeTargetByHeaders_(targetSheet, config, rows) {
  const lastCol = Math.max(targetSheet.getLastColumn(), 1);
  const headers = targetSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h == null ? '' : h).trim());
  const empty = headers.every(h => !h);
  const map = {};
  if (empty) headers.length = 0;
  headers.forEach((h, i) => { const k = normalizeHeader_(h); if (k && !map[k]) map[k] = i + 1; });
  const cols = config.columns.map(([target]) => {
    const k = normalizeHeader_(target);
    if (!map[k]) { headers.push(target); map[k] = headers.length; targetSheet.getRange(1, headers.length).setValue(target); }
    return map[k];
  });
  const oldLast = targetSheet.getLastRow();
  // Ghi từng nhóm cột liền nhau.
  const order = cols.map((c, i) => ({c, i})).sort((a, b) => a.c - b.c), runs = [];
  order.forEach(x => { const r = runs[runs.length - 1]; if (r && r.end === x.c - 1) { r.end = x.c; r.idx.push(x.i); } else runs.push({start: x.c, end: x.c, idx: [x.i]}); });
  runs.forEach(run => {
    const w = run.end - run.start + 1;
    if (rows.length) targetSheet.getRange(2, run.start, rows.length, w).setValues(rows.map(v => run.idx.map(i => v[i])));
    if (oldLast > rows.length + 1) targetSheet.getRange(rows.length + 2, run.start, oldLast - rows.length - 1, w).clearContent();
  });
  const n = Math.max(rows.length, 1);
  const fmt = (names, pattern) => (names || []).forEach(name => {
    const c = map[normalizeHeader_(name)];
    if (c) targetSheet.getRange(2, c, n, 1).setNumberFormat(pattern);
  });
  fmt(config.dateHeaders, 'dd/MM/yyyy');
  fmt(config.timeHeaders, 'hh:mm:ss');
}

function buildLookupByOrderNumber_(config) {
  const sheet = SpreadsheetApp
    .openById(SOURCE_SPREADSHEET_ID)
    .getSheetByName(config.sheetName);
  if (!sheet) throw new Error(`Khong tim thay tab lookup: ${config.sheetName}`);

  const headerMap = getHeaderMap_(sheet);
  const keyColumn = findFirstHeaderColumn_(headerMap, config.sourceKeyHeaders || []);
  const valueColumn = findFirstHeaderColumn_(headerMap, config.lookupValueHeaders || []);
  if (!keyColumn) throw new Error(`Khong tim thay cot khoa trong tab lookup: ${config.sheetName}`);
  if (!valueColumn) throw new Error(`Khong tim thay cot gia tri trong tab lookup: ${config.sheetName}`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const rows = sheet.getRange(2, 1, lastRow - 1, Math.max(keyColumn, valueColumn)).getValues();
  const out = {};
  rows.forEach(row => {
    const key = normalizeLookupKey_(row[keyColumn - 1]);
    const value = row[valueColumn - 1];
    if (key && hasValue_(value) && !Object.prototype.hasOwnProperty.call(out, key)) {
      out[key] = value;
    }
  });
  return out;
}

function findFirstHeaderColumn_(headerMap, headers) {
  for (const header of headers) {
    const column = headerMap[normalizeHeader_(header)];
    if (column) return column;
  }
  return 0;
}

function normalizeLookupKey_(value) {
  return String(value == null ? '' : value)
    .trim()
    .replace(/\.0+$/, '');
}

function clearTargetContent_(sheet, width) {
  const rows = Math.max(sheet.getLastRow(), 1);
  const cols = Math.max(Number(width || sheet.getLastColumn()), 1);
  sheet.getRange(1, 1, rows, cols).clearContent();
}

function applyTargetFormats_(sheet, config, rowCount) {
  (config.dateFormatColumns || []).forEach(col => {
    sheet.getRange(2, col, Math.max(rowCount - 1, 1), 1).setNumberFormat('dd/MM/yyyy');
  });

  (config.timeFormatColumns || []).forEach(col => {
    sheet.getRange(2, col, Math.max(rowCount - 1, 1), 1).setNumberFormat('hh:mm:ss');
  });
}

function rangeContainsValidDate_(range, sheetColumn) {
  const offset = sheetColumn - range.getColumn();
  if (offset < 0 || offset >= range.getNumColumns()) return false;

  return range.getValues().some(row => isRealDate_(row[offset]));
}

function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.reduce((map, header, index) => {
    const key = normalizeHeader_(header);
    if (key && !map[key]) map[key] = index + 1;
    return map;
  }, {});
}

function findHeaderColumn_(sheet, header) {
  return getHeaderMap_(sheet)[normalizeHeader_(header)] || 0;
}

function requireHeaderColumn_(headerMap, header, sheetName) {
  const column = headerMap[normalizeHeader_(header)];
  if (!column) throw new Error(`Khong tim thay header "${header}" trong tab "${sheetName}"`);
  return column;
}

function normalizeHeader_(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function isRealDate_(value) {
  return Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime());
}

function hasValue_(value) {
  return value !== null && value !== '' && typeof value !== 'undefined';
}

function banGiaoWithLock_(fn) {
  return pnWithLock_(fn);
}

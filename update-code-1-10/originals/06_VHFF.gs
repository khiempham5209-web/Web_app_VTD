const SOURCE_SPREADSHEET_ID = '1V39AVE2JfEtMPYqnG1-J75fKPDXU37fCnJ9Na3UXOco';
const TARGET_SPREADSHEET_ID = '1ZLqeo_djwMJofqBCD3Ilk6lIkUmVsEfK7T5Oj9D6wmI';

const DOC_SOURCE_SHEET = 'Chứng từ_FF';
const DOC_TARGET_SHEET = 'Đã bàn giao CT';
const PRODUCT_SOURCE_SHEET = 'Hoàn sản phẩm';
const PRODUCT_TARGET_SHEET = 'Bàn giao SP_Hàng hoàn';

function syncAllBanGiao() {
  syncBanGiaoChungTu();
  syncBanGiaoSanPham();
}

function syncBanGiaoChungTu() {
  syncSheetByDate_({
    sourceSheetName: DOC_SOURCE_SHEET,
    targetSheetName: DOC_TARGET_SHEET,
    targetHeaders: [
      'Ngày lên đơn',
      'Mã đơn GHTK',
      'Khách Hàng',
      'Mã PO',
      'Số đơn hàng',
      'Địa chỉ nhận hàng',
      'Xác thực hóa đơn',
      'Ngày bàn giao CT',
      'Note',
      'Link ảnh',
      'Loại siêu thị',
    ],
    sourceHeaders: [
      'Ngày lên đơn',
      'Mã đơn GHTK',
      'Khách Hàng',
      'Mã PO',
      'Số đơn hàng',
      'Địa chỉ nhận hàng',
      'Xác thực hóa đơn',
      'Ngày bàn giao CT',
      'Note',
      'Link ảnh',
      'Loại siêu thị',
    ],
    handoverHeader: 'Ngày bàn giao CT',
    dateFormatColumns: [1, 8],
  });
}

function syncBanGiaoSanPham() {
  syncSheetWhenHasProduct_({
    sourceSheetName: PRODUCT_SOURCE_SHEET,
    targetSheetName: PRODUCT_TARGET_SHEET,
    targetHeaders: [
      'Ngày nhận trả',
      'Mã đơn',
      'Tên khách hàng',
      'Số đơn hàng',
      'Mã PO',
      'Mã vật tư',
      'Barcode',
      'Tên sản phẩm',
      'Số lượng',
      'Phân loại',
      'Tình trạng FF',
      'Ghi chú',
      'Hạn sử dụng',
      'Kích thước',
      'Khối lượng',
      'Loại siêu thị',
      'Hình ảnh',
      'CBM',
      'Ngày bàn giao',
      'Ngày booking',
    ],
    sourceHeaders: [
      'Ngày hoàn trả',
      'Mã đơn',
      'Tên khách hàng',
      'Số đơn hàng',
      'Mã PO',
      'Mã vật tư',
      'Barcode',
      'Tên sản phẩm',
      'Số lượng',
      'Phân loại',
      'Tình trạng FF',
      'Ghi chú',
      'Hạn sử dụng',
      'Kích thước',
      'Khối lượng',
      'Loại siêu thị',
      'Hình ảnh',
      'CBM',
      'Ngày bàn giao',
    ],
    // Co san pham la day sang dich. Ngay ban giao co the trong va update sau.
    requiredAnyHeaders: ['Mã vật tư', 'Tên sản phẩm', 'Mã đơn'],
    handoverHeader: 'Ngày bàn giao',
    lookupByOrderNumber: {
      sheetName: 'Booking',
      // Booking dung cot Row Labels lam khoa, gia tri cua cot nay la So don hang.
      sourceKeyHeaders: ['Row Labels', 'Số đơn hàng', 'Mã đơn hàng KH', 'Mã đơn hàng', 'OD'],
      lookupValueHeaders: ['Ngày'],
    },
    dateFormatColumns: [1, 19, 20],
    timeFormatColumns: [],
  });
}

function onEdit(e) {
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

function setupBanGiaoTrigger() {
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
  const sourceColumns = config.sourceHeaders.map(header => requireHeaderColumn_(sourceHeaderMap, header, config.sourceSheetName));
  const handoverColumn = requireHeaderColumn_(sourceHeaderMap, config.handoverHeader, config.sourceSheetName);
  const maxSourceCol = Math.max(...sourceColumns, handoverColumn);
  const outputWidth = config.targetHeaders.length;

  const output = [];
  if (lastRow >= 2) {
    const rows = sourceSheet.getRange(2, 1, lastRow - 1, maxSourceCol).getValues();
    rows.forEach((row, index) => {
      const handoverValue = row[handoverColumn - 1];
      if (!isRealDate_(handoverValue)) return;
      output.push({
        sortDate: handoverValue,
        sourceIndex: index,
        values: sourceColumns.map(col => row[col - 1]),
      });
    });
  }

  output.sort((a, b) => {
    const dateDiff = a.sortDate.getTime() - b.sortDate.getTime();
    return dateDiff || a.sourceIndex - b.sourceIndex;
  });

  clearTargetContent_(targetSheet, outputWidth);
  targetSheet.getRange(1, 1, 1, outputWidth).setValues([config.targetHeaders]);
  if (output.length) {
    targetSheet.getRange(2, 1, output.length, outputWidth).setValues(output.map(item => item.values));
  }

  applyTargetFormats_(targetSheet, config, Math.max(output.length + 1, 2));
}

function syncSheetWhenHasProduct_(config) {
  const sourceSheet = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID).getSheetByName(config.sourceSheetName);
  const targetSheet = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID).getSheetByName(config.targetSheetName);
  if (!sourceSheet) throw new Error(`Khong tim thay tab nguon: ${config.sourceSheetName}`);
  if (!targetSheet) throw new Error(`Khong tim thay tab dich: ${config.targetSheetName}`);

  const lastRow = sourceSheet.getLastRow();
  const sourceHeaderMap = getHeaderMap_(sourceSheet);
  const sourceColumns = config.sourceHeaders.map(header => requireHeaderColumn_(sourceHeaderMap, header, config.sourceSheetName));
  const requiredColumns = (config.requiredAnyHeaders || []).map(header => requireHeaderColumn_(sourceHeaderMap, header, config.sourceSheetName));
  const handoverColumn = requireHeaderColumn_(sourceHeaderMap, config.handoverHeader, config.sourceSheetName);
  const lookupMap = config.lookupByOrderNumber
    ? buildLookupByOrderNumber_(config.lookupByOrderNumber)
    : null;
  const sourceOrderColumn = lookupMap
    ? requireHeaderColumn_(sourceHeaderMap, 'Số đơn hàng', config.sourceSheetName)
    : 0;
  const maxSourceCol = Math.max(...sourceColumns, ...requiredColumns, handoverColumn);
  const outputWidth = config.targetHeaders.length;

  const output = [];
  if (lastRow >= 2) {
    const rows = sourceSheet.getRange(2, 1, lastRow - 1, maxSourceCol).getValues();
    rows.forEach((row, index) => {
      const hasProductInfo = requiredColumns.some(col => hasValue_(row[col - 1]));
      if (!hasProductInfo) return;

      const handoverValue = row[handoverColumn - 1];
      output.push({
        sortDate: isRealDate_(handoverValue) ? handoverValue : null,
        sourceIndex: index,
        values: sourceColumns.map(col => row[col - 1]).concat(
          lookupMap
            ? [lookupMap[normalizeLookupKey_(row[sourceOrderColumn - 1])] || '']
            : []
        ),
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

  clearTargetContent_(targetSheet, outputWidth);
  targetSheet.getRange(1, 1, 1, outputWidth).setValues([config.targetHeaders]);
  if (output.length) {
    targetSheet.getRange(2, 1, output.length, outputWidth).setValues(output.map(item => item.values));
  }

  applyTargetFormats_(targetSheet, config, Math.max(output.length + 1, 2));
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
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

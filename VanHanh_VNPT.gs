/**
 * VNPT V1 — VẬN HÀNH SHEET THEO CÁC LUỒNG ĐỘC LẬP (dựa trên FFPL V4.2 của TH)
 *
 * Đặt trong 1 project Apps Script RIÊNG (standalone), không bound vào Google Sheet.
 * Mở sheet bằng openById, trigger cài bằng vnptInstallTriggers().
 * Nếu muốn email báo lỗi: thêm CanhBaoLoi_VNPT.gs (vnptNotifyError) vào cùng project.
 *
 * Khác TH:
 *  - Tab: Tổng hợp / Nhập / Hoàn trả / Đóng tay_chuyển kho / Thông tin sku / Xuất.
 *  - Không có DS đơn all, Kiểm kê, cột TÌNH TRẠNG TỒN KHO.
 *  - Tab Xuất ghi tay; code chỉ đọc để cộng XUẤT FF.
 *  - Cột Tình trạng (Nhập/Hoàn trả) không ảnh hưởng tồn.
 *  - Kích thước nhận 2 chiều (DxR) hoặc 3 chiều (DxRxC); 2 chiều không tính CBM.
 *  - Số dự kiến = Tổng NVL (m) / số ở Tổng hợp!M3. M3 trống → Số dự kiến để trống.
 *
 * Lần đầu: chạy vnptKhoiTaoTongHop() để xóa dữ liệu mẫu ở Tổng hợp rồi dựng lại,
 * sau đó chạy vnptInstallTriggers().
 */

const VNPT = {
  id: '1EOyNJyjUoXXXmOH-OJCD_Vsyauhmm2Uj_Ujnqaj6e-s',
  title: 'TỔNG HỢP TỒN KHO – VNPT',

  sheets: {
    all: 'Tổng hợp',
    sku: 'Thông tin sku',
    in: 'Nhập',
    ff: 'Xuất',
    transfer: 'Đóng tay_chuyển kho',
    returns: 'Hoàn trả'
  },

  headers: {
    'Tổng hợp': 3,
    'Thông tin sku': 3,
    'Nhập': 3,
    'Xuất': 3,
    'Đóng tay_chuyển kho': 3,
    'Hoàn trả': 3
  },

  state: '_VNPT_STATE',
  rangeAll: 'VNPT_ALL',

  // Ô chứa số chia NVL (vd 100 hoặc 1000) trên tab Tổng hợp.
  rollCell: 'M3',
  pollMinutes: 5,

  laneHeaders: {
    in: 'NHẬP',
    ff: 'XUẤT FF',
    hand: 'ĐÓNG TAY',
    transfer: 'CHUYỂN KHO',
    returns: 'HOÀN'
  },

  required: {
    'Tổng hợp': ['SUB SKU','TÊN SKU','NHẬP','XUẤT FF','ĐÓNG TAY','CHUYỂN KHO','HOÀN','TỒN','CBM','Tổng NVL','Số dự kiến'],
    'Thông tin sku': ['SUB SKU','Tên sản phẩm','Đơn vị','Kích thước','Tổng mức tiêu hao'],
    'Nhập': ['SUB SKU','Tên sản phẩm','Ngày','Số lượng'],
    'Xuất': ['THÁNG','NGÀY','SUB SKU','TÊN SẢN PHẨM','SL đơn','SL SP','SL đơn APL','SL SP APL'],
    'Đóng tay_chuyển kho': ['SUB SKU','Tên sản phẩm','Ngày','Số lượng','Loại xuất','Trạng thái'],
    'Hoàn trả': ['Tên sản phẩm','Ngày tháng','Số lượng','Mã đơn','Loại hoàn']
  }
};

/************************************************
 * HÀM CƠ BẢN
 ************************************************/

let vnSsCache_ = null;
function vnss_() {
  if (!vnSsCache_) vnSsCache_ = SpreadsheetApp.openById(VNPT.id);
  return vnSsCache_;
}

function vntext_(v) { return v === null || v === undefined ? '' : String(v).trim(); }
function vnkey_(v) { return vntext_(v).normalize('NFC').replace(/\s+/g, ' ').toLowerCase(); }
function vnblank_(v) { return v === null || v === undefined || (typeof v === 'string' && v.trim() === ''); }
function vneffective_(base, actual) { return vnblank_(actual) ? base : actual; }

function vnnum_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  if (typeof v === 'number' && Number.isFinite(v)) return v;
  const t = vntext_(v);
  const n = t.includes(',') ? Number(t.replace(/\./g, '').replace(',', '.')) : Number(t);
  if (!Number.isFinite(n)) throw new Error('Số lượng không hợp lệ: ' + t);
  return n;
}

function vndate_(value) {
  if (value instanceof Date && !isNaN(value)) return value;
  const text = vntext_(value);
  if (/^\d{4}-\d{2}-\d{2}T.*(?:Z|[+-]\d{2}:?\d{2})$/.test(text)) {
    const instant = new Date(text);
    if (!isNaN(instant)) return instant;
  }
  let m = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s|$)/), key;
  if (m) key = m[3] + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0');
  else {
    m = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s]|$)/);
    if (m) key = m[1] + '-' + m[2].padStart(2, '0') + '-' + m[3].padStart(2, '0');
  }
  if (!key) throw new Error('Ngày phải là Date hoặc dd/MM/yyyy: ' + text);
  const tz = vnss_().getSpreadsheetTimeZone();
  const d = Utilities.parseDate(key, tz, 'yyyy-MM-dd');
  if (isNaN(d) || Utilities.formatDate(d, tz, 'yyyy-MM-dd') !== key) throw new Error('Ngày không hợp lệ: ' + text);
  return d;
}

function vnmonth_(v) {
  return Utilities.formatDate(vndate_(v), vnss_().getSpreadsheetTimeZone(), 'yyyy-MM');
}

function vneq_(a, b) {
  if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
  return a === b;
}

function vnlock_(cb) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Đang có lượt cập nhật, thử lại sau.');
  try { return cb(); } finally { lock.releaseLock(); }
}

function vnensureGrid_(sheet, lastRow, lastCol) {
  if (sheet.getMaxRows() < lastRow) sheet.insertRowsAfter(sheet.getMaxRows(), lastRow - sheet.getMaxRows());
  if (sheet.getMaxColumns() < lastCol) sheet.insertColumnsAfter(sheet.getMaxColumns(), lastCol - sheet.getMaxColumns());
}

/** SUB SKU > 15 chữ số nếu lưu dạng số sẽ bị làm tròn → báo lỗi thay vì khớp sai. */
function vnsku_(raw, where) {
  if (typeof raw === 'number' && !Number.isSafeInteger(raw)) {
    throw new Error('SUB SKU dạng số quá dài, đã bị làm tròn (' + raw + ') tại ' + where +
      '. Định dạng cột SUB SKU là Văn bản thuần và gõ lại mã.');
  }
  return vntext_(raw);
}

/************************************************
 * HEADER VÀ ĐỌC BẢNG
 ************************************************/

function vnmap_(headers, location) {
  const map = {};
  headers.forEach((h, i) => {
    const k = vnkey_(h);
    if (!k) return;
    if (Object.prototype.hasOwnProperty.call(map, k)) throw new Error('Trùng header "' + h + '" ở ' + location);
    map[k] = i;
  });
  return map;
}

function vncol_(table, name, optional) {
  const c = table.h[vnkey_(name)];
  if (c === undefined && !optional) throw new Error('Thiếu header "' + name + '" tại ' + table.sh.getName());
  return c;
}

function vnrequire_(table, headers) { headers.forEach(h => vncol_(table, h)); }

function vntable_(sheetName) {
  const sheet = vnss_().getSheetByName(sheetName);
  if (!sheet) throw new Error('Thiếu tab ' + sheetName);
  const hr = VNPT.headers[sheetName];
  if (!hr) throw new Error('Chưa cấu hình hàng header: ' + sheetName);

  const raw = sheet.getRange(hr, 1, 1, sheet.getLastColumn()).getValues()[0];
  let width = raw.length;
  // Tổng hợp: ô M3 là số cấu hình NVL, không phải header của bảng.
  if (sheetName === VNPT.sheets.all) {
    width = Math.min(width, sheet.getRange(VNPT.rollCell).getColumn() - 1);
  }
  while (width && !vntext_(raw[width - 1])) width--;
  if (!width) throw new Error('Header trống: ' + sheetName);

  const headers = raw.slice(0, width);
  const lastRow = sheet.getLastRow();
  const rows = lastRow > hr ? sheet.getRange(hr + 1, 1, lastRow - hr, width).getValues() : [];
  return { sh: sheet, hr: hr, width: width, headers: headers, h: vnmap_(headers, sheetName), rows: rows };
}

function vnset_(sheet, row, colIndex, value) {
  const r = sheet.getRange(row, colIndex + 1);
  if (r.getFormula()) throw new Error('Không ghi đè công thức: ' + sheet.getName() + '!' + r.getA1Notation());
  if (!vneq_(r.getValue(), value)) r.setValue(value);
}

function vnapplyCells_(sheet, edits) {
  const byCol = new Map();
  edits.forEach(e => {
    if (!byCol.has(e.col)) byCol.set(e.col, new Map());
    byCol.get(e.col).set(e.row, e);
  });
  const blocks = [];
  for (const [col, map] of byCol) {
    const rows = [...map.keys()].sort((a, b) => a - b);
    let s = 0;
    while (s < rows.length) {
      let e = s + 1;
      while (e < rows.length && rows[e] === rows[e - 1] + 1) e++;
      const range = sheet.getRange(rows[s], col + 1, e - s, 1);
      if (range.getFormulas().some(r => r[0]))
        throw new Error('Có công thức trong vùng code quản lý: ' + sheet.getName() + '!' + range.getA1Notation());
      blocks.push({ range: range, values: rows.slice(s, e).map(r => [map.get(r).value]) });
      s = e;
    }
  }
  // Kiểm tra hết rồi mới ghi.
  blocks.forEach(b => b.range.setValues(b.values));
}

function vnvalidateSheet_(sheetName) {
  const req = VNPT.required[sheetName];
  if (req) vnrequire_(vntable_(sheetName), req);
}

/************************************************
 * DANH MỤC SKU
 ************************************************/

function vnparseSize_(text, name) {
  // Bỏ ký tự lạ (dấu kết hợp, khoảng trắng, "cm"...), chỉ giữ số, dấu thập phân và x.
  const clean = vntext_(text).normalize('NFD').replace(/[^0-9.,x×X*]/g, '').replace(/[×X*]/g, 'x');
  if (!clean) return null;
  const parts = clean.split('x');
  if (parts.length < 2 || parts.length > 3 || parts.some(p => !/^\d+(?:[.,]\d+)?$/.test(p)))
    throw new Error('Kích thước không hợp lệ (cần DxR hoặc DxRxC): ' + name);
  // 2 chiều: không có thể tích → không tính CBM.
  if (parts.length === 2) return null;
  return parts.map(p => Number(p.replace(',', '.'))).reduce((a, b) => a * b, 1) / 1000000;
}

function vncatalog_() {
  const t = vntable_(VNPT.sheets.sku);
  vnrequire_(t, VNPT.required[VNPT.sheets.sku]);
  const skuC = vncol_(t, 'SUB SKU'), nameC = vncol_(t, 'Tên sản phẩm'), unitC = vncol_(t, 'Đơn vị');
  const sizeC = vncol_(t, 'Kích thước'), normC = vncol_(t, 'Tổng mức tiêu hao');
  const bySku = new Map(), byName = new Map();

  t.rows.forEach((row, i) => {
    const sku = vnsku_(row[skuC], t.sh.getName() + ' dòng ' + (t.hr + 1 + i));
    const name = vntext_(row[nameC]);
    if (!sku && !name) return;
    if (!sku || !name) throw new Error('Danh mục thiếu mã hoặc tên SKU dòng ' + (t.hr + 1 + i));
    if (bySku.has(sku) || byName.has(vnkey_(name))) throw new Error('Trùng mã/tên trong Thông tin sku: ' + name);
    if (!vnblank_(row[normC]) && vnnum_(row[normC]) < 0) throw new Error('Định mức âm tại SKU: ' + name);
    const p = {
      sku: sku,
      rawSku: row[skuC],
      name: name,
      unit: row[unitC],
      norm: vnblank_(row[normC]) ? null : vnnum_(row[normC]),
      cbm: vnparseSize_(row[sizeC], name)
    };
    bySku.set(sku, p);
    byName.set(vnkey_(name), p);
  });
  return { bySku: bySku, byName: byName };
}

function vnproduct_(catalog, name, sku) {
  const n = vntext_(name), s = vntext_(sku);
  const p = n ? catalog.byName.get(vnkey_(n)) : catalog.bySku.get(s);
  if (!p) throw new Error('Không tìm thấy sản phẩm trong danh mục: ' + (n || s));
  return p;
}

/************************************************
 * CHỌN TÊN → TỰ ĐIỀN SUB SKU / ĐƠN VỊ / THÁNG
 ************************************************/

function vnfillNames_(sheetName, catalog, editedRange) {
  const t = vntable_(sheetName);
  const skuC = vncol_(t, 'SUB SKU'), nameC = vncol_(t, 'Tên sản phẩm');
  const unitC = vncol_(t, 'Đơn vị', true), monthC = vncol_(t, 'Tháng', true), dateC = vncol_(t, 'Ngày', true);
  const statusC = sheetName === VNPT.sheets.transfer ? vncol_(t, 'Trạng thái') : undefined;
  const editedName = editedRange &&
    editedRange.getColumn() <= nameC + 1 && editedRange.getLastColumn() >= nameC + 1;
  const plan = [];

  t.rows.forEach((row, i) => {
    const r = t.hr + 1 + i;
    if (editedRange && (r < editedRange.getRow() || r > editedRange.getLastRow())) return;
    if (statusC !== undefined && vnkey_(row[statusC]) === 'hủy') return;
    const name = vntext_(row[nameC]);
    const sku = vnsku_(row[skuC], sheetName + ' dòng ' + r);
    if (!name && !sku) return;

    // Xóa tên → xóa mã và đơn vị tương ứng.
    if (editedName && !name) {
      const cells = [[skuC, '']];
      if (unitC !== undefined) cells.push([unitC, '']);
      plan.push([r, cells]);
      return;
    }

    const p = vnproduct_(catalog, name, sku);
    const cells = [[skuC, vntext_(p.rawSku)], [nameC, p.name]];
    if (unitC !== undefined) cells.push([unitC, p.unit]);
    if (monthC !== undefined && dateC !== undefined && !vnblank_(row[dateC])) {
      cells.push([monthC, vndate_(vnmonth_(row[dateC]) + '-01')]);
    }
    plan.push([r, cells]);
  });

  // Mã dài phải giữ dạng văn bản, tránh Sheets tự đổi thành số và làm tròn.
  if (plan.length) {
    t.sh.getRange(t.hr + 1, skuC + 1, t.sh.getMaxRows() - t.hr, 1).setNumberFormat('@');
  }
  plan.forEach(([r, cells]) => cells.forEach(([c, v]) => vnset_(t.sh, r, c, v)));
  return vntable_(sheetName);
}

/************************************************
 * TRẠNG THÁI PHÁT SINH CỦA TỪNG LUỒNG (tab ẩn)
 ************************************************/

const VN_STATE_HEADERS = ['Luồng', 'SUB SKU', 'Tháng', 'Số lượng'];

function vnstate_() {
  const ss = vnss_();
  let sh = ss.getSheetByName(VNPT.state);
  if (!sh) {
    sh = ss.insertSheet(VNPT.state);
    sh.getRange(1, 1, 1, 4).setValues([VN_STATE_HEADERS]);
    sh.getRange('B:C').setNumberFormat('@');
    sh.hideSheet();
  } else if (JSON.stringify(sh.getRange(1, 1, 1, 4).getValues()[0]) !== JSON.stringify(VN_STATE_HEADERS)) {
    throw new Error('Tab trạng thái không đúng cấu trúc. Không ghi đè.');
  }
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues() : [];
  return { sh: sh, rows: rows };
}

function vnstateMonth_(v) {
  let t;
  if (v instanceof Date && !isNaN(v)) t = vnmonth_(v);
  else {
    t = vntext_(v).replace(/^m:/, '');
    if (!/^\d{4}-\d{2}$/.test(t)) t = vnmonth_(v);
  }
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(t)) throw new Error('Tháng trạng thái không hợp lệ: ' + String(v));
  return t;
}

function vnstateKey_(row) {
  return JSON.stringify([vntext_(row[0]), vntext_(row[1]), vnstateMonth_(row[2])]);
}

function vnaggregate_(entries) {
  const map = new Map();
  entries.forEach(e => {
    const k = JSON.stringify([e.lane, e.sku, e.month]);
    const item = map.get(k) || { lane: e.lane, sku: e.sku, month: e.month, qty: 0 };
    item.qty += e.qty;
    map.set(k, item);
  });
  return map;
}

function vntotals_(entries, lane) {
  const totals = new Map();
  for (const e of entries.values()) {
    if (e.lane !== lane) continue;
    totals.set(e.sku, (totals.get(e.sku) || 0) + e.qty);
  }
  return totals;
}

function vncommitState_(state, lanes, next) {
  const old = new Map(), dups = [], edits = [], pending = [];
  state.rows.forEach((row, i) => {
    if (!lanes.includes(vntext_(row[0]))) return;
    const k = vnstateKey_(row);
    if (old.has(k)) dups.push(old.get(k));
    old.set(k, { r: row, i: i });
  });
  for (const [k, e] of next) {
    if (!lanes.includes(e.lane)) continue;
    const month = 'm:' + e.month;
    const ex = old.get(k);
    if (ex) {
      if (!vneq_(ex.r[2], month)) edits.push({ row: ex.i + 2, col: 2, value: month });
      if (!vneq_(ex.r[3], e.qty)) edits.push({ row: ex.i + 2, col: 3, value: e.qty });
      old.delete(k);
    } else pending.push([e.lane, e.sku, month, e.qty]);
  }
  for (const ex of old.values()) {
    const month = 'm:' + vnstateMonth_(ex.r[2]);
    if (!vneq_(ex.r[2], month)) edits.push({ row: ex.i + 2, col: 2, value: month });
    if (!vneq_(ex.r[3], 0)) edits.push({ row: ex.i + 2, col: 3, value: 0 });
  }
  dups.forEach(ex => ex.r.slice(0, 4).forEach((v, c) => {
    if (v !== '') edits.push({ row: ex.i + 2, col: c, value: '' });
  }));
  vnapplyCells_(state.sh, edits);
  if (pending.length) {
    const start = state.sh.getLastRow() + 1;
    vnensureGrid_(state.sh, start + pending.length - 1, 4);
    state.sh.getRange(start, 2, pending.length, 2).setNumberFormat('@');
    state.sh.getRange(start, 1, pending.length, 4).setValues(pending);
  }
}

/************************************************
 * VÙNG CODE QUẢN LÝ TRÊN TỔNG HỢP
 ************************************************/

function vnregion_(table, rangeName, keyHeaders, accept) {
  const ss = vnss_();
  let range = ss.getRangeByName(rangeName);
  if (range) {
    if (range.getSheet().getSheetId() !== table.sh.getSheetId() ||
        range.getRow() !== table.hr + 1 || range.getColumn() !== 1) {
      throw new Error('Vùng ' + rangeName + ' lệch header. Cần kiểm tra ranh giới, không tự ghi đè.');
    }
    if (range.getNumColumns() !== table.width) {
      range = table.sh.getRange(table.hr + 1, 1, range.getNumRows(), table.width);
      ss.setNamedRange(rangeName, range);
    }
    return range;
  }

  const keyCols = keyHeaders.map(h => vncol_(table, h));
  let count = 0;
  for (const row of table.rows) {
    if (!keyCols.every(c => vntext_(row[c])) || !accept(row)) break;
    count++;
  }
  if (!count && table.rows[0] && table.rows[0].some(v => v !== '')) {
    throw new Error('Nội dung lạ ngay đầu ' + table.sh.getName() +
      '. Dừng để bảo vệ bảng. Nếu là dữ liệu mẫu cũ, chạy vnptKhoiTaoTongHop().');
  }
  vnensureGrid_(table.sh, table.hr + Math.max(1, count), table.width);
  range = table.sh.getRange(table.hr + 1, 1, Math.max(1, count), table.width);
  ss.setNamedRange(rangeName, range);
  return range;
}

function vnappendOwned_(table, rangeName, values) {
  const ss = vnss_();
  const range = ss.getRangeByName(rangeName);
  const existing = range.getValues();
  const emptyIdx = existing.findIndex(r => r.every(v => v === ''));
  if (emptyIdx >= 0) {
    const r = range.getRow() + emptyIdx;
    values.forEach((v, c) => { if (v !== '' && v !== undefined) vnset_(table.sh, r, c, v); });
    return r;
  }
  const ins = range.getLastRow() + 1;
  if (ins > table.sh.getMaxRows()) table.sh.insertRowsAfter(ins - 1, 1);
  else table.sh.insertRowsBefore(ins, 1);
  table.sh.getRange(ins - 1, 1, 1, table.width).copyTo(
    table.sh.getRange(ins, 1, 1, table.width), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  table.sh.getRange(ins, 1, 1, table.width).setValues([values]);
  ss.setNamedRange(rangeName, table.sh.getRange(table.hr + 1, 1, existing.length + 1, table.width));
  return ins;
}

/************************************************
 * TỔNG HỢP
 ************************************************/

function vnsummary_(catalog) {
  const t = vntable_(VNPT.sheets.all);
  vnrequire_(t, VNPT.required[VNPT.sheets.all]);
  const skuC = vncol_(t, 'SUB SKU');
  const range = vnregion_(t, VNPT.rangeAll, ['SUB SKU'], row => catalog.bySku.has(vntext_(row[skuC])));
  const map = new Map();
  range.getValues().forEach((row, i) => {
    const sku = vnsku_(row[skuC], t.sh.getName() + ' dòng ' + (t.hr + 1 + i));
    if (!sku) return;
    if (map.has(sku)) throw new Error('Trùng SKU trên Tổng hợp: ' + sku);
    map.set(sku, { row: t.hr + 1 + i, values: row });
  });
  return { t: t, map: map };
}

/** Số chia NVL ở Tổng hợp!M3. Trống → null (Số dự kiến để trống). */
function vnrollDivisor_() {
  const v = vnss_().getSheetByName(VNPT.sheets.all).getRange(VNPT.rollCell).getValue();
  if (vnblank_(v)) return null;
  const n = vnnum_(v);
  if (!(n > 0)) throw new Error('Tổng hợp!' + VNPT.rollCell + ' phải là số > 0 (vd 100 hoặc 1000): ' + v);
  return n;
}

function vnderiveValues_(table, values, p, materials, divisor) {
  const col = h => vncol_(table, h), get = h => vnnum_(values[col(h)]);
  const stock = get('NHẬP') + get('HOÀN') - get('XUẤT FF') - get('ĐÓNG TAY') - get('CHUYỂN KHO');
  const result = [
    [col('TỒN'), stock],
    [col('CBM'), p.cbm === null ? '' : stock * p.cbm]
  ];
  if (materials !== false) {
    const exported = get('XUẤT FF') + get('ĐÓNG TAY');
    const nvl = exported <= 0 || p.norm === null ? '' : exported * p.norm / 100; // cm → mét
    const roll = nvl === '' || divisor === null ? '' : nvl / divisor;
    result.push([col('Tổng NVL'), nvl], [col('Số dự kiến'), roll]);
  }
  return result;
}

function vnwriteLane_(catalog, lane, totals, previousKeys) {
  const summary = vnsummary_(catalog), t = summary.t;
  const skuC = vncol_(t, 'SUB SKU'), nameC = vncol_(t, 'TÊN SKU');
  const target = vncol_(t, VNPT.laneHeaders[lane]);
  const divisor = vnrollDivisor_();
  const keys = new Set([...totals.keys(), ...previousKeys, ...summary.map.keys()]);
  const edits = [], additions = [], affected = new Set();

  for (const sku of keys) {
    const p = catalog.bySku.get(sku);
    if (!p) throw new Error('SKU bị xóa khỏi danh mục: ' + sku);
    const ex = summary.map.get(sku);
    if (!ex && !totals.has(sku)) continue;
    const values = ex ? ex.values.slice() : new Array(t.width).fill('');
    if (!ex) {
      values[skuC] = vntext_(p.rawSku);
      values[nameC] = p.name;
      Object.values(VNPT.laneHeaders).forEach(h => values[vncol_(t, h)] = 0);
    }
    const qty = totals.get(sku) || 0;
    const changed = !ex || !vneq_(values[target], qty);
    values[target] = qty;
    const computed = vnderiveValues_(t, values, p, !ex || lane === 'ff' || lane === 'hand', divisor);
    computed.forEach(([c, v]) => values[c] = v);
    if (ex) {
      [target, ...computed.map(x => x[0])].forEach(c => {
        if (!vneq_(ex.values[c], values[c])) edits.push({ row: ex.row, col: c, value: values[c] });
      });
    } else additions.push(values);
    if (changed) affected.add(sku);
  }

  vnapplyCells_(t.sh, edits);
  if (additions.length) {
    const owned = vnss_().getRangeByName(VNPT.rangeAll);
    t.sh.getRange(owned.getRow(), skuC + 1, t.sh.getMaxRows() - owned.getRow() + 1, 1).setNumberFormat('@');
  }
  additions.forEach(row => vnappendOwned_(t, VNPT.rangeAll, row));
  return affected;
}

function vnheadline_() {
  const ss = vnss_();
  const t = vntable_(VNPT.sheets.all);
  const range = ss.getRangeByName(VNPT.rangeAll);
  if (!range) return;
  const skuC = vncol_(t, 'SUB SKU'), stockC = vncol_(t, 'TỒN'), cbmC = vncol_(t, 'CBM');
  let skuCount = 0, totalStock = 0, totalCbm = 0;
  for (const row of range.getValues()) {
    if (!vntext_(row[skuC])) continue;
    skuCount++;
    const s = vnnum_(row[stockC]);
    // Chỉ cộng tồn dương; tồn âm từng SKU vẫn giữ để đối chiếu.
    if (s > 0) {
      totalStock += s;
      totalCbm += Math.max(0, vnnum_(row[cbmC]));
    }
  }
  const now = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  t.sh.getRange('A2').setValue('Cập nhật: ' + now + ' | SKU: ' + skuCount +
    ' | Tổng tồn: ' + totalStock + ' | CBM: ' + totalCbm.toFixed(3));
}

/************************************************
 * ĐỌC RIÊNG TỪNG NGUỒN
 ************************************************/

function vnreadSource_(kind, catalog, editedRange) {
  let table, lanes;
  const entries = [];
  const push = (lane, p, date, qty) => {
    const n = vnnum_(qty);
    if (n < 0) throw new Error('Số lượng âm tại ' + table.sh.getName());
    entries.push({ lane: lane, sku: p.sku, month: vnmonth_(date), qty: n });
  };

  if (kind === 'in' || kind === 'transfer') {
    const transfer = kind === 'transfer';
    const sheetName = transfer ? VNPT.sheets.transfer : VNPT.sheets.in;
    vnvalidateSheet_(sheetName);
    table = vnfillNames_(sheetName, catalog, editedRange);
    lanes = transfer ? ['hand', 'transfer'] : ['in'];
    const name = vncol_(table, 'Tên sản phẩm'), sku = vncol_(table, 'SUB SKU');
    const date = vncol_(table, 'Ngày'), qty = vncol_(table, 'Số lượng');
    const type = transfer ? vncol_(table, 'Loại xuất') : null;
    const status = transfer ? vncol_(table, 'Trạng thái') : null;
    table.rows.forEach((r, i) => {
      if (transfer && vnkey_(r[status]) === 'hủy') return;
      if (!vntext_(r[name]) && !vntext_(r[sku])) return;
      // Chưa có ngày: phiếu nháp, chưa phát sinh.
      if (vnblank_(r[date])) return;
      if (vnblank_(r[qty])) throw new Error(sheetName + ': thiếu số lượng dòng ' + (table.hr + 1 + i));
      let lane = 'in';
      if (transfer) {
        const t = vnkey_(r[type]);
        if (!['đóng tay', 'chuyển kho'].includes(t))
          throw new Error(sheetName + ': Loại xuất không hợp lệ dòng ' + (table.hr + 1 + i));
        lane = t === 'đóng tay' ? 'hand' : 'transfer';
      }
      // Cột Tình trạng không ảnh hưởng tồn.
      push(lane, vnproduct_(catalog, r[name], r[sku]), r[date], r[qty]);
    });

  } else if (kind === 'returns') {
    lanes = ['returns'];
    table = vntable_(VNPT.sheets.returns);
    vnrequire_(table, VNPT.required[VNPT.sheets.returns]);
    const name = vncol_(table, 'Tên sản phẩm'), date = vncol_(table, 'Ngày tháng');
    const qty = vncol_(table, 'Số lượng'), type = vncol_(table, 'Loại hoàn'), order = vncol_(table, 'Mã đơn');
    table.rows.forEach((r, i) => {
      if (vnkey_(r[type]) === 'ecom') return;
      if ([name, date, qty, order].every(c => vnblank_(r[c]))) return;
      if (!vntext_(r[name]) || vnblank_(r[date]) || vnblank_(r[qty]))
        throw new Error('Hoàn trả: thiếu tên/ngày/số lượng dòng ' + (table.hr + 1 + i));
      push('returns', vnproduct_(catalog, r[name], ''), r[date], r[qty]);
    });

  } else if (kind === 'ff') {
    // Tab Xuất ghi tay. Code chỉ tự điền SUB SKU/Tháng theo tên rồi cộng SL SP (ưu tiên SL SP APL).
    lanes = ['ff'];
    vnvalidateSheet_(VNPT.sheets.ff);
    table = vnfillNames_(VNPT.sheets.ff, catalog, editedRange);
    const name = vncol_(table, 'TÊN SẢN PHẨM'), sku = vncol_(table, 'SUB SKU'), date = vncol_(table, 'NGÀY');
    const qty = vncol_(table, 'SL SP'), apl = vncol_(table, 'SL SP APL');
    table.rows.forEach((r, i) => {
      if (r.every(v => vnblank_(v) || vnkey_(v) === 'chưa có')) return;
      if (!vntext_(r[name]) && !vntext_(r[sku])) return;
      const n = vneffective_(r[qty], r[apl]);
      if (vnblank_(r[date]) || vnblank_(n))
        throw new Error('Xuất: thiếu ngày/SL SP dòng ' + (table.hr + 1 + i));
      push('ff', vnproduct_(catalog, r[name], r[sku]), r[date], n);
    });

  } else throw new Error('Luồng không hợp lệ: ' + kind);

  return { entries: entries, lanes: lanes };
}

function vnrunUnlocked_(kind, catalog, editedRange) {
  vnsummary_(catalog); // đích phải hợp lệ trước khi tự điền nguồn
  const source = vnreadSource_(kind, catalog, editedRange);
  const state = vnstate_(), next = vnaggregate_(source.entries);
  for (const lane of source.lanes) {
    const previous = state.rows.filter(r => vntext_(r[0]) === lane).map(r => vntext_(r[1]));
    vnwriteLane_(catalog, lane, vntotals_(next, lane), previous);
  }
  vncommitState_(state, source.lanes, next);
  vnheadline_();
}

function vnrun_(kind, editedRange) {
  return vnguard_('vnrun_:' + kind, () => vnlock_(() => vnrunUnlocked_(kind, vncatalog_(), editedRange)));
}

function vnptRefreshNhap() { return vnrun_('in'); }
function vnptRefreshChuyenDongTay() { return vnrun_('transfer'); }
function vnptRefreshHoan() { return vnrun_('returns'); }
function vnptRefreshXuat() { return vnrun_('ff'); }

/************************************************
 * DANH MỤC / NVL / DÒNG TỔNG
 ************************************************/

function vnptRefreshDanhMuc() {
  return vnguard_('vnptRefreshDanhMuc', () => vnlock_(() => {
    const catalog = vncatalog_(), summary = vnsummary_(catalog), t = summary.t;
    const nameC = vncol_(t, 'TÊN SKU'), divisor = vnrollDivisor_(), edits = [];
    for (const [sku, ex] of summary.map) {
      const p = catalog.bySku.get(sku);
      if (!p) throw new Error('SKU trên Tổng hợp không còn trong danh mục: ' + sku);
      if (!vneq_(ex.values[nameC], p.name)) edits.push({ row: ex.row, col: nameC, value: p.name });
      vnderiveValues_(t, ex.values, p, true, divisor).forEach(([c, v]) => {
        if (!vneq_(ex.values[c], v)) edits.push({ row: ex.row, col: c, value: v });
      });
    }
    vnapplyCells_(t.sh, edits);
    vndropdowns_();
    vnheadline_();
  }));
}

/** Chỉ tính lại Tổng NVL / Số dự kiến (vd sau khi đổi số ở M3). */
function vnptRefreshNVL() {
  return vnguard_('vnptRefreshNVL', () => vnlock_(() => {
    const catalog = vncatalog_(), summary = vnsummary_(catalog), t = summary.t;
    const nvlC = vncol_(t, 'Tổng NVL'), rollC = vncol_(t, 'Số dự kiến');
    const divisor = vnrollDivisor_(), edits = [];
    for (const [sku, item] of summary.map) {
      const p = catalog.bySku.get(sku);
      if (!p) throw new Error('SKU không còn trong danh mục: ' + sku);
      vnderiveValues_(t, item.values, p, true, divisor)
        .filter(([c]) => c === nvlC || c === rollC)
        .forEach(([c, v]) => { if (!vneq_(item.values[c], v)) edits.push({ row: item.row, col: c, value: v }); });
    }
    vnapplyCells_(t.sh, edits);
    vnformatSummary_(t);
  }));
}

function vnptRefreshHeadline() {
  return vnguard_('vnptRefreshHeadline', () => vnlock_(() => vnheadline_()));
}

function vnformatSummary_(t) {
  const owned = vnss_().getRangeByName(VNPT.rangeAll);
  if (!owned) return;
  t.sh.getRange(owned.getRow(), vncol_(t, 'Tổng NVL') + 1, owned.getNumRows(), 1).setNumberFormat('0.######" m"');
  t.sh.getRange(owned.getRow(), vncol_(t, 'Số dự kiến') + 1, owned.getNumRows(), 1).setNumberFormat('0.0');
}

/************************************************
 * DROPDOWN TÊN SẢN PHẨM
 ************************************************/

function vndropdowns_() {
  const master = vntable_(VNPT.sheets.sku);
  const nameC = vncol_(master, 'Tên sản phẩm');
  const source = master.sh.getRange(master.hr + 1, nameC + 1, master.sh.getMaxRows() - master.hr, 1);
  const rule = SpreadsheetApp.newDataValidation().requireValueInRange(source, true).setAllowInvalid(false).build();

  [VNPT.sheets.in, VNPT.sheets.transfer, VNPT.sheets.ff].forEach(name => {
    const t = vntable_(name);
    const skuC = vncol_(t, 'SUB SKU'), prodC = vncol_(t, 'Tên sản phẩm');
    const rows = t.sh.getMaxRows() - t.hr;
    const skuRange = t.sh.getRange(t.hr + 1, skuC + 1, rows, 1);
    skuRange.clearDataValidations();
    skuRange.setNumberFormat('@');
    t.sh.getRange(t.hr + 1, prodC + 1, rows, 1).setDataValidation(rule);
  });
}

/************************************************
 * SỰ KIỆN SỬA SHEET
 ************************************************/

function vnptHandleEdit(e) {
  return vnguard_('vnptHandleEdit', () => {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet(), name = sheet.getName();
    const hr = VNPT.headers[name];
    if (!hr) return;

    if (name === VNPT.sheets.all) {
      // Chỉ phản ứng khi đổi số chia NVL ở M3; các ô khác do code quản lý.
      const cfg = sheet.getRange(VNPT.rollCell);
      if (e.range.getRow() <= cfg.getRow() && e.range.getLastRow() >= cfg.getRow() &&
          e.range.getColumn() <= cfg.getColumn() && e.range.getLastColumn() >= cfg.getColumn()) {
        return vnptRefreshNVL();
      }
      return;
    }

    if (e.range.getRow() <= hr && e.range.getLastRow() >= hr) {
      vnvalidateSheet_(name); // đổi nhầm header phải báo ngay
      if (e.range.getLastRow() <= hr) return;
    }

    if (name === VNPT.sheets.in) return vnrun_('in', e.range);
    if (name === VNPT.sheets.transfer) return vnrun_('transfer', e.range);
    if (name === VNPT.sheets.returns) return vnrun_('returns');
    if (name === VNPT.sheets.ff) return vnrun_('ff', e.range);
    if (name === VNPT.sheets.sku) return vnptRefreshDanhMuc();
  });
}

function vnptHandleChange(e) {
  // Sự kiện cấu trúc không chỉ rõ nguồn; các poll riêng sẽ bắt thay đổi.
  console.log('Thay đổi cấu trúc: ' + (e && e.changeType || 'manual') + '. Các luồng tự kiểm tra riêng.');
}

/************************************************
 * POLL RIÊNG TỪNG NGUỒN
 * (Hoàn trả do App Chứng Từ ghi qua API → onEdit không bắt được, poll sẽ bắt.)
 ************************************************/

function vnpoll_(sheetName, handler) {
  return vnguard_('vnpoll_:' + sheetName, () => {
    const t = vntable_(sheetName);
    const digest = Utilities.base64Encode(Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5, JSON.stringify([t.headers, t.rows]), Utilities.Charset.UTF_8));
    const props = PropertiesService.getScriptProperties();
    const key = 'VNPT_HASH_' + sheetName;
    if (props.getProperty(key) === digest) return;
    handler();
    // Đọc lại sau khi chạy (code có thể vừa tự điền SKU/Tháng) để không chạy lặp.
    const after = vntable_(sheetName);
    props.setProperty(key, Utilities.base64Encode(Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5, JSON.stringify([after.headers, after.rows]), Utilities.Charset.UTF_8)));
  });
}

function vnptPollNhap() { return vnpoll_(VNPT.sheets.in, vnptRefreshNhap); }
function vnptPollChuyenDongTay() { return vnpoll_(VNPT.sheets.transfer, vnptRefreshChuyenDongTay); }
function vnptPollHoan() { return vnpoll_(VNPT.sheets.returns, vnptRefreshHoan); }
function vnptPollXuat() { return vnpoll_(VNPT.sheets.ff, vnptRefreshXuat); }
function vnptPollSku() { return vnpoll_(VNPT.sheets.sku, vnptRefreshDanhMuc); }

/************************************************
 * ĐỐI SOÁT TẤT CẢ — CHỈ CHẠY KHI CHỦ ĐỘNG
 ************************************************/

function vnptRefreshAll() {
  return vnguard_('vnptRefreshAll', () => {
    const errors = [];
    [
      ['Nhập', vnptRefreshNhap],
      ['Xuất', vnptRefreshXuat],
      ['Đóng tay/chuyển kho', vnptRefreshChuyenDongTay],
      ['Hoàn', vnptRefreshHoan]
    ].forEach(([label, fn]) => {
      try { fn(); } catch (err) { errors.push(label + ': ' + err.message); }
    });
    if (errors.length) throw new Error('Các luồng khác đã chạy độc lập. Lỗi: ' + errors.join(' | '));
  });
}

/************************************************
 * KHỞI TẠO LẦN ĐẦU — XÓA DỮ LIỆU MẪU Ở TỔNG HỢP
 *
 * Chỉ xóa nội dung các dòng dưới header (A..L), giữ định dạng, header và ô M3.
 * Có công thức trong vùng này → dừng, không xóa.
 ************************************************/

function vnptKhoiTaoTongHop() {
  vnguard_('vnptKhoiTaoTongHop', () => vnlock_(() => {
    const ss = vnss_();
    const t = vntable_(VNPT.sheets.all);
    vnrequire_(t, VNPT.required[VNPT.sheets.all]);
    vncatalog_(); // danh mục lỗi thì dừng trước khi xóa gì

    if (t.rows.length) {
      const area = t.sh.getRange(t.hr + 1, 1, t.rows.length, t.width);
      if (area.getFormulas().some(r => r.some(Boolean)))
        throw new Error('Có công thức trong vùng dữ liệu Tổng hợp, không tự xóa: ' + area.getA1Notation());
      area.clearContent();
    }
    const named = ss.getRangeByName(VNPT.rangeAll);
    if (named) ss.removeNamedRange(VNPT.rangeAll);

    const st = ss.getSheetByName(VNPT.state);
    if (st && st.getLastRow() > 1) st.getRange(2, 1, st.getLastRow() - 1, 4).clearContent();

    const props = PropertiesService.getScriptProperties();
    Object.values(VNPT.sheets).forEach(n => props.deleteProperty('VNPT_HASH_' + n));

    t.sh.getRange('A1').setValue(VNPT.title);
  }));

  vnptRefreshAll();
  vnguard_('vnptKhoiTaoTongHop:format', () => vnlock_(() => vnformatSummary_(vntable_(VNPT.sheets.all))));
}

/************************************************
 * CÀI TRIGGER (project standalone)
 ************************************************/

function vnptInstallTriggers() {
  return vnguard_('vnptInstallTriggers', () => {
    if (typeof vnptAuthorizeAlerts === 'function') vnptAuthorizeAlerts();
    vnlock_(() => {
      const catalog = vncatalog_();
      const summary = vnsummary_(catalog);
      vnformatSummary_(summary.t);
      vnstate_();
      vndropdowns_();
    });

    const owned = [
      'vnptHandleEdit', 'vnptHandleChange',
      'vnptPollNhap', 'vnptPollChuyenDongTay', 'vnptPollHoan', 'vnptPollXuat', 'vnptPollSku'
    ];
    ScriptApp.getProjectTriggers().forEach(tr => {
      if (owned.includes(tr.getHandlerFunction())) ScriptApp.deleteTrigger(tr);
    });

    const ss = vnss_();
    ScriptApp.newTrigger('vnptHandleEdit').forSpreadsheet(ss).onEdit().create();
    ScriptApp.newTrigger('vnptHandleChange').forSpreadsheet(ss).onChange().create();
    ['vnptPollNhap', 'vnptPollChuyenDongTay', 'vnptPollHoan', 'vnptPollXuat', 'vnptPollSku']
      .forEach(h => ScriptApp.newTrigger(h).timeBased().everyMinutes(VNPT.pollMinutes).create());

    console.log('Đã cài dropdown và trigger VNPT. Mỗi luồng chạy riêng.');
  });
}

/************************************************
 * BÁO LỖI
 ************************************************/

function vnguard_(label, cb) {
  try { return cb(); }
  catch (error) {
    console.error('[VNPT] ' + label + ': ' + (error.stack || error));
    if (!error.vnptReported) {
      error.vnptReported = true;
      if (typeof vnptNotifyError === 'function') {
        try { vnptNotifyError('[VNPT] ' + label, error); }
        catch (mailErr) { console.error('Không gửi được cảnh báo: ' + mailErr); }
      } else console.error('Chưa có CanhBaoLoi_VNPT.gs: chưa gửi cảnh báo email.');
    }
    throw error;
  }
}

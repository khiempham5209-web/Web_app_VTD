/**
 * FFPL V4.2 — VẬN HÀNH SHEET THEO CÁC LUỒNG ĐỘC LẬP
 *
 * Đặt trong VanHanh.gs, cùng project với API.gs.
 * Thay toàn bộ code vận hành cũ; KHÔNG thay API.gs.
 * Code sync giữ ở project riêng.
 * Thêm CanhBaoLoi.gs cùng project để báo lỗi qua email.
 * Định mức SKU tính bằng cm; Tổng NVL xuất ra mét; cuộn mặc định 100 m, giữ số lẻ.
 *
 * Không tạo bảng báo cáo tháng hoặc ô chọn tháng.
 * Không xóa toàn bộ bảng khi cập nhật.
 * Tìm các cột nghiệp vụ theo header.
 *
 * API cũ gọi refreshTongHopAll() sau khi lưu hoàn.
 * Hàm này được giữ tên và chỉ cập nhật luồng HOÀN.
 *
 * refreshAll() chỉ dành cho đối soát chủ động.
 * Sửa dữ liệu thông thường không gọi refreshAll().
 */

const FFPL3 = {
  id: '1TVck35ibO3zJ6lkKpHNyXckgLyjvUxEoLbMa-6kt-Pk',

  sheets: {
    all: 'Tổng hợp all',
    sku: 'Thông tin SKU',
    in: 'Nhập NCC',
    ff: 'BC_Xuất FF',
    orders: 'DS đơn all',
    transfer: 'Đóng tay_Chuyển kho',
    returns: 'Hàng hoàn',
    check: 'Kiểm kê'
  },

  headers: {
    'Tổng hợp all': 3,
    'Thông tin SKU': 3,
    'Nhập NCC': 3,
    'BC_Xuất FF': 3,
    'DS đơn all': 1,
    'Đóng tay_Chuyển kho': 3,
    'Hàng hoàn': 3,
    'Kiểm kê': 2
  },

  state: '_FFPL_V3_STATE',

  minStock: 1000,
  bubbleRollMeters: 100,
  pollMinutes: 5,

  laneHeaders: {
    in: 'NHẬP',
    ff: 'XUẤT FF',
    hand: 'ĐÓNG TAY',
    transfer: 'CHUYỂN KHO',
    returns: 'HOÀN'
  }
};

/************************************************
 * HÀM CƠ BẢN
 ************************************************/

function f3ss_() {
  return SpreadsheetApp.openById(FFPL3.id);
}

function f3text_(value) {
  return value === null || value === undefined
    ? ''
    : String(value).trim();
}

function f3key_(value) {
  return f3text_(value)
    .normalize('NFC')
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

function f3num_(value) {
  if (
    value === '' ||
    value === null ||
    value === undefined
  ) {
    return 0;
  }

  if (
    typeof value === 'number' &&
    Number.isFinite(value)
  ) {
    return value;
  }

  const text = f3text_(value);

  const number = text.includes(',')
    ? Number(text.replace(/\./g, '').replace(',', '.'))
    : Number(text);

  if (!Number.isFinite(number)) {
    throw new Error('Số lượng không hợp lệ: ' + text);
  }

  return number;
}

function f3date_(value) {
  if (value instanceof Date && !isNaN(value)) return value;
  const text=f3text_(value);
  // ISO có timezone phải được hiểu là thời điểm, rồi quy về ngày theo timezone của sheet.
  if (/^\d{4}-\d{2}-\d{2}T.*(?:Z|[+-]\d{2}:?\d{2})$/.test(text)) {
    const instant=new Date(text);
    if (!isNaN(instant)) return instant;
  }
  let match=text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s|$)/),key;
  if (match) key=match[3]+'-'+match[2].padStart(2,'0')+'-'+match[1].padStart(2,'0');
  else {
    match=text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s]|$)/);
    if (match) key=match[1]+'-'+match[2].padStart(2,'0')+'-'+match[3].padStart(2,'0');
  }
  if (!key) throw new Error('Ngày phải là Date hoặc dd/MM/yyyy: '+text);
  const tz=f3ss_().getSpreadsheetTimeZone();
  const date=Utilities.parseDate(key,tz,'yyyy-MM-dd');
  if (isNaN(date) || Utilities.formatDate(date,tz,'yyyy-MM-dd')!==key)
    throw new Error('Ngày không hợp lệ: '+text);
  return date;
}

function f3month_(value) {
  return Utilities.formatDate(
    f3date_(value),
    f3ss_().getSpreadsheetTimeZone(),
    'yyyy-MM'
  );
}

function f3day_(value) {
  return Utilities.formatDate(
    f3date_(value),
    f3ss_().getSpreadsheetTimeZone(),
    'yyyy-MM-dd'
  );
}

function f3eq_(oldValue, newValue) {
  if (
    oldValue instanceof Date &&
    newValue instanceof Date
  ) {
    return oldValue.getTime() === newValue.getTime();
  }

  return oldValue === newValue;
}

/* Khóa NGẮN chỉ cho khoảnh khắc ghi vào tab chung (Tổng hợp all, STATE, Kiểm kê, BC_Xuất FF).
 * Phần đọc/tính của từng luồng chạy ngoài khóa nên các luồng không chặn nhau.
 * Đến lượt ghi mà luồng khác đang ghi -> chờ (tối đa 2 phút), không báo lỗi ngay. */
function f3lock_(callback) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(120000)) {
    throw new Error(
      'Chờ ghi Tổng hợp quá 2 phút, thử lại sau.'
    );
  }

  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

/* Mỗi luồng ghi lại thời điểm BẮT ĐẦU đọc nguồn của lượt đã ghi xong.
 * Lượt đọc nguồn sớm hơn lượt đã ghi -> dữ liệu cũ hơn -> bỏ qua, không ghi đè số mới. */
function f3isStale_(kind, started) {
  const done = Number(PropertiesService.getScriptProperties().getProperty('FFPL3_DONE_' + kind) || 0);
  return done > started;
}

function f3markDone_(kind, started) {
  PropertiesService.getScriptProperties().setProperty('FFPL3_DONE_' + kind, String(started));
}

function f3ensureGrid_(sheet, lastRow, lastColumn) {
  if (sheet.getMaxRows() < lastRow) {
    sheet.insertRowsAfter(
      sheet.getMaxRows(),
      lastRow - sheet.getMaxRows()
    );
  }

  if (sheet.getMaxColumns() < lastColumn) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      lastColumn - sheet.getMaxColumns()
    );
  }
}

/************************************************
 * HEADER VÀ ĐỌC BẢNG
 ************************************************/

function f3map_(headers, location) {
  const map = {};

  headers.forEach((header, index) => {
    const key = f3key_(header);

    if (!key) return;

    if (
      Object.prototype.hasOwnProperty.call(map, key)
    ) {
      throw new Error(
        'Trùng header "' + header + '" ở ' + location
      );
    }

    map[key] = index;
  });

  return map;
}

function f3col_(table, headerName, optional) {
  const column = table.h[f3key_(headerName)];

  if (column === undefined && !optional) {
    throw new Error(
      'Thiếu header "' +
      headerName +
      '" tại ' +
      table.sh.getName()
    );
  }

  return column;
}

function f3table_(sheetName) {
  const sheet = f3ss_().getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Thiếu tab ' + sheetName);
  }

  const headerRow = FFPL3.headers[sheetName];

  if (!headerRow) {
    throw new Error(
      'Chưa cấu hình hàng header: ' + sheetName
    );
  }

  const rawHeaders = sheet
    .getRange(
      headerRow,
      1,
      1,
      sheet.getLastColumn()
    )
    .getValues()[0];

  let width = rawHeaders.length;

  while (
    width &&
    !f3text_(rawHeaders[width - 1])
  ) {
    width--;
  }

  if (!width) {
    throw new Error('Header trống: ' + sheetName);
  }

  const headers = rawHeaders.slice(0, width);
  const headerMap = f3map_(headers, sheetName);

  const lastRow = sheet.getLastRow();

  const rows = lastRow > headerRow
    ? sheet
        .getRange(
          headerRow + 1,
          1,
          lastRow - headerRow,
          width
        )
        .getValues()
    : [];

  return {
    sh: sheet,
    hr: headerRow,
    width: width,
    headers: headers,
    h: headerMap,
    rows: rows
  };
}

function f3set_(sheet, row, columnIndex, value) {
  const range = sheet.getRange(
    row,
    columnIndex + 1
  );

  if (range.getFormula()) {
    throw new Error(
      'Không ghi đè công thức: ' +
      sheet.getName() +
      '!' +
      range.getA1Notation()
    );
  }

  if (!f3eq_(range.getValue(), value)) {
    range.setValue(value);
  }
}

/************************************************
 * DANH MỤC SKU
 ************************************************/

function f3catalog_() {
  const table = f3table_(FFPL3.sheets.sku);

  const skuColumn = f3col_(table, 'SUB SKU');
  const nameColumn = f3col_(table, 'Tên sản phẩm');
  const unitColumn = f3col_(table, 'Đơn vị');
  const sizeColumn = f3col_(table, 'Kích thước');
  const normColumn = f3col_(table, 'Tổng mức tiêu hao');

  const bySku = new Map();
  const byName = new Map();

  table.rows.forEach(row => {
    const sku = f3text_(row[skuColumn]);
    const name = f3text_(row[nameColumn]);

    if (!sku && !name) return;

    if (!sku || !name) {
      throw new Error(
        'Danh mục thiếu mã hoặc tên SKU.'
      );
    }

    if (
      bySku.has(sku) ||
      byName.has(f3key_(name))
    ) {
      throw new Error(
        'Trùng mã/tên trong Thông tin SKU: ' + name
      );
    }

    const sizeText = f3text_(row[sizeColumn])
      .replace(/\s/g, '');

    const dimensions = sizeText.match(
      /^(\d+(?:[.,]\d+)?)[x×](\d+(?:[.,]\d+)?)[x×](\d+(?:[.,]\d+)?)$/i
    );

    if (sizeText && !dimensions) {
      throw new Error(
        'Kích thước không hợp lệ: ' + name
      );
    }

    const cbm = dimensions
      ? dimensions
          .slice(1)
          .map(value =>
            Number(value.replace(',', '.'))
          )
          .reduce(
            (product, value) => product * value,
            1
          ) / 1000000
      : null;

    if (!f3blank_(row[normColumn]) && f3num_(row[normColumn]) < 0)
      throw new Error('Định mức âm tại SKU: '+name);
    const product = {
      sku: sku,
      rawSku: row[skuColumn],
      name: name,
      unit: row[unitColumn],
      norm: f3blank_(row[normColumn]) ? null : f3num_(row[normColumn]),
      cbm: cbm
    };

    bySku.set(sku, product);
    byName.set(f3key_(name), product);
  });

  return {
    bySku: bySku,
    byName: byName
  };
}

function f3product_(catalog, name, sku) {
  const productName = f3text_(name);
  const productSku = f3text_(sku);

  const product = productName
    ? catalog.byName.get(f3key_(productName))
    : catalog.bySku.get(productSku);

  if (!product) {
    throw new Error(
      'Không tìm thấy sản phẩm trong danh mục: ' +
      (productName || productSku)
    );
  }

  return product;
}

/************************************************
 * CHỌN TÊN → TỰ ĐIỀN SUB SKU / ĐƠN VỊ / THÁNG
 ************************************************/

function f3fillNames_(sheetName, catalog, editedRange) {
  const table = f3table_(sheetName);

  const skuColumn = f3col_(table, 'SUB SKU');
  const nameColumn = f3col_(table, 'Tên sản phẩm');

  const unitColumn = f3col_(
    table,
    'Đơn vị',
    true
  );

  const monthColumn = f3col_(
    table,
    'Tháng',
    true
  );

  const dateColumn = f3col_(
    table,
    sheetName === FFPL3.sheets.check
      ? 'Ngày kiểm'
      : 'Ngày',
    true
  );

  const editedName = editedRange &&
    editedRange.getColumn() <= nameColumn + 1 &&
    editedRange.getLastColumn() >= nameColumn + 1;

  const plan = [];

  // Chỉ đưa vào kế hoạch ô thật sự khác giá trị hiện có (quét cả tab mà vẫn ghi ít).
  const changedOnly = (row, cells) => cells.filter(([column, value]) => !f3eq_(row[column], value));

  table.rows.forEach((row, index) => {
    const sheetRow = table.hr + 1 + index;

    // Sửa khi đang lọc: Google chỉ báo vùng liền đầu tiên. Vẫn quét CẢ tab để điền bù
    // các dòng có Tên mà thiếu SKU/Đơn vị/Tháng; ngoài vùng vừa sửa thì không đụng dòng không có Tên.
    const inEdit = !editedRange || (
      sheetRow >= editedRange.getRow() &&
      sheetRow <= editedRange.getLastRow()
    );

    const statusColumn = sheetName === FFPL3.sheets.transfer ? f3col_(table,'Trạng thái') : undefined;
    if (statusColumn !== undefined && f3key_(row[statusColumn]) === 'hủy') return;
    const name = f3text_(row[nameColumn]);
    const sku = f3text_(row[skuColumn]);

    if (!name && !sku) return;
    if (!inEdit && !name) return;

    // Người dùng xóa tên: xóa mã và đơn vị tương ứng.
    if (inEdit && editedName && !name) {
      const cells = [[skuColumn, '']];

      if (unitColumn !== undefined) {
        cells.push([unitColumn, '']);
      }

      const changed = changedOnly(row, cells);
      if (changed.length) plan.push([sheetRow, changed]);
      return;
    }

    const product = f3product_(
      catalog,
      name,
      sku
    );

    const cells = [
      [skuColumn, product.rawSku],
      [nameColumn, product.name]
    ];

    if (unitColumn !== undefined) {
      cells.push([unitColumn, product.unit]);
    }

    if (
      monthColumn !== undefined &&
      dateColumn !== undefined &&
      row[dateColumn] !== ''
    ) {
      const date = f3date_(row[dateColumn]);

      cells.push([
        monthColumn,
        f3date_(f3month_(date)+'-01')
      ]);
    }

    const changed = changedOnly(row, cells);
    if (changed.length) plan.push([sheetRow, changed]);
  });

  // Chỉ ghi sau khi đã kiểm tra toàn bộ các dòng cần xử lý.
  plan.forEach(([sheetRow, cells]) => {
    cells.forEach(([column, value]) => {
      f3set_(
        table.sh,
        sheetRow,
        column,
        value
      );
    });
  });

  return f3table_(sheetName);
}

/************************************************
 * TRẠNG THÁI PHÁT SINH CỦA TỪNG LUỒNG
 *
 * Dùng để nhận biết dữ liệu bị sửa/xóa/chuyển SKU.
 * Không phải bảng báo cáo tháng.
 ************************************************/

function f3state_() {
  const ss = f3ss_();
  let sheet = ss.getSheetByName(FFPL3.state);

  const headers = [
    'Luồng',
    'SUB SKU',
    'Tháng',
    'Số lượng'
  ];

  if (!sheet) {
    sheet = ss.insertSheet(FFPL3.state);

    sheet
      .getRange(1, 1, 1, headers.length)
      .setValues([headers]);

    sheet.hideSheet();
  } else {
    const existingHeaders = sheet
      .getRange(1, 1, 1, headers.length)
      .getValues()[0];

    if (
      JSON.stringify(existingHeaders) !==
      JSON.stringify(headers)
    ) {
      throw new Error(
        'Tab trạng thái không đúng cấu trúc. Không ghi đè.'
      );
    }
  }

  const rows = sheet.getLastRow() > 1
    ? sheet
        .getRange(
          2,
          1,
          sheet.getLastRow() - 1,
          headers.length
        )
        .getValues()
    : [];

  return {
    sh: sheet,
    rows: rows,
    h: f3map_(headers, FFPL3.state)
  };
}

function f3aggregate_(entries) {
  const map = new Map();

  entries.forEach(entry => {
    const key = JSON.stringify([
      entry.lane,
      entry.sku,
      entry.month
    ]);

    const item = map.get(key) || {
      lane: entry.lane,
      sku: entry.sku,
      month: entry.month,
      qty: 0
    };

    item.qty += entry.qty;
    map.set(key, item);
  });

  return map;
}

function f3stateKey_(row) {
  return JSON.stringify([f3text_(row[0]),f3text_(row[1]),f3stateMonth_(row[2])]);
}

function f3totals_(entries, lane) {
  const totals = new Map();

  for (const entry of entries.values()) {
    if (entry.lane !== lane) continue;

    totals.set(
      entry.sku,
      (totals.get(entry.sku) || 0) + entry.qty
    );
  }

  return totals;
}

function f3commitState_(state,lanes,next) {
  const old=new Map(),duplicates=[],edits=[],pending=[];
  // Chỉ quản lý trạng thái của luồng đang chạy; lỗi/khóa của luồng khác không chặn.
  state.rows.forEach((row,index)=>{
    if (!lanes.includes(f3text_(row[0]))) return;
    const key=f3stateKey_(row);
    if (old.has(key)) duplicates.push(old.get(key));
    old.set(key,{r:row,i:index}); // giữ vị trí cuối; số lượng luôn lấy lại từ nguồn hiện tại
  });
  for (const [key,entry] of next) {
    if (!lanes.includes(entry.lane)) continue;
    const month='m:'+entry.month; // tiền tố chống Sheets tự đổi YYYY-MM thành ngày
    const record=[entry.lane,entry.sku,month,entry.qty],existing=old.get(key);
    if (existing) {
      if (!f3eq_(existing.r[2],month)) edits.push({row:existing.i+2,col:2,value:month});
      if (!f3eq_(existing.r[3],entry.qty)) edits.push({row:existing.i+2,col:3,value:entry.qty});
      old.delete(key);
    } else pending.push(record);
  }
  for (const existing of old.values()) {
    const month='m:'+f3stateMonth_(existing.r[2]);
    if (!f3eq_(existing.r[2],month)) edits.push({row:existing.i+2,col:2,value:month});
    if (!f3eq_(existing.r[3],0)) edits.push({row:existing.i+2,col:3,value:0});
  }
  // Chỉ làm trống đúng bản ghi trùng trong tab kỹ thuật, không xóa nguồn hoặc bảng tổng.
  duplicates.forEach(existing=>existing.r.slice(0,4).forEach((v,col)=>{
    if (v!=='') edits.push({row:existing.i+2,col:col,value:''});
  }));
  f3applyCells_(state.sh,edits);
  if (pending.length) {
    const start=state.sh.getLastRow()+1;
    f3ensureGrid_(state.sh,start+pending.length-1,4);
    state.sh.getRange(start,1,pending.length,4).setValues(pending);
  }
}

function f3region_(table, rangeName, keyHeaders, accept) {
  const ss = f3ss_();
  let range = ss.getRangeByName(rangeName);

  if (range) {
    if (
      range.getSheet().getSheetId() !==
        table.sh.getSheetId() ||
      range.getRow() !== table.hr + 1 ||
      range.getColumn() !== 1
    ) {
      throw new Error(
        'Vùng ' +
        rangeName +
        ' lệch header. Cần kiểm tra ranh giới, không tự ghi đè.'
      );
    }

    // Khi thêm Tổng NVL/Số dự kiến, chỉ mở rộng bề ngang cùng vùng dòng.
    if (range.getNumColumns() !== table.width) {
      range = table.sh.getRange(table.hr+1,1,range.getNumRows(),table.width);
      ss.setNamedRange(rangeName,range);
    }
    return range;
  }

  const keyColumns = keyHeaders.map(
    header => f3col_(table, header)
  );

  let count = 0;

  for (const row of table.rows) {
    if (
      !keyColumns.every(
        column => f3text_(row[column])
      ) ||
      !accept(row)
    ) {
      break;
    }

    count++;
  }

  if (
    !count &&
    table.rows[0] &&
    table.rows[0].some(value => value !== '')
  ) {
    throw new Error(
      'Nội dung lạ ngay đầu ' +
      table.sh.getName() +
      '. Dừng để bảo vệ bảng.'
    );
  }

  f3ensureGrid_(
    table.sh,
    table.hr + Math.max(1, count),
    table.width
  );

  range = table.sh.getRange(
    table.hr + 1,
    1,
    Math.max(1, count),
    table.width
  );

  ss.setNamedRange(rangeName, range);

  return range;
}

function f3appendOwned_(table, rangeName, values) {
  const ss = f3ss_();
  const range = ss.getRangeByName(rangeName);

  const existingValues = range.getValues();

  const emptyIndex = existingValues.findIndex(
    row => row.every(value => value === '')
  );

  if (emptyIndex >= 0) {
    const sheetRow = range.getRow() + emptyIndex;

    values.forEach((value, column) => {
      if (value !== '' && value !== undefined) {
        f3set_(
          table.sh,
          sheetRow,
          column,
          value
        );
      }
    });

    return sheetRow;
  }

  const insertRow = range.getLastRow() + 1;

  if (insertRow > table.sh.getMaxRows()) {
    table.sh.insertRowsAfter(insertRow - 1, 1);
  } else {
    table.sh.insertRowsBefore(insertRow, 1);
  }

  table.sh
    .getRange(
      insertRow - 1,
      1,
      1,
      table.width
    )
    .copyTo(
      table.sh.getRange(
        insertRow,
        1,
        1,
        table.width
      ),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );

  table.sh
    .getRange(
      insertRow,
      1,
      1,
      table.width
    )
    .setValues([values]);

  ss.setNamedRange(
    rangeName,
    table.sh.getRange(
      table.hr + 1,
      1,
      existingValues.length + 1,
      table.width
    )
  );

  return insertRow;
}

/************************************************
 * BẢNG TỔNG HỢP ALL
 ************************************************/

function f3summary_(catalog) {
  const table = f3table_(FFPL3.sheets.all);

  [
    'SUB SKU',
    'TÊN SKU',
    'NHẬP',
    'XUẤT FF',
    'ĐÓNG TAY',
    'CHUYỂN KHO',
    'HOÀN',
    'TỒN',
    'CBM',
    'TÌNH TRẠNG TỒN KHO', 'Tổng NVL', 'Số dự kiến'
  ].forEach(header => {
    f3col_(table, header);
  });

  const skuColumn = f3col_(table, 'SUB SKU');

  const range = f3region_(
    table,
    'FFPL_V3_ALL',
    ['SUB SKU'],
    row => catalog.bySku.has(
      f3text_(row[skuColumn])
    )
  );

  const map = new Map();

  range.getValues().forEach((row, index) => {
    const sku = f3text_(row[skuColumn]);

    if (!sku) return;

    if (map.has(sku)) {
      throw new Error(
        'Trùng SKU trên bảng all: ' + sku
      );
    }

    map.set(sku, {
      row: table.hr + 1 + index,
      values: row
    });
  });

  return {
    t: table,
    map: map
  };
}

function f3writeLane_(catalog,lane,totals,previousKeys) {
  const summary=f3summary_(catalog), table=summary.t;
  const skuCol=f3col_(table,'SUB SKU'), nameCol=f3col_(table,'TÊN SKU');
  const target=f3col_(table,FFPL3.laneHeaders[lane]);
  const keys=new Set([...totals.keys(),...previousKeys,...summary.map.keys()]);
  const edits=[], additions=[], affected=new Set();
  for (const sku of keys) {
    const p=catalog.bySku.get(sku);
    if (!p) throw new Error('SKU bị xóa khỏi danh mục: '+sku);
    const existing=summary.map.get(sku);
    if (!existing && !totals.has(sku)) continue;
    const values=existing ? existing.values.slice() : new Array(table.width).fill('');
    if (!existing) {
      values[skuCol]=p.rawSku; values[nameCol]=p.name;
      Object.values(FFPL3.laneHeaders).forEach(h=>values[f3col_(table,h)]=0);
    }
    const quantity=totals.get(sku)||0;
    const changed=!existing || !f3eq_(values[target],quantity);
    values[target]=quantity;
    const computed=f3deriveValues_(table,values,p,!existing || lane==='ff' || lane==='hand');
    computed.forEach(([col,v])=>values[col]=v);
    if (existing) {
      [target,...computed.map(x=>x[0])].forEach(col=>{
        if (!f3eq_(existing.values[col],values[col]))
          edits.push({row:existing.row,col:col,value:values[col]});
      });
    } else additions.push(values);
    if (changed) affected.add(sku);
  }
  // Kiểm tra tất cả ô sẽ ghi trước khi ghi, giữ nguyên ô thủ công/công thức.
  f3applyCells_(table.sh,edits);
  additions.forEach(row=>f3appendOwned_(table,'FFPL_V3_ALL',row));
  return affected;
}

function f3derived_(table,sheetRow,values,product,materials) {
  const edits=f3deriveValues_(table,values,product,materials)
    .filter(([col,v])=>!f3eq_(values[col],v))
    .map(([col,v])=>({row:sheetRow,col:col,value:v}));
  f3applyCells_(table.sh,edits);
}

function f3headline_() {
  const ss = f3ss_();
  const table = f3table_(FFPL3.sheets.all);
  const range = ss.getRangeByName('FFPL_V3_ALL');

  if (!range) return;

  const skuColumn = f3col_(table, 'SUB SKU');
  const stockColumn = f3col_(table, 'TỒN');
  const cbmColumn = f3col_(table, 'CBM');

  let skuCount = 0;
  let totalStock = 0;
  let totalCbm = 0;

  for (const row of range.getValues()) {
    if (!f3text_(row[skuColumn])) continue;

    skuCount++;
    const stock = f3num_(row[stockColumn]);
    // Tổng phía trên chỉ phản ánh lượng hàng tồn dương.
    // Giữ nguyên tồn âm từng SKU để đối chiếu, không bù trừ sang SKU khác.
    if (stock > 0) {
      totalStock += stock;
      totalCbm += Math.max(0, f3num_(row[cbmColumn]));
    }
  }

  const now = Utilities.formatDate(
    new Date(),
    ss.getSpreadsheetTimeZone(),
    'dd/MM/yyyy HH:mm:ss'
  );

  // A2 là ô chú thích trình bày đã có,
  // không phải vị trí cột dữ liệu nghiệp vụ.
  table.sh.getRange('A2').setValue(
    'Cập nhật: ' +
    now +
    ' | SKU: ' +
    skuCount +
    ' | Tổng tồn: ' +
    totalStock +
    ' | CBM: ' +
    totalCbm.toFixed(3)
  );
}

/************************************************
 * ĐỌC RIÊNG TỪNG NGUỒN
 ************************************************/

function f3readSource_(kind, catalog, editedRange) {
  let table, lanes;
  const entries = [];
  const push = (lane,p,date,qty) => {
    const n = f3num_(qty);
    if (n < 0) throw new Error('Số lượng âm tại ' + table.sh.getName());
    entries.push({lane:lane,sku:p.sku,month:f3month_(date),qty:n});
  };
  if (kind === 'in' || kind === 'transfer') {
    const transfer = kind === 'transfer';
    const sheetName = transfer ? FFPL3.sheets.transfer : FFPL3.sheets.in;
    table = f3table_(sheetName);
    f3require_(table, ['SUB SKU','Tên sản phẩm','Ngày','Số lượng'].concat(transfer ? ['Loại xuất','Trạng thái'] : []));
    table = f3fillNames_(sheetName,catalog,editedRange);
    lanes = transfer ? ['hand','transfer'] : ['in'];
    const name = f3col_(table,'Tên sản phẩm'), sku = f3col_(table,'SUB SKU');
    const date = f3col_(table,'Ngày'), qty = f3col_(table,'Số lượng');
    const type = transfer ? f3col_(table,'Loại xuất') : null;
    const status = transfer ? f3col_(table,'Trạng thái') : null;
    table.rows.forEach((r,i) => {
      if (transfer && f3key_(r[status]) === 'hủy') return;
      if (!f3text_(r[name]) && !f3text_(r[sku])) return;
      // Chưa có ngày: phiếu nháp chưa phát sinh nhập/xuất.
      if (f3blank_(r[date])) return;
      if (f3blank_(r[qty])) throw new Error(sheetName + ': thiếu số lượng dòng ' + (table.hr+1+i));
      let lane = 'in';
      if (transfer) {
        const t = f3key_(r[type]);
        if (!['đóng tay','chuyển kho'].includes(t)) throw new Error(sheetName + ': Loại xuất không hợp lệ dòng ' + (table.hr+1+i));
        lane = t === 'đóng tay' ? 'hand' : 'transfer';
      }
      push(lane,f3product_(catalog,r[name],r[sku]),r[date],r[qty]);
    });
  } else if (kind === 'returns') {
    lanes = ['returns']; table = f3table_(FFPL3.sheets.returns);
    f3require_(table,['Tên sản phẩm','Ngày tháng','Số lượng','Mã đơn','Loại hoàn']);
    const name=f3col_(table,'Tên sản phẩm'),date=f3col_(table,'Ngày tháng');
    const qty=f3col_(table,'Số lượng'),type=f3col_(table,'Loại hoàn'),order=f3col_(table,'Mã đơn');
    table.rows.forEach((r,i) => {
      if (f3key_(r[type]) === 'ecom') return;
      if ([name,date,qty,order].every(col=>f3blank_(r[col]))) return;
      if (!f3text_(r[name]) || f3blank_(r[date]) || f3blank_(r[qty]))
        throw new Error('Hàng hoàn: thiếu tên/ngày/số lượng dòng ' + (table.hr+1+i));
      push('returns',f3product_(catalog,r[name],''),r[date],r[qty]);
    });
  } else if (kind === 'ff') {
    lanes=['ff']; table=f3table_(FFPL3.sheets.ff);
    f3require_(table,['TÊN SẢN PHẨM','SUB SKU','NGÀY','SL đơn','SL SP','SL đơn APL','SL SP APL']);
    const name=f3col_(table,'TÊN SẢN PHẨM'),sku=f3col_(table,'SUB SKU'),date=f3col_(table,'NGÀY');
    const qty=f3col_(table,'SL SP'),apl=f3col_(table,'SL SP APL');
    const orders=f3col_(table,'SL đơn'),ordersApl=f3col_(table,'SL đơn APL');
    table.rows.forEach((r,i) => {
      if (r.every(v => f3blank_(v) || f3key_(v) === 'chưa có')) return;
      if (!f3text_(r[name]) && !f3text_(r[sku])) return;
      const n=f3effective_(r[qty],r[apl]);
      const orderCount=f3effective_(r[orders],r[ordersApl]);
      if (f3blank_(r[date]) || f3blank_(n) || f3blank_(orderCount))
        throw new Error('BC xuất: thiếu ngày/số lượng dòng ' + (table.hr+1+i));
      if (f3num_(orderCount)<0) throw new Error('BC xuất: số đơn âm dòng '+(table.hr+1+i));
      // Số đơn thực tế và SL SP thực tế độc lập; không suy ra PCS từ số đơn.
      push('ff',f3product_(catalog,r[name],r[sku]),r[date],n);
    });
  } else throw new Error('Luồng không hợp lệ: '+kind);
  return {entries:entries,lanes:lanes};
}

/* Một luồng = một tab nguồn. Đọc + tự điền + cộng tổng chạy NGOÀI khóa (luồng này không chặn luồng khác);
 * chỉ bước ghi vào Tổng hợp all / STATE / Kiểm kê nằm trong khóa ngắn. */
function f3run_(kind,editedRange) {
  return ffplGuard_("f3run_:"+kind,()=>{
    const started=Date.now();
    const catalog=f3catalog_();
    // Đích phải hợp lệ trước khi tự điền dữ liệu nguồn.
    f3summary_(catalog);
    const source=f3readSource_(kind,catalog,editedRange);
    const next=f3aggregate_(source.entries);
    return f3commitLanes_(kind,catalog,source.lanes,next,started);
  });
}

function f3commitLanes_(kind,catalog,lanes,next,started) {
  return f3lock_(()=>{
    // Có lượt đọc nguồn mới hơn của CÙNG luồng đã ghi xong -> bỏ qua để không ghi đè số mới bằng số cũ.
    if (f3isStale_(kind,started)) return {ok:true,skipped:'stale',kind:kind};
    // Đọc Tổng hợp/STATE ngay trong khóa: TỒN tính từ số mới nhất của các luồng khác.
    const state=f3state_(), affected=new Set();
    for (const lane of lanes) {
      const previous=state.rows.filter(r=>f3text_(r[0])===lane).map(r=>f3text_(r[1]));
      f3writeLane_(catalog,lane,f3totals_(next,lane),previous).forEach(sku=>affected.add(sku));
    }
    f3commitState_(state,lanes,next);
    f3headline_();
    // Kiểm kê là bên đọc tồn: cập nhật đối chiếu cho SKU bị ảnh hưởng, không sửa tồn.
    if (affected.size) f3refreshCheckForSkus_(catalog,affected);
    f3markDone_(kind,started);
    return {ok:true,kind:kind,affected:affected.size};
  });
}

function ffplRefreshNhap() {
  return f3run_('in');
}

function ffplRefreshChuyenDongTay() {
  return f3run_('transfer');
}

function ffplRefreshHoan() {
  return f3run_('returns');
}

function ffplRefreshXuatTuBC() {
  return f3run_('ff');
}

/**
 * Giữ tên hàm để API cũ gọi sau khi lưu hoàn.
 * Chỉ cập nhật HOÀN, không chạy nhập/xuất/chuyển kho.
 */
function refreshTongHopAll() {
  return ffplRefreshHoan();
}

/************************************************
 * DS ĐƠN ALL → BC XUẤT FF
 *
 * Đây chỉ là luồng XUẤT.
 * Không chạy lại nhập/hoàn/đóng tay/chuyển kho.
 ************************************************/

function f3buildBC_(catalog,groupedIn) {
  const dest=f3table_(FFPL3.sheets.ff);
  f3require_(dest,['THÁNG','NGÀY','SUB SKU','TÊN SẢN PHẨM','SL đơn','SL SP','SL đơn APL','SL SP APL','GHI CHÚ']);
  // groupedIn: đã gom DS đơn all ngoài khóa; trong khóa chỉ đọc BC và ghi phần khác.
  const grouped=groupedIn || f3groupOrders_(f3table_(FFPL3.sheets.orders),catalog).groups;
  const skuCol=f3col_(dest,'SUB SKU'),dateCol=f3col_(dest,'NGÀY');
  // Chỉ thay placeholder cũ khi toàn dòng là trống hoặc "chưa có".
  const first=dest.rows[0];
  if (first && first.some(v=>v!=='') && first.every(v=>v==='' || f3key_(v)==='chưa có')) {
    f3applyCells_(dest.sh,first.map((v,col)=>({row:dest.hr+1,col:col,value:''})).filter((e,i)=>first[i]!==''));
  }
  const current=f3table_(FFPL3.sheets.ff);
  const range=f3region_(current,'FFPL_V3_BC',['NGÀY','SUB SKU'],r=>catalog.bySku.has(f3text_(r[skuCol])));
  const old=new Map();
  range.getValues().forEach((r,i)=>{
    if (!f3text_(r[skuCol])) return;
    const key=JSON.stringify([f3day_(r[dateCol]),f3text_(r[skuCol])]);
    if (old.has(key)) throw new Error('Trùng ngày/SKU trong BC xuất.');
    old.set(key,{values:r,row:current.hr+1+i});
  });
  const edits=[],adds=[];
  const headers=['THÁNG','NGÀY','SUB SKU','TÊN SẢN PHẨM','SL đơn','SL SP'];
  for (const [key,item] of grouped) {
    const values={
      'THÁNG':f3date_(item.day.slice(0,7)+'-01'),
      'NGÀY':f3date_(item.day),
      'SUB SKU':item.p.rawSku,'TÊN SẢN PHẨM':item.p.name,
      'SL đơn':item.ods.size,'SL SP':item.qty
    };
    const existing=old.get(key);
    if (existing) {
      headers.forEach(h=>{
        const col=f3col_(current,h),v=values[h];
        if (!f3eq_(existing.values[col],v)) edits.push({row:existing.row,col:col,value:v});
      });
      old.delete(key);
    } else {
      const row=new Array(current.width).fill('');
      headers.forEach(h=>row[f3col_(current,h)]=values[h]);
      adds.push(row);
    }
  }
  for (const existing of old.values()) {
    ['SL đơn','SL SP'].forEach(h=>{
      const col=f3col_(current,h);
      if (!f3eq_(existing.values[col],0)) edits.push({row:existing.row,col:col,value:0});
    });
  }
  f3applyCells_(current.sh,edits);
  adds.forEach(row=>f3appendOwned_(current,'FFPL_V3_BC',row));
  // Không ghi APL hoặc ghi chú; tổng đơn ngày lấy từ Set OD, không cộng SL đơn các SKU.
}

/* DS đơn all (dán file đơn) -> dựng BC_Xuất FF -> tính XUẤT FF. Gom đơn chạy ngoài khóa;
 * khóa ngắn khi ghi BC, rồi luồng XUẤT FF tự đọc BC và ghi Tổng hợp như các luồng khác. */
function refreshBCXuatFF() {
  return ffplGuard_("refreshBCXuatFF",()=>{
    const catalog=f3catalog_();
    f3summary_(catalog);
    const grouped=f3groupOrders_(f3table_(FFPL3.sheets.orders),catalog).groups;
    f3lock_(()=>f3buildBC_(catalog,grouped));
    return f3run_('ff');
  });
}

function refreshKiemKe(editedRange) {
  return ffplGuard_("refreshKiemKe",()=>{
  return f3lock_(() => {
    const catalog = f3catalog_();

    const table = f3fillNames_(
      FFPL3.sheets.check,
      catalog,
      editedRange
    );

    const summary = f3summary_(catalog);

    const skuColumn = f3col_(table, 'SUB SKU');
    const qtyColumn = f3col_(table, 'Số lượng');
    const compareColumn = f3col_(table, 'Đối chiếu');

    const stockColumn = f3col_(
      summary.t,
      'TỒN'
    );

    // Quét cả tab (sửa khi đang lọc vẫn đúng), chỉ ghi ô Đối chiếu khác giá trị hiện có.
    const edits = [];

    table.rows.forEach((row, index) => {
      const sheetRow = table.hr + 1 + index;

      const sku = f3text_(row[skuColumn]);

      if (!sku) return;

      const existing = summary.map.get(sku);

      let comparison;

      if (row[qtyColumn] === '') {
        comparison = 'Chưa nhập số kiểm';
      } else if (!existing) {
        comparison = 'Chưa có tồn sổ';
      } else if (
        f3num_(row[qtyColumn]) ===
        f3num_(existing.values[stockColumn])
      ) {
        comparison = 'Khớp';
      } else {
        comparison =
          'Lệch: tồn ' +
          existing.values[stockColumn] +
          ' / kiểm ' +
          row[qtyColumn];
      }

      if (!f3eq_(row[compareColumn], comparison)) {
        edits.push({row: sheetRow, col: compareColumn, value: comparison});
      }
    });

    f3applyCells_(table.sh, edits);
  });

  });
}

/************************************************
 * CẬP NHẬT DANH MỤC
 ************************************************/

function ffplRefreshDanhMuc() {
  return ffplGuard_("ffplRefreshDanhMuc",()=>{
  return f3lock_(() => {
    const catalog = f3catalog_();
    const summary = f3summary_(catalog);

    const nameColumn = f3col_(
      summary.t,
      'TÊN SKU'
    );

    for (const [sku, existing] of summary.map) {
      const product = catalog.bySku.get(sku);

      if (!product) {
        throw new Error(
          'SKU trên tổng không còn trong danh mục: ' +
          sku
        );
      }

      f3set_(
        summary.t.sh,
        existing.row,
        nameColumn,
        product.name
      );

      f3derived_(
        summary.t,
        existing.row,
        existing.values,
        product
      );
    }

    f3dropdowns_(catalog);
    f3headline_();
  });

  });
}

/************************************************
 * SỰ KIỆN SỬA SHEET
 ************************************************/

function handleSpreadsheetEdit(e) {
  return ffplGuard_("handleSpreadsheetEdit",()=>{
  if (!e || !e.range) return;

  const sheetName = e.range
    .getSheet()
    .getName();

  const headerRow = FFPL3.headers[sheetName];

  if (!headerRow) return;
  if (e.range.getRow() <= headerRow && e.range.getLastRow() >= headerRow) {
    // Đổi nhầm header phải báo ngay, không âm thầm bỏ qua.
    f3validateSheet_(sheetName);
    if (e.range.getLastRow() <= headerRow) return;
  }

  if (sheetName === FFPL3.sheets.in) {
    return f3run_('in', e.range);
  }

  if (sheetName === FFPL3.sheets.transfer) {
    return f3run_('transfer', e.range);
  }

  if (sheetName === FFPL3.sheets.returns) {
    // Hàng hoàn đổi: tính HOÀN ở Tổng hợp + đẩy sang SCUTT. Lỗi bên này không chặn bên kia.
    let error = null;
    try { f3run_('returns'); } catch (err) { error = err; }
    try { if (typeof syncHangHoan === 'function') syncHangHoan(); } catch (err) { error = error || err; }
    if (error) throw error;
    return;
  }

  if (sheetName === FFPL3.sheets.ff) {
    return f3run_('ff');
  }

  if (sheetName === FFPL3.sheets.orders) {
    return refreshBCXuatFF();
  }

  if (sheetName === FFPL3.sheets.check) {
    return refreshKiemKe(e.range);
  }

  if (sheetName === FFPL3.sheets.sku) {
    return ffplRefreshDanhMuc();
  }

  });
}

function handleSpreadsheetChange(e) {
  return ffplGuard_("handleSpreadsheetChange",()=>{
  // Sự kiện cấu trúc không chỉ rõ nguồn đã đổi.
  // Không đoán nguồn và không gọi refreshAll().
  // Các trigger riêng sẽ nhận biết sửa/xóa dữ liệu.
  console.log(
    'Thay đổi cấu trúc: ' +
    (e && e.changeType || 'manual') +
    '. Các luồng tự kiểm tra riêng.'
  );

  });
}

/************************************************
 * BỎ KIỂM TRA THEO GIỜ
 * Không còn trigger 5 phút. Giữ tên hàm rỗng để trigger cũ (nếu chưa cài lại) chạy không lỗi.
 * Chạy ffplInstallTriggers() một lần sau khi dán code để xóa hẳn các trigger này.
 ************************************************/

function f3poll_(sheetName, handler) { return {ok: true, skipped: 'poll đã bỏ'}; }
function ffplPollNhap() { return f3poll_(); }
function ffplPollChuyenDongTay() { return f3poll_(); }
function ffplPollHoan() { return f3poll_(); }
function ffplPollDon() { return f3poll_(); }
function ffplPollBC() { return f3poll_(); }
function ffplPollSku() { return f3poll_(); }

/************************************************
 * APP NHẬP HOÀN (API) -> TÍNH HOÀN + SCUTT
 * App ghi Hàng hoàn bằng code nên trigger sửa không chạy. API ghi xong thì hẹn
 * MỘT lượt chạy một lần (không phải trigger định kỳ) để trả kết quả cho app ngay.
 * Nhiều đơn liền nhau chỉ giữ 1 lượt hẹn; lượt đó tính theo toàn bộ tab nên gồm cả các đơn sau.
 ************************************************/

const FFPL_RETURNS_JOB = 'ffplReturnsJob';

function ffplScheduleReturnsJob() {
  const pending = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === FFPL_RETURNS_JOB);
  if (pending) return {ok: true, scheduled: false};
  ScriptApp.newTrigger(FFPL_RETURNS_JOB).timeBased().after(5000).create();
  return {ok: true, scheduled: true};
}

function ffplReturnsJob() {
  // Xóa lượt hẹn TRƯỚC khi đọc nguồn: đơn nào lưu sau thời điểm này sẽ hẹn được lượt mới.
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === FFPL_RETURNS_JOB)
    .forEach(t => ScriptApp.deleteTrigger(t));
  let error = null;
  try { f3run_('returns'); } catch (err) { error = err; }
  try { if (typeof syncHangHoan === 'function') syncHangHoan(); } catch (err) { error = error || err; }
  if (error) throw error;
}

/************************************************
 * ĐỐI SOÁT TẤT CẢ — CHỈ CHẠY KHI CHỦ ĐỘNG YÊU CẦU
 ************************************************/

function refreshAll() {
  return ffplGuard_("refreshAll",()=>{
  const errors = [];

  const jobs = [
    ['Nhập', ffplRefreshNhap],
    ['Xuất', refreshBCXuatFF],
    ['Đóng tay/chuyển kho', ffplRefreshChuyenDongTay],
    ['Hoàn', ffplRefreshHoan]
  ];

  for (const [label, handler] of jobs) {
    try {
      handler();
    } catch (error) {
      console.error(label + ': ' + error.message);
      errors.push(label + ': ' + error.message);
    }
  }

  if (errors.length) {
    throw new Error(
      'Các luồng khác đã được thử độc lập. Lỗi: ' +
      errors.join(' | ')
    );
  }

  });
}

/************************************************
 * MENU
 ************************************************/

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('FFPL')

      .addItem(
        'Cập nhật nhập',
        'ffplRefreshNhap'
      )

      .addItem(
        'Cập nhật xuất từ DS đơn',
        'refreshBCXuatFF'
      )

      .addItem(
        'Cập nhật xuất từ BC',
        'ffplRefreshXuatTuBC'
      )

      .addItem(
        'Cập nhật đóng tay/chuyển kho',
        'ffplRefreshChuyenDongTay'
      )

      .addItem(
        'Cập nhật hoàn',
        'ffplRefreshHoan'
      )

      .addItem(
        'Đối soát tất cả — chủ động',
        'refreshAll'
      )

      .addToUi();
  } catch (error) {
    // Project standalone có thể không có UI context.
  }
}

/************************************************
 * DROPDOWN TÊN SẢN PHẨM
 ************************************************/

function f3dropdowns_(catalog) {
  const master = f3table_(FFPL3.sheets.sku);

  const nameColumn = f3col_(
    master,
    'Tên sản phẩm'
  );

  const source = master.sh.getRange(
    master.hr + 1,
    nameColumn + 1,
    master.sh.getMaxRows() - master.hr,
    1
  );

  const rule = SpreadsheetApp
    .newDataValidation()
    .requireValueInRange(source, true)
    .setAllowInvalid(false)
    .build();

  const targetSheets = [
    FFPL3.sheets.in,
    FFPL3.sheets.transfer,
    FFPL3.sheets.check
  ];

  for (const sheetName of targetSheets) {
    const table = f3table_(sheetName);

    const skuColumn = f3col_(table, 'SUB SKU');
    const productColumn = f3col_(
      table,
      'Tên sản phẩm'
    );

    const rows = table.sh.getMaxRows() - table.hr;

    table.sh
      .getRange(
        table.hr + 1,
        skuColumn + 1,
        rows,
        1
      )
      .clearDataValidations();

    table.sh
      .getRange(
        table.hr + 1,
        productColumn + 1,
        rows,
        1
      )
      .setDataValidation(rule);
  }
}

/************************************************
 * CÀI TRIGGER
 ************************************************/

function ffplInstallTriggers() {
  return ffplGuard_("ffplInstallTriggers",()=>{
  if (typeof ffplNotifyError!=='function') throw new Error('Cần thêm CanhBaoLoi.gs trước khi cài trigger.');
  if (typeof ffplAuthorizeAlerts==='function') ffplAuthorizeAlerts();
  f3lock_(() => {
    const catalog = f3catalog_();

    const summary=f3summary_(catalog);
    const owned=f3ss_().getRangeByName('FFPL_V3_ALL');
    ['Tổng NVL','Số dự kiến'].forEach(h=>{
      summary.t.sh.getRange(owned.getRow(),f3col_(summary.t,h)+1,owned.getNumRows(),1).setNumberFormat(h==='Tổng NVL' ? '0.######" m"' : '0.0');
    });
    f3state_();
    f3dropdowns_(catalog);
  });

  const ownedHandlers = [
    'handleSpreadsheetEdit',
    'handleSpreadsheetChange',

    // Trigger của bản vận hành trước.
    'ffplPoll',

    // Trigger của bản này.
    'ffplPollNhap',
    'ffplPollChuyenDongTay',
    'ffplPollHoan',
    'ffplPollDon',
    'ffplPollBC',
    'ffplPollSku',

    // Sync hoàn cũ chạy theo MỌI thay đổi của file và giữ khóa -> bỏ; nay gọi khi Hàng hoàn đổi.
    'onSourceChange',
    FFPL_RETURNS_JOB
  ];

  ScriptApp.getProjectTriggers()
    .forEach(trigger => {
      if (
        ownedHandlers.includes(
          trigger.getHandlerFunction()
        )
      ) {
        ScriptApp.deleteTrigger(trigger);
      }
    });

  const ss = f3ss_();

  ScriptApp
    .newTrigger('handleSpreadsheetEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  // Chỉ còn trigger SỬA: sửa tab nào chạy luồng tab đó. Không trigger thay đổi, không trigger theo giờ.
  // App nhập hoàn: API tự hẹn 1 lượt tính HOÀN + SCUTT (ffplReturnsJob).

  console.log(
    'Đã cài dropdown và trigger sửa. ' +
    'Đã xóa trigger 5 phút, trigger thay đổi và onSourceChange cũ. ' +
    'Mỗi tab chạy luồng riêng.'
  );

  });
}

function installTriggers() {
  return ffplInstallTriggers();
}

/** Bổ sung V4: cập nhật chọn lọc, APL, NVL, cảnh báo. */
function f3blank_(v) { return v===null || v===undefined || (typeof v==='string' && v.trim()===''); }
function f3effective_(base,actual) { return f3blank_(actual) ? base : actual; }
function f3require_(table,headers) { headers.forEach(h=>f3col_(table,h)); }
function f3deriveValues_(table,values,p,materials) {
  const col=h=>f3col_(table,h), get=h=>f3num_(values[col(h)]);
  const stock=get('NHẬP')+get('HOÀN')-get('XUẤT FF')-get('ĐÓNG TAY')-get('CHUYỂN KHO');
  const result=[
    [col('TỒN'),stock],
    [col('CBM'),p.cbm===null?'':stock*p.cbm],
    [col('TÌNH TRẠNG TỒN KHO'),stock<=0?'Hết hàng':stock<=FFPL3.minStock?'Sắp hết':'OK']
  ];
  if (materials!==false) {
    const exported=get('XUẤT FF')+get('ĐÓNG TAY');
    const nvl=exported<=0 || p.norm===null ? '' : exported*p.norm/100; // cm → mét
    result.push([col('Tổng NVL'),nvl],[col('Số dự kiến'),nvl===''?'':nvl/FFPL3.bubbleRollMeters]);
  }
  return result;
}
function f3applyCells_(sheet,edits) {
  const byColumn=new Map();
  edits.forEach(e=>{
    if (!byColumn.has(e.col)) byColumn.set(e.col,new Map());
    byColumn.get(e.col).set(e.row,e);
  });
  const blocks=[];
  for (const [col,map] of byColumn) {
    const rows=[...map.keys()].sort((a,b)=>a-b);
    let start=0;
    while (start<rows.length) {
      let end=start+1;
      while (end<rows.length && rows[end]===rows[end-1]+1) end++;
      const range=sheet.getRange(rows[start],col+1,end-start,1);
      if (range.getFormulas().some(r=>r[0]))
        throw new Error('Có công thức trong vùng code quản lý: '+sheet.getName()+'!'+range.getA1Notation());
      blocks.push({range:range,values:rows.slice(start,end).map(r=>[map.get(r).value])});
      start=end;
    }
  }
  blocks.forEach(b=>b.range.setValues(b.values));
}
function f3validateSheet_(sheetName) {
  const required={
    'Tổng hợp all':['SUB SKU','TÊN SKU','NHẬP','XUẤT FF','ĐÓNG TAY','CHUYỂN KHO','HOÀN','TỒN','CBM','TÌNH TRẠNG TỒN KHO','Tổng NVL','Số dự kiến'],
    'Thông tin SKU':['SUB SKU','Tên sản phẩm','Đơn vị','Kích thước','Tổng mức tiêu hao'],
    'Nhập NCC':['SUB SKU','Tên sản phẩm','Ngày','Số lượng'],
    'BC_Xuất FF':['THÁNG','NGÀY','SUB SKU','TÊN SẢN PHẨM','SL đơn','SL SP','SL đơn APL','SL SP APL','GHI CHÚ'],
    'DS đơn all':['Mã OD','Mã sản phẩm','Tên sản phẩm','Số lượng','Thời gian tạo'],
    'Đóng tay_Chuyển kho':['SUB SKU','Tên sản phẩm','Ngày','Số lượng','Loại xuất','Trạng thái'],
    'Hàng hoàn':['Tên sản phẩm','Ngày tháng','Số lượng','Mã đơn','Loại hoàn'],
    'Kiểm kê':['SUB SKU','Tên sản phẩm','Ngày kiểm','Số lượng','Đối chiếu']
  };
  if (required[sheetName]) f3require_(f3table_(sheetName),required[sheetName]);
}
function f3refreshCheckForSkus_(catalog,affected) {
  const table=f3table_(FFPL3.sheets.check);
  f3require_(table,['SUB SKU','Số lượng','Đối chiếu']);
  const summary=f3summary_(catalog);
  const skuCol=f3col_(table,'SUB SKU'), qtyCol=f3col_(table,'Số lượng');
  const compare=f3col_(table,'Đối chiếu'), stockCol=f3col_(summary.t,'TỒN');
  const edits=[];
  table.rows.forEach((r,i)=>{
    const sku=f3text_(r[skuCol]);
    if (!affected.has(sku)) return;
    const existing=summary.map.get(sku);
    const v=f3blank_(r[qtyCol])?'Chưa nhập số kiểm':!existing?'Chưa có tồn sổ':
      f3num_(r[qtyCol])===f3num_(existing.values[stockCol])?'Khớp':
      'Lệch: tồn '+existing.values[stockCol]+' / kiểm '+r[qtyCol];
    if (!f3eq_(r[compare],v)) edits.push({row:table.hr+1+i,col:compare,value:v});
  });
  f3applyCells_(table.sh,edits);
}
function f3runUnlockedFF_(catalog) {
  const source=f3readSource_('ff',catalog),state=f3state_(),next=f3aggregate_(source.entries);
  const previous=state.rows.filter(r=>f3text_(r[0])==='ff').map(r=>f3text_(r[1]));
  const affected=f3writeLane_(catalog,'ff',f3totals_(next,'ff'),previous);
  f3commitState_(state,['ff'],next);
  f3headline_();
  if (affected.size) f3refreshCheckForSkus_(catalog,affected);
}
function ffplGuard_(label,callback) {
  try { return callback(); }
  catch (error) {
    console.error(label+': '+(error.stack || error));
    if (!error.ffplReported) {
      error.ffplReported=true;
      if (typeof ffplNotifyError==='function') {
        try { ffplNotifyError(label,error); } catch (mailError) { console.error('Không gửi được cảnh báo: '+mailError); }
      } else console.error('Thiếu CanhBaoLoi.gs: chưa thể gửi cảnh báo email.');
    }
    throw error;
  }
}

function f3groupOrders_(source,catalog) {
  f3require_(source,['Mã OD','Mã sản phẩm','Tên sản phẩm','Số lượng','Thời gian tạo']);
  const col=h=>f3col_(source,h);
  const od=col('Mã OD'),sku=col('Mã sản phẩm'),name=col('Tên sản phẩm'),qty=col('Số lượng'),date=col('Thời gian tạo');
  const groups=new Map(),daily=new Map();
  source.rows.forEach((r,i)=>{
    if ([od,sku,name,qty,date].every(c=>f3blank_(r[c]))) return;
    if (f3blank_(r[od]) || f3blank_(r[sku]) || f3blank_(r[date]) || f3blank_(r[qty]))
      throw new Error('DS đơn all thiếu Mã OD/SKU/ngày/số lượng dòng '+(source.hr+1+i));
    const p=f3product_(catalog,r[name],r[sku]),day=f3day_(r[date]),order=f3text_(r[od]);
    const quantity=f3num_(r[qty]);
    if (quantity<0) throw new Error('DS đơn all: số lượng âm dòng '+(source.hr+1+i));
    const key=JSON.stringify([day,p.sku]),item=groups.get(key)||{p:p,day:day,qty:0,ods:new Set()};
    item.qty+=quantity; item.ods.add(order); groups.set(key,item);
    if (!daily.has(day)) daily.set(day,new Set());
    daily.get(day).add(order);
  });
  return {groups:groups,daily:daily};
}
/** Trả số đơn duy nhất toàn ngày; không cộng số đơn của các SKU. Không ghi thêm bảng. */
function ffplGetDailyOrderCounts() {
  return ffplGuard_("ffplGetDailyOrderCounts",()=>{
  const data=f3groupOrders_(f3table_(FFPL3.sheets.orders),f3catalog_()),result={};
  for (const [day,ods] of data.daily) result[day]=ods.size;
  return result;

  });
}

/** Nhận cả trạng thái cũ (Date/serial) và mới (m:YYYY-MM). */
function f3stateMonth_(value) {
  let text;
  if (value instanceof Date && !isNaN(value)) text=f3month_(value);
  else if (typeof value==='number' && value>20000 && value<100000) {
    text=new Date(Date.UTC(1899,11,30)+Math.floor(value)*86400000).toISOString().slice(0,7);
  } else {
    text=f3text_(value).replace(/^m:/,'');
    if (!/^\d{4}-\d{2}$/.test(text)) text=f3month_(value);
  }
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(text)) throw new Error('Tháng trạng thái không hợp lệ: '+String(value));
  return text;
}


/** Cập nhật riêng vật liệu: chỉ đọc danh mục + Tổng hợp all, không chạy lại nguồn nhập/xuất/hoàn. */
function ffplRefreshNVL() {
  return ffplGuard_('ffplRefreshNVL',()=>f3lock_(()=>{
    const catalog=f3catalog_(),summary=f3summary_(catalog),table=summary.t;
    const nvlCol=f3col_(table,'Tổng NVL'),rollCol=f3col_(table,'Số dự kiến'),edits=[];
    for (const [sku,item] of summary.map) {
      const p=catalog.bySku.get(sku);
      if (!p) throw new Error('SKU không còn trong danh mục: '+sku);
      f3deriveValues_(table,item.values,p,true)
        .filter(([col])=>col===nvlCol || col===rollCol)
        .forEach(([col,value])=>{
          if (!f3eq_(item.values[col],value)) edits.push({row:item.row,col:col,value:value});
        });
    }
    f3applyCells_(table.sh,edits);
    const range=f3ss_().getRangeByName('FFPL_V3_ALL');
    table.sh.getRange(range.getRow(),nvlCol+1,range.getNumRows(),1).setNumberFormat('0.######" m"');
    table.sh.getRange(range.getRow(),rollCol+1,range.getNumRows(),1).setNumberFormat('0.0');
  }));
}


/** Chỉ cập nhật dòng tổng phía trên, không đọc/tính lại các nguồn giao dịch. */
function ffplRefreshHeadline() {
  return ffplGuard_('ffplRefreshHeadline',()=>f3lock_(()=>f3headline_()));
}


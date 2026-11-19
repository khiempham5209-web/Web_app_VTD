/************************************************
 * SYNC HÀNG HOÀN -> SCUTT HÈ 2026
 * Loại trừ Loại hoàn = Ecom; ánh xạ 8 cột giữ nguyên.
 * Dòng đã có ở đích được giữ lại nếu nguồn đổi sang Ecom.
 * Cài trigger bằng installSyncHangHoanTriggers().
 ************************************************/


const SOURCE_ID = '1TVck35ibO3zJ6lkKpHNyXckgLyjvUxEoLbMa-6kt-Pk';
const SOURCE_SHEET_NAME = 'Hàng hoàn';
const SOURCE_HEADER_ROW = 3;


const DEST_ID = '1WuJGFQG9QK6oSaBdjooLpYzTvT_MJRtYzCZjG3fD74M';
const DEST_SHEET_NAME = 'HN_ SCUTT hè 2026';
const DEST_HEADER_ROW = 1;


const PROP_SOURCE_HASH = 'LAST_SOURCE_HASH_HANG_HOAN';


const SRC_HEADERS = {
  NGAY_THANG: 'Ngày tháng',
  THOI_GIAN: 'Thời gian',
  MA_DON: 'Mã đơn',
  TEN_SP: 'Tên sản phẩm',
  SO_LUONG: 'Số lượng',
  TINH_TRANG: 'Tình trạng',
  DA_TAO_SV: 'Đã tạo SV',
  GHI_CHU: 'Ghi chú',
  LOAI_HOAN: 'Loại hoàn'
};


const DST_HEADERS = {
  NGAY_HOAN_TRA: 'Ngày hoàn trả',
  MA_DON_GHTK: 'Mã đơn GHTK',
  TEN_SP: 'Tên sản phẩm',
  SO_LUONG: 'Số lượng',
  TINH_TRANG: 'Tình trạng',
  FF_NHAN_TON: 'FF nhận tồn',
  DA_TAO_LHNB: 'Đã tạo LHNB',
  NOTE: 'Note'
};


/************************************************
 * HÀM CHẠY CHÍNH
 ************************************************/


/* Chỉ chạy khi Hàng hoàn đổi (sửa tay -> handleSpreadsheetEdit; app nhập -> ffplReturnsJob).
 * Kiểm tra nguồn có đổi NGOÀI khóa; không đổi thì thoát ngay, không giữ khóa của các luồng khác.
 * Có đổi mới vào khóa ngắn để đọc đích + ghi (tránh 2 lượt cùng thêm trùng dòng). */
function syncHangHoan() {
  const srcSh = SpreadsheetApp.openById(SOURCE_ID).getSheetByName(SOURCE_SHEET_NAME);
  if (!srcSh) throw new Error('Không tìm thấy tab nguồn: ' + SOURCE_SHEET_NAME);
  const srcData = srcSh.getDataRange().getValues();
  if (srcData.length <= SOURCE_HEADER_ROW) return;
  const srcHash = 'EXCLUDE_ECOM_V1:' + makeHash_([srcData[SOURCE_HEADER_ROW - 1]].concat(srcData.slice(SOURCE_HEADER_ROW)));
  if (srcHash === PropertiesService.getScriptProperties().getProperty(PROP_SOURCE_HASH)) {
    console.log('Nguồn không thay đổi -> bỏ qua sync');
    return;
  }
  return syncHangHoanLocked_();
}


function syncHangHoanLocked_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(120000)) throw new Error('Chờ ghi SCUTT quá 2 phút, thử lại sau.');


  try {
    const srcSS = SpreadsheetApp.openById(SOURCE_ID);
    const dstSS = SpreadsheetApp.openById(DEST_ID);


    const srcSh = srcSS.getSheetByName(SOURCE_SHEET_NAME);
    const dstSh = dstSS.getSheetByName(DEST_SHEET_NAME);


    if (!srcSh) throw new Error('Không tìm thấy tab nguồn: ' + SOURCE_SHEET_NAME);
    if (!dstSh) throw new Error('Không tìm thấy tab đích: ' + DEST_SHEET_NAME);


    const srcData = srcSh.getDataRange().getValues();
    const dstData = dstSh.getDataRange().getValues();


    if (srcData.length <= SOURCE_HEADER_ROW) return;
    if (dstData.length < DEST_HEADER_ROW) throw new Error('Tab đích chưa có header');


    const srcHeader = srcData[SOURCE_HEADER_ROW - 1];
    const dstHeader = dstData[DEST_HEADER_ROW - 1];


    const srcRowsData = srcData.slice(SOURCE_HEADER_ROW);
    const dstRowsData = dstData.slice(DEST_HEADER_ROW);


    const srcHash = 'EXCLUDE_ECOM_V1:' + makeHash_([srcHeader].concat(srcRowsData));
    const props = PropertiesService.getScriptProperties();
    const lastHash = props.getProperty(PROP_SOURCE_HASH);


    const srcHeaderMap = buildHeaderMap_(srcHeader);
    const dstHeaderMap = buildHeaderMap_(dstHeader);


    validateHeaders_(srcHeaderMap, Object.values(SRC_HEADERS), 'nguồn');
    validateHeaders_(dstHeaderMap, Object.values(DST_HEADERS), 'đích');


    if (srcHash === lastHash) {
      console.log('Nguồn không thay đổi -> bỏ qua sync');
      return;
    }


    const srcRows = buildSourceRows_([srcHeader].concat(srcRowsData), srcHeaderMap);
    const dstMap = buildDestMap_([dstHeader].concat(dstRowsData), dstHeaderMap, DEST_HEADER_ROW);


    let added = 0;
    let updated = 0;
    let skipped = 0;


    const appendRows = [];


    srcRows.forEach(srcObj => {
      const maDon = srcObj.maDon;
      if (!maDon) return;


      const mapped = mapSourceToDest_(srcObj);


      if (dstMap.has(maDon)) {
        const item = dstMap.get(maDon);
        const rowIndex = item.rowIndex;
        const oldRow = item.row;


        if (isDestRowSame_(oldRow, mapped, dstHeaderMap)) {
          skipped++;
          return;
        }


        updateSingleRow_(dstSh, rowIndex, mapped, dstHeaderMap);
        updated++;
      } else {
        appendRows.push(mapped);
        added++;
      }
    });


    if (appendRows.length > 0) {
      appendNewRows_(dstSh, appendRows, dstHeaderMap);
    }


    applyCheckbox_(dstSh, dstHeaderMap[DST_HEADERS.DA_TAO_LHNB] + 1);


    props.setProperty(PROP_SOURCE_HASH, srcHash);


    console.log(`Sync xong | Thêm mới: ${added} | Cập nhật: ${updated} | Bỏ qua: ${skipped}`);


  } finally {
    lock.releaseLock();
  }
}


/************************************************
 * BUILD DATA
 ************************************************/


function buildSourceRows_(data, h) {
  const rows = [];


  for (let i = 1; i < data.length; i++) {
    const r = data[i];


    // Loại trống hoặc Hoàn trả vẫn sync như cũ; chỉ bỏ Ecom.
    if (clean_(r[h[SRC_HEADERS.LOAI_HOAN]]).toLowerCase() === 'ecom') continue;
    const maDon = clean_(r[h[SRC_HEADERS.MA_DON]]);
    if (!maDon) continue;


    rows.push({
      ngayThang: r[h[SRC_HEADERS.NGAY_THANG]],
      thoiGian: r[h[SRC_HEADERS.THOI_GIAN]],
      maDon: maDon,
      tenSP: r[h[SRC_HEADERS.TEN_SP]],
      soLuong: r[h[SRC_HEADERS.SO_LUONG]],
      tinhTrang: r[h[SRC_HEADERS.TINH_TRANG]],
      daTaoSV: r[h[SRC_HEADERS.DA_TAO_SV]],
      ghiChu: r[h[SRC_HEADERS.GHI_CHU]]
    });
  }


  return rows;
}


function buildDestMap_(data, h, headerRow) {
  const map = new Map();


  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const maDon = clean_(r[h[DST_HEADERS.MA_DON_GHTK]]);


    if (maDon && !map.has(maDon)) {
      map.set(maDon, {
        rowIndex: headerRow + i,
        row: r
      });
    }
  }


  return map;
}


function mapSourceToDest_(src) {
  return {
    [DST_HEADERS.NGAY_HOAN_TRA]: src.ngayThang,
    [DST_HEADERS.MA_DON_GHTK]: src.maDon,
    [DST_HEADERS.TEN_SP]: src.tenSP,
    [DST_HEADERS.SO_LUONG]: src.soLuong,
    [DST_HEADERS.TINH_TRANG]: src.tinhTrang,
    [DST_HEADERS.FF_NHAN_TON]: mapFFNhanTon_(src.tinhTrang),
    [DST_HEADERS.DA_TAO_LHNB]: toCheckbox_(src.daTaoSV),
    [DST_HEADERS.NOTE]: src.ghiChu
  };
}


/************************************************
 * WRITE DATA
 ************************************************/


function updateSingleRow_(sheet, rowIndex, mapped, h) {
  Object.keys(mapped).forEach(headerName => {
    const colIndex = h[headerName] + 1;
    sheet.getRange(rowIndex, colIndex).setValue(mapped[headerName]);
  });
}


function appendNewRows_(sheet, rows, h) {
  const startRow = sheet.getLastRow() + 1;
  const maxCol = sheet.getLastColumn();


  const values = rows.map(mapped => {
    const row = new Array(maxCol).fill('');


    Object.keys(mapped).forEach(headerName => {
      row[h[headerName]] = mapped[headerName];
    });


    return row;
  });


  sheet.getRange(startRow, 1, values.length, maxCol).setValues(values);
}


/************************************************
 * COMPARE
 ************************************************/


function isDestRowSame_(oldRow, mapped, h) {
  for (const headerName of Object.keys(mapped)) {
    const oldVal = normalizeCompare_(oldRow[h[headerName]]);
    const newVal = normalizeCompare_(mapped[headerName]);


    if (oldVal !== newVal) return false;
  }


  return true;
}


function normalizeCompare_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd');
  }


  if (v === true) return 'TRUE';
  if (v === false) return 'FALSE';


  return String(v ?? '').trim();
}


/************************************************
 * RULE
 ************************************************/


function mapFFNhanTon_(tinhTrang) {
  const v = removeVietnamese_(clean_(tinhTrang)).toLowerCase();


  if (v === 'binh thuong') {
    return 'Đã nhập';
  }


  if (
    v.includes('mop') ||
    v.includes('loi hong') ||
    v.includes('hong san pham')
  ) {
    return 'Chưa nhập';
  }


  return '';
}


function toCheckbox_(value) {
  if (value === true) return true;
  if (value === false) return false;


  const v = removeVietnamese_(clean_(value)).toLowerCase();


  if (['true', 'da tao', 'x', 'yes', '1'].includes(v)) return true;
  if (['false', 'chua tao', 'no', '0', ''].includes(v)) return false;


  return false;
}


/************************************************
 * TRIGGER
 ************************************************/


function installSyncHangHoanTriggers() {
  // Không còn trigger onChange riêng (chạy theo MỌI thay đổi của file, giữ khóa chặn luồng khác).
  // Sync SCUTT nay được gọi khi Hàng hoàn đổi: sửa tay (handleSpreadsheetEdit) / app nhập (ffplReturnsJob).
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onSourceChange') ScriptApp.deleteTrigger(trigger);
  });
  console.log('Đã xóa trigger onSourceChange. Cài trigger chung bằng ffplInstallTriggers().');
}


function onSourceChange(e) {
  console.log('onChange chạy: ' + JSON.stringify(e));
  syncHangHoan();
}


/************************************************
 * MANUAL
 ************************************************/


function runManual() {
  syncHangHoan();
}


function resetHash() {
  PropertiesService.getScriptProperties().deleteProperty(PROP_SOURCE_HASH);
}


/************************************************
 * HELPERS
 ************************************************/


function buildHeaderMap_(headerRow) {
  const map = {};


  headerRow.forEach((h, i) => {
    const name = clean_(h);
    if (name) map[name] = i;
  });


  return map;
}


function validateHeaders_(map, required, side) {
  const missing = required.filter(h => map[h] === undefined);


  if (missing.length) {
    throw new Error(`Thiếu header ${side}: ${missing.join(', ')}`);
  }
}


function makeHash_(data) {
  const text = JSON.stringify(data.map(row => row.map(v => {
    if (v instanceof Date) {
      return Utilities.formatDate(v, 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd HH:mm:ss');
    }
    return String(v ?? '').trim();
  })));


  const raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    text,
    Utilities.Charset.UTF_8
  );


  return raw.map(b => {
    const n = (b + 256) % 256;
    return n.toString(16).padStart(2, '0');
  }).join('');
}


function applyCheckbox_(sheet, colIndex) {
  const lastRow = sheet.getLastRow();
  if (lastRow < DEST_HEADER_ROW + 1) return;


  const rule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();


  sheet.getRange(DEST_HEADER_ROW + 1, colIndex, lastRow - DEST_HEADER_ROW, 1).setDataValidation(rule);
}


function clean_(v) {
  return String(v ?? '').trim();
}


function removeVietnamese_(str) {
  return String(str ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D');
}


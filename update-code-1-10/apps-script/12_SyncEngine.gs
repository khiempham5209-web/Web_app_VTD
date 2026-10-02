/* PN SYNC ENGINE V2 — đồng bộ tab chính ⇄ tab full (Booking, File đơn, Chứng từ_FF).
 *
 * Quy tắc nghiệp vụ:
 * - Tab full giữ lịch sử mọi tháng. Tab chính Booking/File đơn giữ tháng hiện tại; Chứng từ_FF giữ
 *   tháng hiện tại + đơn tháng cũ chưa nhận chứng từ. Mỗi dòng ở tab chính khớp 100% với dòng cùng ID ở full.
 * - Sửa ở tab chính -> full cập nhật y hệt. Sửa ở full (dòng còn ở tab chính) -> tab chính cập nhật y hệt.
 * - Dòng mới ở tab chính -> thêm vào full. Dòng mới ở full thuộc tháng hiện tại (hoặc chứng từ chưa nhận) -> thêm vào tab chính.
 * - Mỗi ngày: xóa khỏi tab chính dòng tháng cũ đã khớp full (chứng từ: chỉ khi Đã nhận chứng từ / Shop hủy OD).
 *
 * Kỹ thuật:
 * - Cột tìm theo TÊN tiêu đề. Tab chính: cột nghiệp vụ = dãy tiêu đề liền nhau từ cột A (bảng pivot/ghi chú
 *   đặt sau một cột trống không bị đụng tới). Tab full: mọi cột có tiêu đề, trừ cột kỹ thuật.
 *   Cột kỹ thuật tìm theo tên: __PN_ID, __PN_BASE (cả 2 tab), Tháng (tab full).
 * - __PN_BASE = dấu vân tay nội dung tại lần đồng bộ gần nhất, để biết bên nào vừa đổi.
 * - Mỗi lần chạy đọc mỗi tab 1 lần, so sánh trong bộ nhớ, chỉ GHI các dòng khác nhau. Không đổi gì thì không ghi gì.
 * - Hai bên cùng đổi một dòng: bên vừa sửa tay thắng (onEdit); nếu không biết thì tab chính thắng.
 *   Bản bị thay luôn được lưu vào tab ẩn _PN_CONFLICTS, không mất dữ liệu.
 */
const PN_V2 = {
  engine: 'v2',
  tech: ['__PN_ID', '__PN_BASE', 'Tháng'],
  conflictSheet: '_PN_CONFLICTS',
  budgetMs: 270000,
  pairs: {
    'Booking': {date: ['Ngày'], key: ['Row Labels'], ident: ['Row Labels']},
    'File đơn': {date: ['Thời gian tạo đơn'], key: ['Mã đơn'], ident: ['Mã đơn'], repeated: true},
    'Chứng từ_FF': {date: ['Ngày lên đơn'], key: ['Số đơn hàng'], ghtk: ['Mã đơn GHTK'], ident: ['Số đơn hàng', 'Mã đơn GHTK'],
      status: ['Xác thực hóa đơn'], docs: true}
  }
};
let pnV2Start_ = 0;

function pnV2Cfg_(p) { return Object.assign({}, PN_V2.pairs[p.main], {main: p.main, full: p.full}); }
function pnV2IsTech_(h) { return PN_V2.tech.indexOf(pnText_(h)) >= 0; }
function pnV2NewId_(p) { return 'v2:' + p.main + ':' + Utilities.getUuid(); }
function pnV2OverBudget_() { return pnV2Start_ && Date.now() - pnV2Start_ > PN_V2.budgetMs; }

// Bố cục một tab theo tiêu đề dòng 1.
function pnV2Layout_(sh, isFull) {
  const lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(pnText_);
  let width = 0;
  while (width < headers.length && headers[width] && !pnV2IsTech_(headers[width])) width++;
  const cols = {};
  headers.forEach((h, i) => {
    if (!h || pnV2IsTech_(h)) return;
    if (!isFull && i >= width) return;
    const n = pnNorm_(h);
    if (cols[n] != null) throw new Error('Trùng tên cột "' + h + '" ở tab ' + sh.getName() + '; đổi tên một cột để đồng bộ đúng.');
    cols[n] = i;
  });
  const tech = name => { const i = headers.indexOf(name); return i >= 0 ? i : null; };
  return {sh, name: sh.getName(), isFull, headers, width, cols, lastCol,
    id: tech('__PN_ID'), base: tech('__PN_BASE'), month: isFull ? tech('Tháng') : null};
}
function pnV2Col_(L, aliases) {
  for (const a of aliases || []) { const i = L.cols[pnNorm_(a)]; if (i != null) return i; }
  return null;
}
function pnV2ReadWidth_(L) {
  return Math.max(L.lastCol, (L.id == null ? 0 : L.id + 1), (L.base == null ? 0 : L.base + 1), (L.month == null ? 0 : L.month + 1));
}
function pnV2Read_(L) {
  const sh = L.sh, last = sh.getLastRow();
  return last < 2 ? [] : sh.getRange(2, 1, last - 1, pnV2ReadWidth_(L)).getValues();
}
// Giá trị chuẩn hóa để so sánh: ngày dạng text và dạng ngày của Sheets được coi là một.
function pnV2Canon_(v) {
  if (v instanceof Date) {
    if (isNaN(v)) return '';
    const s = Utilities.formatDate(v, PN_FULL.timezone, 'yyyy-MM-dd HH:mm:ss');
    return s.slice(11) === '00:00:00' ? s.slice(0, 10) : s;
  }
  if (v === true || v === false) return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return String(v);
  const t = pnText_(v);
  if (/^(\d{1,2}\/\d{1,2}\/\d{4}|\d{4}-\d{1,2}-\d{1,2})$/.test(t)) return pnIsoDate_(t) || t;
  return t;
}
// Danh sách tên cột so sánh = cột nghiệp vụ của tab chính.
function pnV2Names_(M) { return M.headers.slice(0, M.width).map(pnNorm_); }
function pnV2Sig_(L, row, names) {
  const parts = [];
  names.forEach(n => {
    const i = L.cols[n];
    const c = i == null ? '' : pnV2Canon_(row[i]);
    if (c !== '') parts.push([n, c]);
  });
  return JSON.stringify(parts);
}
function pnV2Has_(L, row, names) { return pnV2Sig_(L, row, names) !== '[]'; }
// Một dòng là HỒ SƠ khi có mã định danh (Booking: Row Labels; File đơn: Mã đơn; Chứng từ: Số đơn hàng hoặc Mã đơn GHTK).
// Dòng trống hoặc chỉ có giá trị mặc định (dropdown, checkbox) không phải hồ sơ: không đồng bộ, không tạo dòng ở full.
function pnV2Rec_(L, cfg, row) {
  return (cfg.ident || []).some(a => { const i = pnV2Col_(L, [a]); return i != null && pnV2Canon_(row[i]) !== ''; });
}
function pnV2Month_(L, cfg, row) {
  const i = pnV2Col_(L, cfg.date);
  return i == null ? '' : pnMonth_(row[i]);
}
function pnV2Key_(L, cfg, row) {
  if (cfg.repeated) return '';
  const i = pnV2Col_(L, cfg.key), k = i == null ? '' : pnText_(row[i]);
  if (k) return (cfg.docs ? 'ORDER:' : '') + k.toLowerCase();
  const g = cfg.ghtk ? pnV2Col_(L, cfg.ghtk) : null;
  return g != null && pnText_(row[g]) ? 'GHTK:' + pnText_(row[g]).toLowerCase() : '';
}
// Dòng tháng cũ đủ điều kiện rời tab chính.
function pnV2Eligible_(L, cfg, row, currentMonth) {
  const month = pnV2Month_(L, cfg, row);
  if (!month || month >= currentMonth) return false;
  if (!cfg.docs) return true;
  const s = pnV2Col_(L, cfg.status);
  return s != null && ['da nhan chung tu', 'shop huy od'].indexOf(pnNorm_(row[s])) >= 0;
}
// Dòng phải có mặt ở tab chính.
function pnV2NeedsMain_(L, cfg, row, currentMonth) {
  const month = pnV2Month_(L, cfg, row);
  if (!month) return !!cfg.docs; // ngày chưa xác định: chứng từ giữ lại để kiểm tra
  if (month >= currentMonth) return true;
  return !!cfg.docs && !pnV2Eligible_(L, cfg, row, currentMonth);
}
// Đảm bảo cột kỹ thuật và đủ cột nghiệp vụ của tab chính có trong full (thêm vào cột trống, không chèn cột).
function pnV2EnsureColumns_(p) {
  const main = pnSheet_(p.main), full = pnSheet_(p.full);
  let M = pnV2Layout_(main, false), F = pnV2Layout_(full, true);
  const freeCol = (L, limit) => {
    for (let i = L.isFull ? 0 : L.width + 1; i < limit; i++) if (!L.headers[i]) return i;
    return Math.max(L.headers.length, L.isFull ? 0 : L.width + 1);
  };
  const setHeader = (L, i, text) => { pnSize_(L.sh, 2, i + 1); L.sh.getRange(1, i + 1).setValue(text); L.headers[i] = text; };
  let changed = false;
  if (M.id == null || M.base == null) {
    // Cột kỹ thuật tab chính: mặc định AD:AE như bản migration.
    if (M.id == null) { const i = M.headers[29] ? freeCol(M, 60) : 29; setHeader(M, i, '__PN_ID'); }
    M = pnV2Layout_(main, false);
    if (M.base == null) { const i = M.headers[30] ? freeCol(M, 60) : 30; setHeader(M, i, '__PN_BASE'); }
    main.hideColumns(pnV2Layout_(main, false).id + 1, 1);
    changed = true;
  }
  if (F.id == null || F.base == null || F.month == null) {
    if (F.month == null) setHeader(F, freeCol(F, 29), 'Tháng');
    if (F.id == null) setHeader(F, F.headers[29] ? freeCol(F, 60) : 29, '__PN_ID');
    F = pnV2Layout_(full, true);
    if (F.base == null) setHeader(F, F.headers[30] ? freeCol(F, 60) : 30, '__PN_BASE');
    changed = true;
  }
  F = pnV2Layout_(full, true);
  const limit = Math.min(F.id == null ? 9999 : F.id, F.base == null ? 9999 : F.base);
  for (const h of M.headers.slice(0, M.width)) {
    if (F.cols[pnNorm_(h)] != null) continue;
    const i = freeCol(F, limit);
    if (i >= limit) throw new Error('Tab ' + p.full + ' hết cột trống để thêm cột "' + h + '".');
    setHeader(F, i, h);
    F = pnV2Layout_(full, true);
    changed = true;
  }
  if (changed) SpreadsheetApp.flush();
  return {M: pnV2Layout_(main, false), F: pnV2Layout_(full, true)};
}
function pnV2ConflictLog_(p, side, id, values) {
  const ss = pnSS_();
  let sh = ss.getSheetByName(PN_V2.conflictSheet);
  if (!sh) { sh = ss.insertSheet(PN_V2.conflictSheet); sh.getRange(1, 1, 1, 5).setValues([['Thời điểm', 'Tab', 'Bên bị thay', 'ID', 'Nội dung bị thay']]); sh.hideSheet(); }
  const row = Math.max(2, sh.getLastRow() + 1);
  pnSize_(sh, row, 5);
  sh.getRange(row, 1, 1, 5).setValues([[new Date(), p.main, side, id, values]]);
}
// Ghi theo từng ô cần đổi; gom dòng liền nhau có cùng tập cột thành một lệnh ghi.
function pnV2Flush_(sh, updates) {
  const rows = Array.from(updates.keys()).sort((a, b) => a - b);
  const items = rows.map(r => {
    const vals = updates.get(r), cols = Object.keys(vals).map(Number).sort((a, b) => a - b), runs = [];
    cols.forEach(c => { const last = runs[runs.length - 1]; if (last && last.end === c - 1) last.end = c; else runs.push({start: c, end: c}); });
    return {r, vals, runs, sig: JSON.stringify(runs)};
  });
  let writes = 0;
  for (let i = 0; i < items.length;) {
    let j = i + 1;
    while (j < items.length && items[j].r === items[j - 1].r + 1 && items[j].sig === items[i].sig) j++;
    const maxCol = items[i].runs.length ? items[i].runs[items[i].runs.length - 1].end + 1 : 1;
    pnSize_(sh, items[j - 1].r, maxCol);
    items[i].runs.forEach(run => {
      const block = items.slice(i, j).map(it => { const out = []; for (let c = run.start; c <= run.end; c++) out.push(it.vals[c]); return out; });
      sh.getRange(items[i].r, run.start + 1, j - i, run.end - run.start + 1).setValues(block);
      writes++;
    });
    i = j;
  }
  return writes;
}
function pnV2Put_(map, row, col, value) {
  if (col == null) return;
  if (!map.has(row)) map.set(row, {});
  map.get(row)[col] = value;
}
// Sao chép định dạng/validation của dòng mẫu (dòng 2) cho các dòng vừa thêm.
// Dòng mới thêm vào chỉ lấy dropdown/checkbox và định dạng số của dòng mẫu (dòng 2), KHÔNG lấy màu/chữ của đơn khác.
function pnV2CopyFormat_(L, start, count, width) {
  if (!count || start <= 2 || width < 1) return;
  const from = L.sh.getRange(2, 1, 1, width), to = L.sh.getRange(start, 1, count, width);
  from.copyTo(to, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  const numbers = from.getNumberFormats()[0];
  to.setNumberFormats(Array.from({length: count}, () => numbers.slice()));
}

function pnV2LastDataRow_(L, cfg, rows) {
  for (let i = rows.length - 1; i >= 0; i--) if (pnV2Rec_(L, cfg, rows[i]) || (L.id != null && pnText_(rows[i][L.id]))) return i + 2;
  return 1;
}

/* Đồng bộ một cặp tab. opts.prefer = {side:'main'|'full', rows:Set(số dòng trên tab đó)}:
 * dòng nằm trong vùng vừa sửa thì bên đó thắng khi hai bên cùng đổi. */
function pnV2SyncPair_(p, opts) {
  opts = opts || {};
  const cfg = pnV2Cfg_(p), month = pnCurrentMonth_();
  const layouts = pnV2EnsureColumns_(p), M = layouts.M, F = layouts.F, names = pnV2Names_(M);
  const legacy = PropertiesService.getScriptProperties().getProperty('PN_ENGINE') !== PN_V2.engine;
  const mRows = pnV2Read_(M), fRows = pnV2Read_(F);
  const mUp = new Map(), fUp = new Map();
  const res = {pair: p.main, pushed: 0, pulled: 0, toFull: 0, toMain: 0, conflicts: 0, ids: 0, mainRows: [], fullRows: []};
  const preferred = (side, row) => opts.prefer && opts.prefer.side === side && opts.prefer.rows && opts.prefer.rows.has(row);

  // Ghi giá trị nghiệp vụ của nguồn sang đích theo tên cột.
  const copy = (src, srcRow, dst, dstUp, dstRowNo, dstRow) => {
    names.forEach(n => {
      const s = src.cols[n], d = dst.cols[n];
      if (d == null) return;
      const v = s == null ? '' : srcRow[s];
      if (dstRow && pnV2Canon_(dstRow[d]) === pnV2Canon_(v)) return;
      pnV2Put_(dstUp, dstRowNo, d, v);
    });
  };
  const setMonth = (rowNo, row) => {
    if (F.month == null) return;
    const m = pnV2Month_(M, cfg, row) || 'CẦN KIỂM TRA';
    pnV2Put_(fUp, rowNo, F.month, m);
  };

  // 1. Chỉ mục full theo ID; dòng full chưa có ID hoặc trùng ID được cấp ID mới.
  const fById = new Map(), fByKey = new Map(), fNoId = [];
  fRows.forEach((r, i) => {
    if (!pnV2Rec_(F, cfg, r)) return;
    const rowNo = i + 2;
    let id = pnText_(r[F.id]);
    if (!id || fById.has(id)) { id = pnV2NewId_(p); fNoId.push(rowNo); pnV2Put_(fUp, rowNo, F.id, id); r[F.id] = id; res.ids++; }
    fById.set(id, {rowNo, r});
    const k = pnV2Key_(F, cfg, r);
    if (k && !fByKey.has(k)) fByKey.set(k, id);
  });

  // Lịch sử full không bao giờ được giảm. Nếu số hồ sơ full giảm (ai đó xóa dòng ở full): báo động,
  // vẫn đồng bộ tiếp (bước 2 tự chép lại vào full các dòng còn ở tab chính), không xóa gì thêm.
  const props = PropertiesService.getScriptProperties(), countKey = 'PN_FULL_COUNT_' + p.main;
  const prevCount = Number(props.getProperty(countKey) || 0);
  if (fById.size < prevCount) {
    const msg = new Date().toISOString() + ' ' + p.full + ' giảm từ ' + prevCount + ' xuống ' + fById.size + ' hồ sơ. Kiểm tra lịch sử phiên bản của Sheet.';
    props.setProperty('PN_FULL_LOSS', msg); console.error(msg);
    pnV2ConflictLog_(p, p.full, 'FULL_LOSS', msg);
    res.fullLoss = msg;
  }
  // 2. Duyệt tab chính.
  const mIds = new Set(), claimed = new Set();
  let fNext = pnV2LastDataRow_(F, cfg, fRows) + 1;
  const fAppendStart = fNext;
  mRows.forEach((r, i) => {
    if (!pnV2Rec_(M, cfg, r)) return;
    const rowNo = i + 2;
    let id = pnText_(r[M.id]);
    if (id && mIds.has(id)) id = ''; // dòng copy/dán kèm ID cũ: coi là hồ sơ mới
    if (!id && !cfg.repeated) {
      // Dòng chưa có ID nhưng đã có hồ sơ cùng mã ở full (chưa được dòng chính nào nhận): nối lại, không tạo trùng.
      const k = pnV2Key_(M, cfg, r), found = k && fByKey.get(k);
      if (found && !claimed.has(found) && !mRows.some(o => pnText_(o[M.id]) === found)) id = found;
    }
    const sigM = pnV2Sig_(M, r, names), hashM = pnHash_(sigM);
    if (!id || !fById.has(id)) {
      // Hồ sơ mới ở tab chính (hoặc full bị mất dòng): thêm vào full, không bao giờ để mất lịch sử.
      id = id || pnV2NewId_(p);
      if (!pnText_(r[M.id]) || pnText_(r[M.id]) !== id) { pnV2Put_(mUp, rowNo, M.id, id); res.ids++; }
      const target = fNext++;
      copy(M, r, F, fUp, target, null);
      pnV2Put_(fUp, target, F.id, id); pnV2Put_(fUp, target, F.base, hashM); setMonth(target, r);
      pnV2Put_(mUp, rowNo, M.base, hashM);
      mIds.add(id); claimed.add(id); res.toFull++; res.mainRows.push(rowNo); res.fullRows.push(target);
      return;
    }
    if (pnText_(r[M.id]) !== id) pnV2Put_(mUp, rowNo, M.id, id);
    mIds.add(id); claimed.add(id);
    const f = fById.get(id), sigF = pnV2Sig_(F, f.r, names), hashF = pnHash_(sigF);
    const baseM = pnText_(r[M.base]), baseF = pnText_(f.r[F.base]);
    if (sigM === sigF) {
      if (baseM !== hashM) pnV2Put_(mUp, rowNo, M.base, hashM);
      if (baseF !== hashM) pnV2Put_(fUp, f.rowNo, F.base, hashM);
      if (F.month != null && pnText_(f.r[F.month]) !== (pnV2Month_(F, cfg, f.r) || 'CẦN KIỂM TRA')) setMonth(f.rowNo, r);
      return;
    }
    // Hai bên khác nhau: xác định bên vừa đổi.
    // Bên vừa đổi theo dấu vân tay lần đồng bộ trước.
    let natural;
    if (legacy) {
      // Lần đầu chạy bản mới: __PN_BASE còn là vân tay kiểu cũ (theo vị trí cột).
      const oldBase = baseM || baseF;
      if (oldBase && pnFingerprint_(f.r, p) === oldBase) natural = 'push';
      else if (oldBase && pnFingerprint_(r, p) === oldBase) natural = 'pull';
    } else {
      const base = baseM || baseF;
      if (base === hashF) natural = 'push';
      else if (base === hashM) natural = 'pull';
    }
    // Vùng vừa sửa tay thắng; nếu không có thì theo bên vừa đổi; cả hai cùng đổi thì tab chính thắng.
    let dir = preferred('main', rowNo) ? 'push' : preferred('full', f.rowNo) ? 'pull' : natural || 'push';
    if (dir !== natural) {
      // Bên thua cũng có thay đổi riêng: lưu nguyên bản bị thay để không mất dữ liệu.
      res.conflicts++;
      pnV2ConflictLog_(p, dir === 'push' ? p.full : p.main, id, dir === 'push' ? sigF : sigM);
    }
    if (dir === 'push') {
      copy(M, r, F, fUp, f.rowNo, f.r);
      setMonth(f.rowNo, r);
      pnV2Put_(fUp, f.rowNo, F.base, hashM); pnV2Put_(mUp, rowNo, M.base, hashM);
      res.pushed++; res.mainRows.push(rowNo); res.fullRows.push(f.rowNo);
    } else {
      copy(F, f.r, M, mUp, rowNo, r);
      pnV2Put_(fUp, f.rowNo, F.base, hashF); pnV2Put_(mUp, rowNo, M.base, hashF);
      res.pulled++; res.mainRows.push(rowNo); res.fullRows.push(f.rowNo);
    }
  });

  // 3. Hồ sơ ở full chưa có ở tab chính nhưng phải có (tháng hiện tại, chứng từ chưa nhận): thêm vào tab chính.
  let mNext = pnV2LastDataRow_(M, cfg, mRows) + 1;
  const mAppendStart = mNext;
  fById.forEach((f, id) => {
    if (mIds.has(id) || !pnV2NeedsMain_(F, cfg, f.r, month)) return;
    const target = mNext++, sigF = pnV2Sig_(F, f.r, names), hashF = pnHash_(sigF);
    copy(F, f.r, M, mUp, target, null);
    pnV2Put_(mUp, target, M.id, id); pnV2Put_(mUp, target, M.base, hashF);
    if (pnText_(f.r[F.base]) !== hashF) pnV2Put_(fUp, f.rowNo, F.base, hashF);
    if (F.month != null && pnText_(f.r[F.month]) !== (pnV2Month_(F, cfg, f.r) || 'CẦN KIỂM TRA')) pnV2Put_(fUp, f.rowNo, F.month, pnV2Month_(F, cfg, f.r) || 'CẦN KIỂM TRA');
    mIds.add(id); res.toMain++; res.mainRows.push(target); res.fullRows.push(f.rowNo);
  });
  // Dòng full chỉ được cấp ID (không thuộc tab chính) cũng cần base và Tháng.
  fNoId.forEach(rowNo => {
    const f = fRows[rowNo - 2];
    if (!fUp.has(rowNo) || fUp.get(rowNo)[F.base] != null) return;
    pnV2Put_(fUp, rowNo, F.base, pnHash_(pnV2Sig_(F, f, names)));
    if (F.month != null) pnV2Put_(fUp, rowNo, F.month, pnV2Month_(F, cfg, f) || 'CẦN KIỂM TRA');
  });

  props.setProperty(countKey, String(fById.size + res.toFull));
  // 4. Ghi: chỉ các ô thay đổi. Dòng mới được chép định dạng của dòng 2.
  res.writes = pnV2Flush_(F.sh, fUp) + pnV2Flush_(M.sh, mUp);
  if (fNext > fAppendStart) pnV2CopyFormat_(F, fAppendStart, fNext - fAppendStart, F.width);
  if (mNext > mAppendStart) pnV2CopyFormat_(M, mAppendStart, mNext - mAppendStart, M.width);
  if (res.writes) { SpreadsheetApp.flush(); pnMarkDirty_(); }
  return res;
}

// Kiểm tra mọi dòng tab chính khớp đúng dòng cùng ID ở full. Dùng sau đồng bộ và trước khi dọn.
function pnV2Audit_(p) {
  const cfg = pnV2Cfg_(p), L = pnV2EnsureColumns_(p), M = L.M, F = L.F, names = pnV2Names_(M);
  const fById = new Map(), errors = [];
  pnV2Read_(F).forEach((r, i) => { const id = pnText_(r[F.id]); if (id) { if (fById.has(id)) errors.push('Full trùng ID dòng ' + (i + 2)); fById.set(id, r); } });
  const seen = new Set();
  pnV2Read_(M).forEach((r, i) => {
    if (!pnV2Rec_(M, cfg, r)) return;
    const id = pnText_(r[M.id]);
    if (!id) { errors.push('Chưa có ID dòng ' + (i + 2)); return; }
    if (seen.has(id)) { errors.push('Trùng ID dòng ' + (i + 2)); return; }
    seen.add(id);
    const f = fById.get(id);
    if (!f) errors.push('Thiếu ở full dòng ' + (i + 2));
    else if (pnV2Sig_(M, r, names) !== pnV2Sig_(F, f, names)) errors.push('Lệch với full dòng ' + (i + 2));
  });
  return {sheet: p.main, ok: !errors.length, errors: errors.slice(0, 50)};
}

// Cột tính bằng code (thay ARRAYFORMULA cũ): Booking!Check, File đơn!Quản lý/Khu vực/Ngay. Chỉ ghi ô khác.
function pnV2ComputeColumns_() {
  if (PropertiesService.getScriptProperties().getProperty('PN_FORMULAS_CONVERTED') !== '1') return 0;
  const factor = pnSheet_('TT Nhập').getRange('L3').getValue();
  const bc = pnRows_(pnSheet_('DS BC'), 15), byDepot = new Map();
  bc.forEach(r => { const k = pnLookupKey_(r[6]); if (!byDepot.has(k)) byDepot.set(k, r); });
  let writes = 0;
  PN_FULL.pairs.slice(0, 2).forEach(p => {
    [p.main, p.full].forEach(name => {
      const L = pnV2Layout_(pnSheet_(name), name === p.full), rows = pnV2Read_(L), up = new Map();
      if (p.main === 'Booking') {
        const d = pnV2Col_(L, ['Ngày']), total = pnV2Col_(L, ['Tổng bánh']), out = pnV2Col_(L, ['Check']);
        if (d == null || total == null || out == null) return;
        if (typeof factor !== 'number' || !Number.isFinite(factor)) throw new Error('TT Nhập!L3 không phải số hợp lệ.');
        rows.forEach((r, i) => {
          const v = r[d] === '' ? '' : pnRoundUp_(Number(r[total] || 0) * factor);
          if (pnV2Canon_(v) !== pnV2Canon_(r[out])) pnV2Put_(up, i + 2, out, v);
        });
      } else {
        const code = pnV2Col_(L, ['Mã đơn']), created = pnV2Col_(L, ['Thời gian tạo đơn']), depot = pnV2Col_(L, ['Kho đích']);
        const qm = pnV2Col_(L, ['Quản lý']), kv = pnV2Col_(L, ['Khu vực']), day = pnV2Col_(L, ['Ngay']);
        if (code == null || created == null || depot == null) return;
        rows.forEach((r, i) => {
          const empty = r[code] === '';
          const match = byDepot.get(pnLookupKey_(r[depot])), iso = pnIsoDate_(r[created]);
          const values = [[qm, empty ? '' : (match ? match[4] : '#N/A')], [kv, empty ? '' : (match ? match[14] : '#N/A')],
            [day, empty ? '' : (iso ? iso.slice(8) + '/' + iso.slice(5, 7) + '/' + iso.slice(0, 4) : '#VALUE!')]];
          values.forEach(([c, v]) => { if (c != null && pnV2Canon_(v) !== pnV2Canon_(r[c])) pnV2Put_(up, i + 2, c, v); });
        });
      }
      writes += pnV2Flush_(L.sh, up);
    });
  });
  return writes;
}

/* Một lượt đồng bộ đầy đủ theo đúng thứ tự phụ thuộc:
 * Booking, File đơn -> script nghiệp vụ ghi Chứng từ_full (đơn mới từ Booking, mã GHTK, khu vực) -> Chứng từ. */
function pnV2SyncAll_(opts) {
  opts = opts || {};
  pnV2Start_ = pnV2Start_ || Date.now();
  const out = {ok: true, pairs: []};
  pnV2ComputeColumns_();
  const booking = pnV2SyncPair_(PN_FULL.pairs[0], opts.pair === 'Booking' ? opts : {});
  const files = pnV2SyncPair_(PN_FULL.pairs[1], opts.pair === 'File đơn' ? opts : {});
  out.pairs.push(booking, files);
  if (pnV2OverBudget_()) { out.ok = false; out.partial = true; return out; }
  // Đơn mới / ngày từ Booking sang Chứng từ_full: chỉ các dòng Booking vừa đổi.
  const bookingRows = (opts.bookingRows || []).concat(booking.mainRows).filter(Boolean).sort((a, b) => a - b);
  pnV2Blocks_(bookingRows).forEach(([a, b]) => {
    syncBookingRowsToChungTuFF_(a, b);
    bookingSyncUpdateTargetDatesFromBookingRows_(a, b);
  });
  if (opts.fullSweep) {
    const last = pnSheet_('Booking').getLastRow();
    if (last > 1) { syncBookingRowsToChungTuFF_(2, last); bookingSyncUpdateTargetDatesFromBookingRows_(2, last); }
  }
  fillMissingGhtkCodesInChungTuFF();
  if (booking.mainRows.length || files.mainRows.length || opts.areaSweep) pnSyncAreaAll_();
  if (pnV2OverBudget_()) { out.ok = false; out.partial = true; return out; }
  out.pairs.push(pnV2SyncPair_(PN_FULL.pairs[2], opts.pair === 'Chứng từ_FF' ? opts : {}));
  if (out.pairs.every(r => r)) PropertiesService.getScriptProperties().setProperty('PN_ENGINE', PN_V2.engine);
  return out;
}
function pnV2Blocks_(rows) {
  const out = [];
  rows.forEach(r => { const last = out[out.length - 1]; if (last && r <= last[1] + 1) last[1] = Math.max(last[1], r); else out.push([r, r]); });
  return out;
}

/* Dọn mỗi ngày: xóa khỏi tab chính dòng tháng cũ đã khớp 100% với full (cùng ID, cùng nội dung).
 * Chứng từ_FF: chỉ dòng Đã nhận chứng từ / Shop hủy OD. Dòng chưa xác định ngày, chưa khớp: giữ lại.
 * Ghi lại phần còn lại liền nhau từ dòng 2 (không deleteRows để không lệch pivot/công thức). */
// Định dạng được dời theo dòng khi dọn (màu, chữ, số, dropdown/checkbox).
const PN_V2_FORMATS_ = [['getBackgrounds', 'setBackgrounds'], ['getFontColors', 'setFontColors'], ['getFontWeights', 'setFontWeights'],
  ['getFontLines', 'setFontLines'], ['getNumberFormats', 'setNumberFormats'], ['getDataValidations', 'setDataValidations']];
// Ô có nội dung thật (bỏ qua ô trống và checkbox chưa tích).
function pnV2Meaningful_(L, row) {
  for (let i = 0; i < L.width; i++) { const c = pnV2Canon_(row[i]); if (c !== '' && c !== 'FALSE') return true; }
  return false;
}
/* Dọn mỗi ngày: xóa khỏi tab chính dòng tháng cũ đã khớp 100% với full (cùng ID, cùng nội dung).
 * Chứng từ_FF: chỉ dòng Đã nhận chứng từ / Shop hủy OD. Dòng chưa xác định ngày, chưa khớp: giữ lại.
 * Dòng đang gõ dở (chưa có mã nhưng có nội dung) được giữ. Dòng mẫu trống (chỉ dropdown/checkbox chưa tích) bị xóa sạch.
 * Dồn phần còn lại lên từ dòng 2 KÈM định dạng của chính dòng đó; các dòng trống phía dưới được xóa sạch
 * giá trị, màu, định dạng, dropdown/checkbox. Chỉ trong vùng bảng (cột A tới cột cuối của bảng) + cột ID ẩn;
 * không deleteRows để không lệch pivot/công thức bên phải (File đơn T:W). */
function pnV2Cleanup_() {
  const month = pnCurrentMonth_(), results = [];
  for (const p of PN_FULL.pairs) {
    const cfg = pnV2Cfg_(p), L = pnV2EnsureColumns_(p), M = L.M, F = L.F, names = pnV2Names_(M);
    const fById = new Map();
    pnV2Read_(F).forEach(r => { const id = pnText_(r[F.id]); if (id) fById.set(id, r); });
    const before = pnV2Read_(M), cols = pnV2CleanupCols_(M);
    let removed = 0, blank = 0;
    const keptIdx = [];
    before.forEach((r, i) => {
      if (!pnV2Rec_(M, cfg, r)) { if (pnV2Meaningful_(M, r)) keptIdx.push(i); else if (pnV2Has_(M, r, names) || (M.id != null && pnText_(r[M.id]))) blank++; return; }
      const f = fById.get(pnText_(r[M.id]));
      const ok = pnV2Eligible_(M, cfg, r, month) && f && pnV2Sig_(M, r, names) === pnV2Sig_(F, f, names);
      if (ok) removed++; else keptIdx.push(i);
    });
    if (!removed) { results.push({sheet: p.main, removed: 0}); continue; }
    const pick = r => cols.map(c => r[c] === undefined ? '' : r[c]);
    const total = before.length, wanted = [];
    for (let i = 0; i < total; i++) wanted.push(i < keptIdx.length ? pick(before[keptIdx[i]]) : cols.map(() => ''));
    // Định dạng của vùng bảng, đọc một lần.
    const block = M.width ? M.sh.getRange(2, 1, total, M.width) : null;
    const formats = block ? PN_V2_FORMATS_.map(([get]) => block[get]()) : [];
    pnV2Backup_(p, M, before.map(pick), cols, wanted);
    PropertiesService.getScriptProperties().setProperty('PN_CLEANUP_PENDING', JSON.stringify({main: p.main, month, engine: PN_V2.engine}));
    pnV2WriteCols_(M, cols, wanted);
    if (block) {
      if (keptIdx.length) {
        const keptRange = M.sh.getRange(2, 1, keptIdx.length, M.width);
        PN_V2_FORMATS_.forEach(([, set], k) => keptRange[set](keptIdx.map(i => formats[k][i])));
      }
      const vacated = total - keptIdx.length;
      if (vacated > 0) {
        const start = keptIdx.length + 2, range = M.sh.getRange(start, 1, vacated, M.width);
        range.clearFormat();
        // Hết dòng giữ lại: giữ dropdown/checkbox ở dòng 2 làm mẫu cho dòng mới.
        if (keptIdx.length) range.clearDataValidations();
        else if (vacated > 1) M.sh.getRange(3, 1, vacated - 1, M.width).clearDataValidations();
      }
    }
    // Đọc lại đúng số dòng đã ghi (kể cả các dòng cuối vừa xóa trống) để đối chiếu.
    const check = M.sh.getRange(2, 1, total, pnV2ReadWidth_(M)).getValues().map(pick);
    if (pnHash_(check) !== pnHash_(wanted)) throw new Error('Đối soát sau dọn không khớp ở ' + p.main + '; giữ khóa phục hồi.');
    PropertiesService.getScriptProperties().deleteProperty('PN_CLEANUP_PENDING');
    results.push({sheet: p.main, removed, blankCleared: blank});
  }
  return {ok: true, results};
}

function pnV2CleanupCols_(M) {
  const cols = [];
  for (let i = 0; i < M.width; i++) cols.push(i);
  if (M.id != null) cols.push(M.id);
  if (M.base != null) cols.push(M.base);
  return cols;
}
function pnV2WriteCols_(M, cols, rows) {
  if (!rows.length) return;
  const runs = [];
  cols.forEach((c, k) => { const last = runs[runs.length - 1]; if (last && last.end === c - 1) { last.end = c; last.k2 = k; } else runs.push({start: c, end: c, k1: k, k2: k}); });
  runs.forEach(run => M.sh.getRange(2, run.start + 1, rows.length, run.end - run.start + 1).setValues(rows.map(r => r.slice(run.k1, run.k2 + 1))));
}
// Sao lưu trước khi dọn: cột A.. = dữ liệu trước dọn, các cột sau = dữ liệu dự kiến sau dọn.
function pnV2Backup_(p, M, before, cols, wanted) {
  const name = '_PN_BACKUP_' + p.main, ss = pnSS_(), sh = ss.getSheetByName(name) || ss.insertSheet(name);
  const w = cols.length, rows = before.map((r, i) => r.concat(wanted[i]));
  pnSize_(sh, rows.length + 1, w * 2);
  sh.getRange(1, 1, 1, w * 2).setValues([cols.map(c => M.headers[c] || '').concat(cols.map(c => 'SAU DỌN: ' + (M.headers[c] || '')))]);
  if (rows.length) sh.getRange(2, 1, rows.length, w * 2).setValues(rows);
  PropertiesService.getScriptProperties().setProperty('PN_BACKUP_' + p.main, JSON.stringify({rows: rows.length, cols, at: new Date().toISOString(), hash: pnHash_(rows), engine: PN_V2.engine}));
  sh.hideSheet();
}

// Lượt dọn bị ngắt giữa chừng: mỗi dòng phải đang là bản "trước" hoặc "sau dự kiến". Có dòng khác (đã sửa tay sau đó)
// hoặc có dòng mới phía dưới thì DỪNG, không ghi đè. Hợp lệ thì hoàn tất lượt dọn.
function pnV2ResumeCleanup_() {
  const props = PropertiesService.getScriptProperties(), raw = props.getProperty('PN_CLEANUP_PENDING');
  if (!raw) return {ok: true, skipped: true};
  const state = JSON.parse(raw), p = pnPair_(state.main), info = JSON.parse(props.getProperty('PN_BACKUP_' + p.main) || 'null');
  if (!info || !info.cols) throw new Error('Thiếu thông tin backup bản mới.');
  const w = info.cols.length;
  const backup = info.rows ? pnSheet_('_PN_BACKUP_' + p.main).getRange(2, 1, info.rows, w * 2).getValues() : [];
  if (pnHash_(backup) !== info.hash) throw new Error('Backup không khớp; không tự phục hồi.');
  const M = pnV2Layout_(pnSheet_(p.main), false), cfg = pnV2Cfg_(p);
  const width = pnV2ReadWidth_(M);
  const current = info.rows ? M.sh.getRange(2, 1, info.rows, width).getValues() : [];
  const pick = r => info.cols.map(c => r[c] === undefined ? '' : r[c]);
  for (let i = 0; i < info.rows; i++) {
    const h = pnHash_(pick(current[i]).map(pnV2Canon_));
    if (h !== pnHash_(backup[i].slice(0, w).map(pnV2Canon_)) && h !== pnHash_(backup[i].slice(w).map(pnV2Canon_)))
      throw new Error('Có sửa tay sau khi dọn lỗi ở dòng ' + (i + 2) + '; cần đối chiếu, không ghi đè.');
  }
  if (pnV2Read_(M).slice(info.rows).some(r => pnV2Rec_(M, cfg, r))) throw new Error('Có dòng mới sau lượt dọn; cần đối chiếu trước.');
  const wanted = backup.map(r => r.slice(w));
  pnV2WriteCols_(M, info.cols, wanted);
  props.deleteProperty('PN_CLEANUP_PENDING');
  return {ok: true, sheet: p.main, completed: true};
}

const PN_VHFF_JOBS_ = ['reject', 'documents', 'products', 'incidents', 'defects'];
function pnV2VhffStep_() {
  const props = PropertiesService.getScriptProperties();
  const dirty = props.getProperty('PN_VHFF_DIRTY');
  let cycle = JSON.parse(props.getProperty('PN_VHFF_CYCLE') || 'null');
  if (!cycle) { if (!dirty) return {ok: true, skipped: true}; cycle = {revision: dirty, next: 0}; }
  const name = PN_VHFF_JOBS_[cycle.next];
  const run = {
    reject: () => syncAllChungTuKhongDatYC(), documents: () => syncBanGiaoChungTu(), products: () => syncBanGiaoSanPham(),
    incidents: () => syncAllSuVuBanhXep_(), defects: () => syncAllHangLoiToGHTK()
  }[name];
  try {
    const result = run();
    if (result && result.ok === false) throw new Error(JSON.stringify(result));
  } catch (err) { props.setProperty('PN_VHFF_ERROR', name + ': ' + String(err.message || err)); throw err; }
  cycle.next++;
  if (cycle.next < PN_VHFF_JOBS_.length) { props.setProperty('PN_VHFF_CYCLE', JSON.stringify(cycle)); return {ok: true, job: name, pending: true}; }
  refreshBanGiaoChungTuDropdown(); syncBanGiaoChungTuBySelectedDate();
  props.deleteProperty('PN_VHFF_CYCLE');
  if (props.getProperty('PN_VHFF_DIRTY') === cycle.revision) {
    props.deleteProperty('PN_VHFF_DIRTY'); props.deleteProperty('PN_VHFF_ERROR'); props.setProperty('PN_VHFF_ACK', cycle.revision);
  }
  return {ok: true, job: name, done: true};
}

/* Liệt kê dòng "không phải hồ sơ" (có ID ẩn nhưng không có mã định danh) mà bản đồng bộ cũ đã tạo,
 * ví dụ dòng chỉ có giá trị mặc định bị đẩy sang full với Tháng = "CẦN KIỂM TRA". Chỉ liệt kê, không xóa. */
function pnListNonRecordRows() {
  const out = [];
  PN_FULL.pairs.forEach(p => {
    const cfg = pnV2Cfg_(p);
    [[p.main, false], [p.full, true]].forEach(([name, isFull]) => {
      const L = pnV2Layout_(pnSheet_(name), isFull), rows = [];
      pnV2Read_(L).forEach((r, i) => { if (L.id != null && pnText_(r[L.id]) && !pnV2Rec_(L, cfg, r)) rows.push(i + 2); });
      out.push({sheet: name, rows: rows.slice(0, 500), count: rows.length});
    });
  });
  console.log(JSON.stringify(out));
  return out;
}

/* Đồng bộ đúng MỘT hồ sơ (dòng rowNo trên tab side) giữa tab chính và full. Chỉ đọc cột ID của tab kia và đúng
 * 2 dòng liên quan, nên giữ khóa rất ngắn. prefer = true: bên side thắng khi khác nhau (vừa sửa tay / API vừa ghi).
 * Trường hợp hiếm (chưa có ID, ID trùng, hồ sơ chưa có ở tab kia) thì chuyển sang đồng bộ cả cặp như cũ. */
function pnV2SyncOne_(p, side, rowNo, prefer) { return pnV2SyncRows_(p, side, [rowNo], prefer); }
// Giá trị một cột (theo tên) của cả tab, đã chuẩn hóa. Dùng để biết script nghiệp vụ vừa đổi những dòng nào.
function pnV2ColSnap_(sh, isFull, aliases) {
  const L = pnV2Layout_(sh, isFull), c = pnV2Col_(L, aliases), last = sh.getLastRow();
  if (c == null || last < 2) return [];
  return sh.getRange(2, c + 1, last - 1, 1).getValues().map(r => pnV2Canon_(r[0]));
}
function pnV2ColDiff_(before, after) {
  const out = [];
  for (let i = 0; i < Math.max(before.length, after.length); i++) if ((before[i] || '') !== (after[i] || '')) out.push(i + 2);
  return out;
}
/* Đồng bộ ĐÚNG các dòng rowList (trên tab side) với tab còn lại. Chỉ đọc: khối dòng vừa đổi, cột ID/mã của tab kia
 * và các dòng tương ứng. Dòng mới (chưa có ID) -> thêm vào tab kia; dòng đã có -> cập nhật đúng dòng cùng ID.
 * prefer = true: bên side thắng khi hai bên khác nhau (vừa sửa tay / API / script vừa ghi). */
function pnV2SyncRows_(p, side, rowList, prefer) {
  const res = {ok: true, changed: 0, added: 0, conflicts: 0, otherRows: []};
  const rowsIn = Array.from(new Set((rowList || []).map(Number).filter(r => r >= 2))).sort((a, b) => a - b);
  if (!rowsIn.length) return res;
  const cfg = pnV2Cfg_(p), month = pnCurrentMonth_(), L = pnV2EnsureColumns_(p), M = L.M, F = L.F, names = pnV2Names_(M);
  const src = side === 'main' ? M : F, dst = side === 'main' ? F : M;
  const rows = rowsIn.filter(r => r <= src.sh.getLastRow());
  if (!rows.length) return res;
  const colVals = (Lx, i) => { const last = Lx.sh.getLastRow(); return last < 2 || i == null ? [] : Lx.sh.getRange(2, i + 1, last - 1, 1).getValues().map(r => r[0]); };
  const r1 = rows[0], r2 = rows[rows.length - 1];
  const block = src.sh.getRange(r1, 1, r2 - r1 + 1, pnV2ReadWidth_(src)).getValues();
  const srcIds = colVals(src, src.id).map(pnText_);
  const dstIds = colVals(dst, dst.id).map(pnText_);
  // Chỉ mục bên kia theo ID; ID xuất hiện 2 lần ở bên kia -> đồng bộ cả cặp (trường hợp hiếm, an toàn).
  const dstById = new Map();
  for (let i = 0; i < dstIds.length; i++) if (dstIds[i]) { if (dstById.has(dstIds[i])) return Object.assign(pnV2SyncPair_(p, prefer ? {prefer: {side, rows: new Set(rows)}} : {}), {fallback: true}); dstById.set(dstIds[i], i + 2); }
  // Dòng cuối có dữ liệu của bên kia (để thêm dòng mới ngay sau). Chỉ đọc khi thật sự cần thêm dòng.
  const lastDataRow = () => {
    const identVals = (cfg.ident || []).map(x => pnV2Col_(dst, [x])).filter(c => c != null).map(c => colVals(dst, c));
    for (let i = Math.max(dstIds.length, ...identVals.map(v => v.length)) - 1; i >= 0; i--)
      if (dstIds[i] || identVals.some(v => pnV2Canon_(v[i]) !== '')) return i + 2;
    return 1;
  };
  // Nối dòng chính chưa có ID với hồ sơ full cùng mã (không phải File đơn), để không tạo trùng.
  let keyMap = null;
  const keyOf = () => {
    if (keyMap) return keyMap;
    keyMap = new Map();
    if (cfg.repeated || side !== 'main') return keyMap;
    const k = pnV2Col_(dst, cfg.key), gk = cfg.ghtk ? pnV2Col_(dst, cfg.ghtk) : null;
    const kv = k == null ? [] : colVals(dst, k), gv = gk == null ? [] : colVals(dst, gk);
    for (let i = 0; i < Math.max(kv.length, gv.length); i++) {
      const key = pnText_(kv[i]) ? (cfg.docs ? 'ORDER:' : '') + pnText_(kv[i]).toLowerCase() : (pnText_(gv[i]) ? 'GHTK:' + pnText_(gv[i]).toLowerCase() : '');
      if (key && dstIds[i] && !keyMap.has(key)) keyMap.set(key, dstIds[i]);
    }
    return keyMap;
  };
  const claimed = new Set(srcIds.filter(Boolean));
  const srcUp = new Map(), dstUp = new Map();
  const M_ = side === 'main' ? {L: M, up: srcUp} : {L: M, up: dstUp}, F_ = side === 'main' ? {L: F, up: dstUp} : {L: F, up: srcUp};
  let next = 0, appendStart = 0;
  const nextRow = () => { if (!next) { next = lastDataRow() + 1; appendStart = next; } return next++; };
  const pending = [];
  rows.forEach(r => {
    const row = block[r - r1];
    if (!pnV2Rec_(src, cfg, row)) return;
    let id = pnText_(row[src.id]);
    if (id && srcIds.indexOf(id) + 2 !== r) id = ''; // dòng copy/dán kèm ID của dòng khác: hồ sơ mới
    if (!id && side === 'main' && !cfg.repeated) {
      const k = pnV2Key_(src, cfg, row), found = k && keyOf().get(k);
      if (found && !claimed.has(found)) { id = found; claimed.add(found); }
    }
    const sigS = pnV2Sig_(src, row, names), hashS = pnHash_(sigS);
    if (id && dstById.has(id)) { pending.push({r, row, id, dstNo: dstById.get(id), sigS, hashS}); if (pnText_(row[src.id]) !== id) pnV2Put_(srcUp, r, src.id, id); return; }
    // Chưa có ở bên kia.
    if (!id) { id = pnV2NewId_(p); pnV2Put_(srcUp, r, src.id, id); }
    else if (pnText_(row[src.id]) !== id) pnV2Put_(srcUp, r, src.id, id);
    if (side === 'full' && !pnV2NeedsMain_(F, cfg, row, month)) {
      // Hồ sơ lịch sử không thuộc tab chính: chỉ bảo đảm ID, base, Tháng ở full.
      pnV2Put_(srcUp, r, F.base, hashS);
      if (F.month != null) pnV2Put_(srcUp, r, F.month, pnV2Month_(F, cfg, row) || 'CẦN KIỂM TRA');
      return;
    }
    const t = nextRow();
    names.forEach(n => { const s = src.cols[n], d = dst.cols[n]; if (d != null) pnV2Put_(dstUp, t, d, s == null ? '' : row[s]); });
    pnV2Put_(dstUp, t, dst.id, id); pnV2Put_(dstUp, t, dst.base, hashS); pnV2Put_(srcUp, r, src.base, hashS);
    if (F.month != null) pnV2Put_(F_.up, side === 'main' ? t : r, F.month, pnV2Month_(src, cfg, row) || 'CẦN KIỂM TRA');
    res.added++; res.otherRows.push(t);
  });
  // Hồ sơ đã có ở cả hai bên: đọc các dòng bên kia (ít thì đọc từng dòng, nhiều thì đọc cả tab một lần).
  if (pending.length) {
    const dstRows = new Map();
    if (pending.length > 20) pnV2Read_(dst).forEach((rw, i) => dstRows.set(i + 2, rw));
    else pending.forEach(x => dstRows.set(x.dstNo, dst.sh.getRange(x.dstNo, 1, 1, pnV2ReadWidth_(dst)).getValues()[0]));
    pending.forEach(x => {
      const drow = dstRows.get(x.dstNo), sigD = pnV2Sig_(dst, drow, names), hashD = pnHash_(sigD);
      const mRow = side === 'main' ? x.row : drow, fRow = side === 'main' ? drow : x.row;
      const mNo = side === 'main' ? x.r : x.dstNo, fNo = side === 'main' ? x.dstNo : x.r;
      const hashM = side === 'main' ? x.hashS : hashD, hashF = side === 'main' ? hashD : x.hashS;
      const mUp = M_.up, fUp = F_.up;
      let finalRow = x.row;
      if (x.sigS === sigD) {
        if (pnText_(mRow[M.base]) !== hashM) pnV2Put_(mUp, mNo, M.base, hashM);
        if (pnText_(fRow[F.base]) !== hashM) pnV2Put_(fUp, fNo, F.base, hashM);
      } else {
        const base = pnText_(mRow[M.base]) || pnText_(fRow[F.base]);
        const natural = base === hashF ? 'push' : base === hashM ? 'pull' : null;
        const dir = prefer ? (side === 'main' ? 'push' : 'pull') : (natural || 'push');
        if (dir !== natural) { res.conflicts++; pnV2ConflictLog_(p, dir === 'push' ? p.full : p.main, x.id, dir === 'push' ? pnV2Sig_(F, fRow, names) : pnV2Sig_(M, mRow, names)); }
        const from = dir === 'push' ? M : F, to = dir === 'push' ? F : M, fromRow = dir === 'push' ? mRow : fRow, toRow = dir === 'push' ? fRow : mRow;
        const toUp = dir === 'push' ? fUp : mUp, toNo = dir === 'push' ? fNo : mNo, hash = dir === 'push' ? hashM : hashF;
        names.forEach(n => {
          const s = from.cols[n], d = to.cols[n];
          if (d == null) return;
          const v = s == null ? '' : fromRow[s];
          if (pnV2Canon_(toRow[d]) !== pnV2Canon_(v)) pnV2Put_(toUp, toNo, d, v);
        });
        pnV2Put_(mUp, mNo, M.base, hash); pnV2Put_(fUp, fNo, F.base, hash);
        finalRow = fromRow; res.changed++;
      }
      // Tháng (tab full) theo nội dung cuối cùng của hồ sơ.
      if (F.month != null) {
        const want = pnV2Month_(finalRow === mRow ? M : F, cfg, finalRow) || 'CẦN KIỂM TRA';
        if (pnText_(fRow[F.month]) !== want) pnV2Put_(fUp, fNo, F.month, want);
      }
      res.otherRows.push(x.dstNo);
    });
  }
  const writes = pnV2Flush_(dst.sh, dstUp) + pnV2Flush_(src.sh, srcUp);
  if (next > appendStart && appendStart) pnV2CopyFormat_(dst, appendStart, next - appendStart, dst.width);
  if (writes) { SpreadsheetApp.flush(); if (res.changed || res.added) pnMarkDirty_(); }
  return res;
}
// Cột tính bằng code cho đúng các dòng vừa đổi (Booking!Check; File đơn!Quản lý/Khu vực/Ngay).
function pnV2ComputeRows_(name, rowList) {
  if (PropertiesService.getScriptProperties().getProperty('PN_FORMULAS_CONVERTED') !== '1') return 0;
  const p = pnPair_(name);
  if (!p || p.main === 'Chứng từ_FF') return 0;
  const rows = Array.from(new Set(rowList || [])).filter(r => r >= 2).sort((a, b) => a - b);
  if (!rows.length) return 0;
  const sh = pnSheet_(name), L = pnV2Layout_(sh, name === p.full), r1 = rows[0], r2 = rows[rows.length - 1];
  if (r1 > sh.getLastRow()) return 0;
  const block = sh.getRange(r1, 1, r2 - r1 + 1, pnV2ReadWidth_(L)).getValues(), up = new Map();
  if (p.main === 'Booking') {
    const d = pnV2Col_(L, ['Ngày']), total = pnV2Col_(L, ['Tổng bánh']), out = pnV2Col_(L, ['Check']);
    if (d == null || total == null || out == null) return 0;
    const factor = pnSheet_('TT Nhập').getRange('L3').getValue();
    if (typeof factor !== 'number' || !Number.isFinite(factor)) throw new Error('TT Nhập!L3 không phải số hợp lệ.');
    rows.forEach(r => { const row = block[r - r1]; const v = row[d] === '' ? '' : pnRoundUp_(Number(row[total] || 0) * factor); if (pnV2Canon_(v) !== pnV2Canon_(row[out])) pnV2Put_(up, r, out, v); });
  } else {
    const code = pnV2Col_(L, ['Mã đơn']), created = pnV2Col_(L, ['Thời gian tạo đơn']), depot = pnV2Col_(L, ['Kho đích']);
    const qm = pnV2Col_(L, ['Quản lý']), kv = pnV2Col_(L, ['Khu vực']), day = pnV2Col_(L, ['Ngay']);
    if (code == null || created == null || depot == null) return 0;
    const byDepot = new Map();
    pnRows_(pnSheet_('DS BC'), 15).forEach(r => { const k = pnLookupKey_(r[6]); if (!byDepot.has(k)) byDepot.set(k, r); });
    rows.forEach(r => {
      const row = block[r - r1], empty = row[code] === '', match = byDepot.get(pnLookupKey_(row[depot])), iso = pnIsoDate_(row[created]);
      [[qm, empty ? '' : (match ? match[4] : '#N/A')], [kv, empty ? '' : (match ? match[14] : '#N/A')],
       [day, empty ? '' : (iso ? iso.slice(8) + '/' + iso.slice(5, 7) + '/' + iso.slice(0, 4) : '#VALUE!')]]
        .forEach(([c, v]) => { if (c != null && pnV2Canon_(v) !== pnV2Canon_(row[c])) pnV2Put_(up, r, c, v); });
    });
  }
  return pnV2Flush_(sh, up);
}
/* Sửa/dán tay trên Booking, File đơn, Chứng từ (chính hoặc full): chỉ xử lý các dòng vừa đổi và những dòng
 * mà script nghiệp vụ đổi theo (đơn chứng từ mới từ Booking, ngày lên đơn, mã GHTK, khu vực). */
function pnV2HandleRows_(name, r1, r2) {
  const p = pnPair_(name), side = name === p.main ? 'main' : 'full', rows = [];
  for (let r = r1; r <= r2; r++) rows.push(r);
  const docs = PN_FULL.pairs[2], docsFull = pnSheet_(docs.full);
  if (p.main === 'Chứng từ_FF') return pnV2SyncRows_(p, side, rows, true);
  pnV2ComputeRows_(name, rows);
  const own = pnV2SyncRows_(p, side, rows, true);
  const docRows = [];
  if (p.main === 'Booking') {
    const lastBefore = docsFull.getLastRow(), datesBefore = pnV2ColSnap_(docsFull, true, ['Ngày lên đơn']);
    if (side === 'main') { syncBookingRowsToChungTuFF_(r1, r2); bookingSyncUpdateTargetDatesFromBookingRows_(r1, r2); }
    else pnSyncBookingFullRows_(r1, r2);
    for (let r = lastBefore + 1; r <= docsFull.getLastRow(); r++) docRows.push(r);
    pnV2ColDiff_(datesBefore, pnV2ColSnap_(docsFull, true, ['Ngày lên đơn'])).forEach(r => docRows.push(r));
  }
  // Mã GHTK: File đơn mới hoặc đơn chứng từ mới có thể được điền mã.
  const ghtkBefore = pnV2ColSnap_(docsFull, true, ['Mã đơn GHTK']);
  fillMissingGhtkCodesInChungTuFF();
  pnV2ColDiff_(ghtkBefore, pnV2ColSnap_(docsFull, true, ['Mã đơn GHTK'])).forEach(r => docRows.push(r));
  const docsRes = pnV2SyncRows_(docs, 'full', docRows, true);
  // Khu vực Booking phụ thuộc Booking và File đơn: chỉ đẩy các dòng Booking vừa đổi khu vực.
  const areaRes = [];
  ['Booking', 'Booking_full'].forEach(bn => {
    if (p.main === 'Booking' && bn !== name) return;
    const sh = pnSheet_(bn), before = pnV2ColSnap_(sh, bn === 'Booking_full', ['Khu vực']), prev = KHUVUC_BOOKING_SYNC_CONFIG.bookingSheetName;
    try { KHUVUC_BOOKING_SYNC_CONFIG.bookingSheetName = bn; syncAllKhuVucBookingToChungTuFF_(); }
    finally { KHUVUC_BOOKING_SYNC_CONFIG.bookingSheetName = prev; }
    const changed = pnV2ColDiff_(before, pnV2ColSnap_(sh, bn === 'Booking_full', ['Khu vực']));
    if (changed.length) areaRes.push(pnV2SyncRows_(PN_FULL.pairs[0], bn === 'Booking' ? 'main' : 'full', changed, true));
  });
  return {ok: true, own, docs: docsRes, area: areaRes};
}

/* Sửa tay đúng lúc khóa bận: lưu lại đúng tab + dòng đã sửa (mỗi lần một khóa riêng, không ghi đè nhau),
 * và xử lý ngay khi có ai giữ được khóa: lần sửa tiếp theo, lần mở Sheet, hoặc lần app lưu đơn kế tiếp. */
function pnQueueEdit_(name, r1, r2) {
  const key = 'PN_EDITQ_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8);
  PropertiesService.getScriptProperties().setProperty(key, JSON.stringify({name, r1, r2}));
}
function pnDrainEditQueue_(maxMs) {
  const props = PropertiesService.getScriptProperties(), all = props.getProperties(), started = Date.now();
  const keys = Object.keys(all).filter(k => k.indexOf('PN_EDITQ_') === 0).sort();
  let done = 0;
  for (const k of keys) {
    if (maxMs && Date.now() - started > maxMs) break;
    if (pnV2OverBudget_()) break;
    let job = null;
    try { job = JSON.parse(all[k]); } catch (err) {}
    if (job && job.name && pnPair_(job.name)) pnV2HandleRows_(job.name, Math.max(2, Number(job.r1) || 2), Math.max(Number(job.r1) || 2, Number(job.r2) || 2));
    props.deleteProperty(k);
    done++;
  }
  return {done, left: keys.length - done};
}
/*************************************************
CODE XOA DU LIEU (Booking / Booking_full / Chứng từ_FF / Chứng từ_full) — file rieng, KHONG sua cac file khac.

Project nay KHONG gan voi Google Sheet nen khong tu hien menu duoc.
Menu "PN Công cụ" + hop thoai nam o PROJECT GAN VOI SHEET (project OCR), file "PN_Menu_Xoa",
goi sang 2 ham duoi day qua Thu vien (ten goi: PN):
  PN.pnMenuXoaXemTruoc(req)  -> chi doc, tra ve so dong se xoa o tung tab.
  PN.pnMenuXoaThucHien(req)  -> xoa that (req.expect = ket qua luc xem truoc; khac -> dung).

- Xoa theo SO DON HANG (Booking: cot "Row Labels", Chung tu: cot "Số đơn hàng") hoac theo NGAY
  (Booking: cot "Ngày", Chung tu: cot "Ngày lên đơn"), chi o cac tab da chon.
- Chi xoa NGUYEN DONG cua dung cac dong khop, trong 1 lan giu khoa cua project nay
  (khoa + Script Properties cua thu vien dung chung voi trigger cua project nay).
*************************************************/

const PN_XOA_TABS = {
  'Booking':       {pair: 0, full: false, key: ['Row Labels'],  date: ['Ngày']},
  'Booking_full':  {pair: 0, full: true,  key: ['Row Labels'],  date: ['Ngày']},
  'Chứng từ_FF':   {pair: 2, full: false, key: ['Số đơn hàng'], date: ['Ngày lên đơn']},
  'Chứng từ_full': {pair: 2, full: true,  key: ['Số đơn hàng'], date: ['Ngày lên đơn']}
};

// Goi tu menu (qua thu vien): chi doc, tra ve so dong se xoa o tung tab.
function pnMenuXoaXemTruoc(req) {
  return pnWithLock_(() => { pnRequireReady_(); return pnXoaPlan_(req).summary; }, 30000);
}

// Goi tu menu (qua thu vien): xoa that. req.expect = so dong tung tab luc xem truoc; khac -> dung.
function pnMenuXoaThucHien(req) {
  return pnWithLock_(() => {
    pnRequireReady_();
    const plan = pnXoaPlan_(req);
    if (JSON.stringify(plan.summary.counts) !== JSON.stringify(req.expect || {}))
      throw new Error('Dữ liệu vừa thay đổi sau khi xem trước. Bấm "Xem trước" lại.');
    // Xoa tab full truoc, roi tab chinh.
    plan.tabs.slice().sort((a, b) => Number(b.cfg.full) - Number(a.cfg.full)).forEach(t => pnXoaDeleteRows_(t.L, t.rows, !t.cfg.full));
    SpreadsheetApp.flush();
    pnRunInvalidate_();
    // Tab full da xoa: cap nhat so ho so da biet, de dong bo khong bao "full giam".
    plan.tabs.filter(t => t.cfg.full && t.rows.length).forEach(t => {
      const p = PN_FULL.pairs[t.cfg.pair], cfg = pnV2Cfg_(p), F = pnV2Layout_(pnSheet_(p.full), true);
      PropertiesService.getScriptProperties().setProperty('PN_FULL_COUNT_' + p.main, String(pnV2Read_(F).filter(r => pnV2Rec_(F, cfg, r)).length));
    });
    pnMarkDirty_();
    console.log('PN Xoa du lieu: ' + JSON.stringify(plan.summary));
    return plan.summary;
  }, 60000);
}

// Lap danh sach dong can xoa (chi doc).
function pnXoaPlan_(req) {
  req = req || {};
  const names = (req.tabs || []).filter(n => PN_XOA_TABS[n]);
  if (!names.length) throw new Error('Chưa chọn tab nào.');
  let match, label;
  if (req.mode === 'ngay') {
    const iso = pnIsoDate_(req.date);
    if (!iso) throw new Error('Ngày không hợp lệ. Nhập dạng dd/MM/yyyy, ví dụ 27/10/2026.');
    label = 'ngày ' + iso.slice(8) + '/' + iso.slice(5, 7) + '/' + iso.slice(0, 4);
    match = (L, cfg, r) => { const c = pnV2Col_(L, cfg.date); return c != null && pnIsoDate_(r[c]) === iso; };
  } else {
    const codes = String(req.codes || '').split(/[\s,;]+/).map(pnText_).filter(Boolean);
    if (!codes.length) throw new Error('Chưa nhập Số đơn hàng.');
    const set = new Set(codes.map(pnNorm_));
    label = codes.length + ' số đơn hàng';
    match = (L, cfg, r) => { const c = pnV2Col_(L, cfg.key); return c != null && set.has(pnNorm_(r[c])); };
    req._codes = codes;
  }
  const found = new Set(), tabs = [], counts = {}, warnings = [], missingCols = [];
  names.forEach(name => {
    const cfg = PN_XOA_TABS[name], sh = pnSheet_(name), L = pnV2Layout_(sh, cfg.full);
    const keyCol = pnV2Col_(L, cfg.key), dateCol = pnV2Col_(L, cfg.date);
    if (keyCol == null || (req.mode === 'ngay' && dateCol == null)) { missingCols.push(name); return; }
    const appCols = cfg.pair === 2 ? ['Link ảnh', 'Thời gian', 'User thao tác'].map(h => pnV2Col_(L, [h])).filter(c => c != null) : [];
    const rows = [];
    let withApp = 0;
    pnV2Read_(L).forEach((r, i) => {
      if (!match(L, cfg, r)) return;
      rows.push(i + 2);
      found.add(pnNorm_(r[keyCol]));
      if (appCols.some(c => pnText_(r[c]))) withApp++;
    });
    counts[name] = rows.length;
    tabs.push({name, cfg, L, rows});
    if (withApp) warnings.push(name + ': ' + withApp + ' dòng đã có dữ liệu app (ảnh / thời gian / user thao tác) cũng sẽ bị xóa.');
  });
  if (missingCols.length) throw new Error('Không thấy cột cần dùng ở tab: ' + missingCols.join(', '));
  [['Booking', 'Booking_full'], ['Chứng từ_FF', 'Chứng từ_full']].forEach(([m, f]) => {
    if (names.includes(m) !== names.includes(f))
      warnings.push('Chỉ chọn ' + (names.includes(m) ? m : f) + ' mà không chọn ' + (names.includes(m) ? f : m) + ': đồng bộ có thể chép lại các dòng vừa xóa từ tab kia.');
  });
  const notFound = (req._codes || []).filter(c => !found.has(pnNorm_(c)));
  return {tabs, summary: {label, counts, total: tabs.reduce((s, t) => s + t.rows.length, 0), notFound, warnings}};
}

// Xoa dong tu duoi len theo nhom lien nhau. Tab chinh co du lieu ben phai bang (ngoai cot ID)
// thi chi xoa o trong bang + cot ID/BASE (giong luot don thang cu); con lai xoa nguyen dong.
function pnXoaDeleteRows_(L, rows, isMain) {
  if (!rows.length) return;
  const side = isMain && pnV2HasSideContent_(L);
  const tech = [L.id, L.base].filter(c => c != null);
  for (let i = rows.length - 1; i >= 0;) {
    let j = i;
    while (j > 0 && rows[j - 1] === rows[j] - 1) j--;
    const start = rows[j], n = rows[i] - rows[j] + 1;
    if (!side) L.sh.deleteRows(start, n);
    else {
      L.sh.getRange(start, 1, n, L.width).deleteCells(SpreadsheetApp.Dimension.ROWS);
      tech.forEach(c => L.sh.getRange(start, c + 1, n, 1).deleteCells(SpreadsheetApp.Dimension.ROWS));
    }
    i = j - 1;
  }
}

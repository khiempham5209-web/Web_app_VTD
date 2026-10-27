/*************************************************
MENU "PN Công cụ" > "Xóa dữ liệu…"  — file rieng, KHONG sua cac file code khac.

- Tich chon tab can xoa: Booking, Booking_full, Chứng từ_FF, Chứng từ_full (1, vai hoac ca 4 tab).
- Xoa theo SO DON HANG (dan 1 hoac nhieu ma; Booking tim cot "Row Labels", Chung tu tim cot "Số đơn hàng")
  hoac theo NGAY (Booking: cot "Ngày"; Chung tu: cot "Ngày lên đơn").
- Chi xoa NGUYEN DONG cua dung cac dong khop, o dung cac tab da tich. Khong dung gi khac.
- "Xem trước" -> thay so dong se xoa o tung tab -> "Xóa" moi xoa that.
  Toan bo trong 1 lan giu khoa (khong luot dong bo nao chen vao giua). Du lieu doi sau khi xem truoc -> dung, khong xoa.
- Tich tab chinh ma khong tich tab full tuong ung (hoac nguoc lai): co canh bao vi dong bo co the chep lai.

Menu hien ra khi mo lai Sheet (ham onOpen). Lan dau bam menu, Google hoi cap quyen.
*************************************************/

const PN_XOA_TABS = {
  'Booking':       {pair: 0, full: false, key: ['Row Labels'],  date: ['Ngày']},
  'Booking_full':  {pair: 0, full: true,  key: ['Row Labels'],  date: ['Ngày']},
  'Chứng từ_FF':   {pair: 2, full: false, key: ['Số đơn hàng'], date: ['Ngày lên đơn']},
  'Chứng từ_full': {pair: 2, full: true,  key: ['Số đơn hàng'], date: ['Ngày lên đơn']}
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('PN Công cụ').addItem('Xóa dữ liệu…', 'pnMenuXoaMo').addToUi();
}

function pnMenuXoaMo() {
  const html = HtmlService.createHtmlOutput(PN_XOA_HTML_).setWidth(560).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Xóa dữ liệu (Booking / Chứng từ)');
}

// Goi tu hop thoai: chi doc, tra ve so dong se xoa o tung tab.
function pnMenuXoaXemTruoc(req) {
  return pnWithLock_(() => { pnRequireReady_(); return pnXoaPlan_(req).summary; }, 30000);
}

// Goi tu hop thoai: xoa that. req.expect = so dong tung tab luc xem truoc; khac -> dung.
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

const PN_XOA_HTML_ = String.raw`<!DOCTYPE html><html><head><base target="_top"><style>
body{font:13px Arial,sans-serif;margin:12px;color:#222}
fieldset{border:1px solid #ccc;border-radius:6px;margin:0 0 10px;padding:8px 10px}
legend{font-weight:bold;padding:0 4px}
label{display:block;margin:3px 0}
textarea{width:100%;height:110px;box-sizing:border-box;font:inherit}
input[type=text]{width:140px;font:inherit;padding:3px}
button{padding:7px 14px;margin-right:6px;border:0;border-radius:4px;background:#1a73e8;color:#fff;cursor:pointer}
button.del{background:#d93025}button:disabled{background:#aaa;cursor:default}
#msg{margin-top:10px;white-space:pre-line}.err{color:#c5221f}.ok{color:#188038}.warn{color:#b06000}
.hide{display:none}
</style></head><body>
<fieldset><legend>1. Chọn tab cần xóa</legend>
<label><input type="checkbox" class="tab" value="Booking"> Booking</label>
<label><input type="checkbox" class="tab" value="Booking_full"> Booking_full</label>
<label><input type="checkbox" class="tab" value="Chứng từ_FF"> Chứng từ_FF</label>
<label><input type="checkbox" class="tab" value="Chứng từ_full"> Chứng từ_full</label>
</fieldset>
<fieldset><legend>2. Xóa theo</legend>
<label><input type="radio" name="mode" value="od" checked> Số đơn hàng</label>
<label><input type="radio" name="mode" value="ngay"> Ngày</label>
<div id="boxOd"><textarea id="codes" placeholder="Dán 1 hoặc nhiều Số đơn hàng, mỗi dòng 1 mã (hoặc cách nhau bằng dấu phẩy)"></textarea></div>
<div id="boxNgay" class="hide">Ngày: <input type="text" id="ngay" placeholder="dd/MM/yyyy"></div>
</fieldset>
<button id="btnXem">Xem trước</button><button id="btnXoa" class="del" disabled>Xóa</button>
<div id="msg"></div>
<script>
var last = null;
function $(id){return document.getElementById(id);}
function req(){
  var tabs=[].slice.call(document.querySelectorAll('.tab')).filter(function(x){return x.checked;}).map(function(x){return x.value;});
  var mode=document.querySelector('input[name=mode]:checked').value;
  return {tabs:tabs, mode:mode, codes:$('codes').value, date:$('ngay').value};
}
function msg(t,c){$('msg').className=c||'';$('msg').textContent=t;}
function reset(){last=null;$('btnXoa').disabled=true;$('btnXoa').textContent='Xóa';}
[].slice.call(document.querySelectorAll('input,textarea')).forEach(function(x){x.addEventListener('input',reset);x.addEventListener('change',reset);});
[].slice.call(document.querySelectorAll('input[name=mode]')).forEach(function(x){x.addEventListener('change',function(){
  var od=document.querySelector('input[name=mode]:checked').value==='od';
  $('boxOd').className=od?'':'hide';$('boxNgay').className=od?'hide':'';});});
$('btnXem').onclick=function(){
  reset();var r=req();$('btnXem').disabled=true;msg('Đang kiểm tra...');
  google.script.run.withSuccessHandler(function(s){
    $('btnXem').disabled=false;
    var lines=['Sẽ xóa theo '+s.label+':'];
    Object.keys(s.counts).forEach(function(k){lines.push('  • '+k+': '+s.counts[k]+' dòng');});
    lines.push('Tổng: '+s.total+' dòng');
    if(s.notFound.length)lines.push('Không tìm thấy ở các tab đã chọn: '+s.notFound.join(', '));
    if(s.warnings.length)lines.push('','⚠️ '+s.warnings.join('\n⚠️ '));
    msg(lines.join('\n'),s.warnings.length?'warn':'');
    if(s.total){last={req:r,expect:s.counts};$('btnXoa').disabled=false;$('btnXoa').textContent='Xóa '+s.total+' dòng';}
  }).withFailureHandler(function(e){$('btnXem').disabled=false;msg('❌ '+e.message,'err');}).pnMenuXoaXemTruoc(r);
};
$('btnXoa').onclick=function(){
  if(!last)return;var r=last.req;r.expect=last.expect;
  $('btnXoa').disabled=true;$('btnXem').disabled=true;msg('Đang xóa...');
  google.script.run.withSuccessHandler(function(s){
    $('btnXem').disabled=false;last=null;$('btnXoa').textContent='Xóa';
    var lines=['✅ Đã xóa theo '+s.label+':'];
    Object.keys(s.counts).forEach(function(k){lines.push('  • '+k+': '+s.counts[k]+' dòng');});
    msg(lines.join('\n'),'ok');
  }).withFailureHandler(function(e){$('btnXem').disabled=false;msg('❌ '+e.message,'err');}).pnMenuXoaThucHien(r);
};
</script></body></html>`;

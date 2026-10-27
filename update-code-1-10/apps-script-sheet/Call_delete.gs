/*************************************************
MENU "Delete" > "Xóa dữ liệu…"  — dan vao PROJECT GAN VOI GOOGLE SHEET (project co OCR).

File nay CHI la giao dien (menu + hop thoai). Code xoa that nam o project PJ_Phạm Nguyên
(file 14_Menu_Xoa), goi qua Thu vien ten goi "PN".

Cai dat (1 lan), trong project OCR:
  1. Thu vien (+) > dan Script ID cua PJ_Phạm Nguyên > Tra cuu
     > Phien ban: chon ban moi nhat (hoac "HEAD (Che do phat trien)") > Ma nhan dang: PN > Them.
  2. Them tep moi "PN_Menu_Xoa", dan file nay vao, Luu.
  3. Tai lai Google Sheet -> menu "Delete" hien canh menu "OCR".
     Lan dau bam menu, Google hoi cap quyen.
*************************************************/

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Delete').addItem('Xóa dữ liệu…', 'pnMenuXoaMo').addToUi();
}

function pnMenuXoaMo() {
  const html = HtmlService.createHtmlOutput(PN_MENU_XOA_HTML_).setWidth(560).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Xóa dữ liệu (Booking / Chứng từ)');
}

// Hop thoai goi 2 ham nay; 2 ham chuyen sang thu vien PN (project PJ_Phạm Nguyên).
function pnMenuXoaGoiXemTruoc(req) { return PN.pnMenuXoaXemTruoc(req); }
function pnMenuXoaGoiThucHien(req) { return PN.pnMenuXoaThucHien(req); }

const PN_MENU_XOA_HTML_ = String.raw`<!DOCTYPE html><html><head><base target="_top"><style>
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
  }).withFailureHandler(function(e){$('btnXem').disabled=false;msg('❌ '+e.message,'err');}).pnMenuXoaGoiXemTruoc(r);
};
$('btnXoa').onclick=function(){
  if(!last)return;var r=last.req;r.expect=last.expect;
  $('btnXoa').disabled=true;$('btnXem').disabled=true;msg('Đang xóa...');
  google.script.run.withSuccessHandler(function(s){
    $('btnXem').disabled=false;last=null;$('btnXoa').textContent='Xóa';
    var lines=['✅ Đã xóa theo '+s.label+':'];
    Object.keys(s.counts).forEach(function(k){lines.push('  • '+k+': '+s.counts[k]+' dòng');});
    msg(lines.join('\n'),'ok');
  }).withFailureHandler(function(e){$('btnXem').disabled=false;msg('❌ '+e.message,'err');}).pnMenuXoaGoiThucHien(r);
};
</script></body></html>`;

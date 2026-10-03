/* Existing API body is pnApiSaveCore_. This wrapper adds durable idempotency and mirroring.
 * A RUNNING upload with no recorded result is deliberately quarantined: retry must not
 * create duplicate Drive files when the previous execution died after createFile().
 */
function pnJournal_() {
  const ss=pnSS_(),name='_PN_REQUESTS';
  let sh=ss.getSheetByName(name);
  if(!sh){sh=ss.insertSheet(name);sh.getRange(1,1,1,4).setValues([['Key','State','Data','Updated']]);sh.hideSheet();}
  return sh;
}
function pnJournalRead_(key) {
  const sh=pnJournal_();
  for(let attempt=0;attempt<2;attempt++){
    if(sh.getLastRow()<2)return null;
    const found=sh.getRange(2,1,sh.getLastRow()-1,1).createTextFinder(key).matchEntireCell(true).useRegularExpression(false).findAll();
    if(!found.length)return null;
    // Nhật ký chỉ thêm dòng: dòng cuối cùng của khóa là trạng thái mới nhất.
    const last=found.reduce((m,c)=>c.getRow()>m.getRow()?c:m,found[0]);
    const r=sh.getRange(last.getRow(),1,1,4).getValues()[0];
    // Dòng vừa bị dời (dọn nhật ký cũ chạy cùng lúc): tìm lại.
    if(String(r[0])!==key)continue;
    return {row:last.getRow(),state:r[1],data:JSON.parse(r[2]||'{}')};
  }
  return null;
}

// Nhật ký chỉ THÊM dòng (appendRow nguyên tử): nhiều đơn ghi cùng lúc không đè nhau, không cần khóa của Sheet.
// Nhật ký chỉ THÊM dòng (appendRow nguyên tử): nhiều đơn ghi cùng lúc không đè nhau, không cần khóa của Sheet.
// flush=false: không ép lưu ngay (Apps Script tự lưu khi lần chạy kết thúc) -> nhanh hơn.
function pnJournalWrite_(key,state,data,flush) {
  const text=JSON.stringify(data);
  if(text.length>45000)throw new Error('Nhật ký vượt kích thước an toàn; chia nhỏ thao tác.');
  pnJournal_().appendRow([key,state,text,new Date()]);
  if(flush!==false)SpreadsheetApp.flush();
}
// Dọn nhật ký: bỏ các dòng cũ hơn 7 ngày ở đầu sheet (nhật ký ghi theo thời gian). Chạy trong lượt dọn đầu ngày.
function pnJournalPrune_() {
  const sh=pnJournal_(),last=sh.getLastRow();
  if(last<2)return 0;
  const limit=Date.now()-7*24*3600*1000,dates=sh.getRange(2,4,Math.min(last-1,20000),1).getValues();
  let n=0;
  while(n<dates.length&&dates[n][0] instanceof Date&&dates[n][0].getTime()<limit)n++;
  if(n)sh.deleteRows(2,n);
  return n;
}

function pnRequestKey_(id) { return 'request:'+pnHash_(pnText_(id)); }
// Chỉ băm nội dung nghiệp vụ. Bỏ: token phiên, các khóa bắt đầu bằng "_" (app gắn _diagRequestId
// mới cho MỖI lần gửi), số dòng, phiên bản app và dữ liệu ảnh base64 (nén lại có thể khác byte).
// Nếu không bỏ, gửi lại cùng một đơn sau khi mất phản hồi bị từ chối "clientId đã dùng cho nội dung khác".
const PN_SIGNATURE_SKIP_=['sessionToken','rowNumber','appVersion','base64','data','dataUrl'];
function pnSignatureData_(v) {
  if(Array.isArray(v))return v.map(pnSignatureData_);
  if(!v||typeof v!=='object')return v;
  const out={};
  Object.keys(v).sort().forEach(k=>{
    if(k.charAt(0)==='_'||PN_SIGNATURE_SKIP_.indexOf(k)>=0)return;
    out[k]=pnSignatureData_(v[k]);
  });
  return out;
}
function pnRequestSignature_(params) {
  return pnHash_(pnSignatureData_(params||{}));
}
function pnRequestNow_() { return pnRequest_?new Date(pnRequest_.startedAt):new Date(); }
function pnPendingRequests_() {
  const latest=new Map();
  pnRows_(pnJournal_(),4).forEach(r=>{ if(/^request:[a-f0-9]+$/.test(String(r[0])))latest.set(String(r[0]),r[1]); });
  return Array.from(latest.entries()).filter(([,s])=>s==='PENDING'||s==='WRITTEN').map(([k])=>k);
}

const PN_UPLOAD_STALE_MS_=7*60*1000; // Apps Script dừng mọi lần chạy sau 6 phút
// Ảnh đã tải trong lượt này hoặc lượt gửi trước (lưu trong nhật ký của đơn) thì dùng lại, không tải lại.
function pnUploadOnce_(stage,fn) {
  if(!pnRequest_)throw new Error('Upload phải qua apiSave_.');
  const uploads=pnRequest_.uploads||(pnRequest_.uploads={});
  if(uploads[stage])return uploads[stage];
  // Đơn gửi bởi bản code cũ: kết quả tải ảnh nằm ở dòng nhật ký riêng.
  if(pnRequest_.legacy){
    const old=pnJournalRead_(pnRequest_.key+':'+stage);
    if(old&&old.state==='DONE')return uploads[stage]=old.data;
  }
  return uploads[stage]=fn();
}

// Tải ảnh TRƯỚC khi vào khóa chung. Khi vào khóa, pnApiSaveCore_ chỉ đọc lại kết quả đã có
// trong nhật ký. Dữ liệu không hợp lệ thì bỏ qua để core báo lỗi như cũ.
function pnPreUpload_(params) {
  const query=clean_(params.maDon||params.query||params.po||params.orderNo);
  const type=norm_(clean_(params.returnType||params.loaiHoan||'Chung tu'));
  const hasProducts=['san pham','product','san pham + chung tu','product-docs'].indexOf(type)>=0;
  const hasDocuments=['chung tu','docs','san pham + chung tu','product-docs'].indexOf(type)>=0;
  if(!hasDocuments&&!hasProducts)return;
  let found;
  try { const sh=sheet_(); found=pnResolveSave_(sh,headerMap_(headers_(sh)),params); } catch(err) { return; }
  const maDonForFolder=clean_(query||found.record.maDon||found.record.po||found.record.orderNo);
  const files=Array.isArray(params.files)?params.files:[];
  if(hasDocuments&&!files.length)return;
  const items=hasProducts?normalizeProductItems_(params.items):[];
  if(hasProducts&&!items.length)return;
  for(let i=0;i<items.length;i++)if(validateProductItem_(items[i],i+1))return;
  if(hasDocuments)pnUploadOnce_('docs',()=>uploadFiles_(maDonForFolder,files));
  items.forEach((item,i)=>pnUploadOnce_('product:'+i,()=>uploadProductFiles_(found.record.orderNo||params.orderNo||maDonForFolder,item)));
}
function pnItemId_(item,index) { return pnRequest_.clientId+':'+(item.clientItemId||('item:'+index)); }
function pnEncodeRow_(row) { return row.map(v=>v instanceof Date?{date:v.toISOString()}:v); }
function pnDecodeRow_(row) { return row.map(v=>v&&typeof v==='object'&&v.date?new Date(v.date):v); }
// Ghi chứng từ vào đúng dòng của đơn ở Chứng từ_full, theo TÊN cột. Chỉ ghi các ô thay đổi; không đụng cột kỹ thuật.
function pnWriteDocumentOnce_(sh,row,col,params,upload,timeText,user) {
  // Lượt gửi trước đã ghi chứng từ (đã lưu trong nhật ký): không ghi lại, để không đè sửa tay sau đó.
  if(pnRequest_.docWritten)return;
  const L=pnV2Layout_(sh,true),width=pnV2ReadWidth_(L),names=Object.keys(L.cols);
  const current=sh.getRange(row,1,1,width).getValues()[0];
  const values=[
    [['xac thuc hoa don'],clean_(params.xacThuc||params.xacThucHoaDon)],
    [['ghi chu chung tu'],String(params.note||params.ghiChu||'')],
    [['trang thai'],clean_(params.status||params.trangThai)],
    [['thoi gian'],timeText],[['user thao tac'],user]
  ];
  if(upload.linkAnh)values.push([['link anh'],upload.linkAnh]);
  // Mã ecom CT: chỉ ghi khi app gửi giá trị, để app cũ không xóa mã đã nhập tay trên Sheet.
  const maEcomCt=pnEcomCt_(params.maEcomCt);
  if(maEcomCt)values.push([['ma ecom ct'],maEcomCt]);
  const up=new Map();
  values.forEach(([aliases,value])=>{ const i=pnV2Col_(L,aliases); if(i!=null&&pnV2Canon_(current[i])!==pnV2Canon_(value))pnV2Put_(up,row,i,value); });
  pnV2Flush_(sh,up);
  pnRequest_.docWritten=true;
}

function pnProductIdColumn_(sh) {
  const headers=sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getDisplayValues()[0];
  const old=headers.indexOf('__PN_ITEM_ID');
  if(old>=0)return old+1;
  const col=headers.length+1;pnSize_(sh,2,col);
  sh.getRange(1,col).setValue('__PN_ITEM_ID');sh.hideColumns(col);
  return col;
}
function pnAppendProductsOnce_(sh,rows,width) {
  const c=pnProductIdColumn_(sh),existing=new Set();
  const last=sh.getLastRow();
  if(last>=2)sh.getRange(2,c,last-1,1).getValues().forEach(r=>{ if(r[0])existing.add(String(r[0])); });
  rows.filter(r=>!existing.has(String(r[c-1]))).forEach(r=>sh.appendRow(r.slice(0,Math.max(width,c))));
}

/* Tìm dòng của đơn trong Chứng từ_full: chỉ đọc các cột mã (Mã đơn GHTK, Mã PO, Số đơn hàng), không đọc cả bảng.
 * Trong cùng một lượt lưu, lần tìm sau dùng lại số dòng và chỉ đọc lại đúng dòng đó (kiểm tra vẫn đúng đơn). */
let pnResolveMemo_=null;
function pnResolveSave_(sh,col,params) {
  const expected=pnText_(params.syncKey),query=pnText_(params.orderNo||params.maDonGhtk||params.maDon||params.query||params.po);
  if(!expected&&!query)throw new Error('Cần khóa đơn; không lưu chỉ bằng số dòng.');
  const L=pnV2Layout_(sh,true),map=headerMap_(L.headers),width=pnV2ReadWidth_(L);
  const cGhtk=firstCol_(map,['ma don ghtk','ma don']),cPo=firstCol_(map,['ma po','po']),cOrder=firstCol_(map,['so don hang','od']);
  const q=norm_(query);
  const hit=(ghtk,po,order)=>{
    ghtk=clean_(ghtk);po=clean_(po);order=clean_(order);
    if(!ghtk&&!order)return false; // không phải hồ sơ
    if(expected)return (order?'ORDER:'+ks_key(order):'GHTK:'+ks_key(ghtk))===expected;
    return [ghtk,po,order].some(v=>norm_(v)===q);
  };
  const result=r=>{ const row=sh.getRange(r,1,1,width).getDisplayValues()[0];
    return {row,rec:recordFromRow_(row,map,r)}; };
  if(pnResolveMemo_&&pnResolveMemo_.params===params&&pnResolveMemo_.sheet===sh.getName()){
    const x=result(pnResolveMemo_.rowNumber);
    if(hit(x.rec.maDonGhtk,x.rec.po,x.rec.orderNo))return {rowNumber:pnResolveMemo_.rowNumber,record:x.rec};
  }
  const last=sh.getLastRow(),rows=[];
  if(last>=2){
    const cols=[cGhtk,cPo,cOrder].filter(Boolean),c1=Math.min.apply(null,cols),c2=Math.max.apply(null,cols);
    const block=sh.getRange(2,c1,last-1,c2-c1+1).getDisplayValues();
    const at=(r,c)=>c?r[c-c1]:'';
    block.forEach((r,i)=>{ if(hit(at(r,cGhtk),at(r,cPo),at(r,cOrder)))rows.push(i+2); });
  }
  if(rows.length!==1)throw new Error(rows.length?'Mã khớp nhiều đơn; dùng Số đơn hàng duy nhất.':'Không tìm thấy đơn trong Chứng từ_full.');
  pnResolveMemo_={params,sheet:sh.getName(),rowNumber:rows[0]};
  return {rowNumber:rows[0],record:result(rows[0]).rec};
}

// Đồng bộ cặp Chứng từ. fullRow: dòng API vừa ghi ở full -> bản API thắng và được đẩy sang Chứng từ_FF.
// Đồng bộ đúng dòng fullRow của Chứng từ_full với Chứng từ_FF. prefer=true: bản API vừa ghi thắng.
// Đồng bộ đúng dòng fullRow của Chứng từ_full với Chứng từ_FF. prefer=true: bản API vừa ghi thắng.
// API KHÔNG chờ luồng Sheet: Sheet đang bận (dán/dọn) thì xếp dòng này cho luồng Sheet tự đẩy sang tab chính.
function pnMirrorDocs_(fullRow,prefer) {
  const p=PN_FULL.pairs[2];
  if(!fullRow)return pnV2SyncPair_(p,{});
  if(pnLockDepth_)return pnV2SyncRows_(p,'full',[fullRow],prefer!==false,{api:true});
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(1)||PropertiesService.getScriptProperties().getProperty('PN_CLEANUP_PENDING')){
    try { lock.releaseLock(); } catch(err) {}
    pnQueueEdit_(p.full,fullRow,fullRow);
    return {ok:true,queued:true};
  }
  pnLockDepth_++;
  try { return pnV2SyncRows_(p,'full',[fullRow],prefer!==false,{api:true}); }
  finally { pnLockDepth_--; lock.releaseLock(); }
}

function pnMirrorDocsSafe_(fullRow,prefer) {
  try { pnMirrorDocs_(fullRow,prefer); return ''; }
  catch(err) {
    const message=String(err&&err.message||err);
    PropertiesService.getScriptProperties().setProperty('PN_MIRROR_ERROR',new Date().toISOString()+' '+message.slice(0,4000));
    console.error('PN mirror (save tiếp tục): '+message);
    return message;
  }
}

function pnBusyResponse_(err) {
  if(!/PN_BUSY/.test(String(err&&err.message||err)))throw err;
  return {ok:false,code:'PN_BUSY',retryable:true,message:'Hệ thống PN đang bận, app sẽ gửi lại sau ít giây.'};
}
/* Luồng API (app lưu đơn): độc lập với thao tác trên Sheet, KHÔNG dùng khóa của Sheet.
 * Nhật ký mỗi đơn 3 lần ghi: PENDING (bắt đầu) -> WRITTEN (ảnh + chứng từ đã ghi) -> DONE (kết quả trả app).
 * - Gửi trùng khi lượt trước đang chạy (< 7 phút): trả PN_BUSY, app gửi lại sau và nhận đúng kết quả.
 * - Gửi lại sau lỗi: dùng lại ảnh đã tải, không ghi lại chứng từ đã ghi.
 * - Chỉ ghi vào đúng dòng của đơn; đẩy đúng dòng đó sang Chứng từ_FF. Không chèn/xóa/dồn dòng. */
function apiSave_(params) {
  params=params||{};
  const clientId=pnText_(params.clientId);
  if(!clientId)return fail_('Thiếu clientId; cập nhật app trước khi lưu.');
  const key=pnRequestKey_(clientId),signature=pnRequestSignature_(params);
  let state=null;
  try {
    pnRequireReady_();
    const saved=pnJournalRead_(key);
    if(saved&&saved.data.signature!==signature)return fail_('clientId đã dùng cho nội dung khác.');
    if(saved&&saved.state==='DONE')return saved.data.response;
    const legacy=!saved&&readSavedRequest_(clientId);
    if(legacy)return legacy;
    const now=Date.now();
    if(saved&&(saved.state==='PENDING'||saved.state==='WRITTEN')&&now-Number(saved.data.touchedAtMs||0)<PN_UPLOAD_STALE_MS_)
      throw new Error('PN_BUSY: Đơn này đang được lưu ở lần gửi trước, app sẽ gửi lại sau.');
    state=saved?saved.data:{signature,clientId,startedAt:new Date().toISOString()};
    if(!state.startedAt)state.startedAt=new Date().toISOString();
    state.touchedAtMs=now;
    pnJournalWrite_(key,'PENDING',state);
    pnRequest_=Object.assign({},state,{key,uploads:Object.assign({},state.uploads||{}),legacy:!!saved&&!saved.data.touchedAtMs});
    if(!state.coreResponse)pnPreUpload_(params);
    return pnSaveLocked_(params,key,state);
  } catch(err) {
    // Lỗi (không phải bận): ghi lại ảnh đã tải để lần gửi lại không tải trùng.
    if(state&&!/PN_BUSY/.test(String(err&&err.message||err))){
      try { state.uploads=pnRequest_&&pnRequest_.uploads||state.uploads;state.error=String(err&&err.message||err).slice(0,500);pnJournalWrite_(key,'FAILED',state,false); } catch(e) {}
    }
    return pnBusyResponse_(err);
  } finally { pnRequest_=null;pnResolveMemo_=null; }
}

function pnSaveLocked_(params,key,state) {
  const clientId=pnText_(params.clientId);
  let fullRow=0;
  try { fullRow=pnResolveSave_(sheet_(),null,params).rowNumber; } catch(err) {}
  // Lấy sửa tay ở Chứng từ_FF (nếu luồng Sheet rảnh) trước khi ghi.
  const mirrorWarning=fullRow&&!state.coreResponse?pnMirrorDocsSafe_(fullRow,false):'';
  let response=state.coreResponse;
  if(!response){
    response=pnApiSaveCore_(params);
    if(!response||!response.ok){ pnJournalWrite_(key,'REJECTED',state,false); return response; }
    state.coreResponse=response;state.uploads=pnRequest_.uploads;state.docWritten=!!pnRequest_.docWritten;state.touchedAtMs=Date.now();
    pnJournalWrite_(key,'WRITTEN',state,false);
  }
  const mirrorAfter=pnMirrorDocsSafe_(response.rowNumber||fullRow,true);
  pnMarkDirty_();
  const record=pnResolveSave_(sheet_(),null,params).record;
  const type=norm_(params.returnType||params.loaiHoan||'Chung tu');
  if(['chung tu','docs','san pham + chung tu','product-docs'].includes(type)){
    // Chỉ chạy sang VHFF khi dòng đã tích "Đã tạo sv" (hàm tự kiểm tra, chưa tích thì thoát ngay).
    try { syncOneChungTuKhongDatYCBySourceRow_(record.rowNumber); }
    catch(err){PropertiesService.getScriptProperties().setProperty('PN_VHFF_ERROR','reject: '+String(err.message||err));}
  }
  response=Object.assign({},response,{
    sourceEpoch:PN_FULL.epoch,syncStatus:'stored',vhffStatus:'pending',
    rowNumber:record.rowNumber,record,
    mirrorWarning:mirrorAfter||mirrorWarning||''
  });
  state.response=response;
  pnJournalWrite_(key,'DONE',state,false);
  rememberSavedRequest_(clientId,response);
  return response;
}

// Mã ecom CT: scan ra link (vd https://i.ghtk.vn/...) hoặc nhập tay -> chỉ giữ 10 chữ số cuối. Ít hơn 10 chữ số thì giữ nguyên.
function pnEcomCt_(v) {
  const s=clean_(v),d=s.replace(/[^0-9]/g,'');
  return d.length>=10?d.slice(-10):s;
}

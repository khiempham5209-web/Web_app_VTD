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
  if(sh.getLastRow()<2)return null;
  const found=sh.getRange(2,1,sh.getLastRow()-1,1).createTextFinder(key).matchEntireCell(true).useRegularExpression(false).findAll();
  if(!found.length)return null;
  // Nhật ký chỉ thêm dòng: dòng cuối cùng của khóa là trạng thái mới nhất.
  const last=found.reduce((m,c)=>c.getRow()>m.getRow()?c:m,found[0]);
  const r=sh.getRange(last.getRow(),1,1,4).getValues()[0];
  return {row:last.getRow(),state:r[1],data:JSON.parse(r[2]||'{}')};
}

// Nhật ký chỉ THÊM dòng (appendRow nguyên tử): nhiều đơn ghi cùng lúc không đè nhau, không cần khóa của Sheet.
function pnJournalWrite_(key,state,data) {
  const text=JSON.stringify(data);
  if(text.length>45000)throw new Error('Nhật ký vượt kích thước an toàn; chia nhỏ thao tác.');
  pnJournal_().appendRow([key,state,text,new Date()]);
  SpreadsheetApp.flush();
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
  return Array.from(latest.entries()).filter(([,s])=>s==='PENDING').map(([k])=>k);
}

const PN_UPLOAD_STALE_MS_=7*60*1000; // Apps Script dừng mọi lần chạy sau 6 phút
function pnUploadOnce_(stage,fn) {
  if(!pnRequest_)throw new Error('Upload phải qua apiSave_.');
  const key=pnRequest_.key+':'+stage;
  // Kiểm tra và đánh dấu RUNNING trong cùng một khóa ngắn, để hai lần gửi trùng không cùng tải ảnh.
  let entry=pnJournalRead_(key);
  if(entry&&entry.state!=='DONE'&&!(entry.state==='RUNNING'&&Date.now()-Number(entry.data&&entry.data.startedAtMs||0)<PN_UPLOAD_STALE_MS_))entry=null;
  if(!entry)pnJournalWrite_(key,'RUNNING',{clientId:pnRequest_.clientId,stage,startedAtMs:Date.now()});
  if(entry&&entry.state==='DONE')return entry.data;
  // Lần gửi trước còn đang tải ảnh: báo bận (app giữ đơn ở hàng chờ và gửi lại), không báo lỗi.
  if(entry)throw new Error('PN_BUSY: Ảnh của đơn này đang được tải ở lần gửi trước, app sẽ gửi lại sau.');
  let result;
  try { result=fn(); }
  catch(err) {
    // Lỗi đã biết (Drive báo lỗi): cho gửi lại ngay. Có thể trùng 1 file ảnh trên Drive, nhưng đơn không bị kẹt.
    pnJournalWrite_(key,'FAILED',{clientId:pnRequest_.clientId,stage,error:String(err&&err.message||err)});
    throw err;
  }
  pnJournalWrite_(key,'DONE',result);
  return result;
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
  const key=pnRequest_.key+':document-write';
  let stage=pnJournalRead_(key);
  if(stage&&stage.state==='DONE')return;
  const L=pnV2Layout_(sh,true),width=pnV2ReadWidth_(L),names=Object.keys(L.cols);
  const current=sh.getRange(row,1,1,width).getValues()[0];
  const sig=r=>pnHash_(pnV2Sig_(L,r,names));
  // Dòng vừa bị sửa tay sau lần chuẩn bị trước: chuẩn bị lại trên nội dung hiện tại (chỉ đè các cột chứng từ).
  if(stage&&sig(current)!==stage.data.before&&sig(current)!==stage.data.after)stage=null;
  if(!stage) {
    const next=current.slice();
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
    values.forEach(([aliases,value])=>{ const i=pnV2Col_(L,aliases); if(i!=null)next[i]=value; });
    const data={before:sig(current),after:sig(next),values:pnEncodeRow_(next)};
    pnJournalWrite_(key,'PREPARED',data);stage={state:'PREPARED',data};
  }
  const fingerprint=sig(current);
  if(fingerprint!==stage.data.before&&fingerprint!==stage.data.after)
    throw new Error('Hồ sơ đã thay đổi trong lúc khôi phục thao tác; không ghi đè bản mới.');
  if(fingerprint!==stage.data.after){
    const next=pnDecodeRow_(stage.data.values),up=new Map();
    names.forEach(n=>{const i=L.cols[n];if(pnV2Canon_(next[i])!==pnV2Canon_(current[i]))pnV2Put_(up,row,i,next[i]);});
    pnV2Flush_(sh,up);
  }
  SpreadsheetApp.flush();
  pnJournalWrite_(key,'DONE',stage.data);
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

function pnResolveSave_(sh,col,params) {
  const expected=pnText_(params.syncKey),query=pnText_(params.orderNo||params.maDonGhtk||params.maDon||params.query||params.po);
  if(!expected&&!query)throw new Error('Cần khóa đơn; không lưu chỉ bằng số dòng.');
  const L=pnV2Layout_(sh,true),cfg=pnV2Cfg_(PN_FULL.pairs[2]);
  const map=headerMap_(L.headers);
  const values=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,pnV2ReadWidth_(L)).getDisplayValues():[];
  const records=values.map((r,i)=>pnV2Rec_(L,cfg,r)?{rowNumber:i+2,record:recordFromRow_(r,map,i+2)}:null).filter(Boolean);
  const matches=records.filter(x=>expected?x.record.syncKey===expected:recordMatchesQuery_(x.record,query));
  if(matches.length!==1)throw new Error(matches.length?'Mã khớp nhiều đơn; dùng Số đơn hàng duy nhất.':'Không tìm thấy đơn trong Chứng từ_full.');
  return matches[0];
}

// Đồng bộ cặp Chứng từ. fullRow: dòng API vừa ghi ở full -> bản API thắng và được đẩy sang Chứng từ_FF.
// Đồng bộ đúng dòng fullRow của Chứng từ_full với Chứng từ_FF. prefer=true: bản API vừa ghi thắng.
function pnMirrorDocs_(fullRow,prefer) {
  if(!fullRow)return pnV2SyncPair_(PN_FULL.pairs[2],{});
  return pnV2SyncRows_(PN_FULL.pairs[2],'full',[fullRow],prefer!==false,{api:true});
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
 * - Chống trùng theo clientId qua nhật ký (chỉ thêm dòng).
 * - Chỉ ghi vào đúng dòng của đơn (tìm theo mã, kiểm tra lại dòng ngay trước khi ghi).
 * - Đẩy đúng dòng đó sang Chứng từ_FF. Không bao giờ chèn/xóa/dồn dòng. */
function apiSave_(params) {
  params=params||{};
  const clientId=pnText_(params.clientId);
  if(!clientId)return fail_('Thiếu clientId; cập nhật app trước khi lưu.');
  const key=pnRequestKey_(clientId),signature=pnRequestSignature_(params);
  try {
    pnRequireReady_();
    const saved=pnJournalRead_(key);
    if(saved&&saved.data.signature!==signature)return fail_('clientId đã dùng cho nội dung khác.');
    if(saved&&saved.state==='DONE')return saved.data.response;
    const legacy=!saved&&readSavedRequest_(clientId);
    if(legacy)return legacy;
    const state=saved?saved.data:{signature,clientId,startedAt:new Date().toISOString()};
    if(!saved)pnJournalWrite_(key,'PENDING',state);
    if(!state.coreResponse){
      pnRequest_=Object.assign({},state,{key});
      try { pnPreUpload_(params); } finally { pnRequest_=null; }
    }
    return pnSaveLocked_(params,key,signature);
  } catch(err) { return pnBusyResponse_(err); }
}

function pnSaveLocked_(params,key,signature) {
    pnRequireReady_();
    const clientId=pnText_(params.clientId),saved=pnJournalRead_(key);
    if(saved&&saved.data.signature!==signature)return fail_('clientId đã dùng cho nội dung khác.');
    if(saved&&saved.state==='DONE')return saved.data.response;
    let fullRow=0;
    try { fullRow=pnResolveSave_(sheet_(),null,params).rowNumber; } catch(err) {}
    const mirrorWarning=fullRow?pnMirrorDocsSafe_(fullRow,false):'';
    const state=saved?saved.data:{signature,clientId,startedAt:new Date().toISOString()};
    pnJournalWrite_(key,'PENDING',state);
    pnRequest_=Object.assign({},state,{key});
    try {
      let response=state.coreResponse;
      if(!response){
        response=pnApiSaveCore_(params);
        if(!response||!response.ok){
          pnJournalWrite_(key,'REJECTED',state);
          return response;
        }
        state.coreResponse=response;pnJournalWrite_(key,'PENDING',state);
      }
      const mirrorAfter=pnMirrorDocsSafe_(state.coreResponse&&state.coreResponse.rowNumber||response.rowNumber,true);
      pnMarkDirty_();
      const record=pnResolveSave_(sheet_(),headerMap_(headers_(sheet_())),params).record;
      const type=norm_(params.returnType||params.loaiHoan||'Chung tu');
      if(['chung tu','docs','san pham + chung tu','product-docs'].includes(type)){
        try { syncOneChungTuKhongDatYCBySourceRow_(record.rowNumber); }
        catch(err){PropertiesService.getScriptProperties().setProperty('PN_VHFF_ERROR','reject: '+String(err.message||err));}
      }
      // VHFF runs in the worker; app success means the request is durably stored, not VHFF acknowledgement.
      response=Object.assign({},response,{
        sourceEpoch:PN_FULL.epoch,syncStatus:'stored',vhffStatus:'pending',
        rowNumber:record.rowNumber,record,
        // Đơn đã lưu vào full; tab chính sẽ được trigger đối soát bù khi admin xử lý xong lỗi này.
        mirrorWarning:mirrorAfter||mirrorWarning||''
      });
      state.response=response;
      pnJournalWrite_(key,'DONE',state);
      rememberSavedRequest_(clientId,response);
      return response;
    } finally { pnRequest_=null; }
}
// Mã ecom CT: scan ra link (vd https://i.ghtk.vn/...) hoặc nhập tay -> chỉ giữ 10 chữ số cuối. Ít hơn 10 chữ số thì giữ nguyên.
function pnEcomCt_(v) {
  const s=clean_(v),d=s.replace(/D/g,'');
  return d.length>=10?d.slice(-10):s;
}

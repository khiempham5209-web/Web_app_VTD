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
  if(found.length>1)throw new Error('Nhật ký trùng khóa; cần đối soát: '+key);
  if(!found.length)return null;
  const r=sh.getRange(found[0].getRow(),1,1,4).getValues()[0];
  return {row:found[0].getRow(),state:r[1],data:JSON.parse(r[2]||'{}')};
}
function pnJournalWrite_(key,state,data) {
  const text=JSON.stringify(data);
  if(text.length>45000)throw new Error('Nhật ký vượt kích thước an toàn; chia nhỏ thao tác.');
  // Ghi nhật ký luôn trong khóa: upload ảnh chạy ngoài khóa chung, hai request không được cấp trùng dòng.
  pnWithLock_(()=>{
    const sh=pnJournal_(),old=pnJournalRead_(key),row=old?old.row:Math.max(2,sh.getLastRow()+1);
    pnSize_(sh,row,4);
    sh.getRange(row,1,1,4).setValues([[key,state,text,new Date()]]);
    SpreadsheetApp.flush();
  });
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
  return pnRows_(pnJournal_(),4).filter(r=>/^request:[a-f0-9]+$/.test(String(r[0]))&&r[1]==='PENDING').map(r=>r[0]);
}
// Một lần tải ảnh treo quá lâu chắc chắn đã chết (Apps Script dừng sau 6 phút): cho phép tải lại.
const PN_UPLOAD_STALE_MS_=10*60*1000;
function pnUploadOnce_(stage,fn) {
  if(!pnRequest_)throw new Error('Upload phải qua apiSave_.');
  const key=pnRequest_.key+':'+stage;
  // Kiểm tra và đánh dấu RUNNING trong cùng một khóa ngắn, để hai lần gửi trùng không cùng tải ảnh.
  const entry=pnWithLock_(()=>{
    const e=pnJournalRead_(key);
    if(e&&e.state==='DONE')return e;
    if(e&&e.state==='RUNNING'&&Date.now()-Number(e.data&&e.data.startedAtMs||0)<PN_UPLOAD_STALE_MS_)return e;
    pnJournalWrite_(key,'RUNNING',{clientId:pnRequest_.clientId,stage,startedAtMs:Date.now()});
    return null;
  });
  if(entry&&entry.state==='DONE')return entry.data;
  if(entry)throw new Error('Ảnh của đơn này đang được tải ở một lần gửi khác. App sẽ thử lại sau.');
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
function pnWriteDocumentOnce_(sh,row,col,params,upload,timeText,user) {
  const key=pnRequest_.key+':document-write',p=PN_FULL.pairs[2];
  let stage=pnJournalRead_(key);
  if(stage&&stage.state==='DONE')return;
  const current=sh.getRange(row,1,1,p.width).getValues()[0];
  if(!stage) {
    const next=current.slice();
    const values=[
      [['xac thuc hoa don'],clean_(params.xacThuc||params.xacThucHoaDon)],
      [['ghi chu chung tu'],String(params.note||params.ghiChu||'')],
      [['trang thai'],clean_(params.status||params.trangThai)],
      [['thoi gian'],timeText],[['user thao tac'],user]
    ];
    if(upload.linkAnh)values.push([['link anh'],upload.linkAnh]);
    // Mã ecom CT (cột L): chỉ ghi khi app gửi giá trị, để app cũ không xóa mã đã nhập tay trên Sheet.
    const maEcomCt=clean_(params.maEcomCt);
    if(maEcomCt)values.push([['ma ecom ct'],maEcomCt]);
    // Ghi theo TÊN cột (header), không theo số cột. Cột nằm ngoài vùng nghiệp vụ A–Q mà API đồng bộ
    // chính/full thì bỏ qua, không làm hỏng cả lần lưu.
    values.forEach(([aliases,value])=>{
      const c=firstCol_(col,aliases);
      if(c&&c<=p.width)next[c-1]=value;
      else if(c)console.warn('Cột '+aliases[0]+' nằm ngoài vùng A–Q, bỏ qua.');
    });
    const data={before:pnFingerprint_(current,p),after:pnFingerprint_(next,p),values:pnEncodeRow_(next)};
    pnJournalWrite_(key,'PREPARED',data);stage={state:'PREPARED',data};
  }
  const fingerprint=pnFingerprint_(current,p);
  if(fingerprint!==stage.data.before&&fingerprint!==stage.data.after)
    throw new Error('Hồ sơ đã thay đổi trong lúc khôi phục thao tác; không ghi đè bản mới.');
  if(fingerprint!==stage.data.after)sh.getRange(row,1,1,p.width).setValues([pnDecodeRow_(stage.data.values)]);
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
  const c=pnProductIdColumn_(sh),existing=new Map();
  pnRows_(sh,width).forEach((r,i)=>{
    if(r[c-1]){if(existing.has(r[c-1]))throw new Error('Trùng clientItemId ở Hoàn sản phẩm');existing.set(r[c-1],i+2);}
  });
  const missing=rows.filter(r=>!existing.has(r[c-1]));
  if(missing.length){
    const start=Math.max(2,sh.getLastRow()+1);pnSize_(sh,start+missing.length-1,width);
    sh.getRange(start,1,missing.length,width).setValues(missing);
  }
}
function pnResolveSave_(sh,col,params) {
  const expected=pnText_(params.syncKey),query=pnText_(params.orderNo||params.maDonGhtk||params.maDon||params.query||params.po);
  if(!expected&&!query)throw new Error('Cần khóa đơn; không lưu chỉ bằng số dòng.');
  const values=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,17).getDisplayValues():[];
  const records=values.map((r,i)=>({rowNumber:i+2,record:recordFromRow_(r,col,i+2)}));
  const matches=records.filter(x=>expected?x.record.syncKey===expected:recordMatchesQuery_(x.record,query));
  if(matches.length!==1)throw new Error(matches.length?'Mã khớp nhiều đơn; dùng Số đơn hàng duy nhất.':'Không tìm thấy đơn trong Chứng từ_full.');
  return matches[0];
}
function pnMirrorDocs_() {
  pnReconcilePair_(PN_FULL.pairs[2]);
  pnMaterializeDocs_();
  pnEnsureFullFilter_(PN_FULL.pairs[2]);
}
// Đối soát chính/full khi lưu đơn: lỗi ở dòng KHÁC (xung đột, full giảm dòng...) không được chặn
// việc lưu của mọi máy. Ghi lỗi lại cho admin; trigger 5 phút vẫn báo và thử lại.
function pnMirrorDocsSafe_() {
  try { pnMirrorDocs_(); return ''; }
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
function apiSave_(params) {
  params=params||{};
  const clientId=pnText_(params.clientId);
  if(!clientId)return fail_('Thiếu clientId; cập nhật app trước khi lưu.');
  const key=pnRequestKey_(clientId),signature=pnRequestSignature_(params);
  try {
    // Bước 1 (khóa ngắn): chống trùng, đăng ký yêu cầu.
    const begin=pnWithLock_(()=>{
      pnRequireReady_();
      const saved=pnJournalRead_(key);
      if(saved&&saved.data.signature!==signature)return {response:fail_('clientId đã dùng cho nội dung khác.')};
      if(saved&&saved.state==='DONE')return {response:saved.data.response};
      // Preserve successes recorded by the pre-migration API.
      const legacy=!saved&&readSavedRequest_(clientId);
      if(legacy)return {response:legacy};
      const state=saved?saved.data:{signature,clientId,startedAt:new Date().toISOString()};
      if(!saved)pnJournalWrite_(key,'PENDING',state);
      return {state};
    });
    if(begin.response)return begin.response;
    // Bước 2 (ngoài khóa chung): tải ảnh lên Drive. Phần chậm nhất không còn chặn đơn của máy khác.
    if(!begin.state.coreResponse){
      pnRequest_=Object.assign({},begin.state,{key});
      try { pnPreUpload_(params); } finally { pnRequest_=null; }
    }
    // Bước 3 (khóa): ghi Sheet; ảnh đã có trong nhật ký nên core không tải lại.
    return pnWithLock_(()=>pnSaveLocked_(params,key,signature));
  } catch(err) { return pnBusyResponse_(err); }
}
function pnSaveLocked_(params,key,signature) {
    pnRequireReady_();
    const clientId=pnText_(params.clientId),saved=pnJournalRead_(key);
    if(saved&&saved.data.signature!==signature)return fail_('clientId đã dùng cho nội dung khác.');
    if(saved&&saved.state==='DONE')return saved.data.response;
    const mirrorWarning=pnMirrorDocsSafe_();
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
      const mirrorAfter=pnMirrorDocsSafe_();
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

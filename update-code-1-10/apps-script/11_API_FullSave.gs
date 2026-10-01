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
  const sh=pnJournal_(),old=pnJournalRead_(key),row=old?old.row:Math.max(2,sh.getLastRow()+1);
  const text=JSON.stringify(data);
  if(text.length>45000)throw new Error('Nhật ký vượt kích thước an toàn; chia nhỏ thao tác.');
  pnSize_(sh,row,4);
  sh.getRange(row,1,1,4).setValues([[key,state,text,new Date()]]);
  SpreadsheetApp.flush();
}
function pnRequestKey_(id) { return 'request:'+pnHash_(pnText_(id)); }
function pnRequestSignature_(params) {
  const data=Object.assign({},params);
  ['sessionToken','_silent','rowNumber'].forEach(k=>delete data[k]);
  return pnHash_(data);
}
function pnRequestNow_() { return pnRequest_?new Date(pnRequest_.startedAt):new Date(); }
function pnPendingRequests_() {
  return pnRows_(pnJournal_(),4).filter(r=>/^request:[a-f0-9]+$/.test(String(r[0]))&&r[1]==='PENDING').map(r=>r[0]);
}
function pnUploadOnce_(stage,fn) {
  if(!pnRequest_)throw new Error('Upload phải qua apiSave_.');
  const key=pnRequest_.key+':'+stage,entry=pnJournalRead_(key);
  if(entry&&entry.state==='DONE')return entry.data;
  if(entry&&entry.state==='RUNNING')
    throw new Error('Lần tải ảnh trước chưa xác định kết quả. Giữ nguyên yêu cầu '+pnRequest_.clientId+'; kiểm tra nhật ký trước khi thử lại.');
  pnJournalWrite_(key,'RUNNING',{clientId:pnRequest_.clientId,stage});
  const result=fn();
  pnJournalWrite_(key,'DONE',result);
  return result;
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
    values.forEach(([aliases,value])=>setArrayByAliases_(next,col,aliases,value));
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
function apiSave_(params) {
  return pnWithLock_(()=>{
    pnRequireReady_();
    params=params||{};
    const clientId=pnText_(params.clientId);
    if(!clientId)return fail_('Thiếu clientId; cập nhật app trước khi lưu.');
    const key=pnRequestKey_(clientId),signature=pnRequestSignature_(params),saved=pnJournalRead_(key);
    if(saved&&saved.data.signature!==signature)return fail_('clientId đã dùng cho nội dung khác.');
    if(saved&&saved.state==='DONE')return saved.data.response;
    // Preserve successes recorded by the pre-migration API.
    const legacy=!saved&&readSavedRequest_(clientId);
    if(legacy)return legacy;
    pnMirrorDocs_();
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
      pnMirrorDocs_();
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
        rowNumber:record.rowNumber,record
      });
      state.response=response;
      pnJournalWrite_(key,'DONE',state);
      rememberSavedRequest_(clientId,response);
      return response;
    } finally { pnRequest_=null; }
  });
}

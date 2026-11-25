/* Equivalent of the four observed ARRAYFORMULAs. TT Nhập!L1 stays authoritative.
 * DS BC and TT Nhập formulas themselves are not removed.
 */
function pnRoundUp_(v) { return Math.sign(v)*Math.ceil(Math.abs(v)); }
// Hệ số tính Check: ưu tiên TT Nhập!L1; L1 trống/không phải số thì dùng L3 (2 ô giữ cùng công thức).
function pnCheckFactor_() {
  const sh=pnSheet_('TT Nhập');
  for(const a of ['L1','L3']){const v=sh.getRange(a).getValue();if(typeof v==='number'&&Number.isFinite(v))return v;}
  throw new Error('TT Nhập!L1 và L3 đều không phải số hợp lệ.');
}
function pnLookupKey_(v) { return String(v == null ? '' : v).toLowerCase(); }
function pnFormulaValues_(p,rows,factor,bcRows) {
  if(p.main==='Booking') {
    if(typeof factor!=='number'||!Number.isFinite(factor))throw new Error('TT Nhập!L1 không phải số hợp lệ.');
    return rows.map(r=>[r[0]===''?'':pnRoundUp_(Number(r[9]||0)*factor)]);
  }
  const byDepot=new Map();
  bcRows.forEach(r=>{const key=pnLookupKey_(r[6]);if(!byDepot.has(key))byDepot.set(key,r);});
  return rows.map(r=>{
    if(r[0]==='')return ['','',''];
    const match=byDepot.get(pnLookupKey_(r[6])),date=pnIsoDate_(r[1]);
    return [match?match[4]:'#N/A',match?match[14]:'#N/A',date?date.slice(8)+'/'+date.slice(5,7)+'/'+date.slice(0,4):'#VALUE!'];
  });
}
function pnFormulaPlans_(includeFull) {
  SpreadsheetApp.flush();
  const factor=pnCheckFactor_();
  const bc=pnRows_(pnSheet_('DS BC'),15);
  const plans=[];
  PN_FULL.pairs.slice(0,2).forEach(p=>{
    (includeFull?[p.main,p.full]:[p.main]).forEach(name=>{
      const sh=pnSheet_(name),rows=pnRows_(sh,p.width);
      plans.push({p,sh,rows,col:16,values:pnFormulaValues_(p,rows,factor,bc)});
    });
  });
  return plans;
}
function pnExpectedFormula_(sheet,col) {
  const formulas={
    'Booking:16': '=arrayformula(if(A2:A="";"";ROUNDUP(J2:J*\'TT Nhập\'!$L$3;0)))',
    'File đơn:16': '=arrayformula(if(A2:A="";"";xlookup(G2:G;\'DS BC\'!$G$2:$G;\'DS BC\'!$E2:$E)))',
    'File đơn:17': '=arrayformula(if(A2:A="";"";xlookup(G2:G;\'DS BC\'!$G$2:$G;\'DS BC\'!$O2:$O)))',
    'File đơn:18': '=arrayformula(if(A2:A="";"";text(B2:B;"dd/mm/yyyy")))'
  };
  return formulas[sheet+':'+col];
}
function pnConvertFormulas_() {
  const plans=pnFormulaPlans_();
  const normalized=x=>String(x).replace(/\s/g,'').toLowerCase();
  // Preflight all columns before removing any anchor.
  plans.forEach(({p,sh,rows,values,col})=>{
    const count=p.main==='Booking'?1:3;
    for(let c=0;c<count;c++){
      const formula=sh.getRange(2,col+c).getFormula();
      if(formula && normalized(formula)!==normalized(pnExpectedFormula_(p.main,col+c)))
        throw new Error('Công thức khác bản đã kiểm tra: '+p.main+' cột '+(col+c));
      if(formula)rows.forEach((r,i)=>{
        if(String(r[col+c-1])!==String(values[i][c]))
          throw new Error('Code/công thức chưa khớp: '+p.main+' dòng '+(i+2)+' cột '+(col+c));
      });
    }
  });
  plans.forEach(({p,sh,values,col})=>{
    const count=p.main==='Booking'?1:3;
    // Snapshot original formula for rollback; leave unrelated formulas/pivots alone.
    for(let c=0;c<count;c++){
      const cell=sh.getRange(2,col+c),formula=cell.getFormula();
      if(formula){
        PropertiesService.getScriptProperties().setProperty('PN_FORMULA_'+p.main+'_'+(col+c),formula);
        cell.clearContent();
      }
    }
    SpreadsheetApp.flush();
    if(values.length)sh.getRange(2,col,values.length,count).setValues(values);
  });
  PropertiesService.getScriptProperties().setProperty('PN_FORMULAS_CONVERTED','1');
}
function pnComputeColumns_() {
  if(PropertiesService.getScriptProperties().getProperty('PN_FORMULAS_CONVERTED')!=='1')
    throw new Error('Chưa chuyển công thức; chạy migration trước.');
  for(const plan of pnFormulaPlans_(true)){
    const count=plan.p.main==='Booking'?1:3,updates=[];
    plan.values.forEach((v,i)=>{
      if(v.some((x,c)=>String(x)!==String(plan.rows[i][plan.col+c-1])))updates.push({row:i+2,values:v});
    });
    pnWriteBlocks_(plan.sh,updates,plan.col,count);
  }
}
// Bộ đồng bộ v2 tự cấp ID cho dòng mới ở full trong cùng lượt đồng bộ; giữ tên hàm cho các script cũ gọi.
function pnStampFullIds_(p) { return 0; }

/* Sửa tay trên Sheet: KHÔNG khóa chung. Mỗi lần sửa chỉ xử lý đúng các dòng vừa sửa (và dòng script đổi theo),
 * chạy song song với người khác và với app. Chỉ bước thêm dữ liệu MỚI vào cuối một tab đi qua khóa ngắn.
 * Đang dọn tháng (5h sáng) hoặc khóa thêm dòng bận quá lâu: lưu đúng dòng vừa sửa, tự chạy bù sau ~20 giây. */
function pnHandleEdit(e) {
  if(!e||!e.range)return;
  const sh=e.range.getSheet(),name=sh.getName(),p=pnPair_(name);
  const watched=['Booking','Booking_full','File đơn','File đơn_full','Chứng từ_FF','Chứng từ_full','DS BC','TT Nhập','DS SKU','Hàng lỗi','Bàn giao chứng từ','Hoàn sản phẩm','Sự vụ'];
  if(watched.indexOf(name)<0)return;
  // Capture the original manual edit before the two tabs are synchronized or queued.
  if ((name === 'Chứng từ_FF' || name === 'Chứng từ_full') &&
      typeof pnAppLogSheetEvents_ === 'function') {
    pnAppLogSafe_(() => pnLogOriginalSvEdit_(e, sh));
  }
  const r1=Math.max(2,e.range.getRow()),r2=Math.max(r1,e.range.getLastRow());
  const later=()=>{
    if(p)pnQueueEdit_(name,r1,r2);
    else pnScheduleJob_(pnEditJobs_(name));
    return {ok:false,busy:true,queued:true};
  };
  if(pnCleaning_())return later();
  try {
    return pnRun_(()=>{
      pnRequireReady_();
      pnV2Start_=Date.now();
      if(name==='Booking'){const col=bookingSyncHeaderMap_(pnHeaders_(sh));bookingSyncFillMissingBookingDates_(sh,col,r1,r2);}
      if(name==='Chứng từ_FF'||name==='Chứng từ_full')bookingSyncRepairControlColumnsForEditedTargetRange_(e.range);
      if(name==='TT Nhập'||name==='DS SKU')onEditSyncTTNhapDSSKU(e);
      if(name==='Hàng lỗi'||name==='DS SKU')onEditSyncHangLoiGHTK(e);
      // Như logic cũ: sửa Hoàn sản phẩm -> cập nhật Bàn giao SP_Hàng hoàn; sửa Sự vụ -> cập nhật Sự vụ bánh xẹp (file đích).
      if(name==='Hoàn sản phẩm'&&r2>=2)pnRunTarget_('products',()=>syncBanGiaoSanPham());
      // Sửa Mã đơn GHTK / Mã ecom CT / Link ảnh ở Chứng từ -> bổ sung cột Chứng từ thu hồi (Bàn giao SP_Hàng hoàn).
      if((name==='Chứng từ_FF'||name==='Chứng từ_full')&&pnEditTouchesRecall_(sh,e.range))pnRunTarget_('products',()=>syncBanGiaoSanPham());
      if(name==='Sự vụ'&&r2>=2)pnRunTarget_('incidents',()=>syncAllSuVuBanhXep_());
      let result=null;
      if(p){
        // Booking / File đơn / Chứng từ (chính hoặc full): chỉ các dòng vừa sửa/dán và dòng script đổi theo.
        result=pnV2HandleRows_(name,r1,r2,{c1:e.range.getColumn(),c2:e.range.getLastColumn()});
        // Sửa khi đang lọc: Google chỉ báo vùng liền đầu tiên -> bắt bù các dòng phía dưới bị lệch ở đúng các cột vừa sửa.
        try { const caught=pnV2CatchFilteredEdit_(name,r1,r2,e.range.getColumn(),e.range.getLastColumn()); if(result&&caught)result.filtered=caught; }
        catch(err) { console.error('Filtered edit: '+String(err&&err.message||err)); }
      } else if(name==='DS BC'||name==='TT Nhập'){
        // Chỉ khi hệ số Check (TT Nhập!L1) / bảng kho (DS BC) THỰC SỰ đổi: tính lại đúng cột đó, chỉ ghi ô khác.
        result=pnV2RecalcSettings_(name);
      }
      if(name==='Bàn giao chứng từ'&&e.range.getA1Notation()==='A1'){
        // Lỗi dựng biên bản không được làm hỏng các việc đồng bộ khác trong cùng lần sửa.
        try { refreshBanGiaoChungTuDropdown();syncBanGiaoChungTuBySelectedDate(); }
        catch(err) { console.error('Biên bản bàn giao: '+String(err&&err.message||err)); }
      }
      pnAfterActivity_();
      return result;
    });
  } catch(err) {
    // Khóa thêm dòng bận quá 30 giây (hiếm): không bỏ qua, lưu đúng dòng đã sửa và tự chạy bù.
    if(/PN_BUSY/.test(String(err&&err.message||err)))return later();
    throw err;
  } finally { pnV2Start_=0; }
}
function pnLogOriginalSvEdit_(e, sheet) {
  const edit = e.range;
  if (edit.getLastRow() < 2) return;
  const col = rejectSyncHeaderMap_(rejectSyncHeaders_(sheet));
  const checkbox = rejectSyncFirstCol_(col, ['da tao sv', 'đã tạo sv']);
  if (!checkbox || edit.getColumn() > checkbox || edit.getLastColumn() < checkbox) return;

  const single = edit.getNumRows() === 1 && edit.getNumColumns() === 1;
  if (single && (!rejectSyncIsTrue_(e.value) || rejectSyncIsTrue_(e.oldValue))) return;

  const first = Math.max(2, edit.getRow());
  const count = edit.getLastRow() - first + 1;
  const rows = sheet.getRange(first, 1, count, sheet.getLastColumn()).getDisplayValues();
  const sourceTab = sheet.getName();
  const at = new Date().toISOString();
  let editor = '';
  try { editor = e.user ? String(e.user.getEmail() || '').trim() : ''; }
  catch (err) { console.warn('PN SV editor unavailable'); }
  const column = sheet.getRange(1, checkbox).getA1Notation().replace(/1$/, '');

  rows.forEach((row, index) => {
    const source = rejectSyncRecordFromRowValues_(row, col);
    if (!source.maDonGhtk || (!single && !rejectSyncIsTrue_(source.daTaoSv))) return;
    pnAppLogSheetEvents_(source, [{
      at, origin: 'SHEET',
      state: single ? 'SV_CHECKED' : 'SV_TRUE_PASTED',
      editor: editor || 'Không xác định',
      sourceTab, sourceCell: column + (first + index)
    }]);
  });
}

// Tab không thuộc cặp chính/full: việc VHFF cần chạy bù nếu lần sửa bị hoãn.
function pnEditJobs_(name) {
  if(name==='Hoàn sản phẩm')return ['products'];
  if(name==='Sự vụ')return ['incidents'];
  if(name==='Hàng lỗi'||name==='DS SKU')return ['defects'];
  return [];
}
/* TT Nhập / DS BC: chỉ khi giá trị dùng để tính THỰC SỰ đổi (hệ số Check; Kho đích -> Quản lý/Khu vực của DS BC).
 * Tính lại đúng cột đó (Booking!Check hoặc File đơn!Quản lý/Khu vực/Ngay) ở cả tab chính và full, chỉ ghi ô khác.
 * File đơn đổi Khu vực -> cập nhật Khu vực của đúng các đơn đó ở Booking. Không đồng bộ lại toàn bộ. */
function pnV2RecalcSettings_(name) {
  const props=PropertiesService.getScriptProperties();
  const only=name==='TT Nhập'?'Booking':'File đơn',key=name==='TT Nhập'?'PN_SET_CHECK':'PN_SET_DSBC';
  const sig=name==='TT Nhập'?'F:'+pnCheckFactor_():pnHash_(pnRows_(pnSheet_('DS BC'),15).map(r=>[pnLookupKey_(r[6]),r[4],r[14]]));
  if(props.getProperty(key)===sig)return {ok:true,skipped:'unchanged'};
  const out={rows:{}};
  const writes=pnV2ComputeColumns_(only,out);
  const res={ok:true,writes,rows:{}};
  Object.keys(out.rows).forEach(k=>{res.rows[k]=out.rows[k].length;});
  if(only==='File đơn'){
    const orders=new Set();
    ['File đơn','File đơn_full'].forEach(n=>{
      const rs=(out.rows[n]||[]).slice().sort((a,b)=>a-b);
      pnV2Blocks_(rs).forEach(([a,b])=>pnV2FileDonOrders_(n,[a,b]).forEach(o=>orders.add(o)));
    });
    res.area=[];
    ['Booking','Booking_full'].forEach(bn=>{
      const changed=pnV2AreaForOrders_(bn,orders);
      if(changed.length)res.area.push(pnV2SyncRows_(PN_FULL.pairs[0],bn==='Booking'?'main':'full',changed,true));
    });
  }
  props.setProperty(key,sig);
  return res;
}

function pnSyncBookingFullRows_(start,end) {
  const previous=BOOKING_SYNC_CONFIG.bookingSheetName;
  try {
    BOOKING_SYNC_CONFIG.bookingSheetName='Booking_full';
    syncBookingRowsToChungTuFF_(start,end);
    bookingSyncUpdateTargetDatesFromBookingRows_(start,end);
  } finally { BOOKING_SYNC_CONFIG.bookingSheetName=previous; }
}
function pnSyncAreaAll_() {
  const previous=KHUVUC_BOOKING_SYNC_CONFIG.bookingSheetName;
  try {
    for(const name of ['Booking','Booking_full']){
      KHUVUC_BOOKING_SYNC_CONFIG.bookingSheetName=name;
      syncAllKhuVucBookingToChungTuFF_();
    }
  } finally { KHUVUC_BOOKING_SYNC_CONFIG.bookingSheetName=previous; }
}
// Chạy tay khi cần đối soát lại toàn bộ (ví dụ sau khi dán dữ liệu bằng script khác). KHÔNG cài trigger theo giờ.
function pnScheduledReconcile() {
  try {
    return pnWithLock_(()=>{
      pnRequireReady_();
      pnV2Start_=Date.now();
      return pnV2SyncAll_({});
    });
  } catch(err) {
    if(/PN_BUSY/.test(String(err&&err.message||err)))return {ok:false,busy:true};
    throw err;
  } finally { pnV2Start_=0; }
}
/* Sau mỗi lần sửa/mở Sheet: chỉ việc nhẹ. Không quét lại toàn bộ, không chạy vòng VHFF, không dọn tháng
 * (dọn tháng chạy 5h sáng: pnDon5h).
 * - Chứng từ ở full chưa từng được đồng bộ (lượt trước bị dừng giữa chừng): đẩy sang tab chính (tối đa 2 phút/lần).
 * - Còn lần sửa trong hàng chờ: hẹn lượt chạy bù riêng (không làm chen trong lần sửa của người khác). */
function pnAfterActivity_() {
  const props=PropertiesService.getScriptProperties();
  if(!pnV2OverBudget_()&&Date.now()-Number(props.getProperty('PN_HEAL_LAST')||0)>120000){
    props.setProperty('PN_HEAL_LAST',String(Date.now()));
    try { pnV2HealOrphans_(); } catch(err) { console.error('Heal: '+String(err&&err.message||err)); }
  }
  if(Object.keys(props.getProperties()).some(k=>k.indexOf('PN_EDITQ_')===0))pnScheduleJob_([]);
}
// Chạy tay: đồng bộ VHFF hết các phần còn lại ngay.
function pnRunVhffNow() {
  return pnWithLock_(()=>{
    pnRequireReady_();pnV2Start_=Date.now();
    try {
      const steps=[];
      for(let i=0;i<PN_VHFF_JOBS_.length+1&&!pnV2OverBudget_();i++){
        const r=pnV2VhffStep_();steps.push(r);
        if(r.skipped||r.done)break;
      }
      return {ok:true,steps};
    } finally { pnV2Start_=0; }
  });
}

// Giữ tên hàm cho tương thích; không còn được cài trigger theo giờ.
function pnVhffWorker() { return pnRunVhffNow(); }

function pnSyncVHFF_() {
  pnRequireReady_();
  const props=PropertiesService.getScriptProperties(),sourceBefore=pnSourceRevision_();
  if(props.getProperty('PN_VHFF_SOURCE_ACK')!==sourceBefore&&!props.getProperty('PN_VHFF_DIRTY'))pnMarkDirty_();
  const revision=props.getProperty('PN_VHFF_DIRTY');
  if(!revision)return {ok:true,skipped:true};
  // Fail the whole acknowledgement if any required destination fails.
  // Existing field mappings/manual Ecom preservation stay in the nine business scripts.
  const audit=pnAudit_();
  if(!audit.ok)throw new Error('Full chưa khớp; hoãn VHFF.');
  const jobs=[
    ['reject',()=>syncAllChungTuKhongDatYC()],
    ['documents',()=>syncBanGiaoChungTu()],
    ['products',()=>syncBanGiaoSanPham()],
    ['incidents',()=>syncAllSuVuBanhXep_()],
    ['defects',()=>syncAllHangLoiToGHTK()]
  ],results=[];
  for(const [name,job] of jobs) {
    try {
      const result=job();
      if(result&&result.ok===false)throw new Error(JSON.stringify(result));
      results.push({name,ok:true});
    } catch(err) {
      props.setProperty('PN_VHFF_ERROR',name+': '+String(err.message||err));
      throw err;
    }
  }
  refreshBanGiaoChungTuDropdown();
  syncBanGiaoChungTuBySelectedDate();
  if(pnSourceRevision_()!==sourceBefore){pnMarkDirty_();throw new Error('Nguồn đổi trong lúc sync VHFF; chạy lại trước khi dọn.');}
  if(props.getProperty('PN_VHFF_DIRTY')===revision){
    props.deleteProperty('PN_VHFF_DIRTY');props.deleteProperty('PN_VHFF_ERROR');
    props.setProperty('PN_VHFF_ACK',revision);
    props.setProperty('PN_VHFF_SOURCE_ACK',sourceBefore);
  }
  return {ok:true,results};
}
function pnSourceRevision_() {
  const specs=PN_FULL.pairs.map(p=>[p.full,p.width]).concat([['Hoàn sản phẩm',0],['Sự vụ',0],['Hàng lỗi',0]]);
  return pnHash_(specs.map(([name,width])=>{const sh=pnSheet_(name);return [name,pnRows_(sh,width||sh.getLastColumn())];}));
}
function pnSyncVHFF() { return pnWithLock_(()=>pnSyncVHFF_()); }
// Chỉ còn trigger theo SỰ KIỆN: sửa Sheet và mở Sheet. Không có trigger chạy theo giờ.
function pnInstallTriggers() {
  pnRequireReady_();
  const old=new Set([
    'onEditSyncBookingToChungTuFF','onChangeSyncBookingToChungTuFF','onEditSyncChungTuKhongDatYC',
    'onEditSyncBanGiaoChungTu','onOpenSyncBanGiaoChungTu','onEditSyncTTNhapDSSKU',
    'handleBanGiaoEdit','onEditSyncHangLoiGHTK','onEditSyncSuVuBanhXep','onEditSyncSuvuBanhXep',
    'onEditSyncKhuVucBookingToChungTuFF','onChangeSyncKhuVucBookingToChungTuFF',
    'pnHandleEdit','pnScheduledReconcile','pnCleanupDaily','pnOpen','pnVhffWorker'
  ]);
  ScriptApp.getProjectTriggers().filter(t=>old.has(t.getHandlerFunction())).forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('pnHandleEdit').forSpreadsheet(PN_FULL.spreadsheetId).onEdit().create();
  ScriptApp.newTrigger('pnOpen').forSpreadsheet(PN_FULL.spreadsheetId).onOpen().create();
  // 5h sáng (lúc không ai làm): dọn tháng cũ + dọn dòng trống cuối tab. Thay cho trigger pnDonDongTrong cũ.
  ScriptApp.getProjectTriggers().filter(t=>['pnDonDongTrong','pnDon5h'].indexOf(t.getHandlerFunction())>=0).forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('pnDon5h').timeBased().atHour(5).everyDays(1).inTimezone('Asia/Ho_Chi_Minh').create();
  const props=PropertiesService.getScriptProperties();
  props.setProperty('PN_FULL_CLEANUP','enabled');
  ['PN_VHFF_CYCLE','PN_DAILY_DONE','PN_SYNC_REQUESTED'].forEach(k=>props.deleteProperty(k));
  return 'Đã cài trigger: khi sửa Sheet, khi mở Sheet, và 5h sáng (dọn tháng cũ + dòng trống). Không khóa chung.';
}

function pnOpen() {
  if(pnCleaning_())return {ok:false,busy:true};
  try {
    return pnRun_(()=>{
      pnRequireReady_();
      pnV2Start_=Date.now();
      const r=onOpenSyncBanGiaoChungTu();
      pnAfterActivity_();
      return r;
    });
  } catch(err) {
    if(/PN_BUSY/.test(String(err&&err.message||err)))return {ok:false,busy:true};
    throw err;
  } finally { pnV2Start_=0; }
}

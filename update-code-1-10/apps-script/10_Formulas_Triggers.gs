/* Equivalent of the four observed ARRAYFORMULAs. TT Nhập!L3 stays authoritative.
 * DS BC and TT Nhập formulas themselves are not removed.
 */
function pnRoundUp_(v) { return Math.sign(v)*Math.ceil(Math.abs(v)); }
function pnLookupKey_(v) { return String(v == null ? '' : v).toLowerCase(); }
function pnFormulaValues_(p,rows,factor,bcRows) {
  if(p.main==='Booking') {
    if(typeof factor!=='number'||!Number.isFinite(factor))throw new Error('TT Nhập!L3 không phải số hợp lệ.');
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
  const factor=pnSheet_('TT Nhập').getRange('L3').getValue();
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

// Sửa tay trên Sheet: chỉ xử lý các dòng vừa sửa (bên vừa sửa thắng nếu hai bên cùng đổi), rồi đẩy tiếp
// sang các tab phụ thuộc. Đọc mỗi tab một lần, chỉ ghi ô khác. Khóa bận thì để lượt định kỳ làm, không báo lỗi.
function pnHandleEdit(e) {
  if(!e||!e.range)return;
  const sh=e.range.getSheet(),name=sh.getName(),p=pnPair_(name);
  const watched=['Booking','Booking_full','File đơn','File đơn_full','Chứng từ_FF','Chứng từ_full','DS BC','TT Nhập','DS SKU','Hàng lỗi','Bàn giao chứng từ'];
  if(watched.indexOf(name)<0)return;
  try {
    return pnWithLock_(()=>{
      pnRequireReady_();
      pnV2Start_=Date.now();
      const r1=Math.max(2,e.range.getRow()),r2=Math.max(r1,e.range.getLastRow());
      if(name==='Booking'){const col=bookingSyncHeaderMap_(pnHeaders_(sh));bookingSyncFillMissingBookingDates_(sh,col,r1,r2);}
      if(name==='Chứng từ_FF'||name==='Chứng từ_full')bookingSyncRepairControlColumnsForEditedTargetRange_(e.range);
      if(name==='TT Nhập'||name==='DS SKU')onEditSyncTTNhapDSSKU(e);
      if(name==='Hàng lỗi'||name==='DS SKU')onEditSyncHangLoiGHTK(e);
      let result=null;
      if(p||name==='DS BC'||name==='TT Nhập'){
        const rows=new Set();if(p)for(let r=r1;r<=r2;r++)rows.add(r);
        result=pnV2SyncAll_(p?{pair:p.main,prefer:{side:name===p.main?'main':'full',rows}}:{});
        if(name==='Booking_full')pnSyncBookingFullRows_(r1,r2);
      }
      if(name==='Bàn giao chứng từ'&&e.range.getA1Notation()==='A1'){refreshBanGiaoChungTuDropdown();syncBanGiaoChungTuBySelectedDate();}
      pnAfterActivity_();
      return result;
    });
  } catch(err) {
    if(/PN_BUSY/.test(String(err&&err.message||err))){PropertiesService.getScriptProperties().setProperty('PN_SYNC_REQUESTED',String(Date.now()));return {ok:false,busy:true};}
    throw err;
  } finally { pnV2Start_=0; }
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
// Lưới an toàn 5 phút: bắt thay đổi không qua onEdit (script, API, dán từ nơi khác). Không chạy VHFF ở đây.
// Chạy tay khi cần đối soát lại toàn bộ (ví dụ sau khi dán dữ liệu bằng script khác). KHÔNG cài trigger theo giờ.
function pnScheduledReconcile() {
  try {
    return pnWithLock_(()=>{
      pnRequireReady_();
      pnV2Start_=Date.now();
      const result=pnV2SyncAll_({});
      if(!result.partial)PropertiesService.getScriptProperties().deleteProperty('PN_SYNC_REQUESTED');
      return result;
    });
  } catch(err) {
    if(/PN_BUSY/.test(String(err&&err.message||err)))return {ok:false,busy:true};
    throw err;
  } finally { pnV2Start_=0; }
}
// Việc định kỳ chạy theo sự kiện (sửa Sheet, mở Sheet), trong cùng khóa và trong thời gian cho phép:
// - Dọn tháng cũ: 1 lần/ngày, ở lần đầu tiên trong ngày.
// - Lần sửa trước bị bỏ qua vì khóa bận: đồng bộ bù.
// - VHFF: khi có thay đổi, mỗi lần làm 1 phần (cách nhau ít nhất 2 phút) để không giữ khóa lâu.
function pnAfterActivity_() {
  const props=PropertiesService.getScriptProperties();
  const today=Utilities.formatDate(new Date(),PN_FULL.timezone,'yyyy-MM-dd');
  if(props.getProperty('PN_DAILY_DONE')!==today&&!pnV2OverBudget_()){
    if(props.getProperty('PN_FULL_CLEANUP')==='enabled'){
      // Nhẹ: chỉ đồng bộ dòng thay đổi rồi dọn (vài giây). Không quét lại toàn bộ Booking để app không phải chờ.
      const sync=pnV2SyncAll_({});
      if(sync.partial)return;
      pnV2Cleanup_();
    }
    props.setProperty('PN_DAILY_DONE',today);
  }
  if(props.getProperty('PN_SYNC_REQUESTED')&&!pnV2OverBudget_()){
    const sync=pnV2SyncAll_({});
    if(!sync.partial)props.deleteProperty('PN_SYNC_REQUESTED');
  }
  const last=Number(props.getProperty('PN_VHFF_LAST_STEP')||0);
  if((props.getProperty('PN_VHFF_DIRTY')||props.getProperty('PN_VHFF_CYCLE'))&&Date.now()-last>120000&&!pnV2OverBudget_()){
    props.setProperty('PN_VHFF_LAST_STEP',String(Date.now()));
    try { pnV2VhffStep_(); } catch(err) { console.error('VHFF: '+String(err&&err.message||err)); }
  }
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
  PropertiesService.getScriptProperties().setProperty('PN_FULL_CLEANUP','enabled');
  return 'Đã cài trigger: khi sửa Sheet và khi mở Sheet. Không có trigger theo giờ. Dọn tháng cũ chạy 1 lần/ngày ở lần sửa/mở đầu tiên.';
}

function pnOpen() {
  try {
    return pnWithLock_(()=>{
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

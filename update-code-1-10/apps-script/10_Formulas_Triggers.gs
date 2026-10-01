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
function pnStampFullIds_(p) {
  const sh=pnSheet_(p.full),rows=pnRows_(sh,31),updates=[];
  rows.forEach((r,i)=>{
    if(!pnHasData_(r,p)||r[29])return;
    updates.push({row:i+2,values:['full:'+p.main+':'+Utilities.getUuid(),pnFingerprint_(r,p)]});
  });
  pnWriteBlocks_(sh,updates,30,2);
}
function pnHandleEdit(e) {
  if(!e||!e.range)return;
  return pnWithLock_(()=>{
    pnRequireReady_();
    const name=e.range.getSheet().getName(),p=pnPair_(name);
    // Manual edits are authoritative changes, not a blind full-row overwrite.
    if(p&&name===p.full)pnStampFullIds_(p);
    if(name==='Booking') {
      const sh=e.range.getSheet(),col=bookingSyncHeaderMap_(pnHeaders_(sh));
      bookingSyncFillMissingBookingDates_(sh,col,Math.max(2,e.range.getRow()),e.range.getLastRow());
    }
    pnMirrorAll();
    if(name==='Booking'||name==='Booking_full') {
      if(name==='Booking') {
        syncBookingRowsToChungTuFF_(Math.max(2,e.range.getRow()),e.range.getLastRow());
        bookingSyncUpdateTargetDatesFromBookingRows_(Math.max(2,e.range.getRow()),e.range.getLastRow());
      } else pnSyncBookingFullRows_(Math.max(2,e.range.getRow()),e.range.getLastRow());
    }
    if(['Booking','Booking_full','File đơn','File đơn_full','DS BC','TT Nhập'].includes(name)){
      fillMissingGhtkCodesInChungTuFF();
      pnSyncAreaAll_();
    }
    if(name==='Chứng từ_FF')bookingSyncRepairControlColumnsForEditedTargetRange_(e.range);
    if(name==='Chứng từ_full')bookingSyncRepairControlColumnsForEditedTargetRange_(e.range);
    if(name==='TT Nhập'||name==='DS SKU')onEditSyncTTNhapDSSKU(e);
    if(name==='Hàng lỗi'||name==='DS SKU')onEditSyncHangLoiGHTK(e);
    pnStampFullIds_(PN_FULL.pairs[2]);
    pnMirrorAll();
    pnMarkDirty_();
    if(name==='Bàn giao chứng từ'&&e.range.getA1Notation()==='A1'){
      refreshBanGiaoChungTuDropdown();syncBanGiaoChungTuBySelectedDate();
    }
  });
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
function pnScheduledReconcile() {
  return pnWithLock_(()=>{
    pnRequireReady_();
    pnMirrorAll();
    // Repairs script/API imports that do not fire onEdit.
    const sh=pnSheet_('Booking_full');
    if(sh.getLastRow()>1){
      pnSyncBookingFullRows_(2,sh.getLastRow());
    }
    fillMissingGhtkCodesInChungTuFF();
    pnSyncAreaAll_();
    pnStampFullIds_(PN_FULL.pairs[2]);
    pnMirrorAll();
    return pnSyncVHFF_();
  });
}
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
function pnInstallTriggers() {
  pnRequireReady_();
  const old=new Set([
    'onEditSyncBookingToChungTuFF','onChangeSyncBookingToChungTuFF','onEditSyncChungTuKhongDatYC',
    'onEditSyncBanGiaoChungTu','onOpenSyncBanGiaoChungTu','onEditSyncTTNhapDSSKU',
    'handleBanGiaoEdit','onEditSyncHangLoiGHTK','onEditSyncSuVuBanhXep','onEditSyncSuvuBanhXep',
    'onEditSyncKhuVucBookingToChungTuFF','onChangeSyncKhuVucBookingToChungTuFF',
    'pnHandleEdit','pnScheduledReconcile','pnCleanupDaily','pnOpen'
  ]);
  ScriptApp.getProjectTriggers().filter(t=>old.has(t.getHandlerFunction())).forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('pnHandleEdit').forSpreadsheet(PN_FULL.spreadsheetId).onEdit().create();
  ScriptApp.newTrigger('pnOpen').forSpreadsheet(PN_FULL.spreadsheetId).onOpen().create();
  ScriptApp.newTrigger('pnScheduledReconcile').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('pnCleanupDaily').timeBased().atHour(1).everyDays(1).inTimezone(PN_FULL.timezone).create();
  return 'Đã cài trigger trong project này. Dọn chỉ chạy sau pnEnableCleanup.';
}
function pnOpen() { pnRequireReady_();return onOpenSyncBanGiaoChungTu(); }

/* Full-tab extension. Install all nine scripts in the existing API project.
 * Business data is never deleted from *_full. Original source files are bundled separately.
 * AD:AE are hidden identity/base-signature columns; File đơn!T:W is untouched.
 */
const PN_FULL = {
  spreadsheetId: '1V39AVE2JfEtMPYqnG1-J75fKPDXU37fCnJ9Na3UXOco',
  epoch: 'pn-full-v1', timezone: 'Asia/Saigon', idCol: 30, baseCol: 31,
  pairs: [
    {main:'Booking', full:'Booking_full', width:16, date:0, key:1},
    {main:'File đơn', full:'File đơn_full', width:18, date:1, key:0, repeated:true},
    {main:'Chứng từ_FF', full:'Chứng từ_full', width:17, date:0, key:4, docs:true}
  ]
};
let pnLockDepth_ = 0;
let pnRequest_ = null;
function pnSS_() { return SpreadsheetApp.openById(PN_FULL.spreadsheetId); }
function pnSheet_(name) {
  const sh = pnSS_().getSheetByName(name);
  if (!sh) throw new Error('Thiếu tab: ' + name);
  return sh;
}
function pnPair_(name) { return PN_FULL.pairs.find(p => p.main === name || p.full === name); }
function pnText_(v) { return String(v == null ? '' : v).trim(); }
function pnNorm_(v) { return pnText_(v).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/\s+/g,' '); }
function pnHash_(v) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(v), Utilities.Charset.UTF_8)
    .map(x => ('0' + ((x + 256) % 256).toString(16)).slice(-2)).join('');
}
function pnWithLock_(fn, waitMs) {
  if (pnLockDepth_) return fn();
  const lock = LockService.getScriptLock();
  // Lỗi có mã PN_BUSY để apiSave_ trả về "đang bận, gửi lại" thay vì lỗi khó hiểu.
  if (!lock.tryLock(waitMs || 30000)) throw new Error('PN_BUSY: Hệ thống PN đang bận (khóa), thử lại sau.');
  pnLockDepth_++;
  try { return fn(); } finally { pnLockDepth_--; lock.releaseLock(); }
}

// API, đọc dữ liệu, VHFF chỉ dùng tab full -> KHÔNG bị chặn bởi việc dọn tab chính.
function pnRequireReady_() {
  if (PropertiesService.getScriptProperties().getProperty('PN_FULL_READY') !== PN_FULL.epoch)
    throw new Error('Nguồn full chưa đối soát xong. Chạy pnMigrateFull trước khi chuyển API.');
}
// Trước khi đụng tab chính: nếu còn lượt dọn dở (bản cũ), tự sửa khi đang giữ khóa của Sheet.
function pnMainReady_() {
  const raw = PropertiesService.getScriptProperties().getProperty('PN_CLEANUP_PENDING');
  if (!raw) return;
  if (pnLockDepth_ && JSON.parse(raw).engine === PN_V2.engine) { pnV2ResumeCleanup_(); return; }
  throw new Error('PN_BUSY: Lượt dọn tab chính trước bị gián đoạn; mở/sửa Sheet để tự hoàn tất hoặc chạy pnResumeCleanup.');
}

function pnSize_(sh, rows, cols) {
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
}
// Tiêu đề dùng cho map cột của các script: tab chính = khối nghiệp vụ liền nhau từ A; tab full = mọi cột.
function pnHeaders_(sh) {
  const pair = pnPair_(sh.getName());
  if (!pair) return sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getDisplayValues()[0];
  const L = pnV2Layout_(sh, sh.getName() === pair.full);
  return L.isFull ? L.headers : L.headers.slice(0, L.width);
}

function pnRows_(sh, width) {
  return sh.getLastRow() < 2 ? [] : sh.getRange(2,1,sh.getLastRow()-1,width).getValues();
}
function pnHasData_(r,p) { return r.slice(0,p.width).some(v => v !== '' && v != null && v !== false); }
function pnKey_(r,p) {
  const key = pnText_(r[p.key]);
  if(key) return (p.docs ? 'ORDER:' : '') + key.toLowerCase();
  return p.docs && pnText_(r[1]) ? 'GHTK:' + pnText_(r[1]).toLowerCase() : '';
}
function pnIsoDate_(value) {
  if (value instanceof Date && !isNaN(value)) return Utilities.formatDate(value,PN_FULL.timezone,'yyyy-MM-dd');
  const s=pnText_(value);
  let m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s|$)/);
  if(m) m=[m[0],m[3],m[2],m[1]];
  else m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T]|$)/);
  if(!m) return '';
  const y=+m[1],mo=+m[2],d=+m[3],test=new Date(Date.UTC(y,mo-1,d));
  if(test.getUTCFullYear()!==y || test.getUTCMonth()!==mo-1 || test.getUTCDate()!==d) return '';
  return m[1]+'-'+String(mo).padStart(2,'0')+'-'+String(d).padStart(2,'0');
}
function pnMonth_(value) { return pnIsoDate_(value).slice(0,7); }
function pnCurrentMonth_() { return Utilities.formatDate(new Date(),PN_FULL.timezone,'yyyy-MM'); }
function pnEligible_(row,p,currentMonth) {
  if(!pnKey_(row,p)) return false;
  const month=pnMonth_(row[p.date]);
  if(!month || month >= currentMonth) return false;
  return !p.docs || ['da nhan chung tu','shop huy od'].includes(pnNorm_(row[6]));
}
function pnNeedsMain_(row,p,currentMonth) {
  const month=pnMonth_(row[p.date]);
  // Unknown dates/statuses are kept. Never silently discard exceptions.
  return !month || month >= currentMonth || !pnEligible_(row,p,currentMonth);
}
function pnFingerprint_(row,p) {
  // Typed values are retained in Sheets; hash equates date text with a true date for business date columns.
  return pnHash_(row.slice(0,p.width).map((v,i) => {
    if(i===p.date || (p.docs && i===7)) return pnIsoDate_(v) || pnText_(v);
    return v instanceof Date ? Utilities.formatDate(v,PN_FULL.timezone,'yyyy-MM-dd HH:mm:ss') : v;
  }));
}
function pnSetupPair_(p) {
  const main=pnSheet_(p.main), full=pnSheet_(p.full);
  [main,full].forEach(sh => pnSize_(sh,2,31));
  const headers=main.getRange(1,1,1,p.width).getValues()[0];
  if(headers.some(x=>!pnText_(x))) throw new Error('Header nguồn thiếu: '+p.main);
  const old=full.getRange(1,1,1,p.width).getValues()[0];
  if(old.some(pnText_) && old.some((x,i)=>pnNorm_(x)!==pnNorm_(headers[i])))
    throw new Error('Header full khác nguồn: '+p.full);
  full.getRange(1,1,1,p.width+1).setValues([headers.concat(['Tháng'])]);
  [main,full].forEach(sh => {
    const meta=sh.getRange(1,30,1,2).getValues()[0];
    if(meta.some((v,i)=>v && v!==['__PN_ID','__PN_BASE'][i])) throw new Error('AD:AE đã có dữ liệu: '+sh.getName());
    sh.getRange(1,30,1,2).setValues([['__PN_ID','__PN_BASE']]);
    sh.hideColumns(30,2);
  });
  full.setFrozenRows(1);
}
function pnWriteBlocks_(sh, updates, col, width) {
  const sorted=updates.slice().sort((a,b)=>a.row-b.row);
  for(let i=0;i<sorted.length;) {
    const start=sorted[i].row, values=[sorted[i].values]; let j=i+1;
    while(j<sorted.length && sorted[j].row===start+values.length) values.push(sorted[j++].values);
    pnSize_(sh,start+values.length-1,col+width-1);
    sh.getRange(start,col,values.length,width).setValues(values);
    i=j;
  }
}
function pnSetBusiness_(sh,row,values,p) {
  pnSize_(sh,row,31);
  // Mirror formatting from the primary so getDisplayValues keeps dates/money/checkboxes compatible.
  const source=pnSheet_(p.main);
  source.getRange(2,1,1,p.width).copyTo(sh.getRange(row,1,1,p.width),SpreadsheetApp.CopyPasteType.PASTE_FORMAT,false);
  source.getRange(2,1,1,p.width).copyTo(sh.getRange(row,1,1,p.width),SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,false);
  sh.getRange(row,1,1,p.width).setValues([values.slice(0,p.width)]);
}
function pnPlanPair_(p,primary,history) {
  const byId=new Map(),byKey=new Map(),used=new Set(),issues=[];
  history.forEach((r,i)=>{
    const id=pnText_(r[29]),key=pnKey_(r,p);
    if(!id && pnHasData_(r,p)) { issues.push('Full có dòng chưa có ID: '+(i+2)); return; }
    if(id){if(byId.has(id))issues.push('Full trùng ID: '+id);byId.set(id,{r,index:i});}
    if(key&&!p.repeated){if(byKey.has(key))issues.push('Full trùng khóa: '+key);byKey.set(key,{r,index:i});}
  });
  const mainKeys=new Set(),changes=[];
  primary.forEach((r,i)=>{
    if(!pnHasData_(r,p))return;
    const key=pnKey_(r,p),givenId=pnText_(r[29]);
    if(!p.repeated&&key){if(mainKeys.has(key))issues.push('Chính trùng khóa '+key);mainKeys.add(key);}
    const seed='seed:'+p.main+':'+(i+2)+':'+pnFingerprint_(r,p);
    let found=givenId ? byId.get(givenId) : ((!p.repeated&&key ? byKey.get(key) : null)||byId.get(seed));
    if(givenId && !found){issues.push('ID nguồn không còn trong full: '+givenId);return;}
    if(!p.repeated&&key&&byKey.has(key)&&(!found||byKey.get(key).index!==found.index)){
      issues.push('Khóa mới đã thuộc hồ sơ full khác: '+key);return;
    }
    if(found && used.has(found.index)){issues.push('Hai dòng chính cùng hồ sơ full');return;}
    if(found)used.add(found.index);
    const ph=pnFingerprint_(r,p),base=pnText_(r[30]),fh=found?pnFingerprint_(found.r,p):'';
    let direction='equal';
    if(!found)direction='append';
    else if(ph!==fh) {
      if(base===ph)direction='pull';
      else if(base===fh)direction='push';
      else {issues.push('Xung đột hai bản '+p.main+' dòng '+(i+2));return;}
    }
    changes.push({mainRow:i+2,fullRow:found?found.index+2:0,id:found?pnText_(found.r[29]):seed,
      direction,main:r,full:found&&found.r,values:(direction==='pull'?found.r:r).slice(0,p.width)});
  });
  return {changes,issues};
}
function pnReconcilePair_(p) {
  const main=pnSheet_(p.main),full=pnSheet_(p.full);
  const primary=pnRows_(main,31),history=pnRows_(full,31);
  const props=PropertiesService.getScriptProperties(),countKey='PN_FULL_COUNT_'+p.main;
  const count=history.filter(r=>pnHasData_(r,p)).length;
  if(count<Number(props.getProperty(countKey)||0))throw new Error('Full bị giảm số hồ sơ; dừng để kiểm tra: '+p.full);
  const plan=pnPlanPair_(p,primary,history);
  if(plan.issues.length)throw new Error(plan.issues.slice(0,12).join('\n'));
  let next=Math.max(2,full.getLastRow()+1),changed=0;
  const writes=[],pulls=[],metas=[];
  // Commit full first; primary identity checkpoint follows. A failed retry matches business keys.
  for(const c of plan.changes) {
    const fh=pnFingerprint_(c.values,p);
    if(c.direction==='append'||c.direction==='push'){
      c.fullRow=c.fullRow||next++;
      const output=new Array(31).fill('');
      c.values.forEach((v,i)=>output[i]=v);
      output[p.width]=pnMonth_(c.values[p.date])||'CẦN KIỂM TRA';
      output[29]=c.id;output[30]=fh;
      writes.push({row:c.fullRow,values:output});
      changed++;
    }
    if(c.direction==='pull'){
      pulls.push({row:c.mainRow,values:c.values});
      changed++;
    }
    if(pnText_(c.main[29])!==c.id || pnText_(c.main[30])!==fh)
      metas.push({row:c.mainRow,values:[c.id,fh]});
  }
  pnWriteBlocks_(full,writes,1,31);
  // Preserve each source row's types/formats; never repaint history from row 2.
  pnCopyMappedFormats_(main,full,plan.changes.filter(c=>c.direction==='append'||c.direction==='push'),p.width);
  pnWriteBlocks_(main,pulls,1,p.width);
  pnWriteBlocks_(main,metas,30,2);
  // Refresh Month for edits made directly on full, without rewriting business data.
  const fullRows=pnRows_(full,31),months=[];
  fullRows.forEach((r,i)=>{if(pnHasData_(r,p)){
    const month=pnMonth_(r[p.date])||'CẦN KIỂM TRA';
    if(r[p.width]!==month)months.push({row:i+2,values:[month]});
  }});
  pnWriteBlocks_(full,months,p.width+1,1);
  props.setProperty(countKey,String(fullRows.filter(r=>pnHasData_(r,p)).length));
  if(changed)pnMarkDirty_();
  return {pair:p.main,changed};
}
function pnMaterializeDocs_() {
  const p=PN_FULL.pairs[2],main=pnSheet_(p.main),full=pnSheet_(p.full),rows=pnRows_(main,31);
  const ids=new Set(rows.map(r=>pnText_(r[29])).filter(Boolean));
  let next=2; rows.forEach((r,i)=>{if(pnHasData_(r,p))next=i+3;});
  const business=[],metadata=[];
  pnRows_(full,31).forEach(r=>{
    if(!pnHasData_(r,p)||!pnKey_(r,p)||!r[29]||ids.has(r[29])||!pnNeedsMain_(r,p,pnCurrentMonth_()))return;
    business.push({row:next,values:r.slice(0,p.width)});
    metadata.push({row:next,values:[r[29],pnFingerprint_(r,p)]});
    ids.add(r[29]);next++;
  });
  if(business.length){
    const start=business[0].row;
    pnSize_(main,next,31);
    main.getRange(2,1,1,p.width).copyTo(main.getRange(start,1,business.length,p.width),SpreadsheetApp.CopyPasteType.PASTE_FORMAT,false);
    main.getRange(2,1,1,p.width).copyTo(main.getRange(start,1,business.length,p.width),SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,false);
    pnWriteBlocks_(main,business,1,p.width);
    pnWriteBlocks_(main,metadata,30,2);
  }
}
function pnMarkDirty_() {
  const props=PropertiesService.getScriptProperties(),version=String(Date.now())+':'+Utilities.getUuid();
  props.setProperty('PN_VHFF_DIRTY',version);
  props.setProperty('PN_DOCOPS_DATA_VERSION',version);
}
function pnAudit_() {
  const results=PN_FULL.pairs.map(p=>{
    const primary=pnRows_(pnSheet_(p.main),31),history=pnRows_(pnSheet_(p.full),31),byId=new Map();
    const errors=[];history.forEach(r=>{if(r[29]){if(byId.has(r[29]))errors.push('Trùng ID full');byId.set(r[29],r);}});
    primary.forEach((r,i)=>{
      if(!pnHasData_(r,p))return;
      const other=byId.get(r[29]);
      if(!other||pnFingerprint_(r,p)!==pnFingerprint_(other,p))errors.push('Chưa khớp dòng '+(i+2));
    });
    return {sheet:p.main,primary:primary.filter(r=>pnHasData_(r,p)).length,full:history.filter(r=>pnHasData_(r,p)).length,errors};
  });
  return {ok:results.every(r=>!r.errors.length),results};
}
function pnAuditFull() { return pnWithLock_(()=>{const r=PN_FULL.pairs.map(pnV2Audit_);console.log(JSON.stringify(r));return {ok:r.every(x=>x.ok),results:r};}); }

function pnCopyMappedFormats_(source,target,changes,width) {
  for(let i=0;i<changes.length;) {
    const first=changes[i];let j=i+1;
    while(j<changes.length&&changes[j].mainRow===first.mainRow+j-i&&changes[j].fullRow===first.fullRow+j-i)j++;
    const from=source.getRange(first.mainRow,1,j-i,width),to=target.getRange(first.fullRow,1,j-i,width);
    from.copyTo(to,SpreadsheetApp.CopyPasteType.PASTE_FORMAT,false);
    from.copyTo(to,SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,false);
    i=j;
  }
}
function pnFormatOnlyEquivalent_(a,b) {
  if(a instanceof Date && b instanceof Date)return a.getTime()===b.getTime();
  if(a===b)return true;
  const date=a instanceof Date?a:b instanceof Date?b:null;
  const number=a instanceof Date?b:a;
  if(!date||typeof number!=='number'||!Number.isFinite(number))return false;
  const local=Utilities.formatDate(date,PN_FULL.timezone,'yyyy-MM-dd HH:mm:ss');
  const serial=(Date.parse(local.replace(' ','T')+'Z')-Date.UTC(1899,11,30))/86400000;
  return Math.abs(serial-number)<1e-8;
}
// Recovery for the initial migration's row-2 format bug. No business values/IDs are written.
function pnRepairMigrationFormats() {
  return pnWithLock_(()=>{
    const props=PropertiesService.getScriptProperties();
    if(props.getProperty('PN_FULL_READY')||props.getProperty('PN_CLEANUP_PENDING'))
      throw new Error('Chỉ dùng phục hồi migration chưa hoàn tất, chưa dọn.');
    const plans=[],pendingMigration=[];
    PN_FULL.pairs.forEach(p=>{
      const main=pnSheet_(p.main),full=pnSheet_(p.full),history=pnRows_(full,31),byId=new Map();
      history.forEach((r,i)=>{if(r[29]){if(byId.has(r[29]))throw new Error('Trùng ID full');byId.set(r[29],{r,row:i+2});}});
      const changes=[];
      pnRows_(main,31).forEach((r,i)=>{
        if(!pnHasData_(r,p))return;
        if(!pnText_(r[29])) {
          if(pnText_(r[30]))throw new Error('Có checkpoint nhưng mất ID: '+p.main+' dòng '+(i+2));
          pendingMigration.push({sheet:p.main,row:i+2});
          return;
        }
        const match=byId.get(r[29]);
        if(!match)throw new Error('Thiếu ID full: '+p.main+' dòng '+(i+2));
        if(pnFingerprint_(r,p)===pnFingerprint_(match.r,p))return;
        // Reconstruct full's typed values without accepting any changed business value.
        // A primary-only edit may wait for normal three-way reconciliation.
        const typedFull=match.r.slice();let formatChanged=false;
        r.slice(0,p.width).forEach((v,c)=>{
          if((v instanceof Date)!==(match.r[c] instanceof Date)&&pnFormatOnlyEquivalent_(v,match.r[c])) {
            typedFull[c]=v;formatChanged=true;
          }
        });
        const base=pnText_(r[30]);
        if(!base||match.r[30]!==base||pnFingerprint_(typedFull,p)!==base)
          throw new Error('Có thay đổi dữ liệu thật; không tự sửa: '+p.main+' dòng '+(i+2));
        if(pnFingerprint_(r,p)!==base)pendingMigration.push({sheet:p.main,row:i+2,reason:'primaryChangedOnly'});
        if(formatChanged)changes.push({mainRow:i+2,fullRow:match.row});
      });
      plans.push({p,main,full,changes});
    });
    plans.forEach(x=>pnCopyMappedFormats_(x.main,x.full,x.changes,x.p.width));
    SpreadsheetApp.flush();
    // Only uncheckpointed rows may remain pending; never weaken the migration audit.
    const audit=pnAudit_();
    const unexpected=audit.results.flatMap(result=>result.errors.filter(error=>
      !pendingMigration.some(x=>x.sheet===result.sheet&&error==='Chưa khớp dòng '+x.row)));
    if(unexpected.length)throw new Error(JSON.stringify(audit));
    const result={ok:true,scope:'formatRepair',migrationReady:audit.ok,pendingMigration,audit};
    console.log(JSON.stringify(result));return result;
  });
}
// One entry point for the interrupted migration with Booking metadata shifted to Q:R.
function pnRecoverFullMigration() {
  return pnWithLock_(()=>{
    const props=PropertiesService.getScriptProperties();
    if(props.getProperty('PN_FULL_READY')||props.getProperty('PN_CLEANUP_PENDING'))
      throw new Error('Chỉ phục hồi migration chưa hoàn tất.');
    const p=PN_FULL.pairs[0],main=pnSheet_(p.main),full=pnSheet_(p.full);
    const headers=main.getRange(1,17,1,2).getValues()[0];
    const rows=pnRows_(main,31),history=pnRows_(full,31),byId=new Map(),updates=[];
    history.forEach(r=>{if(r[29]){if(byId.has(r[29]))throw new Error('Trùng ID full');byId.set(r[29],r);}});
    const virtual=rows.map(r=>r.slice());
    if(headers[0]==='__PN_ID'&&headers[1]==='__PN_BASE') {
      virtual.forEach((r,i)=>{
        if(!pnHasData_(r,p)||!pnText_(r[16]))return;
        const id=pnText_(r[16]),base=pnText_(r[17]),other=byId.get(id);
        if(!other||!base||base!==other[30]||pnKey_(r,p)!==pnKey_(other,p))
          throw new Error('ID Q:R không khớp hồ sơ full: Booking dòng '+(i+2));
        if(r[29]||r[30]) {
          if(r[29]!==id||r[30]!==base)throw new Error('AD:AE khác Q:R: Booking dòng '+(i+2));
          return;
        }
        r[29]=id;r[30]=base;updates.push({row:i+2,values:[id,base]});
      });
    }
    // Preflight all pairs using virtual restored identities and equivalent Date types.
    // Real two-sided edits, mismatched keys and duplicate IDs stop before any write.
    PN_FULL.pairs.forEach(pair=>{
      const primary=pair.main===p.main?virtual:pnRows_(pnSheet_(pair.main),31);
      const archived=(pair.main===p.main?history:pnRows_(pnSheet_(pair.full),31)).map(r=>r.slice());
      const indexed=new Map();archived.forEach(r=>{if(r[29])indexed.set(r[29],r);});
      primary.forEach(r=>{
        const other=indexed.get(r[29]);if(!other)return;
        r.slice(0,pair.width).forEach((v,c)=>{
          if((v instanceof Date)!==(other[c] instanceof Date)&&pnFormatOnlyEquivalent_(v,other[c]))other[c]=v;
        });
      });
      const plan=pnPlanPair_(pair,primary,archived);
      if(plan.issues.length)throw new Error('Chưa ghi dữ liệu. '+plan.issues.slice(0,12).join('\n'));
    });
    // Copy only verified metadata to the original fixed columns; retain Q:R as evidence.
    pnSize_(main,Math.max(2,main.getLastRow()),31);
    main.getRange(1,30,1,2).setValues([['__PN_ID','__PN_BASE']]);
    pnWriteBlocks_(main,updates,30,2);
    SpreadsheetApp.flush();
    pnRepairMigrationFormats();
    const result=pnMigrateFull();
    console.log('RECOVERY_COMPLETE: '+JSON.stringify(result));
    return result;
  });
}
function pnMigrateFull() {
  return pnWithLock_(()=>{
    // Does not clear primary or publish external reports.
    PN_FULL.pairs.forEach(pnSetupPair_);
    pnConvertFormulas_();
    PN_FULL.pairs.forEach(pnReconcilePair_);
    PN_FULL.pairs.forEach(pnEnsureFullFilter_);
    const audit=pnAudit_();
    if(!audit.ok)throw new Error(JSON.stringify(audit));
    PropertiesService.getScriptProperties().setProperty('PN_FULL_READY',PN_FULL.epoch);
    pnMarkDirty_();
    console.log(JSON.stringify(audit));
    return audit;
  });
}
// Chạy tay một lượt đồng bộ đầy đủ và trả kết quả đối soát từng tab.
function pnMirrorAll() {
  return pnWithLock_(()=>{
    pnRequireReady_();
    const started=!pnV2Start_; if(started)pnV2Start_=Date.now();
    try {
      // Sửa mã số dài bị Sheets đổi thành số ở tab chính trước khi so sánh.
      const longCodes=pnV2RepairLongCodes_();
      const sync=pnV2SyncAll_({});
      const audit=PN_FULL.pairs.map(pnV2Audit_);
      return {ok:audit.every(a=>a.ok),longCodes,sync,results:audit};
    } finally { if(started)pnV2Start_=0; }
  });
}

function pnEnsureFullFilter_(p) {
  const sh=pnSheet_(p.full);
  sh.hideColumns(p.width+2,31-p.width-1);
  const old=sh.getFilter(),needed=Math.max(5000,sh.getLastRow()+1000);
  if(old&&old.getRange().getNumRows()>sh.getLastRow()&&old.getRange().getNumColumns()===31)return;
  const criteria=[];
  if(old){
    for(let col=1;col<=old.getRange().getNumColumns();col++)criteria.push(old.getColumnFilterCriteria(col));
    old.remove();
  }
  pnSize_(sh,needed,31);
  // Include hidden IDs in the filter range, so a filter sort cannot detach identity from data.
  const filter=sh.getRange(1,1,needed,31).createFilter();
  criteria.forEach((c,i)=>{if(c)filter.setColumnFilterCriteria(i+1,c);});
}
// Xem trước: các dòng tháng cũ sẽ rời tab chính (đã khớp full; chứng từ chỉ khi đã nhận/hủy).
function pnCleanupPreview() {
  return pnWithLock_(()=>{
    pnRequireReady_();
    const month=pnCurrentMonth_();
    return {month,audit:PN_FULL.pairs.map(pnV2Audit_),candidates:PN_FULL.pairs.map(p=>{
      const cfg=pnV2Cfg_(p),M=pnV2EnsureColumns_(p).M;
      return {sheet:p.main,rows:pnV2Read_(M).map((r,i)=>pnV2Rec_(M,cfg,r)&&pnV2Eligible_(M,cfg,r,month)?i+2:0).filter(Boolean)};
    })};
  });
}

// 1 giờ sáng: đồng bộ đủ (kể cả quét toàn bộ Booking -> Chứng từ một lần/ngày), rồi dọn tháng cũ.
// Không phụ thuộc VHFF: dọn chỉ bỏ khỏi tab chính dòng đã có y hệt ở full.
function pnCleanupDaily() {
  return pnWithLock_(()=>{
    if(PropertiesService.getScriptProperties().getProperty('PN_FULL_CLEANUP')!=='enabled')
      return {ok:false,reason:'Dọn tự động đang tắt. Bật bằng pnEnableCleanup.'};
    pnRequireReady_();
    pnV2Start_=Date.now();
    try {
      const sync=pnV2SyncAll_({fullSweep:true,areaSweep:true});
      if(sync.partial)return {ok:false,reason:'Đồng bộ chưa xong trong thời gian cho phép; hoãn dọn.'};
      return pnV2Cleanup_();
    } finally { pnV2Start_=0; }
  });
}

function pnResumeCleanup() {
  return pnWithLock_(()=>{
    const props=PropertiesService.getScriptProperties(),raw=props.getProperty('PN_CLEANUP_PENDING');
    if(!raw)return {ok:true,skipped:true};
    if(JSON.parse(raw).engine===PN_V2.engine)return pnV2ResumeCleanup_();
    return pnResumeCleanupLegacy_();
  });
}
function pnResumeCleanupLegacy_() {
  return pnWithLock_(()=>{
    const props=PropertiesService.getScriptProperties(),raw=props.getProperty('PN_CLEANUP_PENDING');
    if(!raw)return {ok:true,skipped:true};
    const state=JSON.parse(raw),p=pnPair_(state.main),info=JSON.parse(props.getProperty('PN_BACKUP_'+p.main)||'null');
    if(!info)throw new Error('Thiếu thông tin backup.');
    const backup=pnSheet_('_PN_BACKUP_'+p.main).getRange(2,1,info.rows,31).getValues();
    if(pnHash_(backup)!==info.hash)throw new Error('Backup không khớp; không tự phục hồi.');
    const kept=backup.filter(r=>pnHasData_(r,p)&&!pnEligible_(r,p,state.month)),sh=pnSheet_(p.main);
    const current=sh.getRange(2,1,info.rows,31).getValues();
    const wanted=Array.from({length:info.rows},(_,i)=>kept[i]||new Array(31).fill(''));
    // Each range write is atomic; accept only before/after states of the interrupted writes.
    // Never overwrite a manual edit made after interruption.
    for(let i=0;i<info.rows;i++){
      for(const [start,end] of [[0,p.width],[29,31]]){
        const h=pnHash_(current[i].slice(start,end));
        if(h!==pnHash_(backup[i].slice(start,end))&&h!==pnHash_(wanted[i].slice(start,end)))
          throw new Error('Có sửa tay sau khi dọn lỗi ở dòng '+(i+2)+'; cần đối chiếu, không ghi đè.');
      }
    }
    if(pnRows_(sh,31).slice(info.rows).some(r=>pnHasData_(r,p)))throw new Error('Có dòng mới sau lượt dọn; cần đối chiếu trước.');
    sh.getRange(2,1,info.rows,p.width).setValues(wanted.map(r=>r.slice(0,p.width)));
    sh.getRange(2,30,info.rows,2).setValues(wanted.map(r=>r.slice(29,31)));
    const after=sh.getRange(2,1,info.rows,31).getValues();
    if(after.some((r,i)=>pnHash_([r.slice(0,p.width),r.slice(29,31)])!==pnHash_([wanted[i].slice(0,p.width),wanted[i].slice(29,31)])))
      throw new Error('Phục hồi chưa khớp; giữ khóa.');
    props.deleteProperty('PN_CLEANUP_PENDING');
    return {ok:true,sheet:p.main,removed:backup.filter(r=>pnEligible_(r,p,state.month)).length};
  });
}

function pnEnableCleanup() {
  PropertiesService.getScriptProperties().setProperty('PN_FULL_CLEANUP','enabled');
  return pnCleanupPreview();
}

function pnDisableCleanup() { PropertiesService.getScriptProperties().deleteProperty('PN_FULL_CLEANUP'); }
function pnBackupCleanup_(p,rows) {
  const name='_PN_BACKUP_'+p.main,ss=pnSS_(),sh=ss.getSheetByName(name)||ss.insertSheet(name);
  pnSize_(sh,rows.length+1,31);
  sh.getRange(1,1,1,31).setValues([pnSheet_(p.main).getRange(1,1,1,31).getValues()[0]]);
  if(rows.length)sh.getRange(2,1,rows.length,31).setValues(rows);
  PropertiesService.getScriptProperties().setProperty('PN_BACKUP_'+p.main,JSON.stringify({rows:rows.length,at:new Date().toISOString(),hash:pnHash_(rows)}));
  sh.hideSheet();
}

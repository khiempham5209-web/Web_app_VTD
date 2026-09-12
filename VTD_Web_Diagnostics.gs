/* Add as a NEW file in the existing VTD Apps Script project. */
function vtdDiagnosticsReceive(params) {
  const auth=vtdApp_auth_({sessionToken:String(params.sessionToken||'')});
  if(!auth.allowed)return {ok:false,message:'Diagnostics authentication required'};
  const batch=Array.isArray(params.events)?params.events.slice(0,20):[];
  const clean=function(v,n){return String(v==null?'':v).replace(/^[=+@-]/,"'").slice(0,n);};
  const allowed=['localStorageBytes','online','visible','sessionAgeMs','reason','errorName','queueCount','action','durationMs','sameSession','line','restored','diagnosticsVersion'];
  const rows=batch.map(function(e){const detail={};allowed.forEach(function(k){const v=(e.detail||{})[k];if(typeof v==='boolean'||typeof v==='number')detail[k]=v;else if(typeof v==='string')detail[k]=v.slice(0,80);});return [new Date(),clean(e.id,100),clean(e.time,40),clean(auth.email,160),clean(e.user,160),clean(e.device,100),clean(e.version,30),clean(e.kind,60),JSON.stringify(detail)];});
  if(!rows.length)return {ok:true};
  const lock=LockService.getScriptLock();if(!lock.tryLock(5000))return {ok:false,message:'Diagnostics busy'};
  try{const ss=vtdApp_ss_();let sh=ss.getSheetByName('_VTD_WEB_DIAGNOSTICS');if(!sh){sh=ss.insertSheet('_VTD_WEB_DIAGNOSTICS');sh.appendRow(['ReceivedAt','EventId','ClientTime','AuthenticatedUser','ReportedUser','Device','Version','Event','Detail']);}
    // Event IDs let analysts identify retries without counting them twice.
    sh.getRange(sh.getLastRow()+1,1,rows.length,9).setValues(rows);
    return {ok:true,accepted:rows.length};
  }finally{lock.releaseLock();}
}

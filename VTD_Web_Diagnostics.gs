/* Add as a NEW file in the existing VTD Apps Script project. */
function vtdDiagnosticsReceive(params) {
  const auth=vtdApp_auth_({sessionToken:String(params.sessionToken||'')});
  if(!auth.allowed)return {ok:false,message:'Diagnostics authentication required'};
  const batch=Array.isArray(params.events)?params.events.slice(0,20):[];
  const clean=function(v,n){return String(v==null?'':v).replace(/^[=+@-]/,function(c){return "'"+c;}).slice(0,n);};
  const longFields={htmlTitle:180,errorExcerpt:600,diagnosticConclusion:400};
  const allowed=['retryAfter','responseRequestId','htmlTitle','errorExcerpt','evidenceCategory','conclusionBasis','diagnosticConclusion','evidenceTruncated','method','requestLength','responseFormat','serverOk','requestId','transport','requestHost','endpointId','concurrentRequests','redirected','responseType','responseLength','pageCategory','protocolPresent','servicePresent','jsonDiagnosticsVersion','status','contentType','responseHost','label','localStorageBytes','online','visible','sessionAgeMs','reason','errorName','queueCount','action','durationMs','sameSession','line','restored','diagnosticsVersion'];
  const rows=batch.filter(e=>e && typeof e.id==='string' && e.id.length>0 && e.id.length<=100).map(function(e){const detail={};allowed.forEach(function(k){const v=(e.detail||{})[k];if(typeof v==='boolean'||typeof v==='number')detail[k]=v;else if(typeof v==='string')detail[k]=v.slice(0,longFields[k]||80);});return [new Date(),clean(e.id,100),clean(e.time,40),clean(auth.email,160),clean(e.user,160),clean(e.device,100),clean(e.version,30),clean(e.kind,60),JSON.stringify(detail)];});
  if(!rows.length)return {ok:true};
  const lock=LockService.getScriptLock();if(!lock.tryLock(5000))return {ok:false,message:'Diagnostics busy'};
  try{const ss=vtdApp_ss_();let sh=ss.getSheetByName('_VTD_WEB_DIAGNOSTICS');if(!sh){sh=ss.insertSheet('_VTD_WEB_DIAGNOSTICS');sh.appendRow(['ReceivedAt','EventId','ClientTime','AuthenticatedUser','ReportedUser','Device','Version','Event','Detail']);}
    // Atomic dedup within the most recent 5,000 rows, including duplicates in this batch.
    const last=sh.getLastRow(), start=Math.max(2,last-4999);
    const previous=last>=2?sh.getRange(start,2,last-start+1,5).getValues():[];
    const key=r=>JSON.stringify([r[0],r[2],r[4]]);
    const seen=new Set(previous.map(key));
    const fresh=rows.filter(r=>{const k=JSON.stringify([r[1],r[3],r[5]]);if(seen.has(k))return false;seen.add(k);return true;});
    if(fresh.length)sh.getRange(last+1,1,fresh.length,9).setValues(fresh);
    return {ok:true,accepted:fresh.length,duplicates:rows.length-fresh.length};
  }finally{lock.releaseLock();}
}

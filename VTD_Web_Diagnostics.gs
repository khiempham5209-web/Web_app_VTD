/* Add as a NEW file in the existing VTD Apps Script project. */
function vtdDiagnosticsReceive(params) {
  params=params||{};
  if(JSON.stringify(params).length>100000)return {ok:false,message:'Diagnostics batch too large'};
  const token=String(params.sessionToken||'');
  const auth=token?vtdApp_auth_({sessionToken:token}):{allowed:false};
  const unverified=!auth.allowed;
  if(unverified && params.preAuthDiagnostics!==true)return {ok:false,message:'Diagnostics authentication required'};
  const publicKinds=new Set(['diagnostics_loaded','json_diagnostics_loaded','request_diagnostics_loaded','page_show','page_hide','login_screen','login_response_ok','login_client_failed','api_failed','api_invalid_response','json_diagnostic','request_started','request_headers','request_network_error','response_read_error','response_inspected','request_finished','javascript_error','unhandled_rejection']);
  const batch=Array.isArray(params.events)?params.events.slice(0,20):[];
  const clean=function(v,n){return String(v==null?'':v).replace(/^[=+@-]/,function(c){return "'"+c;}).slice(0,n);};
  const longFields={htmlTitle:180,errorExcerpt:600,diagnosticConclusion:400};
  const allowed=['retryAfter','responseRequestId','htmlTitle','errorExcerpt','evidenceCategory','conclusionBasis','diagnosticConclusion','evidenceTruncated','method','requestLength','responseFormat','serverOk','requestId','transport','requestHost','endpointId','concurrentRequests','redirected','responseType','responseLength','pageCategory','protocolPresent','servicePresent','jsonDiagnosticsVersion','status','contentType','responseHost','label','localStorageBytes','online','visible','sessionAgeMs','reason','errorName','queueCount','action','durationMs','sameSession','line','restored','diagnosticsVersion'];
  const rows=batch.filter(e=>e && (!unverified || publicKinds.has(e.kind)) && typeof e.id==='string' && e.id.length>0 && e.id.length<=100).map(function(e){const detail={};allowed.forEach(function(k){const v=(e.detail||{})[k];if(typeof v==='boolean'||typeof v==='number')detail[k]=v;else if(typeof v==='string')detail[k]=v.slice(0,longFields[k]||80);});detail.identityVerified=!unverified;detail.ingestMode=unverified?'pre_auth':'authenticated';return [new Date(),clean(e.id,100),clean(e.time,40),clean(unverified?'UNVERIFIED':auth.email,160),clean(unverified?'':e.user,160),clean(e.device,100),clean(e.version,30),clean(e.kind,60),JSON.stringify(detail)];});
  if(!rows.length)return {ok:true};
  const lock=LockService.getScriptLock();if(!lock.tryLock(5000))return {ok:false,message:'Diagnostics busy'};
  try{
    // One bounded state value, under the same script lock. No credential validation bypass.
    if(unverified){
      const props=PropertiesService.getScriptProperties(), key='VTD_PREAUTH_DIAG_RATE_V1';
      const minute=Math.floor(Date.now()/60000);
      let rate;try{rate=JSON.parse(props.getProperty(key)||'null');}catch(_){}
      if(!rate || rate.minute!==minute)rate={minute,total:0,devices:{}};
      const devices=Array.from(new Set(rows.map(r=>String(r[5]||'unknown'))));
      if(rate.total>=30 || devices.some(d=>(Number(rate.devices['d:'+d])||0)>=3))return {ok:false,message:'Diagnostics rate limited',retryAfterSeconds:60};
      rate.total++;devices.forEach(d=>{rate.devices['d:'+d]=(Number(rate.devices['d:'+d])||0)+1;});
      props.setProperty(key,JSON.stringify(rate));
    }
    const ss=vtdApp_ss_();let sh=ss.getSheetByName('_VTD_WEB_DIAGNOSTICS');if(!sh){sh=ss.insertSheet('_VTD_WEB_DIAGNOSTICS');sh.appendRow(['ReceivedAt','EventId','ClientTime','AuthenticatedUser','ReportedUser','Device','Version','Event','Detail']);}
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

/* Add this file to VTD and PN projects. Requires the small entry-point hook supplied with this release. */
function apiDiagCategory_(value) {
  const s=String(value || '').toLowerCase();
  if(/quota|too many|limit exceeded|quá nhiều/.test(s))return 'quota_or_rate_limit';
  if(/permission|access denied|authorization|quyền/.test(s))return 'permission';
  if(/timeout|timed out|execution time|quá lâu/.test(s))return 'timeout';
  if(/json|unexpected token/.test(s))return 'json_parse';
  if(/đăng nhập|authentication|session/.test(s))return 'authentication';
  if(/not found|does not exist|không tìm thấy/.test(s))return 'not_found';
  return 'other';
}
function apiDiagRun_(service, method, e, run) {
  let body={}, malformed=false;
  try{body=JSON.parse(e && e.postData && e.postData.contents || '{}');}catch(_){malformed=true;}
  if(!body || typeof body!=='object')body={};
  const p=body.params || {}, query=e && e.parameter || {};
  const id=String(p._diagRequestId || query._diagRequestId || '');
  if(!/^diag_[a-zA-Z0-9_-]{1,90}$/.test(id) || body.action==='clientDiagnostics')return run();
  const start=Date.now();
  const safe=v=>String(v || '').replace(/[^a-zA-Z0-9_.-]/g,'').slice(0,90);
  const d={requestId:id,service:service,method:method,action:safe(body.action || query.action || query._diagAction),requestLength:String(e && e.postData && e.postData.contents || '').length,requestMalformed:malformed,serverDiagnosticsVersion:'2'};
  const log=(level,value)=>{try{console[level](value);}catch(_){}};
  log('log',JSON.stringify(Object.assign({stage:'received'},d)));
  try {
    const response=run();
    d.stage='response_ready';
    let text='';try{text=response && typeof response.getContent==='function' ? response.getContent() : '';}catch(_){d.inspectionFailed=true;}
    d.responseLength=text.length;
    try{const result=JSON.parse(text);d.responseFormat='json';d.serverOk=result.ok===true;d.errorCategory=result.ok===false?apiDiagCategory_(result.message):'';d.protocolPresent=Object.prototype.hasOwnProperty.call(result,'protocol');
      // Function/line frames only. Never store the error message or business values.
      d.stackFrames=(String(result.message || '').match(/\bat [A-Za-z0-9_.$]+ \([^\n)]*:\d+(?::\d+)?\)/g)||[]).slice(0,5).map(x=>x.replace(/[^a-zA-Z0-9_.$(): /-]/g,'')).join(' | ').slice(0,500);
    }catch(_){d.responseFormat=/^\s*</.test(text)?'html':'non_json';}
    return response;
  }catch(err){d.stage='handler_threw';d.errorCategory=apiDiagCategory_(err);d.errorName=safe(err && err.name);throw err;}
  finally {
    d.durationMs=Date.now()-start;
    log('log',JSON.stringify(d));
    // Diagnostics must not replace or fail the business response.
    try {
      const lock=LockService.getScriptLock();
      if(lock.tryLock(200))try {
        const ss=service==='VTD'?vtdApp_ss_():SpreadsheetApp.openById(CONFIG.spreadsheetId);
        let sh=ss.getSheetByName('_API_REQUEST_DIAGNOSTICS');
        if(!sh){sh=ss.insertSheet('_API_REQUEST_DIAGNOSTICS');sh.appendRow(['ServerTime','RequestId','Service','Method','Action','Stage','DurationMs','Detail']);}
        // Bounded rolling log, no deletion of existing rows. Header remains at row 1.
        const key='API_DIAG_ROW_V2';const props=PropertiesService.getScriptProperties();
        const next=(Number(props.getProperty(key)||0)%5000)+1;
        if(sh.getMaxRows()<next+1)sh.insertRowsAfter(sh.getMaxRows(),next+1-sh.getMaxRows());
        sh.getRange(next+1,1,1,8).setValues([[new Date(),id,service,method,d.action,d.stage,d.durationMs,JSON.stringify(d)]]);
        props.setProperty(key,String(next));
      }finally{lock.releaseLock();}
      else log('warn','API_DIAG_WRITE_SKIPPED_LOCK '+id);
    }catch(_){log('warn','API_DIAG_WRITE_FAILED '+id);}
  }
}

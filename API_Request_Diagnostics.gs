
// Bounded evidence only; never store complete HTML or successful business JSON.
function vtdDiagRedact_(value, limit) {
  return String(value || '').slice(0,100000)
    .replace(/<script\b[^>]*>[\s\S]*?<\/script\s*>/gi,' ')
    .replace(/<style\b[^>]*>[\s\S]*?<\/style\s*>/gi,' ')
    .replace(/<!--[\s\S]*?-->/g,' ').replace(/<[^>]*>/g,' ')
    .replace(/&(?:nbsp|lt|gt|quot|amp|#\d+);/gi,' ')
    .replace(/https?:\/\/[^\s<>"']+/gi,'[URL]')
    .replace(/[\w.+-]+@[\w.-]+\.[a-z]{2,}/gi,'[EMAIL]')
    .replace(/((?:sessionToken|access_token|id_token|password|passwd|authorization|cookie|token)\s*[=:]\s*)[^\s,;]+/gi,'$1[REDACTED]')
    .replace(/Bearer\s+\S+/gi,'Bearer [REDACTED]')
    .replace(/[A-Za-z0-9_-]{28,}/g,'[ID]').replace(/\b\d{7,}\b/g,'[NUMBER]')
    .replace(/\s+/g,' ').trim().slice(0,limit||600);
}
function vtdDiagEvidence_(text, status) {
  const raw=String(text||'').slice(0,100000);
  const title=(raw.match(/<title\b[^>]*>([\s\S]*?)<\/title>/i)||[])[1]||'';
  const plain=vtdDiagRedact_(raw,10000);
  const rules=[
    ['quota_or_rate_limit',/service invoked too many times|quota exceeded|too many requests|rate limit|dịch vụ bị gọi quá nhiều/i],
    ['google_login',/sign in to continue|sign in with your google account|đăng nhập để tiếp tục/i],
    ['access_denied',/access denied|permission denied|you need access|authorization is required/i],
    ['not_found',/the file you have requested does not exist|page not found|404[ .:-]*not found/i],
    ['bad_request',/malformed|400[ .:-]*bad request|invalid request/i],
    ['javascript_reference',/ReferenceError:\s*[A-Za-z_$][\w$]* is not defined/i],
    ['javascript_type',/TypeError:\s*Cannot read properties of (?:undefined|null)|TypeError:\s*[A-Za-z_$][\w.$]* is not a function/i],
    ['json_parse',/SyntaxError:\s*Unexpected token|Unexpected end of JSON input/i],
    ['timeout',/exceeded maximum execution time|execution timed out/i],
    ['server_error',/internal server error|temporarily unavailable/i]
  ];
  let match,category='unknown';for(const rule of rules){const m=plain.match(rule[1]);if(m){category=rule[0];match=m[0];break;}}
  return {htmlTitle:vtdDiagRedact_(title,180),errorExcerpt:match||'',evidenceCategory:category,
    conclusionBasis:match?'response_text':'insufficient_evidence',
    diagnosticConclusion:match?'Phản hồi có thông báo: '+category+'; cần đối chiếu server để xác định nguồn lỗi.':'Chưa xác định nguyên nhân; HTTP '+(status||'không rõ')+' không đủ bằng chứng.',
    evidenceTruncated:String(text||'').length>100000};
}

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
  const d={requestId:id,service:service,method:method,action:safe(body.action || query.action || query._diagAction),requestLength:String(e && e.postData && e.postData.contents || '').length,requestMalformed:malformed,serverDiagnosticsVersion:'3'};
  const log=(level,value)=>{try{console[level](value);}catch(_){}};
  log('log',JSON.stringify(Object.assign({stage:'received'},d)));
  try {
    const response=run();
    d.stage='response_ready';
    let text='';try{text=response && typeof response.getContent==='function' ? response.getContent() : '';}catch(_){d.inspectionFailed=true;}
    d.responseLength=text.length;
    try{const result=JSON.parse(text);d.responseFormat='json';d.serverOk=result.ok===true;d.errorCategory=result.ok===false?apiDiagCategory_(result.message):'';d.protocolPresent=Object.prototype.hasOwnProperty.call(result,'protocol');
      // Function/line frames only. Never store the error message or business values.
      if(result.ok===false){d.errorEvidence=vtdDiagEvidence_(result.message,0);}
      d.stackFrames=(String(result.message || '').match(/\bat [A-Za-z0-9_.$]+ \([^\n)]*:\d+(?::\d+)?\)/g)||[]).slice(0,5).map(x=>x.replace(/[^a-zA-Z0-9_.$(): /-]/g,'')).join(' | ').slice(0,500);
    }catch(_){d.responseFormat=/^\s*</.test(text)?'html':'non_json';}
    return response;
  }catch(err){d.stage='handler_threw';d.errorCategory=apiDiagCategory_(err);d.errorName=safe(err && err.name);d.errorEvidence=vtdDiagEvidence_(String(err),0);d.stackFrames=(String(err && err.stack || '').match(/\bat [A-Za-z0-9_.$]+ \([^\n)]*:\d+(?::\d+)?\)/g)||[]).slice(0,5).map(x=>vtdDiagRedact_(x,120)).join(' | ');throw err;}
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

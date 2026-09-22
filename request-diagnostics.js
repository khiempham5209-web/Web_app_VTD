/* Temporary request tracing. No business bodies, image data or credentials leave the existing APIs. */
(function(){
  'use strict';
  if(window.vtdRequestDiagnostics)return;
  let count=0, seq=0;
  const meta=new WeakMap();
  const emit=(kind,d)=>{try{window.vtdDiagnostics?.record(kind,d);}catch(_){}};
  const id=()=> 'diag_'+Date.now().toString(36)+'_'+(++seq)+'_'+Math.random().toString(36).slice(2,9);
  const safe=v=>String(v||'').replace(/[^a-zA-Z0-9_.-]/g,'').slice(0,80);
  const fetchOld=window.fetch;
  window.fetch=function(url,options){
    let u,payload;
    try{u=new URL(typeof url==='string'?url:url.url);if(u.hostname!=='script.google.com'||!options||typeof options.body!=='string')return fetchOld.apply(this,arguments);payload=JSON.parse(options.body);if(!payload.action||payload.action==='clientDiagnostics')return fetchOld.apply(this,arguments);}catch(_){return fetchOld.apply(this,arguments);}
    const requestId=payload.params?._diagRequestId || id(), start=Date.now();
    let hash=2166136261;for(const c of u.origin+u.pathname)hash=Math.imul(hash^c.charCodeAt(0),16777619);
    const d={requestId,action:safe(payload.action),method:String(options.method||'GET'),transport:'web_fetch',endpointId:(hash>>>0).toString(16),requestHost:u.hostname,requestLength:options.body.length,concurrentRequests:++count};
    payload.params=Object.assign({},payload.params,{_diagRequestId:requestId});
    const opts=Object.assign({},options,{body:JSON.stringify(payload)});
    emit('request_started',d);
    u.searchParams.set('_diagRequestId',requestId);u.searchParams.set('_diagAction',safe(payload.action));
    let pending;try{pending=fetchOld.call(this,u.href,opts);}catch(e){count--;emit('request_network_error',Object.assign({},d,{errorName:safe(e.name),durationMs:Date.now()-start}));throw e;}
    return Promise.resolve(pending).then(r=>{Object.assign(d,{durationMs:Date.now()-start,redirected:!!r.redirected,responseType:r.type||'',retryAfter:r.headers?.get('retry-after')||'',responseRequestId:r.headers?.get('x-request-id')||''});meta.set(r,d);emit('request_headers',Object.assign({},d,{status:r.status,durationMs:Date.now()-start,redirected:r.redirected}));return r;},e=>{emit('request_network_error',Object.assign({},d,{errorName:safe(e.name),reason:e.name==='AbortError'?'aborted_or_timeout':'network',durationMs:Date.now()-start}));throw e;}).finally(()=>count--);
  };
  const readOld=window.readApiJson;
  if(readOld)window.readApiJson=async function(r,label,action){
    const d=Object.assign({},meta.get(r),{label,action,status:r.status,contentType:r.headers?.get('content-type')||''});
    try{d.responseHost=new URL(r.url).hostname;}catch(_){}
    try{return window.parseApiResponse(await r.text(),d);}catch(e){emit('response_read_error',Object.assign({},d,{errorName:safe(e.name)}));throw e;}
  };
  // Attach correlation to native requests without replacing the Android bridge.
  const transportOld=window.apiTransport;
  if(transportOld)window.apiTransport=function(action,params){
    if(!window.VTDNative?.api || action==='uploadImages'||action==='uploadImage')return transportOld.apply(this,arguments);
    const requestId=id(),start=Date.now(),before=new Set(Object.keys(nativeCallbacks));
    const d={requestId,action:safe(action),transport:'native_bridge',method:'POST'};
    const p=transportOld.call(this,action,Object.assign({},params,{_diagRequestId:requestId}));
    for(const key of Object.keys(nativeCallbacks))if(!before.has(key))nativeCallbacks[key].diag=d;
    emit('request_started',d);
    return p.then(result=>{emit('request_finished',Object.assign({},d,{durationMs:Date.now()-start,serverOk:result?.ok===true}));return result;});
  };
  let nativeContext=null;
  const resolveOld=window.__vtdNativeResolve;
  if(resolveOld)window.__vtdNativeResolve=function(key,text){const previous=nativeContext;nativeContext=nativeCallbacks[key]?.diag||null;try{return resolveOld.apply(this,arguments);}finally{nativeContext=previous;}};
  const parseOld=window.parseApiResponse;
  if(parseOld)window.parseApiResponse=function(text,metadata){
    const d=Object.assign({},nativeContext||{},metadata||{});
    const result=parseOld.call(this,text,d);
    let raw;try{raw=JSON.parse(text);}catch(_){}
    const format=!raw?( /^\s*</.test(text)?'html':'non_json'):'json';
    emit('response_inspected',Object.assign({},d,{responseFormat:format,responseLength:String(text).length,serverOk:result?.ok===true,protocolPresent:raw?Object.prototype.hasOwnProperty.call(raw,'protocol'):false,servicePresent:raw?Object.prototype.hasOwnProperty.call(raw,'service'):false,reason:result?.responseCategory || (result?.ok===false?'api_error':'')}));
    return result;
  };
  window.vtdRequestDiagnostics={version:'2'};
  emit('request_diagnostics_loaded',{diagnosticsVersion:'2'});
})();

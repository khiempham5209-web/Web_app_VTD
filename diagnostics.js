/* Read-only diagnostics. No business payloads, images, passwords or tokens in events. */
(function () {
  'use strict';
  if (window.VTDNative || window.vtdDiagnostics) return;
  const key='vtd_diagnostics_v1', originalFetch=window.fetch.bind(window);
  let events=[], busy=false, nextSend=0, loginAt=0, context=null;
  try { events=JSON.parse(sessionStorage.getItem(key)||'[]'); if(!Array.isArray(events))events=[]; } catch (_) {}
  function persist(){try{sessionStorage.setItem(key,JSON.stringify(events));}catch(_){} }
  function reason(e){const s=String(e && (e.message||e.name)||e||'');return /quota/i.test(s)?'quota':/đăng nhập|dang nhap/i.test(s)?'auth_message':/Load failed|fetch|network/i.test(s)?'network':/timeout|quá lâu/i.test(s)?'timeout':'other';}
  function snapshot(){let bytes=0;try{for(let i=0;i<localStorage.length;i++){const k=localStorage.key(i);bytes+=2*(k.length+(localStorage.getItem(k)||'').length);}}catch(_){}return {localStorageBytes:bytes,online:navigator.onLine,visible:!document.hidden,sessionAgeMs:loginAt?Date.now()-loginAt:null};}
  function record(kind,detail){try{events.push({id:Date.now().toString(36)+'_'+Math.random().toString(36).slice(2),time:new Date().toISOString(),kind,user:typeof currentEmail==='function'?currentEmail():'',device:typeof deviceId==='function'?deviceId():'',version:typeof APP_VERSION==='string'?APP_VERSION:'',detail:Object.assign(snapshot(),detail||{})});events=events.slice(-80);persist();schedule();}catch(_){} }
  let timer;
  function schedule(){if(!timer)timer=setTimeout(()=>{timer=null;flush();},2000);}
  async function flush(){if(busy||!events.length||Date.now()<nextSend||!navigator.onLine||typeof live==='undefined'||!live.token)return;busy=true;const batch=events.slice(0,20);const controller=new AbortController();const timeout=setTimeout(()=>controller.abort(),20000);try{const r=await originalFetch(activeApiUrl(),{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action:'clientDiagnostics',params:{sessionToken:live.token,events:batch}}),signal:controller.signal});const res=await r.json();if(res.ok){const ids=new Set(batch.map(e=>e.id));events=events.filter(e=>!ids.has(e.id));persist();nextSend=0;}else nextSend=Date.now()+60000;}catch(_){nextSend=Date.now()+60000;}finally{clearTimeout(timeout);busy=false;}}
  function wrapAsync(name,kind){const fn=window[name];if(typeof fn!=='function')return;window[name]=async function(){try{return await fn.apply(this,arguments);}catch(e){record(kind,{reason:reason(e),errorName:String(e&&e.name||'Error').slice(0,60)});throw e;}};}
  // Preserve synchronous return values and thrown errors.
  const save=window.saveLocalQueue;if(save)window.saveLocalQueue=function(q){try{return save.apply(this,arguments);}catch(e){record('queue_write_failed',{reason:reason(e),queueCount:Array.isArray(q)?q.length:0});throw e;}};
  wrapAsync('queueImagePut','image_write_failed');wrapAsync('queueImageStore','image_transaction_failed');
  wrapAsync('loginDemo','login_client_failed');
  const transport=window.apiTransport;if(transport)window.apiTransport=function(action,params){const sent=live.token;const start=Date.now();return transport.apply(this,arguments).then(res=>{if(action==='login'&&res&&res.ok){loginAt=Date.now();record('login_response_ok');}else if(res&&res.ok===false)record('api_failed',{action:String(action).slice(0,60),reason:reason(res.message),durationMs:Date.now()-start,sameSession:sent===live.token});return res;});};
  const handler=window.handleApiResponse;if(handler)window.handleApiResponse=function(action,res){const previous=context;context={action:String(action).slice(0,60),reason:reason(res&&res.message)};try{return handler.apply(this,arguments);}finally{context=previous;}};
  const out=window.logout;if(out)window.logout=function(){record('manual_logout');return out.apply(this,arguments);};
  const show=window.showLogin;if(show)window.showLogin=function(){record('login_screen',context||{reason:'day_boot_or_other'});return show.apply(this,arguments);};
  window.addEventListener('error',e=>record('javascript_error',{reason:reason(e.error),errorName:String(e.error&&e.error.name||'Error').slice(0,60),line:e.lineno||0}));
  window.addEventListener('unhandledrejection',e=>record('unhandled_rejection',{reason:reason(e.reason),errorName:String(e.reason&&e.reason.name||'Error').slice(0,60)}));
  window.addEventListener('online',()=>{nextSend=0;flush();});
  window.addEventListener('pageshow',e=>record('page_show',{restored:!!e.persisted}));
  window.addEventListener('pagehide',()=>{record('page_hide');persist();});
  setInterval(flush,60000);
  window.vtdDiagnostics={flush,pending:()=>events.length};
  record('diagnostics_loaded',{diagnosticsVersion:'1'});
})();

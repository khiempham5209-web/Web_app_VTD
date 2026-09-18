const fs=require('fs'),vm=require('vm'),assert=require('assert/strict');
const source=fs.readFileSync(process.argv[2],'utf8');
const events=[];let calls=0;
const response={status:404,type:'cors',redirected:true,url:'https://script.googleusercontent.com/macros/echo?secret=TOKEN',headers:{get:()=> 'text/html'},text:async()=>'<html>404 not found SECRET CUSTOMER TOKEN</html>'};
const w={fetch:async()=>{calls++;return response},readApiJson:()=>{},parseApiResponse:(t)=>{try{return JSON.parse(t)}catch{return {ok:false,code:'INVALID_API_RESPONSE',responseCategory:'not_found'}}},vtdDiagnostics:{record:(kind,detail)=>events.push({kind,detail})}};
vm.runInNewContext(source,{window:w,URL,WeakMap,Date,Math,Promise});
(async()=>{
const r=await w.fetch('https://script.google.com/macros/s/DEPLOY/exec?token=TOKEN',{method:'POST',body:'SECRET'});
const result=await w.readApiJson(r,'PN','skuSync');
assert.equal(result.code,'INVALID_API_RESPONSE');assert.equal(calls,1);
const e=events.find(e=>e.kind==='json_diagnostic').detail;
assert.equal(e.status,404);assert.equal(e.redirected,true);assert.equal(e.pageCategory,'not_found');assert.equal(e.concurrentRequests,1);
assert(!JSON.stringify(events).match(/SECRET|CUSTOMER|TOKEN|DEPLOY/));
const good={ok:true,protocol:1,records:[]};assert.deepEqual(w.parseApiResponse(JSON.stringify(good),{action:'skuSync'}),good);
w.parseApiResponse('{"ok":true,"service":"PRIVATE"}',{action:'skuSync'});
assert.equal(events.at(-1).detail.pageCategory,'json_protocol_mismatch');assert.equal(events.at(-1).detail.servicePresent,true);assert(!JSON.stringify(events).includes('PRIVATE'));
console.log('PASS: response unchanged, no extra calls, HTML classification, protocol mismatch, sensitive data excluded');
})().catch(e=>{console.error(e);process.exitCode=1});

const fs=require('fs'),vm=require('vm'),assert=require('assert/strict');
const app=fs.readFileSync('index.html','utf8'),pn=fs.readFileSync('ChungTuFF_WebApi.gs','utf8'),vtd=fs.readFileSync('Code_chinh_VTD.gs','utf8');
function extract(name){const re=new RegExp('    (?:async )?function '+name+'\\(');const start=app.search(re);if(start<0)throw Error(name);return app.slice(start,app.indexOf('\n    }',start)+6);}
let passed=0;function test(name,fn){return Promise.resolve().then(fn).then(()=>{passed++;console.log('PASS '+name);});}
const context={console,Date,Math,Set,Map};vm.createContext(context);vm.runInContext(pn,context);
const c={docOpsState:{productNotes:[]},docops:()=>{},PN_PRODUCT_NOTES:['Bình thường','Hàng cận date','Xẹp hơi','Hết hạn sử dụng','Vỏ bị xé/hở','Khác']};vm.createContext(c);
for(const n of ['pnSyncKey','keySyncDiff','pnExpiryVisible','pnValidExpiry','togglePnNote','docOpsProductNoteValue','validateDocOpsProduct'])vm.runInContext(extract(n),c);
function sheet(data){return {getLastRow:()=>data.length,getLastColumn:()=>data[0].length,getName:()=> 'Chứng từ_FF',getDataRange:()=>({getDisplayValues:()=>data}),getRange:(r,col,rows=1,cols=1)=>({getDisplayValues:()=>data.slice(r-1,r-1+rows).map(row=>row.slice(col-1,col-1+cols))})};}
(async()=>{
 await test('Syntax: both APIs and all inline scripts',()=>{new vm.Script(pn);new vm.Script(vtd);for(const m of app.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g))new vm.Script(m[1]);});
 const headers=Array.from({length:30},(_,i)=>'extra '+i);headers[28]='Số đơn hàng';headers[26]='Mã đơn GHTK';headers[25]='Khách hàng';
 const row=Array(30).fill('');row[28]='2913101';row[26]='1680000000';row[25]='Khách hàng thử';
 const data=[headers,row],sh=sheet(data);context.sheet_=()=>sh;
 let manifest;
 await test('PN reads reordered headers beyond column W',()=>{manifest=context.apiKeySync_({});assert.equal(manifest.entries[0].key,'ORDER:2913101');assert.equal(context.apiPage_({}).records[0].orderNo,'2913101');assert.equal(context.apiLookup_({query:'1680000000'}).record.orderNo,'2913101');});
 await test('Direct lookup/page fingerprint equals manifest',()=>{assert.equal(context.apiPage_({}).records[0].contentVersion,manifest.entries[0].version);});
 await test('Manual GHTK edit invalidates source version without apiSave',()=>{row[26]='1699999999';assert.notEqual(context.apiKeySync_({}).version,manifest.version);assert.equal(context.apiKeySync_({version:manifest.version,keys:['ORDER:2913101']}).changed,true);});
 await test('Return order without order number uses GHTK text including dot',()=>{row[28]='';row[26]='TH090926.1';assert.equal(context.apiKeySync_({}).entries[0].key,'GHTK:th090926.1');});
 await test('Blank keys are reported; blank rows excluded',()=>{row[26]='';data.push(Array(30).fill(''));const r=context.apiKeySync_({});assert.equal(r.invalidCount,1);assert.equal(r.total,0);data.pop();});
 await test('Duplicate keys detected rather than silently overwritten',()=>{row[28]='2913101';data.push(row.slice());assert.equal(context.apiKeySync_({}).duplicateCount,1);data.pop();});
 await test('2000 cached + 30 new downloads only 30; row move is mapping only',()=>{const rows=Array.from({length:2000},(_,i)=>({syncKey:'ORDER:'+i,contentVersion:'v1',rowNumber:i+2}));const entries=Array.from({length:2030},(_,i)=>({key:'ORDER:'+i,version:'v1',rowNumber:i+3}));const d=c.keySyncDiff(entries,rows);assert.equal(d.missing.length,30);assert.equal(d.changed.length,0);assert.equal(d.mapped.length,2000);entries[5].version='v2';assert.equal(c.keySyncDiff(entries,rows).changed.length,1);});
 await test('Bình thường is exclusive; abnormal notes can mix',()=>{c.togglePnNote('Bình thường',true);c.togglePnNote('Xẹp hơi',true);c.togglePnNote('Hết hạn sử dụng',true);assert.equal(c.docOpsState.productNotes.join(';'),'Xẹp hơi;Hết hạn sử dụng');assert.equal(c.pnExpiryVisible(),true);c.togglePnNote('Bình thường',true);assert.equal(c.docOpsState.productNotes.join(';'),'Bình thường');assert.equal(c.pnExpiryVisible(),false);});
 await test('Expiry optional; real DD/MM/YYYY validation including leap years',()=>{for(const x of ['', '29/02/2024','10/09/2026'])assert.equal(c.pnValidExpiry(x),true);for(const x of ['29/02/2025','31/04/2026','1/9/2026','2026-09-10'])assert.equal(c.pnValidExpiry(x),false);});
 const valid={name:'Test',barcode:'123',material:'M1',status:'Bình thường',quantity:1,note:'Xẹp hơi',classification:'Hoàn trả',notes:['Xẹp hơi'],lengthCm:1,widthCm:1,heightCm:1,weightGram:1,images:[{}],expiryDate:''};
 await test('Required classification, optional expiry, Other text validation',()=>{assert.equal(c.validateDocOpsProduct(valid,'SKU 1'),'');assert.match(c.validateDocOpsProduct({...valid,classification:''},'SKU 1'),/Phân loại/);assert.match(c.validateDocOpsProduct({...valid,notes:['Khác'],otherNote:''},'SKU 1'),/Khác/);});
 await test('API normalizes and saves new product fields, preserves legacy compatibility',()=>{const item=context.normalizeProductItems_([{...valid,productSchema:2,notes:['Hết hạn sử dụng','Khác'],otherNote:'rách',expiryDate:'10/09/2026'}])[0];assert.equal(context.validateProductItem_(item,1),'');assert.equal(item.note,'Hết hạn sử dụng; Khác: rách');assert.equal(item.classification,'Hoàn trả');assert.equal(item.expiryDate,'10/09/2026');const old=context.normalizeProductItems_([{...valid,classification:'',productSchema:0}])[0];assert.equal(context.validateProductItem_(old,1),'');});
 await test('SKU API handles >5000 codes, coalesces identical rows and detects conflicts',()=>{
  const codes=[['Tên sản phẩm','Loại','Mã vật tư'],...Array.from({length:5030},(_,i)=>['Name '+i,'Type','M'+i])];const ctx={console};vm.createContext(ctx);vm.runInContext(vtd,ctx);ctx.vtdApp_requireAction_=()=>null;ctx.vtdApp_ss_=()=>({getId:()=> 'test',getSheetByName:()=>sheet(codes)});
  const first=ctx.vtdApp_masterSkuKeySync({});assert.equal(first.total,5030);const last=ctx.vtdApp_masterSkuKeySync({offset:5000,version:first.version});assert.equal(last.entries.length,30);codes[1][0]='Changed';assert.notEqual(ctx.vtdApp_masterSkuKeySync({}).version,first.version);codes.push(codes[1].slice());assert.equal(ctx.vtdApp_masterSkuKeySync({}).duplicateCount,0);assert.equal(ctx.vtdApp_masterSkuKeySync({}).total,5030);codes.push(['Conflicting name','Type',codes[1][2]]);assert.equal(ctx.vtdApp_masterSkuKeySync({}).duplicateCount,1);
 });
 await test('Stale lookup response cannot leave loading stuck or replace scanned order',async()=>{
  const state={};let finish;const ctx={docOpsState:state,renderDocOpsLookup:()=>{},docOpsFindCached:async q=>q==='2913101'?{orderNo:q}:null,docOpsApi:()=>new Promise(r=>finish=r),docOpsRememberRecords:async()=>{},keySyncSchedule:()=>{},keySyncProgress:async()=>{}};vm.createContext(ctx);vm.runInContext(extract('lookupDocOpsOrder'),ctx);const first=ctx.lookupDocOpsOrder('11111');await new Promise(setImmediate);await ctx.lookupDocOpsOrder('2913101');finish({ok:true,record:{orderNo:'11111'}});await first;assert.equal(state.loading,false);assert.equal(state.record.orderNo,'2913101');
 });
 await test('Reconcile commit failure cannot report done',async()=>{
  const progress=[];const ctx={Map,Set,keySyncProgress:async(t,p)=>progress.push(p),keySyncManifest:async()=>({total:1,version:'v',entries:[{key:'ORDER:1',version:'a',rowNumber:2}],invalidCount:0,duplicateCount:0}),keySyncRead:async()=>[],keySyncDiff:c.keySyncDiff,keySyncCall:async()=>({records:[{syncKey:'ORDER:1',orderNo:'1',contentVersion:'a'}]}),pnSyncKey:c.pnSyncKey,keySyncWrite:async()=>{throw Error('QuotaExceededError');}};vm.createContext(ctx);vm.runInContext(extract('keySyncReconcile'),ctx);await assert.rejects(()=>ctx.keySyncReconcile('docOpsLookup'),/Quota/);assert.equal(progress.some(p=>p.status==='done'),false);
 });
 await test('Complete reconcile downloads only 30 new records, commits, and next run downloads zero',async()=>{
   const old=Array.from({length:2000},(_,i)=>({syncKey:'ORDER:'+i,orderNo:String(i),rowNumber:i+2,contentVersion:'v1'}));
   const records=[...old,...Array.from({length:30},(_,i)=>({syncKey:'ORDER:'+(2000+i),orderNo:String(2000+i),rowNumber:2002+i,contentVersion:'v1'}))];
   const snapshot=context.ks_pack(records,[],'test-pn',2031), stores={records:new Map(old.map(r=>[r.syncKey,{...r}]))}, reports=[], requested=[];
   const ctx={Map,Set,Date,docOpsLookupMemory:[],keySyncDiff:c.keySyncDiff,pnSyncKey:c.pnSyncKey,
     keySyncRead:async name=>Array.from(stores[name].values()),
     keySyncWrite:async(name,rows,deletes=[])=>{for(const key of deletes)stores[name].delete(key);for(const row of rows)stores[name].set(row.syncKey,{...row});},
     keySyncCall:async(type,params)=>{if(params.keys)requested.push(...params.keys);const res=context.ks_reply(snapshot,params);if(!res.ok)throw Error(res.message);return res;},
     keySyncProgress:async(type,p)=>{reports.push(p);return p;},docOpsLookupWriteMeta:async()=>{},keySyncPublishProducts:async()=>{}};
   vm.createContext(ctx);vm.runInContext(extract('keySyncManifest')+'\n'+extract('keySyncReconcile'),ctx);
   assert.equal(await ctx.keySyncReconcile('docOpsLookup'),true);assert.equal(requested.length,30);assert.equal(stores.records.size,2030);assert.equal(reports.at(-1).status,'done');
   requested.length=0;assert.equal(await ctx.keySyncReconcile('docOpsLookup'),true);assert.equal(requested.length,0);
   records[0].contentVersion='v2';records[0].maDonGhtk='added-later';Object.assign(snapshot,context.ks_pack(records,[],'test-pn',2031));
   assert.equal(await ctx.keySyncReconcile('docOpsLookup'),true);assert.equal(requested.length,1);assert.equal(stores.records.get('ORDER:0').maDonGhtk,'added-later');
 });
 await test('Web camera appends same-named iPhone photos in every image bucket; cancel keeps photos',async()=>{
  const ctx={live:{},renderFilePreview:()=>{},toast:()=>{},readFileAsDataUrl:async f=>'data:image/jpeg;base64,'+f.bytes,
    splitDataUrl:url=>({mimeType:'image/jpeg',base64:url.split(',')[1]}),nativeImagePickerMap:{camera:['camera','docOpsProductFiles','preview']},openFilePicker:()=>{}};
  vm.createContext(ctx);vm.runInContext(extract('captureFiles')+'\n'+extract('retakeFiles'),ctx);
  for(const bucket of ['docOpsProductFiles','docOpsFiles','productFiles','docFiles','returnThFiles']){
    await ctx.captureFiles({files:[{name:'image.jpg',bytes:'first'}],id:'camera'},bucket,'preview');
    await ctx.captureFiles({files:[{name:'image.jpg',bytes:'second'}],id:'camera'},bucket,'preview');
    assert.equal(ctx.live[bucket].length,2);assert.equal(ctx.live[bucket][0].base64,'first');assert.equal(ctx.live[bucket][1].base64,'second');
    ctx.retakeFiles(bucket,'preview','camera');await ctx.captureFiles({files:[],id:'camera'},bucket,'preview');assert.equal(ctx.live[bucket].length,2);
  }
 });
 await test('Photo read error retains existing photos; cleared draft rejects late photo',async()=>{
   let finish;const ctx={live:{docFiles:[{name:'existing'}]},renderFilePreview:()=>{},toast:()=>{},readFileAsDataUrl:async()=>{throw Error('read error');},splitDataUrl:()=>({mimeType:'image/jpeg',base64:'abc'})};
   vm.createContext(ctx);vm.runInContext(extract('captureFiles'),ctx);
   await ctx.captureFiles({files:[{}],id:'camera'},'docFiles','preview');assert.equal(ctx.live.docFiles.length,1);assert.equal(ctx.live.capturePending.docFiles,0);
   ctx.readFileAsDataUrl=()=>new Promise(r=>finish=r);const pending=ctx.captureFiles({files:[{}],id:'camera'},'docFiles','preview');await new Promise(setImmediate);ctx.live.docFiles=[];finish('data:image/jpeg;base64,abc');await pending;assert.equal(ctx.live.docFiles.length,0);
 });
 await test('Notes remain a compact dropdown with multiple selected options, not inline checkboxes',()=>{
   const ctx={docOpsState:{productNotes:['Xẹp hơi','Hết hạn sử dụng'],notesOpen:true},PN_PRODUCT_NOTES:c.PN_PRODUCT_NOTES,html:x=>String(x)};vm.createContext(ctx);vm.runInContext(extract('pnNotesHtml'),ctx);
   const markup=ctx.pnNotesHtml();assert.equal(markup.includes('type="checkbox"'),false);assert.ok(markup.includes('aria-multiselectable="true"'));assert.equal((markup.match(/aria-selected="true"/g)||[]).length,2);
 });
 await test('Cache email sends every received error report and retries failed sends',()=>{
   let count=0, fail=false;const props=new Map();const ctx={console,Date,Session:{getEffectiveUser:()=>({getEmail:()=> 'fulfillment.wms.3pl@gmail.com'})},LockService:{getScriptLock:()=>({tryLock:()=>true,releaseLock:()=>{}})},PropertiesService:{getScriptProperties:()=>({getProperty:k=>props.get(k),setProperty:(k,v)=>props.set(k,v)})},Utilities:{formatDate:()=> 'today'},MailApp:{sendEmail:m=>{assert.equal(m.to,'khiempham5209@gmail.com');if(fail)throw Error('mail unavailable');count++;}}};vm.createContext(ctx);vm.runInContext(vtd,ctx);
   const p={status:'error',cacheType:'masterSku',lastError:'duplicate',email:'worker',deviceId:'one'};
   ctx.vtdApp_cacheErrorEmail_({...p,status:'done'});assert.equal(count,0);
   ctx.vtdApp_cacheErrorEmail_(p);ctx.vtdApp_cacheErrorEmail_({...p,deviceId:'two'});assert.equal(count,2);
   fail=true;assert.throws(()=>ctx.vtdApp_cacheErrorEmail_({...p,lastError:'network'}));fail=false;ctx.vtdApp_cacheErrorEmail_({...p,lastError:'network'});assert.equal(count,3);
 });
 await test('PN SKU delta by material coalesces duplicates and sends zero unchanged records',()=>{
   const ctx={console};vm.createContext(ctx);vm.runInContext(pn,ctx);
   const data=[['Mã vật tư','Tên vật tư','Barcode'],['M1','Name','123'],['M1','Name','123']];ctx.skuSheet_=()=>sheet(data);
   const first=ctx.apiSkuSync_({});assert.equal(first.total,1);assert.equal(first.records.length,1);
   const versions=Object.fromEntries(first.records.map(r=>[r.syncKey,r.contentVersion]));assert.equal(ctx.apiSkuSync_({versions}).records.length,0);
   data.push(['M2','New','456']);assert.equal(ctx.apiSkuSync_({versions}).records.length,1);
   data.push(['M1','Conflict','999']);assert.equal(ctx.apiSkuSync_({}).ok,false);
 });
 await test('PN SKU timeout and commit failure retain old in-memory cache',async()=>{
   const old=[{syncKey:'MATERIAL:m1',material:'M1',name:'Old',contentVersion:'a',rowNumber:2}];let failCommit=false;
   const ctx={Map,Set,Object,Date,docOpsState:{},docOpsSkuMemory:old.slice(),docOpsApiUrl:()=> 'api',hydrateDocOpsSkuCache:async()=>{},keySyncRead:async store=>store==='pnSku'?old:undefined,keySyncState:{progress:{}},cacheDailyClaim:async()=>true,keySyncProgress:async()=>{},view:'home',docOpsApi:async()=>{if(!failCommit)throw Error('timeout');return {ok:true,protocol:1,total:1,entries:[{key:'MATERIAL:m1',version:'b',rowNumber:3}],records:[{syncKey:'MATERIAL:m1',material:'M1',name:'New',contentVersion:'b'}]};},keySyncWrite:async()=>{throw Error('quota');}};
   vm.createContext(ctx);vm.runInContext(extract('loadDocOpsSkuCache'),ctx);
   assert.equal(await ctx.loadDocOpsSkuCache(),false);assert.equal(ctx.docOpsSkuMemory[0].name,'Old');
   failCommit=true;assert.equal(await ctx.loadDocOpsSkuCache(),false);assert.equal(ctx.docOpsSkuMemory[0].name,'Old');
 });
 await test('Daily claim persists across reload and failures; manual run and next day allowed',async()=>{
   const rows=new Map();let day='11/09/2026';
   const db={transaction:()=>{const tx={};tx.objectStore=()=>({get:key=>{const req={};setImmediate(()=>{req.result=rows.get(key);req.onsuccess();setImmediate(()=>tx.oncomplete());});return req;},put:row=>rows.set(row.key,row)});return tx;}};
   const ctx={keySyncDb:async()=>db,todayApiDateText:()=>day};vm.createContext(ctx);vm.runInContext(extract('cacheDailyClaim'),ctx);
   assert.equal(await ctx.cacheDailyClaim('docOpsSku'),true);assert.equal(await ctx.cacheDailyClaim('docOpsSku'),false);
   vm.runInContext(extract('cacheDailyClaim'),ctx);assert.equal(await ctx.cacheDailyClaim('docOpsSku'),false);
   assert.equal(await ctx.cacheDailyClaim('docOpsSku',true),true);day='12/09/2026';assert.equal(await ctx.cacheDailyClaim('docOpsSku'),true);
 });
 await test('Update config saves server first and always reports errors without clearing input',async()=>{
   const nodes={updateLatestVersion:{value:'5.9.3'},updateApkUrl:{value:'https://example.test/app.apk'},updateFolderUrl:{value:''},updateNote:{value:'New'},saveUpdateConfigButton:{},updateSaveStatus:{style:{}}};let calls=0,mode='success';const messages=[];
   const ctx={live:{},appConfig:{update:{latestVersion:'5.8'}},appUpdateConfig:()=>ctx.appConfig.update,canAdminConfig:()=>true,el:id=>nodes[id],toast:m=>messages.push(m),api:async()=>{calls++;if(mode==='throw')throw Error('network');if(mode==='denied')return {ok:false,message:'Denied'};return {ok:true};},mergeSystemConfig:()=>{},saveAppConfig:()=>{if(mode==='quota')throw Error('QuotaExceeded');}};vm.createContext(ctx);vm.runInContext(extract('saveUpdateConfig'),ctx);
   mode='quota';await ctx.saveUpdateConfig();assert.equal(calls,1);assert.match(messages.at(-1),/Đã lưu lên hệ thống/);assert.equal(nodes.saveUpdateConfigButton.disabled,false);
   ctx.appConfig.update={latestVersion:'5.8'};mode='denied';await ctx.saveUpdateConfig();assert.equal(ctx.appConfig.update.latestVersion,'5.8');assert.match(messages.at(-1),/Denied/);
   mode='throw';await ctx.saveUpdateConfig();assert.match(messages.at(-1),/network/);assert.equal(ctx.live.updateConfigSaving,false);
   mode='success';await ctx.saveUpdateConfig();assert.match(messages.at(-1),/Đã lưu cấu hình cập nhật lên hệ thống/);
 });
 await test('Staff login returns update, theme and PN endpoint without admin permission',()=>{
   const ctx={console,Logger:{log:()=>{}},Utilities:{getUuid:()=> 'test-token'},CacheService:{getScriptCache:()=>({get:()=>null})}};vm.createContext(ctx);vm.runInContext(vtd,ctx);
   ctx.vtdApp_permissionConfig_=()=>({pins:{'staff@test':'testpass'},adminEmails:[],users:{'staff@test':{allowed:true,role:'staff'}}});
   ctx.vtdApp_markPasswordFirstLogin_=()=>({});ctx.vtdApp_recordLogin_=()=>{};ctx.vtdApp_sessionPut_=()=>{};ctx.vtdApp_sessionGet_=()=>({});ctx.vtdApp_vnDayKey_=()=> 'day';ctx.vtdApp_debugLog_=()=>{};ctx.vtdApp_authShape_=()=>({email:'staff@test',role:'staff'});
   ctx.vtdApp_publicSystemConfig_=()=>({update:{latestVersion:'5.9.3'},themeConfig:{version:'new'},docOpsApiUrl:'pn-endpoint'});
   const res=ctx.vtdApp_login({email:'staff@test',pin:'testpass'});assert.equal(res.ok,true);assert.equal(res.config.system.update.latestVersion,'5.9.3');assert.equal(res.config.system.themeConfig.version,'new');assert.equal(res.config.system.docOpsApiUrl,'pn-endpoint');
 });
 await test('Background starts Master SKU before PN and claims daily slot after idle',async()=>{
   const events=[];const ctx={Date,Math,keySyncState:{},live:{token:'t'},navigator:{onLine:true},document:{hidden:false},docOpsApiUrl:()=> 'pn',keySyncRead:async()=>null,todayApiDateText:()=> 'day',keySyncProgress:async(t)=>events.push('report:'+t),keySyncIdle:async()=>events.push('idle'),cacheDailyClaim:async t=>{events.push('claim:'+t);return true;},keySyncReconcile:async t=>events.push('sync:'+t),loadDocOpsSkuCache:async()=>events.push('pnSku'),keySyncSchedule:()=>{}};vm.createContext(ctx);vm.runInContext(extract('keySyncBackground'),ctx);await ctx.keySyncBackground();assert.equal(events[0],'report:masterSku');assert.ok(events.indexOf('idle')<events.indexOf('claim:masterSku'));assert.ok(events.indexOf('sync:masterSku')<events.indexOf('pnSku'));
 });
 await test('Idle ignores unrelated background requests and times out on blocked UI',async()=>{
   let now=0;const ctx={Date:{now:()=>now},keySyncState:{pending:9,lastRequest:0,lastInput:-10000},live:{token:'t',loadingCount:0,syncRunning:false},docOpsState:{},navigator:{onLine:true},document:{hidden:false},setTimeout:fn=>{now+=91000;fn();}};vm.createContext(ctx);vm.runInContext(extract('keySyncIdle'),ctx);await ctx.keySyncIdle();assert.equal(now,0);ctx.live.loadingCount=1;await assert.rejects(()=>ctx.keySyncIdle(),/90/);
 });
 await test('Staff cache status is visible; manual download requires explicit permission',()=>{
   let allowed=false;const ctx={canAdminConfig:()=>false,canAction:a=>allowed && a==='downloadCache',keyedProgress:()=>({status:'error',lastError:'Test failure'}),cacheStatusLabel:x=>x,html:x=>String(x)};vm.createContext(ctx);vm.runInContext(extract('ownCachePanelHtml'),ctx);assert.ok(ctx.ownCachePanelHtml().includes('Test failure'));assert.equal(ctx.ownCachePanelHtml().includes('onclick="downloadOwnCache()"'),false);allowed=true;assert.ok(ctx.ownCachePanelHtml().includes('onclick="downloadOwnCache()"'));
 });
 await test('Repeated scheduling cannot postpone an already earlier cache start',()=>{
   const timers=[];const ctx={Date:{now:()=>1000},keySyncState:{},setTimeout:(fn,ms)=>{timers.push(ms);return timers.length;},clearTimeout:()=>{},keySyncBackground:()=>{}};vm.createContext(ctx);vm.runInContext(extract('keySyncSchedule'),ctx);ctx.keySyncSchedule(1000);ctx.keySyncSchedule(30000);assert.equal(timers.length,1);ctx.keySyncSchedule(100);assert.equal(timers.length,2);
 });
 await test('First real master download replaces sample products even while input is open',async()=>{
   let applied=false;const rows=[{code:'real'}];const ctx={keySyncRead:async()=>rows,view:'input',keySyncState:{active:true},live:{},products:[{code:'sample'}],applyKeyedProducts:()=>{applied=true;}};vm.createContext(ctx);vm.runInContext(extract('keySyncPublishProducts'),ctx);await ctx.keySyncPublishProducts();assert.equal(applied,true);ctx.live.masterSkuPublished=true;applied=false;await ctx.keySyncPublishProducts();assert.equal(applied,false);assert.equal(ctx.live.keyedProductsPending,rows);
 });
 console.log('TOTAL '+passed+' tests passed');
})().catch(err=>{console.error(err);process.exitCode=1;});

const http=require('http'),fs=require('fs');
const harness=`
    (async function qaBoot(){
      apiTransport=async()=>({ok:true});docOpsApiTransport=async()=>({ok:false,message:'QA offline'});
      el('login').classList.add('hidden');el('app').classList.remove('hidden');
      live.user={email:'qa@local.test',role:'Admin'};view='docops';
      docOpsState.skuLoaded=true;docOpsState.returnType='Sản phẩm';
      docOpsState.selectedSku={material:'TEST-M1',barcode:'123456789',name:'SKU kiểm thử',lengthCm:'10',widthCm:'10',heightCm:'10',weightGram:'100'};
      docOpsState.productSearch='SKU kiểm thử';docOpsState.lengthCm='10';docOpsState.widthCm='10';docOpsState.heightCm='10';docOpsState.weightGram='100';
      docops();el('title').textContent='Kiểm thử PN 5.9 — không kết nối Sheet';
      const result=document.createElement('pre');result.id='qa-result';result.style='white-space:pre-wrap;padding:10px;background:#eaf8ed';el('app').prepend(result);
      const log=[];const check=(ok,label)=>{if(!ok)throw Error(label);log.push('PASS '+label);result.textContent=log.join('\\n');};
      try {
        await keySyncWrite('records',[{syncKey:'ORDER:QA',orderNo:'QA',rowNumber:2}]);
        check((await keySyncRead('records','ORDER:QA')).orderNo==='QA','IndexedDB ghi và đọc theo khóa đơn');
        await keySyncWrite('records',[],['ORDER:ROLLBACK']);let rejected=false;try{await keySyncWrite('records',[{syncKey:'ORDER:ROLLBACK',orderNo:'ROLLBACK'},{}]);}catch(err){rejected=true;}
        await new Promise(r=>setTimeout(r,20));
        check(rejected && !(await keySyncRead('records','ORDER:ROLLBACK')),'Lỗi ghi phải rollback cả giao dịch');
        await keySyncWrite('records',[{syncKey:'ORDER:QA',orderNo:'QA',rowNumber:30,maDonGhtk:'GHTK-QA'}]);
        check((await keySyncRead('records')).filter(r=>r.syncKey==='ORDER:QA').length===1,'Đổi row và bổ sung GHTK không nhân đôi đơn');
        const draft=docOpsProductDraftItem();draft.quantity='1';draft.status='Bình thường';draft.classification='Hàng thu hồi';draft.notes=['Hết hạn sử dụng','Khác'];draft.otherNote='Kiểm thử';draft.note='Hết hạn sử dụng; Khác: Kiểm thử';draft.expiryDate='10/09/2026';draft.images=[{id:'qa-image'}];
        await keySyncWrite('meta',[{key:'qa-offline-product',item:draft}]);const saved=(await keySyncRead('meta','qa-offline-product')).item;
        check(saved.classification==='Hàng thu hồi' && saved.expiryDate==='10/09/2026' && saved.notes.length===2,'Các trường SKU giữ nguyên qua lưu DB');
        check(!validateDocOpsProduct(saved,'SKU 1'),'SKU mới hợp lệ với ngày và ghi chú nhiều lựa chọn');
        await new Promise((resolve,reject)=>{
          const req=indexedDB.open(DOC_OPS_LOOKUP_DB,1);
          req.onupgradeneeded=()=>{req.result.createObjectStore('records',{keyPath:'rowNumber'});req.result.createObjectStore('meta',{keyPath:'key'});};
          req.onerror=()=>reject(req.error);
          req.onsuccess=()=>{const db=req.result,tx=db.transaction(['records','meta'],'readwrite');tx.objectStore('records').put({rowNumber:8,orderNo:'QA-MIGRATION',maDonGhtk:'QA-GHTK'});tx.objectStore('meta').put({key:'sync',lastRow:10});tx.oncomplete=()=>{db.close();resolve();};tx.onerror=()=>reject(tx.error);};
        });
        await keySyncWrite('meta',[],['sync']);await hydrateDocOpsLookupCache();
        check(!!(await keySyncRead('records','ORDER:qa-migration')) && docOpsLookupMetaMemory.lastRow===10,'Chuyển cache row 5.8 sang khóa đơn, giữ mốc tải thêm');
      }catch(err){result.textContent=log.join('\\n')+'\\nFAIL '+err.message;console.error(err);}
    })();
`;
http.createServer((req,res)=>{
 if(req.url==='/favicon.ico'){res.writeHead(204);res.end();return;}
 let html=fs.readFileSync('index.html','utf8').replace('setInterval(enforceSession, 60000);','/* QA: session expiry disabled on isolated localhost fixture */');const start=html.indexOf('    (function bootLive(){'),end=html.indexOf('    window.addEventListener("pageshow"',start);
 html=html.slice(0,start)+harness+html.slice(end);
 res.writeHead(200,{'Content-Type':'text/html; charset=utf-8','Content-Security-Policy':"default-src 'self' data: blob:; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; connect-src 'none'; img-src 'self' data: blob:"});res.end(html);
}).listen(8599,'127.0.0.1',()=>console.log('QA server http://127.0.0.1:8599 — external requests blocked'));

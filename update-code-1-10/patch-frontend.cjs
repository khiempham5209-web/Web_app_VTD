const fs=require('fs'),path=require('path'),crypto=require('crypto');
const root=__dirname,source='originals/frontend-index.html';
let html=fs.readFileSync(path.join(root,source),'utf8');
const original=html;
function patch(a,b){if(!html.includes(a))throw Error('Missing frontend anchor: '+a.slice(0,100));html=html.replace(a,b);}
patch('      const cleanPayload = Object.assign({}, payload || {}, {\n        maDon,',
 '      const cleanPayload = Object.assign({}, payload || {}, {\n        syncKey: record.syncKey || pnSyncKey(record),\n        orderNo: record.orderNo || record.soDonHang || "",\n        maDonGhtk: record.maDonGhtk || "",\n        maDon,');
patch('        const remoteLastRow = Math.max(1, Number(manifest.lastRow || 1));',
 '        const sourceEpoch = String(manifest.sourceEpoch || "legacy");\n'+
 '        if (sourceEpoch === "pn-full-v1" || sourceEpoch !== String(docOpsLookupMetaMemory.sourceEpoch || "legacy")) {\n'+
 '          if (!await keySyncReconcile("docOpsLookup")) throw new Error("Chưa đối soát xong nguồn chứng từ mới; giữ nguyên cache và hàng đợi.");\n'+
 '          await docOpsLookupWriteMeta({sourceEpoch});\n'+
 '          if (sourceEpoch === "pn-full-v1") {\n'+
 '            await loadDocOpsToday(true);\n'+
 '            return true; // keyed incremental sync survives filtering/sorting and relocation\n'+
 '          }\n'+
 '        }\n'+
 '        const remoteLastRow = Math.max(1, Number(manifest.lastRow || 1));');
patch('        version=res.version; all.push(...res.entries);',
 '        if (first.sourceEpoch !== res.sourceEpoch) throw new Error("Nguồn API thay đổi trong khi tải; sẽ đối soát lại.");\n'+
 '        version=res.version; all.push(...res.entries);');
patch('        await docOpsLookupWriteMeta({lastRow:manifest.lastRow,count:docOpsLookupMemory.length});',
 '        await docOpsLookupWriteMeta({lastRow:manifest.lastRow,count:docOpsLookupMemory.length,sourceEpoch:String(manifest.sourceEpoch || "legacy")});');
patch('            if (res.dataVersion) await docOpsLookupWriteMeta({dataVersion:String(res.dataVersion), updatedAt:new Date().toISOString()});',
 '            if (res.record) {\n'+
 '              try { await docOpsRememberRecords([res.record]); }\n'+
 '              catch (cacheError) { keySyncSchedule(30000); }\n'+
 '            }\n'+
 '            if (res.dataVersion) await docOpsLookupWriteMeta({dataVersion:String(res.dataVersion), updatedAt:new Date().toISOString()});');
fs.mkdirSync(path.join(root,'frontend'),{recursive:true});
fs.writeFileSync(path.join(root,'frontend','index.html'),html);
fs.writeFileSync(path.join(root,'frontend','SOURCE.json'),JSON.stringify({source,sha256:crypto.createHash('sha256').update(original).digest('hex'),output:'index.html'},null,2));
console.log('Created frontend candidate; source and offline queues unchanged.');

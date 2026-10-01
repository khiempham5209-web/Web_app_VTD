const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const root = __dirname;
const sources = [
 ['01_API.gs','f1fcd8aa-2ee2-4f03-bb79-3de77b41ace5'],
 ['02_Reject.gs','fec7ac80-47a4-4d22-9250-7c9b34a7762e'],
 ['03_Booking.gs','faf66a8b-66e7-405c-89d8-f6f822c339e4'],
 ['04_Handover.gs','4653a741-25dc-4e84-89f9-de4ae27c605a'],
 ['05_SKU.gs','30174d06-c826-4658-aed7-e7953ac41123'],
 ['06_VHFF.gs','0898c9c2-e478-4523-b156-587faf3b9d3a'],
 ['07_Defect.gs','3b005606-06b8-4464-8b53-64916b15b4fb'],
 ['08_Incident.gs','2737a61d-2342-47ab-a1bf-46270c7f0fb8'],
 ['09_Area.gs','3ba0db59-0cbf-4705-85ca-ae33cd3b9d9c']
];
fs.mkdirSync(path.join(root,'originals'),{recursive:true});
fs.mkdirSync(path.join(root,'apps-script'),{recursive:true});
function replace(s,a,b){if(!s.includes(a))throw Error('Missing patch anchor: '+a.slice(0,100));return s.replace(a,b);}
const manifest=[];
for(const [name,id] of sources){
 const input=path.join(root,'originals',name);
 const bytes=fs.readFileSync(input);
 manifest.push({file:name,source:'originals/'+name,attachmentId:id,sha256:crypto.createHash('sha256').update(bytes).digest('hex')});
 let s=bytes.toString('utf8').replace(/\r\n/g,'\n');
 // All modified scripts run in the API project. One dispatcher owns triggers.
 s=s.replace(/function (installSync\w+Trigger|setupBanGiaoTrigger)\(\) \{/g,'function $1() { return pnInstallTriggers(); /* superseded */\n');
 s=s.replace(/function onEdit\(e\)/g,'function pnLegacyBanGiaoOnEdit(e)');
 // Fixed business headers exclude the pivot at File đơn!T:W and technical metadata.
 s=s.replace(/sh\.getRange\(1, 1, 1, sh\.getLastColumn\(\)\)\.getDisplayValues\(\)\[0\]/g,'pnHeaders_(sh)');
 if(name==='01_API.gs'){
  s=replace(s,'sheetName: "Chứng từ_FF"','sheetName: "Chứng từ_full"');
  s=replace(s,'function apiSave_(params) {','function pnApiSaveCore_(params) {');
  s=replace(s,'  const found = resolveRecordForSave_(sh, col, rowNumber, query);','  const found = pnResolveSave_(sh, col, params);');
  s=replace(s,'  const now = new Date();','  const now = pnRequestNow_();');
  const docStart=s.indexOf('  if (hasDocuments) {\n    setCellByAliases_');
  const docEnd=s.indexOf('\n\n  const productRows = [];',docStart);
  if(docStart<0||docEnd<0)throw Error('Missing document write block');
  s=s.slice(0,docStart)+'  if (hasDocuments) pnWriteDocumentOnce_(sh, found.rowNumber, col, params, upload, timeText, user);'+s.slice(docEnd);
  s=replace(s,'  const upload = hasDocuments ? uploadFiles_(maDonForFolder, files)', '  const upload = hasDocuments ? pnUploadOnce_("docs", () => uploadFiles_(maDonForFolder, files))');
  s=replace(s,'items.map(item => uploadProductFiles_(found.record.orderNo || params.orderNo || maDonForFolder, item))','items.map((item, i) => pnUploadOnce_("product:" + i, () => uploadProductFiles_(found.record.orderNo || params.orderNo || maDonForFolder, item)))');
  s=replace(s,'      productRows.push(values);','      values[pnProductIdColumn_(productSheet) - 1] = pnItemId_(item, index);\n      productRows.push(values);');
  s=replace(s,'    const productHeaders = headers_(productSheet);','    pnProductIdColumn_(productSheet);\n    const productHeaders = headers_(productSheet);');
  s=replace(s,'      const startRow = Math.max(2, productSheet.getLastRow() + 1);\n      productSheet.getRange(startRow, 1, productRows.length, productHeaders.length).setValues(productRows);','      pnAppendProductsOnce_(productSheet, productRows, productHeaders.length);');
  s=replace(s,'    clientItemId: clean_(item.clientItemId),','    clientItemId: clean_(item.clientItemId),');
  s=replace(s,'  if (hasDocuments) callRejectSyncLocal_(found.rowNumber);','  pnMarkDirty_(); // VHFF worker retries independently; never hide a failed sync.');
  s=replace(s,'  rememberSavedRequest_(clientId, response);','  // Durable request journal is finalized by apiSave_ after mirroring.');
  s=replace(s,'  const sh = sheet_();\n  return ok_({','  pnRequireReady_();\n  const sh = sheet_();\n  return ok_({\n    sourceEpoch: PN_FULL.epoch,');
  s=replace(s,'  const sh = ss.getSheetByName(CONFIG.sheetName);','  pnRequireReady_();\n  const sh = ss.getSheetByName(CONFIG.sheetName);');
  s=replace(s,'  data.ok = true;','  data.ok = true;\n  data.sourceEpoch = PN_FULL.epoch;');
  s=replace(s,'const common = {ok:true, protocol:1,','const common = {ok:true, sourceEpoch:PN_FULL.epoch, protocol:1,');
  s=replace(s,'  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();\n  const q = norm_(query);','  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();\n  const q = norm_(query);\n  const matches = values.map((row,i) => ({row,i})).filter(x => [pick_(x.row,col,["ma don ghtk","ma don"]),pick_(x.row,col,["ma po","po"]),pick_(x.row,col,["so don hang","od"])].some(v => norm_(v) === q));\n  if(matches.length > 1) throw new Error("Mã tra cứu khớp nhiều đơn; hãy dùng Số đơn hàng hoặc Mã GHTK duy nhất.");');
 }
 if(name==='02_Reject.gs'){
  s=replace(s,'sourceSheetName: "Chứng từ_FF"','sourceSheetName: "Chứng từ_full"');
  s=replace(s,'orderSheetName: "File đơn"','orderSheetName: "File đơn_full"');
  s=replace(s,'  const sources = rejectSyncSourceRecordsByRows_(rowNumbers);',
   '  const sources = rejectSyncSourceRecordsByRows_(rowNumbers);\n'+
   '  const pnKeys=sources.map(r=>rejectSyncNorm_(r.maDonGhtk)).filter(Boolean);\n'+
   '  if(new Set(pnKeys).size!==pnKeys.length)throw new Error("Chứng từ trùng Mã GHTK; hoãn sync để tránh ghi sai đích.");');
 }
 if(name==='03_Booking.gs'){
  s=replace(s,'orderSheetName: "File đơn"','orderSheetName: "File đơn_full"');
  s=replace(s,'targetSheetName: "Chứng từ_FF"','targetSheetName: "Chứng từ_full"');
  s=replace(s,'  bookingSyncApplyControlColumnsBatch_(target, targetCol, nextRow, output.length);',
   '  bookingSyncApplyControlColumnsBatch_(target, targetCol, nextRow, output.length);\n  pnStampFullIds_(PN_FULL.pairs[2]);\n  pnMarkDirty_();');
  // The original Booking reader remains primary. Full backfill is handled explicitly.
 }
 if(name==='04_Handover.gs')s=replace(s,'sourceSheetName: "Chứng từ_FF"','sourceSheetName: "Chứng từ_full"');
 if(name==='06_VHFF.gs'){
  s=replace(s,"const DOC_SOURCE_SHEET = 'Chứng từ_FF'","const DOC_SOURCE_SHEET = 'Chứng từ_full'");
  s=replace(s,"sheetName: 'Booking'","sheetName: 'Booking_full'");
  s=s.replace(/  clearTargetContent_\(targetSheet, outputWidth\);/g,'  pnRequireReady_();\n  const pnOldLast = targetSheet.getLastRow();');
  s=s.replace(/  applyTargetFormats_\(targetSheet, config, Math.max\(output.length \+ 1, 2\)\);/g,'  if (pnOldLast > output.length + 1) targetSheet.getRange(output.length + 2, 1, pnOldLast - output.length - 1, outputWidth).clearContent();\n  applyTargetFormats_(targetSheet, config, Math.max(output.length + 1, 2));');
 }
 if(name==='08_Incident.gs'){
  s=replace(s,'orderSheetId: 1754451801','orderSheetId: 872196600');
  s=replace(s,'orderSheetName: "File đơn"','orderSheetName: "File đơn_full"');
  s=replace(s,'  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];','  const headers = pnHeaders_(sh);');
  s=replace(s,'  suVuBanhXepClearTarget_(target, targetWidth);','  pnRequireReady_();\n  const pnOldLast = target.getLastRow();');
  s=replace(s,'  suVuBanhXepFormatTarget_(target, outputRows.length, outputHeaders.length);',
   '  if(pnOldLast > outputRows.length + 1) target.getRange(outputRows.length + 2, 1, pnOldLast - outputRows.length - 1, targetWidth).clearContent();\n  suVuBanhXepFormatTarget_(target, outputRows.length, outputHeaders.length);');
 }
 if(name==='07_Defect.gs'){
  s=replace(s,'  defectSyncClearTargetData_(target);','  pnRequireReady_();\n  const pnOldLast = target.getLastRow();');
  s=replace(s,'  defectSyncFormatTarget_(target, output.length);',
   '  if(pnOldLast > output.length + 1) target.getRange(output.length + 2, 1, pnOldLast - output.length - 1, DEFECT_SYNC_CONFIG.targetColumnCount).clearContent();\n  defectSyncFormatTarget_(target, output.length);');
 }
 if(name==='09_Area.gs')s=replace(s,'orderSheetName: "File đơn"','orderSheetName: "File đơn_full"');
 // Re-entrant lock adapter allows dispatcher -> existing business functions.
 s=s.replace(/  const lock = LockService.getScriptLock\(\);\n  lock.waitLock\(30000\);\n  try \{\n    return fn\(\);\n  \} finally \{\n    lock.releaseLock\(\);\n  \}/g,'  return pnWithLock_(fn);');
 s=s.replace(/function (suVuBanhXepSyncWithLock_|khuVucBookingSyncWithLock_)\(fn\) \{[\s\S]*?\n\}\n/g,'function $1(fn) { return pnWithLock_(fn); }\n');
 fs.writeFileSync(path.join(root,'apps-script',name),'// FULL-TABS V1: install with 00/10/11 helpers; follow README, not legacy installation comments below.\n'+s);
}
fs.writeFileSync(path.join(root,'SOURCE_MANIFEST.json'),JSON.stringify(manifest,null,2));
console.log('Built 9 scripts; original bytes preserved.');

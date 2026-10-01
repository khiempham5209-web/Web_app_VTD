const fs=require('fs'),path=require('path'),vm=require('vm'),crypto=require('crypto'),assert=require('assert/strict');
const root=path.join(__dirname,'..'),dir=path.join(root,'apps-script');
const files=fs.readdirSync(dir).filter(x=>x.endsWith('.gs')).sort();
const code=files.map(f=>fs.readFileSync(path.join(dir,f),'utf8')).join('\n');
new vm.Script(code);
let passed=0;
function test(name,fn){try{fn();passed++;console.log('PASS '+name);}catch(e){console.error('FAIL '+name);throw e;}}
const headers={
 'Booking':['Ngày','Row Labels','Ship to location','Customer name','Cust po number','Ship to address1','Quận','Giao nhận','Chuỗi','Tổng bánh','Số Thùng','Thành Tiền','CBM','Note','Khu vực','Check'],
 'File đơn':['Mã đơn','Thời gian tạo đơn','Mã shop','Tỉnh thành địa chỉ giao','Tỉnh thành địa chỉ lấy','Kho nguồn','Kho đích','Loại shop','Loại dịch vụ','Mã KH','Mã đơn hàng KH','Danh sách tên sản phẩm (số lượng)(cân nặng)','Cân nặng đơn qua trụ lần đầu','Vị trí giao','Mã chứng từ phát sinh','Quản lý','Khu vực','Ngay'],
 'Chứng từ_FF':['Ngày lên đơn','Mã đơn GHTK','Khách Hàng','Mã PO','Số đơn hàng','Địa chỉ nhận hàng','Xác thực hóa đơn','Ngày bàn giao CT','Đã tạo sv','Ghi chú chứng từ','Note','Mã ecom CT','Link ảnh','Loại siêu thị','Trạng thái','Thời gian','User thao tác'],
 'Hoàn sản phẩm':['Ngày hoàn trả','Thời gian','Mã đơn','Tên khách hàng','Số đơn hàng','Mã PO','Mã vật tư','Barcode','Tên sản phẩm','Số lượng','Tình trạng FF','Ghi chú','Phân loại','Hạn sử dụng','Kích thước','Khối lượng','Loại siêu thị','Hình ảnh','User thao tác','CBM','Ngày bàn giao'],
 'DS SKU':['Mã vật tư','Barcode','Tên vật tư','Quy cách','Kích thước','Khối lượng'],
};
function colNo(s){return [...s].reduce((v,c)=>v*26+c.charCodeAt(0)-64,0);}
class Range{
 constructor(sh,r,c,n=1,w=1){this.sh=sh;this.r=r;this.c=c;this.n=n;this.w=w;}
 getValues(){return Array.from({length:this.n},(_,i)=>Array.from({length:this.w},(_,j)=>this.sh.rows[this.r+i-1]?.[this.c+j-1]??''));}
 getDisplayValues(){return this.getValues().map(r=>r.map(v=>v instanceof Date?fmt(v,'dd/MM/yyyy'):v===true?'TRUE':v===false?'FALSE':String(v)));}
 getValue(){return this.getValues()[0][0];} getDisplayValue(){return this.getDisplayValues()[0][0];}
 setValues(rows){assert.equal(rows.length,this.n);for(let i=0;i<this.n;i++){assert.equal(rows[i].length,this.w);const r=this.sh.rows[this.r+i-1]??=[];rows[i].forEach((v,j)=>{r[this.c+j-1]=v;});}return this;}
 setValue(v){return this.setValues([[v]]);}
 clearContent(){return this.setValues(Array.from({length:this.n},()=>Array(this.w).fill('')));}
 getFormula(){return this.sh.formulas.get(this.r+':'+this.c)||'';}
 getDataValidation(){return {kind:'mock'};}
 getRow(){return this.r;}getColumn(){return this.c;}getNumRows(){return this.n;}getNumColumns(){return this.w;}
 getLastRow(){return this.r+this.n-1;} getSheet(){return this.sh;}
 getA1Notation(){return 'A1';}
 copyTo(){return this;}
 createFilter(){const range=this,filter={getRange:()=>range,getColumnFilterCriteria:()=>null,setColumnFilterCriteria(){return this;},remove(){range.sh.filter=null;}};this.sh.filter=filter;return filter;}
 createTextFinder(q){const range=this;return {matchEntireCell(){return this;},useRegularExpression(){return this;},findAll(){return range.getValues().flatMap((r,i)=>String(r[0])===q?[new Range(range.sh,range.r+i,range.c)]:[]);}}}
}
for(const method of ['setNumberFormat','setDataValidation','clearDataValidations','clearFormat','setBorder','setFontFamily','setFontSize','setVerticalAlignment','setFontWeight','setHorizontalAlignment','setWrap','merge','breakApart'])Range.prototype[method]=function(){return this;};
class Sheet{
 constructor(name,rows=[]){this.name=name;this.rows=rows;this.maxRows=1000;this.maxCols=40;this.formulas=new Map();}
 getName(){return this.name;} getLastRow(){let n=this.rows.length;while(n&&!this.rows[n-1]?.some(x=>x!==''&&x!=null))n--;return n;}
 getFilter(){return this.filter||null;}
 getLastColumn(){return Math.max(1,...this.rows.map(r=>{let n=r?.length||0;while(n&&(r[n-1]===''||r[n-1]==null))n--;return n;}));}
 getMaxRows(){return this.maxRows;}getMaxColumns(){return this.maxCols;}
 insertRowsAfter(x,n){this.maxRows+=n;}insertColumnsAfter(x,n){this.maxCols+=n;}
 getRange(r,c,n,w){if(typeof r==='string'){const m=r.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);return new Range(this,+m[2],colNo(m[1]),m[4]?+m[4]-m[2]+1:1,m[3]?colNo(m[3])-colNo(m[1])+1:1);}return new Range(this,r,c,n,w);}
 hideColumns(){}hideSheet(){}setFrozenRows(){}autoResizeRows(){}getSheetId(){return ({'File đơn_full':872196600,'Sự vụ':529840243,'Sự vụ bánh xẹp':318091884})[this.name]||1;}
}
function fmt(d,pattern){
 const parts=Object.fromEntries(new Intl.DateTimeFormat('en-GB',{timeZone:'Asia/Ho_Chi_Minh',year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',second:'2-digit',hourCycle:'h23'}).formatToParts(d).map(x=>[x.type,x.value]));
 return pattern.replace(/yyyy|MM|dd|HH|mm|ss|(?<![a-zA-Z])d(?![a-zA-Z])|(?<![a-zA-Z])M(?![a-zA-Z])/g,x=>({yyyy:parts.year,MM:parts.month,dd:parts.day,HH:parts.hour,mm:parts.minute,ss:parts.second,d:String(+parts.day),M:String(+parts.month)})[x]);
}
function environment(){
 const sheets=new Map(),props={},targets=new Map();
 const ss={getSheetByName:n=>sheets.get(n),insertSheet:n=>{const s=new Sheet(n);sheets.set(n,s);return s;},getSheets:()=>[...sheets.values()]};
 for(const name of ['Chứng từ không đạt YC','Đã bàn giao CT','Bàn giao SP_Hàng hoàn','Hàng lỗi (chuyển sang GHTK)','Sự vụ bánh xẹp'])targets.set(name,new Sheet(name));
 const targetSS={getSheetByName:n=>targets.get(n),getSheets:()=>[...targets.values()]};
 Object.entries(headers).forEach(([n,h])=>sheets.set(n,new Sheet(n,[h.slice()])));
 for(const n of ['Booking_full','File đơn_full','Chứng từ_full','TT Nhập','DS BC','Bàn giao chứng từ','Hàng lỗi','Sự vụ'])sheets.set(n,new Sheet(n));
 sheets.get('TT Nhập').getRange('L3').setValue(0.035893879030581485);
 const depot=Array(15).fill('');depot[4]='Manager';depot[6]='Depot';depot[14]='Nội thành';
 sheets.get('DS BC').rows=[Array(15).fill('Header'),depot];
 const context=vm.createContext({console,Date,Map,Set,Math,JSON,Number,String,Boolean,Array,Object,Error,
  SpreadsheetApp:{openById:id=>id==='1ZLqeo_djwMJofqBCD3Ilk6lIkUmVsEfK7T5Oj9D6wmI'?targetSS:ss,flush(){},CopyPasteType:{PASTE_FORMAT:1,PASTE_DATA_VALIDATION:2},
   newDataValidation:()=>({requireValueInList(){return this;},requireValueInRange(){return this;},setAllowInvalid(){return this;},build(){return {};}})},
  Utilities:{getUuid:()=>crypto.randomUUID(),DigestAlgorithm:{SHA_256:'sha256',MD5:'md5'},Charset:{UTF_8:'utf8'},
   computeDigest:(a,s)=>[...crypto.createHash(a).update(s).digest()],formatDate:(d,t,p)=>fmt(d,p)},
  PropertiesService:{getScriptProperties:()=>({getProperty:k=>props[k]??null,setProperty:(k,v)=>{props[k]=v;},deleteProperty:k=>delete props[k],getProperties:()=>({...props})})},
  LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){},tryLock(){return true;}})},
  Session:{getActiveUser:()=>({getEmail:()=> 'test@example.invalid'})}
 });
 vm.runInContext(code,context);
 vm.runInContext("pnCurrentMonth_=()=> '2026-10';",context);
 return {c:context,sheets,targets,props,run:s=>vm.runInContext(s,context)};
}
function doc(order,date,status){return [date,'G'+order,'Customer','PO'+order,String(order),'Address',status,'',false,'','','','','Winmart','Bình thường','',''];}
function seed(e){
 e.sheets.get('Chứng từ_FF').rows.push(doc(1,'01/10/2026','Chưa nhận chứng từ'),doc(2,'01/09/2026','Chưa nhận chứng từ'),doc(3,'01/09/2026','Đã nhận chứng từ'),doc(4,'01/09/2026','Shop hủy OD'));
 const booking=Array(16).fill('');booking[0]='01/09/2026';booking[1]='2';booking[9]=32;booking[15]=2;
 e.sheets.get('Booking').rows.push(booking);
 const fd=Array(18).fill('');fd[0]='G2';fd[1]='2026-09-01 19:00:00';fd[6]='Depot';fd[10]='2';
 fd[15]='Manager';fd[16]='Nội thành';fd[17]='01/09/2026';
 e.sheets.get('File đơn').rows.push(fd,fd.map((x,i)=>i===13?'Changed position':x));
 e.sheets.get('File đơn').getRange(2,20,1,4).setValues([['pivot','must','stay','here']]);
}
test('All Apps Script files compile together without duplicate const declarations',()=>{});
test('Original files are byte-preserved and SHA256 matches source',()=>{
 const manifest=JSON.parse(fs.readFileSync(path.join(root,'SOURCE_MANIFEST.json')));
 manifest.forEach(m=>assert.equal(crypto.createHash('sha256').update(fs.readFileSync(path.join(root,'originals',m.file))).digest('hex'),m.sha256));
});
test('Month parsing validates dates, time zones, year rollover',()=>{
 const e=environment();
 assert.equal(e.c.pnMonth_('31/02/2026'),'');assert.equal(e.c.pnMonth_('2026-09-30 23:59:59'),'2026-09');
 assert.equal(e.c.pnMonth_(new Date('2026-09-30T18:00:00Z')),'2026-10');
 assert.equal(e.c.pnMonth_('01/01/2027'),'2027-01');
});
test('Current month retained; old received/cancelled removed; pending/unknown retained',()=>{
 const e=environment(),p=e.run('PN_FULL.pairs[2]');
 for(const status of ['Đã nhận chứng từ','Shop hủy OD']){
  assert.equal(e.c.pnEligible_(doc(1,'01/09/2026',status),p,'2026-10'),true);
  assert.equal(e.c.pnEligible_(doc(1,'01/10/2026',status),p,'2026-10'),false);
 }
 for(const status of ['Chưa nhận chứng từ','','Other'])assert.equal(e.c.pnEligible_(doc(1,'01/09/2026',status),p,'2026-10'),false);
});
test('Formula equivalents preserve ROUNDUP sign, first XLOOKUP match and missing result',()=>{
 const e=environment(),bp=e.run('PN_FULL.pairs[0]'),fp=e.run('PN_FULL.pairs[1]');
 assert.equal(e.c.pnRoundUp_(-1.2),-2);assert.equal(e.c.pnRoundUp_(0),0);
 const b=Array(16).fill('');b[0]='x';b[9]=32;
 assert.equal(e.c.pnFormulaValues_(bp,[b],0.035893879030581485,[])[0][0],2);
 const f=Array(18).fill('');f[0]='g';f[1]='2026-09-30 10:00:00';f[6]='DEPOT';
 const a=Array(15).fill('');a[6]='Depot';a[4]='First';a[14]='Area';
 const second=a.slice();second[4]='Second';
 assert.equal(e.c.pnFormulaValues_(fp,[f],0,[a,second])[0][0],'First');
 assert.equal(e.c.pnFormulaValues_(fp,[f],0,[])[0][0],'#N/A');
});
test('Migration retains repeated File đơn rows and all 17 document fields; rerun is idempotent',()=>{
 const e=environment();seed(e);assert.equal(e.c.pnMigrateFull().ok,true);
 assert.equal(e.sheets.get('File đơn_full').getLastRow(),3);
 assert.equal(e.sheets.get('Chứng từ_full').getLastRow(),5);
 assert.equal(e.sheets.get('File đơn').getRange(2,20).getValue(),'pivot');
 assert.equal(e.c.pnMigrateFull().ok,true);assert.equal(e.sheets.get('File đơn_full').getLastRow(),3);
 assert.equal(e.sheets.get('Chứng từ_full').getRange(2,18).getValue(),'2026-10');
});
test('Retry after full write but before primary checkpoint does not duplicate repeated rows',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();
 e.sheets.get('File đơn').getRange(2,30,2,2).clearContent();
 e.c.pnReconcilePair_(e.run('PN_FULL.pairs[1]'));
 assert.equal(e.sheets.get('File đơn_full').getLastRow(),3);
});
test('Primary edits push; full edits pull; conflicting edits fail closed',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();const p=e.run('PN_FULL.pairs[2]');
 e.sheets.get('Chứng từ_FF').getRange(2,10).setValue('from main');e.c.pnReconcilePair_(p);
 assert.equal(e.sheets.get('Chứng từ_full').getRange(2,10).getValue(),'from main');
 e.sheets.get('Chứng từ_full').getRange(2,10).setValue('from full');e.c.pnReconcilePair_(p);
 assert.equal(e.sheets.get('Chứng từ_FF').getRange(2,10).getValue(),'from full');
 e.sheets.get('Chứng từ_full').getRange(2,10).setValue('A');e.sheets.get('Chứng từ_FF').getRange(2,10).setValue('B');
 assert.throws(()=>e.c.pnReconcilePair_(p),/Xung đột/);
});
test('Cleanup refuses when disabled or VHFF fails; success preserves pivot and full',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();
 assert.equal(e.c.pnCleanupDaily().ok,false);
 e.props.PN_FULL_CLEANUP='enabled';
 e.run("pnSyncVHFF_=()=>{throw new Error('VHFF unavailable');};");
 assert.throws(()=>e.c.pnCleanupDaily(),/VHFF unavailable/);
 assert.equal(e.sheets.get('Chứng từ_FF').getRange(5,5).getValue(),'4');
 e.run("pnSyncVHFF_=()=>{PropertiesService.getScriptProperties().deleteProperty('PN_VHFF_DIRTY');return {ok:true};};");
 const r=e.c.pnCleanupDaily();assert.equal(r.ok,true);
 assert.equal(e.sheets.get('Chứng từ_full').getLastRow(),5);
 assert.equal(e.sheets.get('Chứng từ_FF').getRange(4,5).getValue(),'');
 assert.equal(e.sheets.get('Chứng từ_FF').getRange(3,5).getValue(),'2');
 assert.equal(e.sheets.get('File đơn').getRange(2,20).getValue(),'pivot');
 assert.equal(e.c.pnAuditFull().ok,true);
 assert.equal(e.c.pnCleanupDaily().ok,true);
});
test('API full source lists old rows after primary cleanup; save ignores stale rowNumber',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();
 e.sheets.get('Chứng từ_FF').getRange(4,1,2,17).clearContent();
 e.sheets.get('Chứng từ_FF').getRange(4,30,2,2).clearContent();
 assert.equal(e.c.apiPage_({}).records.length,4);
 assert.equal(e.c.apiKeySync_({}).total,4);
 const found=e.c.pnResolveSave_(e.c.sheet_(),e.c.headerMap_(e.c.headers_(e.c.sheet_())),{maDon:'G3',rowNumber:2});
 assert.equal(found.record.orderNo,'3');assert.equal(found.rowNumber,4);
 assert.throws(()=>e.c.pnResolveSave_(e.c.sheet_(),{}, {rowNumber:2}),/khóa đơn/);
});
test('Duplicate API clientId returns same result, uploads once, mirrors full and main',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();
 e.run("var uploadCount=0; uploadFiles_=()=>{uploadCount++;return {linkAnh:'https://example.invalid/photo',files:[{id:'1'}],folderUrl:''};};");
 const p={clientId:'test-request',maDon:'G1',rowNumber:999,returnType:'Chứng từ',xacThuc:'Đã nhận chứng từ',status:'Bình thường',note:'OK',files:[{base64:'AA=='}]};
 assert.equal(e.c.apiSave_(p).ok,true);assert.equal(e.c.apiSave_(p).ok,true);
 assert.equal(e.run('uploadCount'),1);
 assert.equal(e.sheets.get('Chứng từ_full').getRange(2,7).getValue(),'Đã nhận chứng từ');
 assert.equal(e.sheets.get('Chứng từ_FF').getRange(2,7).getValue(),'Đã nhận chứng từ');
 assert.equal(e.c.apiSave_({...p,note:'changed'}).ok,false);
});
test('Upload uncertainty is quarantined; same request never automatically uploads twice',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();
 e.run("var uploadCount=0; uploadFiles_=()=>{uploadCount++;throw new Error('lost response');};");
 const p={clientId:'uncertain',maDon:'G1',returnType:'Chứng từ',files:[{base64:'AA=='}]};
 assert.throws(()=>e.c.apiSave_(p),/lost response/);
 assert.throws(()=>e.c.apiSave_(p),/chưa xác định/);
 assert.equal(e.run('uploadCount'),1);
});
test('Product rows deduplicate by request item ID',()=>{
 const e=environment();const sh=e.sheets.get('Hoàn sản phẩm'),c=e.c.pnProductIdColumn_(sh),r=Array(c).fill('');r[2]='G1';r[c-1]='request:item';
 e.c.pnAppendProductsOnce_(sh,[r],c);e.c.pnAppendProductsOnce_(sh,[r],c);
 assert.equal(sh.getLastRow(),2);
});
test('Request retry survives session-token renewal and stale row relocation',()=>{
 const e=environment(),a={clientId:'x',maDon:'G1',sessionToken:'old',rowNumber:2,_silent:true},b={...a,sessionToken:'new',rowNumber:900};
 assert.equal(e.c.pnRequestSignature_(a),e.c.pnRequestSignature_(b));
});
test('Later retry of completed document stage cannot overwrite a newer note',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();
 e.run("pnRequest_={key:'request:test',clientId:'test'};");
 const sh=e.c.sheet_(),col=e.c.headerMap_(e.c.headers_(sh)),params={xacThuc:'Đã nhận chứng từ',note:'first'};
 e.c.pnWriteDocumentOnce_(sh,2,col,params,{linkAnh:'url'},'01/10/2026 10:00:00','user');
 sh.getRange(2,10).setValue('newer');
 e.c.pnWriteDocumentOnce_(sh,2,col,params,{linkAnh:'url'},'01/10/2026 10:00:00','user');
 assert.equal(sh.getRange(2,10).getValue(),'newer');
});
test('Full archive edits remain visible in VHFF handover after primary is cleared',()=>{
 const e=environment();seed(e);
 e.sheets.get('Chứng từ_FF').getRange(4,8).setValue(new Date('2026-09-20T00:00:00+07:00'));
 e.c.pnMigrateFull();
 e.c.syncBanGiaoChungTu();
 const dest=e.targets.get('Đã bàn giao CT');
 assert.equal(dest.getLastRow(),2);assert.equal(dest.getRange(2,5).getValue(),'3');
 e.sheets.get('Chứng từ_FF').getRange(4,1,2,17).clearContent();
 e.sheets.get('Chứng từ_FF').getRange(4,30,2,2).clearContent();
 e.sheets.get('Chứng từ_full').getRange(4,11).setValue('archived note');
 e.c.syncBanGiaoChungTu();
 assert.equal(dest.getLastRow(),2);assert.equal(dest.getRange(2,9).getValue(),'archived note');
});
test('VHFF reject sync uses historical File đơn and preserves both manual Ecom columns',()=>{
 const e=environment();seed(e);e.sheets.get('Chứng từ_FF').getRange(3,9).setValue(true);e.c.pnMigrateFull();
 e.c.syncAllChungTuKhongDatYC();
 const dest=e.targets.get('Chứng từ không đạt YC');
 const h=dest.getRange(1,1,1,dest.getLastColumn()).getValues()[0];
 const a=h.indexOf('Ecom FF -> Kho đích')+1,b=h.indexOf('Ecom Kho đích -> FF')+1;
 assert.ok(a>0&&b>0);dest.getRange(2,a).setValue('manual A');dest.getRange(2,b).setValue('manual B');
 e.sheets.get('File đơn').getRange(2,1,2,18).clearContent();
 e.c.syncAllChungTuKhongDatYC();
 assert.equal(dest.getLastRow(),2);assert.equal(dest.getRange(2,a).getValue(),'manual A');assert.equal(dest.getRange(2,b).getValue(),'manual B');
 assert.equal(dest.getRange(2,4).getValue(),'Depot');
});
test('VHFF returns retain booking date after primary Booking is cleared',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();
 const h=headers['Hoàn sản phẩm'],r=Array(h.length).fill('');
 r[h.indexOf('Số đơn hàng')]='2';r[h.indexOf('Mã vật tư')]='SKU';r[h.indexOf('Tên sản phẩm')]='Product';
 e.sheets.get('Hoàn sản phẩm').rows.push(r);e.sheets.get('Booking').getRange(2,1,1,16).clearContent();
 e.c.syncBanGiaoSanPham();
 assert.equal(e.targets.get('Bàn giao SP_Hàng hoàn').getRange(2,20).getValue(),'01/09/2026');
});
test('VHFF worker acknowledges all destinations and notices later archive-only edits',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();
 assert.equal(e.c.pnSyncVHFF().ok,true);assert.equal(e.props.PN_VHFF_DIRTY,undefined);
 const before=e.props.PN_VHFF_SOURCE_ACK;
 e.sheets.get('Chứng từ_FF').getRange(4,1,2,17).clearContent();e.sheets.get('Chứng từ_FF').getRange(4,30,2,2).clearContent();
 e.sheets.get('Chứng từ_full').getRange(4,10).setValue('late change');
 assert.equal(e.c.pnSyncVHFF().ok,true);assert.notEqual(e.props.PN_VHFF_SOURCE_ACK,before);
});
test('Interrupted cleanup blocks API and resumes without detaching identities',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();const p=e.run('PN_FULL.pairs[2]'),sh=e.sheets.get('Chứng từ_FF');
 const before=e.c.pnRows_(sh,31);e.c.pnBackupCleanup_(p,before);
 e.props.PN_CLEANUP_PENDING=JSON.stringify({main:p.main,month:'2026-10'});
 // Simulate successful business compaction and crash before identity write.
 const kept=[before[0],before[1]];
 sh.getRange(2,1,2,17).setValues(kept.map(r=>r.slice(0,17)));sh.getRange(4,1,2,17).clearContent();
 assert.throws(()=>e.c.apiInit_(),/gián đoạn/);
 assert.equal(e.c.pnResumeCleanup().ok,true);assert.equal(e.props.PN_CLEANUP_PENDING,undefined);
 assert.equal(e.c.pnAuditFull().ok,true);
});
test('Interrupted cleanup refuses to overwrite a subsequent manual edit',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();const p=e.run('PN_FULL.pairs[2]'),sh=e.sheets.get('Chứng từ_FF');
 e.c.pnBackupCleanup_(p,e.c.pnRows_(sh,31));e.props.PN_CLEANUP_PENDING=JSON.stringify({main:p.main,month:'2026-10'});
 sh.getRange(2,10).setValue('new manual note');
 assert.throws(()=>e.c.pnResumeCleanup(),/sửa tay/);
 assert.equal(sh.getRange(2,10).getValue(),'new manual note');
});
test('Full filter includes hidden IDs and unexpected loss of history is blocked',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();
 assert.equal(e.sheets.get('Chứng từ_full').getFilter().getRange().getNumColumns(),31);
 e.sheets.get('Chứng từ_full').getRange(5,1,1,31).clearContent();
 assert.throws(()=>e.c.pnMirrorAll(),/giảm số hồ sơ/);
});
test('Archived received/cancelled rows stay out of main, old pending returns to main',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();
 e.sheets.get('Chứng từ_FF').getRange(4,1,2,17).clearContent();e.sheets.get('Chứng từ_FF').getRange(4,30,2,2).clearContent();
 e.c.pnMaterializeDocs_();assert.equal(e.sheets.get('Chứng từ_FF').getRange(4,5).getValue(),'');
 e.sheets.get('Chứng từ_full').getRange(4,7).setValue('Chưa nhận chứng từ');
 e.c.pnMaterializeDocs_();assert.equal(e.sheets.get('Chứng từ_FF').getRange(4,5).getValue(),'3');
 assert.equal(e.sheets.get('Chứng từ_FF').getRange(5,5).getValue(),'');
});
test('Full source init carries epoch; frontend candidate parses and retains offline queue',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();assert.equal(e.c.apiInit_().sourceEpoch,'pn-full-v1');
 const html=fs.readFileSync(path.join(root,'frontend','index.html'),'utf8');
 for(const m of html.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script>/gi)){
  if(/src=|application\/(?:ld\+)?json|importmap/i.test(m[1]))continue;
  new vm.Script(m[2]);
 }
 assert.match(html,/sourceEpoch !== String\(docOpsLookupMetaMemory.sourceEpoch/);
 assert.match(html,/syncKey: record.syncKey \|\| pnSyncKey\(record\)/);
 assert.doesNotMatch(html,/indexedDB\.deleteDatabase/);
});
test('Mixed date-formatted Booking cells survive migration and interrupted-format recovery',()=>{
 const originalCopy=Range.prototype.copyTo;
 const serial=d=>(Date.parse(fmt(d,'yyyy-MM-dd HH:mm:ss').replace(' ','T')+'Z')-Date.UTC(1899,11,30))/86400000;
 Range.prototype.copyTo=function(target,type){
  if(type===1)for(let i=0;i<target.n;i++)for(let j=0;j<target.w;j++){
   const source=this.sh.getRange(this.r+i%this.n,this.c+j).getValue(),cell=target.sh.getRange(target.r+i,target.c+j),value=cell.getValue();
   if(source instanceof Date&&typeof value==='number')cell.setValue(new Date(Date.UTC(1899,11,30)+value*86400000-7*3600000));
   else if(!(source instanceof Date)&&value instanceof Date)cell.setValue(serial(value));
  }
  return this;
 };
 try{
  const e=environment();seed(e);const main=e.sheets.get('Booking'),full=e.sheets.get('Booking_full');
  main.rows.push(main.rows[1].slice());main.rows[2][1]='mixed';main.rows[2][10]=new Date('2026-07-03T17:00:00Z');
  e.c.pnMigrateFull();assert.equal(e.c.pnAudit_().ok,true);
  assert(full.rows[2][10] instanceof Date);
  delete e.props.PN_FULL_READY;full.rows[2][10]=serial(main.rows[2][10]);
  assert.equal(e.c.pnAudit_().ok,false);
  e.c.pnRepairMigrationFormats();assert.equal(e.c.pnAudit_().ok,true);
  assert.equal(main.rows[2][10].getTime(),full.rows[2][10].getTime());
  full.rows[2][10]=serial(main.rows[2][10])+1;
  assert.throws(()=>e.c.pnRepairMigrationFormats(),/thay đổi dữ liệu thật/);
 }finally{Range.prototype.copyTo=originalCopy;}
});
test('Formatting equivalence does not hide changed numbers, strings or dates',()=>{
 const e=environment();
 const date=new Date(Date.UTC(1899,11,30)+46206*86400000-7*3600000);
 assert.equal(e.c.pnFormatOnlyEquivalent_(date,46206),true);
 assert.equal(e.c.pnFormatOnlyEquivalent_(date,46207),false);
 assert.equal(e.c.pnFormatOnlyEquivalent_('46206',46206),false);
 assert.equal(e.c.pnFormatOnlyEquivalent_(2,3),false);
});
test('Repair defers uncheckpointed rows but refuses a missing known ID',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();delete e.props.PN_FULL_READY;
 const main=e.sheets.get('Booking');const fresh=main.rows[1].slice(0,16);fresh[1]='new-order';main.rows.push(fresh);
 let result=e.c.pnRepairMigrationFormats();assert.equal(result.ok,true);assert.equal(result.migrationReady,false);
 assert.equal(result.pendingMigration.length,1);assert.equal(result.pendingMigration[0].row,3);
 assert.equal(main.rows[2][29],undefined);assert.equal(e.c.pnAudit_().ok,false);
 e.c.pnMigrateFull();assert.equal(e.c.pnAudit_().ok,true);
 delete e.props.PN_FULL_READY;main.rows[2][29]='missing-known-id';
 assert.throws(()=>e.c.pnRepairMigrationFormats(),/Thiếu ID full/);
 main.rows[2][29]='';assert.throws(()=>e.c.pnRepairMigrationFormats(),/checkpoint nhưng mất ID/);
});
test('Repair defers primary-only blank derived column; migration restores it without losing late Booking',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();delete e.props.PN_FULL_READY;
 const main=e.sheets.get('Booking'),full=e.sheets.get('Booking_full');
 main.rows[1][15]='';const late=main.rows[1].slice(0,16);late[0]='30/09/2026';late[1]='late-order';main.rows.push(late);
 const result=e.c.pnRepairMigrationFormats();assert.equal(result.ok,true);assert.equal(result.pendingMigration.length,2);
 assert.equal(main.rows[1][15],'');assert.equal(full.rows[1][15],2);
 e.c.pnMigrateFull();assert.equal(e.c.pnAudit_().ok,true);assert.equal(main.rows[1][15],2);
 assert.equal(full.rows.filter(r=>r[1]==='late-order').length,1);
 delete e.props.PN_FULL_READY;main.rows[1][15]='';full.rows[1][9]=999;
 assert.throws(()=>e.c.pnRepairMigrationFormats(),/thay đổi dữ liệu thật/);
});
console.log('RESULT '+passed+' tests passed.');

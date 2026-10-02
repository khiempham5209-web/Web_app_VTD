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
 copyTo(to,type){if(type===2){const m=this.sh.fmt?.dv;if(m){const fm=(to.sh.fmt??={}),d=(fm.dv??=new Map());for(let i=0;i<to.n;i++)for(let j=0;j<to.w;j++){const v=m.get((this.r+(i%this.n))+':'+(this.c+j));if(v!=null)d.set((to.r+i)+':'+(to.c+j),v);}}}return this;}
 createFilter(){const range=this,filter={getRange:()=>range,getColumnFilterCriteria:()=>null,setColumnFilterCriteria(){return this;},remove(){range.sh.filter=null;}};this.sh.filter=filter;return filter;}
 createTextFinder(q){const range=this;return {matchEntireCell(){return this;},useRegularExpression(){return this;},findAll(){return range.getValues().flatMap((r,i)=>String(r[0])===q?[new Range(range.sh,range.r+i,range.c)]:[]);}}}
}
const FMT_KINDS={Backgrounds:'bg',FontColors:'fc',FontWeights:'fw',FontLines:'fl',NumberFormats:'nf',DataValidations:'dv'};
for(const [name,k] of Object.entries(FMT_KINDS)){
 Range.prototype['get'+name]=function(){const m=this.sh.fmt?.[k]||new Map();return Array.from({length:this.n},(_,i)=>Array.from({length:this.w},(_,j)=>m.get((this.r+i)+':'+(this.c+j))??null));};
 Range.prototype['set'+name]=function(rows){assert.equal(rows.length,this.n);const fm=(this.sh.fmt??={});const m=(fm[k]??=new Map());rows.forEach((row,i)=>{assert.equal(row.length,this.w);row.forEach((v,j)=>{if(v==null)m.delete((this.r+i)+':'+(this.c+j));else m.set((this.r+i)+':'+(this.c+j),v);});});return this;};
}
Range.prototype.clearFormat=function(){for(const k of ['bg','fc','fw','fl','nf']){const m=this.sh.fmt?.[k];if(!m)continue;for(let i=0;i<this.n;i++)for(let j=0;j<this.w;j++)m.delete((this.r+i)+':'+(this.c+j));}return this;};
Range.prototype.clearDataValidations=function(){const m=this.sh.fmt?.dv;if(m)for(let i=0;i<this.n;i++)for(let j=0;j<this.w;j++)m.delete((this.r+i)+':'+(this.c+j));return this;};
for(const method of ['setNumberFormat','setDataValidation','setBorder','setFontFamily','setFontSize','setVerticalAlignment','setFontWeight','setHorizontalAlignment','setWrap','merge','breakApart'])Range.prototype[method]=function(){return this;};
class Sheet{
 constructor(name,rows=[]){this.name=name;this.rows=rows;this.maxRows=1000;this.maxCols=40;this.formulas=new Map();}
 getName(){return this.name;} getLastRow(){let n=this.rows.length;while(n&&!this.rows[n-1]?.some(x=>x!==''&&x!=null))n--;return n;}
 getFilter(){return this.filter||null;}
 getLastColumn(){return Math.max(1,...this.rows.map(r=>{let n=r?.length||0;while(n&&(r[n-1]===''||r[n-1]==null))n--;return n;}));}
 getMaxRows(){return this.maxRows;}getMaxColumns(){return this.maxCols;}
 insertRowsAfter(x,n){this.maxRows+=n;}insertColumnsAfter(x,n){this.maxCols+=n;}
 getRange(r,c,n,w){if(typeof r==='string'){const m=r.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);return new Range(this,+m[2],colNo(m[1]),m[4]?+m[4]-m[2]+1:1,m[3]?colNo(m[3])-colNo(m[1])+1:1);}return new Range(this,r,c,n,w);}
 appendRow(row){const n=this.getLastRow()+1;this.rows[n-1]=row.slice();return this;}
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
 // Quy tắc mới: dọn không phụ thuộc VHFF; chỉ bỏ dòng tháng cũ đã khớp full.
 e.run("pnSyncVHFF_=()=>{throw new Error('VHFF must not be required');};");
 assert.equal(e.sheets.get('Chứng từ_FF').getRange(5,5).getValue(),'4');
 const r=e.c.pnCleanupDaily();assert.equal(r.ok,true);
 assert.equal(e.sheets.get('Booking').getRange(2,2).getValue(),'','old-month Booking row left main');
 assert.equal(e.sheets.get('Booking_full').getRange(2,2).getValue(),'2','Booking history kept');
 assert.equal(e.sheets.get('File đơn').getRange(2,1).getValue(),'','old-month File đơn rows left main');
 assert.equal(e.sheets.get('File đơn_full').getLastRow(),3,'File đơn history kept');
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
test('Drive upload error can be resent; upload in progress returns PN_BUSY; a dead upload is retried after 7 minutes',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();
 e.run("var uploadCount=0; uploadFiles_=()=>{uploadCount++;if(uploadCount===1)throw new Error('lost response');return {linkAnh:'https://example.invalid/p',files:[{id:'1'}],folderUrl:''};};");
 const p={clientId:'u1',maDon:'G1',returnType:'Chứng từ',files:[{base64:'AA=='}]};
 assert.throws(()=>e.c.apiSave_(p),/lost response/);
 assert.equal(e.c.apiSave_(p).ok,true,'resend after a Drive error is not stuck');
 assert.equal(e.run('uploadCount'),2);
 const q={clientId:'u2',maDon:'G1',returnType:'Chứng từ',files:[{base64:'AA=='}]};
 const stageKey=e.c.pnRequestKey_('u2')+':docs';
 e.c.pnJournalWrite_(stageKey,'RUNNING',{startedAtMs:Date.now()});
 const busy=e.c.apiSave_(q);assert.equal(busy.code,'PN_BUSY','upload in progress -> retryable busy, app keeps the order pending');
 e.c.pnJournalWrite_(stageKey,'RUNNING',{startedAtMs:Date.now()-11*60*1000});
 assert.equal(e.c.apiSave_(q).ok,true,'dead upload older than 10 minutes is retried');
});
test('Resend with new _diagRequestId, appVersion and re-encoded image returns the saved result',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();
 e.run("var uploadCount=0; uploadFiles_=()=>{uploadCount++;return {linkAnh:'https://example.invalid/p',files:[{id:'1'}],folderUrl:''};};");
 const a={clientId:'r1',maDon:'G1',returnType:'Chứng từ',xacThuc:'Đã nhận chứng từ',note:'OK',_diagRequestId:'diag_1',appVersion:'5.9.17',files:[{fileName:'a.jpg',base64:'AA=='}]};
 const b={...a,_diagRequestId:'diag_2',appVersion:'5.9.18',files:[{fileName:'a.jpg',base64:'AAAA'}]};
 const first=e.c.apiSave_(a),second=e.c.apiSave_(b);
 assert.equal(first.ok,true);assert.equal(second.ok,true);
 assert.equal(second.time,first.time,'second call returns the stored result');
 assert.equal(e.run('uploadCount'),1);
 assert.equal(e.c.apiSave_({...a,note:'changed'}).ok,false,'different business content is still rejected');
});
test('API save touches only its own order; a two-sided edit elsewhere is resolved by the next sync (main wins, losing copy logged)',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();
 e.run("uploadFiles_=()=>({linkAnh:'https://example.invalid/p',files:[{id:'1'}],folderUrl:''});");
 const main=e.sheets.get('Chứng từ_FF'),full=e.sheets.get('Chứng từ_full');
 const rowOf=(sh,code)=>sh.rows.findIndex(r=>r&&r[1]===code)+1;
 const mainG1=rowOf(main,'G1');assert.ok(mainG1>1,'G1 still in main');
 const mainRow=mainG1,fullRow=rowOf(full,'G1');
 main.getRange(mainRow,10).setValue('main edit');full.getRange(fullRow,10).setValue('full edit');
 const g2Main=rowOf(main,'G2');
 assert.ok(g2Main>1,'G2 still in main');
 const res=e.c.apiSave_({clientId:'c1',maDon:'G2',returnType:'Chứng từ',xacThuc:'Đã nhận chứng từ',files:[{base64:'AA=='}]});
 assert.equal(res.ok,true,'save of another order succeeds');
 assert.equal(res.mirrorWarning,'');
 // API chỉ đồng bộ đúng đơn vừa lưu (G2): đơn G1 đang lệch không bị API đụng tới.
 assert.equal(full.getRange(rowOf(full,'G1'),10).getValue(),'full edit','API does not touch other orders');
 // Lần đồng bộ sau xử lý G1: hai bên cùng sửa -> tab chính thắng, bản full bị thay lưu ở _PN_CONFLICTS.
 e.c.pnScheduledReconcile();
 assert.equal(full.getRange(rowOf(full,'G1'),10).getValue(),'main edit');
 assert.equal(main.getRange(rowOf(main,'G1'),10).getValue(),'main edit');
 assert.match(JSON.stringify(e.sheets.get('_PN_CONFLICTS').rows),/full edit/);
 assert.equal(full.getRange(rowOf(full,'G2'),7).getValue(),'Đã nhận chứng từ');
});
test('App save is independent of the Sheet lane: succeeds while the Sheet lock is held, writes only its order, both tabs',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();e.c.pnMirrorAll();
 e.run("uploadFiles_=()=>({linkAnh:'https://example.invalid/p',files:[{id:'1'}],folderUrl:''});");
 e.run("LockService={getScriptLock:()=>({tryLock(){return false;},waitLock(){throw new Error('locked');},releaseLock(){}})};");
 const res=e.c.apiSave_({clientId:'lane',maDon:'G1',returnType:'Chứng từ',xacThuc:'Đã nhận chứng từ',files:[{base64:'AA=='}]});
 assert.equal(res.ok,true,'API does not wait for the Sheet lock');
 const main=e.sheets.get('Chứng từ_FF'),full=e.sheets.get('Chứng từ_full');
 assert.equal(full.getRange(rowOf(full,4,'1'),7).getValue(),'Đã nhận chứng từ');
 assert.equal(main.getRange(rowOf(main,4,'1'),7).getValue(),'Đã nhận chứng từ');
});
test('Photos upload outside the shared lock',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();
 e.run("var depthAtUpload=-1; uploadFiles_=()=>{depthAtUpload=pnLockDepth_;return {linkAnh:'https://example.invalid/p',files:[{id:'1'}],folderUrl:''};};");
 assert.equal(e.c.apiSave_({clientId:'o1',maDon:'G1',returnType:'Chứng từ',files:[{base64:'AA=='}]}).ok,true);
 assert.equal(e.run('depthAtUpload'),0);
});
test('Mã ecom CT is written when sent and never cleared by an app that does not send it',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();
 e.run("uploadFiles_=()=>({linkAnh:'https://example.invalid/p',files:[{id:'1'}],folderUrl:''});");
 const main=e.sheets.get('Chứng từ_FF'),full=e.sheets.get('Chứng từ_full');
 const rowOf=(sh,code)=>sh.rows.findIndex(r=>r&&r[1]===code)+1;
 assert.equal(e.c.apiSave_({clientId:'e1',maDon:'G1',returnType:'Chứng từ',maEcomCt:'EC123',files:[{base64:'AA=='}]}).ok,true);
 assert.equal(full.getRange(rowOf(full,'G1'),12).getValue(),'EC123');
 assert.equal(main.getRange(rowOf(main,'G1'),12).getValue(),'EC123');
 main.getRange(rowOf(main,'G2'),12).setValue('KEEP');
 assert.equal(e.c.apiSave_({clientId:'e2',maDon:'G2',returnType:'Chứng từ',files:[{base64:'AA=='}]}).ok,true);
 assert.equal(full.getRange(rowOf(full,'G2'),12).getValue(),'KEEP');
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
test('Full filter includes hidden IDs; loss of history raises an alert and is restored from main',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();
 assert.equal(e.sheets.get('Chứng từ_full').getFilter().getRange().getNumColumns(),31);
 e.c.pnMirrorAll();
 const lost=e.sheets.get('Chứng từ_full').rows[4].slice(0,17);
 e.sheets.get('Chứng từ_full').getRange(5,1,1,31).clearContent();
 // Không chặn đồng bộ; báo động và chép lại vào full dòng còn ở tab chính.
 const r=e.c.pnMirrorAll();
 assert.match(String(e.props.PN_FULL_LOSS||''),/giảm/);
 assert.equal(r.ok,true,'main and full match again');
 assert.ok(e.sheets.get('Chứng từ_full').rows.some(x=>x&&x[4]===lost[4]&&x[1]===lost[1]),'lost history row restored from main');
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
test('One-step recovery restores shifted IDs, preserves primary edits and late orders',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();delete e.props.PN_FULL_READY;
 const main=e.sheets.get('Booking');main.rows[0][16]='__PN_ID';main.rows[0][17]='__PN_BASE';
 main.rows[1][16]=main.rows[1][29];main.rows[1][17]=main.rows[1][30];main.rows[1][29]='';main.rows[1][30]='';
 main.rows[1][14]='Nội thành';
 const late=main.rows[1].slice(0,16);late[1]='late';late[0]='30/09/2026';main.rows.splice(1,0,late);
 const result=e.c.pnRecoverFullMigration();assert.equal(result.ok,true);
 assert.equal(main.rows[2][29],main.rows[2][16]);assert.equal(main.rows[2][14],'Nội thành');
 assert.equal(e.sheets.get('Booking_full').rows.filter(r=>r[1]==='late').length,1);
});
test('One-step recovery refuses wrong moved identity before writing metadata',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();delete e.props.PN_FULL_READY;
 const main=e.sheets.get('Booking');main.rows[0][16]='__PN_ID';main.rows[0][17]='__PN_BASE';
 main.rows[1][16]=main.rows[1][29];main.rows[1][17]=main.rows[1][30];main.rows[1][29]='';main.rows[1][30]='';main.rows[1][1]='wrong-order';
 assert.throws(()=>e.c.pnRecoverFullMigration(),/không khớp hồ sơ/);assert.equal(main.rows[1][29],'');
});

// ===== Bộ đồng bộ v2: theo tên cột, chỉ xử lý dòng thay đổi, chính ⇄ full khớp 100% =====
function v2env(){const e=environment();seed(e);e.c.pnMigrateFull();return e;}
function auditOk(e){const r=e.run('PN_FULL.pairs.map(pnV2Audit_)');r.forEach(x=>assert.equal(x.ok,true,x.sheet+': '+x.errors.join('; ')));}
function rowOf(sh,col,val){return sh.rows.findIndex((r,i)=>i>0&&r&&String(r[col])===String(val))+1;}
function countWrites(fn){const orig=Range.prototype.setValues;let n=0;Range.prototype.setValues=function(rows){n++;return orig.call(this,rows);};try{fn();}finally{Range.prototype.setValues=orig;}return n;}
function edit(e,sheet,row,col,value){const sh=e.sheets.get(sheet);sh.getRange(row,col).setValue(value);return e.c.pnHandleEdit({range:sh.getRange(row,col)});}
function docRow(order,date,status){const r=doc(order,date,status);return r;}
test('V2 upgrade from legacy fingerprints: first run only rewrites base/ID, no business change; second run writes nothing',()=>{
 const e=v2env();
 const snap=['Booking','Booking_full','File đơn','File đơn_full','Chứng từ_FF','Chứng từ_full'].map(n=>JSON.stringify(e.sheets.get(n).rows.map(r=>(r||[]).slice(0,18))));
 e.c.pnMirrorAll();
 const after=['Booking','Booking_full','File đơn','File đơn_full','Chứng từ_FF','Chứng từ_full'].map(n=>JSON.stringify(e.sheets.get(n).rows.map(r=>(r||[]).slice(0,18))));
 assert.deepEqual(after,snap,'no business value changed by the upgrade');
 assert.equal(e.props.PN_ENGINE,'v2');auditOk(e);
 assert.equal(countWrites(()=>e.c.pnMirrorAll()),0,'unchanged data -> zero writes');
});
test('V2 columns matched by header name: reordering main columns causes no false change',()=>{
 const e=v2env();e.c.pnMirrorAll();
 const main=e.sheets.get('Chứng từ_FF');
 // Đổi chỗ cột "Khách Hàng"(C) và "Mã PO"(D) ở tab chính, kèm dữ liệu.
 main.rows.forEach(r=>{if(r){const t=r[2];r[2]=r[3];r[3]=t;}});
 const fullBefore=JSON.stringify(e.sheets.get('Chứng từ_full').rows.map(r=>(r||[]).slice(0,17)));
 e.c.pnMirrorAll();
 assert.equal(JSON.stringify(e.sheets.get('Chứng từ_full').rows.map(r=>(r||[]).slice(0,17))),fullBefore,'full untouched');
 auditOk(e);
});
test('V2 manual edit: main -> full and full -> main (current month), only edited row written',()=>{
 const e=v2env();e.c.pnMirrorAll();
 const main=e.sheets.get('Chứng từ_FF'),full=e.sheets.get('Chứng từ_full');
 const m=rowOf(main,4,'1'),f=rowOf(full,4,'1');
 edit(e,'Chứng từ_FF',m,10,'ghi chú từ tab chính');
 assert.equal(full.getRange(f,10).getValue(),'ghi chú từ tab chính');
 edit(e,'Chứng từ_full',f,11,'note từ full');
 assert.equal(main.getRange(m,11).getValue(),'note từ full');
 auditOk(e);
});
test('V2 new rows: main -> full with ID and Tháng; full row of current month -> main',()=>{
 const e=v2env();e.c.pnMirrorAll();
 const main=e.sheets.get('Chứng từ_FF'),full=e.sheets.get('Chứng từ_full');
 const nr=main.getLastRow()+1;main.getRange(nr,1,1,17).setValues([docRow(7,'02/10/2026','Chưa nhận chứng từ')]);
 e.c.pnHandleEdit({range:main.getRange(nr,1,1,17)});
 const fr=rowOf(full,4,'7');assert.ok(fr>1,'order 7 in full');
 assert.equal(full.getRange(fr,18).getValue(),'2026-10');assert.ok(String(full.getRange(fr,30).getValue()).length>5,'full has ID');
 const fnew=full.getLastRow()+1;full.getRange(fnew,1,1,17).setValues([docRow(8,'03/10/2026','Chưa nhận chứng từ')]);
 e.c.pnScheduledReconcile();
 assert.ok(rowOf(main,4,'8')>1,'order 8 copied to main');
 auditOk(e);
});
test('V2 empty or default-only rows are not records and are never copied to full',()=>{
 const e=v2env();e.c.pnMirrorAll();
 const main=e.sheets.get('Chứng từ_FF'),full=e.sheets.get('Chứng từ_full');
 const before=full.getLastRow(),nr=main.getLastRow()+1;
 const blank=Array(17).fill('');blank[6]='Chưa nhận chứng từ';blank[14]='Bình thường';blank[8]=false;
 main.getRange(nr,1,1,17).setValues([blank]);
 e.c.pnScheduledReconcile();
 assert.equal(full.getLastRow(),before,'no row added to full');
});
test('V2 copy-pasted row with the old hidden ID becomes a new record; original untouched',()=>{
 const e=v2env();e.c.pnMirrorAll();
 const main=e.sheets.get('Chứng từ_FF'),full=e.sheets.get('Chứng từ_full');
 const src=rowOf(main,4,'1'),copy=main.rows[src-1].slice();copy[4]='9';copy[1]='G9';
 const nr=main.getLastRow()+1;main.rows[nr-1]=copy;
 e.c.pnScheduledReconcile();
 assert.ok(rowOf(full,4,'9')>1,'copy added as new full record');
 assert.notEqual(main.getRange(nr,30).getValue(),main.getRange(src,30).getValue(),'copy got its own ID');
 assert.equal(full.getRange(rowOf(full,4,'1'),2).getValue(),'G1','original record unchanged');
 auditOk(e);
});
test('V2 new main column is added to full by name and synced',()=>{
 const e=v2env();e.c.pnMirrorAll();
 const main=e.sheets.get('Chứng từ_FF'),full=e.sheets.get('Chứng từ_full');
 main.getRange(1,18).setValue('Cột mới');main.getRange(2,18).setValue('giá trị mới');
 e.c.pnScheduledReconcile();
 const col=full.rows[0].indexOf('Cột mới')+1;assert.ok(col>18,'header added in a free column of full');
 assert.equal(full.getRange(rowOf(full,1,main.getRange(2,2).getValue()),col).getValue(),'giá trị mới');
 auditOk(e);
});
test('V2 daily cleanup: received/cancelled old docs and old Booking/File đơn leave main; unreceived stays; reverted status returns',()=>{
 const e=v2env();e.props.PN_FULL_CLEANUP='enabled';
 assert.equal(e.c.pnCleanupDaily().ok,true);
 const main=e.sheets.get('Chứng từ_FF'),full=e.sheets.get('Chứng từ_full');
 assert.ok(rowOf(main,4,'2')>1,'old unreceived kept');
 assert.equal(rowOf(main,4,'3'),0,'old received removed');assert.equal(rowOf(main,4,'4'),0,'old cancelled removed');
 assert.ok(rowOf(full,4,'3')>1&&rowOf(full,4,'4')>1,'history kept in full');
 assert.equal(e.sheets.get('File đơn').getRange(2,20).getValue(),'pivot','pivot untouched');
 edit(e,'Chứng từ_full',rowOf(full,4,'3'),7,'Chưa nhận chứng từ');
 assert.ok(rowOf(main,4,'3')>1,'reverted to unreceived -> back to main');
 auditOk(e);
});
test('V2 API save writes only the right order, identical in full and main, even with reordered columns',()=>{
 const e=v2env();e.c.pnMirrorAll();
 e.run("uploadFiles_=()=>({linkAnh:'https://example.invalid/p',files:[{id:'1'}],folderUrl:''});");
 const main=e.sheets.get('Chứng từ_FF'),full=e.sheets.get('Chứng từ_full');
 main.rows.forEach(r=>{if(r){const t=r[9];r[9]=r[10];r[10]=t;}}); // đổi chỗ "Ghi chú chứng từ" và "Note" ở tab chính
 e.c.pnMirrorAll();
 const others=()=>JSON.stringify([main,full].map(sh=>sh.rows.filter((r,i)=>i>0&&r&&String(r[4])!=='2').map(r=>r.slice(0,17))));
 const before=others();
 const res=e.c.apiSave_({clientId:'v2api',maDon:'G2',returnType:'Chứng từ',xacThuc:'Đã nhận chứng từ',note:'ghi chú app',maEcomCt:'EC-9',files:[{base64:'AA=='}]});
 assert.equal(res.ok,true);
 assert.equal(others(),before,'no other order changed');
 const fr=rowOf(full,4,'2'),mr=rowOf(main,4,'2');
 assert.equal(full.getRange(fr,7).getValue(),'Đã nhận chứng từ');assert.equal(main.getRange(mr,7).getValue(),'Đã nhận chứng từ');
 assert.equal(full.getRange(fr,10).getValue(),'ghi chú app');assert.equal(main.getRange(mr,11).getValue(),'ghi chú app','note by header in moved main column');
 assert.equal(full.getRange(fr,12).getValue(),'EC-9');assert.equal(main.getRange(mr,12).getValue(),'EC-9');
 auditOk(e);
});
test('V2 busy lock: scheduled run and edit return quietly; edit queues the exact rows',()=>{
 const e=v2env();
 e.run("LockService={getScriptLock:()=>({tryLock(){return false;},waitLock(){},releaseLock(){}})};");
 assert.equal(e.c.pnScheduledReconcile().busy,true);
 const sh=e.sheets.get('Chứng từ_FF');
 assert.equal(e.c.pnHandleEdit({range:sh.getRange(2,10)}).busy,true);
 assert.ok(Object.keys(e.props).some(k=>k.indexOf('PN_EDITQ_')===0),'edited rows queued, not dropped');
});
test('V2 pipeline: new Booking row creates the order in Chứng từ_full and Chứng từ_FF; new File đơn fills GHTK in both',()=>{
 const e=v2env();e.c.pnMirrorAll();
 const bk=e.sheets.get('Booking'),nr=bk.getLastRow()+1,row=Array(16).fill('');
 row[0]='02/10/2026';row[1]='77';row[3]='Customer X';row[4]='PO77';row[5]='Addr';row[9]=10;
 bk.getRange(nr,1,1,16).setValues([row]);e.c.pnHandleEdit({range:bk.getRange(nr,1,1,16)});
 const main=e.sheets.get('Chứng từ_FF'),full=e.sheets.get('Chứng từ_full');
 assert.ok(rowOf(e.sheets.get('Booking_full'),1,'77')>1,'booking in Booking_full');
 assert.ok(rowOf(full,4,'77')>1,'order in Chứng từ_full');assert.ok(rowOf(main,4,'77')>1,'order in Chứng từ_FF');
 const fd=e.sheets.get('File đơn'),fr=fd.getLastRow()+1,frow=Array(18).fill('');
 frow[0]='GH77';frow[1]='2026-10-02 10:00:00';frow[6]='Depot';frow[10]='77';
 fd.getRange(fr,1,1,18).setValues([frow]);e.c.pnHandleEdit({range:fd.getRange(fr,1,1,18)});
 assert.equal(full.getRange(rowOf(full,4,'77'),2).getValue(),'GH77');
 assert.equal(main.getRange(rowOf(main,4,'77'),2).getValue(),'GH77');
 auditOk(e);
});
test('V2 first run after upgrade picks the side that really changed (legacy base)',()=>{
 const e=v2env();
 const main=e.sheets.get('Chứng từ_FF'),full=e.sheets.get('Chứng từ_full');
 main.getRange(rowOf(main,4,'1'),10).setValue('sửa ở chính trước nâng cấp');
 full.getRange(rowOf(full,4,'2'),11).setValue('sửa ở full trước nâng cấp');
 e.c.pnMirrorAll();
 assert.equal(full.getRange(rowOf(full,4,'1'),10).getValue(),'sửa ở chính trước nâng cấp');
 assert.equal(main.getRange(rowOf(main,4,'2'),11).getValue(),'sửa ở full trước nâng cấp');
 assert.ok(!e.sheets.get('_PN_CONFLICTS'),'no false conflict');
 auditOk(e);
});
test('V2 edited side wins when both sides changed; losing copy is logged',()=>{
 const e=v2env();e.c.pnMirrorAll();
 const main=e.sheets.get('Chứng từ_FF'),full=e.sheets.get('Chứng từ_full');
 const m=rowOf(main,4,'1'),f=rowOf(full,4,'1');
 main.getRange(m,10).setValue('bản chính');
 edit(e,'Chứng từ_full',f,10,'bản full vừa sửa tay');
 assert.equal(main.getRange(m,10).getValue(),'bản full vừa sửa tay');
 assert.match(JSON.stringify(e.sheets.get('_PN_CONFLICTS').rows),/bản chính/);
 auditOk(e);
});
test('V2 scale: with 3000 rows per tab, Sheet calls stay constant (not per row); one edit writes only that row',()=>{
 const e=environment();seed(e);
 const docs=e.sheets.get('Chứng từ_FF'),bk=e.sheets.get('Booking'),fd=e.sheets.get('File đơn');
 for(let i=100;i<3100;i++){
  docs.rows.push(doc(i,'0'+(1+i%9)+'/10/2026','Chưa nhận chứng từ'));
  const b=Array(16).fill('');b[0]='01/10/2026';b[1]=String(i);b[9]=i%40;bk.rows.push(b);
  const f=Array(18).fill('');f[0]='G'+i;f[1]='2026-10-01 08:00:00';f[6]='Depot';f[10]=String(i);fd.rows.push(f);
 }
 e.c.pnMigrateFull();e.c.pnMirrorAll();
 const calls={read:0,write:0};
 const gv=Range.prototype.getValues,gd=Range.prototype.getDisplayValues,sv=Range.prototype.setValues;
 Range.prototype.getValues=function(){calls.read++;return gv.call(this);};
 Range.prototype.getDisplayValues=function(){calls.read++;return gd.call(this);};
 Range.prototype.setValues=function(r){calls.write++;return sv.call(this,r);};
 try {
  e.c.pnScheduledReconcile();
  assert.equal(calls.write,0,'no change -> no write');
  assert.ok(calls.read<80,'reads independent of row count: '+calls.read);
  calls.read=0;calls.write=0;
  const m=rowOf(docs,4,'2500');docs.getRange(m,10).setValue('sửa 1 dòng');
  e.c.pnHandleEdit({range:docs.getRange(m,10)});
  assert.ok(calls.write<=6,'one edited row -> few writes: '+calls.write);
  assert.ok(calls.read<80,'reads independent of row count: '+calls.read);
 } finally { Range.prototype.getValues=gv;Range.prototype.getDisplayValues=gd;Range.prototype.setValues=sv; }
 assert.equal(e.sheets.get('Chứng từ_full').getRange(rowOf(e.sheets.get('Chứng từ_full'),4,'2500'),10).getValue(),'sửa 1 dòng');
 auditOk(e);
});
test('No time trigger: first Sheet edit of the day runs the daily cleanup once; later edits do not',()=>{
 const e=v2env();e.props.PN_FULL_CLEANUP='enabled';e.c.pnMirrorAll();
 const main=e.sheets.get('Chứng từ_FF');
 edit(e,'Chứng từ_FF',rowOf(main,4,'1'),10,'sửa đầu ngày');
 assert.equal(rowOf(main,4,'3'),0,'old received order cleaned on first edit of the day');
 assert.ok(e.props.PN_DAILY_DONE,'daily marker set');
 let cleaned=0;e.run("var __c=pnV2Cleanup_;pnV2Cleanup_=function(){globalThis.__n=(globalThis.__n||0)+1;return __c();};");
 edit(e,'Chứng từ_FF',rowOf(main,4,'1'),10,'sửa lần 2');
 assert.equal(e.run('globalThis.__n||0'),0,'no second cleanup the same day');
 auditOk(e);
});
test('Cleanup moves each row together with its colours/dropdowns; vacated rows are wiped clean; new rows take no colour',()=>{
 const e=v2env();e.props.PN_FULL_CLEANUP='enabled';e.c.pnMirrorAll();
 const main=e.sheets.get('Chứng từ_FF');
 const rG3=rowOf(main,4,'3'),rG2=rowOf(main,4,'2'),last=main.getLastRow();
 main.getRange(rG3,1,1,17).setBackgrounds([Array(17).fill('#ff0000')]);       // đơn sẽ bị dọn: tô đỏ
 main.getRange(rG2,1,1,17).setBackgrounds([Array(17).fill('#00ff00')]);       // đơn giữ lại: tô xanh
 main.getRange(2,1,last-1,17).setDataValidations(Array.from({length:last-1},()=>Array(17).fill('dropdown')));
 main.rows.push(Array(17).fill(''));main.rows[main.rows.length-1][8]=false;   // dòng mẫu trống chỉ có checkbox
 main.getRange(main.rows.length,1,1,17).setDataValidations([Array(17).fill('dropdown')]);
 e.c.pnCleanupDaily();
 const g2=rowOf(main,4,'2');
 assert.equal(main.getRange(g2,1).getBackgrounds()[0][0],'#00ff00','kept row keeps its own colour');
 const colours=main.getRange(2,1,main.getLastRow()+3,17).getBackgrounds().flat();
 assert.ok(!colours.includes('#ff0000'),'colour of removed order not left on any row');
 const firstEmpty=main.getLastRow()+1;
 assert.equal(main.getRange(firstEmpty,1).getDataValidations()[0][0],null,'vacated row has no dropdown/checkbox');
 assert.equal(main.getRange(firstEmpty,1).getBackgrounds()[0][0],null,'vacated row has no colour');
 const nr=main.getLastRow()+1;main.getRange(nr,1,1,17).setValues([doc(55,'05/10/2026','Chưa nhận chứng từ')]);
 const full=e.sheets.get('Chứng từ_full'),fr=full.getLastRow()+1;full.getRange(fr,1,1,17).setValues([doc(56,'05/10/2026','Chưa nhận chứng từ')]);
 e.c.pnScheduledReconcile();
 const r56=rowOf(main,4,'56');
 assert.equal(main.getRange(r56,1).getBackgrounds()[0][0],null,'new row from full takes no colour of another order');
 assert.equal(main.getRange(r56,1).getDataValidations()[0][0],'dropdown','new row takes the template dropdown');
 auditOk(e);
});
test('API save and a manual edit on Chứng từ read only the ID column + the one row of the main tab (3000 rows)',()=>{
 const e=environment();seed(e);const docs=e.sheets.get('Chứng từ_FF');
 for(let i=100;i<3100;i++)docs.rows.push(doc(i,'01/10/2026','Chưa nhận chứng từ'));
 e.c.pnMigrateFull();e.c.pnMirrorAll();
 e.run("uploadFiles_=()=>({linkAnh:'https://example.invalid/p',files:[{id:'1'}],folderUrl:''});");
 let cells=0;const gv=Range.prototype.getValues,gd=Range.prototype.getDisplayValues;
 Range.prototype.getValues=function(){if(this.sh.name==='Chứng từ_FF')cells+=this.n*this.w;return gv.call(this);};
 Range.prototype.getDisplayValues=function(){if(this.sh.name==='Chứng từ_FF')cells+=this.n*this.w;return gd.call(this);};
 try {
  const res=e.c.apiSave_({clientId:'scale-api',maDon:'G2500',returnType:'Chứng từ',xacThuc:'Đã nhận chứng từ',files:[{base64:'AA=='}]});
  assert.equal(res.ok,true);
  assert.ok(cells<3*3100,'API read of main tab is ~ID column only, not whole tab: '+cells);
  cells=0;
  const m=rowOf(docs,4,'2600');docs.getRange(m,10).setValue('sửa tay');e.c.pnHandleEdit({range:docs.getRange(m,10)});
  assert.ok(cells<3*3100,'manual edit read of main tab is ~ID column only: '+cells);
 } finally { Range.prototype.getValues=gv;Range.prototype.getDisplayValues=gd; }
 const full=e.sheets.get('Chứng từ_full');
 assert.equal(docs.getRange(rowOf(docs,4,'2500'),7).getValue(),'Đã nhận chứng từ','main updated from API write');
 assert.equal(full.getRange(rowOf(full,4,'2600'),10).getValue(),'sửa tay','manual edit reached full');
 auditOk(e);
});
test('Pasting 20 Booking rows into 3000 existing: only those rows go to Booking_full; their orders appear in both Chứng từ tabs; few writes',()=>{
 const e=environment();seed(e);
 const bk=e.sheets.get('Booking'),docs=e.sheets.get('Chứng từ_FF');
 for(let i=100;i<3100;i++){const b=Array(16).fill('');b[0]='01/10/2026';b[1]=String(i);b[9]=5;bk.rows.push(b);docs.rows.push(doc(i,'01/10/2026','Chưa nhận chứng từ'));}
 e.c.pnMigrateFull();e.c.pnMirrorAll();
 const start=bk.getLastRow()+1,paste=[];
 for(let i=0;i<20;i++){const b=Array(16).fill('');b[0]='02/10/2026';b[1]=String(9000+i);b[3]='Cust';b[4]='PO'+i;b[5]='Addr';b[9]=10;paste.push(b);}
 bk.getRange(start,1,20,16).setValues(paste);
 let writes=0;const sv=Range.prototype.setValues;Range.prototype.setValues=function(r){writes++;return sv.call(this,r);};
 try { e.c.pnHandleEdit({range:bk.getRange(start,1,20,16)}); } finally { Range.prototype.setValues=sv; }
 const bf=e.sheets.get('Booking_full'),full=e.sheets.get('Chứng từ_full');
 for(let i=0;i<20;i++){
  assert.ok(rowOf(bf,1,String(9000+i))>1,'booking '+(9000+i)+' in Booking_full');
  assert.ok(rowOf(full,4,String(9000+i))>1,'order in Chứng từ_full');
  assert.ok(rowOf(docs,4,String(9000+i))>1,'order in Chứng từ_FF');
 }
 assert.ok(writes<=25,'batched writes, not per row of the whole tab: '+writes);
 auditOk(e);
});
test('Edit made while the lock is busy is not lost: the edited rows sync at the next Sheet activity, app save does not touch it',()=>{
 const e=v2env();e.c.pnMirrorAll();
 e.run("uploadFiles_=()=>({linkAnh:'https://example.invalid/p',files:[{id:'1'}],folderUrl:''});");
 const main=e.sheets.get('Chứng từ_FF'),full=e.sheets.get('Chứng từ_full');
 e.run("var __lock=LockService;LockService={getScriptLock:()=>({tryLock(){return false;},waitLock(){},releaseLock(){}})};");
 const m=rowOf(main,4,'1');main.getRange(m,12).setValue('EC-TAY');
 const r=e.c.pnHandleEdit({range:main.getRange(m,12)});
 assert.equal(r.queued,true);
 assert.notEqual(full.getRange(rowOf(full,4,'1'),12).getValue(),'EC-TAY','not yet synced while busy');
 e.run("LockService=__lock;");
 assert.equal(e.c.apiSave_({clientId:'q1',maDon:'G2',returnType:'Chứng từ',files:[{base64:'AA=='}]}).ok,true);
 assert.notEqual(full.getRange(rowOf(full,4,'1'),12).getValue(),'EC-TAY','app save does not process Sheet edits');
 const g=rowOf(main,4,'2');e.c.pnHandleEdit({range:main.getRange(g,13)});
 assert.equal(full.getRange(rowOf(full,4,'1'),12).getValue(),'EC-TAY','queued edit synced by the next Sheet activity');
 assert.ok(!Object.keys(e.props).some(k=>k.indexOf('PN_EDITQ_')===0),'queue drained');
 auditOk(e);
});
test('Journal is append-only: latest state wins, interleaved requests do not overwrite each other',()=>{
 const e=environment();seed(e);e.c.pnMigrateFull();
 e.c.pnJournalWrite_('request:aa','PENDING',{n:1});
 e.c.pnJournalWrite_('request:bb','PENDING',{n:2});
 e.c.pnJournalWrite_('request:aa','DONE',{n:3});
 assert.equal(e.c.pnJournalRead_('request:aa').state,'DONE');
 assert.equal(e.c.pnJournalRead_('request:aa').data.n,3);
 assert.equal(e.c.pnJournalRead_('request:bb').data.n,2);
 assert.deepEqual(Array.from(e.c.pnPendingRequests_()),['request:bb']);
});
test('Mã ecom CT keeps only the last 10 digits',()=>{
 const e=environment();
 assert.equal(e.c.pnEcomCt_('https://i.ghtk.vn/S22843210.MB1.A12.1234567890'),'1234567890');
 assert.equal(e.c.pnEcomCt_(' 0987654321 '),'0987654321');
 assert.equal(e.c.pnEcomCt_('AB123'),'AB123');
});
console.log('RESULT '+passed+' tests passed.');

const fs=require('fs'),path=require('path'),cp=require('child_process'),crypto=require('crypto');
const root=__dirname;
const manifest=JSON.parse(fs.readFileSync(path.join(root,'SOURCE_MANIFEST.json')));
const outputs=[];
for(const m of manifest) {
 const result=cp.spawnSync('git',['diff','--no-index','--no-ext-diff','--',path.join(root,'originals',m.file),path.join(root,'apps-script',m.file)],{encoding:'utf8'});
 if(result.status!==0&&result.status!==1)throw new Error(result.stderr||'git diff unavailable');
 outputs.push(result.stdout);
}
const app=cp.spawnSync('git',['diff','--no-index','--no-ext-diff','--',path.join(root,'originals','frontend-index.html'),path.join(root,'frontend','index.html')],{encoding:'utf8'});
if(app.status!==0&&app.status!==1)throw new Error(app.stderr);
outputs.push(app.stdout);
fs.writeFileSync(path.join(root,'CHANGES.patch'),outputs.join('\n'));
const files=['README.md','FORMULA_AUDIT.json','frontend/index.html',...fs.readdirSync(path.join(root,'apps-script')).filter(x=>x.endsWith('.gs')).map(x=>'apps-script/'+x)];
fs.writeFileSync(path.join(root,'OUTPUT_MANIFEST.json'),JSON.stringify(files.map(file=>({file,sha256:crypto.createHash('sha256').update(fs.readFileSync(path.join(root,file))).digest('hex')})),null,2));
console.log('Review diff and output checksums written.');

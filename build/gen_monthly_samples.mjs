import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import {fileURLToPath, pathToFileURL} from 'node:url';
import {execFileSync} from 'node:child_process';
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const runtime=createRequire(path.join(root,'build/out/monthly-samples/loader.cjs'));
const {Workbook,SpreadsheetFile,FileBlob}=await import(pathToFileURL(runtime.resolve('@oai/artifact-tool')).href);
const iconv={decode:(b)=>new TextDecoder('shift_jis',{fatal:true}).decode(b),encode:(s)=>execFileSync('C:/Users/ynisi/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe',['-c','import sys;sys.stdout.buffer.write(sys.stdin.buffer.read().decode("utf-8").encode("cp932"))'],{input:Buffer.from(s,'utf8'),windowsHide:true})};
const base=path.join(root,'samples/current');
const input=path.join(base,'data');
const updateName='追加CSVデータ/5月分（3件追加）';
const update=path.join(base,updateName);
await fs.mkdir(update,{recursive:true});
const cfg=JSON.parse(await fs.readFile(path.join(root,'configs/sample/settings.json'),'utf8'));
const sourceNames=Object.fromEntries(Object.entries(cfg.data.tables).map(([k,v])=>[k,v.file]));
const csvKeys=['TXN','PAY','DEL'];
const columns={TXN:29,PAY:15,DEL:2};
const rows={};
for(const key of csvKeys){
  const book=await Workbook.fromCSV(iconv.decode(await fs.readFile(path.join(input,sourceNames[key])),'cp932'),{sheetName:'入力'});
  rows[key]=book.worksheets.getItemAt(0).getRangeByIndexes(0,0,6,columns[key]).values;
}
const original=sourceNames.APP.endsWith('.csv')
  ? await Workbook.fromCSV(iconv.decode(await fs.readFile(path.join(input,sourceNames.APP))),{sheetName:'入力'})
  : await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(input,sourceNames.APP)));
rows.APP=original.worksheets.getItemAt(0).getRange('A1:S6').values;
const clone=(v)=>JSON.parse(JSON.stringify(v));
const month=(value,n)=> typeof value==='string' ? value.replaceAll('FT260901',`FT260${n}15`).replaceAll('2026/09/01',`2026/0${n}/15`).replaceAll('2026/09/30',`2026/0${n}/${n===4?'30':'31'}`) : value;
const card=n=>'AB'+String(10000000+n)+'CD';
const application=n=>'受付'+(n%2?'AA':'BB')+'26-'+String(n).padStart(6,'0');
function newRow(row,n){
  const letter=String.fromCharCode(64+n);
  return row.map(v=>typeof v==='string' ? v.replaceAll(card(1),card(n)).replaceAll(application(1),application(n))
    .replaceAll('案件A','案件'+letter).replaceAll('利用者A','利用者'+letter).replaceAll('申込者A','申込者'+letter)
    .replaceAll('TEST0001','TEST'+String(n).padStart(4,'0')).replaceAll('USER0001','USER'+String(n).padStart(4,'0'))
    .replaceAll('TESTPAY0001','TESTPAY'+String(n).padStart(4,'0')).replaceAll('TESTBILL0001','TESTBILL'+String(n).padStart(4,'0'))
    .replaceAll('TESTMASTER0001','TESTMASTER'+String(n).padStart(4,'0')).replaceAll('sample1@','sample'+n+'@') : v);
}
const april=clone(rows),may={};
for(const key of csvKeys) april[key]=april[key].map(row=>row.map(v=>month(v,4)));
for(let i=1;i<6;i++) april.APP[i][17]='2026/04/10';
for(const key of ['TXN','PAY']){
  may[key]=[rows[key][0],clone(rows[key][2]),clone(rows[key][3]),...[9,10,11].map(n=>newRow(rows[key][1],n))].map(row=>row.map(v=>month(v,5)));
}
may.PAY[1][4]='決済済';
may.APP=[rows.APP[0],clone(rows.APP[2]),newRow(rows.APP[1],3),...[9,10,11].map(n=>newRow(rows.APP[1],n))];
for(let i=1;i<6;i++) may.APP[i][17]=i<3?'2026/04/10':'2026/05/10';
may.DEL=[rows.DEL[0],...[2,3,9,10,11].map(n=>[card(n),application(n)])];
const nameFor=(key,n)=>sourceNames[key].replace(/(?:_[45]月分)?\.[^.]+$/u,'.csv');
const remarksIndex=rows.APP[0].indexOf('意見欄');
if(remarksIndex<0) throw new Error('意見欄 column missing');
for(const [n,set] of [[4,april],[5,may]]){
  for(let i=1;i<set.APP.length;i++){
    set.APP[i][remarksIndex]=`${n}月分の申込内容と提出資料を確認しました。\n連絡は平日の午後を希望しています。\n日程が変わる場合は、候補日を二つ案内してください。\n次回は変更後の計画を確認します。以上は架空の内容です。`;
  }
}
async function writeSet(set,n,directory){
  for(const key of [...csvKeys,'APP']){
    const csv=set[key].map(row=>row.map(v=>{const s=v==null?'':String(v);return /[",\r\n]/u.test(s)?'"'+s.replaceAll('"','""')+'"':s;}).join(',')).join('\r\n')+'\r\n';
    await fs.writeFile(path.join(directory,nameFor(key,n)),iconv.encode(csv,'cp932'));
  }
}
await writeSet(april,4,input);
await writeSet(may,5,update);
for(const [key,table] of Object.entries(cfg.data.tables)) table.file=nameFor(key,4);
cfg.data.tables.APP.encoding='shift_jis';
delete cfg.data.tables.APP.sheet;
await fs.writeFile(path.join(root,'configs/sample/settings.json'),JSON.stringify(cfg,null,2)+'\n');
await fs.writeFile(path.join(root,'settings.json'),JSON.stringify(cfg,null,2)+'\n');
for(const [key,old] of Object.entries(sourceNames)) if(old!==nameFor(key,4)) await fs.unlink(path.join(input,old));
console.log('Created April and May input sets: four files x five rows each. May contains B/C plus three new cases I/J/K.');
console.log(update);

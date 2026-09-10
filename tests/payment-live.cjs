'use strict';
const fs=require('fs');
const path=require('path');
const {CdpClient,waitFor,delay}=require('../build/webview2_cdp');
const [port,evidence]=process.argv.slice(2);
const info=JSON.parse(fs.readFileSync(path.join(evidence,'case.json'),'utf8').replace(/^\uFEFF/,''));
const results=[];
function check(name,value){results.push({name,pass:!!value});if(!value)throw Error(name);console.log('PASS '+name);}
async function connect(dialog){
  const until=Date.now()+45000;
  while(Date.now()<until){try{
    const list=await(await fetch(`http://127.0.0.1:${port}/json/list`)).json();
    const target=list.find(x=>x.type==='page'&&x.url.includes('reader-data-viewer.local')&&x.url.includes('#dialog')===dialog);
    if(target){const client=new CdpClient(target.webSocketDebuggerUrl);await client.open();await client.send('Runtime.enable');return client;}
  }catch(e){}await delay(60);}throw Error('Reader did not start');
}
async function main(){let app,modal;try{
  app=await connect(false);modal=await connect(true);
  const ready=()=>waitFor(app,`document.querySelector('#b-upd[aria-disabled=false]')!==null`,20000,'main enabled');
  const click=(client,selector)=>client.evaluate(`document.querySelector(${JSON.stringify(selector)}).click();true`);
  const visible=id=>waitFor(modal,`document.querySelector('#${id}.show')!==null`,10000,id);
  const capture=async(client,name)=>{await delay(200);const shot=await client.send('Page.captureScreenshot',{format:'png',captureBeyondViewport:false});fs.writeFileSync(path.join(evidence,name+'.png'),Buffer.from(shot.data,'base64'));};
  const search=async key=>{await app.evaluate(`(()=>{const n=document.querySelector('#input');n.textContent=${JSON.stringify(key)};n.dispatchEvent(new Event('input',{bubbles:true}));})()`);await click(app,'#b-search');};
  const masked=['s1.item1.row0','s1.item1.row1','s1.item1.row2','s1.item1.row3','s1.item1.row4','s1.item1.row5','s1.item1.row6','s2.value','s3.value'];
  const deleteOnce=async()=>{
    await click(app,'#b-del');await visible('v-del');
    await click(modal,'#v-del [data-modal-default=true]');
    await waitFor(app,`document.querySelector('#b-upd[aria-disabled=true]')!==null`,10000,'deletion started');
    await ready();
  };
  // Keep the deliberately different candidate fixture when startup offers an update.
  for(let i=0;i<150;i++){
    if(await app.evaluate(`document.querySelector('#b-upd[aria-disabled=false]')!==null`))break;
    await modal.evaluate(`(()=>{const buttons=document.querySelectorAll('#v-send.show .foot .btn');if(buttons.length===2)buttons[1].click();})()`);
    await delay(100);
  }
  await ready();await capture(app,'main-empty');
  check('legacy JSON layout loaded',await app.evaluate(`document.querySelector('[data-bind="s1.item0.row0"]')!==null`));
  await search(info.paidKey);await visible('v-cand');
  check('two real candidate rows',await modal.evaluate(`document.querySelectorAll('#v-cand tbody tr').length===2`));
  await capture(modal,'candidates');
  await modal.evaluate(`document.querySelector('#v-cand tbody tr').dispatchEvent(new MouseEvent('dblclick',{bubbles:true}));true`);
  await ready();await waitFor(app,`document.querySelector('#judge').textContent==='済'`,10000,'paid');
  check('paid detail bindings',await app.evaluate(`${JSON.stringify(masked)}.every(id=>document.querySelector('[data-bind="'+id+'"]').textContent.length>0)`));
  await capture(app,'main-paid');
  await click(app,'#b-work');await waitFor(app,`document.querySelector('[data-bind="s5.value"]').textContent==='未送信 1 件'`,10000,'one pending');
  await deleteOnce();
  check('unsent done status cannot delete a shared unprocessed row',await app.evaluate(`document.querySelector('[data-bind="s6.segment1"]').textContent==='台帳件数 100' && document.querySelector('[data-bind="s5.value"]').textContent==='未送信 1 件'`));
  await click(app,'#b-send');await visible('v-send');await click(modal,'#v-send .foot .btn');
  await ready();await waitFor(app,`document.querySelector('[data-bind="s5.value"]').textContent==='未送信 0 件'`,10000,'sent');
  fs.copyFileSync(info.ledger,path.join(evidence,'sent-ledger.xlsx'));
  check('sent changes clear the display',await app.evaluate(`document.querySelector('[data-bind="s1.item1.row1"]').textContent===''`));
  await search(info.noAppKey);await waitFor(app,`document.querySelector('#judge').textContent==='済'`,10000,'paid without APP');
  check('paid without APP remains searchable',await app.evaluate(`document.querySelector('[data-bind="s1.item1.row1"]').textContent!=='' && document.querySelector('[data-bind="s1.item1.row3"]').textContent===''`));
  await capture(app,'main-no-application');
  await search(info.missingKey);await waitFor(app,`document.querySelector('#judge').textContent==='未決済'`,10000,'no payment');
  check('unpaid masks nine details',await app.evaluate(`${JSON.stringify(masked)}.every(id=>document.querySelector('[data-bind="'+id+'"]').textContent==='')`));
  await capture(app,'main-unpaid');
  await click(app,'#b-set');await visible('v-set');await capture(modal,'settings');
  await click(modal,'#v-set [data-modal-default=true]');await ready();
  check('settings save retains legacy layout',!!JSON.parse(fs.readFileSync(info.settings,'utf8').replace(/^\uFEFF/,'' )).screen.sections);
  await click(app,'#b-out');await visible('v-out');await capture(modal,'export');
  const exportPath=await modal.evaluate(`document.querySelector('#v-out [data-field=exportPath]').textContent.trim()`);
  await click(modal,'#v-out [data-modal-default=true]');await ready();
  const exported=path.isAbsolute(exportPath)?exportPath:path.join(info.scratch,exportPath);
  for(let i=0;i<100&&!fs.existsSync(exported);i++)await delay(100);
  check('export creates a nonempty CSV',fs.existsSync(exported)&&fs.statSync(exported).size>0);
  fs.copyFileSync(exported,path.join(evidence,'export.csv'));
  for(const [button,id,count] of [['#b-upd','v-upd',100],['#b-del','v-del',99]]){
    await click(app,button);await visible(id);await capture(modal,id==='v-upd'?'update':'delete');
    const overlaps=await modal.evaluate(`(()=>{const body=document.querySelector('#${id} .body');return Array.from(body.querySelectorAll('fieldset')).filter(f=>{const l=f.querySelector('legend').getBoundingClientRect(),r=f.getBoundingClientRect();return l.top<r.top||l.bottom>r.bottom;}).length;})()`);
    check(id+' legends stay inside frames',overlaps===0);
    await click(modal,`#${id} [data-modal-default=true]`);
    await waitFor(app,`document.querySelector('#b-upd[aria-disabled=true]')!==null`,10000,'process started');
    // Source update may show an informational reset summary on this fixture.
    for(let i=0;i<100;i++){
      if(await app.evaluate(`document.querySelector('#b-upd[aria-disabled=false]')!==null`))break;
      await modal.evaluate(`(()=>{const b=document.querySelector('#v-shared.show .foot .btn')||document.querySelector('#v-send.show .foot .btn');if(b)b.click();})()`);
      await delay(100);
    }
    await ready();check(id+' row count',await app.evaluate(`document.querySelector('[data-bind="s6.segment1"]').textContent===${JSON.stringify('台帳件数 '+String(count))}`));
  }
  fs.copyFileSync(info.ledger,path.join(evidence,'deleted-ledger.xlsx'));
  await capture(app,'main-after-delete');
}finally{
  fs.writeFileSync(path.join(evidence,'ui-results.json'),JSON.stringify(results,null,2));
  if(modal)modal.close();if(app){try{await app.evaluate(`window.chrome.webview.postMessage({type:'window',command:'close'});true`);}catch(e){}app.close();}
}}
main().catch(e=>{console.error(e);process.exitCode=1;});

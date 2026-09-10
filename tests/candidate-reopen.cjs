'use strict';
// CDP drives only after the companion PowerShell check proves native visibility.
const fs = require('fs'), path = require('path');
const [evidence, root] = process.argv.slice(2);
const info = JSON.parse(fs.readFileSync(path.join(evidence, 'case.json'), 'utf8').replace(/^\uFEFF/, ''));
const { CdpClient, delay, press, waitFor } = require(path.join(root, 'build/webview2_cdp.js'));
let app, modal, requestId = 0;
const results = [];
async function connect(dialog) {
  const until = Date.now() + 45000;
  while (Date.now() < until) {
    try {
      const list = await (await fetch(`http://127.0.0.1:${info.port}/json/list`, { signal: AbortSignal.timeout(2000) })).json();
      const t = list.find(x => x.type === 'page' && x.url.includes('reader-data-viewer.local') && x.url.includes('#dialog') === dialog);
      if (t) { const c = new CdpClient(t.webSocketDebuggerUrl); await c.open(); await c.send('Runtime.enable'); return c; }
    } catch (_) {}
    await delay(60);
  }
  throw Error('Reader did not start');
}
async function native(kind, cycle) {
  const id = ++requestId;
  fs.writeFileSync(path.join(evidence, 'native-request.json'), JSON.stringify({ id, kind, cycle }));
  const until = Date.now() + 15000;
  while (Date.now() < until) {
    try {
      const r = JSON.parse(fs.readFileSync(path.join(evidence, 'native-reply.json'), 'utf8'));
      if (r.id === id) {
        results.push(r);
        if (!r.pass) throw Error('NATIVE FAILURE ' + kind + ' cycle ' + cycle + ': ' + JSON.stringify(r));
        console.log('PASS native ' + kind + ' cycle ' + cycle); return;
      }
    } catch (e) { if (String(e.message).startsWith('NATIVE FAILURE')) throw e; }
    await delay(60);
  }
  throw Error('Native inspector did not reply');
}
async function click(client, selector, count = 1) {
  const point = await client.evaluate(`(()=>{const r=document.querySelector(${JSON.stringify(selector)}).getBoundingClientRect();return {x:r.x+r.width/2,y:r.y+r.height/2};})()`);
  for (let i = 1; i <= count; i++) {
    await client.send('Input.dispatchMouseEvent', { type: 'mousePressed', ...point, button: 'left', clickCount: i });
    await client.send('Input.dispatchMouseEvent', { type: 'mouseReleased', ...point, button: 'left', clickCount: i });
  }
}
async function main() {
  try {
    app = await connect(false); modal = await connect(true);
    const ready = `document.querySelector('#b-upd[aria-disabled=false]')!==null`;
    for (let i = 0; i < 200; i++) {
      if (await app.evaluate(ready)) break;
      await modal.evaluate(`(()=>{const b=document.querySelectorAll('#v-send.show .foot .btn');if(b.length===2)b[1].click();})()`);
      await delay(100);
    }
    await waitFor(app, ready, 20000, 'startup ready');
    await native('ready', -1);
    for (let n = 0; n < 8; n++) {
      const key = n % 4 === 2 ? info.otherKey : info.paidKey;
      await app.evaluate(`(()=>{const e=document.querySelector('#input');e.textContent=${JSON.stringify(key)};e.dispatchEvent(new Event('input',{bubbles:true}));e.focus();})()`);
      if (n % 2) await press(app, 'Enter'); else await click(app, '#b-search');
      await waitFor(modal, `!!document.querySelector('#v-cand.show')`, 10000, 'candidate DOM');
      await native('dialog', n);
      if (n < 2) {
        const shot = await modal.send('Page.captureScreenshot', { format: 'png' });
        fs.writeFileSync(path.join(evidence, 'candidate-' + n + '.png'), Buffer.from(shot.data, 'base64'));
      }
      if (n % 3 === 0) await click(modal, '#v-cand tbody tr', 2);
      else { await click(modal, '#v-cand tbody tr'); if (n % 3 === 1) await click(modal, '#v-cand .foot .btn'); else await press(modal, 'Enter'); }
      await waitFor(modal, `!document.querySelector('#v-cand.show')`, 10000, 'candidate closed');
      const judgment = `(document.querySelector('#judge')||document.querySelector('[data-judgment=paymentStatus] .ok'))?.textContent`;
      await waitFor(app, judgment + `==='済'`, 10000, 'selected result displayed');
      await native('ready', n);
      // The old token must not resize or re-open a closed native surface.
      if (n === 7) {
        const before = await modal.evaluate('window.innerWidth');
        await modal.evaluate(`window.chrome.webview.postMessage({type:'dialogSize',token:-1,width:33,height:44,title:'stale'});true`);
        await delay(100);
        if (await modal.evaluate('window.innerWidth') !== before) throw Error('Stale size message changed the closed dialog');
      }
      await click(app, '#b-clear');
      await waitFor(app, judgment + `!=='済'`, 10000, 'result cleared');
      await native('ready', n + 100);
    }
  } finally {
    fs.writeFileSync(path.join(evidence, 'results.json'), JSON.stringify(results, null, 2));
    if (modal) {
      try { await modal.evaluate(`(()=>{const b=document.querySelector('.veil.show .tb .cl');if(b)b.click();})()`); } catch (_) {}
      modal.close();
    }
    if (app) { try { await app.evaluate(`window.chrome.webview.postMessage({type:'window',command:'close'});true`); } catch (_) {} app.close(); }
  }
}
main().catch(e => { console.error(e.stack); process.exitCode = 1; });

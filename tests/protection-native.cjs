'use strict';
const fs = require('fs'), path = require('path');
const [evidence, root] = process.argv.slice(2);
const info = JSON.parse(fs.readFileSync(path.join(evidence, 'case.json'), 'utf8'));
const { CdpClient, delay, press, waitFor } = require(path.join(root, 'build/webview2_cdp.js'));
let app, modal, sequence = 0;
const results = [];
const originalConfig = fs.readFileSync(info.settings, 'utf8');
function check(name, pass) { results.push({ name, pass: !!pass }); if (!pass) throw Error(name); console.log('PASS ' + name); }
async function connect(dialog) {
  const until = Date.now() + 50000;
  while (Date.now() < until) {
    try {
      const list = await (await fetch(`http://127.0.0.1:${info.port}/json/list`, { signal: AbortSignal.timeout(2000) })).json();
      const target = list.find(x => x.type === 'page' && x.url.includes('reader-data-viewer.local') && x.url.includes('#dialog') === dialog);
      if (target) { const c = new CdpClient(target.webSocketDebuggerUrl); await c.open(); await c.send('Runtime.enable'); return c; }
    } catch (_) {}
    await delay(100);
  }
  throw Error('Reader did not start');
}
async function native(kind, label, options = {}) {
  const id = ++sequence;
  fs.writeFileSync(path.join(evidence, 'native-request.json'), JSON.stringify({ id, kind, label, ...options }));
  const until = Date.now() + 12000;
  while (Date.now() < until) {
    let reply;
    try { reply = JSON.parse(fs.readFileSync(path.join(evidence, 'native-reply.json'), 'utf8')); } catch (_) {}
    if (reply && reply.id === id) {
      results.push(reply);
      if (!reply.pass) throw Error('Native ' + label + ': ' + JSON.stringify(reply));
      check(label + ' never topmost', !reply.topmost);
      console.log('PASS native ' + label); return reply;
    }
    await delay(60);
  }
  throw Error('Native inspector timed out: ' + label);
}
async function click(client, selector, count = 1) {
  const point = await client.evaluate(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});e.scrollIntoView({block:'nearest'});const r=e.getBoundingClientRect();return {x:r.x+r.width/2,y:r.y+r.height/2};})()`);
  for (let n = 1; n <= count; n++) {
    await client.send('Input.dispatchMouseEvent', { type: 'mousePressed', ...point, button: 'left', clickCount: n });
    await client.send('Input.dispatchMouseEvent', { type: 'mouseReleased', ...point, button: 'left', clickCount: n });
  }
}
const ready = () => waitFor(app, `!!document.querySelector('#b-upd[aria-disabled=false]')`, 15000, 'ready');
const visible = id => waitFor(modal, `!!document.querySelector('#${id}.show')`, 10000, id);
const hidden = id => waitFor(modal, `!document.querySelector('#${id}.show')`, 10000, id + ' closed');
function detections() {
  const log = path.join(info.scratch, 'data/ReaderDataViewer.log');
  return fs.readFileSync(log, 'utf8').split(/\r?\n/).filter(line => /\tdetect\t/.test(line));
}
async function searched(key, hits) {
  const until = Date.now() + 12000;
  while (Date.now() < until) {
    const log = fs.readFileSync(path.join(info.scratch, 'data/ReaderDataViewer.log'), 'utf8');
    if (log.split(/\r?\n/).some(line => line.includes('key=' + key + ' source=detect') && line.includes('hits=' + hits + ' '))) return;
    await delay(80);
  }
  throw Error('Missing automatic search result for ' + key);
}
async function rejectWrite(action) {
  const before = fs.readFileSync(info.ledger);
  const changed = JSON.parse(originalConfig); changed.data.ledger.protectStates = ['FALSE'];
  fs.writeFileSync(info.settings, JSON.stringify(changed, null, 2));
  if (action === 'restore') {
    await click(app, '#b-archive'); await visible('v-archive');
    await native('dialog', 'guard-restore-selection');
    await click(modal, '#v-archive input[type=checkbox]');
    await click(modal, '#v-archive [data-modal-default=true]'); await visible('v-send');
    await click(modal, '#v-send [data-modal-default=true]');
  } else if (action === 'send') {
    await click(app, '#b-send'); await visible('v-send');
    await native('dialog', 'guard-send-confirmation');
    await click(modal, '#v-send [data-modal-default=true]');
  } else {
    const id = action === 'update' ? 'v-upd' : 'v-del';
    await click(app, action === 'update' ? '#b-upd' : '#b-del'); await visible(id);
    await native('dialog', 'guard-' + action + '-preparation');
    await click(modal, '#' + id + ' [data-modal-default=true]');
  }
  await waitFor(modal, `!!document.querySelector('#v-send.show') && document.querySelector('#v-send .body').textContent.includes('台帳作成時')`, 15000, action + ' definition error');
  await native('dialog', 'guard-' + action + '-refused', { capture: true });
  check(action + ' direct JSON mutation leaves XLSX bytes unchanged', before.equals(fs.readFileSync(info.ledger)));
  fs.writeFileSync(info.settings, originalConfig);
  await click(modal, '#v-send [data-modal-default=true]'); await hidden('v-send'); await ready();
}
async function automaticPaidOnly() {
  const countExpression = `Number(document.querySelector('#sn').textContent.match(/[0-9]+/)[0])`;
  const pending = expected => waitFor(app, `${countExpression}===${expected} && document.querySelector('#b-work').getAttribute('aria-disabled')==='false'`, 10000, 'pending ' + expected);
  const marked = () => app.evaluate(`document.querySelector('#b-work').getAttribute('aria-pressed')==='true'`);
  async function read(key, hits, label) {
    const before = detections().length;
    await native('source', label + '-source', { text: key });
    const until = Date.now() + 12000;
    while (detections().length === before && Date.now() < until) await delay(60);
    check(label + ' accepted once', detections().length === before + 1);
    await searched(key, hits);
    if (hits > 1) { await visible('v-cand'); await native('dialog', label, { capture: true }); }
    else { await ready(); await native('ready', label, { capture: true }); }
  }
  async function blank() { await native('source', 'clear-source-' + sequence, { text: '' }); await delay(400); }
  await read(info.zero, 0, 'automatic-zero-no-state');
  check('zero result adds no state', await app.evaluate(`${countExpression}===0`));
  await read(info.unpaid, 1, 'automatic-unpaid'); await pending(0);
  check('unpaid is looked up and remains unprocessed', !await marked() && await app.evaluate(`document.querySelector('.band .ok').textContent==='未決済'`));
  await click(app, '#b-work'); await pending(1);
  check('manual mark still permits unpaid', await marked());
  await click(app, '#b-work'); await visible('v-send'); await click(modal, '#v-send [data-modal-default=true]'); await hidden('v-send'); await pending(0);
  check('manual reset still permits unpaid', !await marked());
  await app.evaluate(`(()=>{const e=document.querySelector('#input');e.textContent=${JSON.stringify(info.manualPaid)};e.dispatchEvent(new Event('input',{bubbles:true}));document.querySelector('#b-search').click();})()`);
  await waitFor(app, `document.querySelector('.band .ok').textContent==='済'`, 10000, 'manual paid display');
  await ready(); await delay(200);
  check('manual paid lookup adds no automatic state', !await marked() && await app.evaluate(`${countExpression}===0`));
  await read(info.single, 1, 'automatic-paid-single'); await pending(1);
  check('confirmed paid single becomes processed', await marked());
  const repeated = detections().length;
  await native('source', 'automatic-paid-same', { text: info.single }); await delay(600);
  check('same reading neither toggles nor repeats', await marked() && detections().length === repeated && await app.evaluate(`${countExpression}===1`));
  await read(info.multiple, 2, 'automatic-mixed-unselected');
  check('no candidate is completed before confirmation', await app.evaluate(`${countExpression}===1`));
  await click(modal, '#v-cand .foot .btn:last-child'); await hidden('v-cand'); await ready();
  check('candidate cancel adds no state', await app.evaluate(`${countExpression}===1`));
  await blank(); await read(info.multiple, 2, 'automatic-mixed-unpaid-choice');
  await click(modal, '#v-cand tbody tr:nth-child(2)', 2); await hidden('v-cand'); await ready(); await delay(200);
  check('selected unpaid candidate remains unprocessed', !await marked() && await app.evaluate(`${countExpression}===1 && document.querySelector('.band .ok').textContent==='未決済'`));
  check('previous completion notice is cleared for new lookup', await app.evaluate('!document.body.textContent.includes("処理済にしました")'));
  await native('ready', 'selected-unpaid-candidate', { capture: true });
  await blank(); await read(info.multiple, 2, 'automatic-mixed-paid-choice');
  await click(modal, '#v-cand tbody tr:first-child'); await press(modal, 'Enter'); await hidden('v-cand'); await pending(2);
  check('only selected paid candidate becomes processed', await marked());
  await native('ready', 'selected-paid-candidate', { capture: true });
  await click(app, '#b-send'); await visible('v-send'); await click(modal, '#v-send [data-modal-default=true]'); await hidden('v-send'); await pending(0); await ready();
  await native('ready', 'automatic-paid-sent', { capture: true });
  check('automatic changes can be sent', await app.evaluate(`${countExpression}===0`));
}
async function main() {
  try {
    app = await connect(false); modal = await connect(true);
    for (let n = 0; n < 200; n++) {
      if (await app.evaluate(`!!document.querySelector('#b-upd[aria-disabled=false]')`)) break;
      await modal.evaluate(`(()=>{const b=document.querySelectorAll('#v-send.show .foot .btn');if(b.length===2)b[1].click();})()`);
      await delay(100);
    }
    await ready();
    await native('ready', 'startup', { capture: true });
    if (info.automaticOnly) { await automaticPaidOnly(); return; }
    await click(app, '#b-archive'); await visible('v-archive');
    await native('dialog', 'archive-full-row', { capture: true });
    check('20 archived rows listed', await modal.evaluate(`document.querySelectorAll('#v-archive tbody:first-of-type tr[data-archive-index]').length===20`));
    check('all archived values displayed', await modal.evaluate(`JSON.stringify(Array.from(document.querySelectorAll('#v-archive .archive-details td'),x=>x.textContent))===${JSON.stringify(JSON.stringify(['処理済', ...info.archivedLine.split('\t')]))}`));
    await click(modal, '#v-archive input[type=checkbox]');
    await click(modal, '#v-archive [data-modal-default=true]'); await visible('v-send');
    await native('dialog', 'restore-confirm', { capture: true });
    await click(modal, '#v-send [data-modal-default=true]'); await ready();
    await waitFor(app, `document.body.textContent.includes('81')`, 10000, 'restored row count');
    await native('ready', 'restore-complete');
    await click(app, '#b-archive'); await visible('v-archive');
    await native('dialog', 'archive-reopened', { capture: true });
    check('19 archived rows after restore', await modal.evaluate(`document.querySelectorAll('#v-archive tr[data-archive-index]').length===19`));
    await click(modal, '#v-archive .foot .btn:last-child'); await hidden('v-archive'); await ready();
    for (const action of ['update', 'delete', 'restore']) await rejectWrite(action);
    await app.evaluate(`(()=>{const e=document.querySelector('#input');e.textContent=${JSON.stringify(info.manualPending)};e.dispatchEvent(new Event('input',{bubbles:true}));document.querySelector('#b-search').click();})()`);
    await delay(300); await click(app, '#b-work');
    await waitFor(app, `document.querySelector('#sn').textContent.includes('1')`, 10000, 'one unsent change');
    await rejectWrite('send');
    check('rejected send keeps local pending change', await app.evaluate(`document.querySelector('#sn').textContent.includes('1')`));
    const presentation = JSON.parse(originalConfig);
    presentation.screen.workState.states[1].text = '保存済み（表示変更の検査）';
    fs.writeFileSync(info.settings, JSON.stringify(presentation));
    const beforeAllowedSend = fs.readFileSync(info.ledger);
    await click(app, '#b-send'); await visible('v-send'); await click(modal, '#v-send [data-modal-default=true]');
    await waitFor(app, `document.querySelector('#sn').textContent.includes('0')`, 15000, 'permitted display edit send'); await ready();
    check('display-only JSON edit permits actual send', !fs.readFileSync(info.ledger).equals(beforeAllowedSend));
    fs.writeFileSync(info.settings, originalConfig);
    if (info.guardsOnly) return;
    const before = detections().length;
    await native('source', 'ungranted-source', { text: info.single }); await searched(info.single, 1); await ready();
    const ungranted = await native('ready', 'automatic-without-permission', { capture: true });
    const last = detections().at(-1);
    check('ungranted activation result agrees with OS and search continues', last.includes('foreground=accepted') === (ungranted.foreground === ungranted.windows[0].handle));
    await native('source', 'ungranted-same', { text: info.single }); await delay(600);
    await native('away', 'ungranted-value-not-retried');
    check('ungranted read not repeatedly accepted', detections().length === before + 1);
    await native('source', 'clear-for-authorized-read', { text: '' }); await delay(350);
    await native('source', 'single-source-permission', { text: info.single, allow: true }); await delay(400); await ready();
    const normal = await native('main', 'automatic-single-permitted', { capture: true });
    check('single new read accepted once after empty', detections().length === before + 2);
    await native('source', 'same-source', { text: info.single, minimize: true }); await delay(900);
    const same = await native('away', 'same-value-does-not-raise');
    check('same value leaves Reader minimized', same.minimized && detections().length === before + 2);
    await native('source', 'invalid-source', { text: 'not a valid number' }); await delay(600);
    const invalid = await native('away', 'invalid-does-not-raise');
    check('invalid value not accepted', invalid.minimized && detections().length === before + 2);
    await native('source', 'multiple-source-permission', { text: info.multiple, allow: true }); await searched(info.multiple, 2); await visible('v-cand');
    const multiple = await native('dialog', 'automatic-multiple', { foreground: true, capture: true });
    check('minimized Reader restored', !multiple.minimized);
    await click(modal, '#v-cand tbody tr', 2); await hidden('v-cand'); await ready();
    await native('main', 'multiple-selected', { capture: true });
    await click(app, '#b-clear');
    await native('source', 'blank-source', { text: '' }); await delay(350);
    await native('source', 'reread-source-permission', { text: info.multiple, allow: true }); await visible('v-cand');
    await native('dialog', 'automatic-reread', { foreground: true, capture: true });
    await click(modal, '#v-cand tbody tr'); await press(modal, 'Enter'); await hidden('v-cand'); await ready();
    await native('main', 'reread-selected');
    await native('source', 'zero-source-permission', { text: info.zero, minimize: true, allow: true }); await searched(info.zero, 0);
    const zero = await native('main', 'automatic-zero', { capture: true });
    check('normal window size preserved after minimize', (normal.windows[0].right - normal.windows[0].left) === (zero.windows[0].right - zero.windows[0].left) && (normal.windows[0].bottom - normal.windows[0].top) === (zero.windows[0].bottom - zero.windows[0].top));
    await native('source', 'manual-source', { text: info.zero });
    const manualBefore = detections().length;
    await app.evaluate(`(()=>{const e=document.querySelector('#input');e.textContent=${JSON.stringify(info.single)};e.dispatchEvent(new Event('input',{bubbles:true}));document.querySelector('#b-search').click();})()`);
    await ready(); await delay(400);
    await native('away', 'manual-single-does-not-raise');
    check('manual search adds no detection', detections().length === manualBefore);
    await native('source', 'maximized-source', { text: '', maximize: true }); await delay(350);
    await native('source', 'maximized-read-source-permission', { text: info.single, minimize: true, allow: true });
    await delay(800);
    const maximized = await native('main', 'automatic-maximized', { capture: true });
    check('maximized placement survives minimize and automatic read', maximized.maximized && !maximized.minimized);
    const lines = detections();
    check('six new valid automatic reads', lines.length === before + 6);
    check('foreground result logged for every accepted reading', lines.every(line => line.includes('foreground=accepted') || line.includes('foreground=refused')));
    fs.writeFileSync(path.join(evidence, 'detections.json'), JSON.stringify(lines, null, 2));
  } finally {
    fs.writeFileSync(info.settings, originalConfig);
    fs.writeFileSync(path.join(evidence, 'results.json'), JSON.stringify(results, null, 2));
    if (modal) { try { await modal.evaluate(`(()=>{const b=document.querySelector('.veil.show .tb .cl');if(b)b.click();})()`); } catch (_) {} modal.close(); }
    if (app) { try { await app.evaluate(`window.chrome.webview.postMessage({type:'window',command:'close'});true`); } catch (_) {} app.close(); }
  }
}
main().catch(error => { console.error(error.stack); process.exitCode = 1; });

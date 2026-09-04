'use strict';

const delay = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds));

class CdpClient {
  constructor(url) {
    this.nextId = 1;
    this.pending = new Map();
    this.socket = new WebSocket(url);
    this.socket.addEventListener('message', (event) => {
      const message = JSON.parse(String(event.data));
      const pending = this.pending.get(message.id);
      if (!pending) { return; }
      this.pending.delete(message.id);
      if (message.error) { pending.reject(new Error(JSON.stringify(message.error))); }
      else { pending.resolve(message.result); }
    });
  }

  async open() {
    await new Promise((resolve, reject) => {
      this.socket.addEventListener('open', resolve, { once: true });
      this.socket.addEventListener('error', reject, { once: true });
    });
  }

  send(method, params) {
    const id = this.nextId++;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.socket.send(JSON.stringify({ id, method, params: params || {} }));
    });
  }

  async evaluate(expression) {
    const response = await this.send('Runtime.evaluate', {
      expression,
      returnByValue: true,
      awaitPromise: true
    });
    if (response.exceptionDetails) {
      const detail = response.exceptionDetails.exception;
      throw new Error(detail && detail.description ? detail.description : response.exceptionDetails.text);
    }
    return response.result.value;
  }

  close() {
    try { this.socket.close(); } catch (_) { }
  }
}

async function targetFor(port, timeoutMs) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(`http://127.0.0.1:${port}/json/list`);
      const entries = await response.json();
      const page = entries.find((entry) => entry.type === 'page');
      if (page) { return page; }
    } catch (_) { }
    await delay(20);
  }
  throw new Error(`WebView2 DevTools target did not start on port ${port}`);
}

async function connect(port, timeoutMs) {
  const target = await targetFor(port, timeoutMs || 60000);
  const client = new CdpClient(target.webSocketDebuggerUrl);
  await client.open();
  await client.send('Runtime.enable');
  return client;
}

async function waitFor(client, expression, timeoutMs, description) {
  const started = Date.now();
  let lastError = null;
  while (Date.now() - started < timeoutMs) {
    try {
      const value = await client.evaluate(expression);
      if (value) { return value; }
    } catch (error) {
      lastError = error;
    }
    await delay(20);
  }
  throw new Error(`Timed out waiting for ${description || expression}` +
    (lastError ? `; last error: ${lastError.message}` : ''));
}

const harnessSource = `(() => {
  function install() {
    if (window.__rdvTestHookInstalled) { return true; }
    if (!(window.chrome && window.chrome.webview &&
          typeof window.chrome.webview.postMessage === 'function')) { return false; }
    const webview = window.chrome.webview;
    const nativePost = webview.postMessage.bind(webview);
    window.__rdvTestIncoming = [];
    window.__rdvTestOutgoing = [];
    window.__rdvTestNativePost = nativePost;
    window.__rdvTestHeld = null;
    webview.addEventListener('message', (event) => {
      window.__rdvTestIncoming.push(event.data);
    });
    webview.postMessage = (message) => {
      window.__rdvTestOutgoing.push(message);
      if (message && message.type === 'modalShown') {
        window.__rdvTestHeld = message;
        return;
      }
      return nativePost(message);
    };
    window.__rdvTestHookInstalled = true;
    return true;
  }
  if (!install() && !window.__rdvTestHookTimer) {
    window.__rdvTestHookTimer = window.setInterval(install, 10);
  }
})()`;

async function installHarness(client, timeoutMs) {
  await client.send('Page.enable');
  await client.send('Page.addScriptToEvaluateOnNewDocument', { source: harnessSource });
  try { await client.evaluate(harnessSource); } catch (_) { }
  await waitFor(client,
    `location.hostname === 'reader-data-viewer.local' && window.__rdvTestHookInstalled === true`,
    timeoutMs || 60000,
    'the Reader DOM test hook');
}

function keyDefinition(key) {
  const definitions = {
    Enter: ['Enter', 13],
    Escape: ['Escape', 27],
    Tab: ['Tab', 9],
    ArrowUp: ['ArrowUp', 38],
    ArrowDown: ['ArrowDown', 40],
    ArrowLeft: ['ArrowLeft', 37],
    ArrowRight: ['ArrowRight', 39],
    Home: ['Home', 36],
    End: ['End', 35],
    PageUp: ['PageUp', 33],
    PageDown: ['PageDown', 34],
    Space: ['Space', 32]
  };
  return definitions[key] || [key, key.length === 1 ? key.toUpperCase().charCodeAt(0) : 0];
}

async function press(client, key, modifiers) {
  const definition = keyDefinition(key);
  const params = {
    key: key === 'Space' ? ' ' : key,
    code: definition[0],
    windowsVirtualKeyCode: definition[1],
    nativeVirtualKeyCode: definition[1],
    modifiers: Number(modifiers) || 0
  };
  await client.send('Input.dispatchKeyEvent', Object.assign({ type: 'rawKeyDown' }, params));
  await client.send('Input.dispatchKeyEvent', Object.assign({ type: 'keyUp' }, params));
}

function decodeExpression(value) {
  return Buffer.from(value || '', 'base64').toString('utf8');
}

async function main() {
  const port = Number(process.argv[2]);
  const command = process.argv[3] || '';
  if (!Number.isInteger(port) || port < 1) { throw new Error('A DevTools port is required.'); }
  const client = await connect(port, 60000);
  try {
    if (command === 'install') {
      await installHarness(client, Number(process.argv[4]) || 60000);
      return { installed: true, url: await client.evaluate('location.href') };
    }
    if (command === 'eval') {
      return await client.evaluate(decodeExpression(process.argv[4]));
    }
    if (command === 'wait') {
      const expression = decodeExpression(process.argv[4]);
      return await waitFor(client, expression, Number(process.argv[5]) || 30000, expression);
    }
    if (command === 'press') {
      await press(client, process.argv[4], Number(process.argv[5]) || 0);
      return true;
    }
    throw new Error(`Unknown command: ${command}`);
  } finally {
    client.close();
  }
}

module.exports = {
  CdpClient,
  connect,
  delay,
  harnessSource,
  installHarness,
  press,
  waitFor
};

if (require.main === module) {
  main().then((value) => {
    process.stdout.write(JSON.stringify(value === undefined ? null : value));
  }).catch((error) => {
    process.stderr.write((error.stack || String(error)) + '\n');
    process.exitCode = 1;
  });
}

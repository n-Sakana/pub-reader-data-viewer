'use strict';

const {
  connect,
  installHarness,
  waitFor
} = require('./webview2_cdp');

const port = Number(process.argv[2]);
const singleKey = process.argv[3];
const scenario = process.argv[4];
const expectedByScenario = {
  'emphasis-max': {
    width: 1000, height: 760, gap: '8px', padding: '8px',
    bodyFont: '10pt', keyFont: '36pt', judgmentFont: '36pt', unsearchedFont: '36pt'
  },
  'gap-zero': {
    width: 780, height: 671, gap: '0px', padding: '0px',
    bodyFont: '10pt', keyFont: '15pt', judgmentFont: '15pt', unsearchedFont: '12pt'
  }
};

let passed = 0;
let failed = 0;

function detailOf(value) {
  if (typeof value === 'string') { return value; }
  try { return JSON.stringify(value); } catch (_) { return String(value); }
}

function check(name, condition, detail) {
  if (condition) {
    passed++;
    process.stdout.write(`  ok   ${name}\n`);
  } else {
    failed++;
    process.stdout.write(`  FAIL ${name} (${detailOf(detail)})\n`);
  }
}

function fits(widths) {
  return widths.scrollWidth <= widths.clientWidth + 1;
}

async function main() {
  const expected = expectedByScenario[scenario];
  if (!Number.isInteger(port) || !singleKey || singleKey.length !== 8 || !expected) {
    throw new Error('port, eight-character key, and known layout scenario are required');
  }

  const client = await connect(port, 60000);
  try {
    await installHarness(client, 60000);
    await waitFor(client,
      `document.querySelector('.stage.runtime #b-search[aria-disabled=false]') !== null`,
      60000,
      `the ready ${scenario} Reader DOM`);

    const layout = await client.evaluate(`(() => {
      const stage = document.querySelector('.stage');
      const win = stage.querySelector('.win');
      const client = stage.querySelector('.client');
      const stack = stage.querySelector('.stack');
      const widths = (node) => ({ scrollWidth: node.scrollWidth, clientWidth: node.clientWidth });
      const stageRect = stage.getBoundingClientRect();
      const style = getComputedStyle(stage);
      const clientStyle = getComputedStyle(client);
      const stackRect = stack.getBoundingClientRect();
      return {
        viewport: { width: innerWidth, height: innerHeight },
        stage: { width: stageRect.width, height: stageRect.height },
        variables: {
          width: style.getPropertyValue('--card-width').trim(),
          gap: style.getPropertyValue('--card-gap').trim(),
          stackGap: getComputedStyle(stack).gap,
          padding: clientStyle.padding,
          bodyFont: style.getPropertyValue('--body-font').trim(),
          keyFont: style.getPropertyValue('--key-font').trim(),
          judgmentFont: style.getPropertyValue('--judgment-font').trim(),
          unsearchedFont: style.getPropertyValue('--unsearched-font').trim()
        },
        document: widths(document.documentElement),
        stageWidth: widths(stage),
        windowWidth: widths(win),
        clientWidth: widths(client),
        stackWidth: widths(stack),
        sectionsInside: Array.from(stack.children).every((node) => {
          const rect = node.getBoundingClientRect();
          return rect.left >= stackRect.left - 0.75 && rect.right <= stackRect.right + 0.75;
        })
      };
    })()`);

    check(`${scenario} settings reach the live CSS variables`,
      layout.variables.width === `${expected.width}px` &&
        layout.variables.gap === expected.gap && layout.variables.stackGap === expected.gap &&
        layout.variables.padding === expected.padding &&
        layout.variables.bodyFont === expected.bodyFont &&
        layout.variables.keyFont === expected.keyFont &&
        layout.variables.judgmentFont === expected.judgmentFont &&
        layout.variables.unsearchedFont === expected.unsearchedFont &&
        Math.abs(layout.stage.width - expected.width) < 0.75 &&
        Math.abs(layout.stage.height - expected.height) < 0.75,
      { expected, actual: layout });
    check(`${scenario} document has no horizontal overflow`, fits(layout.document), layout.document);
    check(`${scenario} stage and window have no horizontal overflow`,
      fits(layout.stageWidth) && fits(layout.windowWidth),
      { stage: layout.stageWidth, window: layout.windowWidth });
    check(`${scenario} scrollable client has no horizontal overflow`,
      fits(layout.clientWidth),
      { client: layout.clientWidth, stack: layout.stackWidth });
    check(`${scenario} sections stay inside both stack edges`,
      layout.sectionsInside, layout);

    await client.evaluate(`(() => {
      const input = document.querySelector('#input');
      input.textContent = ${JSON.stringify(singleKey)};
      input.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText' }));
      input.focus();
      document.querySelector('#b-search').click();
      return true;
    })()`);
    await waitFor(client,
      `document.querySelector('.keyvalue').textContent.trim() === ${JSON.stringify(singleKey)}`,
      30000,
      `the complete ${scenario} key value`);

    const key = await client.evaluate(`(() => {
      const node = document.querySelector('.keyvalue');
      const rect = node.getBoundingClientRect();
      const style = getComputedStyle(node);
      const range = document.createRange();
      range.selectNodeContents(node);
      const text = range.getBoundingClientRect();
      return {
        value: node.textContent.trim(),
        characters: Array.from(node.textContent.trim()).length,
        scrollWidth: node.scrollWidth,
        clientWidth: node.clientWidth,
        rect: { left: rect.left, right: rect.right, width: rect.width },
        textRect: { left: text.left, right: text.right, width: text.width },
        padding: { left: parseFloat(style.paddingLeft), right: parseFloat(style.paddingRight) },
        fontSize: style.fontSize,
        overflow: style.overflow
      };
    })()`);
    check(`${scenario} renders the complete eight-character key`,
      key.value === singleKey && key.characters === 8, key);
    check(`${scenario} key value has no scrollWidth clipping`,
      key.scrollWidth <= key.clientWidth + 1, key);
    check(`${scenario} rendered key text stays inside the field content box`,
      key.textRect.left >= key.rect.left + key.padding.left - 1 &&
        key.textRect.right <= key.rect.right - key.padding.right + 1 &&
        key.textRect.width <= key.clientWidth - key.padding.left - key.padding.right + 1,
      key);

    process.stdout.write(`\n${passed} passed, ${failed} failed (${scenario})\n`);
    process.stdout.write(failed ? 'RESULT: FAIL\n' : 'RESULT: PASS\n');
    if (passed + failed !== 8 || failed) { process.exitCode = 1; }
  } finally {
    try {
      await client.evaluate(`window.__rdvTestNativePost({ type: 'window', command: 'close' }); true`);
    } catch (_) { }
    client.close();
  }
}

main().catch((error) => {
  process.stderr.write((error.stack || String(error)) + '\n');
  process.exitCode = 1;
});

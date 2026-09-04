'use strict';

const {
  connect,
  installHarness,
  press,
  waitFor
} = require('./webview2_cdp');

const port = Number(process.argv[2]);
const singleKey = process.argv[3];
const singleIdentity = process.argv[4];
const multiKey = process.argv[5];

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

async function main() {
  if (!Number.isInteger(port) || !singleKey || !singleIdentity || !multiKey) {
    throw new Error('port, single key, single identity, and multi key are required');
  }

  const client = await connect(port, 60000);
  try {
    await installHarness(client, 60000);
    await waitFor(client,
      `document.querySelector('.stage.runtime #b-search[aria-disabled=false]') !== null`,
      60000,
      'the ready Reader DOM');

    async function evaluate(expression) {
      return client.evaluate(expression);
    }

    async function openModal(button, veil) {
      await evaluate(`(() => {
        window.__rdvTestHeld = null;
        const button = document.querySelector(${JSON.stringify(button)});
        button.focus();
        button.click();
        return true;
      })()`);
      await waitFor(client,
        `document.querySelector(${JSON.stringify(veil)}).classList.contains('show')`,
        30000,
        `${veil} to open`);
    }

    async function closeModal(veil) {
      await evaluate(`document.querySelector(${JSON.stringify(veil + ' .foot .btn:last-child')}).click(); true`);
      await waitFor(client,
        `!document.querySelector(${JSON.stringify(veil)}).classList.contains('show')`,
        30000,
        `${veil} to close`);
    }

    async function modalFit(veil) {
      return evaluate(`(() => {
        const stage = document.querySelector('.stage').getBoundingClientRect();
        const dialog = document.querySelector(${JSON.stringify(veil + ' .dlg')});
        const body = dialog.querySelector('.body');
        const rect = dialog.getBoundingClientRect();
        return {
          inside: rect.left >= stage.left - 0.75 && rect.right <= stage.right + 0.75 &&
            rect.top >= stage.top - 0.75 && rect.bottom <= stage.bottom + 0.75,
          bodyHorizontal: body.scrollWidth <= body.clientWidth + 1,
          rect: [rect.left, rect.top, rect.right, rect.bottom],
          bodyWidth: [body.scrollWidth, body.clientWidth]
        };
      })()`);
    }

    process.stdout.write('\nmain window DOM\n');
    const mainDom = await evaluate(`(() => {
      const stage = document.querySelector('.stage');
      const win = stage.querySelector('.win');
      const client = stage.querySelector('.client');
      const stack = stage.querySelector('.stack');
      const rect = (node) => {
        const value = node.getBoundingClientRect();
        return { left: value.left, top: value.top, right: value.right,
          bottom: value.bottom, width: value.width, height: value.height };
      };
      const children = Array.from(stack.children);
      const childRects = children.map(rect);
      const gaps = childRects.slice(1).map((value, index) => value.top - childRects[index].bottom);
      const stackRect = rect(stack);
      const actions = Array.from(stage.querySelectorAll('[data-action]'));
      const status = Array.from(stage.querySelectorAll('.sb .sp'));
      return {
        href: location.href,
        host: location.hostname,
        protocol: location.protocol,
        runtime: stage.classList.contains('runtime'),
        scaffolding: document.querySelectorAll('.demo,.notes').length,
        productWindows: document.querySelectorAll('.stage > .win').length,
        windowChildren: Array.from(win.children).map((node) => node.className),
        stackOrder: children.map((node) => node.className),
        groups: stack.querySelectorAll('fieldset.gb').length,
        stageRect: rect(stage),
        winRect: rect(win),
        clientRect: rect(client),
        stackRect,
        gaps,
        childrenInside: childRects.every((value) =>
          value.left >= stackRect.left - 0.75 && value.right <= stackRect.right + 0.75),
        style: {
          width: stage.style.getPropertyValue('--card-width').trim(),
          maxWidth: getComputedStyle(stage).maxWidth,
          gap: stage.style.getPropertyValue('--card-gap').trim(),
          columnGap: getComputedStyle(stage.querySelector('.cols')).gap,
          padding: getComputedStyle(client).padding,
          font: stage.style.getPropertyValue('--ff'),
          body: stage.style.getPropertyValue('--body-font').trim(),
          key: stage.style.getPropertyValue('--key-font').trim(),
          judgment: stage.style.getPropertyValue('--judgment-font').trim(),
          unsearched: stage.style.getPropertyValue('--unsearched-font').trim()
        },
        overflow: {
          document: [document.documentElement.scrollWidth, document.documentElement.clientWidth,
            document.documentElement.scrollHeight, document.documentElement.clientHeight],
          stage: [stage.scrollWidth, stage.clientWidth, stage.scrollHeight, stage.clientHeight],
          client: [client.scrollWidth, client.clientWidth, client.scrollHeight, client.clientHeight],
          win: [win.scrollWidth, win.clientWidth, win.scrollHeight, win.clientHeight]
        },
        actionIds: actions.map((node) => node.id),
        actionRoles: actions.map((node) => node.getAttribute('role')),
        actionStates: actions.map((node) => [node.getAttribute('aria-disabled'), node.tabIndex]),
        input: (() => {
          const node = document.querySelector('#input');
          return { role: node.getAttribute('role'), editable: node.contentEditable,
            disabled: node.getAttribute('aria-disabled'), max: node.getAttribute('data-max-length'),
            label: node.getAttribute('aria-label') };
        })(),
        focus: document.activeElement.id,
        work: (() => {
          const node = document.querySelector('#b-work');
          return { pressed: node.getAttribute('aria-pressed'), text: node.textContent };
        })(),
        status: { role: stage.querySelector('.sb').getAttribute('role'),
          live: stage.querySelector('.sb').getAttribute('aria-live'),
          count: status.length, flex: status.map((node) => getComputedStyle(node).flex) },
        bandLabel: stage.querySelector('.band-label').textContent
      };
    })()`);

    check('the target is the trusted WebView2 page',
      mainDom.host === 'reader-data-viewer.local' && mainDom.protocol === 'https:', mainDom.href);
    check('runtime DOM replaces the mock scaffolding',
      mainDom.runtime && mainDom.scaffolding === 0, mainDom);
    check('the stage contains one product window', mainDom.productWindows === 1, mainDom.productWindows);
    check('the window owns title, client, and status in order',
      mainDom.windowChildren.join('|') === 'tb|client|sb', mainDom.windowChildren);
    check('the stack follows the configured screen order',
      mainDom.stackOrder.join('|') === 'gb|cols|gb|gb|band|commandbar|sep|sendbar', mainDom.stackOrder);
    check('the main record surface has five framed groups', mainDom.groups === 5, mainDom.groups);
    check('card width is applied from settings',
      mainDom.style.width === '780px' && mainDom.style.maxWidth === '780px' &&
        Math.abs(mainDom.stageRect.width - 780) < 0.75, mainDom.style);
    check('one configured gap drives sections and columns',
      mainDom.style.gap === '8px' && mainDom.style.columnGap === '8px' &&
        mainDom.gaps.every((value) => Math.abs(value - 8) < 0.75),
      { style: mainDom.style, gaps: mainDom.gaps });
    check('card padding is applied from settings', mainDom.style.padding === '8px', mainDom.style.padding);
    check('font family is applied from settings', mainDom.style.font.includes('Meiryo UI'), mainDom.style.font);
    check('all four font-size settings reach CSS variables',
      [mainDom.style.body, mainDom.style.key, mainDom.style.judgment,
        mainDom.style.unsearched].join('|') === '10pt|15pt|15pt|12pt', mainDom.style);
    check('the configured start viewport is the live DOM viewport',
      Math.abs(mainDom.stageRect.width - 780) < 0.75 &&
        Math.abs(mainDom.stageRect.height - 671) < 0.75, mainDom.stageRect);
    check('the page has no horizontal or vertical scrollbar',
      mainDom.overflow.document[0] <= mainDom.overflow.document[1] &&
        mainDom.overflow.document[2] <= mainDom.overflow.document[3] + 1, mainDom.overflow.document);
    check('the main client fits without scrolling',
      mainDom.overflow.client[0] <= mainDom.overflow.client[1] &&
        mainDom.overflow.client[2] <= mainDom.overflow.client[3] + 1, mainDom.overflow.client);
    check('the window and stage do not overflow their boxes',
      mainDom.overflow.stage[0] <= mainDom.overflow.stage[1] &&
        mainDom.overflow.stage[2] <= mainDom.overflow.stage[3] + 1 &&
        mainDom.overflow.win[0] <= mainDom.overflow.win[1] &&
        mainDom.overflow.win[2] <= mainDom.overflow.win[3] + 1, mainDom.overflow);
    check('stack children stay inside both horizontal edges', mainDom.childrenInside,
      { stack: mainDom.stackRect, order: mainDom.stackOrder });
    check('all eight configured actions exist once',
      mainDom.actionIds.join('|') ===
        'b-search|b-clear|b-work|b-out|b-upd|b-del|b-set|b-send', mainDom.actionIds);
    check('configured actions expose button semantics',
      mainDom.actionRoles.every((value) => value === 'button'), mainDom.actionRoles);
    check('ready actions are enabled and keyboard reachable',
      mainDom.actionStates.every((value) => value[0] === 'false' && value[1] === 0), mainDom.actionStates);
    check('the search field exposes an enabled textbox',
      mainDom.input.role === 'textbox' && mainDom.input.editable === 'true' &&
        mainDom.input.disabled === 'false' && mainDom.input.max === '64' &&
        mainDom.input.label === '番号1(N):', mainDom.input);
    check('initial keyboard focus is in the search field', mainDom.focus === 'input', mainDom.focus);
    check('work state exposes a toggle state and configured text',
      mainDom.work.pressed === 'false' && mainDom.work.text === '未処理(W)', mainDom.work);
    check('status is a four-part polite live region',
      mainDom.status.role === 'status' && mainDom.status.live === 'polite' &&
        mainDom.status.count === 4 && mainDom.status.flex[2].startsWith('1 1'), mainDom.status);
    check('the configured judgment-band label is present',
      mainDom.bandLabel === 'ステータス１', mainDom.bandLabel);

    process.stdout.write('\nstate rendering and keyboard modal\n');
    await evaluate(`(() => {
      const input = document.querySelector('#input');
      input.textContent = ${JSON.stringify(singleKey)};
      input.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText' }));
      input.focus();
      document.querySelector('#b-search').click();
      return true;
    })()`);
    await waitFor(client,
      `Array.from(document.querySelectorAll('.fld')).some((node) => node.textContent.trim() === ${JSON.stringify(singleIdentity)})`,
      30000,
      'the single search result');
    const single = await evaluate(`(() => {
      const muted = Array.from(document.querySelectorAll('.fld.tone-muted'));
      return { identity: Array.from(document.querySelectorAll('.fld')).some((node) =>
          node.textContent.trim() === ${JSON.stringify(singleIdentity)}),
        muted: muted.map((node) => node.textContent.trim()),
        errors: document.querySelectorAll('.fld.tone-error').length,
        judgment: document.querySelector('#judge').textContent,
        workDisabled: document.querySelector('#b-work').getAttribute('aria-disabled') };
    })()`);
    check('a real search renders the oracle identity', single.identity, single);
    check('muted fields retain their visible placeholder text',
      single.muted.length > 0 && single.muted.every((value) => value.length > 0), single.muted);
    check('a valid record has no error field and keeps work enabled',
      single.errors === 0 && single.workDisabled === 'false', single);

    await evaluate(`(() => {
      const input = document.querySelector('#input');
      input.textContent = 'x';
      input.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText' }));
      input.focus();
      return true;
    })()`);
    await press(client, 'Enter', 0);
    await waitFor(client,
      `document.querySelector('#v-send').classList.contains('show')`,
      30000,
      'the invalid-key warning');
    const warning = await evaluate(`(() => {
      const veil = document.querySelector('#v-send');
      const dialog = veil.querySelector('.dlg');
      return { role: dialog.getAttribute('role'), modal: dialog.getAttribute('aria-modal'),
        buttons: veil.querySelectorAll('.foot .btn').length,
        body: veil.querySelector('.modal-message').innerText,
        focus: document.activeElement.textContent };
    })()`);
    check('Enter on the search field opens the real warning dialog',
      warning.body.includes('^[0-9]{8}$'), warning);
    check('the warning exposes modal dialog semantics',
      warning.role === 'dialog' && warning.modal === 'true' && warning.buttons === 1, warning);
    check('the warning default button receives focus', warning.focus === 'OK', warning.focus);
    await press(client, 'Tab', 0);
    const tabForward = await evaluate(`document.activeElement.classList.contains('cl')`);
    await press(client, 'Tab', 8);
    const tabBack = await evaluate(`document.activeElement.hasAttribute('data-modal-default')`);
    check('Tab and Shift+Tab remain inside the modal', tabForward && tabBack,
      { tabForward, tabBack });
    await press(client, 'Escape', 0);
    await waitFor(client,
      `!document.querySelector('#v-send').classList.contains('show')`,
      30000,
      'the warning to close');
    check('Escape closes the modal and restores search focus',
      await evaluate(`document.activeElement.id === 'input'`),
      await evaluate(`document.activeElement.id`));

    process.stdout.write('\nsettings DOM\n');
    await openModal('#b-set', '#v-set');
    const settingsFit = await modalFit('#v-set');
    const settings = await evaluate(`(() => {
      const veil = document.querySelector('#v-set');
      const dialog = veil.querySelector('.dlg');
      const spin = veil.querySelector('[data-field=candidateRows]');
      return { label: dialog.getAttribute('aria-label'), role: dialog.getAttribute('role'),
        modal: dialog.getAttribute('aria-modal'),
        legends: Array.from(veil.querySelectorAll('.body > fieldset > legend')).map((node) => node.textContent),
        edits: Array.from(veil.querySelectorAll('[role=textbox][contenteditable=true]')).map((node) => node.dataset.field),
        readOnly: [veil.querySelector('[data-target-summary]'), veil.querySelector('[data-target-read]')]
          .map((node) => [node.contentEditable, node.getAttribute('role')]),
        nativeInputs: veil.querySelectorAll('input[type=number],input[type=date]').length,
        spin: { role: spin.getAttribute('role'), min: spin.getAttribute('aria-valuemin'),
          max: spin.getAttribute('aria-valuemax'), now: spin.getAttribute('aria-valuenow') },
        steps: Array.from(veil.querySelectorAll('.number-step')).map((node) => node.getAttribute('aria-label')),
        buttons: Array.from(veil.querySelectorAll('.body .btn')).map((node) => node.textContent) };
    })()`);
    check('settings opens as the labelled modal dialog',
      settings.label === '設定' && settings.role === 'dialog' && settings.modal === 'true', settings);
    check('settings dialog stays inside the stage without horizontal clipping',
      settingsFit.inside && settingsFit.bodyHorizontal, settingsFit);
    check('settings has the three configured groups',
      settings.legends.join('|') === '場所|検索|監視対象', settings.legends);
    check('settings paths and pattern are editable textboxes',
      settings.edits.join('|') === 'dataDir|ledger|log|pattern', settings.edits);
    check('watch target summaries remain read-only DOM fields',
      settings.readOnly.every((value) => value[0] !== 'true' && value[1] === null), settings.readOnly);
    check('candidate count is a custom accessible spinbutton',
      settings.nativeInputs === 0 && settings.spin.role === 'spinbutton' &&
        settings.spin.min === '1' && settings.spin.max === '1000' && settings.spin.now === '100' &&
        settings.steps.join('|') === '1 増やす|1 減らす', settings);
    await evaluate(`document.querySelector('#v-set [data-field=candidateRows]').focus(); true`);
    await press(client, 'ArrowUp', 0);
    const stepped = await evaluate(`(() => {
      const node = document.querySelector('#v-set [data-field=candidateRows]');
      return [node.textContent, node.getAttribute('aria-valuenow')];
    })()`);
    check('ArrowUp changes both spinbutton value and ARIA value',
      stepped[0] === '101' && stepped[1] === '101', stepped);
    check('settings exposes browse, picker, OK, and cancel actions',
      settings.buttons.join('|') === '参照...|参照...|参照...|画面から選ぶ|OK|キャンセル', settings.buttons);
    await evaluate(`(() => {
      document.querySelector('#v-set [data-field=pattern]').textContent = '[';
      document.querySelector('#v-set .foot [data-modal-default=true]').click();
      return true;
    })()`);
    await waitFor(client,
      `!document.querySelector('#v-set .setting-error').hidden`,
      30000,
      'the .NET regex validation');
    const regexError = await evaluate(`(() => {
      const veil = document.querySelector('#v-set');
      const alert = veil.querySelector('.setting-error');
      return { shown: veil.classList.contains('show'), role: alert.getAttribute('role'),
        text: alert.textContent, focus: document.activeElement.dataset.field };
    })()`);
    check('invalid .NET regex stays open with a focused alert',
      regexError.shown && regexError.role === 'alert' && regexError.text.length > 0 &&
        regexError.focus === 'pattern', regexError);
    await press(client, 'Escape', 0);
    await waitFor(client, `!document.querySelector('#v-set').classList.contains('show')`,
      30000, 'settings to cancel');
    check('settings cancel restores focus to its launch button',
      await evaluate(`document.activeElement.id === 'b-set'`),
      await evaluate(`document.activeElement.id`));

    process.stdout.write('\nexport DOM\n');
    await openModal('#b-out', '#v-out');
    const exportFit = await modalFit('#v-out');
    const exportStart = await evaluate(`(() => {
      const veil = document.querySelector('#v-out');
      const lists = veil.querySelectorAll('[role=listbox]');
      return { label: veil.querySelector('.dlg').getAttribute('aria-label'),
        fit: [veil.querySelector('.dlg').getAttribute('role'), veil.querySelector('.dlg').getAttribute('aria-modal')],
        listLabels: Array.from(lists).map((node) => node.getAttribute('aria-label')),
        counts: Array.from(lists).map((node) => node.querySelectorAll('[role=option]').length),
        selected: Array.from(lists[1].querySelectorAll('[role=option]')).map((node) => node.dataset.ref),
        destination: (() => { const node = veil.querySelector('[data-field=exportPath]');
          return [node.getAttribute('role'), node.contentEditable, node.textContent]; })() };
    })()`);
    check('export opens as a fitted labelled modal',
      exportStart.label === 'テーブル出力' && exportStart.fit.join('|') === 'dialog|true' &&
        exportFit.inside && exportFit.bodyHorizontal, { exportStart, exportFit });
    check('export has two labelled listboxes with 16 and 5 options',
      exportStart.listLabels.join('|') === '出力できる項目|出力する項目' &&
        exportStart.counts.join('|') === '16|5', exportStart);
    check('export starts with the five configured fields in order',
      exportStart.selected.join('|') === 'B.key1|B.key2|A.a_name|B.b_status|$work', exportStart.selected);

    await evaluate(`(() => {
      const item = document.querySelector('#v-out [aria-label="出力できる項目"] [role=option]');
      item.click();
      item.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
      return true;
    })()`);
    let listCounts = await evaluate(`Array.from(document.querySelectorAll('#v-out [role=listbox]')).map((node) => node.children.length)`);
    check('double-click moves an available field to selected', listCounts.join('|') === '15|6', listCounts);
    await evaluate(`(() => {
      const items = document.querySelectorAll('#v-out [aria-label="出力する項目"] [role=option]');
      const item = items[items.length - 1];
      item.click();
      item.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
      return true;
    })()`);
    listCounts = await evaluate(`Array.from(document.querySelectorAll('#v-out [role=listbox]')).map((node) => node.children.length)`);
    check('double-click moves a selected field back', listCounts.join('|') === '16|5', listCounts);
    await evaluate(`(() => {
      const list = document.querySelector('#v-out [aria-label="出力できる項目"]');
      const item = list.lastElementChild;
      item.click(); item.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
      Array.from(document.querySelectorAll('#v-out .btn')).find((node) =>
        node.textContent === '既定に戻す').click();
      return true;
    })()`);
    const resetFields = await evaluate(`Array.from(document.querySelectorAll('#v-out [aria-label="出力する項目"] [role=option]')).map((node) => node.dataset.ref)`);
    check('reset restores the five configured fields in order',
      resetFields.join('|') === 'B.key1|B.key2|A.a_name|B.b_status|$work', resetFields);

    await evaluate(`(() => {
      const value = document.querySelector('#v-out [data-field=filterFirst]');
      value.textContent = 'SAMPLE';
      Array.from(document.querySelectorAll('#v-out .fgrid > .btn')).find((node) => node.textContent === '追加').click();
      return true;
    })()`);
    await waitFor(client, `document.querySelectorAll('#v-out .f3 tbody tr').length === 1`,
      30000, 'the text export filter');
    const textFilter = await evaluate(`Array.from(document.querySelectorAll('#v-out .f3 tbody tr td')).map((node) => node.textContent)`);
    check('a text filter is validated into the visible condition table',
      textFilter[1] === '含む' && textFilter[2] === 'SAMPLE', textFilter);
    await evaluate(`Array.from(document.querySelectorAll('#v-out .btn')).find((node) => node.textContent === '削除').click(); true`);
    await waitFor(client, `document.querySelectorAll('#v-out .f3 tbody tr').length === 0`,
      30000, 'the text filter removal');
    check('the selected text filter can be removed',
      await evaluate(`document.querySelectorAll('#v-out .f3 tbody tr').length === 0`), 'row remains');

    await evaluate(`(() => {
      const select = document.querySelector('#v-out select');
      select.value = 'B.b_date';
      select.dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    })()`);
    const dateEditor = await evaluate(`(() => {
      const veil = document.querySelector('#v-out');
      return { operators: Array.from(veil.querySelectorAll('select'))[1].options.length,
        operator: Array.from(veil.querySelectorAll('select'))[1].value,
        dates: veil.querySelectorAll('.date-value').length,
        native: veil.querySelectorAll('input[type=date],input[type=number]').length };
    })()`);
    check('a date filter uses range and two custom date editors',
      dateEditor.operators === 1 && dateEditor.operator === 'range' &&
        dateEditor.dates === 2 && dateEditor.native === 0, dateEditor);
    const firstDate = await evaluate(`document.querySelector('#v-out .date-value').textContent`);
    await evaluate(`document.querySelector('#v-out .date-value').focus(); true`);
    await press(client, 'Enter', 0);
    await waitFor(client, `document.querySelector('#v-out .date-editor').classList.contains('open')`,
      30000, 'the calendar');
    const calendar = await evaluate(`(() => {
      const popup = document.querySelector('#v-out .calendar-popup');
      return { role: popup.getAttribute('role'), label: popup.getAttribute('aria-label'),
        days: popup.querySelectorAll('.calendar-day').length,
        selected: popup.querySelectorAll('.calendar-day[aria-pressed=true]').length,
        focus: document.activeElement.className };
    })()`);
    check('the custom calendar is a 42-day dialog with one selected day',
      calendar.role === 'dialog' && calendar.label === 'カレンダー' &&
        calendar.days === 42 && calendar.selected === 1 &&
        calendar.focus.includes('selected'), calendar);
    await press(client, 'ArrowRight', 0);
    await new Promise((resolve) => setTimeout(resolve, 50));
    const movedDate = await evaluate(`(() => ({ value: document.querySelector('#v-out .date-value').textContent,
      selected: document.querySelectorAll('#v-out .calendar-day[aria-pressed=true]').length,
      focus: document.activeElement.classList.contains('selected') }))()`);
    check('calendar arrow keys move the selected date and focus',
      movedDate.value !== firstDate && movedDate.selected === 1 && movedDate.focus, movedDate);
    await press(client, 'Escape', 0);
    await waitFor(client, `!document.querySelector('#v-out .date-editor').classList.contains('open')`,
      30000, 'the calendar to close');
    check('calendar Escape closes only the calendar and restores its button',
      await evaluate(`document.querySelector('#v-out').classList.contains('show') &&
        document.activeElement.classList.contains('date-value')`),
      await evaluate(`document.activeElement.className`));

    await evaluate(`(() => {
      const select = document.querySelector('#v-out select');
      select.value = 'B.b_qty';
      select.dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    })()`);
    const numberEditor = await evaluate(`(() => {
      const veil = document.querySelector('#v-out');
      return { operator: Array.from(veil.querySelectorAll('select'))[1].value,
        values: veil.querySelectorAll('.filter-value-host [role=textbox]').length,
        dates: veil.querySelectorAll('.date-editor').length,
        native: veil.querySelectorAll('input[type=date],input[type=number]').length };
    })()`);
    check('a number filter uses range and two custom numeric textboxes',
      numberEditor.operator === 'range' && numberEditor.values === 2 &&
        numberEditor.dates === 0 && numberEditor.native === 0, numberEditor);
    await evaluate(`(() => {
      const values = document.querySelectorAll('#v-out .filter-value-host [data-field]');
      values[0].textContent = '10'; values[1].textContent = '2';
      Array.from(document.querySelectorAll('#v-out .fgrid > .btn')).find((node) => node.textContent === '追加').click();
      return true;
    })()`);
    await waitFor(client, `!document.querySelector('#v-out .setting-error').hidden`,
      30000, 'the reversed number-range validation');
    const reversed = await evaluate(`(() => {
      const alert = document.querySelector('#v-out .setting-error');
      return { role: alert.getAttribute('role'), text: alert.textContent,
        rows: document.querySelectorAll('#v-out .f3 tbody tr').length,
        shown: document.querySelector('#v-out').classList.contains('show') };
    })()`);
    check('the .NET bridge rejects a reversed numeric range inline',
      reversed.role === 'alert' && reversed.text.length > 0 && reversed.rows === 0 && reversed.shown, reversed);
    await evaluate(`(() => {
      const values = document.querySelectorAll('#v-out .filter-value-host [data-field]');
      values[0].textContent = '2'; values[1].textContent = '10';
      Array.from(document.querySelectorAll('#v-out .fgrid > .btn')).find((node) => node.textContent === '追加').click();
      return true;
    })()`);
    await waitFor(client, `document.querySelectorAll('#v-out .f3 tbody tr').length === 1`,
      30000, 'the valid number range');
    const validNumber = await evaluate(`Array.from(document.querySelectorAll('#v-out .f3 tbody tr td')).map((node) => node.textContent)`);
    check('a valid numeric range is rendered canonically',
      validNumber[1] === '範囲' && validNumber[2] === '2 ～ 10', validNumber);
    await evaluate(`Array.from(document.querySelectorAll('#v-out .btn')).find((node) => node.textContent === '削除').click(); true`);
    check('export destination is an editable CSV textbox',
      exportStart.destination[0] === 'textbox' && exportStart.destination[1] === 'true' &&
        /\.csv$/i.test(exportStart.destination[2]), exportStart.destination);
    await press(client, 'Escape', 0);
    await waitFor(client, `!document.querySelector('#v-out').classList.contains('show')`,
      30000, 'export to cancel');
    check('export cancel restores focus to its launch button',
      await evaluate(`document.activeElement.id === 'b-out'`),
      await evaluate(`document.activeElement.id`));

    process.stdout.write('\nprocess and candidate DOM\n');
    await openModal('#b-upd', '#v-upd');
    const updateFit = await modalFit('#v-upd');
    const update = await evaluate(`(() => {
      const veil = document.querySelector('#v-upd');
      const tables = veil.querySelectorAll('table');
      return { legends: Array.from(veil.querySelectorAll('.body > fieldset > legend')).map((node) => node.textContent),
        rows: Array.from(tables).map((node) => node.querySelectorAll('tbody tr').length),
        heads: Array.from(tables).map((node) => Array.from(node.querySelectorAll('thead th')).map((cell) => cell.childNodes[0].textContent)),
        firstStep: Array.from(tables[1].querySelectorAll('tbody tr:first-child td')).map((node) => node.textContent),
        outputs: Array.from(veil.querySelectorAll('.body > fieldset:last-of-type .kv')).map((node) => node.textContent),
        execute: (() => { const node = veil.querySelector('[data-modal-default=true]');
          return [node.textContent, node.getAttribute('aria-disabled'), node.title]; })(),
        jobControls: veil.querySelectorAll('[data-field*=job],select').length };
    })()`);
    check('update has three fixed groups and no job picker',
      update.legends.length === 3 && update.jobControls === 0 &&
        update.legends[1] === '処理内容' && update.legends[2] === '書き出し先', update);
    check('update input and step tables are 3x5 and 3x7',
      update.rows.join('|') === '3|3' && update.heads[0].length === 5 && update.heads[1].length === 7,
      { rows: update.rows, heads: update.heads });
    check('update step text comes from operation, key, condition, and output',
      update.firstStep[1] === '結合' && update.firstStep[4] === '番号1 = 番号1' &&
        update.firstStep[5] === '左外部' && update.firstStep[6] === '中間1', update.firstStep);
    check('update has three outputs and an enabled default action',
      update.outputs.length === 3 && update.execute[0] === '実行' &&
        update.execute[1] === 'false' && updateFit.inside && updateFit.bodyHorizontal,
      { update, updateFit });
    await closeModal('#v-upd');

    await openModal('#b-del', '#v-del');
    const deleteFit = await modalFit('#v-del');
    const deleting = await evaluate(`(() => {
      const veil = document.querySelector('#v-del');
      const tables = veil.querySelectorAll('table');
      const execute = veil.querySelector('[data-modal-default=true]');
      return { rows: Array.from(tables).map((node) => node.querySelectorAll('tbody tr').length),
        heads: Array.from(tables).map((node) => node.querySelectorAll('thead th').length),
        firstHead: tables[0].querySelector('thead th').childNodes[0].textContent,
        execute: [execute.textContent, execute.getAttribute('aria-disabled'), execute.title] };
    })()`);
    check('delete input and step tables are 2x5 and 4x7',
      deleting.rows.join('|') === '2|4' && deleting.heads.join('|') === '5|7' &&
        deleting.firstHead === '指定', deleting);
    check('delete has an enabled action and an unclipped dialog',
      deleting.execute[0] === '削除する' && deleting.execute[1] === 'false' &&
        deleteFit.inside && deleteFit.bodyHorizontal, { deleting, deleteFit });
    await closeModal('#v-del');

    await evaluate(`(() => {
      const input = document.querySelector('#input');
      input.textContent = ${JSON.stringify(multiKey)};
      input.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText' }));
      input.focus();
      document.querySelector('#b-search').click();
      return true;
    })()`);
    await waitFor(client, `document.querySelector('#v-cand').classList.contains('show')`,
      30000, 'the candidate list');
    const candidateFit = await modalFit('#v-cand');
    const candidate = await evaluate(`(() => {
      const veil = document.querySelector('#v-cand');
      const rows = veil.querySelectorAll('tbody tr');
      return { label: veil.querySelector('.dlg').getAttribute('aria-label'),
        hint: veil.querySelector('.hint').textContent,
        rows: rows.length, columns: veil.querySelectorAll('thead th').length,
        grips: veil.querySelectorAll('thead .gr').length,
        options: veil.querySelectorAll('tbody tr[role=option]').length,
        selected: veil.querySelectorAll('tbody tr[aria-selected=true]').length,
        focus: document.activeElement.dataset.index };
    })()`);
    check('candidate modal reports all five matching records',
      candidate.label === '候補一覧' && candidate.hint.includes('該当 5 件') &&
        candidate.rows === 5 && candidateFit.inside && candidateFit.bodyHorizontal,
      { candidate, candidateFit });
    check('candidate table has ten resizable configured columns',
      candidate.columns === 10 && candidate.grips === 10, candidate);
    check('candidate rows expose one keyboard selection',
      candidate.options === 5 && candidate.selected === 1 && candidate.focus === '0', candidate);
    await press(client, 'ArrowDown', 0);
    const movedCandidate = await evaluate(`(() => {
      const row = document.querySelector('#v-cand tbody tr[aria-selected=true]');
      return { selected: row.dataset.index, focus: document.activeElement.dataset.index,
        identity: row.querySelector('td:nth-child(2)').textContent };
    })()`);
    check('candidate ArrowDown moves selection and focus together',
      movedCandidate.selected === '1' && movedCandidate.focus === '1', movedCandidate);
    await press(client, 'Enter', 0);
    await waitFor(client, `!document.querySelector('#v-cand').classList.contains('show')`,
      30000, 'the candidate acceptance');
    await waitFor(client,
      `Array.from(document.querySelectorAll('.fld')).some((node) =>
        node.textContent.trim() === ${JSON.stringify(movedCandidate.identity)})`,
      30000,
      'the focused candidate record');
    const accepted = await evaluate(`!document.querySelector('#v-cand').classList.contains('show')`);
    const acceptedIdentity = await evaluate(`Array.from(document.querySelectorAll('.fld')).some((node) =>
      node.textContent.trim() === ${JSON.stringify(movedCandidate.identity)})`);
    check('candidate Enter accepts and renders the focused record',
      accepted && acceptedIdentity, { accepted, identity: movedCandidate.identity, acceptedIdentity });

    process.stdout.write(`\n${passed} passed, ${failed} failed\n`);
    process.stdout.write(failed ? 'RESULT: FAIL\n' : 'RESULT: PASS\n');
    if (failed) { process.exitCode = 1; }
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

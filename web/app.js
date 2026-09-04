(function () {
  'use strict';

  var stage = document.querySelector('.stage');
  var screen = null;
  var state = null;
  var input = null;
  var statusSegments = [];
  var currentModal = null;
  var currentToken = 0;
  var modalReturnFocus = null;
  var settingsContent = null;
  var exportContent = null;

  function post(message) {
    if (window.chrome && window.chrome.webview) {
      window.chrome.webview.postMessage(message);
    }
  }

  function element(tag, className, text) {
    var node = document.createElement(tag);
    if (className) { node.className = className; }
    if (text !== undefined && text !== null) { node.textContent = text; }
    return node;
  }

  function setMnemonic(node, text) {
    node.textContent = '';
    text = text || '';
    var match = /&([A-Za-z0-9])/.exec(text);
    if (!match) { node.textContent = text; return; }
    var at = match.index;
    node.appendChild(document.createTextNode(text.slice(0, at)));
    var underline = element('u', '', match[1]);
    node.appendChild(underline);
    node.appendChild(document.createTextNode(text.slice(at + 2)));
  }

  function px(value) { return String(Number(value) || 0) + 'px'; }

  function box(value) {
    if (!value || !value.length) { return ''; }
    return value.map(px).join(' ');
  }

  function activate(node, callback) {
    node.addEventListener('click', function (event) {
      if (node.getAttribute('aria-disabled') === 'true') { return; }
      callback(event);
    });
    node.addEventListener('keydown', function (event) {
      if ((event.key === 'Enter' || event.key === ' ') &&
          node.getAttribute('aria-disabled') !== 'true') {
        event.preventDefault();
        callback(event);
      }
    });
  }

  function button(definition, extraClass) {
    var action = definition.action || '';
    var node = element('div', (action === 'workState' ? 'tog' : 'btn') +
      (definition.primary ? ' def' : '') + (extraClass ? ' ' + extraClass : ''));
    node.tabIndex = 0;
    node.setAttribute('role', 'button');
    node.setAttribute('data-action', action);
    node.setAttribute('data-job', definition.job || '');
    node.id = actionId(action);
    if (definition.tip) { node.title = definition.tip; node.setAttribute('aria-label', definition.tip); }
    setMnemonic(node, definition.text || action);
    activate(node, function () {
      post({ type: 'action', name: action, job: definition.job || '', key: keyValue() });
    });
    return node;
  }

  function actionId(action) {
    var ids = {
      search: 'b-search', clear: 'b-clear', workState: 'b-work',
      tableExport: 'b-out', updateRecords: 'b-upd', deleteRecords: 'b-del',
      settings: 'b-set', sendChanges: 'b-send', refreshLedger: 'b-refresh'
    };
    return ids[action] || ('b-' + action);
  }

  function applyMargin(node, definition) {
    if (definition.margin) { node.style.margin = box(definition.margin); }
  }

  function boundField(id, className) {
    var field = element('div', 'fld ' + (className || ''));
    field.setAttribute('data-bind', id);
    return field;
  }

  function fieldset(title) {
    var set = element('fieldset', 'gb');
    set.appendChild(element('legend', '', title || ''));
    return set;
  }

  function renderKeyPanel(definition) {
    var set = fieldset(definition.title);
    var line = element('div', 'keyline');
    var figure = element('div', 'keyfigure');
    figure.appendChild(element('div', 'lab', definition.label || ''));
    figure.appendChild(boundField(definition.value, 'keyvalue'));
    line.appendChild(figure);
    var actions = element('div', 'keyactions');
    var label = element('label');
    setMnemonic(label, definition.inputLabel || '');
    actions.appendChild(label);
    input = boundField('', 'inp');
    input.id = 'input';
    input.contentEditable = 'true';
    input.tabIndex = 0;
    input.setAttribute('role', 'textbox');
    input.setAttribute('aria-label', (definition.inputLabel || '').replace(/&/g, ''));
    input.setAttribute('data-max-length', String(definition.maxLength || 64));
    input.style.width = px(definition.inputWidth);
    input.style.flex = 'none';
    if (definition.placeholder) { input.setAttribute('data-placeholder', definition.placeholder); }
    input.addEventListener('input', onKeyInput);
    input.addEventListener('keydown', function (event) {
      if (event.key === 'Enter') {
        event.preventDefault();
        post({ type: 'action', name: 'search', job: '', key: keyValue() });
      }
    });
    actions.appendChild(input);
    (definition.buttons || []).forEach(function (entry) { actions.appendChild(button(entry)); });
    line.appendChild(actions);
    set.appendChild(line);
    applyMargin(set, definition);
    return set;
  }

  function onKeyInput() {
    var max = Number(input.getAttribute('data-max-length')) || 64;
    var value = input.textContent.replace(/[\r\n]/g, '');
    if (value.length > max) { value = value.slice(0, max); input.textContent = value; placeCaretEnd(input); }
    post({ type: 'key', value: value });
  }

  function keyValue() { return input ? input.textContent.trim() : ''; }

  function placeCaretEnd(node) {
    var range = document.createRange();
    range.selectNodeContents(node);
    range.collapse(false);
    var selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
  }

  function renderFieldList(definition) {
    var set = fieldset(definition.title);
    (definition.rows || []).forEach(function (row) {
      var line = element('div', 'row dynamic-row');
      line.style.setProperty('--row-height', px(definition.rowHeight));
      line.style.setProperty('--label-width', px(definition.labelWidth));
      line.appendChild(element('label', '', row.label || ''));
      line.appendChild(boundField(row.value, 'r v'));
      set.appendChild(line);
    });
    applyMargin(set, definition);
    return set;
  }

  function renderTextBox(definition) {
    var set = fieldset(definition.title);
    var value = boundField(definition.value, 'v txt dynamic-text');
    value.style.setProperty('--text-height', px((Number(definition.lines) || 1) * 14 + 8));
    set.appendChild(value);
    applyMargin(set, definition);
    return set;
  }

  function renderColumns(definition) {
    var columns = element('div', 'cols');
    columns.setAttribute('data-stack-below', String(definition.stackBelow || 0));
    columns.style.setProperty('--section-gap', px(definition.gap));
    (definition.items || []).forEach(function (entry, index) {
      var item = renderSection(entry);
      var weight = definition.weights && definition.weights[index] ? definition.weights[index] : 1;
      item.style.flex = String(weight) + ' 1 0';
      columns.appendChild(item);
    });
    applyMargin(columns, definition);
    return columns;
  }

  function renderStatusBand(definition) {
    var band = element('div', 'band');
    band.style.setProperty('--band-height', px(definition.height));
    band.setAttribute('data-judgment', definition.judgment);
    var center = element('span');
    center.style.flex = '1';
    center.style.textAlign = 'center';
    var result = element('span', 'ok unsearched');
    result.id = 'judge';
    var sub = element('span');
    sub.id = 'jsub';
    center.appendChild(result);
    center.appendChild(sub);
    band.appendChild(center);
    applyMargin(band, definition);
    return band;
  }

  function renderSendBar(definition) {
    var bar = element('div', 'sendbar');
    bar.style.setProperty('--send-height', px(definition.height));
    var count = boundField(definition.value, 'sendn');
    count.className = 'sendn';
    count.id = 'sn';
    bar.appendChild(count);
    (definition.buttons || []).forEach(function (entry) { bar.appendChild(button(entry)); });
    applyMargin(bar, definition);
    return bar;
  }

  function renderCommandBar(definition) {
    var bar = element('div', 'commandbar');
    (definition.buttons || []).forEach(function (entry) { bar.appendChild(button(entry)); });
    applyMargin(bar, definition);
    return bar;
  }

  function renderStatusBar(definition) {
    var bar = element('div', 'sb');
    bar.style.setProperty('--status-height', px(definition.height));
    bar.setAttribute('role', 'status');
    bar.setAttribute('aria-live', 'polite');
    statusSegments = [];
    (definition.segments || []).forEach(function (entry, index) {
      var segment = element('div', 'sp');
      segment.setAttribute('data-prefix', entry.prefix || '');
      segment.setAttribute('data-bind', entry.value);
      if (entry.clock) { segment.setAttribute('data-clock', 'true'); }
      if (entry.bold) { segment.style.fontWeight = 'bold'; }
      if (entry.dot) { segment.setAttribute('data-dot', 'true'); }
      if (index === 2) { segment.style.flex = '1 1 0'; }
      else { segment.style.flex = '0 1 auto'; }
      statusSegments.push(segment);
      bar.appendChild(segment);
    });
    bar.appendChild(element('div', 'grip'));
    return bar;
  }

  function renderSection(definition) {
    if (definition.type === 'keyPanel') { return renderKeyPanel(definition); }
    if (definition.type === 'columns') { return renderColumns(definition); }
    if (definition.type === 'fieldList') { return renderFieldList(definition); }
    if (definition.type === 'textBox') { return renderTextBox(definition); }
    if (definition.type === 'statusBand') { return renderStatusBand(definition); }
    if (definition.type === 'sendBar') { return renderSendBar(definition); }
    return element('div');
  }

  function renderTitle(definition) {
    var bar = element('div', 'tb');
    bar.appendChild(element('div', 'ic', 'R'));
    bar.appendChild(element('div', 'ttl', definition && definition.brand ? definition.brand : 'Reader Data Viewer'));
    var tags = element('div', 'tb-tags');
    if (definition) {
      (definition.tags || []).forEach(function (entry) {
        tags.appendChild(element('div', 'tb-tag ' + (entry.look || ''), entry.text));
      });
      (definition.buttons || []).forEach(function (entry) { tags.appendChild(button(entry, 'sm')); });
    }
    bar.appendChild(tags);
    [['minimize', '_'], ['maximize', '□'], ['close', '✕']].forEach(function (entry) {
      var control = element('div', 'bx' + (entry[0] === 'close' ? ' cl' : ''), entry[1]);
      control.tabIndex = 0;
      control.setAttribute('role', 'button');
      control.setAttribute('aria-label', entry[0]);
      activate(control, function () { post({ type: 'window', command: entry[0] }); });
      bar.appendChild(control);
    });
    bar.addEventListener('mousedown', function (event) {
      if (event.button === 0 && !event.target.closest('.bx,.btn,.tog')) {
        post({ type: 'window', command: 'drag' });
      }
    });
    bar.addEventListener('dblclick', function (event) {
      if (!event.target.closest('.bx,.btn,.tog')) { post({ type: 'window', command: 'maximize' }); }
    });
    return bar;
  }

  function renderScreen(definition) {
    screen = definition;
    stage.classList.add('runtime');
    var card = definition.card;
    stage.style.setProperty('--card-width', px(card.width));
    stage.style.setProperty('--card-gap', px(card.gap));
    var padding = card.padding || [0, 0, 0, 0];
    stage.style.setProperty('--card-pad-t', px(padding[0]));
    stage.style.setProperty('--card-pad-r', px(padding[1]));
    stage.style.setProperty('--card-pad-b', px(padding[2]));
    stage.style.setProperty('--card-pad-l', px(padding[3]));
    stage.style.setProperty('--ff', '"' + String(card.font).replace(/"/g, '') + '","Segoe UI",sans-serif');
    stage.style.setProperty('--body-font', String(card.fontSize) + 'pt');
    stage.style.setProperty('--key-font', String(card.keyValueFontSize) + 'pt');
    stage.style.setProperty('--key-height', px(Math.max(34, Math.ceil(Number(card.keyValueFontSize) * 4 / 3 * 1.35 + 6))));
    stage.style.setProperty('--judgment-font', String(card.judgmentFontSize) + 'pt');
    stage.style.setProperty('--judgment-height', px(Math.max(30, Math.ceil(Number(card.judgmentFontSize) * 4 / 3 * 1.25 + 8))));
    stage.style.setProperty('--unsearched-font', String(card.unsearchedFontSize) + 'pt');

    var windowNode = stage.querySelector('.win');
    windowNode.textContent = '';
    var titleDefinition = null;
    var statusDefinition = null;
    definition.sections.forEach(function (entry) {
      if (entry.type === 'titleBar') { titleDefinition = entry; }
      if (entry.type === 'statusBar') { statusDefinition = entry; }
    });
    windowNode.appendChild(renderTitle(titleDefinition));
    var client = element('div', 'client');
    var stack = element('div', 'stack');
    var commandsPlaced = false;
    definition.sections.forEach(function (entry) {
      if (entry.type === 'titleBar') { return; }
      if (entry.type === 'statusBar') { return; }
      if (entry.type === 'sendBar' && statusDefinition && statusDefinition.buttons && statusDefinition.buttons.length) {
        stack.appendChild(renderCommandBar(statusDefinition));
        stack.appendChild(element('div', 'sep'));
        commandsPlaced = true;
      }
      stack.appendChild(renderSection(entry));
    });
    if (!commandsPlaced && statusDefinition && statusDefinition.buttons && statusDefinition.buttons.length) {
      stack.appendChild(renderCommandBar(statusDefinition));
    }
    client.appendChild(stack);
    windowNode.appendChild(client);
    if (statusDefinition) { windowNode.appendChild(renderStatusBar(statusDefinition)); }
    updateColumns();
  }

  function updateColumns() {
    Array.prototype.forEach.call(stage.querySelectorAll('.cols'), function (node) {
      var limit = Number(node.getAttribute('data-stack-below')) || 0;
      node.classList.toggle('stacked', limit > 0 && stage.clientWidth < limit);
    });
  }

  function applyState(next) {
    state = next;
    var values = next.values || {};
    Object.keys(values).forEach(function (id) {
      Array.prototype.forEach.call(stage.querySelectorAll('[data-bind="' + cssEscape(id) + '"]'), function (node) {
        var item = values[id];
        var prefix = node.getAttribute('data-prefix') || '';
        var visible = item.tone === 1 ? '' : (item.text || '');
        var text = prefix + (node.getAttribute('data-dot') === 'true' ? '● ' : '') + visible;
        node.textContent = text;
        node.classList.toggle('tone-muted', item.tone === 1);
        node.classList.toggle('tone-error', item.tone === 2);
        node.title = text;
      });
    });
    if (input && document.activeElement !== input && input.textContent !== (next.key || '')) {
      input.textContent = next.key || '';
    }
    Array.prototype.forEach.call(stage.querySelectorAll('[data-action]'), function (node) {
      var action = node.getAttribute('data-action');
      var enabled = action === 'settings' ? true : action === 'workState' ? next.workEnabled : next.opsEnabled;
      node.classList.toggle('dis', !enabled);
      node.setAttribute('aria-disabled', enabled ? 'false' : 'true');
      node.tabIndex = enabled ? 0 : -1;
    });
    var work = stage.querySelector('#b-work');
    if (work) {
      setMnemonic(work, next.workText || '');
      work.classList.toggle('on', !!next.workDown);
      work.setAttribute('aria-pressed', next.workDown ? 'true' : 'false');
    }
    var pending = stage.querySelector('#sn');
    if (pending) { pending.classList.toggle('warn', Number(next.pending) > 0); }
    var judgments = next.judgments || {};
    Object.keys(judgments).forEach(function (id) {
      var band = stage.querySelector('[data-judgment="' + cssEscape(id) + '"]');
      if (!band) { return; }
      var result = judgments[id];
      var head = band.querySelector('.ok');
      var sub = band.querySelector('#jsub');
      head.className = 'ok ' + (result.look || 'unsearched');
      head.textContent = result.text || '';
      sub.textContent = result.sub ? ' ' + result.sub : '';
    });
    if (statusSegments.length) {
      statusSegments.forEach(function (node) { node.classList.remove('notice', 'error'); });
      if (next.notice) {
        var target = statusSegments[Math.min(2, statusSegments.length - 1)];
        target.textContent = next.notice;
        target.title = next.notice;
        target.classList.add('notice');
        if (next.noticeError) { target.classList.add('error'); }
      }
    }
  }

  function cssEscape(value) {
    return String(value).replace(/(["\\])/g, '\\$1');
  }

  function setDisabled(node, disabled) {
    node.classList.toggle('dis', !!disabled);
    node.setAttribute('aria-disabled', disabled ? 'true' : 'false');
    node.tabIndex = disabled ? -1 : 0;
  }

  function modalButton(text, primary, callback, disabled) {
    var node = element('div', 'btn' + (primary ? ' def' : ''));
    setMnemonic(node, text);
    node.setAttribute('role', 'button');
    setDisabled(node, !!disabled);
    activate(node, callback);
    return node;
  }

  function modalShell(veilId, title) {
    closeVeils(false);
    var veil = document.getElementById(veilId);
    var dialog = veil.querySelector('.dlg');
    dialog.setAttribute('role', 'dialog');
    dialog.setAttribute('aria-modal', 'true');
    dialog.setAttribute('aria-label', title || '');
    dialog.querySelector('.ttl').textContent = title || '';
    var close = dialog.querySelector('.tb .cl');
    close.tabIndex = 0;
    close.setAttribute('role', 'button');
    close.setAttribute('aria-label', 'close');
    close.onclick = function () { finishModal({ ok: false }); };
    close.onkeydown = function (event) {
      if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); finishModal({ ok: false }); }
    };
    var body = dialog.querySelector('.body');
    body.textContent = '';
    veil.classList.add('show');
    currentModal = veil;
    return { veil: veil, dialog: dialog, body: body };
  }

  function finishModal(result) {
    if (!currentToken) { return; }
    var token = currentToken;
    closeVeils(true);
    post({ type: 'modalResult', token: token, result: result || { ok: false } });
  }

  function closeVeils(restore) {
    Array.prototype.forEach.call(stage.querySelectorAll('.veil'), function (veil) { veil.classList.remove('show'); });
    currentModal = null;
    if (restore && modalReturnFocus && document.contains(modalReturnFocus)) { modalReturnFocus.focus(); }
  }

  function openConfirm(content) {
    var shell = modalShell('v-send', content.title);
    var message = element('div', 'modal-message');
    message.appendChild(element('div', 'modal-symbol' + (content.ask ? '' : ' warn'), content.ask ? '?' : '!'));
    var text = element('div');
    text.style.flex = '1';
    text.style.minWidth = '0';
    String(content.body || '').split(/\r?\n/).forEach(function (line, index) {
      var part = element('div', '', line || ' ');
      if (index) { part.style.marginTop = '7px'; }
      text.appendChild(part);
    });
    message.appendChild(text);
    shell.body.appendChild(message);
    var foot = element('div', 'foot');
    if (content.ask) {
      foot.appendChild(modalButton('はい(&Y)', true, function () { finishModal({ ok: true }); }));
      foot.appendChild(modalButton('いいえ(&N)', false, function () { finishModal({ ok: false }); }));
    } else {
      foot.appendChild(modalButton('OK', true, function () { finishModal({ ok: true }); }));
    }
    shell.body.appendChild(foot);
  }

  function tableNode(columns, rows, options) {
    var holder = element('div', 'lv' + (options.readOnly ? ' ro' : ''));
    if (options.maxHeight) { holder.style.maxHeight = px(options.maxHeight); }
    var table = element('table');
    var head = element('thead');
    var headRow = element('tr');
    columns.forEach(function (column) {
      var th = element('th', column.align === 'right' ? 'r' : '', column.header || '');
      if (column.width) { th.style.width = px(column.width); }
      if (options.headerHeight) { th.style.height = px(options.headerHeight); }
      headRow.appendChild(th);
    });
    head.appendChild(headRow);
    table.appendChild(head);
    var body = element('tbody');
    rows.forEach(function (row, rowIndex) {
      var tr = element('tr');
      tr.tabIndex = options.readOnly ? -1 : 0;
      tr.setAttribute('data-index', String(rowIndex));
      tr.setAttribute('aria-selected', rowIndex === options.selected ? 'true' : 'false');
      if (options.rowHeight) { tr.style.height = px(options.rowHeight); }
      row.forEach(function (cell, index) {
        var column = columns[index] || {};
        var td = element('td', column.align === 'right' ? 'r' : '');
        if (cell && typeof cell === 'object') {
          var visible = cell.tone === 1 ? '' : (cell.text || '');
          if (column.render === 'tag') { td.appendChild(element('span', 'tag ' + (cell.look || ''), visible)); }
          else { td.textContent = visible; }
          if (cell.tone === 1 || column.muted) { td.style.color = '#808080'; }
          if (cell.tone === 2) { td.classList.add('bad'); }
          if (column.bold) { td.style.fontWeight = 'bold'; }
        } else { td.textContent = cell === undefined || cell === null ? '' : String(cell); }
        tr.appendChild(td);
      });
      if (!options.readOnly) {
        activate(tr, function () {
          Array.prototype.forEach.call(body.children, function (other) { other.setAttribute('aria-selected', 'false'); });
          tr.setAttribute('aria-selected', 'true');
          if (options.onSelect) { options.onSelect(rowIndex); }
        });
        tr.addEventListener('keydown', function (event) {
          if (event.key === 'Enter' && options.onAccept) { options.onAccept(rowIndex); }
        });
        tr.addEventListener('dblclick', function () { if (options.onAccept) { options.onAccept(rowIndex); } });
      }
      body.appendChild(tr);
    });
    table.appendChild(body);
    holder.appendChild(table);
    makeResizable(holder);
    return holder;
  }

  function openCandidates(content, shared, tellOnly) {
    var shell = modalShell(shared ? 'v-shared' : 'v-cand', content.title);
    shell.dialog.style.width = px(content.width || 744);
    var hint = content.hint || '';
    if (!shared) { hint += '  該当 ' + String(content.total || 0) + ' 件'; }
    shell.body.appendChild(element('div', 'hint', hint));
    var selected = content.selected >= 0 ? content.selected : (content.rows.length ? 0 : -1);
    var list = tableNode(content.columns, content.rows, {
      readOnly: !content.selectable,
      maxHeight: content.maxHeight,
      rowHeight: content.rowHeight,
      headerHeight: content.headerHeight,
      selected: selected,
      onSelect: function (index) { selected = index; },
      onAccept: function (index) { if (content.selectable) { finishModal({ ok: true, index: index }); } }
    });
    shell.body.appendChild(list);
    var foot = element('div', 'foot');
    foot.appendChild(modalButton('OK', true, function () {
      finishModal(content.selectable ? { ok: selected >= 0, index: selected } : { ok: true });
    }, content.selectable && selected < 0));
    if (!tellOnly) { foot.appendChild(modalButton('キャンセル', false, function () { finishModal({ ok: false }); })); }
    shell.body.appendChild(foot);
  }

  function openUnmatched(content) {
    var shell = modalShell('v-unm', content.title);
    shell.body.appendChild(element('div', 'hint', content.body));
    shell.body.appendChild(tableNode([
      { header: '#', width: 42, align: 'right' },
      { header: 'キー', width: 210 },
      { header: '理由' }
    ], content.rows, { readOnly: true }));
    var foot = element('div', 'foot');
    foot.appendChild(modalButton('OK', true, function () { finishModal({ ok: true }); }));
    shell.body.appendChild(foot);
  }

  function openProcess(content, deleting) {
    var shell = modalShell(deleting ? 'v-del' : 'v-upd', content.title);
    shell.body.appendChild(element('div', 'hint', content.hint));
    var inputs = fieldset(content.inputTitle);
    inputs.appendChild(tableNode([
      { header: deleting ? '指定' : '表', width: 40 },
      { header: 'ファイル', width: 150 },
      { header: 'キー', width: 62 },
      { header: '行数', width: 74, align: 'right' },
      { header: '検証' }
    ], content.inputs.map(function (entry) {
      return [entry.id, entry.file, entry.key, entry.rows, { text: entry.validation, tone: entry.valid ? 0 : 2 }];
    }), { readOnly: true }));
    shell.body.appendChild(inputs);
    shell.body.appendChild(element('div', 'process-gap'));
    var steps = fieldset('処理内容');
    steps.appendChild(tableNode([
      { header: '#', width: 26, align: 'right' }, { header: '操作', width: 46 },
      { header: '対象1', width: 64 }, { header: '対象2', width: 76 },
      { header: 'キー', width: 120 }, { header: '条件', width: 240 }, { header: '出力' }
    ], content.steps, { readOnly: true, maxHeight: 128 }));
    shell.body.appendChild(steps);
    shell.body.appendChild(element('div', 'process-gap'));
    var output = fieldset('書き出し先');
    ['パス', 'ファイル名', '最終更新'].forEach(function (label, index) {
      var row = element('div', 'kv');
      row.appendChild(element('label', '', label));
      var value = element('div', 'fld', content.output[index] || '');
      value.style.flex = '1';
      row.appendChild(value);
      output.appendChild(row);
    });
    shell.body.appendChild(output);
    var foot = element('div', 'foot');
    foot.appendChild(modalButton(content.executeText, true, function () { finishModal({ ok: true }); }, !content.canRun));
    foot.appendChild(modalButton('キャンセル', false, function () { finishModal({ ok: false }); }));
    shell.body.appendChild(foot);
  }

  function editable(value, field, width) {
    var node = element('div', 'fld inp', value || '');
    node.contentEditable = 'true';
    node.tabIndex = 0;
    node.setAttribute('role', 'textbox');
    node.setAttribute('data-field', field || '');
    node.style.flex = width ? 'none' : '1';
    if (width) { node.style.width = px(width); }
    return node;
  }

  function browseButton(field, kind, valueNode) {
    return modalButton('参照...', false, function () {
      post({ type: 'browse', token: currentToken, field: field, kind: kind, value: valueNode.textContent.trim() });
    });
  }

  function openSettings(content) {
    settingsContent = content;
    var shell = modalShell('v-set', content.title);
    shell.body.appendChild(element('div', 'hint', content.hint));
    var paths = fieldset('場所');
    [['データ', 'dataDir', 'folder'], ['統合台帳', 'ledger', 'ledger'], ['ログ', 'log', 'log']].forEach(function (entry) {
      var row = element('div', 'kv');
      row.appendChild(element('label', '', entry[0]));
      var edit = editable(content[entry[1]], entry[1]);
      row.appendChild(edit);
      row.appendChild(browseButton(entry[1], entry[2], edit));
      paths.appendChild(row);
    });
    shell.body.appendChild(paths);
    shell.body.appendChild(element('div', 'process-gap'));
    var search = fieldset('検索');
    var patternRow = element('div', 'kv');
    patternRow.appendChild(element('label', '', '番号の形式'));
    patternRow.appendChild(editable(content.pattern, 'pattern'));
    search.appendChild(patternRow);
    var countRow = element('div', 'kv');
    countRow.appendChild(element('label', '', '候補の表示件数'));
    countRow.appendChild(editable(String(content.candidateRows), 'candidateRows', 70));
    search.appendChild(countRow);
    shell.body.appendChild(search);
    shell.body.appendChild(element('div', 'process-gap'));
    var watch = fieldset('監視対象');
    var targetRow = element('div', 'kv');
    targetRow.appendChild(element('label', '', '対象'));
    var target = element('div', 'fld', content.target.summary || '');
    target.setAttribute('data-target-summary', 'true');
    target.style.flex = '1';
    targetRow.appendChild(target);
    var pick = modalButton('画面から選ぶ', false, startPicker);
    pick.id = 'b-pick';
    targetRow.appendChild(pick);
    watch.appendChild(targetRow);
    var readRow = element('div', 'kv');
    readRow.appendChild(element('label', '', '読み取り'));
    var read = element('div', 'fld', content.target.read || '');
    read.setAttribute('data-target-read', 'true');
    read.style.flex = '1';
    readRow.appendChild(read);
    watch.appendChild(readRow);
    shell.body.appendChild(watch);
    var error = element('div', 'setting-error');
    error.setAttribute('role', 'alert');
    error.hidden = true;
    shell.body.appendChild(error);
    var foot = element('div', 'foot');
    foot.appendChild(modalButton('OK', true, function () { submitSettings(shell.body, error); }));
    foot.appendChild(modalButton('キャンセル', false, function () { finishModal({ ok: false }); }));
    shell.body.appendChild(foot);
  }

  function submitSettings(body, error) {
    function value(name) {
      var node = body.querySelector('[data-field="' + name + '"]');
      return node ? node.textContent.trim() : '';
    }
    var dataDir = value('dataDir');
    var ledger = value('ledger');
    var log = value('log');
    var pattern = value('pattern');
    var candidateRows = Number(value('candidateRows'));
    var invalid = null;
    if (!dataDir || !ledger || !log) { invalid = '場所は空欄にできません。'; }
    else {
      try { new RegExp(pattern); } catch (exception) { invalid = '番号の形式を確認してください。'; }
    }
    if (!invalid && (!Number.isInteger(candidateRows) || candidateRows < 1 || candidateRows > 1000)) {
      invalid = '候補の表示件数は 1～1000 の整数にしてください。';
    }
    if (invalid) {
      error.textContent = invalid;
      error.hidden = false;
      var first = !dataDir ? 'dataDir' : !ledger ? 'ledger' : !log ? 'log' :
        invalid.indexOf('番号の形式') >= 0 ? 'pattern' : 'candidateRows';
      body.querySelector('[data-field="' + first + '"]').focus();
      return;
    }
    finishModal({ ok: true, dataDir: dataDir, ledger: ledger, log: log,
      pattern: pattern, candidateRows: candidateRows });
  }

  function startPicker() {
    if (!currentModal) { return; }
    currentModal.classList.remove('show');
    var shell = modalShell('v-pick', '画面から選ぶ');
    shell.dialog.querySelector('.tb .cl').onclick = function () { post({ type: 'pickerCancel' }); };
    var how = element('div', 'how', '対象の欄にカーソルを合わせて Ctrl + Shift を押す');
    shell.body.appendChild(how);
    var values = element('div', 'picker-values');
    ['type', 'automationId', 'className', 'name', 'process', 'read'].forEach(function (field, index) {
      var labels = ['種類', 'AutomationId', 'クラス名', '名前', 'プロセス', '読み取り'];
      var row = element('div', 'kv');
      row.appendChild(element('label', '', labels[index]));
      var value = element('div', 'fld', '---');
      value.setAttribute('data-picker-field', field);
      row.appendChild(value);
      values.appendChild(row);
    });
    shell.body.appendChild(values);
    var foot = element('div', 'foot2');
    foot.appendChild(element('div', 'esc', 'Esc で中止'));
    foot.appendChild(modalButton('閉じる', false, function () {
      post({ type: 'pickerCancel' });
    }));
    shell.body.appendChild(foot);
    post({ type: 'picker', token: currentToken });
  }

  function pickerPreview(preview) {
    if (!currentModal || currentModal.id !== 'v-pick') { return; }
    Object.keys(preview || {}).forEach(function (field) {
      var node = currentModal.querySelector('[data-picker-field="' + field + '"]');
      if (node) { node.textContent = preview[field] || '---'; node.title = preview[field] || ''; }
    });
  }

  function pickerResult(target) {
    if (target && settingsContent) { settingsContent.target = target; }
    openSettings(settingsContent);
  }

  function openExport(content) {
    exportContent = content;
    var shell = modalShell('v-out', content.title);
    shell.body.appendChild(element('div', 'hint', content.hint));
    var byRef = {};
    content.fields.forEach(function (field) { byRef[field.ref] = field; });
    var selected = content.defaults.filter(function (reference) { return !!byRef[reference]; });
    var available = content.fields.map(function (field) { return field.ref; })
      .filter(function (reference) { return selected.indexOf(reference) < 0; });
    var picker = element('div');
    picker.style.display = 'flex';
    picker.style.gap = 'var(--card-gap)';
    picker.style.alignItems = 'stretch';
    var left = element('div'); left.style.flex = '1 1 0'; left.style.minWidth = '0';
    left.appendChild(element('div', 'lab', '出力できる項目'));
    var leftList = element('div', 'lb'); leftList.style.height = '180px'; left.appendChild(leftList);
    var mover = element('div'); mover.style.display = 'flex'; mover.style.flexDirection = 'column';
    mover.style.justifyContent = 'center'; mover.style.gap = 'var(--card-gap)';
    var right = element('div'); right.style.flex = '1 1 0'; right.style.minWidth = '0';
    var rightHead = element('div'); rightHead.style.display = 'flex'; rightHead.style.alignItems = 'center';
    var rightTitle = element('span', 'lab'); rightTitle.style.flex = '1'; rightHead.appendChild(rightTitle);
    var reset = modalButton('既定に戻す', false, function () {
      selected = content.defaults.filter(function (reference) { return !!byRef[reference]; });
      available = content.fields.map(function (field) { return field.ref; })
        .filter(function (reference) { return selected.indexOf(reference) < 0; });
      drawLists();
    });
    reset.classList.add('sm'); rightHead.appendChild(reset); right.appendChild(rightHead);
    var rightList = element('div', 'lb'); rightList.style.height = '180px'; right.appendChild(rightList);
    function listItem(reference, list, source, destination) {
      var item = element('div', '', byRef[reference].label);
      item.tabIndex = 0; item.setAttribute('role', 'option'); item.setAttribute('data-ref', reference);
      activate(item, function () {
        Array.prototype.forEach.call(list.children, function (other) { other.setAttribute('aria-selected', 'false'); });
        item.setAttribute('aria-selected', 'true');
      });
      item.addEventListener('dblclick', function () { moveItem(list, source, destination); });
      return item;
    }
    function drawLists() {
      leftList.textContent = ''; rightList.textContent = '';
      available.forEach(function (reference) { leftList.appendChild(listItem(reference, leftList, available, selected)); });
      selected.forEach(function (reference) { rightList.appendChild(listItem(reference, rightList, selected, available)); });
      rightTitle.textContent = '出力する項目（' + selected.length + '）';
    }
    function moveItem(list, source, destination) {
      var chosen = list.querySelector('[aria-selected=true]');
      if (!chosen) { return; }
      var reference = chosen.getAttribute('data-ref');
      var index = source.indexOf(reference);
      if (index >= 0) { source.splice(index, 1); destination.push(reference); drawLists(); }
    }
    var moveRight = modalButton('▶', false, function () { moveItem(leftList, available, selected); });
    moveRight.classList.add('sm'); moveRight.style.width = '34px';
    var moveLeft = modalButton('◀', false, function () { moveItem(rightList, selected, available); });
    moveLeft.classList.add('sm'); moveLeft.style.width = '34px';
    mover.appendChild(moveRight); mover.appendChild(moveLeft);
    picker.appendChild(left); picker.appendChild(mover); picker.appendChild(right);
    shell.body.appendChild(picker);
    drawLists();
    shell.body.appendChild(element('div', 'process-gap'));
    var filters = [];
    var filterSet = fieldset('絞り込み条件（すべてに一致）');
    var grid = element('div', 'fgrid');
    grid.appendChild(element('div', 'fh', '項目'));
    grid.appendChild(element('div', 'fh', '条件'));
    var valueHead = element('div', 'fh', '値'); valueHead.style.gridColumn = '3 / span 3'; grid.appendChild(valueHead);
    grid.appendChild(element('div'));
    var fieldSelect = selectNode(content.fields.map(function (field) { return { value: field.ref, text: field.label }; }));
    var operatorSelect = selectNode([]);
    var first = editable('', 'filterFirst'); first.style.flex = '';
    var mark = element('div', 'tilde', '～');
    var last = editable('', 'filterLast'); last.style.flex = '';
    var add = modalButton('追加', false, addFilter); add.classList.add('sm');
    grid.appendChild(fieldSelect); grid.appendChild(operatorSelect); grid.appendChild(first); grid.appendChild(mark); grid.appendChild(last); grid.appendChild(add);
    var listHost = element('div', 'row3');
    var filterTable = tableNode([{ header: '項目', width: 180 }, { header: '条件', width: 110 }, { header: '値' }], [], { readOnly: false });
    filterTable.classList.add('f3'); listHost.appendChild(filterTable); grid.appendChild(listHost);
    var remove = modalButton('削除', false, removeFilter); remove.classList.add('sm', 'row3b'); grid.appendChild(remove);
    filterSet.appendChild(grid); shell.body.appendChild(filterSet);
    fieldSelect.addEventListener('change', updateOperators);
    function updateOperators() {
      var field = byRef[fieldSelect.value];
      operatorSelect.textContent = '';
      var entries = field.kind === 'text' ? [
        ['contains', '含む'], ['equals', '等しい'], ['startsWith', '始まる'], ['notContains', '含まない']
      ] : [['range', '範囲']];
      entries.forEach(function (entry) { var option = element('option', '', entry[1]); option.value = entry[0]; operatorSelect.appendChild(option); });
      mark.classList.toggle('off', field.kind === 'text');
      last.classList.toggle('off', field.kind === 'text');
      last.contentEditable = field.kind === 'text' ? 'false' : 'true';
    }
    function addFilter() {
      var field = byRef[fieldSelect.value];
      var firstValue = first.textContent.trim();
      var lastValue = last.textContent.trim();
      if (!firstValue || (field.kind !== 'text' && !lastValue)) { first.focus(); return; }
      var filter = { field: field.ref, operator: operatorSelect.value, first: firstValue,
        last: field.kind === 'text' ? '' : lastValue };
      filters.push(filter); redrawFilters();
    }
    function redrawFilters() {
      var replacement = tableNode([{ header: '項目', width: 180 }, { header: '条件', width: 110 }, { header: '値' }],
        filters.map(function (entry) {
          var option = operatorSelect.querySelector('option[value="' + entry.operator + '"]');
          return [byRef[entry.field].label, option ? option.textContent : entry.operator,
            entry.last ? entry.first + ' ～ ' + entry.last : entry.first];
        }), { readOnly: false, selected: filters.length - 1 });
      replacement.classList.add('f3'); listHost.replaceChild(replacement, listHost.firstChild); filterTable = replacement;
    }
    function removeFilter() {
      var row = filterTable.querySelector('tbody tr[aria-selected=true]');
      var index = row ? Number(row.getAttribute('data-index')) : filters.length === 1 ? 0 : -1;
      if (index >= 0) { filters.splice(index, 1); redrawFilters(); }
    }
    updateOperators();
    var pathRow = element('div', 'kv export-path'); pathRow.style.marginTop = '9px';
    pathRow.appendChild(element('label', '', '出力先'));
    var destination = editable(content.destination, 'exportPath'); pathRow.appendChild(destination);
    pathRow.appendChild(browseButton('exportPath', 'export', destination)); shell.body.appendChild(pathRow);
    var error = element('div', 'setting-error'); error.hidden = true; error.setAttribute('role', 'alert'); shell.body.appendChild(error);
    var foot = element('div', 'foot');
    foot.appendChild(modalButton('OK', true, function () {
      if (!selected.length || !destination.textContent.trim()) {
        error.textContent = '出力する項目と出力先を指定してください。'; error.hidden = false; return;
      }
      finishModal({ ok: true, path: destination.textContent.trim(), fields: selected, filters: filters });
    }));
    foot.appendChild(modalButton('キャンセル', false, function () { finishModal({ ok: false }); }));
    shell.body.appendChild(foot);
  }

  function selectNode(entries) {
    var select = element('select', 'fld inp');
    select.style.width = '100%'; select.style.padding = '0 2px';
    entries.forEach(function (entry) { var option = element('option', '', entry.text); option.value = entry.value; select.appendChild(option); });
    return select;
  }

  function makeResizable(holder) {
    var table = holder.querySelector('table');
    if (!table || table.getAttribute('data-rz')) { return; }
    table.setAttribute('data-rz', '1');
    Array.prototype.forEach.call(table.querySelectorAll('thead th'), function (th) {
      var grip = element('div', 'gr');
      th.appendChild(grip);
      grip.addEventListener('mousedown', function (event) {
        event.preventDefault(); event.stopPropagation();
        var start = event.clientX; var width = th.getBoundingClientRect().width;
        function move(moveEvent) { th.style.width = px(Math.max(24, Math.round(width + moveEvent.clientX - start))); }
        function up() { document.removeEventListener('mousemove', move); document.removeEventListener('mouseup', up); document.body.classList.remove('rzing'); }
        document.body.classList.add('rzing'); document.addEventListener('mousemove', move); document.addEventListener('mouseup', up);
      });
    });
  }

  function patchModal(field, value) {
    if (!currentModal) { return; }
    var node = currentModal.querySelector('[data-field="' + cssEscape(field) + '"]');
    if (node) { node.textContent = value || ''; }
  }

  function openModal(message) {
    modalReturnFocus = document.activeElement;
    currentToken = Number(message.token) || 0;
    var modal = message.modal;
    var content = message.content || {};
    if (modal === 'confirm') { openConfirm(content); }
    else if (modal === 'candidates') { openCandidates(content, false, false); }
    else if (modal === 'shared') { openCandidates(content, true, false); }
    else if (modal === 'sharedTell') { openCandidates(content, true, true); }
    else if (modal === 'unmatched') { openUnmatched(content); }
    else if (modal === 'update') { openProcess(content, false); }
    else if (modal === 'delete') { openProcess(content, true); }
    else if (modal === 'settings') { openSettings(content); }
    else if (modal === 'export') { openExport(content); }
    window.setTimeout(function () {
      var focus = currentModal && currentModal.querySelector('[contenteditable=true],[tabindex="0"]');
      if (focus) { focus.focus(); }
      post({ type: 'modalShown', token: currentToken });
    }, 0);
  }

  function onMessage(event) {
    var message = event.data || {};
    if (message.type === 'init') {
      renderScreen(message.screen);
      applyState(message.state);
      post({ type: 'ready' });
    } else if (message.type === 'state') { applyState(message.state); }
    else if (message.type === 'keySet') {
      if (input) { input.textContent = message.value || ''; }
    } else if (message.type === 'modalOpen') { openModal(message); }
    else if (message.type === 'modalPatch' && Number(message.token) === currentToken) { patchModal(message.field, message.value); }
    else if (message.type === 'pickerPreview') { pickerPreview(message.preview); }
    else if (message.type === 'pickerResult' && Number(message.token) === currentToken) { pickerResult(message.target); }
  }

  document.addEventListener('keydown', function (event) {
    if (event.key === 'Escape' && currentModal && currentModal.id !== 'v-pick') {
      event.preventDefault(); finishModal({ ok: false }); return;
    }
    if (event.key === 'Tab' && currentModal) {
      var items = Array.prototype.filter.call(currentModal.querySelectorAll('[contenteditable=true],[tabindex="0"],select'), function (node) {
        return node.offsetParent !== null && node.getAttribute('aria-disabled') !== 'true';
      });
      if (!items.length) { return; }
      var index = items.indexOf(document.activeElement);
      if (event.shiftKey && index <= 0) { event.preventDefault(); items[items.length - 1].focus(); }
      else if (!event.shiftKey && index === items.length - 1) { event.preventDefault(); items[0].focus(); }
      return;
    }
    if (event.altKey && !event.ctrlKey && !event.metaKey && !currentModal) {
      var key = event.key.toUpperCase();
      var match = null;
      Array.prototype.some.call(stage.querySelectorAll('[data-action]'), function (node) {
        var underline = node.querySelector('u');
        if (underline && underline.textContent.toUpperCase() === key && node.getAttribute('aria-disabled') !== 'true') { match = node; return true; }
        return false;
      });
      if (match) { event.preventDefault(); match.click(); }
    }
  });

  window.addEventListener('resize', updateColumns);
  window.setInterval(function () {
    var now = new Date();
    var value = [now.getHours(), now.getMinutes(), now.getSeconds()].map(function (part) {
      return String(part).padStart(2, '0');
    }).join(':');
    Array.prototype.forEach.call(stage.querySelectorAll('[data-clock=true]'), function (node) { node.textContent = value; });
  }, 1000);
  if (window.chrome && window.chrome.webview) { window.chrome.webview.addEventListener('message', onMessage); }
  window.rdvBridge = { version: 1, render: renderScreen, state: applyState };
})();

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
  var pendingSettings = null;
  var pendingFilter = null;
  var activeCalendar = null;
  var mainFocusDone = false;

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
        if (input.getAttribute('aria-disabled') !== 'true') {
          post({ type: 'action', name: 'search', job: '', key: keyValue() });
        }
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
    band.appendChild(element('span', 'band-label', definition.label || ''));
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
    var controlHeight = Math.max(20, Math.ceil(Number(card.fontSize) * 4 / 3) + 7);
    stage.style.setProperty('--control-height', px(controlHeight));
    stage.style.setProperty('--control-button-width', px(controlHeight));
    stage.style.setProperty('--number-editor-width', px(Math.max(
      controlHeight * 3,
      Math.ceil(Number(card.fontSize) * 4 / 3) * 5)));
    stage.style.setProperty('--calendar-cell', px(Math.max(
      controlHeight,
      Math.ceil(Number(card.fontSize) * 4 / 3) + Number(card.gap) * 2)));

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
        var visible = item.text || '';
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
      var enabled = action === 'settings' ? true : action === 'workState' ?
        next.opsEnabled && next.workEnabled : next.opsEnabled;
      node.classList.toggle('dis', !enabled);
      node.setAttribute('aria-disabled', enabled ? 'false' : 'true');
      node.tabIndex = enabled ? 0 : -1;
    });
    if (input) {
      var inputEnabled = !!next.opsEnabled;
      input.contentEditable = inputEnabled ? 'true' : 'false';
      input.classList.toggle('dis', !inputEnabled);
      input.setAttribute('aria-disabled', inputEnabled ? 'false' : 'true');
      input.tabIndex = inputEnabled ? 0 : -1;
      if (!inputEnabled && document.activeElement === input) { input.blur(); }
      if (inputEnabled && !mainFocusDone && !currentModal) {
        mainFocusDone = true;
        window.setTimeout(function () {
          if (input && input.getAttribute('aria-disabled') !== 'true' && !currentModal) {
            input.focus();
            placeCaretEnd(input);
          }
        }, 0);
      }
    }
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
    if (primary) { node.setAttribute('data-modal-default', 'true'); }
    setDisabled(node, !!disabled);
    activate(node, callback);
    return node;
  }

  function modalShell(veilId, title) {
    closeCalendar(false);
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
    pendingSettings = null;
    pendingFilter = null;
    closeCalendar(false);
    closeVeils(true);
    post({ type: 'modalResult', token: token, result: result || { ok: false } });
  }

  function closeVeils(restore) {
    Array.prototype.forEach.call(stage.querySelectorAll('.veil'), function (veil) { veil.classList.remove('show'); });
    currentModal = null;
    activeCalendar = null;
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
    function selectRow(index, focus) {
      var children = body.children;
      if (!children.length) { return; }
      index = Math.max(0, Math.min(children.length - 1, index));
      Array.prototype.forEach.call(children, function (other, otherIndex) {
        var selected = otherIndex === index;
        other.setAttribute('aria-selected', selected ? 'true' : 'false');
        other.tabIndex = selected ? 0 : -1;
      });
      if (options.onSelect) { options.onSelect(index); }
      if (focus) {
        children[index].focus();
        children[index].scrollIntoView({ block: 'nearest' });
      }
    }
    rows.forEach(function (row, rowIndex) {
      var tr = element('tr');
      tr.tabIndex = options.readOnly ? -1 : rowIndex === options.selected ? 0 : -1;
      if (!options.readOnly) { tr.setAttribute('role', 'option'); }
      tr.setAttribute('data-index', String(rowIndex));
      tr.setAttribute('aria-selected', rowIndex === options.selected ? 'true' : 'false');
      if (options.rowHeight) { tr.style.height = px(options.rowHeight); }
      row.forEach(function (cell, index) {
        var column = columns[index] || {};
        var td = element('td', column.align === 'right' ? 'r' : '');
        if (cell && typeof cell === 'object') {
          var visible = cell.text || '';
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
          selectRow(rowIndex, false);
        });
        tr.addEventListener('keydown', function (event) {
          var next = rowIndex;
          if (event.key === 'ArrowUp') { next = rowIndex - 1; }
          else if (event.key === 'ArrowDown') { next = rowIndex + 1; }
          else if (event.key === 'Home') { next = 0; }
          else if (event.key === 'End') { next = rows.length - 1; }
          else if (event.key === 'Enter' && options.onAccept) {
            event.preventDefault();
            options.onAccept(rowIndex);
            return;
          } else { return; }
          event.preventDefault();
          selectRow(next, true);
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
    var okButton = null;
    var list = tableNode(content.columns, content.rows, {
      readOnly: !content.selectable,
      maxHeight: content.maxHeight,
      rowHeight: content.rowHeight,
      headerHeight: content.headerHeight,
      selected: selected,
      onSelect: function (index) {
        selected = index;
        if (okButton) { setDisabled(okButton, false); }
      },
      onAccept: function (index) { if (content.selectable) { finishModal({ ok: true, index: index }); } }
    });
    if (!shared || content.rows.length) { shell.body.appendChild(list); }
    var foot = element('div', 'foot');
    okButton = modalButton('OK', true, function () {
      finishModal(content.selectable ? { ok: selected >= 0, index: selected } : { ok: true });
    }, content.selectable && selected < 0);
    foot.appendChild(okButton);
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
    var execute = modalButton(content.executeText, true, function () { finishModal({ ok: true }); }, !content.canRun);
    if (!content.canRun && content.cannotRunText) {
      execute.title = content.cannotRunText;
      execute.setAttribute('aria-label', content.executeText + ': ' + content.cannotRunText);
    }
    foot.appendChild(execute);
    foot.appendChild(modalButton('キャンセル', false, function () { finishModal({ ok: false }); }));
    shell.body.appendChild(foot);
  }

  function editable(value, field, width) {
    var node = element('div', 'fld inp', value || '');
    node.contentEditable = 'true';
    node.tabIndex = 0;
    node.setAttribute('role', 'textbox');
    node.setAttribute('data-field', field || '');
    node.setAttribute('aria-multiline', 'false');
    node.style.flex = width ? 'none' : '1';
    if (width) { node.style.width = px(width); }
    node.addEventListener('input', function () {
      var clean = node.textContent.replace(/[\r\n]/g, '');
      if (clean !== node.textContent) { node.textContent = clean; placeCaretEnd(node); }
    });
    node.addEventListener('keydown', function (event) {
      if (event.key === 'Enter') { event.preventDefault(); }
    });
    return node;
  }

  function browseButton(field, kind, valueNode) {
    return modalButton('参照...', false, function () {
      post({ type: 'browse', token: currentToken, field: field, kind: kind, value: valueNode.textContent.trim() });
    });
  }

  function numberEditor(value, field, minimum, maximum) {
    var root = element('div', 'number-editor');
    var edit = editable(String(value), field);
    edit.classList.add('number-value');
    edit.setAttribute('role', 'spinbutton');
    edit.setAttribute('aria-valuemin', String(minimum));
    edit.setAttribute('aria-valuemax', String(maximum));
    edit.style.flex = '1';
    function read(fallback) {
      var parsed = Number(edit.textContent.trim());
      return Number.isInteger(parsed) ? parsed : fallback;
    }
    function write(next) {
      next = Math.max(minimum, Math.min(maximum, Number(next) || minimum));
      edit.textContent = String(next);
      edit.setAttribute('aria-valuenow', String(next));
      return next;
    }
    function step(delta) {
      write(read(minimum) + delta);
      edit.focus();
      placeCaretEnd(edit);
    }
    edit.addEventListener('input', function () {
      var clean = edit.textContent.replace(/[^0-9]/g, '');
      if (clean !== edit.textContent) { edit.textContent = clean; placeCaretEnd(edit); }
      var parsed = Number(clean);
      if (Number.isInteger(parsed) && parsed > maximum) { write(maximum); }
      else if (clean.length && Number.isInteger(parsed) && parsed < minimum) { write(minimum); }
      else if (Number.isInteger(parsed)) { edit.setAttribute('aria-valuenow', String(parsed)); }
    });
    edit.addEventListener('blur', function () { write(read(minimum)); });
    edit.addEventListener('keydown', function (event) {
      var delta = 0;
      if (event.key === 'ArrowUp') { delta = 1; }
      else if (event.key === 'ArrowDown') { delta = -1; }
      else if (event.key === 'PageUp') { delta = 10; }
      else if (event.key === 'PageDown') { delta = -10; }
      else if (event.key === 'Home') { event.preventDefault(); write(minimum); return; }
      else if (event.key === 'End') { event.preventDefault(); write(maximum); return; }
      if (delta) { event.preventDefault(); step(delta); }
    });
    var steps = element('div', 'number-steps');
    [['▲', 1, '1 増やす'], ['▼', -1, '1 減らす']].forEach(function (definition) {
      var control = element('div', 'number-step', definition[0]);
      control.setAttribute('role', 'button');
      control.setAttribute('aria-label', definition[2]);
      control.tabIndex = -1;
      control.addEventListener('mousedown', function (event) { event.preventDefault(); });
      control.addEventListener('click', function () { step(definition[1]); });
      steps.appendChild(control);
    });
    root.appendChild(edit);
    root.appendChild(steps);
    root.normalizeValue = function () { return write(read(minimum)); };
    write(value);
    return root;
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
    var candidateEditor = numberEditor(content.candidateRows, 'candidateRows', 1, 1000);
    countRow.appendChild(candidateEditor);
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
    foot.appendChild(modalButton('OK', true, function () {
      candidateEditor.normalizeValue();
      submitSettings(shell.body, error);
    }));
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
    pendingSettings = { ok: true, dataDir: dataDir, ledger: ledger, log: log,
      pattern: pattern, candidateRows: Number(value('candidateRows')) };
    error.hidden = true;
    post({ type: 'settingsSubmit', token: currentToken, dataDir: dataDir,
      ledger: ledger, log: log, pattern: pattern,
      candidateRows: pendingSettings.candidateRows });
  }

  function settingsValidation(message) {
    if (!pendingSettings || Number(message.token) !== currentToken ||
        !currentModal || currentModal.id !== 'v-set') { return; }
    if (message.ok) {
      var result = pendingSettings;
      pendingSettings = null;
      finishModal(result);
      return;
    }
    var error = currentModal.querySelector('.setting-error');
    if (error) { error.textContent = message.error || ''; error.hidden = false; }
    var target = currentModal.querySelector('[data-field="' + cssEscape(message.field || '') + '"]');
    if (target) { target.focus(); }
    pendingSettings = null;
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

  function padNumber(value, count) {
    var text = String(Math.abs(value));
    while (text.length < count) { text = '0' + text; }
    return text;
  }

  function formatDate(value, format) {
    format = format || 'yyyyMMdd';
    var months = ['January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December'];
    var days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    var result = '';
    for (var index = 0; index < format.length;) {
      var character = format.charAt(index);
      if (character === "'" || character === '"') {
        var quote = character;
        index++;
        while (index < format.length && format.charAt(index) !== quote) {
          if (format.charAt(index) === '\\' && index + 1 < format.length) { index++; }
          result += format.charAt(index++);
        }
        if (index < format.length) { index++; }
        continue;
      }
      if (character === '\\' && index + 1 < format.length) {
        result += format.charAt(index + 1); index += 2; continue;
      }
      if (character === '%' && index + 1 < format.length) { index++; character = format.charAt(index); }
      var end = index + 1;
      while (end < format.length && format.charAt(end) === character) { end++; }
      var count = end - index;
      var year = value.getFullYear();
      if (character === 'y') {
        result += count <= 2 ? padNumber(year % 100, count) : padNumber(year, count);
      } else if (character === 'M') {
        result += count === 1 ? String(value.getMonth() + 1) : count === 2 ?
          padNumber(value.getMonth() + 1, 2) : count === 3 ?
            months[value.getMonth()].slice(0, 3) : months[value.getMonth()];
      } else if (character === 'd') {
        result += count === 1 ? String(value.getDate()) : count === 2 ?
          padNumber(value.getDate(), 2) : count === 3 ?
            days[value.getDay()].slice(0, 3) : days[value.getDay()];
      } else if (character === 'H') {
        result += count === 1 ? String(value.getHours()) : padNumber(value.getHours(), 2);
      } else if (character === 'h') {
        var hour = value.getHours() % 12 || 12;
        result += count === 1 ? String(hour) : padNumber(hour, 2);
      } else if (character === 'm') {
        result += count === 1 ? String(value.getMinutes()) : padNumber(value.getMinutes(), 2);
      } else if (character === 's') {
        result += count === 1 ? String(value.getSeconds()) : padNumber(value.getSeconds(), 2);
      } else if (character === 'f' || character === 'F') {
        var fraction = padNumber(value.getMilliseconds(), 3);
        while (fraction.length < count) { fraction += '0'; }
        fraction = fraction.slice(0, count);
        result += character === 'F' ? fraction.replace(/0+$/, '') : fraction;
      } else if (character === 't') {
        var designator = value.getHours() < 12 ? 'AM' : 'PM';
        result += count === 1 ? designator.charAt(0) : designator;
      } else { result += format.slice(index, end); }
      index = end;
    }
    return result;
  }

  function closeCalendar(restoreFocus) {
    if (!activeCalendar) { return; }
    var root = activeCalendar;
    activeCalendar = null;
    root.classList.remove('open');
    root.querySelector('.date-value').setAttribute('aria-expanded', 'false');
    if (restoreFocus) { root.querySelector('.date-value').focus(); }
  }

  function dateEditor(format, field) {
    var root = element('div', 'dt date-editor');
    var display = element('div', 'fld inp date-value');
    display.tabIndex = 0;
    display.setAttribute('role', 'button');
    display.setAttribute('aria-haspopup', 'dialog');
    display.setAttribute('aria-expanded', 'false');
    display.setAttribute('aria-label', '日付を選ぶ');
    display.setAttribute('data-field', field || '');
    var chosen = new Date();
    var shownMonth = new Date(chosen.getFullYear(), chosen.getMonth(), 1);
    var popup = element('div', 'calendar-popup');
    popup.setAttribute('role', 'dialog');
    popup.setAttribute('aria-label', 'カレンダー');
    root.appendChild(display);
    root.appendChild(popup);
    function setChosen(next) {
      chosen.setFullYear(next.getFullYear(), next.getMonth(), next.getDate());
      shownMonth = new Date(chosen.getFullYear(), chosen.getMonth(), 1);
      display.textContent = formatDate(chosen, format);
    }
    function changeDay(amount) {
      var next = new Date(chosen.getTime());
      next.setDate(next.getDate() + amount);
      setChosen(next);
      if (root.classList.contains('open')) { drawCalendar(true); }
    }
    function changeMonth(amount) {
      shownMonth = new Date(shownMonth.getFullYear(), shownMonth.getMonth() + amount, 1);
      drawCalendar(false);
    }
    function drawCalendar(focusChosen) {
      popup.textContent = '';
      var head = element('div', 'calendar-head');
      var previous = modalButton('◀', false, function () { changeMonth(-1); });
      previous.classList.add('calendar-nav');
      var title = element('div', 'calendar-title', shownMonth.getFullYear() + '年' +
        String(shownMonth.getMonth() + 1) + '月');
      var next = modalButton('▶', false, function () { changeMonth(1); });
      next.classList.add('calendar-nav');
      head.appendChild(previous); head.appendChild(title); head.appendChild(next);
      popup.appendChild(head);
      var grid = element('div', 'calendar-grid');
      ['日', '月', '火', '水', '木', '金', '土'].forEach(function (name) {
        grid.appendChild(element('div', 'calendar-week', name));
      });
      var start = new Date(shownMonth.getFullYear(), shownMonth.getMonth(), 1 - shownMonth.getDay());
      for (var offset = 0; offset < 42; offset++) {
        (function () {
          var date = new Date(start.getFullYear(), start.getMonth(), start.getDate() + offset);
          var same = date.getFullYear() === chosen.getFullYear() &&
            date.getMonth() === chosen.getMonth() && date.getDate() === chosen.getDate();
          var day = element('div', 'calendar-day' +
            (date.getMonth() === shownMonth.getMonth() ? '' : ' outside') +
            (same ? ' selected' : ''), String(date.getDate()));
          day.setAttribute('role', 'button');
          day.setAttribute('aria-label', formatDate(date, 'yyyy-MM-dd'));
          day.setAttribute('aria-pressed', same ? 'true' : 'false');
          day.tabIndex = same ? 0 : -1;
          activate(day, function () { setChosen(date); closeCalendar(true); });
          day.addEventListener('keydown', function (event) {
            var delta = 0;
            if (event.key === 'ArrowLeft') { delta = -1; }
            else if (event.key === 'ArrowRight') { delta = 1; }
            else if (event.key === 'ArrowUp') { delta = -7; }
            else if (event.key === 'ArrowDown') { delta = 7; }
            else if (event.key === 'PageUp') { event.preventDefault(); changeMonth(-1); return; }
            else if (event.key === 'PageDown') { event.preventDefault(); changeMonth(1); return; }
            else { return; }
            event.preventDefault();
            changeDay(delta);
          });
          grid.appendChild(day);
        })();
      }
      popup.appendChild(grid);
      if (focusChosen) {
        window.setTimeout(function () {
          var selected = popup.querySelector('.calendar-day.selected');
          if (selected) { selected.focus(); }
        }, 0);
      }
    }
    function open() {
      if (activeCalendar && activeCalendar !== root) { closeCalendar(false); }
      shownMonth = new Date(chosen.getFullYear(), chosen.getMonth(), 1);
      drawCalendar(false);
      root.classList.add('open');
      display.setAttribute('aria-expanded', 'true');
      activeCalendar = root;
      var selected = popup.querySelector('.calendar-day.selected');
      if (selected) { selected.focus(); }
    }
    display.addEventListener('click', open);
    display.addEventListener('keydown', function (event) {
      if (event.key === 'Enter' || event.key === ' ' || (event.altKey && event.key === 'ArrowDown')) {
        event.preventDefault(); open();
      } else if (event.key === 'ArrowUp' || event.key === 'ArrowDown') {
        event.preventDefault(); changeDay(event.key === 'ArrowUp' ? 1 : -1);
      }
    });
    setChosen(chosen);
    return root;
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
    var leftSelected = null;
    var rightSelected = null;
    var picker = element('div');
    picker.style.display = 'flex';
    picker.style.gap = 'var(--card-gap)';
    picker.style.alignItems = 'stretch';
    var left = element('div'); left.style.flex = '1 1 0'; left.style.minWidth = '0';
    left.appendChild(element('div', 'lab', '出力できる項目'));
    var leftList = element('div', 'lb'); leftList.style.height = '180px';
    leftList.setAttribute('role', 'listbox'); leftList.setAttribute('aria-label', '出力できる項目');
    left.appendChild(leftList);
    var mover = element('div'); mover.style.display = 'flex'; mover.style.flexDirection = 'column';
    mover.style.justifyContent = 'center'; mover.style.gap = 'var(--card-gap)';
    var right = element('div'); right.style.flex = '1 1 0'; right.style.minWidth = '0';
    var rightHead = element('div'); rightHead.style.display = 'flex'; rightHead.style.alignItems = 'center';
    var rightTitle = element('span', 'lab'); rightTitle.style.flex = '1'; rightHead.appendChild(rightTitle);
    var reset = modalButton('既定に戻す', false, function () {
      selected = content.defaults.filter(function (reference) { return !!byRef[reference]; });
      available = content.fields.map(function (field) { return field.ref; })
        .filter(function (reference) { return selected.indexOf(reference) < 0; });
      leftSelected = null;
      rightSelected = null;
      drawLists();
    });
    reset.classList.add('sm'); rightHead.appendChild(reset); right.appendChild(rightHead);
    var rightList = element('div', 'lb'); rightList.style.height = '180px';
    rightList.setAttribute('role', 'listbox'); rightList.setAttribute('aria-label', '出力する項目');
    right.appendChild(rightList);
    function selectedReference(side) { return side === 'left' ? leftSelected : rightSelected; }
    function setSelectedReference(side, reference) {
      if (side === 'left') { leftSelected = reference; }
      else { rightSelected = reference; }
    }
    function selectListItem(list, side, reference, focus) {
      setSelectedReference(side, reference);
      Array.prototype.forEach.call(list.children, function (other) {
        var chosen = other.getAttribute('data-ref') === reference;
        other.setAttribute('aria-selected', chosen ? 'true' : 'false');
        other.tabIndex = chosen ? 0 : -1;
        if (chosen && focus) { other.focus(); other.scrollIntoView({ block: 'nearest' }); }
      });
    }
    function listItem(reference, list, source, destination, side) {
      var item = element('div', '', byRef[reference].label);
      var chosen = selectedReference(side) === reference;
      item.tabIndex = chosen ? 0 : -1;
      item.setAttribute('role', 'option'); item.setAttribute('data-ref', reference);
      item.setAttribute('aria-selected', chosen ? 'true' : 'false');
      activate(item, function () {
        selectListItem(list, side, reference, false);
      });
      item.addEventListener('keydown', function (event) {
        var index = source.indexOf(reference);
        var next = index;
        if (event.key === 'ArrowUp') { next = index - 1; }
        else if (event.key === 'ArrowDown') { next = index + 1; }
        else if (event.key === 'Home') { next = 0; }
        else if (event.key === 'End') { next = source.length - 1; }
        else { return; }
        event.preventDefault();
        next = Math.max(0, Math.min(source.length - 1, next));
        selectListItem(list, side, source[next], true);
      });
      item.addEventListener('dblclick', function () { moveItem(list, source, destination, side); });
      return item;
    }
    function prepareList(list, source, side) {
      list.tabIndex = selectedReference(side) ? -1 : 0;
      list.onkeydown = function (event) {
        if (event.target !== list || !source.length ||
            ['ArrowUp', 'ArrowDown', 'Home', 'End'].indexOf(event.key) < 0) { return; }
        event.preventDefault();
        var index = event.key === 'End' || event.key === 'ArrowUp' ? source.length - 1 : 0;
        selectListItem(list, side, source[index], true);
        list.tabIndex = -1;
      };
    }
    function drawLists() {
      leftList.textContent = ''; rightList.textContent = '';
      available.forEach(function (reference) {
        leftList.appendChild(listItem(reference, leftList, available, selected, 'left'));
      });
      selected.forEach(function (reference) {
        rightList.appendChild(listItem(reference, rightList, selected, available, 'right'));
      });
      prepareList(leftList, available, 'left');
      prepareList(rightList, selected, 'right');
      rightTitle.textContent = '出力する項目（' + selected.length + '）';
    }
    function moveItem(list, source, destination, side) {
      var reference = selectedReference(side);
      if (!reference) { return; }
      var index = source.indexOf(reference);
      if (index >= 0) {
        source.splice(index, 1); destination.push(reference);
        setSelectedReference(side, source.length ? source[Math.min(index, source.length - 1)] : null);
        drawLists();
      }
    }
    var moveRight = modalButton('▶', false, function () { moveItem(leftList, available, selected, 'left'); });
    moveRight.classList.add('sm'); moveRight.style.width = '34px';
    var moveLeft = modalButton('◀', false, function () { moveItem(rightList, selected, available, 'right'); });
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
    var firstHost = element('div', 'filter-value-host');
    var mark = element('div', 'tilde', '～');
    var lastHost = element('div', 'filter-value-host');
    var add = modalButton('追加', false, addFilter); add.classList.add('sm');
    grid.appendChild(fieldSelect); grid.appendChild(operatorSelect); grid.appendChild(firstHost);
    grid.appendChild(mark); grid.appendChild(lastHost); grid.appendChild(add);
    var listHost = element('div', 'row3');
    var filterTable = tableNode([{ header: '項目', width: 180 }, { header: '条件', width: 110 }, { header: '値' }], [], { readOnly: false });
    filterTable.classList.add('f3'); listHost.appendChild(filterTable); grid.appendChild(listHost);
    var remove = modalButton('削除', false, removeFilter); remove.classList.add('sm', 'row3b'); grid.appendChild(remove);
    filterSet.appendChild(grid); shell.body.appendChild(filterSet);
    fieldSelect.addEventListener('change', updateOperators);
    var operatorLabels = {
      contains: '含む', equals: '等しい', startsWith: '始まる',
      notContains: '含まない', range: '範囲'
    };
    function editorValue(host, name) {
      var node = host.querySelector('[data-field="' + name + '"]');
      return node ? node.textContent.trim() : '';
    }
    function updateOperators() {
      var field = byRef[fieldSelect.value];
      closeCalendar(false);
      operatorSelect.textContent = '';
      var entries = field.kind === 'text' ? ['contains', 'equals', 'startsWith', 'notContains'] : ['range'];
      entries.forEach(function (code) {
        var option = element('option', '', operatorLabels[code]); option.value = code;
        operatorSelect.appendChild(option);
      });
      firstHost.textContent = '';
      lastHost.textContent = '';
      if (field.kind === 'date') {
        firstHost.appendChild(dateEditor(field.format, 'filterFirst'));
        lastHost.appendChild(dateEditor(field.format, 'filterLast'));
      } else {
        firstHost.appendChild(editable('', 'filterFirst'));
        if (field.kind !== 'text') { lastHost.appendChild(editable('', 'filterLast')); }
      }
      mark.classList.toggle('off', field.kind === 'text');
      lastHost.classList.toggle('off', field.kind === 'text');
    }
    function addFilter() {
      var field = byRef[fieldSelect.value];
      var firstValue = editorValue(firstHost, 'filterFirst');
      var lastValue = editorValue(lastHost, 'filterLast');
      pendingFilter = { field: field.ref, operator: operatorSelect.value, first: firstValue,
        last: field.kind === 'text' ? '' : lastValue };
      post({ type: 'validateExportFilter', token: currentToken, field: field.ref,
        first: pendingFilter.first, last: pendingFilter.last });
    }
    function redrawFilters(selectedIndex, focus) {
      if (selectedIndex === undefined) { selectedIndex = filters.length - 1; }
      var replacement = tableNode([{ header: '項目', width: 180 }, { header: '条件', width: 110 }, { header: '値' }],
        filters.map(function (entry) {
          return [byRef[entry.field].label, operatorLabels[entry.operator] || entry.operator,
            entry.last ? entry.first + ' ～ ' + entry.last : entry.first];
        }), { readOnly: false, selected: selectedIndex });
      replacement.classList.add('f3'); listHost.replaceChild(replacement, listHost.firstChild); filterTable = replacement;
      if (focus) {
        var row = filterTable.querySelector('tbody tr[aria-selected=true]');
        if (row) { row.focus(); row.scrollIntoView({ block: 'nearest' }); }
      }
    }
    function removeFilter() {
      var row = filterTable.querySelector('tbody tr[aria-selected=true]');
      var index = row ? Number(row.getAttribute('data-index')) : filters.length === 1 ? 0 : -1;
      if (index >= 0) {
        filters.splice(index, 1);
        redrawFilters(filters.length ? Math.min(index, filters.length - 1) : -1, filters.length > 0);
      }
    }
    function exportFilterValidation(message) {
      if (!pendingFilter || Number(message.token) !== currentToken ||
          !currentModal || currentModal.id !== 'v-out') { return; }
      if (!message.ok) {
        error.textContent = message.error || '';
        error.hidden = false;
        var target = firstHost.querySelector('[data-field]');
        if (target) { target.focus(); }
        pendingFilter = null;
        return;
      }
      pendingFilter.first = message.first === undefined ? pendingFilter.first : message.first;
      pendingFilter.last = message.last === undefined ? pendingFilter.last : message.last;
      filters.push(pendingFilter);
      pendingFilter = null;
      error.hidden = true;
      redrawFilters(filters.length - 1, true);
    }
    updateOperators();
    var pathRow = element('div', 'kv export-path'); pathRow.style.marginTop = '9px';
    pathRow.appendChild(element('label', '', '出力先'));
    var destination = editable(content.destination, 'exportPath'); pathRow.appendChild(destination);
    pathRow.appendChild(browseButton('exportPath', 'export', destination)); shell.body.appendChild(pathRow);
    var error = element('div', 'setting-error'); error.hidden = true; error.setAttribute('role', 'alert'); shell.body.appendChild(error);
    var foot = element('div', 'foot');
    foot.appendChild(modalButton('OK', true, function () {
      if (!selected.length) {
        error.textContent = '出力する項目を 1 つ以上選んでください。'; error.hidden = false; return;
      }
      finishModal({ ok: true, path: destination.textContent.trim(), fields: selected, filters: filters });
    }));
    foot.appendChild(modalButton('キャンセル', false, function () { finishModal({ ok: false }); }));
    shell.body.appendChild(foot);
    shell.dialog.exportFilterValidation = exportFilterValidation;
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
    pendingSettings = null;
    pendingFilter = null;
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
      var focus = currentModal && currentModal.querySelector(
        '.body [data-autofocus=true],.body [contenteditable=true],' +
        '.body tbody tr[aria-selected=true],.body [data-modal-default=true],' +
        '.body select,.body [tabindex="0"]');
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
    else if (message.type === 'settingsValidation') { settingsValidation(message); }
    else if (message.type === 'exportFilterValidation' && currentModal && currentModal.querySelector('.dlg').exportFilterValidation) {
      currentModal.querySelector('.dlg').exportFilterValidation(message);
    }
    else if (message.type === 'pickerPreview') { pickerPreview(message.preview); }
    else if (message.type === 'pickerResult' && Number(message.token) === currentToken) { pickerResult(message.target); }
  }

  document.addEventListener('keydown', function (event) {
    if (event.key === 'Escape' && activeCalendar) {
      event.preventDefault(); closeCalendar(true); return;
    }
    if (event.key === 'Escape' && currentModal && currentModal.id !== 'v-pick') {
      event.preventDefault(); finishModal({ ok: false }); return;
    }
    if (event.key === 'Enter' && currentModal && !event.altKey && !event.ctrlKey && !event.metaKey) {
      var fromText = !!(event.target && event.target.closest &&
        event.target.closest('[contenteditable=true]'));
      if (!event.defaultPrevented || fromText) {
        var primary = currentModal.querySelector('.body [data-modal-default=true][aria-disabled=false]');
        if (primary) { event.preventDefault(); primary.click(); return; }
      }
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
    if (event.key === 'Enter' && !currentModal && !event.defaultPrevented && input &&
        input.getAttribute('aria-disabled') !== 'true') {
      event.preventDefault();
      post({ type: 'action', name: 'search', job: '', key: keyValue() });
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

  document.addEventListener('mousedown', function (event) {
    if (activeCalendar && !activeCalendar.contains(event.target)) { closeCalendar(false); }
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

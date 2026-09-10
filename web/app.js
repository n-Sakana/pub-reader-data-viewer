(function () {
  'use strict';

  function logClientError(kind, message, file, line, column, stack) {
    try { post({ type: 'clientError', kind: kind, message: String(message || ''),
      file: file || '', line: line || 0, column: column || 0, stack: String(stack || '') }); }
    catch (ignored) { /* The native process failure event covers a lost bridge. */ }
  }
  window.addEventListener('error', function (event) {
    logClientError('error', event.message, event.filename, event.lineno, event.colno, event.error && event.error.stack);
  });
  window.addEventListener('unhandledrejection', function (event) {
    logClientError('promise', event.reason, '', 0, 0, event.reason && event.reason.stack);
  });

  var stage = document.querySelector('.stage');
  var input = null;
  var statusSegments = [];
  var currentModal = null;
  var currentToken = 0;
  var modalReturnFocus = null;
  var settingsContent = null;
  var pendingSettings = null;
  var activeCalendar = null;
  var mainFocusDone = false;

  var dialogMode = window.location.hash === '#dialog';
  var dialogSizePending = false;
  if (dialogMode) { document.documentElement.classList.add('dialogmode'); }

  // The dialog window opens off screen at a generous size; report what the
  // dialog measured so the host can fit the window to it and centre it.
  // A dialog can be replaced in place (設定 -> 画面から選ぶ) or grow when a
  // message appears, so watch the size instead of reporting once on open.
  function reportDialogSize() {
    var veil = document.querySelector('.veil.show');
    var dialog = veil ? veil.querySelector('.dlg') : null;
    if (!dialog) { return; }
    var box = dialog.getBoundingClientRect();
    var width = Math.ceil(box.width);
    var height = Math.ceil(box.height);
    if (width < 1 || height < 1) { return; }
    // Opening hides and moves the native window, even when its size still
    // fits. Always acknowledge that opening before suppressing resize loops.
    if (!dialogSizePending && Math.abs(width - window.innerWidth) <= 1 &&
        Math.abs(height - window.innerHeight) <= 1) { return; }
    dialogSizePending = false;
    var title = dialog.querySelector('.tb .ttl');
    post({
      type: 'dialogSize',
      token: currentToken,
      width: width,
      height: height,
      title: title ? title.textContent : ''
    });
  }

  function watchDialogSize() {
    if (!dialogMode) { return; }
    var stage = document.querySelector('.stage');
    if (stage && window.ResizeObserver) {
      new window.ResizeObserver(function () { reportDialogSize(); }).observe(stage);
    }
    // If the window ends up a size the dialog did not ask for -- a late report
    // from a dialog that has already been replaced -- say so again. The dialog
    // is a fixed width, so this settles after one correction.
    window.addEventListener('resize', function () {
      window.clearTimeout(reportDialogSize.pending);
      reportDialogSize.pending = window.setTimeout(reportDialogSize, 60);
    });
  }
  if (dialogMode) { document.addEventListener('DOMContentLoaded', watchDialogSize); watchDialogSize(); }

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

  function fieldset(title) {
    var set = element('fieldset', 'gb');
    set.appendChild(element('legend', '', title || ''));
    return set;
  }

  function onKeyInput() {
    var max = Number(input.getAttribute('data-max-length')) || 64;
    var value = input.textContent.replace(/[\r\n\t]/g, '');
    if (value.length > max) { value = value.slice(0, max); }
    // Keep the submitted value and the visible DOM identical, even after paste.
    if (input.textContent !== value || input.children.length) { input.textContent = value; placeCaretEnd(input); }
    post({ type: 'key', value: value });
  }

  function keyValue() { return input ? input.textContent.replace(/[\r\n\t]/g, '').trim() : ''; }

  function placeCaretEnd(node) {
    var range = document.createRange();
    range.selectNodeContents(node);
    range.collapse(false);
    var selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
  }

  var fixedScreenBound = false;
  function renderScreen(definition) {
    if (!definition || definition.fixed !== true) { throw new Error('The fixed HTML screen requires compact settings.'); }
    stage.classList.add('runtime');
    Array.prototype.forEach.call(stage.querySelectorAll('.win [data-action]'), function (node) {
      var action = node.getAttribute('data-action');
      node.setAttribute('data-job', definition.actions[action] || '');
    });
    if (fixedScreenBound) { return; }
    fixedScreenBound = true;
    input = stage.querySelector('#input');
    statusSegments = Array.prototype.slice.call(stage.querySelectorAll('.win .sb .sp'));
    Array.prototype.forEach.call(stage.querySelectorAll('.win [data-action]'), function (node) {
      activate(node, function () {
        post({ type: 'action', name: node.getAttribute('data-action'), job: node.getAttribute('data-job') || '', key: keyValue() });
      });
    });
    input.addEventListener('input', function (event) { if (!event.isComposing) { onKeyInput(); } });
    input.addEventListener('compositionend', onKeyInput);
    input.addEventListener('paste', function (event) {
      event.preventDefault();
      var text = (event.clipboardData || window.clipboardData).getData('text/plain').replace(/[\r\n\t]/g, '');
      document.execCommand('insertText', false, text);
      onKeyInput();
    });
    input.addEventListener('keydown', function (event) {
      if (event.key === 'Enter') {
        event.preventDefault();
        if (input.getAttribute('aria-disabled') !== 'true') {
          post({ type: 'action', name: 'search', job: '', key: keyValue() });
        }
      }
    });
  }

  function updateColumns() { /* The fixed layout responds through CSS. */ }

  function fixedCandidatePresentation() {
    var table = document.querySelector('#fixed-candidate-columns').content.querySelector('table');
    return {
      title: table.dataset.title, hint: table.dataset.hint,
      width: Number(table.dataset.width), maxHeight: Number(table.dataset.maxHeight),
      rowHeight: Number(table.dataset.rowHeight), headerHeight: Number(table.dataset.headerHeight),
      columns: Array.prototype.map.call(table.querySelectorAll('th'), function (th) {
        return { header: th.textContent, width: Number(th.dataset.width) || 0,
          align: th.dataset.align || 'left', muted: th.dataset.muted === 'true', render: th.dataset.render || 'text' };
      })
    };
  }

  function applyState(next) {
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
    if (dialogMode) {
      dialogSizePending = true;
      requestAnimationFrame(function () { requestAnimationFrame(reportDialogSize); });
    }
    return { veil: veil, dialog: dialog, body: body };
  }

  function finishModal(result) {
    if (!currentToken) { return; }
    var token = currentToken;
    pendingSettings = null;
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
    var presentation = fixedCandidatePresentation();
    if (!shared) { content.title = presentation.title; content.hint = presentation.hint; }
    content.width = presentation.width;
    content.maxHeight = presentation.maxHeight;
    content.rowHeight = presentation.rowHeight;
    content.headerHeight = presentation.headerHeight;
    content.columns = presentation.columns;
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
    foot.appendChild(modalButton(content.discardText || '一覧の未送信変更を破棄', false,
      function () { finishModal({ ok: true, discard: true }); }));
    foot.appendChild(modalButton('OK', true, function () { finishModal({ ok: true, discard: false }); }));
    shell.body.appendChild(foot);
  }

  function installFixedBody(shell, name, content) {
    shell.body.appendChild(document.querySelector('#fixed-' + name).content.cloneNode(true));
    Array.prototype.forEach.call(shell.body.querySelectorAll('[data-text]'), function (node) {
      node.textContent = content[node.dataset.text] || '';
    });
    activate(shell.body.querySelector('[data-command=cancel]'), function () { finishModal({ ok: false }); });
  }

  function openArchive(content) {
    var shell = modalShell('v-archive', content.title);
    shell.body.appendChild(element('div', 'hint', content.hint));
    var selected = {};
    (content.selected || []).forEach(function (index) { selected[index] = true; });
    function indices() { return Object.keys(selected).map(Number); }
    var summary = element('div', 'hint');
    shell.body.appendChild(summary);
    var list = element('div', 'lv archive-list');
    var table = element('table'), thead = element('thead'), head = element('tr');
    ['復元', '識別キー', '確認状態', '削除日時'].forEach(function (title) { head.appendChild(element('th', '', title)); });
    thead.appendChild(head); table.appendChild(thead);
    var body = element('tbody'); table.appendChild(body); list.appendChild(table); shell.body.appendChild(list);
    var details = fieldset('削除時の全内容'); details.classList.add('archive-details');
    var detailTable = element('table'); details.appendChild(detailTable); shell.body.appendChild(details);
    function detail(row) {
      detailTable.textContent = '';
      content.labels.forEach(function (label, index) {
        var tr = element('tr'); tr.appendChild(element('th', '', label)); tr.appendChild(element('td', '', row.values[index])); detailTable.appendChild(tr);
      });
    }
    var restoreButton;
    function update() {
      var count = indices().length;
      summary.textContent = '保管 ' + content.total + ' 件 / 選択 ' + count + ' 件 / ' + (content.page + 1) + ' ページ';
      if (restoreButton) { setDisabled(restoreButton, count === 0); }
    }
    content.rows.forEach(function (row) {
      var tr = element('tr'); tr.setAttribute('data-archive-index', row.index);
      var choice = element('input'); choice.type = 'checkbox'; choice.checked = !!selected[row.index];
      choice.setAttribute('aria-label', row.identity + ' を復元対象にする');
      choice.onchange = function () { if (choice.checked) { selected[row.index] = true; } else { delete selected[row.index]; } detail(row); update(); };
      var td = element('td'); td.appendChild(choice); tr.appendChild(td);
      var identity = element('td');
      var show = modalButton(row.identity, false, function () { detail(row); }); show.classList.add('archive-detail-button'); identity.appendChild(show); tr.appendChild(identity);
      tr.appendChild(element('td', '', row.state));
      tr.appendChild(element('td', '', new Date(row.deletedAt).toLocaleString('ja-JP')));
      body.appendChild(tr);
    });
    if (content.rows.length) { detail(content.rows[0]); }
    else { details.appendChild(element('div', 'hint', '保管中のレコードはありません。')); }
    var foot = element('div', 'foot');
    foot.appendChild(modalButton('このページを選択', false, function () {
      content.rows.forEach(function (row) { selected[row.index] = true; });
      Array.prototype.forEach.call(body.querySelectorAll('input[type=checkbox]'), function (box) { box.checked = true; });
      update();
    }, !content.rows.length));
    foot.appendChild(modalButton('前へ', false, function () { finishModal({ page: content.page - 1, selected: indices() }); }, content.page === 0));
    foot.appendChild(modalButton('次へ', false, function () { finishModal({ page: content.page + 1, selected: indices() }); }, (content.page + 1) * content.pageSize >= content.total));
    restoreButton = modalButton('選択したレコードを復元', true, function () { finishModal({ ok: true, selected: indices() }); });
    foot.appendChild(restoreButton);
    foot.appendChild(modalButton('閉じる', false, function () { finishModal({ ok: false }); }));
    shell.body.appendChild(foot); update();
  }

  function openProcess(content, deleting) {
    var shell = modalShell(deleting ? 'v-del' : 'v-upd', content.title);
    installFixedBody(shell, 'process', content);
    shell.body.querySelector('[data-slot=inputs]').appendChild(tableNode([
      { header: deleting ? '指定' : '表', width: 40 }, { header: 'ファイル', width: 150 },
      { header: 'キー', width: 62 }, { header: '行数', width: 74, align: 'right' }, { header: '検証' }
    ], content.inputs.map(function (entry) {
      return [entry.id, entry.file, entry.key, entry.rows, { text: entry.validation, tone: entry.valid ? 0 : 2 }];
    }), { readOnly: true }));
    shell.body.querySelector('[data-slot=steps]').appendChild(tableNode([
      { header: '#', width: 26, align: 'right' }, { header: '操作', width: 46 },
      { header: '対象1', width: 64 }, { header: '対象2', width: 76 }, { header: 'キー', width: 120 },
      { header: '条件', width: 240 }, { header: '出力' }
    ], content.steps, { readOnly: true, maxHeight: 128 }));
    Array.prototype.forEach.call(shell.body.querySelectorAll('[data-output]'), function (node) {
      node.textContent = content.output[Number(node.dataset.output)] || '';
    });
    var execute = shell.body.querySelector('[data-command=execute]');
    setMnemonic(execute, content.executeText);
    setDisabled(execute, !content.canRun);
    activate(execute, function () { finishModal({ ok: true }); });
    if (!content.canRun && content.cannotRunText) {
      execute.title = content.cannotRunText;
      execute.setAttribute('aria-label', content.executeText + ': ' + content.cannotRunText);
    }
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
    installFixedBody(shell, 'settings', content);
    [['dataDir','folder'], ['ledger','ledger'], ['log','log']].forEach(function (entry) {
      var row = shell.body.querySelector('[data-row="' + entry[0] + '"]');
      var edit = editable(content[entry[0]], entry[0]);
      row.appendChild(edit); row.appendChild(browseButton(entry[0], entry[1], edit));
    });
    shell.body.querySelector('[data-row=pattern]').appendChild(editable(content.pattern, 'pattern'));
    var candidateEditor = numberEditor(content.candidateRows, 'candidateRows', 1, 1000);
    shell.body.querySelector('[data-row=candidateRows]').appendChild(candidateEditor);
    shell.body.querySelector('[data-target-summary]').textContent = content.target.summary || '';
    shell.body.querySelector('[data-target-read]').textContent = content.target.read || '';
    activate(shell.body.querySelector('#b-pick'), startPicker);
    var error = shell.body.querySelector('.setting-error');
    activate(shell.body.querySelector('[data-command=execute]'), function () {
      candidateEditor.normalizeValue(); submitSettings(shell.body, error);
    });
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
    chosen.setHours(0, 0, 0, 0);
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

  function exportFieldPicker(content) {
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
    var leftHead = element('div', 'listhead');
    leftHead.appendChild(element('span', 'lab', '出力できる項目'));
    left.appendChild(leftHead);
    var leftList = element('div', 'lb'); leftList.style.height = '180px';
    leftList.setAttribute('role', 'listbox'); leftList.setAttribute('aria-label', '出力できる項目');
    left.appendChild(leftList);
    var mover = element('div'); mover.style.display = 'flex'; mover.style.flexDirection = 'column';
    mover.style.justifyContent = 'center'; mover.style.gap = 'var(--card-gap)';
    var right = element('div'); right.style.flex = '1 1 0'; right.style.minWidth = '0';
    var rightHead = element('div', 'listhead');
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
    drawLists();
    return { node: picker, values: function () { return selected.slice(); } };
  }


  function exportFilterEditor(content, token, isActive, showError) {
    var byRef = {};
    content.fields.forEach(function (field) { byRef[field.ref] = field; });
    var pendingFilter = null;
    var filters = [];
    var filterSet = document.querySelector('#fixed-filter').content.firstElementChild.cloneNode(true);
    var grid = filterSet.querySelector('.fgrid');
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
    filterSet.appendChild(grid);
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
      post({ type: 'validateExportFilter', token: token, field: field.ref,
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
      if (!pendingFilter || Number(message.token) !== token || !isActive()) { return; }
      if (!message.ok) {
        showError(message.error || '');
        var target = firstHost.querySelector('[data-field]');
        if (target) { target.focus(); }
        pendingFilter = null;
        return;
      }
      pendingFilter.first = message.first === undefined ? pendingFilter.first : message.first;
      pendingFilter.last = message.last === undefined ? pendingFilter.last : message.last;
      filters.push(pendingFilter);
      pendingFilter = null;
      showError(null);
      redrawFilters(filters.length - 1, true);
    }
    updateOperators();
    return { node: filterSet, values: function () { return filters.slice(); }, validate: exportFilterValidation };
  }

  function exportDestination(content) {
    var pathRow = element('div', 'kv export-path'); pathRow.style.marginTop = '9px';
    pathRow.appendChild(element('label', '', '出力先'));
    var safeRow = element('label', 'export-safe');
    var excelSafe = element('input'); excelSafe.type = 'checkbox'; excelSafe.checked = true;
    excelSafe.id = 'export-excel-safe'; safeRow.appendChild(excelSafe);
    safeRow.appendChild(document.createTextNode(content.excelSafeText || 'Excel向けに数式を無効化'));
    var destination = editable(content.destination, 'exportPath'); pathRow.appendChild(destination);
    pathRow.appendChild(browseButton('exportPath', 'export', destination));
    return { nodes: [safeRow, pathRow], path: function () { return destination.textContent.trim(); },
      excelSafe: function () { return excelSafe.checked; } };
  }

  function openExport(content) {
    var shell = modalShell('v-out', content.title);
    installFixedBody(shell, 'export', content);
    var token = currentToken;
    var error = shell.body.querySelector('.setting-error');
    function showError(message) { error.textContent = message || ''; error.hidden = message === null; }
    function isActive() { return currentToken === token && currentModal === shell.veil; }

    var fields = exportFieldPicker(content);
    var filters = exportFilterEditor(content, token, isActive, showError);
    var destination = exportDestination(content);
    shell.body.querySelector('[data-slot=fields]').appendChild(fields.node);
    shell.body.querySelector('[data-slot=filters]').appendChild(filters.node);
    destination.nodes.forEach(function (node) { shell.body.querySelector('[data-slot=destination]').appendChild(node); });
    activate(shell.body.querySelector('[data-command=execute]'), function () {
      var selected = fields.values();
      if (!selected.length) { showError('出力する項目を 1 つ以上選んでください。'); return; }
      var path = destination.path();
      if (!/\.csv$/i.test(path)) { showError('出力先には .csv ファイルを指定してください。'); return; }
      finishModal({ ok: true, path: path, fields: selected, filters: filters.values(), excelSafe: destination.excelSafe() });
    });
    shell.dialog.exportFilterValidation = filters.validate;
  }

  function selectNode(entries) {
    var select = element('select', 'fld inp');
    select.style.width = '100%';
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
    var modal = message.modal;
    var content = message.content || {};
    if (modal === 'confirm') { openConfirm(content); }
    else if (modal === 'candidates') { openCandidates(content, false, false); }
    else if (modal === 'shared') { openCandidates(content, true, false); }
    else if (modal === 'sharedTell') { openCandidates(content, true, true); }
    else if (modal === 'unmatched') { openUnmatched(content); }
    else if (modal === 'archive') { openArchive(content); }
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
    if (event.isComposing || event.keyCode === 229) { event.stopImmediatePropagation(); }
  }, true);

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
      var items = Array.prototype.filter.call(currentModal.querySelectorAll('[contenteditable=true],[tabindex="0"],select,input:not([disabled])'), function (node) {
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

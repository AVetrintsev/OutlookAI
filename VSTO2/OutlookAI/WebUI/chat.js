/* ============================================================
   OutlookAI Chat - WebView2 surface controller (Phase 2)
   Public API (all on window.outlookai):
     appendUserMessage(text)
     appendAssistantMessage(messageId, initialText)
     appendTextDelta(messageId, delta)
     finalizeAssistantMessage(messageId, opts)
      appendToolCallCard(callId, name, argsJson)
      updateToolCallCard(callId, ok, summary, resultJson)
      onFileSaved(messageId, fileInfo)
      onExportError(messageId, error)
      appendAuditRow(text)
     showError(message)
     setComposerEnabled(enabled, isStopVisible)
     clear()
     applyTheme(themeName)              // "light" | "dark" | "high-contrast"
     setContextStrip({subject, recipients, thread})
   Host -> JS:
     C# calls these via WebView2.ExecuteScriptAsync("outlookai.X(...)").
   JS -> Host:
     window.chrome.webview.postMessage(JSON.stringify({type:..., payload:...}))
      Message types: 'send', 'custom_action', 'stop', 'clear', 'copy', 'toolCardClicked',
                     'open_file', 'open_item', 'reveal_in_explorer', 'export_pdf'
   ============================================================ */

(function() {
  'use strict';

  // -- DOM refs ----------------------------------------------------
  var $messages = document.getElementById('messages');
  var $input = document.getElementById('composerInput');
  var $btnSend = document.getElementById('btnSend');
  var $btnStop = document.getElementById('btnStop');
  var $btnClear = document.getElementById('btnClear');
  var $btnCopy = document.getElementById('btnCopy');
  var $composerResizer = document.getElementById('composerResizer');
  var $quickActions = document.getElementById('quickActions');
  var $actionWorkspace = document.getElementById('actionWorkspace');
  var $recommendationsSection = document.getElementById('recommendationsSection');
  var $recommendationsLoading = document.getElementById('recommendationsLoading');
  var $actionRecommendations = document.getElementById('actionRecommendations');
  var $actionGroups = document.getElementById('actionGroups');
  var $actionGroupPopover = document.getElementById('actionGroupPopover');
  var $actionGroupEditor = document.getElementById('actionGroupEditor');
  var $groupEditorTitle = document.getElementById('groupEditorTitle');
  var $groupEditorActions = document.getElementById('groupEditorActions');
  var $btnGroupEditorClose = document.getElementById('btnGroupEditorClose');
  var $btnGroupAddAction = document.getElementById('btnGroupAddAction');
  var $btnGroupImport = document.getElementById('btnGroupImport');
  var $btnGroupExport = document.getElementById('btnGroupExport');
  var $btnGroupReset = document.getElementById('btnGroupReset');
  var $groupImportFile = document.getElementById('groupImportFile');
  var $groupImportPanel = document.getElementById('groupImportPanel');
  var $groupImportText = document.getElementById('groupImportText');
  var $btnGroupImportFile = document.getElementById('btnGroupImportFile');
  var $btnGroupImportCancel = document.getElementById('btnGroupImportCancel');
  var $btnGroupImportApply = document.getElementById('btnGroupImportApply');
  var $customActionDialog = document.getElementById('customActionDialog');
  var $customActionName = document.getElementById('customActionName');
  var $customActionHint = document.getElementById('customActionHint');
  var $customActionPrompt = document.getElementById('customActionPrompt');
  var $customActionSource = document.getElementById('customActionSource');
  var $customActionFilter = document.getElementById('customActionFilter');
  var $customActionPeriod = document.getElementById('customActionPeriod');
  var $customActionMaxItems = document.getElementById('customActionMaxItems');
  var $customActionManualFrom = document.getElementById('customActionManualFrom');
  var $customActionManualTo = document.getElementById('customActionManualTo');
  var $customActionFilterField = document.getElementById('customActionFilterField');
  var $customActionPeriodField = document.getElementById('customActionPeriodField');
  var $customActionMaxItemsField = document.getElementById('customActionMaxItemsField');
  var $customActionManualFromField = document.getElementById('customActionManualFromField');
  var $customActionManualToField = document.getElementById('customActionManualToField');
  var $customActionFullBodies = document.getElementById('customActionFullBodies');
  var $customActionFullBodiesField = document.getElementById('customActionFullBodiesField');
  var $customActionAttachments = document.getElementById('customActionAttachments');
  var $customActionAttachmentsField = document.getElementById('customActionAttachmentsField');
  var $customActionOutput = document.getElementById('customActionOutput');
  var $customActionToolsField = document.getElementById('customActionToolsField');
  var $customActionTools = document.getElementById('customActionTools');
  var $customActionTitle = document.getElementById('customActionTitle');
  var $btnCustomActionClose = document.getElementById('btnCustomActionClose');
  var $btnCustomActionCancel = document.getElementById('btnCustomActionCancel');
  var $btnCustomActionSave = document.getElementById('btnCustomActionSave');
  var $customActionMenu = document.getElementById('customActionMenu');
  var $assistantMessageMenu = document.getElementById('assistantMessageMenu');
  var $btnAssistantExportPdf = document.getElementById('btnAssistantExportPdf');
  var $btnCustomActionEdit = document.getElementById('btnCustomActionEdit');
  var $btnCustomActionDelete = document.getElementById('btnCustomActionDelete');
  var $customActionDeleteDialog = document.getElementById('customActionDeleteDialog');
  var $customActionDeleteText = document.getElementById('customActionDeleteText');
  var $btnCustomActionDeleteCancel = document.getElementById('btnCustomActionDeleteCancel');
  var $btnCustomActionDeleteConfirm = document.getElementById('btnCustomActionDeleteConfirm');

  // -- Bridge to host ----------------------------------------------
  function postToHost(obj) {
    try {
      if (window.chrome && window.chrome.webview && window.chrome.webview.postMessage) {
        window.chrome.webview.postMessage(JSON.stringify(obj));
      } else {
        // Dev fallback: log to console.
        console.log('[hostPost]', obj);
      }
    } catch (err) {
      console.error('postToHost failed', err);
    }
  }

  function fallbackEscapeHtml(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function escapeAttr(s) {
    return fallbackEscapeHtml(s);
  }

  function cssEscape(s) {
    if (window.CSS && typeof window.CSS.escape === 'function') {
      return window.CSS.escape(String(s));
    }
    return String(s).replace(/(["\\\]\[])/g, '\\$1');
  }

  function renderMarkdown(src) {
    if (window.markdown && typeof window.markdown.render === 'function') {
      return window.markdown.render(src);
    }
    if (!src) return '';
    return '<p>' + fallbackEscapeHtml(src).replace(/\n/g, '<br>') + '</p>';
  }

  // -- DOM-builder helpers ----------------------------------------
  function elt(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined && text !== null) e.textContent = text;
    return e;
  }

  function formatBytes(bytes) {
    var n = Number(bytes);
    if (!isFinite(n) || n < 0) return '';
    if (n < 1024) return Math.round(n) + ' B';
    if (n < 1024 * 1024) return (n / 1024).toFixed(1) + ' KB';
    if (n < 1024 * 1024 * 1024) return (n / 1024 / 1024).toFixed(1) + ' MB';
    return (n / 1024 / 1024 / 1024).toFixed(1) + ' GB';
  }

  function formatLabel(format) {
    switch (String(format || '').toLowerCase()) {
      case 'xlsx': return 'Книга Excel';
      case 'pdf': return 'PDF-документ';
      default: return format ? String(format).toUpperCase() : 'Файл';
    }
  }

  function filenameFromPath(path) {
    var value = String(path || '');
    var slash = Math.max(value.lastIndexOf('\\'), value.lastIndexOf('/'));
    return slash >= 0 ? value.substring(slash + 1) : value;
  }

  function postFileAction(type, path) {
    if (!path) return;
    postToHost({
      type: type,
      payload: { path: path }
    });
  }

  function postItemAction(type, id) {
    if (!id) return;
    postToHost({
      type: type,
      payload: { id: id }
    });
  }

  function findAssistantMessage(messageId) {
    if (messageId !== undefined && messageId !== null && messageId !== '') {
      var selector = '[data-message-id="' + cssEscape(messageId) + '"]';
      var matched = null;
      try { matched = document.querySelector(selector); } catch (e) { matched = null; }
      if (matched) return matched;
    }

    var messages = document.querySelectorAll('.msg-assistant');
    return messages.length ? messages[messages.length - 1] : null;
  }

  function appendFileCardToMessage(messageId, fileInfo) {
    if (!fileInfo || !fileInfo.path) return;

    var msgEl = findAssistantMessage(messageId);
    if (!msgEl) return;

    var attach = msgEl.querySelector('.msg-attachments');
    if (!attach) {
      attach = elt('div', 'msg-attachments');
      msgEl.appendChild(attach);
    }

    var filePath = String(fileInfo.path);
    var filename = fileInfo.filename || filenameFromPath(filePath) || 'Сохранённый файл';
    var format = String(fileInfo.format || '').toLowerCase();
    var bytes = formatBytes(fileInfo.bytes);
    var label = formatLabel(fileInfo.format);
    var sub = bytes ? (bytes + ' \u00B7 ' + label) : label;

    var card = elt('div', 'file-card');
    card.setAttribute('data-format', escapeAttr(format));

    card.appendChild(elt('div', 'file-card-icon'));

    var meta = elt('div', 'file-card-meta');
    var name = elt('div', 'file-card-name', filename);
    name.title = filename;
    meta.appendChild(name);
    meta.appendChild(elt('div', 'file-card-sub', sub));
    card.appendChild(meta);

    var actions = elt('div', 'file-card-actions');
    var open = elt('button', 'file-card-btn', 'Открыть');
    open.type = 'button';
    open.addEventListener('click', function() {
      postFileAction('open_file', filePath);
    });

    var reveal = elt('button', 'file-card-btn', 'Показать в папке');
    reveal.type = 'button';
    reveal.addEventListener('click', function() {
      postFileAction('reveal_in_explorer', filePath);
    });

    actions.appendChild(open);
    actions.appendChild(reveal);
    card.appendChild(actions);

    attach.appendChild(card);
    scrollToBottom();
  }

  function appendDraftCardToMessage(messageId, draftInfo) {
    if (!draftInfo || !draftInfo.id) return;

    var msgEl = findAssistantMessage(messageId);
    if (!msgEl) return;

    var attach = msgEl.querySelector('.msg-attachments');
    if (!attach) {
      attach = elt('div', 'msg-attachments');
      msgEl.appendChild(attach);
    }

    var title = draftInfo.title || '\u0427\u0435\u0440\u043d\u043e\u0432\u0438\u043a Outlook';
    var location = draftInfo.location || 'Outlook';
    var card = elt('div', 'file-card');
    card.setAttribute('data-format', 'draft');
    card.appendChild(elt('div', 'file-card-icon'));

    var meta = elt('div', 'file-card-meta');
    var name = elt('div', 'file-card-name', title);
    name.title = title;
    meta.appendChild(name);
    meta.appendChild(elt('div', 'file-card-sub', location));
    card.appendChild(meta);

    var actions = elt('div', 'file-card-actions');
    var open = elt('button', 'file-card-btn', '\u041e\u0442\u043a\u0440\u044b\u0442\u044c');
    open.type = 'button';
    open.addEventListener('click', function() {
      postItemAction('open_item', draftInfo.id);
    });
    actions.appendChild(open);
    card.appendChild(actions);

    attach.appendChild(card);
    scrollToBottom();
  }

  var retryableExportErrorCodes = [
    'file_locked',
    'webview2_missing',
    'path_timeout',
    'pdf_render_timeout'
  ];

  function exportErrorCode(err) {
    try {
      if (!err || typeof err === 'string') return '';
      if (err.error && typeof err.error === 'object') return String(err.error.code || '');
      return String(err.error || '');
    } catch (e) {
      return '';
    }
  }

  function isRetryableExportError(err) {
    return retryableExportErrorCodes.indexOf(exportErrorCode(err)) >= 0;
  }

  function renderInlineErrorCard(messageId, err) {
    var msgEl = findAssistantMessage(messageId);
    if (!msgEl) return false;

    var retryMessageId = messageId;
    if (msgEl.dataset && msgEl.dataset.messageId) {
      retryMessageId = msgEl.dataset.messageId;
    }

    var attach = msgEl.querySelector('.msg-attachments');
    if (!attach) {
      attach = elt('div', 'msg-attachments');
      msgEl.appendChild(attach);
    }

    var card = elt('div', 'error-card');
    card.appendChild(elt('div', 'error-card-icon'));

    var body = elt('div', 'error-card-body');
    body.appendChild(elt('div', 'error-card-title', 'Не удалось экспортировать'));
    var detail = elt('div', 'error-card-detail');
    detail.textContent = exportErrorMessage(err);
    body.appendChild(detail);

    var code = exportErrorCode(err);
    var actions = elt('div', 'error-card-actions');
    var hasActions = false;

    if (isRetryableExportError(err) && retryMessageId !== undefined && retryMessageId !== null && retryMessageId !== '') {
      var retryEntry = assistantMessages[retryMessageId];
      if (retryEntry && retryEntry.complete) {
        var retry = elt('button', 'error-card-btn', 'Повторить');
        retry.type = 'button';
        retry.addEventListener('click', function() {
          handleExportPdf(retryMessageId);
        });
        actions.appendChild(retry);
        hasActions = true;
      }
    }

    if (code === 'webview2_missing') {
      var install = elt('a', 'error-card-link', 'Установить WebView2 Runtime');
      install.href = 'https://developer.microsoft.com/microsoft-edge/webview2/';
      install.target = '_blank';
      install.rel = 'noopener noreferrer';
      actions.appendChild(install);
      hasActions = true;
    }

    if (hasActions) body.appendChild(actions);
    card.appendChild(body);
    attach.appendChild(card);

    resetExportButton(retryMessageId);
    scrollToBottom();
    return true;
  }

  function renderToolResultIfHandled(messageId, resultJson) {
    if (!resultJson) return false;
    try {
      var obj = (typeof resultJson === 'string') ? JSON.parse(resultJson) : resultJson;
      if (obj && obj.result_type === 'file_saved') {
        appendFileCardToMessage(obj.message_id || obj.messageId || messageId, obj);
        return true;
      }
    } catch (e) { /* fall through to existing path */ }
    return false;
  }

  function normalizePdfExportToolError(resultJson) {
    if (!resultJson) return null;
    try {
      var obj = (typeof resultJson === 'string') ? JSON.parse(resultJson) : resultJson;
      if (!obj || !obj.error || typeof obj.error !== 'object') return null;
      return {
        error: obj.error.code || 'pdf_render_failed',
        detail: obj.error.message || obj.error.detail || 'Неизвестная ошибка экспорта.'
      };
    } catch (e) {
      return null;
    }
  }

  function exportErrorMessage(error) {
    var fallback = 'Неизвестная ошибка экспорта.';
    try {
      if (!error) return fallback;
      if (typeof error === 'string') return error || fallback;
      if (error.error && typeof error.error === 'object') {
        return String(error.error.message || error.error.detail || error.error.code || fallback);
      }
      var detail = error.detail || error.message || error.error;
      if (detail) return String(detail);
      return fallback;
    } catch (e) {
      return fallback;
    }
  }

  function scrollToBottom() {
    $messages.scrollTop = $messages.scrollHeight;
  }

  function deriveFilenameHint(markdown) {
    var text = String(markdown || '');
    var lines = text.split(/\r?\n/);
    for (var i = 0; i < lines.length; i++) {
      var m = /^#{1,3}\s+(.+?)\s*$/.exec(lines[i]);
      if (m && m[1]) {
        var heading = m[1].trim();
        heading = heading.replace(/\s+#+\s*$/, '').trim();
        if (heading) return heading.substring(0, 60);
      }
    }
    return 'Отчёт OutlookAI';
  }

  function setExportButtonPending(messageId, pending) {
    var entry = assistantMessages[messageId];
    if (!entry) return;
    entry.exportPending = !!pending;
    if (contextAssistantMessageId === messageId && $btnAssistantExportPdf) {
      $btnAssistantExportPdf.disabled = !!pending;
      $btnAssistantExportPdf.dataset.exportPending = pending ? '1' : '0';
      $btnAssistantExportPdf.textContent = pending ? 'Сохранение...' : 'Сохранить как PDF';
    }
  }

  function resetExportButton(messageId) {
    var entry = assistantMessages[messageId];
    if (!entry) return;
    entry.exportPending = false;
    if (contextAssistantMessageId === messageId && $btnAssistantExportPdf) {
      $btnAssistantExportPdf.dataset.exportPending = '0';
      $btnAssistantExportPdf.textContent = 'Сохранить как PDF';
      $btnAssistantExportPdf.disabled = !entry.complete;
    }
  }

  function handleExportPdf(messageId) {
    var entry = assistantMessages[messageId];
    var msgEl = entry && entry.container;
    if (!msgEl && messageId !== undefined && messageId !== null && messageId !== '') {
      try { msgEl = document.querySelector('[data-message-id="' + cssEscape(messageId) + '"]'); } catch (e) { msgEl = null; }
    }
    if (!entry || !msgEl || !entry.complete) return;
    if (entry.exportPending) return;

    var markdown = entry.raw;
    if (markdown === undefined || markdown === null || markdown === '') {
      markdown = entry.content ? entry.content.innerText : '';
    }
    markdown = String(markdown || '');

    setExportButtonPending(messageId, true);
    postToHost({
      type: 'export_pdf',
      payload: {
        message_id: messageId,
        filename_hint: deriveFilenameHint(markdown),
        content_markdown: markdown
      }
    });
  }

  var assistantMessages = {}; // id -> { container, content, raw }
  var toolCards = {};         // callId -> element
  var customActionsById = {};
  var editingCustomActionId = null;
  var contextCustomActionId = null;
  var contextAssistantMessageId = null;
  var pendingDeleteCustomActionId = null;
  var customActionToolCatalog = [];
  var actionCatalogGroups = [];
  var recommendationsEnabled = false;
  var activeActionGroupId = null;
  var editingActionGroupId = null;

  function toolDisplayName(name) {
    var labels = {
      outlook_get_current_compose_state: 'Открытое письмо',
      outlook_get_current_selection: 'Выбранные письма',
      outlook_list_folders: 'Список папок',
      outlook_search_messages: 'Поиск писем',
      outlook_read_message: 'Чтение письма',
      outlook_read_messages: 'Чтение нескольких писем',
      outlook_count_messages: 'Подсчет писем',
      outlook_aggregate_messages: 'Группировка писем',
      outlook_list_recent_threads_with: 'Недавние переписки',
      outlook_export_excel: 'Экспорт Excel',
      outlook_export_pdf: 'Экспорт PDF',
      outlook_export_search_results: 'Экспорт результатов поиска',
      outlook_create_draft: 'Создание черновика',
      outlook_mark_as_read: 'Изменение статуса прочтения',
      outlook_flag_message: 'Флаг письма',
      outlook_set_category: 'Категория письма'
    };
    return labels[name] || name;
  }

  function renderCustomActionTools(selectedNames) {
    if (!$customActionTools || !$customActionToolsField) return;
    var selected = {};
    (selectedNames || []).forEach(function(name) { selected[name] = true; });
    while ($customActionTools.firstChild) $customActionTools.removeChild($customActionTools.firstChild);
    setFieldVisible($customActionToolsField, customActionToolCatalog.length > 0);

    customActionToolCatalog.forEach(function(tool) {
      var label = elt('label', 'custom-action-tool');
      label.title = tool.description || tool.name;
      var checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.value = tool.name;
      checkbox.checked = !!selected[tool.name];
      checkbox.dataset.toolName = tool.name;
      label.appendChild(checkbox);
      label.appendChild(elt('span', 'custom-action-tool-name', toolDisplayName(tool.name)));
      label.appendChild(elt(
        'span',
        'custom-action-tool-access' + (tool.is_write ? ' write' : ''),
        tool.is_write ? 'изменяет Outlook' : 'чтение'));
      $customActionTools.appendChild(label);
    });
  }

  function selectedCustomActionTools() {
    if (!$customActionTools) return [];
    var result = [];
    var checked = $customActionTools.querySelectorAll('input[type="checkbox"]:checked');
    for (var i = 0; i < checked.length; i++) {
      result.push(checked[i].value);
    }
    return result;
  }

  function setFieldVisible(field, visible) {
    if (field) field.hidden = !visible;
  }

  function updateCustomActionFieldVisibility() {
    var source = $customActionSource ? $customActionSource.value : 'current_selection';
    var isFolderSource = source === 'current_folder' || source === 'all_folders';
    var hasMultipleItems = source !== 'current_open_message';
    var isManualRange = isFolderSource && $customActionPeriod && $customActionPeriod.value === 'manual';
    var attachmentsAffectContext = !isFolderSource
      || ($customActionFullBodies && $customActionFullBodies.checked);

    setFieldVisible($customActionFilterField, isFolderSource);
    setFieldVisible($customActionPeriodField, isFolderSource);
    setFieldVisible($customActionMaxItemsField, hasMultipleItems);
    setFieldVisible($customActionManualFromField, isManualRange);
    setFieldVisible($customActionManualToField, isManualRange);
    setFieldVisible($customActionFullBodiesField, true);
    setFieldVisible($customActionAttachmentsField, attachmentsAffectContext);
  }

  function toLocalDateTimeInput(value) {
    if (!value) return '';
    var date = new Date(value);
    if (isNaN(date.getTime())) return '';
    var offsetMs = date.getTimezoneOffset() * 60000;
    return new Date(date.getTime() - offsetMs).toISOString().slice(0, 16);
  }

  function toIsoDateTime(value) {
    if (!value) return null;
    var date = new Date(value);
    return isNaN(date.getTime()) ? null : date.toISOString();
  }

  function openCustomActionDialog(action) {
    if (!$customActionDialog) return;
    action = action || null;
    editingCustomActionId = action && action.id ? action.id : null;
    if (action && action.group_id) editingActionGroupId = action.group_id;
    $customActionTitle.textContent = editingCustomActionId ? 'Изменить действие' : 'Новое действие';
    $btnCustomActionSave.textContent = editingCustomActionId ? 'Сохранить' : 'Добавить';
    $customActionName.value = action && action.label || '';
    $customActionHint.value = action && action.description || '';
    $customActionPrompt.value = action && action.action_prompt || '';
    $customActionSource.value = action && action.source || 'current_selection';
    $customActionFilter.value = action && action.read_filter || 'all';
    $customActionPeriod.value = action && action.time_range || 'today';
    $customActionMaxItems.value = action && action.max_items || '20';
    $customActionManualFrom.value = toLocalDateTimeInput(action && action.manual_from);
    $customActionManualTo.value = toLocalDateTimeInput(action && action.manual_to);
    $customActionFullBodies.checked = action ? action.include_full_bodies !== false : true;
    $customActionAttachments.checked = !!(action && action.include_attachments);
    var output = action && action.output || 'chat';
    if (output === 'create_draft') output = 'create_reply';
    $customActionOutput.value = output;
    renderCustomActionTools(action && action.allowed_tools || []);
    updateCustomActionFieldVisibility();
    $customActionDialog.hidden = false;
    try { $customActionName.focus(); } catch (e) { /* best-effort */ }
  }

  function closeCustomActionDialog() {
    if ($customActionDialog) $customActionDialog.hidden = true;
    editingCustomActionId = null;
  }

  function closeCustomActionMenu() {
    if ($customActionMenu) $customActionMenu.hidden = true;
    contextCustomActionId = null;
  }

  function closeAssistantMessageMenu() {
    if ($assistantMessageMenu) $assistantMessageMenu.hidden = true;
    contextAssistantMessageId = null;
  }

  function openAssistantMessageMenu(messageId, clientX, clientY) {
    var entry = assistantMessages[messageId];
    if (!$assistantMessageMenu || !entry || !entry.complete) return;
    closeCustomActionMenu();
    contextAssistantMessageId = messageId;
    $assistantMessageMenu.hidden = false;
    if ($btnAssistantExportPdf) {
      $btnAssistantExportPdf.title = 'Save message as PDF';
      $btnAssistantExportPdf.disabled = !!entry.exportPending;
      $btnAssistantExportPdf.dataset.exportPending = entry.exportPending ? '1' : '0';
      $btnAssistantExportPdf.textContent = entry.exportPending
        ? 'Сохранение...'
        : 'Сохранить как PDF';
    }
    var width = $assistantMessageMenu.offsetWidth || 180;
    var height = $assistantMessageMenu.offsetHeight || 40;
    var left = Math.max(4, Math.min(clientX, window.innerWidth - width - 4));
    var top = Math.max(4, Math.min(clientY, window.innerHeight - height - 4));
    $assistantMessageMenu.style.left = left + 'px';
    $assistantMessageMenu.style.top = top + 'px';
  }

  function openCustomActionMenu(actionId, clientX, clientY) {
    if (!$customActionMenu || !customActionsById[actionId]) return;
    closeAssistantMessageMenu();
    contextCustomActionId = actionId;
    $customActionMenu.hidden = false;
    var width = $customActionMenu.offsetWidth || 140;
    var height = $customActionMenu.offsetHeight || 72;
    var left = Math.max(4, Math.min(clientX, window.innerWidth - width - 4));
    var top = Math.max(4, Math.min(clientY, window.innerHeight - height - 4));
    $customActionMenu.style.left = left + 'px';
    $customActionMenu.style.top = top + 'px';
  }

  function openDeleteConfirmation(actionId) {
    var action = customActionsById[actionId];
    if (!action || !$customActionDeleteDialog) return;
    pendingDeleteCustomActionId = actionId;
    $customActionDeleteText.textContent = 'Действие «' + action.label + '» будет удалено без возможности восстановления.';
    $customActionDeleteDialog.hidden = false;
    try { $btnCustomActionDeleteCancel.focus(); } catch (e) { /* best-effort */ }
  }

  function closeDeleteConfirmation() {
    pendingDeleteCustomActionId = null;
    if ($customActionDeleteDialog) $customActionDeleteDialog.hidden = true;
  }

  function saveCustomActionFromDialog() {
    var title = ($customActionName && $customActionName.value || '').trim();
    var prompt = ($customActionPrompt && $customActionPrompt.value || '').trim();
    if (!title || !prompt) {
      api.showError('Заполните название и промпт действия.');
      return;
    }

    var maxItems = parseInt($customActionMaxItems && $customActionMaxItems.value || '20', 10);
    if (!isFinite(maxItems) || maxItems < 1) maxItems = 20;
    if (maxItems > 100) maxItems = 100;
    var source = $customActionSource ? $customActionSource.value : 'current_selection';
    var isFolderSource = source === 'current_folder' || source === 'all_folders';
    var readFilter = isFolderSource && $customActionFilter ? $customActionFilter.value : 'all';
    var timeRange = isFolderSource && $customActionPeriod ? $customActionPeriod.value : 'today';
    var manualFrom = timeRange === 'manual'
      ? toIsoDateTime($customActionManualFrom && $customActionManualFrom.value)
      : null;
    var manualTo = timeRange === 'manual'
      ? toIsoDateTime($customActionManualTo && $customActionManualTo.value)
      : null;
    var includeAttachments = (!isFolderSource || ($customActionFullBodies && $customActionFullBodies.checked))
      ? !!($customActionAttachments && $customActionAttachments.checked)
      : false;
    if (timeRange === 'manual' && (!manualFrom || !manualTo)) {
      api.showError('Укажите начало и конец ручного периода.');
      return;
    }
    if (manualFrom && manualTo && new Date(manualFrom).getTime() > new Date(manualTo).getTime()) {
      api.showError('Начало периода должно быть раньше его окончания.');
      return;
    }
    var allowedTools = selectedCustomActionTools();

    postToHost({
      type: 'custom_action_create',
      payload: {
        id: editingCustomActionId,
        group_id: editingActionGroupId,
        title: title,
        description: ($customActionHint && $customActionHint.value || '').trim(),
        prompt: prompt,
        source: source,
        read_filter: readFilter,
        time_range: timeRange,
        manual_from: manualFrom,
        manual_to: manualTo,
        max_items: source === 'current_open_message' ? 1 : maxItems,
        include_full_bodies: !!($customActionFullBodies && $customActionFullBodies.checked),
        include_attachments: includeAttachments,
        output: $customActionOutput ? $customActionOutput.value : 'chat',
        allow_tools: allowedTools.length > 0,
        allowed_tools: allowedTools
      }
    });
    closeCustomActionDialog();
  }

  function actionById(actionId) {
    return customActionsById[actionId] || null;
  }

  function groupById(groupId) {
    for (var i = 0; i < actionCatalogGroups.length; i++) {
      if (actionCatalogGroups[i].id === groupId) return actionCatalogGroups[i];
    }
    return null;
  }

  function runCustomAction(actionId) {
    if (!actionId) return;
    closeActionGroupPopover();
    postToHost({ type: 'custom_action', payload: { id: actionId } });
  }

  function closeActionGroupPopover() {
    activeActionGroupId = null;
    if ($actionGroupPopover) {
      $actionGroupPopover.hidden = true;
      $actionGroupPopover.innerHTML = '';
    }
  }

  function renderActionGroupPopover(group) {
    if (!$actionGroupPopover || !group) return;
    $actionGroupPopover.innerHTML = '';
    (group.actions || []).forEach(function(action) {
      var button = elt('button', 'group-popover-action');
      button.type = 'button';
      button.appendChild(elt('span', 'group-popover-title', action.label));
      if (action.description) {
        button.appendChild(elt('span', 'group-popover-description', action.description));
      }
      button.addEventListener('click', function() { runCustomAction(action.id); });
      $actionGroupPopover.appendChild(button);
    });
    var footer = elt('div', 'group-popover-footer');
    var edit = elt('button', 'btn btn-ghost', 'Изменить');
    edit.type = 'button';
    edit.addEventListener('click', function() {
      closeActionGroupPopover();
      openGroupEditor(group.id);
    });
    footer.appendChild(edit);
    $actionGroupPopover.appendChild(footer);
    $actionGroupPopover.hidden = false;
    activeActionGroupId = group.id;
  }

  function setRecommendationState(state, ids) {
    if (!$recommendationsSection || !$actionRecommendations) return;
    $actionRecommendations.innerHTML = '';
    $recommendationsLoading.hidden = state !== 'loading';
    if (state === 'disabled' || state === 'empty' || state === 'error') {
      $recommendationsSection.hidden = true;
      return;
    }
    (ids || []).slice(0, 5).forEach(function(id) {
      var action = actionById(id);
      if (!action) return;
      var button = elt('button', 'qa-chip', action.label);
      button.type = 'button';
      button.title = action.description || action.action_prompt || '';
      button.addEventListener('click', function() { runCustomAction(action.id); });
      $actionRecommendations.appendChild(button);
    });
    $recommendationsSection.hidden =
      state !== 'loading' && !$actionRecommendations.children.length;
  }

  function renderGroupButtons() {
    if (!$actionGroups) return;
    $actionGroups.innerHTML = '';
    actionCatalogGroups.forEach(function(group) {
      var button = elt('button', 'qa-chip action-group-btn', group.title);
      button.type = 'button';
      button.addEventListener('click', function() {
        if (activeActionGroupId === group.id) closeActionGroupPopover();
        else renderActionGroupPopover(group);
      });
      $actionGroups.appendChild(button);
    });
  }

  function openGroupEditor(groupId) {
    var group = groupById(groupId);
    if (!group || !$actionGroupEditor) return;
    editingActionGroupId = groupId;
    $groupEditorTitle.textContent = group.title;
    if ($btnGroupReset) $btnGroupReset.hidden = group.can_reset === false;
    $actionGroupEditor.hidden = false;
    renderGroupEditor();
  }

  function closeGroupEditor() {
    if ($actionGroupEditor) $actionGroupEditor.hidden = true;
    if ($groupImportPanel) $groupImportPanel.hidden = true;
    editingActionGroupId = null;
  }

  function reorderGroupAction(actionId, delta) {
    var group = groupById(editingActionGroupId);
    if (!group) return;
    var index = group.actions.findIndex(function(action) { return action.id === actionId; });
    var target = index + delta;
    if (index < 0 || target < 0 || target >= group.actions.length) return;
    var temp = group.actions[index];
    group.actions[index] = group.actions[target];
    group.actions[target] = temp;
    renderGroupEditor();
    postToHost({
      type: 'custom_action_reorder',
      payload: {
        group_id: group.id,
        action_ids: group.actions.map(function(action) { return action.id; })
      }
    });
  }

  function renderGroupEditor() {
    var group = groupById(editingActionGroupId);
    if (!group || !$groupEditorActions) return;
    $groupEditorActions.innerHTML = '';
    group.actions.forEach(function(action, index) {
      var row = elt('div', 'group-editor-row');
      var text = elt('div');
      text.appendChild(elt('div', 'group-editor-row-title', action.label));
      text.appendChild(elt('div', 'group-editor-row-description', action.description || ''));
      row.appendChild(text);
      var controls = elt('div', 'group-editor-row-actions');
      var up = elt('button', 'group-row-btn', '↑');
      up.type = 'button'; up.title = 'Выше'; up.disabled = index === 0;
      up.addEventListener('click', function() { reorderGroupAction(action.id, -1); });
      var down = elt('button', 'group-row-btn', '↓');
      down.type = 'button'; down.title = 'Ниже'; down.disabled = index === group.actions.length - 1;
      down.addEventListener('click', function() { reorderGroupAction(action.id, 1); });
      var edit = elt('button', 'group-row-btn', '✎');
      edit.type = 'button'; edit.title = 'Изменить';
      edit.addEventListener('click', function() { openCustomActionDialog(action); });
      var exp = elt('button', 'group-row-btn', '⇩');
      exp.type = 'button'; exp.title = 'Экспорт YAML';
      exp.addEventListener('click', function() {
        postToHost({ type: 'custom_action_export', payload: { id: action.id } });
      });
      var del = elt('button', 'group-row-btn danger-text', '×');
      del.type = 'button'; del.title = 'Удалить';
      del.addEventListener('click', function() { openDeleteConfirmation(action.id); });
      controls.appendChild(up); controls.appendChild(down); controls.appendChild(edit);
      controls.appendChild(exp); controls.appendChild(del);
      row.appendChild(controls);
      $groupEditorActions.appendChild(row);
    });
  }

  // -- Public API --------------------------------------------------
  var api = {
    appendUserMessage: function(text) {
      var node = elt('div', 'msg msg-user');
      var content = elt('div', 'msg-content');
      content.innerHTML = renderMarkdown(text);
      node.appendChild(content);
      $messages.appendChild(node);
      scrollToBottom();
    },

    appendAssistantMessage: function(id, initialText) {
      var node = elt('div', 'msg msg-assistant');
      node.dataset.messageId = id;
      node.dataset.state = 'streaming';
      node.classList.add('is-streaming');
      var content = elt('div', 'msg-content');
      content.innerHTML = renderMarkdown(initialText || '');
      node.appendChild(content);
      var loading = elt('div', 'msg-loading');
      loading.setAttribute('role', 'status');
      loading.setAttribute('aria-label', 'Формируется ответ');
      loading.appendChild(elt('span', 'msg-loading-dots', ''));
      node.appendChild(loading);
      node.addEventListener('contextmenu', function(e) {
        e.preventDefault();
        if (!assistantMessages[id] || !assistantMessages[id].complete) return;
        openAssistantMessageMenu(id, e.clientX, e.clientY);
      });
      $messages.appendChild(node);
      assistantMessages[id] = {
        container: node,
        content: content,
        raw: initialText || '',
        complete: false,
        loading: loading,
        exportPending: false
      };
      if (initialText) {
        node.classList.add('has-content');
      }
      scrollToBottom();
    },

    appendTextDelta: function(id, delta) {
      var entry = assistantMessages[id];
      if (!entry) {
        api.appendAssistantMessage(id, delta);
        return;
      }
      if (!delta) return;
      entry.raw += delta;
      entry.container.classList.add('has-content');
      // For streaming, render the running text as markdown. This is
      // cheap enough for typical email-length replies.
      entry.content.innerHTML = renderMarkdown(entry.raw);
      scrollToBottom();
    },

    finalizeAssistantMessage: function(id, opts) {
      var entry = assistantMessages[id];
      if (!entry) return;
      opts = opts || {};
      entry.complete = true;
      entry.container.dataset.state = 'complete';
      entry.container.classList.remove('is-streaming');
      if (!entry.raw) {
        entry.content.textContent = opts.stopped
          ? 'Ответ остановлен.'
          : (opts.error ? 'Не удалось получить ответ.' : '');
      }
      if (opts.stopped) entry.container.classList.add('msg-stopped');
      if (opts.error) entry.container.classList.add('msg-error');
      resetExportButton(id);
      scrollToBottom();
    },

    // Compact "working on it" status line. Replaces the previous verbose
    // expandable tool cards (the chat got too cluttered - user feedback).
    // Each tool call shows as a single italic line like
    //   ... Searching messages
    //   ... Reading message
    //   ... Listing folders
    // When the tool completes, the line is REMOVED entirely (we don't
    // need a permanent record - the model uses the result to produce
    // the actual assistant text, which is what the user reads).
    //
    // Friendly verbs per tool name. Anything not in the map falls back
    // to "Working on it..." so unknown future tools still render fine.
    appendToolCallCard: function(callId, name, argsJson) {
      var verb = ({
        outlook_get_current_compose_state: 'Читаю контекст письма',
        outlook_get_current_selection:     'Читаю текущее выделение',
        outlook_list_folders:              'Получаю список папок',
        outlook_search_messages:           'Ищу сообщения',
        outlook_read_message:              'Читаю сообщение',
        outlook_count_messages:            'Считаю сообщения',
        outlook_list_recent_threads_with:  'Ищу недавние переписки',
        outlook_create_draft:              'Создаю черновик',
        outlook_mark_as_read:              'Отмечаю как прочитанное',
        outlook_flag_message:              'Ставлю флаг на сообщение',
        outlook_set_category:              'Назначаю категорию',
      })[name] || 'Выполняю действие';

      var row = elt('div', 'tool-status');
      row.dataset.callId = callId;
      row.dataset.toolName = String(name || '');
      row.dataset.verb = verb;
      row.dataset.startedAt = String(Date.now());
      row.textContent = '\u2026 ' + verb;
      $messages.appendChild(row);

      // Live time counter. Updates the status line every second with the
      // elapsed seconds so the user can see the tool is making progress
      // rather than guessing whether Outlook has frozen. Counter math
      // uses Date.now() - startedAt so the displayed value is always
      // wall-clock-correct even if setInterval is throttled when the
      // WebView2 loses focus.
      var tick = function() {
        if (!row.parentNode) { return; }
        var ms = Date.now() - parseInt(row.dataset.startedAt, 10);
        var secs = Math.max(0, Math.floor(ms / 1000));
        if (secs > 0) {
          row.textContent = '\u2026 ' + verb + ' (' + secs + 's)';
        }
      };
      var tickId = setInterval(function() {
        if (!row.parentNode) { clearInterval(tickId); return; }
        tick();
      }, 1000);
      row.dataset.tickId = String(tickId);

      // Force an immediate refresh when the WebView2 regains focus.
      // WebView2 throttles setInterval while hidden; without this, the
      // counter appears to lag (or to "reset" perceptually) until the
      // next tick after focus returns.
      var visHandler = function() {
        if (document.visibilityState === 'visible') tick();
      };
      document.addEventListener('visibilitychange', visHandler);
      window.addEventListener('focus', visHandler);
      row.dataset.hasVisHandler = '1';
      row._oai_visHandler = visHandler;

      toolCards[callId] = row;
      scrollToBottom();
    },

    updateToolCallCard: function(callId, ok, summary, resultJson) {
      var row = toolCards[callId];
      if (!row) return;
      // Stop the live time counter.
      var tickId = parseInt(row.dataset.tickId || '0', 10);
      if (tickId) clearInterval(tickId);
      // Drop the visibility-change handler (one was attached per row).
      if (row._oai_visHandler) {
        try {
          document.removeEventListener('visibilitychange', row._oai_visHandler);
          window.removeEventListener('focus', row._oai_visHandler);
        } catch (e) { /* best-effort */ }
        row._oai_visHandler = null;
      }
      // On completion, drop the status line. Errors stick around as a
      // muted single-line error so the user sees that something failed
      // without the full JSON dump.
      if (ok) {
        renderToolResultIfHandled(callId, resultJson);
        if (row.parentNode) row.parentNode.removeChild(row);
      } else {
        var pdfExportError = row.dataset.toolName === 'outlook_export_pdf' ? normalizePdfExportToolError(resultJson) : null;
        if (pdfExportError && renderInlineErrorCard(null, pdfExportError)) {
          if (row.parentNode) row.parentNode.removeChild(row);
        } else {
          row.classList.add('tool-status-err');
          row.textContent = '\u26A0 ' + (summary || 'ошибка инструмента');
        }
      }
      delete toolCards[callId];
      scrollToBottom();
    },

    appendAuditRow: function(text) {
      var row = elt('div', 'audit-row');
      var time = elt('span', 'audit-time', new Date().toLocaleTimeString());
      row.appendChild(time);
      row.appendChild(document.createTextNode(text));
      $messages.appendChild(row);
      scrollToBottom();
    },

    onFileSaved: function(messageId, fileInfo) {
      appendFileCardToMessage(messageId, fileInfo);
      resetExportButton(messageId);
    },

    onDraftCreated: function(messageId, draftInfo) {
      appendDraftCardToMessage(messageId, draftInfo);
    },

    onExportError: function(messageId, error) {
      if (!renderInlineErrorCard(messageId, error)) {
        resetExportButton(messageId);
        api.showError(exportErrorMessage(error));
      }
    },

    handleExportPdf: handleExportPdf,

    showError: function(message) {
      var node = elt('div', 'msg msg-assistant msg-error');
      var content = elt('div', 'msg-content');
      content.textContent = message || 'Произошла ошибка.';
      node.appendChild(content);
      $messages.appendChild(node);
      scrollToBottom();
    },

    setComposerEnabled: function(enabled, isStopVisible) {
      $input.disabled = !enabled;
      $btnSend.disabled = !enabled;
      $btnSend.hidden = !!isStopVisible;
      $btnStop.hidden = !isStopVisible;
      $btnClear.disabled = !enabled && !isStopVisible;
      $btnCopy.disabled = false; // always allow copy
      if (enabled) $input.focus();
    },

    clear: function() {
      // Stop any live tool-status time counters and unbind any
      // visibility-change handlers before tossing their DOM.
      for (var cid in toolCards) {
        try {
          var row = toolCards[cid];
          var tickId = row && parseInt(row.dataset && row.dataset.tickId || '0', 10);
          if (tickId) clearInterval(tickId);
          if (row && row._oai_visHandler) {
            try {
              document.removeEventListener('visibilitychange', row._oai_visHandler);
              window.removeEventListener('focus', row._oai_visHandler);
            } catch (e) { /* best-effort */ }
            row._oai_visHandler = null;
          }
        } catch (e) { /* best-effort */ }
      }
      $messages.innerHTML = '';
      assistantMessages = {};
      toolCards = {};
      closeAssistantMessageMenu();
    },

    applyTheme: function(themeName) {
      var allowed = ['light', 'dark', 'high-contrast'];
      var theme = allowed.indexOf(themeName) >= 0 ? themeName : 'light';
      document.body.classList.remove('theme-light', 'theme-dark', 'theme-high-contrast');
      document.body.classList.add('theme-' + theme);
    },

    setContextStrip: function(ctx) {
      // Context still goes to the model; it is intentionally not rendered.
    },

    setActionCatalog: function(payload) {
      payload = payload || {};
      actionCatalogGroups = payload.groups || [];
      recommendationsEnabled = payload.recommendations_enabled === true;
      customActionToolCatalog = payload.tools || [];
      customActionsById = {};
      actionCatalogGroups.forEach(function(group) {
        (group.actions || []).forEach(function(action) {
          action.group_id = group.id;
          customActionsById[action.id] = action;
        });
      });
      if ($quickActions) {
        $quickActions.innerHTML = '';
        $quickActions.hidden = true;
      }
      if ($actionWorkspace) $actionWorkspace.hidden = false;
      closeActionGroupPopover();
      renderGroupButtons();
      if (!recommendationsEnabled) {
        setRecommendationState('disabled', []);
      } else if ((payload.recommendation_ids || []).length) {
        setRecommendationState('ready', payload.recommendation_ids);
      }
      if (editingActionGroupId && !$actionGroupEditor.hidden) {
        var current = groupById(editingActionGroupId);
        if (current) {
          $groupEditorTitle.textContent = current.title;
          renderGroupEditor();
        } else {
          closeGroupEditor();
        }
      }
    },

    setActionRecommendations: function(ids, loading) {
      setRecommendationState(
        !recommendationsEnabled
          ? 'disabled'
          : (loading ? 'loading' : ((ids || []).length ? 'ready' : 'empty')),
        ids || []);
    },

    setActionRecommendationState: function(state, ids) {
      setRecommendationState(
        recommendationsEnabled ? (state || 'empty') : 'disabled',
        ids || []);
    },

    showActionImportPreview: function(yaml, replacements) {
      var items = replacements || [];
      var message = items.length
        ? ('Будут заменены действия:\n\n' + items.join('\n') + '\n\nПродолжить импорт?')
        : 'Добавить действия из YAML?';
      if (window.confirm(message)) {
        postToHost({ type: 'custom_action_import_apply', payload: { yaml: yaml || '' } });
        if ($groupImportPanel) $groupImportPanel.hidden = true;
      }
    },

    /**
     * Render the row of quick-action chips above the composer. Each chip
     * is { label, prompt, id?, type? }. Clicking a custom-action chip
     * posts custom_action; regular chips pre-fill the textarea with
     * the prompt and (by default) immediately sends - the Phase 3a
     * InboxCopilot behavior.
     *
     * Phase 4: callers can pass options = { autoSubmit: false } to
     * disable auto-send and prefill only, so the user can edit
     * [placeholders] in the template before sending. The Reports pane
     * uses this mode.
     *
     * Calling with [] empties the row.
     */
    setQuickActions: function(chips, options) {
      if (!$quickActions) return;
      if ($actionWorkspace) $actionWorkspace.hidden = true;
      $quickActions.hidden = false;
      var autoSubmit = !(options && options.autoSubmit === false);
      customActionToolCatalog = options && options.customActionTools || [];
      while ($quickActions.firstChild) $quickActions.removeChild($quickActions.firstChild);
      customActionsById = {};
      closeCustomActionMenu();
      if (options && options.allowCustomActionManagement) {
        var add = document.createElement('button');
        add.type = 'button';
        add.className = 'qa-chip qa-chip-add';
        add.textContent = '+';
        add.title = 'Добавить действие';
        add.addEventListener('click', function() {
          openCustomActionDialog(null);
        });
        $quickActions.appendChild(add);
      }
      (chips || []).forEach(function(chip) {
        var btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'qa-chip';
        btn.textContent = chip.label;
        btn.title = chip.prompt;
        if (chip.type === 'custom_action') {
          btn.className += ' qa-chip-custom';
          btn.dataset.actionId = chip.id || '';
          customActionsById[chip.id] = chip;
          btn.addEventListener('contextmenu', function(evt) {
            evt.preventDefault();
            evt.stopPropagation();
            openCustomActionMenu(chip.id, evt.clientX, evt.clientY);
          });
        }
        btn.addEventListener('click', function() {
          if (chip.type === 'custom_action' && chip.id) {
            postToHost({
              type: 'custom_action',
              payload: {
                id: chip.id
              }
            });
            return;
          }
          $input.value = chip.prompt;
          if (autoSubmit) {
            sendInput();
          } else {
            // Prefill-only: focus the input and move the caret to the
            // end so the user can immediately edit the [placeholders].
            try { $input.focus(); } catch (e) { /* best-effort */ }
            try {
              var len = $input.value.length;
              if ($input.setSelectionRange) $input.setSelectionRange(len, len);
            } catch (e) { /* best-effort */ }
          }
        });
        $quickActions.appendChild(btn);
      });
    },

    // Kept temporarily for compatibility with hosts that still push options.
    setReasoningOptions: function() {},

    // Used by C# during dev to verify the bridge is live before
    // pushing real messages. Returns a string from JS that C# can
    // assert on with await ExecuteScriptAsync("outlookai.ping()").
    ping: function() { return 'pong'; }
  };

  window.outlookai = window.outlookai || {};
  for (var apiName in api) {
    if (Object.prototype.hasOwnProperty.call(api, apiName)) {
      window.outlookai[apiName] = api[apiName];
    }
  }

  // -- Composer event wiring --------------------------------------
  function sendInput() {
    // Guard against double-send when a previous turn is still running.
    // The host disables $input + $btnSend when a turn is in flight, but
    // a fast Enter or a race between WebMessageReceived and the C# side
    // re-enabling the composer can still let a second submission through
    // and produce two concurrent tool-status rows.
    if ($input.disabled || $btnSend.disabled) return;
    var text = $input.value.trim();
    if (!text) return;
    $input.value = '';
    postToHost({
      type: 'send',
      payload: {
        text: text
      }
    });
  }

  function setComposerInputHeight(height) {
    var minHeight = 64;
    var maxHeight = Math.max(minHeight, Math.min(300, Math.floor(window.innerHeight * 0.45)));
    var nextHeight = Math.max(minHeight, Math.min(maxHeight, Math.round(height)));
    $input.style.height = nextHeight + 'px';
    if ($composerResizer) {
      $composerResizer.setAttribute('aria-valuemin', String(minHeight));
      $composerResizer.setAttribute('aria-valuemax', String(maxHeight));
      $composerResizer.setAttribute('aria-valuenow', String(nextHeight));
    }
  }

  if ($composerResizer) {
    $composerResizer.addEventListener('pointerdown', function(e) {
      if (e.button !== 0) return;
      e.preventDefault();
      var startY = e.clientY;
      var startHeight = $input.getBoundingClientRect().height;
      $composerResizer.classList.add('is-dragging');
      try { $composerResizer.setPointerCapture(e.pointerId); } catch (ignore) {}

      function move(moveEvent) {
        setComposerInputHeight(startHeight + startY - moveEvent.clientY);
      }
      function stop() {
        $composerResizer.classList.remove('is-dragging');
        $composerResizer.removeEventListener('pointermove', move);
        $composerResizer.removeEventListener('pointerup', stop);
        $composerResizer.removeEventListener('pointercancel', stop);
      }

      $composerResizer.addEventListener('pointermove', move);
      $composerResizer.addEventListener('pointerup', stop);
      $composerResizer.addEventListener('pointercancel', stop);
    });
    $composerResizer.addEventListener('keydown', function(e) {
      if (e.key !== 'ArrowUp' && e.key !== 'ArrowDown') return;
      e.preventDefault();
      var delta = e.key === 'ArrowUp' ? 16 : -16;
      setComposerInputHeight($input.getBoundingClientRect().height + delta);
    });
    setComposerInputHeight($input.getBoundingClientRect().height);
  }

  $btnSend.addEventListener('click', sendInput);
  $btnStop.addEventListener('click', function() {
    postToHost({ type: 'stop' });
  });
  $btnClear.addEventListener('click', function() {
    postToHost({ type: 'clear' });
  });
  $btnCopy.addEventListener('click', function() {
    postToHost({ type: 'copy' });
  });
  if ($btnAssistantExportPdf) {
    $btnAssistantExportPdf.addEventListener('click', function() {
      var messageId = contextAssistantMessageId;
      if (!messageId) return;
      handleExportPdf(messageId);
      closeAssistantMessageMenu();
    });
  }
  window.addEventListener('focus', function() {
    postToHost({ type: 'theme_request' });
  });
  document.addEventListener('visibilitychange', function() {
    if (!document.hidden) postToHost({ type: 'theme_request' });
  });
  window.setInterval(function() {
    if (!document.hidden) postToHost({ type: 'theme_request' });
  }, 3000);
  if ($btnCustomActionClose) $btnCustomActionClose.addEventListener('click', closeCustomActionDialog);
  if ($btnCustomActionCancel) $btnCustomActionCancel.addEventListener('click', closeCustomActionDialog);
  if ($btnCustomActionSave) $btnCustomActionSave.addEventListener('click', saveCustomActionFromDialog);
  if ($customActionSource) $customActionSource.addEventListener('change', updateCustomActionFieldVisibility);
  if ($customActionPeriod) $customActionPeriod.addEventListener('change', updateCustomActionFieldVisibility);
  if ($customActionFullBodies) $customActionFullBodies.addEventListener('change', updateCustomActionFieldVisibility);
  if ($btnCustomActionEdit) {
    $btnCustomActionEdit.addEventListener('click', function() {
      var action = customActionsById[contextCustomActionId];
      closeCustomActionMenu();
      if (action) openCustomActionDialog(action);
    });
  }
  if ($btnCustomActionDelete) {
    $btnCustomActionDelete.addEventListener('click', function() {
      var actionId = contextCustomActionId;
      closeCustomActionMenu();
      if (actionId) openDeleteConfirmation(actionId);
    });
  }
  if ($btnCustomActionDeleteCancel) {
    $btnCustomActionDeleteCancel.addEventListener('click', closeDeleteConfirmation);
  }
  if ($btnCustomActionDeleteConfirm) {
    $btnCustomActionDeleteConfirm.addEventListener('click', function() {
      var actionId = pendingDeleteCustomActionId;
      closeDeleteConfirmation();
      if (actionId) {
        postToHost({
          type: 'custom_action_delete',
          payload: { id: actionId }
        });
      }
    });
  }
  if ($btnGroupEditorClose) $btnGroupEditorClose.addEventListener('click', closeGroupEditor);
  if ($btnGroupAddAction) {
    $btnGroupAddAction.addEventListener('click', function() {
      openCustomActionDialog({ group_id: editingActionGroupId });
    });
  }
  if ($btnGroupExport) {
    $btnGroupExport.addEventListener('click', function() {
      postToHost({ type: 'custom_action_export', payload: {} });
    });
  }
  if ($btnGroupReset) {
    $btnGroupReset.addEventListener('click', function() {
      var group = groupById(editingActionGroupId);
      if (!group) return;
      if (window.confirm('Вернуть группу «' + group.title + '» к базовому набору действий?')) {
        postToHost({
          type: 'custom_action_reset_group',
          payload: { group_id: group.id }
        });
      }
    });
  }
  if ($btnGroupImport) {
    $btnGroupImport.addEventListener('click', function() {
      if ($groupImportPanel) $groupImportPanel.hidden = false;
      if ($groupImportText) $groupImportText.focus();
      if ($groupImportFile) $groupImportFile.value = '';
    });
  }
  if ($groupImportFile) {
    $groupImportFile.addEventListener('change', function() {
      var file = $groupImportFile.files && $groupImportFile.files[0];
      if (!file) return;
      var reader = new FileReader();
      reader.onload = function() { $groupImportText.value = String(reader.result || ''); };
      reader.readAsText(file);
    });
  }
  if ($btnGroupImportFile) {
    $btnGroupImportFile.addEventListener('click', function() {
      if ($groupImportFile) $groupImportFile.click();
    });
  }
  if ($btnGroupImportCancel) {
    $btnGroupImportCancel.addEventListener('click', function() {
      $groupImportPanel.hidden = true;
      $groupImportText.value = '';
    });
  }
  if ($btnGroupImportApply) {
    $btnGroupImportApply.addEventListener('click', function() {
      var yaml = ($groupImportText.value || '').trim();
      if (!yaml) {
        api.showError('Вставьте YAML или выберите файл.');
        return;
      }
      postToHost({ type: 'custom_action_import_preview', payload: { yaml: yaml } });
    });
  }
  if ($customActionDeleteDialog) {
    $customActionDeleteDialog.addEventListener('click', function(e) {
      if (e.target === $customActionDeleteDialog) closeDeleteConfirmation();
    });
  }

  // Enter sends, Shift+Enter inserts a newline (standard chat UX).
  $input.addEventListener('keydown', function(e) {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      sendInput();
    }
  });

  document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape' && $customActionDialog && !$customActionDialog.hidden) {
      e.preventDefault();
      closeCustomActionDialog();
      return;
    }
    if (e.key === 'Escape' && $actionGroupEditor && !$actionGroupEditor.hidden) {
      e.preventDefault();
      closeGroupEditor();
      return;
    }
    if (e.key === 'Escape' && $actionGroupPopover && !$actionGroupPopover.hidden) {
      e.preventDefault();
      closeActionGroupPopover();
      return;
    }
    if (e.key === 'Escape' && $customActionDeleteDialog && !$customActionDeleteDialog.hidden) {
      e.preventDefault();
      closeDeleteConfirmation();
      return;
    }
    if (e.key === 'Escape' && $customActionMenu && !$customActionMenu.hidden) {
      e.preventDefault();
      closeCustomActionMenu();
      return;
    }
    if (e.key === 'Escape' && $assistantMessageMenu && !$assistantMessageMenu.hidden) {
      e.preventDefault();
      closeAssistantMessageMenu();
    }
  });

  document.addEventListener('click', function(e) {
    if ($actionGroupPopover && !$actionGroupPopover.hidden &&
        !$actionGroupPopover.contains(e.target) &&
        (!$actionGroups || !$actionGroups.contains(e.target))) {
      closeActionGroupPopover();
    }
    if ($customActionMenu && !$customActionMenu.hidden && !$customActionMenu.contains(e.target)) {
      closeCustomActionMenu();
    }
    if ($assistantMessageMenu && !$assistantMessageMenu.hidden &&
        !$assistantMessageMenu.contains(e.target)) {
      closeAssistantMessageMenu();
    }
  });

  window.addEventListener('blur', function() {
    closeCustomActionMenu();
    closeAssistantMessageMenu();
  });

  // Tell the host we're ready so it can push the initial context strip
  // + apply the theme before the first user message.
  postToHost({ type: 'ready' });
})();

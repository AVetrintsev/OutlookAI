(function() {
  'use strict';
  var $ = function(id) { return document.getElementById(id); };
  var categories = { people: 'Сотрудники и роли', rules: 'Процессы и правила', communication: 'Коммуникация' };
  var items = [], pinned = [], activeCategory = 'people', editing = null, dirty = false;
  var aiId = '', aiBusy = false, aiRequest = null, importToken = null, accepted = false;

  function post(type, payload) {
    if (window.chrome && window.chrome.webview) window.chrome.webview.postMessage({ type: type, payload: payload || {} });
  }
  function node(tag, className, text) {
    var el = document.createElement(tag);
    if (className) el.className = className;
    if (text != null) el.textContent = text;
    return el;
  }
  function button(text, action, title) {
    var el = node('button', 'btn btn-ghost', text);
    el.type = 'button';
    if (title) { el.title = title; el.setAttribute('aria-label', title); }
    el.addEventListener('click', action);
    return el;
  }
  function status(text, error) {
    $('skillStatus').textContent = text || '';
    $('skillStatus').hidden = !text;
    $('skillStatus').classList.toggle('is-error', !!error);
  }
  function setAiBusy(value) {
    aiBusy = !!value; $('skillAiSubmit').disabled = aiBusy; $('skillAiBack').disabled = aiBusy;
    $('skillAiSubmit').textContent = aiBusy ? 'Подготовка…' : 'Подготовить черновик'; $('skillAiCancel').hidden = !aiBusy;
  }
  function show(view) {
    ['skillCatalogView', 'skillEditor', 'skillAiForm', 'skillTransfer'].forEach(function(id) { $(id).hidden = id !== view; });
    status('');
  }
  function discard() { return !dirty || window.confirm('Закрыть черновик без сохранения?'); }
  function open() {
    $('skillSettings').hidden = false;
    document.body.classList.add('skills-open');
    if (!editing && !aiBusy) show('skillCatalogView');
    post('skills_request');
    $('skillClose').focus();
  }
  function close() {
    if (!discard()) return;
    if (aiBusy) post('skills_ai_cancel');
    aiRequest = null;
    setAiBusy(false);
    dirty = false; editing = null;
    $('skillSettings').hidden = true;
    document.body.classList.remove('skills-open');
    $('btnSkills').focus();
  }
  function back() {
    if (!discard()) return;
    dirty = false; editing = null;
    show('skillCatalogView');
    $('skillCreate').focus();
  }
  function date(value) { return value ? String(value).slice(0, 10) : ''; }
  function active(skill) {
    var now = new Date();
    var today = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + String(now.getDate()).padStart(2, '0');
    return skill.enabled && (!skill.valid_from || date(skill.valid_from) <= today) && (!skill.valid_until || date(skill.valid_until) >= today);
  }
  function render() {
    $('skillCategories').replaceChildren();
    Object.keys(categories).forEach(function(key) {
      var count = items.filter(function(item) { return item.category === key; }).length;
      var tab = button(categories[key] + ' · ' + count, function() { activeCategory = key; render(); });
      tab.classList.add('skill-category');
      tab.setAttribute('aria-pressed', String(activeCategory === key));
      $('skillCategories').appendChild(tab);
    });
    $('skillList').replaceChildren();
    var filtered = items.filter(function(item) { return item.category === activeCategory; });
    if (!filtered.length) $('skillList').appendChild(node('p', 'skill-empty', 'В этом разделе пока нет скиллов. Добавьте правило вручную или опишите его ИИ.'));
    filtered.forEach(function(skill) {
      var card = node('article', 'skill-card');
      var head = node('div', 'skill-card-heading');
      head.appendChild(node('strong', '', skill.name));
      head.appendChild(node('span', 'skill-state', !skill.enabled ? 'Выключен' : active(skill) ? 'Активен' : 'Вне периода'));
      card.appendChild(head);
      card.appendChild(node('p', 'skill-card-description', skill.description));
      var scope = [skill.project && 'Проект: ' + skill.project, skill.client && 'Клиент: ' + skill.client,
        skill.valid_until && 'До ' + date(skill.valid_until)].filter(Boolean).join(' · ');
      if (scope) card.appendChild(node('div', 'skill-card-scope', scope));
      var actions = node('div', 'skill-card-actions');
      var pin = button(pinned.indexOf(skill.id) >= 0 ? 'Закреплён' : 'Закрепить', function() {
        var next = pinned.filter(function(id) { return id !== skill.id; });
        if (pinned.indexOf(skill.id) < 0) next.push(skill.id);
        if (next.length > 5) { status('Можно закрепить до 5 скиллов.', true); return; }
        post('skills_pin', { ids: next });
      });
      pin.disabled = !active(skill);
      pin.setAttribute('aria-pressed', String(pinned.indexOf(skill.id) >= 0));
      actions.appendChild(pin);
      actions.appendChild(button('Изменить', function() { edit(skill); }));
      actions.appendChild(button('ИИ', function() { openAi(skill.id); }, 'Подготовить обновление скилла с ИИ'));
      actions.appendChild(button(skill.enabled ? 'Выключить' : 'Включить', function() {
        var changed = JSON.parse(JSON.stringify(skill)); changed.enabled = !changed.enabled;
        post('skills_save', { skill: changed, expected_revision: skill.revision });
      }));
      actions.appendChild(button('Удалить', function() {
        if (window.confirm('Удалить скилл «' + skill.name + '»?')) post('skills_delete', { id: skill.id, revision: skill.revision });
      }));
      card.appendChild(actions);
      $('skillList').appendChild(card);
    });
    $('pinnedSkills').replaceChildren();
    var names = pinned.map(function(id) { return items.find(function(skill) { return skill.id === id; }); }).filter(Boolean);
    $('pinnedSkills').hidden = !names.length;
    if (names.length) $('pinnedSkills').appendChild(button('Скиллы: ' + names.map(function(skill) { return skill.name; }).join(', '), open));
  }
  function personField(container, label, value, key, multiline) {
    var wrap = node('label', 'field-label', label);
    var input = node(multiline ? 'textarea' : 'input', 'field-input');
    input.value = value || ''; input.dataset.personField = key;
    if (multiline) input.rows = 2;
    if (key === 'email') { input.type = 'email'; input.required = true; }
    if (key === 'name') input.required = true;
    wrap.appendChild(input); container.appendChild(wrap);
  }
  function addPerson(person) {
    person = person || {};
    var row = node('fieldset', 'skill-person');
    row.appendChild(node('legend', '', person.name || 'Сотрудник'));
    personField(row, 'Имя', person.name, 'name');
    personField(row, 'Email — основной идентификатор', person.email, 'email');
    personField(row, 'Роль', person.role, 'role');
    personField(row, 'Другие имена и адреса (через запятую)', (person.aliases || []).join(', '), 'aliases');
    personField(row, 'Обязанности (по одной на строку)', (person.responsibilities || []).join('\n'), 'responsibilities', true);
    row.appendChild(button('Убрать сотрудника', function() { row.remove(); dirty = true; }));
    $('skillPeople').appendChild(row);
  }
  function edit(skill, draft) {
    if (!discard()) return;
    editing = JSON.parse(JSON.stringify(skill || { id: '', revision: 0, category: activeCategory, enabled: true, people: [], sources: [] }));
    show('skillEditor');
    ['name', 'description', 'instructions', 'project', 'client'].forEach(function(key) {
      $('skill' + key.charAt(0).toUpperCase() + key.slice(1)).value = editing[key] || '';
    });
    $('skillCategory').value = editing.category || activeCategory;
    $('skillTriggers').value = (editing.triggers || []).join(', ');
    $('skillValidFrom').value = date(editing.valid_from);
    $('skillValidUntil').value = date(editing.valid_until);
    $('skillEnabled').checked = editing.enabled !== false;
    $('skillPeople').replaceChildren();
    (editing.people || []).forEach(addPerson);
    $('skillSources').replaceChildren();
    if ((editing.sources || []).length) {
      var sources = node('details', ''); sources.appendChild(node('summary', '', 'Источники · ' + editing.sources.length));
      editing.sources.forEach(function(source) {
        sources.appendChild(node('p', '', [source.subject, source.from, date(source.date), source.note].filter(Boolean).join(' · ')));
      });
      $('skillSources').appendChild(sources);
    }
    $('skillDraftNote').replaceChildren(); $('skillDraftNote').hidden = !draft;
    if (draft) {
      $('skillDraftNote').appendChild(node('strong', '', 'Черновик ИИ — проверьте перед сохранением'));
      if (draft.reason) $('skillDraftNote').appendChild(node('p', '', draft.reason));
      (draft.uncertainties || []).forEach(function(text) { $('skillDraftNote').appendChild(node('p', '', text)); });
      var original = items.find(function(item) { return item.id === editing.id; });
      if (original) {
        var before = node('details', ''); before.appendChild(node('summary', '', 'Сохранённая версия для сравнения'));
        before.appendChild(node('pre', 'skill-original', original.instructions));
        (original.people || []).forEach(function(person) {
          before.appendChild(node('p', '', [person.name, person.email, person.role, (person.responsibilities || []).join('; ')].filter(Boolean).join(' · ')));
        });
        $('skillDraftNote').appendChild(before);
      }
    }
    dirty = !!draft; $('skillSave').disabled = false; $('skillName').focus();
  }
  function split(value, separator) { return (value || '').split(separator).map(function(text) { return text.trim(); }).filter(Boolean); }
  function save(event) {
    event.preventDefault();
    if (!$('skillEditor').reportValidity() || !editing) return;
    var skill = JSON.parse(JSON.stringify(editing));
    ['name', 'description', 'instructions', 'project', 'client'].forEach(function(key) {
      skill[key] = $('skill' + key.charAt(0).toUpperCase() + key.slice(1)).value.trim();
    });
    skill.category = $('skillCategory').value; skill.enabled = $('skillEnabled').checked;
    skill.triggers = split($('skillTriggers').value, /[,\n]+/);
    skill.valid_from = $('skillValidFrom').value || null; skill.valid_until = $('skillValidUntil').value || null;
    skill.people = Array.from($('skillPeople').children).map(function(row) {
      var person = {};
      row.querySelectorAll('[data-person-field]').forEach(function(input) { person[input.dataset.personField] = input.value.trim(); });
      person.aliases = split(person.aliases, /[,\n]+/); person.responsibilities = split(person.responsibilities, /\n+/);
      return person;
    });
    $('skillSave').disabled = true;
    post('skills_save', { skill: skill, expected_revision: editing.revision || 0 });
  }
  function openAi(id) {
    if (!discard()) return;
    dirty = false; editing = null; aiId = id || ''; show('skillAiForm');
    $('skillAiDescription').value = ''; $('skillAiDescription').focus();
    $('skillAiContext').checked = false;
  }
  function transfer(mode) {
    show('skillTransfer'); importToken = null;
    $('skillTransferText').readOnly = mode === 'export'; $('skillTransferText').value = '';
    $('skillImportFile').hidden = mode === 'export'; $('skillImportFile').value = '';
    $('skillTransferLabel').textContent = mode === 'export' ? 'Экспорт скиллов' : 'Импорт скиллов';
    $('skillImportCheck').hidden = mode === 'export'; $('skillDownload').hidden = mode !== 'export';
    $('skillImportApply').hidden = true; $('skillImportPreview').replaceChildren();
    if (mode === 'export') post('skills_export');
  }
  function receive(message) {
    if (!message) return;
    if (message.request_id && message.request_id !== aiRequest) return;
    if (message.type === 'catalog') {
      items = message.skills || []; pinned = message.pinned_ids || []; accepted = message.disclosure_accepted;
      $('skillNotice').hidden = accepted; render();
    } else if (message.type === 'error') {
      status(message.message, true); $('skillSave').disabled = false;
      if (message.request_id) setAiBusy(false);
    }
    else if (message.type === 'saved') { dirty = false; editing = null; show('skillCatalogView'); status('Скилл сохранён.'); }
    else if (message.type === 'draft') {
      if ($('skillSettings').hidden || !aiRequest) return;
      dirty = false; edit(message.draft.skill, message.draft);
    }
    else if (message.type === 'ai_busy') {
      setAiBusy(message.busy);
    } else if (message.type === 'ai_cancelled') status('Подготовка черновика отменена.');
    else if (message.type === 'export') $('skillTransferText').value = message.json || '';
    else if (message.type === 'imported') { show('skillCatalogView'); status('Импорт завершён.'); }
    else if (message.type === 'import_preview') {
      importToken = message.token; $('skillImportPreview').replaceChildren();
      $('skillImportPreview').appendChild(node('p', '', 'Новые скиллы будут добавлены. Отметьте только те совпадения, которые нужно заменить.'));
      (message.items || []).forEach(function(item) {
        var row = node('label', 'skill-import-row');
        if (item.conflict) { var check = node('input', ''); check.type = 'checkbox'; check.value = item.id; row.appendChild(check); }
        row.appendChild(node('span', '', item.name + (item.conflict ? ' — заменить сохранённый' : ' — новый')));
        $('skillImportPreview').appendChild(row);
      });
      $('skillImportApply').hidden = false;
    }
  }
  $('btnSkills').addEventListener('click', open); $('skillClose').addEventListener('click', close);
  $('skillRefresh').addEventListener('click', function() { post('skills_request'); });
  $('skillAcceptNotice').addEventListener('click', function() { post('skills_disclosure'); });
  $('skillCreate').addEventListener('click', function() { edit(null); });
  $('skillCreateAi').addEventListener('click', function() { openAi(''); });
  $('skillAddPerson').addEventListener('click', function() { addPerson(); dirty = true; });
  $('skillEditor').addEventListener('input', function() { dirty = true; });
  $('skillEditor').addEventListener('submit', save); $('skillEditorCancel').addEventListener('click', back);
  $('skillAiBack').addEventListener('click', back);
  $('skillAiCancel').addEventListener('click', function() { post('skills_ai_cancel'); });
  $('skillAiForm').addEventListener('submit', function(event) {
    event.preventDefault();
    if (aiBusy) return;
    if (!accepted) { status('Прочитайте уведомление выше и нажмите «Понятно».', true); return; }
    if (!$('skillAiForm').reportValidity()) return;
    aiRequest = 'draft_' + Date.now() + '_' + Math.random().toString(36).slice(2);
    setAiBusy(true);
    post('skills_ai_draft', { request_id: aiRequest, id: aiId, description: $('skillAiDescription').value, include_context: $('skillAiContext').checked });
  });
  $('skillImport').addEventListener('click', function() { transfer('import'); });
  $('skillExport').addEventListener('click', function() { transfer('export'); });
  $('skillTransferBack').addEventListener('click', back);
  $('skillImportCheck').addEventListener('click', function() { post('skills_import_preview', { json: $('skillTransferText').value }); });
  $('skillTransferText').addEventListener('input', function() { importToken = null; $('skillImportApply').hidden = true; });
  $('skillImportFile').addEventListener('change', function() {
    var file = this.files[0]; if (!file) return;
    if (file.size > 8 * 1024 * 1024) { status('Файл превышает 8 МБ.', true); return; }
    var reader = new FileReader();
    reader.onload = function() { $('skillTransferText').value = String(reader.result); importToken = null; $('skillImportApply').hidden = true; };
    reader.onerror = function() { status('Не удалось прочитать файл.', true); };
    reader.readAsText(file);
  });
  $('skillImportApply').addEventListener('click', function() {
    if (!importToken) return;
    var ids = Array.from($('skillImportPreview').querySelectorAll('input:checked')).map(function(input) { return input.value; });
    post('skills_import_apply', { token: importToken, replace_ids: ids });
  });
  $('skillDownload').addEventListener('click', function() {
    var url = URL.createObjectURL(new Blob([$('skillTransferText').value], { type: 'application/json;charset=utf-8' }));
    var link = node('a', ''); link.href = url; link.download = 'outlookai-skills.json'; link.click();
    window.setTimeout(function() { URL.revokeObjectURL(url); }, 1000);
  });
  document.addEventListener('keydown', function(event) {
    if (event.key === 'Escape' && !$('skillSettings').hidden) { event.preventDefault(); event.stopImmediatePropagation(); close(); }
    if (event.key === 'Tab' && !$('skillSettings').hidden) {
      var focusable = Array.from($('skillSettings').querySelectorAll('button, input, textarea, select, summary, [tabindex="0"]'))
        .filter(function(el) { return !el.disabled && el.getClientRects().length; });
      var first = focusable[0], last = focusable[focusable.length - 1];
      if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus(); }
      else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus(); }
    }
  });
  window.outlookSkills = {
    receive: receive,
    showProposal: function(draft) { open(); edit(draft.skill, draft); },
    attachProposal: function(container, result) {
      if (!result || !result.skill_proposal || !result.skill_proposal.skill) return;
      container.appendChild(button('Просмотреть обновление скилла', function() { window.outlookSkills.showProposal(result.skill_proposal); }));
    }
  };
})();

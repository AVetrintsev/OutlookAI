// Run with Node + Playwright; only a local fake WebView bridge is used. No AI requests.
const assert = require('node:assert/strict');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');

(async () => {
  const browser = await chromium.launch({ headless: true, channel: 'msedge' });
  const page = await browser.newPage({ viewport: { width: 380, height: 760 } });
  const errors = [];
  page.on('pageerror', error => errors.push(error.message));
  page.on('dialog', dialog => dialog.accept());
  await page.addInitScript(() => {
    window.hostMessages = [];
    window.chrome = window.chrome || {};
    window.chrome.webview = {
      postMessage: message => window.hostMessages.push(message),
      addEventListener: () => {}, removeEventListener: () => {}
    };
  });
  try {
    await page.goto(pathToFileURL(path.resolve(__dirname, '../../OutlookAI/WebUI/index.html')).href);
    const receive = message => page.evaluate(value => window.outlookSkills.receive(value), message);
    const posts = type => page.evaluate(value => window.hostMessages.filter(item => item.type === value), type);
    const skill = { id: 'people-1', revision: 3, category: 'people', name: 'Команда проекта',
      description: 'Ответственные за договоры и сроки согласования', instructions: 'Сопоставляйте сотрудников по email.', enabled: true,
      people: [{ name: 'Иван Иванов', email: 'ivan@example.com', role: 'Владелец договора', responsibilities: ['Согласовать договор'], aliases: [] }],
      project: 'Новый офис', sources: [{ subject: 'Назначение владельца', from: 'lead@example.com', date: '2026-09-05', note: 'Подтвердить область ответственности' }] };
    await receive({ type: 'catalog', skills: [skill], pinned_ids: [], disclosure_accepted: false });
    await page.locator('#btnSkills').click();
    assert.equal(await page.locator('#skillNotice').isVisible(), true);
    await page.locator('#skillCreateAi').click();
    await page.locator('#skillAiDescription').fill('Добавить правило согласования');
    await page.locator('#skillAiSubmit').click();
    assert.equal((await posts('skills_ai_draft')).length, 0, 'Disclosure gates AI requests');
    await page.locator('#skillAcceptNotice').click();
    assert.equal((await posts('skills_disclosure')).length, 1);
    await receive({ type: 'catalog', skills: [skill], pinned_ids: [], disclosure_accepted: true });
    await page.locator('#skillAiBack').click();
    await page.locator('#skillList button', { hasText: 'Закрепить' }).click();
    assert.deepEqual((await posts('skills_pin')).at(-1).payload.ids, [skill.id]);

    for (const theme of ['light', 'dark', 'high-contrast']) {
      for (const width of [320, 380, 600]) {
        await page.setViewportSize({ width, height: 760 });
        await page.evaluate(theme => { document.body.className = 'theme-' + theme + ' skills-open'; }, theme);
        const overflow = await page.evaluate(() => document.documentElement.scrollWidth > innerWidth);
        assert.equal(overflow, false, `${theme} / ${width}px must not scroll horizontally`);
      }
      await page.setViewportSize({ width: 380, height: 760 });
      await page.screenshot({ path: path.resolve(__dirname, '../../../out/validation/skills-' + theme + '.png') });
    }
    await page.locator('#skillList button', { hasText: 'Изменить' }).click();
    await page.setViewportSize({ width: 320, height: 640 });
    assert.equal(await page.evaluate(() => document.documentElement.scrollWidth > innerWidth), false);
    await page.locator('#skillSave').scrollIntoViewIfNeeded();
    await page.screenshot({ path: path.resolve(__dirname, '../../../out/validation/skills-editor.png') });
    await page.setViewportSize({ width: 380, height: 760 });
    await page.locator('#skillName').fill('Команда договоров');
    await page.locator('#skillSave').click();
    const save = (await posts('skills_save')).at(-1).payload;
    assert.equal(save.expected_revision, 3);
    assert.equal(save.skill.name, 'Команда договоров');
    assert.equal(save.skill.people[0].email, 'ivan@example.com');
    await receive({ type: 'error', message: 'Конфликт версии' });
    assert.equal(await page.locator('#skillSave').isEnabled(), true);
    assert.equal(await page.locator('#skillName').inputValue(), 'Команда договоров', 'Conflict preserves edits');
    await page.locator('#skillEditorCancel').click();

    await page.locator('#skillImport').click();
    await page.locator('#skillTransferText').fill('{"schema_version":1,"skills":[]}');
    await page.locator('#skillImportCheck').click();
    await receive({ type: 'import_preview', token: 'preview-1', items: [{ id: skill.id, name: skill.name, conflict: true }, { id: 'new', name: 'Новое правило', conflict: false }] });
    assert.equal(await page.locator('#skillImportPreview input:checked').count(), 0, 'Conflict replacements are opt-in');
    await page.locator('#skillImportApply').click();
    assert.deepEqual((await posts('skills_import_apply')).at(-1).payload.replace_ids, []);
    await page.locator('#skillTransferText').fill('{}');
    assert.equal(await page.locator('#skillImportApply').isVisible(), false, 'Changing source invalidates preview');
    await page.locator('#skillTransferBack').click();

    await page.locator('#skillCreateAi').click();
    await page.locator('#skillAiDescription').fill('Уточнить, кто ведёт договор');
    await page.locator('#skillAiSubmit').click();
    const request = (await posts('skills_ai_draft')).at(-1).payload;
    assert.equal(request.include_context, false, 'Mail context is opt-in per request');
    await receive({ type: 'draft', request_id: request.request_id, draft: { skill, reason: 'В письме указан новый владелец', uncertainties: ['Проверить дату начала роли'] } });
    assert.equal(await page.locator('#skillDraftNote').isVisible(), true);
    assert.equal((await posts('skills_save')).length, 1, 'AI cannot save its draft automatically');
    await receive({ type: 'ai_busy', request_id: request.request_id, busy: false });
    await page.locator('#skillEditorCancel').click();
    await page.locator('#skillCreateAi').click();
    await page.locator('#skillAiDescription').fill('Ещё одно правило');
    await page.locator('#skillAiSubmit').click();
    const cancelled = (await posts('skills_ai_draft')).at(-1).payload;
    await page.locator('#skillClose').click();
    await receive({ type: 'draft', request_id: cancelled.request_id, draft: { skill, uncertainties: [] } });
    assert.equal(await page.locator('#skillSettings').isVisible(), false, 'Late result cannot reopen settings');
    await page.locator('#btnSkills').click();
    assert.equal(await page.locator('#skillCatalogView').isVisible(), true);

    await receive({ type: 'catalog', skills: [{ ...skill, name: '<img src=x onerror="window.xss=1">' }], pinned_ids: [], disclosure_accepted: true });
    assert.equal(await page.locator('#skillList img').count(), 0);
    assert.equal(await page.evaluate(() => window.xss), undefined);
    assert.deepEqual(errors, []);
    console.log('PASS: skills UI — disclosure, pins, editing, import conflicts, AI drafts/cancellation, XSS, 9 theme/width combinations');
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });

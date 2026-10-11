/* 回歸測試（安全性）：
     S1 案件檔夾帶 HTML／程式碼：開啟、自動保存還原、構件快照皆不得執行，載重值清理為數字
     S2 圖說匯入：Gemini 回傳之字串一律當文字顯示（以 Playwright 攔截 generativelanguage 端點）
     S3 構件總表 CSV：= + - @ 開頭之儲存格前置單引號（公式注入） */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs'), os = require('os'), path = require('path');
const {openApp} = require('./helpers');

let app, page;
test.before(async () => { app = await openApp(); page = app.page; });
test.after(async () => { if(app){ assert.deepEqual(app.errors, [], '頁面不得有 JavaScript 錯誤'); await app.close(); } });

const XSS = '1"><img src=x onerror="window.__xss=(window.__xss||0)+1">';
const pwned = () => page.evaluate(() => window.__xss || 0);
const waitApp = () => page.waitForFunction(() => typeof MODEL !== 'undefined' && MODEL && typeof applyPreset === 'function');

test('S1：惡意案件檔不執行、載重值清理為數字，自動保存重新載入後亦不執行', async () => {
  const c = await page.evaluate(() => { document.getElementById('tabPylon').click(); applyPreset('bldg60'); return caseSnapshot(); });
  c.loads = [{name: XSS, pos: '"><img src=x onerror="window.__xss=1">', Pu: XSS, Mux: '12.5', Muy: {}, Vux: 'NaN', Vuy: Infinity, Tu: 0, evil: XSS}];
  c.bloads = [{name: 7, Mpos: XSS, Mneg: '3', Vu: {}, Tu: null}];
  c.members = [{id: XSS, name: XSS, tab: XSS, saved: XSS, snap: {...c, members: undefined, _member: true,
    loads: [{name: 'm', Pu: XSS, Mux: 1}], bloads: [{name: 'b', Mu: '4', Mneg: XSS}]}}];
  const f = path.join(os.tmpdir(), `rcsd-xss-${process.pid}.json`);
  fs.writeFileSync(f, JSON.stringify(c));
  await page.setInputFiles('#caseFile', f);
  // H9：有未存檔變更時先詢問（點「直接開啟」）；構件清單詢問時以案件檔取代
  await page.waitForFunction(() => { const ov = document.getElementById('askOv');
    if(ov.classList.contains('open')){ const b = [...document.querySelectorAll('#askAct button')].find(x => /直接開啟|以案件檔取代/.test(x.textContent)); if(b) b.click(); }
    return LOADS.length === 1 && LOADS[0].Mux === 12.5; }, null, {polling: 100});
  fs.unlinkSync(f);
  await page.evaluate(() => { document.getElementById('tabBeam').click(); drawBLoads(); document.getElementById('tabPylon').click(); });
  await page.waitForTimeout(300);
  assert.equal(await pwned(), 0, '開啟案件檔不得執行夾帶的程式碼');
  const s = await page.evaluate(() => ({L: LOADS[0], B: BLOADS[0], M: MEMBERS[0],
    imgs: document.querySelectorAll('#loadTbl img, #bLoadTbl img').length,
    pu: document.querySelector('#loadTbl input[data-k="Pu"]').value, mpos: document.querySelector('#bLoadTbl input[data-k="Mpos"]').value,
    name: document.querySelector('#loadTbl input[data-k="name"]').value}));
  assert.deepEqual(s.L, {name: XSS, Pu: 0, Mux: 12.5, Muy: 0, Vux: 0, Vuy: 0, Tu: 0}, '數值欄 Number() 化、非有限值歸零、未知欄位與非法位置移除');
  assert.deepEqual(s.B, {name: '7', Mpos: 0, Mneg: 3, Vu: 0, Tu: 0}, '名稱強制為字串');
  assert.equal(s.imgs, 0); assert.equal(s.pu, '0'); assert.equal(s.mpos, '0'); assert.equal(s.name, XSS, '名稱照原文當文字顯示');
  assert.equal(typeof s.M.id, 'string'); assert.equal(s.M.tab, 'pylon');
  assert.deepEqual(s.M.snap.loads[0], {name: 'm', Pu: 0, Mux: 1, Muy: 0, Vux: 0, Vuy: 0, Tu: 0}, '構件快照之載重同樣清理');
  assert.deepEqual(s.M.snap.bloads[0], {name: 'b', Mpos: 4, Mneg: 0, Vu: 0, Tu: 0}, '舊版 Mu 對應 Mpos');

  // 開啟構件、構件清單顯示
  await page.evaluate(() => { memRender(); memOpen(0); memRender(); });
  assert.equal(await pwned(), 0, '構件清單與開啟構件不得執行');

  // 直接寫入未清理的自動保存與構件清單 → 重新載入
  await page.evaluate(([c, XSS]) => {
    const raw = {...c, members: []}; raw.loads = [{name: 'x', Pu: XSS, Mux: 0, Muy: 0, Vux: 0, Vuy: 0, Tu: 0}];
    raw.bloads = [{name: 'y', Mpos: XSS, Mneg: 0, Vu: 0, Tu: 0}];
    CASE_READY = false; clearTimeout(CASE_T);
    localStorage.setItem(CASE_KEY, JSON.stringify(raw));
    localStorage.setItem(MEM_KEY, JSON.stringify([{id: XSS, name: XSS, tab: 'pylon', saved: XSS, snap: raw}]));
  }, [c, XSS]);
  await page.reload(); await waitApp();
  await page.evaluate(() => { document.getElementById('setupSkip') && document.getElementById('setupSkip').click(); document.getElementById('tabBeam').click(); drawBLoads(); memRender(); });
  await page.waitForTimeout(500);
  assert.equal(await pwned(), 0, '自動保存還原不得執行');
  const r = await page.evaluate(() => ({pu: LOADS[0].Pu, mpos: BLOADS[0].Mpos, mem: MEMBERS[0].snap.loads[0].Pu}));
  assert.deepEqual(r, {pu: 0, mpos: 0, mem: 0});
  await page.evaluate(() => { localStorage.removeItem(CASE_KEY); localStorage.removeItem(MEM_KEY); MEMBERS = []; document.getElementById('tabPylon').click(); applyPreset('bldg60'); });
});

test('S2：Gemini 判讀結果中的 HTML 一律以文字顯示', async () => {
  const H = s => `<img src=x onerror="window.__xss=(window.__xss||0)+1">${s}`;
  const res = {sheet_note: H('sheet'), unit_note: H('unit'), sections: [{
    name: H('name'), member_type: H('type'), secType: 'rect', B: '100"><img src=x onerror="window.__xss=9">', H: 120, D: null,
    tie_type: 'hoop', tw: null, covO: 5, covI: null, fc: 350, fy: 4200, fyt: 4200, dbl: '0',
    barSize_as_drawn: H('D32'), tieSize_as_drawn: H('D13'), n_bars_total: 24, n_bars_source: H('src'),
    bars_note_as_drawn: H('note'), low_confidence: [], evidence: H('evidence')}]};
  let url = '';
  await page.route(/generativelanguage\.googleapis\.com/, r => { url = r.request().url();
    return r.fulfill({status: 200, contentType: 'application/json',
      body: JSON.stringify({candidates: [{content: {parts: [{text: JSON.stringify(res)}]}}], usageMetadata: {promptTokenCount: H('1')}})}); });
  await page.evaluate(() => document.getElementById('btnImport').click());
  await page.setInputFiles('#impFile', {name: 'a.png', mimeType: 'image/png', buffer: Buffer.from('89504e47', 'hex')});
  await page.fill('#impKey', 'k'); await page.fill('#impModel', 'gemini-x.1'); await page.check('#impConsent');
  await page.click('#impGo');
  await page.waitForFunction(() => document.getElementById('impResult').style.display !== 'none');
  await page.waitForTimeout(300);
  assert.equal(await pwned(), 0, '模型輸出不得執行');
  const o = await page.evaluate(() => ({imgs: document.querySelectorAll('#impResult img, #impMsg img').length,
    text: document.getElementById('impResult').textContent}));
  assert.equal(o.imgs, 0, '不得產生任何 <img>');
  for(const k of ['sheet', 'unit', 'name', 'type', 'evidence', 'D32', 'src', 'note']) assert.ok(o.text.includes(H(k)), `「${k}」以原文文字顯示`);
  assert.match(url, /\/models\/gemini-x\.1:generateContent$/);
  await page.unroute(/generativelanguage\.googleapis\.com/);
  await page.evaluate(() => document.getElementById('impClose').click());
});

test('S3：構件總表 CSV 中 = + - @ 開頭之儲存格前置單引號', async () => {
  const names = ['=HYPERLINK("http://x","y")', '+1+1', '-2+3', '@SUM(A1)', '一般柱'];
  await page.evaluate(names => { MEMBERS = []; document.getElementById('tabPylon').click(); applyPreset('bldg60');
    for(const n of names){ document.getElementById('memName').value = n; memAdd(); } }, names);
  const [dl] = await Promise.all([page.waitForEvent('download', {timeout: 60000}), page.evaluate(() => memCsv())]);
  const csv = fs.readFileSync(await dl.path(), 'utf8').replace(/^﻿/, '');
  const first = csv.split('\r\n').slice(1).map(l => l.startsWith('"') ? l.slice(1).split('"')[0] : l.split(',')[0]);
  assert.deepEqual(first, ["'=HYPERLINK(", "'+1+1", "'-2+3", "'@SUM(A1)", '一般柱']);
  assert.ok(csv.includes(`"'=HYPERLINK(""http://x"",""y"")"`), '含引號之儲存格以雙引號包覆並跳脫');
  const cells = await page.evaluate(() => [csvCell('=1'), csvCell(-1.5), csvCell('a,b'), csvCell('\t1'), csvCell('ok')]);
  assert.deepEqual(cells, ["'=1", "'-1.5", '"a,b"', "'\t1", 'ok']);
  await page.evaluate(() => { MEMBERS = []; memSave(); memRender(); });
});

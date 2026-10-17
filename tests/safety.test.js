/* 簽證安全機制回歸（驗證與安全性流程手冊第八章）：
     S-1 D/C 顯示無條件進位
     S-2 超出驗證範圍（f'c、f_y、f_yt、輕質）列「需確認」，計算書標示
     S-3 「從圖說匯入」填入且未修改之欄位列「需確認」，修改或套用範本後解除
     S-4 主要結果 NaN 判不合格
     S-5 有側移 δ_s 以本柱代替樓層（未輸入 Q）列「需確認」，輸入 Q 後為參考
     S-6 水工模式之 S_d 列「未涵蓋」
     S-7 計算書「適用範圍、驗證與簽證責任」、工程判斷與簽證技師欄（網頁與 Excel） */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs'), os = require('os'), path = require('path');
const ExcelJS = require('exceljs');
const {openApp} = require('./helpers');

let app, page;
test.before(async () => { app = await openApp(); page = app.page; });
test.after(async () => { if(app){ assert.deepEqual(app.errors, [], '頁面不得有 JavaScript 錯誤'); await app.close(); } });

const state = () => page.evaluate(() => ({
  conf: CONF_ITEMS.pylon.map(q => q.t.replace(/<[^>]+>/g, '')),
  items: colAlertItems(MODEL).map(a => ({lv: a.lv, t: a.t.replace(/<[^>]+>/g, '')}))}));
// 範本不重設 f_y、λ、環境：每次先還原為預設再改
const col = (inputs = {}) => page.evaluate(inputs => { document.getElementById('tabPylon').click(); applyPreset('bldg60');
  for(const [k, v] of Object.entries({fy: 4200, fyt: 4200, dLight: 1, envWater: 'normal', ...inputs})) document.getElementById(k).value = v; render(); }, inputs);

test('S-1：D/C 顯示無條件進位至小數 3 位', async () => {
  const r = await page.evaluate(() => [fdc(1.0004), fdc(0.5151), fdc(0.5), fdc(0.9999), fdc(1), fdc(Infinity)]);
  assert.deepEqual(r, ['1.001', '0.516', '0.500', '1.000', '1.000', '∞']);
});

test('S-2：超出驗證範圍列需確認，計算書標示', async () => {
  await col();
  let s = await state();
  assert.ok(!s.conf.some(t => /超出驗證案例涵蓋範圍/.test(t)), 'f′c 280、f_y 4200 在範圍內');
  for(const [k, v, re] of [['fc', 490, /f'c = 490/], ['fc', 175, /f'c = 175/], ['fy', 5600, /主筋 f_y = 5,600/], ['dLight', 0.75, /輕質/]]){
    await col({[k]: v}); s = await state();
    const t = s.conf.find(x => /超出驗證案例涵蓋範圍/.test(x));
    assert.ok(t && re.test(t), `${k} = ${v} 應列需確認（${t}）`);
  }
  await col({fc: 490});
  const rp = await page.evaluate(() => { buildReport(MODEL); return document.getElementById('report').textContent; });
  assert.match(rp, /本構件超出驗證範圍/);
  await col();
  const rp2 = await page.evaluate(() => { buildReport(MODEL); return document.getElementById('report').textContent; });
  assert.match(rp2, /本構件在驗證範圍內/);
});

test('S-3：圖說匯入（AI 判讀）之欄位未修改前列需確認', async () => {
  await col();
  await page.evaluate(() => { IMP_SRC.set('B', document.getElementById('B').value); IMP_SRC.set('covO', document.getElementById('covO').value); render(); });
  let s = await state(), t = s.conf.find(x => /從圖說匯入/.test(x));
  assert.ok(t && /外寬/.test(t) && /外側保護層/.test(t), `兩欄應列名（${t}）`);
  await page.evaluate(() => { document.getElementById('B').value = 65; render(); });
  s = await state(); t = s.conf.find(x => /從圖說匯入/.test(x));
  assert.ok(t && !/外寬/.test(t) && /外側保護層/.test(t), `修改後之外寬不再列入（${t}）`);
  await page.evaluate(() => applyPreset('bldg60'));
  s = await state();
  assert.ok(!s.conf.some(x => /從圖說匯入/.test(x)), '套用範本後解除');
  // 梁模組之欄位不影響柱
  await page.evaluate(() => { IMP_SRC.set('bw', document.getElementById('bw').value); render(); });
  s = await state();
  assert.ok(!s.conf.some(x => /從圖說匯入/.test(x)), '梁欄位不列於柱');
  await page.evaluate(() => IMP_SRC.clear());
});

test('S-4：主要結果 NaN 判不合格', async () => {
  await col();
  const r = await page.evaluate(() => { const A = []; safetyAlerts('pylon', MODEL.m, {'Y 向 V_s 需求': NaN, 'D/C': 0.5, '無限': Infinity}, (lv, t) => A.push({lv, t}));
    return A.filter(a => /數值計算異常/.test(a.t)); });
  assert.equal(r.length, 1); assert.equal(r[0].lv, 'bad'); assert.match(r[0].t, /Y 向 V_s 需求/); assert.doesNotMatch(r[0].t, /無限/);
});

test('S-5：有側移 δ_s 以本柱代替樓層 → 需確認；輸入 Q 後為參考', async () => {
  await col({lu: 600, slK: 1.5, slKy: 1.5});
  let s = await state();
  assert.ok(s.conf.some(t => /工程判斷（須確認）有側移構架/.test(t)), `未輸入 Q 時列需確認（${s.conf.join(' / ')}）`);
  await col({lu: 600, slK: 1.5, slKy: 1.5, slQ: 0.1});
  s = await state();
  assert.ok(!s.conf.some(t => /δs? ?以本柱/.test(t)), '輸入 Q 後不列需確認');
  assert.ok(s.items.some(a => a.lv === 'info' && /Q 法/.test(a.t)), '改列 Q 法說明');
  await col({lu: 600, slK: 1.5, slKy: 1.5});
  const rp = await page.evaluate(() => { buildReport(MODEL); return document.getElementById('report').textContent; });
  assert.match(rp, /工程判斷：\(1\) 有側移 δs? ?以本柱/);
});

test('S-6：水工模式之 S_d 列未涵蓋（柱、梁）', async () => {
  await col({envWater: 'water'});
  let s = await state();
  assert.ok(s.conf.some(t => /須另行檢核：ACI 350 環境耐久係數 S_d/.test(t)), '柱');
  const b = await page.evaluate(() => { document.getElementById('tabBeam').click(); applyBPreset('pondWall'); renderBeam();
    return CONF_ITEMS.beam.map(q => q.t.replace(/<[^>]+>/g, '')); });
  assert.ok(b.some(t => /S_d/.test(t)), '版');
  await page.evaluate(() => { document.getElementById('envWater').value = 'normal'; document.getElementById('tabPylon').click(); });
});

test('S-7：計算書簽證責任節、工程判斷與簽證技師欄（網頁、Excel）', async () => {
  await col();
  const rp = await page.evaluate(() => { buildReport(MODEL); const el = document.getElementById('report');
    return {t: el.textContent, sign: el.querySelector('.rp-sign').textContent}; });
  assert.match(rp.t, /適用範圍、驗證與簽證責任/);
  assert.match(rp.t, /簽證技師確認事項/);
  assert.match(rp.t, /箍筋肢橫向間距參考 ACI 318-19/, '矩形柱列工程判斷');
  assert.match(rp.sign, /簽證技師：.*科別：.*執業執照字號：.*簽章：/);
  const bm = await page.evaluate(() => { document.getElementById('tabBeam').click(); applyBPreset('pondWall'); renderBeam(); buildBeamReport(BMODEL);
    const t = document.getElementById('report').textContent; document.getElementById('tabPylon').click(); return t; });
  assert.match(bm, /適用範圍、驗證與簽證責任/); assert.match(bm, /Gergely-Lutz/);
  // Excel：簽證責任段落與簽證技師列
  await col();
  await page.evaluate(() => confAll('pylon'));
  const [dl] = await Promise.all([page.waitForEvent('download', {timeout: 120000}), page.click('#btnXlsx')]);
  const f = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'rcsd-sf-')), 'c.xlsx'); await dl.saveAs(f);
  const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(f);
  let txt = ''; wb.eachSheet(ws => ws.eachRow(r => r.eachCell(c => { const v = c.value; txt += (v && typeof v === 'object' ? (v.richText ? v.richText.map(x => x.text).join('') : '') : String(v ?? '')) + '\n'; })));
  assert.match(txt, /簽證責任：本工具為斷面層級之輔助計算/);
  assert.match(txt, /簽證技師：.*執業執照字號/);
});

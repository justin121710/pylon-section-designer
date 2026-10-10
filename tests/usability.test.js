/* 回歸測試（M9、U1～U8）：
     M9 建築柱 V_c 歸零依 §18.4.6.2.1 (a)(b)
     U1 判定三級：需確認項目未勾選前不得列印
     U2 1280×720 參數表至少 10 列、說明卡可收合
     U3 多構件清單：加入、批次總表、開啟還原
     U4 計算書／案件檔記錄版本與雜湊；舊案件缺欄位提示；案件檔保存斷面型式與構件名稱
     U5 圖說匯入：金鑰預設不記住、上傳前須同意外傳
     U6 函式庫自 vendor/ 載入，不連 CDN
     U7 開場精靈「不再顯示」、狀態列直接切換設計依據
     U8 載重表 SI 輸入與單位防呆 */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const {openApp} = require('./helpers');

let app, page;
test.before(async () => { app = await openApp(); page = app.page; });
test.after(async () => { if(app){ assert.deepEqual(app.errors, [], '頁面不得有 JavaScript 錯誤'); await app.close(); } });
const near = (a, b, rel, msg) => assert.ok(Math.abs(a - b) <= rel*Math.abs(b), `${msg}：實際 ${a}，期望 ${b}`);

test('M9：V_c 歸零須 (a) V_e ≥ ½V_des 且 (b) P_u < A_gf′c/20（§18.4.6.2.1）', async () => {
  const run = (Vuy, rule) => page.evaluate(([Vuy, rule]) => {
    document.getElementById('tabPylon').click(); applyPreset('bldg60'); document.getElementById('vcRule').value = rule;
    LOADS = [{name:'A', Pu:30, Mux:20, Muy:0, Vux:0, Vuy, Tu:0}]; drawLoads(); render();
    const c = MODEL.shY.all[0]; return {zeroed: c.V.zeroed, Ve: MODEL.shY.Ve, Vdes: c.Vdes, lim: 0.05*MODEL.m.fc*MODEL.m.Ag};
  }, [Vuy, rule]);
  let r = await run(10, 'ab');
  assert.ok(30e3 < r.lim); assert.ok(r.Ve >= 0.5*r.Vdes); assert.equal(r.zeroed, true, '地震剪力控制且低軸壓：V_c = 0');
  r = await run(200, 'ab');
  assert.ok(r.Ve < 0.5*r.Vdes, '非地震剪力控制'); assert.equal(r.zeroed, false, '(a) 不成立：保留 V_c');
  r = await run(200, 'b');
  assert.equal(r.zeroed, true, '選「僅依軸力」（保守）時歸零');
  await page.evaluate(() => { document.getElementById('vcRule').value = 'ab'; applyPreset('bldg60'); });
});

test('U1：需確認項目未勾選時不得列印；全部確認後可列印', async () => {
  await page.evaluate(() => { window.__printed = 0; window.print = () => { window.__printed++; }; document.getElementById('tabPylon').click(); applyPreset('bldg60'); });
  const n = await page.evaluate(() => ({all: CONF_ITEMS.pylon.length, pend: confPending('pylon').length, vd: document.getElementById('verdict').innerText}));
  assert.ok(n.all > 0 && n.pend === n.all); assert.match(n.vd, /需確認/);
  await page.click('#btnPrint'); await page.waitForTimeout(400);
  const d = await page.evaluate(() => ({open: document.getElementById('askOv').classList.contains('open'), t: document.getElementById('askTtl').textContent}));
  assert.ok(d.open && /需確認/.test(d.t), '應跳出「尚有需確認項目」');
  await page.evaluate(() => document.querySelector('#askAct button').click());       // 取消
  await page.evaluate(() => confAll('pylon'));
  await page.click('#btnPrint'); await page.waitForTimeout(800);
  assert.equal(await page.evaluate(() => window.__printed), 1, '全部確認後應可列印');
  // 換構件後須重新確認
  assert.ok(await page.evaluate(() => { applyPreset('pond80'); return confPending('pylon').length > 0; }));
});

test('U2：1280×720 時參數表至少顯示 10 列，說明卡可收合', async () => {
  await page.setViewportSize({width:1280, height:720}); await page.waitForTimeout(300);
  const r = await page.evaluate(() => { document.getElementById('tabPylon').click();
    const h = document.querySelector('#sheetP .pin-body').getBoundingClientRect().height; return Math.floor(h/39); });
  assert.ok(r >= 10, `只顯示 ${r} 列`);
  await page.evaluate(() => document.querySelector('#detailP .det-tg').click());
  assert.ok(await page.evaluate(() => document.querySelector('#app .pin-mid').classList.contains('det-off')));
  await page.evaluate(() => document.querySelector('#detailP .det-tg').click());
  await page.setViewportSize({width:1440, height:900});
});

test('U3：構件清單加入、批次總表與開啟還原', async () => {
  const r = await page.evaluate(async () => {
    MEMBERS = []; MEM_RES = {};
    document.getElementById('tabPylon').click();
    applyPreset('bldg60'); document.getElementById('memName').value = 'C1'; render(); memAdd();
    applyPreset('pier100'); document.getElementById('memName').value = 'P1'; render(); memAdd();
    document.getElementById('tabBeam').click(); applyBPreset('deckT'); document.getElementById('bMemName').value = 'G1'; renderBeam(); memAdd();
    await memBatch();
    const res = MEMBERS.map(m => [m.name, m.tab, +MEM_RES[m.id].dc.toFixed(3)]);
    const back = [TAB, document.getElementById('bMemName').value];
    memOpen(0); const opened = [TAB, document.getElementById('memName').value, document.getElementById('secType').value, MODEL.s.B];
    return {res, back, opened, n: MEMBERS.length};
  });
  assert.equal(r.n, 3);
  assert.deepEqual(r.res.map(x => x[0]), ['C1', 'P1', 'G1']);
  assert.ok(r.res.every(x => Number.isFinite(x[2]) && x[2] > 0));
  assert.deepEqual(r.back, ['beam', 'G1'], '批次計算後回到原構件');
  assert.deepEqual(r.opened, ['pylon', 'C1', 'solid', 60], '開啟構件還原其名稱、斷面型式與尺寸');
  await page.evaluate(() => { MEMBERS = []; MEM_RES = {}; });
});

test('U4：計算書記錄版本與雜湊；舊案件缺欄位提示；案件檔保存斷面型式與構件名稱', async () => {
  const r = await page.evaluate(() => { document.getElementById('tabPylon').click(); applyPreset('box300'); document.getElementById('memName').value = 'T1'; render();
    buildReport(MODEL); const t = document.getElementById('report').textContent;
    const snap = caseSnapshot(), ih = inputHash();
    applyPreset('bldg60'); document.getElementById('memName').value = ''; render();
    caseApply(JSON.parse(JSON.stringify(snap)));
    const restored = [document.getElementById('secType').value, document.getElementById('memName').value, inputHash() === ih];
    const old = JSON.parse(JSON.stringify(snap)); delete old.values.vcRule; delete old.appVersion; old.ver = 1; caseApply(old);
    const notice = CASE_NOTICE; CASE_NOTICE = null;
    return {t, ver: APP_VERSION, ch: codeHash(), snapVer: snap.appVersion, restored, notice};
  });
  assert.ok(r.t.includes('v' + r.ver) && r.t.includes(r.ch), '計算書含版本與程式雜湊');
  assert.match(r.t, /輸入 SHA-256 [0-9a-f]{16}/);
  assert.equal(r.snapVer, r.ver);
  assert.deepEqual(r.restored, ['box', 'T1', true], '開啟案件還原斷面型式、構件名稱，輸入雜湊不變');
  assert.ok(r.notice && r.notice.fields.some(x => /V_c 歸零/.test(x)), '舊案件應列出採預設值之欄位');
});

test('U5：圖說匯入預設不記住金鑰，上傳前須勾選同意外傳', async () => {
  await page.evaluate(() => { try{ localStorage.removeItem('pylon_gemini_key'); }catch(e){} document.getElementById('btnImport').click(); });
  const st = await page.evaluate(() => ({rem: document.getElementById('impRemember').checked, consent: document.getElementById('impConsent').checked, model: document.getElementById('impModel').tagName}));
  assert.deepEqual(st, {rem: false, consent: false, model: 'INPUT'});
  await page.setInputFiles('#impFile', {name: 'a.png', mimeType: 'image/png', buffer: Buffer.from('89504e470d0a1a0a', 'hex')});
  let reqs = 0; const onReq = q => { if(/generativelanguage/.test(q.url())) reqs++; }; page.on('request', onReq);
  await page.fill('#impKey', 'TESTKEY'); await page.click('#impGo'); await page.waitForTimeout(300);
  page.off('request', onReq);
  const msg = await page.evaluate(() => document.getElementById('impOv').innerText);
  assert.match(msg, /資料外傳告知/); assert.equal(reqs, 0, '未同意前不得上傳');
  assert.equal(await page.evaluate(() => localStorage.getItem('pylon_gemini_key')), null, '金鑰不得預設保存');
  await page.evaluate(() => document.getElementById('impOv').classList.remove('open'));
});

test('U6、U7：函式庫自 vendor/ 載入；「不再顯示」記住；狀態列直接切換設計依據', async () => {
  const cdn = [], onReq = q => { if(/cdn\.jsdelivr|unpkg|cdnjs/.test(q.url())) cdn.push(q.url()); };
  page.on('request', onReq);
  await page.evaluate(() => localStorage.setItem('rcsd.setup.noShow', '1'));
  await page.reload(); await page.waitForFunction(() => typeof MODEL !== 'undefined' && MODEL && window.katex && window.renderMathInElement);
  await page.evaluate(() => ensureXlsxLib());
  page.off('request', onReq);
  assert.deepEqual(cdn, [], '不得向 CDN 請求函式庫');
  assert.equal(await page.evaluate(() => document.getElementById('setupOv').classList.contains('open')), false, '設定不再顯示後不跳出精靈');
  const r = await page.evaluate(() => { document.getElementById('tabPylon').click(); applyPreset('bldg60');
    const sb = document.getElementById('sbCode'); sb.value = 'bridge'; sb.dispatchEvent(new Event('change'));
    const out = [document.getElementById('codeBasis').value, MODEL.opt.code]; sb.value = 'bldg'; sb.dispatchEvent(new Event('change')); return out; });
  assert.deepEqual(r, ['bridge', 'bridge']);
  await page.evaluate(() => localStorage.removeItem('rcsd.setup.noShow'));
});

test('U8：載重表 SI 輸入（kN、kN·m）換算為 tf；CSV 依單位解讀；單位防呆', async () => {
  const r = await page.evaluate(() => { document.getElementById('tabPylon').click(); applyPreset('bldg60');
    const u = document.getElementById('loadUnit'); u.value = 'kN'; u.dispatchEvent(new Event('change'));
    const shown = +document.querySelector('#loadTbl tbody input[data-k="Pu"]').value;
    const inp = document.querySelector('#loadTbl tbody input[data-k="Mux"]'); inp.value = '98.0665'; inp.dispatchEvent(new Event('input'));
    const mux = LOADS[0].Mux;
    const c = parseLoadCsv('A,980.665,98.0665,0,0,0,0', ['Pu','Mux','Muy','Vux','Vuy','Tu'], {pos:true}).ok;
    document.getElementById('csvBox').value = 'A,980.665,98.0665,0,0,0,0';
    return {shown, pu0: PRESETS.bldg60.loads[0].Pu, mux, head: document.querySelector('#loadTbl thead').innerText};
  });
  near(r.shown, r.pu0*9.80665, 1e-6, 'P_u 以 kN 顯示'); near(r.mux, 10, 1e-9, '98.0665 kN·m = 10 tf·m');
  assert.match(r.head, /kN·m/);
  await page.evaluate(() => { applyLoadCsv('pylon'); }); await page.waitForTimeout(300);
  await page.evaluate(() => { const b = [...document.querySelectorAll('#askAct button')].find(x => /取代/.test(x.textContent)); if(b) b.click(); });
  await page.waitForTimeout(300);
  const c = await page.evaluate(() => LOADS.map(L => [L.Pu, L.Mux]));
  near(c[0][0], 100, 1e-9, 'CSV P_u（kN → tf）'); near(c[0][1], 10, 1e-9, 'CSV M_ux');
  const w = await page.evaluate(() => { const u = document.getElementById('loadUnit'); u.value = 'tf'; u.dispatchEvent(new Event('change'));
    LOADS = [{name:'kN 誤當 tf', Pu: 2600*9.80665, Mux: 18*9.8, Muy: 0, Vux: 0, Vuy: 0, Tu: 0}]; drawLoads(); render();
    return colAlertItems(MODEL).some(a => /疑似單位錯誤/.test(a.t)); });
  assert.ok(w, '軸力遠超過 P_o 應提醒疑似單位錯誤');
  await page.evaluate(() => applyPreset('bldg60'));
});

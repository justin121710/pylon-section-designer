/* 回歸測試（第二輪嚴格審查 S1～S3）：
     S1 梁正、負彎矩兩檢核斷面皆判定：判定列、警示、計算書、構件清單、Excel 取兩者最不利
     S2 耐震梁 ρ ≤ min((f'c+100)/(4f_y), 0.025)，上下兩面皆檢核 ρ 與 A_s,min（土木401-112 §18.3.3.1）
     S3 瀏覽器列印（Ctrl+P）前依目前輸入重建計算書；未完成時改印「不得送審」；輸入一變即清空舊計算書
   Excel 部分需要 LibreOffice（soffice）重算；未安裝時略過該段。 */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs'), os = require('os'), path = require('path'), {execFileSync} = require('child_process');
const ExcelJS = require('exceljs');
const {openApp} = require('./helpers');

const hasSoffice = (() => { try{ execFileSync('soffice', ['--version'], {stdio:'ignore', timeout:60000}); return true; }catch(e){ return false; } })();
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'rcsd-rv2-'));
const val = v => (v && typeof v === 'object') ? ('result' in v ? v.result : v.richText ? v.richText.map(t => t.text).join('') : v) : v;
/* 「檢核表」A 欄名稱 → 該列（C 欄值、E 欄判定） */
function row(wb, label){
  const ws = wb.getWorksheet('檢核表'); let out;
  ws.eachRow(r => { if(out === undefined && String(val(r.getCell(1).value) || '').trim() === label) out = {v: val(r.getCell(3).value), judge: val(r.getCell(5).value)}; });
  assert.notEqual(out, undefined, `檢核表找不到「${label}」`);
  return out;
}

let app, page;
test.before(async () => { app = await openApp(); page = app.page; });
test.after(async () => { if(app){ assert.deepEqual(app.errors, [], '頁面不得有 JavaScript 錯誤'); await app.close(); } });

const exportAs = async name => {
  await page.evaluate(() => confAll(TAB==='beam' ? 'beam' : 'pylon'));   // H7：需確認未完成時匯出受阻
  const [dl] = await Promise.all([page.waitForEvent('download', {timeout:120000}), page.click('#btnXlsx')]);
  await dl.saveAs(path.join(TMP, name + '.xlsx'));
};
/* 建築 RC 大梁，頂筋改 2-#8：支承負彎矩 D/C ≈ 2.04，跨中正彎矩 D/C ≈ 0.60 */
const weakTop = sign => page.evaluate(sign => { document.getElementById('tabBeam').click(); applyBPreset('bldgBeam');
  document.getElementById('bTopSize').value = '#8'; document.getElementById('bNtop').value = '2'; BSIGN = sign; renderBeam(); }, sign);
const beamState = () => page.evaluate(() => {
  const vd = [...document.querySelectorAll('#verdict .vd')].map(b => ({st: b.className.replace('vd ', ''), t: b.textContent.replace(/\s+/g, ' ')}));
  buildBeamReport(BMODEL);
  return {cur: BMODEL.ctrl.dc, alt: BMODEL.alt && BMODEL.alt.B.ctrl.dc, gov: beamDcGov(BMODEL).dc, vd,
    bad: beamAlertItems(BMODEL).filter(a => a.lv === 'bad').map(a => optPlain(a)), rep: document.getElementById('report').textContent,
    mem: memSummary().dc};
});

test('S1：畫面停在跨中正彎矩時，支承負彎矩斷面不足仍判不合格（判定列、警示、計算書、構件清單）', async () => {
  for(const sign of ['pos', 'neg']){
    await weakTop(sign);
    const s = await beamState();
    assert.ok(s.gov > 2 && s.gov < 2.1, `兩斷面最不利 D/C ≈ 2.04（${s.gov}）`);
    const flex = s.vd.find(v => v.t.includes('撓曲 D/C'));
    assert.equal(flex.st, 'bad', `檢核 ${sign}：撓曲 D/C 膠囊應為不合格`);
    assert.ok(flex.t.includes('2.04'), flex.t);
    assert.equal(s.vd.find(v => v.t.includes('不合格')).st, 'bad');
    assert.ok(s.bad.some(t => /撓曲不足：D\/C = 2\.04/.test(t)), '警示列出負彎矩斷面撓曲不足');
    assert.ok(s.rep.includes('4.0 正、負彎矩兩檢核斷面之撓曲判定'), '計算書列出兩斷面判定');
    assert.ok(s.rep.includes('不合格（NG）') && !s.rep.includes('合格（OK）'), '計算書總判定為不合格');
    assert.ok(s.mem > 2, '構件清單之撓曲 D/C 取兩斷面最不利');
  }
  // 兩斷面皆撓曲不足：兩者皆列出（不因訊息相同而合併）
  const both = await page.evaluate(() => { applyBPreset('bldgBeam'); document.getElementById('bTopSize').value = '#8'; document.getElementById('bNtop').value = '2';
    document.getElementById('bNbot').value = '2'; BLOADS = [{name:'A', Mpos:60, Mneg:60, Vu:20, Tu:0}]; drawBLoads(); BSIGN = 'pos'; renderBeam();
    return beamAlertItems(BMODEL).filter(a => a.lv === 'bad').map(a => optPlain(a)).filter(t => t.includes('撓曲不足')); });
  assert.equal(both.length, 2, `兩斷面之撓曲不足各列一項：${both.join(' / ')}`);
  assert.ok(both.some(t => t.startsWith('【支承負彎矩斷面】')));
  const pos = await (async () => { await weakTop('pos'); return beamState(); })();
  assert.ok(pos.bad.some(t => t.startsWith('【支承負彎矩斷面】撓曲不足')), '另一斷面之項目標示斷面名稱');
  assert.ok(pos.cur < 0.7, '目前檢核斷面（跨中）本身 D/C ≈ 0.60');
});

test('S1：另一檢核斷面無彎矩需求（M_u = 0）時不判定，不因未配頂筋而判不合格', async () => {
  const s = await page.evaluate(() => { document.getElementById('tabBeam').click(); applyBPreset('bldgBeam');
    document.getElementById('bSeismic').value = '0'; document.getElementById('bNtop').value = '0';
    BLOADS.forEach(L => { L.Mneg = 0; }); drawBLoads(); BSIGN = 'pos'; renderBeam();
    return {has: BMODEL.alt.has, bad: beamAlertItems(BMODEL).filter(a => a.lv === 'bad').map(a => optPlain(a)),
      info: beamAlertItems(BMODEL).some(a => a.lv === 'info' && a.t.includes('【支承負彎矩斷面】各組合之 M_u 皆為 0'))}; });
  assert.equal(s.has, false);
  assert.deepEqual(s.bad.filter(t => t.startsWith('【')), [], '另一斷面不得產生不合格');
  assert.ok(s.info, '註明另一斷面 M_u = 0 不判定');
});

test("S2：耐震梁 ρ 上限 (f'c+100)/(4f_y)，上下兩面皆檢核 ρ 與 A_s,min", async () => {
  // f'c = 210：上限 310/16800 = 1.845% < 2.5%
  const s = await page.evaluate(() => { document.getElementById('tabBeam').click(); applyBPreset('bldgBeam');
    document.getElementById('bfc').value = 210; document.getElementById('bNbot').value = 10; BSIGN = 'neg'; renderBeam();
    const F = BMODEL.seisFace;
    return {rM: F.rhoMax, rhoB: F.bot.rho, bad: beamAlertItems(BMODEL).filter(a => a.lv === 'bad').map(a => optPlain(a)),
      rep: (buildBeamReport(BMODEL), document.getElementById('report').textContent)}; });
  assert.ok(Math.abs(s.rM - 310/16800) < 1e-9, `上限 ${s.rM}`);
  assert.ok(s.rhoB > s.rM && s.rhoB < 0.025, `底筋 ρ = ${s.rhoB} 介於上限與 0.025 之間`);
  assert.ok(s.bad.some(t => t.startsWith('耐震梁底筋鋼筋比')), '目前檢核負彎矩斷面時，底筋 ρ 超限仍判不合格');
  assert.ok(s.rep.includes('底筋鋼筋比') && s.rep.includes('NG'), '計算書第 7 節列出底筋 ρ 不合格');

  // 頂筋量不足：即使所有組合 M_u⁻ = 0 仍須 ≥ A_s,min（上下兩面）
  const t = await page.evaluate(() => { applyBPreset('bldgBeam'); document.getElementById('bTopSize').value = '#4'; document.getElementById('bNtop').value = '2';
    BLOADS.forEach(L => { L.Mneg = 0; }); drawBLoads(); BSIGN = 'pos'; renderBeam();
    return beamAlertItems(BMODEL).filter(a => a.lv === 'bad').map(a => optPlain(a)); });
  assert.ok(t.some(x => x.startsWith('耐震梁頂筋量不足')), '頂筋 A_s < A_s,min 判不合格');
});

test('S1、S2：Excel 檢核彙總含另一檢核斷面與上下兩面耐震檢核，重算後與網頁一致', {skip: !hasSoffice && '未安裝 LibreOffice（soffice）'}, async () => {
  await weakTop('pos');
  const w1 = await page.evaluate(() => ({alt: BMODEL.alt.B.ctrl.dc}));
  await exportAs('s1');
  const w2 = await page.evaluate(() => { applyBPreset('bldgBeam'); document.getElementById('bfc').value = 210; document.getElementById('bNbot').value = 10;
    BSIGN = 'pos'; renderBeam(); return {rhoB: BMODEL.seisFace.bot.rho, rM: BMODEL.seisFace.rhoMax}; });
  await exportAs('s2');
  const out = path.join(TMP, 'out');
  execFileSync('soffice', ['--headless', '--convert-to', 'xlsx', '--outdir', out, path.join(TMP, 's1.xlsx'), path.join(TMP, 's2.xlsx')], {stdio:'ignore', timeout:300000});
  const load = async n => { const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(path.join(out, n + '.xlsx')); return wb; };

  const b1 = await load('s1');
  const dco = row(b1, '另一檢核斷面撓曲 D/C').v;
  assert.ok(Math.abs(dco - w1.alt) <= 0.005*w1.alt, `另一斷面 D/C：Excel ${dco}，網頁 ${w1.alt}`);
  assert.equal(row(b1, '另一檢核斷面：撓曲 Mu／φMn（tf·m）').judge, 'FAIL');
  assert.equal(row(b1, '撓曲 Mu／φMn（tf·m）').judge, 'PASS', '目前檢核斷面本身通過');
  assert.equal(row(b1, '總判定').judge, 'NG', '總判定取兩斷面最不利');

  const b2 = await load('s2');
  const rb = row(b2, "耐震：底筋 ρ ≦ min((f'c+100)/(4fy), 0.025)");
  assert.equal(rb.judge, 'FAIL');
  assert.ok(Math.abs(row(b2, '耐震梁拉力鋼筋比上限').v - w2.rM) < 1e-9);
  assert.ok(Math.abs(row(b2, '底筋鋼筋比').v - w2.rhoB) <= 0.005*w2.rhoB, '底筋 ρ 與網頁一致');
  assert.equal(row(b2, '總判定').judge, 'NG');
});

test('S3：輸入改變即清空舊計算書；Ctrl+P 前依目前輸入重建，未完成時改印「不得送審」', async () => {
  const r = await page.evaluate(() => {
    document.getElementById('tabPylon').click(); applyPreset('bldg60'); confAll('pylon');
    setPreview(true); setPreview(false);
    document.getElementById('B').value = 90; render();
    const cleared = document.getElementById('report').innerHTML === '';
    // 需確認未完成 → 整頁「不得送審」，不含任何計算內容
    const pend = confPending('pylon').length;
    window.dispatchEvent(new Event('beforeprint'));
    const held = document.getElementById('report').textContent;
    window.dispatchEvent(new Event('afterprint'));
    const afterHeld = document.getElementById('report').innerHTML;
    // 全部確認後 → 依目前輸入（B = 90）重建
    confAll('pylon');
    window.dispatchEvent(new Event('beforeprint'));
    const ok = document.getElementById('report').textContent;
    window.dispatchEvent(new Event('afterprint'));
    // 輸入不完整 → 不得送審
    document.getElementById('B').value = ''; render();
    window.dispatchEvent(new Event('beforeprint'));
    const blocked = document.getElementById('report').textContent;
    window.dispatchEvent(new Event('afterprint'));
    document.getElementById('B').value = 60; render();
    return {cleared, pend, held, afterHeld, ok, blocked, hash: inputHash()};
  });
  assert.ok(r.cleared, '關閉預覽後修改輸入，#report 應清空');
  assert.ok(r.pend > 0);
  assert.ok(r.held.includes('不得送審') && !r.held.includes('斷面與材料'), '需確認未完成：只印「不得送審」');
  assert.equal(r.afterHeld, '', '列印結束後移除「不得送審」頁');
  assert.ok(r.ok.includes('90') && r.ok.includes('斷面與材料') && !r.ok.includes('不得送審'), '依目前輸入重建計算書');
  assert.ok(r.blocked.includes('不得送審') && r.blocked.includes('輸入不完整'), '輸入不完整時不得印出');
});

/* ---------------- 第二輪審查 H1～H9 ---------------- */
const pick = re => page.evaluate(re => { const b = [...document.querySelectorAll('#askAct button')].find(x => new RegExp(re).test(x.textContent)); if(!b) return false; b.click(); return true; }, re.source);
const askOpen = () => page.waitForFunction(() => document.getElementById('askOv').classList.contains('open'));

test('H1：柱 M_pr 取軸力範圍內最大值（組合軸力跨越平衡點）', async () => {
  const r = await page.evaluate(() => { document.getElementById('tabPylon').click(); applyPreset('bldg60');
    const ps = [620, 560, 150, 60]; LOADS.forEach((L,i) => { L.Pu = ps[i]; }); drawLoads(); render();
    return {Ve: MODEL.shX.Ve/1000, PMn: MODEL.shX.PMn}; });
  assert.ok(r.Ve > 68.5 && r.Ve < 69.7, `V_e ≈ 69.1 tf（${r.Ve}）`);
  assert.ok(r.PMn > 60e3 && r.PMn < 620e3 && ![60e3,150e3,560e3,620e3].includes(r.PMn), '峰值軸力落在組合之間');
});

test('H2：特殊抗彎矩構架柱尺寸限制 §18.4.2.1', async () => {
  const r = await page.evaluate(() => { applyPreset('bldg60'); document.getElementById('B').value = 25; document.getElementById('H').value = 70; render();
    const bad = colAlertItems(MODEL).filter(a => a.lv==='bad').map(a => optPlain(a));
    buildReport(MODEL); return {bad, rep: document.getElementById('report').textContent}; });
  assert.ok(r.bad.some(t => t.startsWith('柱斷面最小尺度不足')));
  assert.ok(r.bad.some(t => t.startsWith('柱斷面尺度比不足')));
  assert.ok(r.rep.includes('§18.4.2.1(a)') && r.rep.includes('不合格（NG）'));
  await page.evaluate(() => applyPreset('bldg60'));
});

test('H3：梁即時活載撓度 Δ_L = Δ(D+L) − Δ(D)，Δ(D) 以 I_e(M_D)', async () => {
  const r = await page.evaluate(() => { document.getElementById('tabBeam').click(); applyBPreset('bldgJoist'); BSIGN = 'pos'; renderBeam();
    const D = BMODEL.defl, dl = (M, I) => D.K*M*BMODEL.S.L**2/(BMODEL.m.Ec*I);
    return {dL: D.dL, exp: dl(D.Mtot, D.Ie) - dl(D.MD, D.IeD), old: dl(D.Mtot, D.Ie) - dl(D.MD, D.Ie), IeD: D.IeD, Ie: D.Ie}; });
  assert.ok(Math.abs(r.dL - r.exp) < 1e-9);
  assert.ok(r.IeD >= r.Ie && r.dL >= r.old, '以 I_e(M_D) 計 Δ(D) 不低於舊法');
});

test('H5：柱計算書列出檢核彙總、總判定、強柱弱梁與主筋比', async () => {
  const t = await page.evaluate(() => { document.getElementById('tabPylon').click(); applyPreset('bldg60'); buildReport(MODEL); return document.getElementById('report').textContent; });
  for(const k of ['檢核彙總與總判定', '總判定：', '強柱弱梁', '主筋比', '撓曲＋軸力 D/C']) assert.ok(t.includes(k), k);
});

test('H6：需確認狀態綁定輸入、存入案件與構件、批次計算後保留、印於計算書', async () => {
  const r = await page.evaluate(async () => { applyPreset('bldg60'); render(); confAll('pylon');
    MEMBERS = []; document.getElementById('memName').value = 'C1'; memAdd(); await memBatch();
    const afterBatch = confPending('pylon').length, memPend = MEM_RES[MEMBERS[0].id].pend;
    document.getElementById('B').value = 90; render(); const changed = confPending('pylon').length;
    document.getElementById('B').value = 60; render(); const back = confPending('pylon').length;
    const snap = JSON.parse(JSON.stringify(caseSnapshot())); applyPreset('bldg60'); render(); const cleared = confPending('pylon').length;
    caseApply(snap); const reopened = confPending('pylon').length;
    buildReport(MODEL); const rep = document.getElementById('report').textContent;
    MEMBERS = []; memSave(); memRender();
    return {afterBatch, memPend, changed, back, cleared, reopened, rep: rep.includes('設計者確認事項') && rep.includes('已確認')}; });
  assert.equal(r.afterBatch, 0, '批次計算後不清除'); assert.equal(r.memPend, 0, '構件快照帶確認狀態');
  assert.ok(r.changed > 0, '輸入改變（含無數字之未涵蓋項目）須重新確認'); assert.equal(r.back, 0, '改回原值確認恢復');
  assert.ok(r.cleared > 0); assert.equal(r.reopened, 0, '案件檔保存確認狀態');
  assert.ok(r.rep, '計算書列出設計者確認事項');
});

test('H7：需確認未完成時 Excel 匯出受阻，可改匯出標示「不得送審」之草稿', async () => {
  await page.evaluate(() => { applyPreset('bldg60'); render(); });
  assert.ok(await page.evaluate(() => confPending('pylon').length > 0));
  let dl = null; const onDl = d => { dl = d; }; page.on('download', onDl);
  await page.click('#btnXlsx'); await askOpen(); await pick(/^取消$/); await page.waitForTimeout(300);
  assert.equal(dl, null, '取消時不產生檔案');
  const p = page.waitForEvent('download', {timeout:120000});
  await page.click('#btnXlsx'); await askOpen(); await pick(/匯出草稿/);
  const d = await p; page.off('download', onDl);
  const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(await d.path());
  assert.match(wb.getWorksheet('檢核表').headerFooter.oddHeader || '', /不得送審/);
  assert.match(String(wb.getWorksheet('檢核表').getCell('A1').value), /草稿.*不得送審/);
});

test('H8：切換為 kN 時可選「重新解讀」，表頭與說明列只標一種單位', async () => {
  const r = await page.evaluate(async () => { applyPreset('bldg60'); const pu0 = LOADS[0].Pu;
    const u = document.getElementById('loadUnit'); u.value = 'kN'; u.dispatchEvent(new Event('change'));
    [...document.querySelectorAll('#askAct button')].find(x => /重新解讀/.test(x.textContent)).click(); await new Promise(ok => setTimeout(ok, 0));
    const out = {pu0, pu: LOADS[0].Pu, shown: +document.querySelector('#loadTbl tbody input[data-k="Pu"]').value,
      head: document.querySelector('#loadTbl thead th:nth-child(3)').textContent, note: document.querySelector('.lu-f').textContent};
    u.value = 'tf'; u.dispatchEvent(new Event('change'));
    [...document.querySelectorAll('#askAct button')].find(x => /僅換算顯示/.test(x.textContent)).click(); await new Promise(ok => setTimeout(ok, 0));
    applyPreset('bldg60'); return out; });
  assert.ok(Math.abs(r.pu - r.pu0/9.80665) < 1e-9, '內部載重改以 kN 解讀');
  assert.ok(Math.abs(r.shown - r.pu0) < 1e-6, '表內數字不變');
  assert.ok(r.head.includes('kN') && !r.head.includes('tf'), `表頭：${r.head}`);
  assert.equal(r.note, 'kN');
});

test('H9：開啟案件前確認未存檔變更；構件清單取代後與儲存同步；刪除構件須確認', async () => {
  await page.evaluate(() => { applyPreset('bldg60'); render(); MEMBERS = [];
    for(const n of ['A','B','C','D']){ document.getElementById('memName').value = n; memAdd(); }
    caseMark(); document.getElementById('H').value = 123; render(); });
  const c = await page.evaluate(() => { const x = caseSnapshot(); x.members = []; x.values.H = '60'; return x; });
  const f = path.join(os.tmpdir(), `rcsd-h9-${process.pid}.json`); fs.writeFileSync(f, JSON.stringify(c));
  await page.setInputFiles('#caseFile', f);
  await askOpen();
  assert.match(await page.evaluate(() => document.getElementById('askTtl').textContent), /覆蓋目前的輸入/);
  await pick(/^取消$/); await page.waitForTimeout(200);
  assert.equal(await page.evaluate(() => document.getElementById('H').value), '123', '取消時不覆蓋');
  await page.setInputFiles('#caseFile', f); await askOpen(); await pick(/直接開啟/);
  await page.waitForFunction(() => /構件清單/.test(document.getElementById('askTtl').textContent) && document.getElementById('askOv').classList.contains('open'));
  await pick(/以案件檔取代/);
  await page.waitForFunction(() => document.getElementById('H').value === '60');
  const s = await page.evaluate(() => ({n: MEMBERS.length, stored: JSON.parse(localStorage.getItem(MEM_KEY)).length}));
  assert.deepEqual(s, {n: 0, stored: 0}, '構件清單與儲存同步（重新整理不會回來）');
  // 刪除構件須確認
  await page.evaluate(() => { document.getElementById('memName').value = 'X'; memAdd(); memRender(); document.querySelector('.mem-del').click(); });
  await askOpen(); await pick(/^取消$/); await page.waitForTimeout(100);
  assert.equal(await page.evaluate(() => MEMBERS.length), 1);
  await page.evaluate(() => document.querySelector('.mem-del').click()); await askOpen(); await pick(/^刪除$/); await page.waitForTimeout(100);
  assert.equal(await page.evaluate(() => MEMBERS.length), 0);
  fs.unlinkSync(f);
});

test('H1、H2、H4：Excel 之 V_e 取軸力範圍最大值、柱尺寸檢核；修改輸入後雙軸 D/C 與總判定改為須回網頁重算', {skip: !hasSoffice && '未安裝 LibreOffice（soffice）'}, async () => {
  const w = await page.evaluate(() => { document.getElementById('tabPylon').click(); applyPreset('bldg60');
    const ps = [620, 560, 150, 60]; LOADS.forEach((L,i) => { L.Pu = ps[i]; }); drawLoads(); render(); return MODEL.shX.Ve/1000; });
  await exportAs('h4');
  const src = new ExcelJS.Workbook(); await src.xlsx.readFile(path.join(TMP, 'h4.xlsx'));
  const ws = src.getWorksheet('檢核表'); let hr = 0;
  ws.eachRow(r => { if(!hr && String(val(r.getCell(1).value)).startsWith('外高 H')) hr = r.number; });
  ws.getCell('C' + hr).value = 25;                  // 60×25：b/h = 0.417，最小尺度 25 < 30
  await src.xlsx.writeFile(path.join(TMP, 'h4b.xlsx'));
  const out = path.join(TMP, 'out4');
  execFileSync('soffice', ['--headless', '--convert-to', 'xlsx', '--outdir', out, path.join(TMP, 'h4.xlsx'), path.join(TMP, 'h4b.xlsx')], {stdio:'ignore', timeout:300000});
  const load = async n => { const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(path.join(out, n + '.xlsx')); return wb; };
  const a = await load('h4'), b = await load('h4b');
  const ve = row(a, 'X 向容量設計剪力 Ve').v;
  assert.ok(Math.abs(ve - w) <= 0.01*w, `Excel V_e ${ve}，網頁 ${w}`);
  assert.equal(row(a, '撓曲 D/C 狀態').v, '有效');
  assert.match(String(row(b, '撓曲 D/C 狀態').v), /須回網頁重算/);
  assert.equal(row(b, '撓曲＋軸力 D/C').judge, 'N/A');
  assert.match(String(row(b, '總判定').judge), /須回網頁重算/);
  assert.equal(row(b, '柱斷面最小尺度（cm）').judge, 'FAIL');
  await page.evaluate(() => applyPreset('bldg60'));
});

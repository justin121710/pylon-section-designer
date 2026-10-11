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

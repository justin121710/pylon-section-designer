/* 回歸測試（構件細部：多段續接與定尺長度、箍筋延伸入接頭／基礎、梁剪力分區、梁表皮鋼筋）
   Excel 部分需要 LibreOffice（soffice）重算；未安裝時略過該段。 */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs'), os = require('os'), path = require('path'), {execFileSync} = require('child_process');
const ExcelJS = require('exceljs');
const {openApp} = require('./helpers');

const hasSoffice = (() => { try{ execFileSync('soffice', ['--version'], {stdio:'ignore', timeout:60000}); return true; }catch(e){ return false; } })();
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'rcsd-det-'));
const val = v => (v && typeof v === 'object') ? ('result' in v ? v.result : v.richText ? v.richText.map(t => t.text).join('') : v) : v;
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
  await page.evaluate(() => confAll(TAB==='beam' ? 'beam' : 'pylon'));
  const [dl] = await Promise.all([page.waitForEvent('download', {timeout:120000}), page.click('#btnXlsx')]);
  await dl.saveAs(path.join(TMP, name + '.xlsx'));
};
const recalc = (names, out) => {
  execFileSync('soffice', ['--headless', '--convert-to', 'xlsx', '--outdir', path.join(TMP, out), ...names.map(n => path.join(TMP, n + '.xlsx'))], {stdio:'ignore', timeout:300000});
  return async n => { const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(path.join(TMP, out, n + '.xlsx')); return wb; };
};

/* ---------------- 多段續接與定尺長度 ---------------- */
const colStock = (key, set) => page.evaluate(([key, set]) => { document.getElementById('tabPylon').click(); applyPreset(key);
  for(const [id, v] of Object.entries(set || {})) document.getElementById(id).value = v;
  const M = render(), P = M.devPlan, rows = devSchedule(P, M);
  return {spl: P.splices.map(q => q.a), stag: P.stag, lap: P.lap, zone: P.zone, A0: P.A0, A1: P.A1, ok: P.stockOK, max: P.maxPiece,
    lens: rows.map(r => r.len), ids: rows.map(r => r.id), chk: P.checks.find(c => /定尺/.test(c.item)) || null,
    lapZ: devLapRegions(P).length, lapDense: (P.lay.lapZones || []).length}; }, [key, set]);

test('定尺：長構材依定尺增設續接，各段下料長 ≤ 定尺，續接處在可續接範圍內', async () => {
  for(const k of ['pier200', 'box300', 'box600', 'pileC120']){
    const r = await colStock(k);
    assert.ok(r.ok, `${k}：定尺可滿足`);
    assert.ok(r.spl.length >= 2, `${k}：須增設續接（${r.spl.length} 處）`);
    for(const L of r.lens) assert.ok(L <= 1200, `${k}：下料長 ${L} ≤ 1200`);
    for(let i=1;i<r.spl.length;i++) assert.ok(r.spl[i] - r.spl[i-1] >= r.zone - 1e-9, `${k}：相鄰續接區段不重疊`);
    for(const a of r.spl) assert.ok(a >= r.A0 - 1e-9 && a + r.zone <= r.A1 + 1e-9, `${k}：續接處 ${a} 在 [${r.A0}, ${r.A1}] 內`);
    assert.equal(r.lapZ, r.spl.length, `${k}：軸向配置標示每處搭接區`);
    assert.equal(r.chk.lv, 'ok');
  }
  // 短柱：單一續接，下料表與原作法相同（底部插筋＋柱主筋兩段）
  const b = await colStock('pier100');
  assert.equal(b.spl.length, 1); assert.deepEqual(b.ids, ['C1', 'C2']);
  // 0＝不限：不增設、不列檢核
  const u = await colStock('box600', {dStock: 0});
  assert.equal(u.spl.length, 1); assert.equal(u.chk, null);
});

test('定尺：錯開搭接兩組各段皆 ≤ 定尺；機械式續接無重疊；定尺過短判不合格', async () => {
  const s = await colStock('box600', {dStagger: '1'});
  assert.ok(s.stag && s.ok); for(const L of s.lens) assert.ok(L <= 1200, `錯開：${L} ≤ 1200`);
  assert.ok(s.ids.some(i => /-2$/.test(i)), '第二組另列');
  const m = await colStock('box600', {dSplice: 'm3'});
  assert.ok(m.ok && m.spl.length >= 3); for(const L of m.lens) assert.ok(L <= 1200);
  const t = await colStock('box600', {dStock: 200});
  assert.equal(t.ok, false); assert.equal(t.chk.lv, 'bad');
  // 建築耐震柱：續接僅能在中央半高；淨高 14 m 之柱以 6 m 定尺無法於範圍內完成 → 不合格
  const bl = await colStock('bldg60', {lu: 1400, dStock: 600});
  assert.equal(bl.ok, false, '建築柱中央半高內無法依定尺續接');
  await page.evaluate(() => applyPreset('bldg60'));
});

test('定尺：Excel 續接處數與最長下料長同網頁（LibreOffice 重算）', {skip: !hasSoffice && '未安裝 LibreOffice（soffice）'}, async () => {
  const cases = [['pier200', {}], ['box600', {}], ['box600', {dStagger: '1'}], ['pileC120', {}], ['box600', {dStock: 200}], ['pier100', {}]];
  const web = [];
  for(let i=0;i<cases.length;i++){
    const r = await colStock(...cases[i]);
    web.push(r); await exportAs('st' + i);
  }
  const load = recalc(cases.map((_, i) => 'st' + i), 'outst');
  for(let i=0;i<cases.length;i++){
    const wb = await load('st' + i), r = web[i], tag = cases[i][0] + JSON.stringify(cases[i][1]);
    // LibreOffice 以布林值存 --AND() 之 0／1，ExcelJS 讀 FALSE 時無 result → 視為 0
    const okV = row(wb, '定尺長度符合').v;
    assert.equal(okV && typeof okV === 'object' ? 0 : Number(okV), r.ok ? 1 : 0, `${tag}：判定`);
    if(r.ok){
      assert.equal(row(wb, '續接處數（含增設）').v, r.spl.length, `${tag}：續接處數`);
      const mx = row(wb, '主筋最長下料長').v;
      assert.ok(Math.abs(Math.ceil(mx - 1e-9) - r.max) <= 1, `${tag}：最長下料 Excel ${mx}、網頁 ${r.max}`);
    }
  }
  await page.evaluate(() => applyPreset('bldg60'));
});

/* ---------------- 箍筋延伸入接頭／基礎 ---------------- */
const colJ = (key, set) => page.evaluate(([key, set]) => { document.getElementById('tabPylon').click(); applyPreset(key);
  for(const [id, v] of Object.entries(set || {})) document.getElementById(id).value = v;
  const M = render(), P = M.devPlan;
  return {j: P.jext, ld: Math.ceil(P.D.ld - 1e-9), s1: M.fin.sUse, hTop: P.X.hTop, hBot: P.X.hBot, cov: P.X.cov, nEnds: M.lay.nEnds,
    chk: P.checks.filter(c => /箍筋延伸/.test(c.item)).map(c => c.lv), svg: svgColElev(M, true), info: elevInfoCol(M)}; }, [key, set]);

test('箍筋延伸：橋梁塑鉸區延伸入基礎／帽梁 ≥ l_d（§5.3.5）、建築柱頂接頭 h（§18.5.3.1）、基腳邊緣條件（§18.10.2.3）', async () => {
  const b = await colJ('pier100');
  const bot = b.j.find(q => q.end==='bot');
  assert.ok(bot && bot.need === b.ld, `橋梁柱底延伸 = l_d（${bot && bot.need}／${b.ld}）`);
  assert.equal(bot.s, b.s1); assert.equal(bot.n, Math.ceil(b.ld/b.s1 - 1e-9));
  assert.equal(!!b.j.find(q => q.end==='top'), b.nEnds === 2, '柱頂僅兩端塑鉸時延伸');
  assert.match(b.svg, /延伸入基礎/); assert.match(b.info, /延伸入基礎/);
  const sh = await colJ('pier100', {dHBot: 40});
  assert.equal(sh.j.find(q => q.end==='bot').ok, false, '基礎過淺判不合格'); assert.ok(sh.chk.includes('bad'));
  const n = await colJ('pier100', {ph: '0'});
  assert.equal(n.j.length, 0, '非塑鉸區不延伸');
  const c = await colJ('bldg60');
  assert.deepEqual(c.j.map(q => [q.end, q.need]), [['top', c.hTop]], '建築：柱頂接頭深 h；柱底貫穿續接屬下層柱');
  const f = await colJ('bldg60', {dAncBot: 'hook90', dFtgEdge: '1'});
  assert.equal(f.j.find(q => q.end==='bot').need, f.ld, '建築柱錨入基腳且近邊緣：延伸 ≥ l_d');
  const p = await colJ('pileC120');
  assert.equal(p.j.length, 0, '基樁不適用');
  await page.evaluate(() => applyPreset('bldg60'));
});

test('箍筋延伸：Excel 需求長度與數量同網頁（LibreOffice 重算）', {skip: !hasSoffice && '未安裝 LibreOffice（soffice）'}, async () => {
  const cases = [['pier100', {}], ['pier100', {phEnds: '2'}], ['bldg60', {}], ['bldg60', {dAncBot: 'hook90', dFtgEdge: '1'}]];
  const web = [];
  for(let i=0;i<cases.length;i++){ web.push(await colJ(...cases[i])); await exportAs('jx' + i); }
  const load = recalc(cases.map((_, i) => 'jx' + i), 'outjx');
  for(let i=0;i<cases.length;i++){
    const wb = await load('jx' + i), r = web[i], tag = cases[i][0] + JSON.stringify(cases[i][1]);
    const num = v => v && typeof v === 'object' ? 0 : Number(v);
    for(const [end, kN, kNum] of [['bot', '柱底箍筋延伸入基礎長度', '柱底延伸段箍筋數量（另計）'], ['top', '柱頂箍筋延伸入接頭長度', '柱頂延伸段箍筋數量（另計）']]){
      const q = r.j.find(x => x.end === end);
      assert.equal(num(row(wb, kN).v), q ? q.need : 0, `${tag} ${end}：長度`);
      assert.equal(num(row(wb, kNum).v), q ? q.n : 0, `${tag} ${end}：數量`);
    }
  }
  await page.evaluate(() => applyPreset('bldg60'));
});

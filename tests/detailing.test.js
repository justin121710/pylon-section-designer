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

/* ---------------- 梁剪力筋依剪力圖分區 ---------------- */
const beamZ = (key, set, loads) => page.evaluate(([key, set, loads]) => { document.getElementById('tabBeam').click(); applyBPreset(key);
  if(loads) BLOADS.forEach(l => Object.assign(l, loads));
  drawBLoads();
  for(const [id, v] of Object.entries(set || {})) document.getElementById(id).value = v;
  const B = renderBeam();
  return {s1: B.fin.sUse, s2: B.fin2 ? B.fin2.sUse : null, x: B.fin2 ? B.fin2.x : null, Vx: B.fin2 ? B.fin2.Vx/1000 : null, zMid: !!B.zMid,
    Ve: B.sh.Ve/1000, total: B.lay ? B.lay.total : null, zones: B.lay ? B.lay.zones.map(z => ({tag:z.tag, a:z.a, b:z.b, s:z.s})) : [], ln: B.S.ln, h: B.S.h}; }, [key, set, loads]);

test('剪力分區：一般梁中央 l_n/2 以 l_n/4 處剪力、耐震梁 2h 外以 2h 處剪力設計；懸臂與無效益時不分區', async () => {
  const hiV = {Vu: 150};
  const a = await beamZ('gradeBeam', {}, hiV);
  assert.equal(a.zMid, false); assert.equal(a.zones.length, 1, '預設不分區');
  const z = await beamZ('gradeBeam', {bShearZ: '1', bVmid: 20}, hiV);
  assert.ok(z.zMid && z.s2 > z.s1, `中央區間距放大（${z.s1} → ${z.s2}）`);
  assert.ok(Math.abs(z.Vx - (150 - (150 - 20)*0.5)) < 1e-6, `l_n/4 處 V = 85（${z.Vx}）`);
  assert.deepEqual(z.zones.map(q => q.tag), ['端部區', '中央區', '端部區']);
  assert.ok(Math.abs(z.zones[0].b - z.ln/4) < 1e-6);
  assert.ok(z.total < a.total, `支數減少（${a.total} → ${z.total}）`);
  const sz = await beamZ('gradeBeam', {bShearZ: '1', bVmid: 20, bSeismic: '1', bVg: 40}, hiV);
  const Vm = Math.max(sz.Ve - 40, 20);
  assert.equal(sz.x, 2*sz.h);
  assert.ok(Math.abs(sz.Vx - (150 - (150 - Vm)*Math.min(1, 2*sz.h/(sz.ln/2)))) < 1e-6, `耐震梁 2h 處 V（${sz.Vx}）`);
  const s0 = await beamZ('gradeBeam', {bSeismic: '1', bVg: 40}, hiV);
  assert.ok(sz.s2 >= s0.s2, '耐震梁分區後加密區外間距不小於不分區');
  const c = await beamZ('gradeBeam', {bShearZ: '1', bVmid: 20, bDeflCase: 'cantUDL'}, hiV);
  assert.equal(c.zMid, false, '懸臂不分區');
  const p = await beamZ('gradeBeam', {bShearZ: '1'});
  assert.equal(p.zMid, false, '間距由 d/2 控制時無效益，維持全長同一間距');
  await page.evaluate(() => { applyBPreset('bldgBeam'); document.getElementById('tabPylon').click(); });
});

test('剪力分區：Excel 間距與支數同網頁（LibreOffice 重算）', {skip: !hasSoffice && '未安裝 LibreOffice（soffice）'}, async () => {
  const hiV = {Vu: 150};
  const cases = [[{}, hiV], [{bShearZ: '1', bVmid: 20}, hiV], [{bShearZ: '1', bVmid: 20, bSeismic: '1', bVg: 40}, hiV], [{bShearZ: '1'}, null]];
  const web = [];
  for(let i=0;i<cases.length;i++){ web.push(await beamZ('gradeBeam', ...cases[i])); await exportAs('bz' + i); }
  const load = recalc(cases.map((_, i) => 'bz' + i), 'outbz');
  for(let i=0;i<cases.length;i++){
    const wb = await load('bz' + i), r = web[i], tag = JSON.stringify(cases[i][0]);
    assert.equal(Number(row(wb, '全長合計數量').v), r.total, `${tag}：支數`);
    if(r.s2 != null) assert.equal(Number(row(wb, '加密區外採用間距').v), r.s2, `${tag}：s₂`);
  }
  await page.evaluate(() => { applyBPreset('bldgBeam'); document.getElementById('tabPylon').click(); });
});

/* ---------------- 梁表層鋼筋（§9.7.2.3） ---------------- */
const beamSk = (key, set, loads) => page.evaluate(([key, set, loads]) => { document.getElementById('tabBeam').click(); applyBPreset(key);
  if(loads) BLOADS = loads.map(l => ({...l}));
  drawBLoads();
  for(const [id, v] of Object.entries(set || {})) document.getElementById(id).value = v;
  const B = renderBeam(), K = B.skin;
  return {h: B.S.h, skin: K ? {nSide: K.nSide, sLim: K.sLim, sAct: K.sAct, pos: K.pos, neg: K.neg, ys: K.ys, total: K.total, size: K.size} : null,
    sched: devSchedule(B.devPlan, B).filter(r => r.id === 'S1').map(r => r.n), svg: svgBeam(B, true), y1: B.R.layers[0].y, yT: B.R.layers[B.R.layers.length-1].y}; }, [key, set, loads]);

test('表層鋼筋：h > 90 cm 兩側面配置、間距 ≤ 表 24.3.2 之 s、受拉面依彎矩方向；h ≤ 90 與版不適用', async () => {
  const both = [{name:'U', Mpos:60, Mneg:120, Vu:40, Tu:0}], posOnly = [{name:'U', Mpos:60, Mneg:0, Vu:40, Tu:0}];
  const a = await beamSk('gradeBeam', {bh: 150}, both);
  assert.ok(a.skin, 'h = 150 須配置'); assert.ok(a.skin.pos && a.skin.neg);
  assert.ok(a.skin.sAct <= a.skin.sLim + 1e-9, `間距 ${a.skin.sAct} ≤ ${a.skin.sLim}`);
  assert.equal(a.skin.nSide, Math.ceil((a.yT - a.y1)/a.skin.sLim - 1e-9) - 1, '全深：⌈範圍/s⌉ − 1');
  assert.deepEqual(a.sched, [a.skin.total], '下料表列入');
  const p = await beamSk('gradeBeam', {bh: 150}, posOnly);
  assert.ok(p.skin.pos && !p.skin.neg);
  assert.ok(Math.max(...p.skin.ys) <= p.h/2 + 1e-9, '正彎矩：僅距底面 h/2 內');
  assert.equal(p.skin.nSide, Math.ceil((p.h/2 - p.y1)/p.skin.sLim - 1e-9));
  const n = await beamSk('gradeBeam', {bh: 90}, both);
  assert.equal(n.skin, null, 'h = 90 不需配置');
  const sl = await beamSk('pondTop');
  assert.equal(sl.skin, null, '版不適用');
  await page.evaluate(() => { applyBPreset('bldgBeam'); document.getElementById('tabPylon').click(); });
});

test('表層鋼筋：Excel 每側支數與間距同網頁（LibreOffice 重算）', {skip: !hasSoffice && '未安裝 LibreOffice（soffice）'}, async () => {
  const both = [{name:'U', Mpos:60, Mneg:120, Vu:40, Tu:0}], posOnly = [{name:'U', Mpos:60, Mneg:0, Vu:40, Tu:0}], negOnly = [{name:'U', Mpos:0, Mneg:120, Vu:40, Tu:0}];
  const cases = [[{bh: 150}, both], [{bh: 150}, posOnly], [{bh: 180, bSkinSize: '#5'}, negOnly], [{bh: 90}, both]];
  const web = [];
  for(let i=0;i<cases.length;i++){ web.push(await beamSk('gradeBeam', ...cases[i])); await exportAs('sk' + i); }
  const load = recalc(cases.map((_, i) => 'sk' + i), 'outsk');
  for(let i=0;i<cases.length;i++){
    const wb = await load('sk' + i), r = web[i], tag = JSON.stringify(cases[i][0]) + JSON.stringify(cases[i][1][0]);
    const num = v => v && typeof v === 'object' ? 0 : Number(v);   // 公式結果為 0 時 ExcelJS 無 result
    assert.equal(num(row(wb, '表層鋼筋每側支數').v), r.skin ? r.skin.nSide : 0, `${tag}：支數`);
    if(r.skin) assert.ok(Math.abs(Number(row(wb, '表層鋼筋實際間距').v) - r.skin.sAct) < 0.01, `${tag}：間距`);
  }
  await page.evaluate(() => { applyBPreset('bldgBeam'); document.getElementById('tabPylon').click(); });
});

/* ---------------- 橋梁選項：續接錯開（§5.3.5）、錨定依解說（C5.5） ---------------- */
const brOpt = (key, set) => page.evaluate(([key, set]) => { document.getElementById('tabPylon').click(); applyPreset(key);
  for(const [id, v] of Object.entries(set || {})) document.getElementById(id).value = v;
  const M = render(), P = M.devPlan;
  return {stg: (P.checks.find(c => /續接錯開/.test(c.item)) || {}).lv || null, bot: P.bot.q && P.bot.q.len, top: P.top.q && P.top.q.len,
    ld: P.D.ld, db: P.D.db, bldg: P.bldg}; }, [key, set]);

test('橋梁選項：續接錯開預設不強制；強制時未錯開判不合格、機械式為注意；錨定依解說取 max(1.25l_d, 24d_b)', async () => {
  assert.equal((await brOpt('pier100')).stg, null, '預設不檢核');
  assert.equal((await brOpt('pier100', {dBrStag: '1'})).stg, 'bad');
  assert.equal((await brOpt('pier100', {dBrStag: '1', dStagger: '1'})).stg, 'ok');
  assert.equal((await brOpt('pier100', {dBrStag: '1', dSplice: 'm3'})).stg, 'warn');
  assert.equal((await brOpt('bldg60', {dBrStag: '1'})).stg, null, '建築不適用');
  const a = await brOpt('pier100'), c = await brOpt('pier100', {dBrAnc: '1'});
  const need = Math.ceil(Math.max(1.25*c.ld, 24*c.db) - 1e-9);
  assert.equal(c.bot, need); assert.equal(c.top, need);
  assert.ok(c.bot > a.bot, `依解說較長（${a.bot} → ${c.bot}）`);
  await page.evaluate(() => applyPreset('bldg60'));
});

test('橋梁選項：Excel 錨定長度與錯開判定同網頁（LibreOffice 重算）', {skip: !hasSoffice && '未安裝 LibreOffice（soffice）'}, async () => {
  const cases = [{}, {dBrAnc: '1'}, {dBrStag: '1'}, {dBrStag: '1', dStagger: '1'}];
  const web = [];
  for(let i=0;i<cases.length;i++){ web.push(await brOpt('pier100', cases[i])); await exportAs('br' + i); }
  const load = recalc(cases.map((_, i) => 'br' + i), 'outbr');
  const J = {bad:'FAIL', ok:'PASS'};
  for(let i=0;i<cases.length;i++){
    const wb = await load('br' + i), r = web[i], tag = JSON.stringify(cases[i]);
    assert.equal(Number(row(wb, '底端錨定長度').v), r.bot, `${tag}：底端錨定`);
    assert.equal(row(wb, '橋梁主筋續接錯開').judge, r.stg ? J[r.stg] : 'N/A', `${tag}：錯開判定`);
  }
  await page.evaluate(() => applyPreset('bldg60'));
});

/* ---------------- 橋墩施工縫（§5.6） ---------------- */
const cjRun = (key, set) => page.evaluate(([key, set]) => { document.getElementById('tabPylon').click(); applyPreset(key);
  for(const [id, v] of Object.entries(set || {})) document.getElementById(id).value = v;
  const M = render(), P = M.devPlan, ck = P.checks.filter(c => /施工縫/.test(c.item)).map(c => c.lv);
  return {cj: P.cj, ck, Ast: M.Ast, fy: M.m.fy, fc: M.m.fc, Ag: M.Ag, phi: M.m.phivB, la: P.la}; }, [key, set]);

test('施工縫：V_j = φμ(A_vf f_y + 0.75P_n) 對設計剪力；搭接處判不合格；建築與預設不檢核', async () => {
  assert.equal((await cjRun('pier100')).cj, null, '預設無施工縫');
  assert.equal((await cjRun('bldg60', {dCj: 'base'})).cj, null, '建築不適用');
  const r = await cjRun('pier100', {dCj: 'base'});
  const Vn = Math.min(1.0*(r.Ast*r.fy + 0.75*r.cj.Pmin), Math.min(0.2*r.fc, 56.2)*r.Ag);
  assert.ok(Math.abs(r.cj.Vj - r.phi*Vn) < 1e-6, 'V_j 式');
  assert.equal(r.ck[0], r.cj.Vj >= r.cj.Vreq - 1e-6 ? 'ok' : 'bad'); assert.equal(r.ck[1], 'ok');
  const m6 = await cjRun('pier100', {dCj: 'base', dCjMu: '0.6'});
  assert.ok(m6.cj.Vj < r.cj.Vj, 'μ = 0.6λ 較小');
  const lap = await cjRun('pier100', {dCj: 'h', dCjH: Math.round(r.la + 20)});
  assert.equal(lap.cj.inLap, true); assert.equal(lap.ck[1], 'bad');
  await page.evaluate(() => applyPreset('bldg60'));
});

test('施工縫：Excel 同網頁（LibreOffice 重算）', {skip: !hasSoffice && '未安裝 LibreOffice（soffice）'}, async () => {
  const r0 = await cjRun('pier100', {dCj: 'base'});
  const cases = [{dCj: 'base'}, {dCj: 'base', dCjMu: '0.6'}, {dCj: 'h', dCjH: Math.round(r0.la + 20)}];
  const web = [];
  for(let i=0;i<cases.length;i++){ web.push(await cjRun('pier100', cases[i])); await exportAs('cj' + i); }
  const load = recalc(cases.map((_, i) => 'cj' + i), 'outcj');
  for(let i=0;i<cases.length;i++){
    const wb = await load('cj' + i), r = web[i], tag = JSON.stringify(cases[i]);
    assert.ok(Math.abs(Number(row(wb, '施工縫設計剪力強度 Vj').v) - r.cj.Vj/1000) < 0.05, `${tag}：Vj`);
    assert.ok(Math.abs(Number(row(wb, '施工縫剪力需求').v) - r.cj.Vreq/1000) < 0.05, `${tag}：需求`);
    assert.equal(row(wb, '施工縫不得設於主筋搭接處').judge, r.cj.inLap ? 'FAIL' : 'PASS', `${tag}：搭接處`);
  }
  await page.evaluate(() => applyPreset('bldg60'));
});

/* ---------------- 耐震梁：梁寬超出柱寬（§18.3.2.1(c)）、跨內降伏斷面 2h 加密（§18.3.4.1(b)） ---------------- */
const bSeis = set => page.evaluate(set => { document.getElementById('tabBeam').click(); applyBPreset('bldgBeam');
  for(const [id, v] of Object.entries(set || {})) document.getElementById(id).value = v;
  const B = renderBeam(), W = beamWidthOver(B.S);
  return {W, total: B.lay.total, zones: B.lay.zones.map(z => ({tag:z.tag, a:z.a, b:z.b, s:z.s, n:z.n})), h: B.S.h, s1: B.fin.sUse,
    unc: beamUncovered(B), hingeZ: B.hingeZ}; }, set);

test('耐震梁：梁寬超出柱寬每側 ≤ min(c₂, 0.75c₁)；跨內降伏斷面兩側 2h 以 s₁ 加密', async () => {
  const n = await bSeis({});
  assert.equal(n.W, null, 'c₂ 未輸入不檢核'); assert.ok(n.unc.some(t => /18\.3\.2\.1\(c\)/.test(t)));
  const ok = await bSeis({bC2: 40, bw: 60, bdHc: 60}), bad = await bSeis({bC2: 20, bw: 80, bdHc: 20});
  assert.equal(ok.W.ok, true); assert.ok(Math.abs(ok.W.ex - 10) < 1e-9);
  assert.equal(bad.W.ok, false, `每側 ${bad.W.ex} ＞ ${bad.W.lim}`);
  assert.ok(!ok.unc.some(t => /18\.3\.2\.1\(c\)/.test(t)), '已輸入 c₂ 不列未涵蓋');
  const h = await bSeis({bHingeX: 370});
  const z = h.zones.find(q => q.tag === '跨內加密區');
  assert.ok(z, '軸向配置有跨內加密區'); assert.equal(z.s, h.s1);
  assert.ok(z.a <= 370 - 2*h.h + 1e-6 && z.b >= 370 + 2*h.h - 1e-6, '涵蓋 x ± 2h');
  assert.ok(h.total > n.total, `支數增加（${n.total} → ${h.total}）`);
  await page.evaluate(() => { applyBPreset('bldgBeam'); document.getElementById('tabPylon').click(); });
});

test('耐震梁：Excel 支數與柱寬判定同網頁（LibreOffice 重算）', {skip: !hasSoffice && '未安裝 LibreOffice（soffice）'}, async () => {
  const cases = [{bHingeX: 370}, {bC2: 20, bw: 80, bdHc: 20}, {bC2: 40}];
  const web = [];
  for(let i=0;i<cases.length;i++){ web.push(await bSeis(cases[i])); await exportAs('bs' + i); }
  const load = recalc(cases.map((_, i) => 'bs' + i), 'outbs');
  for(let i=0;i<cases.length;i++){
    const wb = await load('bs' + i), r = web[i], tag = JSON.stringify(cases[i]);
    assert.equal(Number(row(wb, '全長合計數量（含搭接段加密）').v), r.total, `${tag}：支數`);
    assert.equal(row(wb, '耐震：梁寬超出柱寬每側（cm）').judge, r.W ? (r.W.ok ? 'PASS' : 'FAIL') : 'N/A', `${tag}：柱寬`);
  }
  await page.evaluate(() => { applyBPreset('bldgBeam'); document.getElementById('tabPylon').click(); });
});

/* ---------------- 基樁最小配置鋼筋樁長（表 18.10.5.7.1） ---------------- */
const pileRun = set => page.evaluate(set => { document.getElementById('tabPylon').click(); applyPreset('pileC120');
  for(const [id, v] of Object.entries(set || {})) document.getElementById(id).value = v;
  const M = render(), P = M.devPlan, PL = M.pileLen, c = P.checks.find(q => /基樁最小配置/.test(q.item));
  return {PL, lv: c ? c.lv : null, total: M.lay.total, pos0: Math.min(...M.lay.pos), L: M.opt.pile.L, D: M.s.D,
    ld: Math.ceil(P.D.ld - 1e-9), spl: P.splices.map(q => q.a), A0: P.A0, sched: devSchedule(P, M).map(r => r.id + ':' + r.st)}; }, set);

test('基樁：最小配置鋼筋樁長（第一、二類 max(L/2, 305, 3D, 撓曲長度 + l_d)；第三類全長，註[1]）與依最小長度終止', async () => {
  const a = await pileRun({});
  assert.equal(a.PL.Lmin, Math.max(a.L/2, 305, 3*a.D), '撓曲長度未輸入');
  assert.equal(a.lv, 'warn', '撓曲長度未輸入列注意'); assert.equal(a.PL.cut, false);
  const b = await pileRun({pileLf: 2000});
  assert.equal(b.PL.Lmin, Math.min(b.L, Math.ceil(Math.max(b.L/2, 305, 3*b.D, 2000 + b.ld) - 1e-9))); assert.equal(b.lv, 'ok');
  const c = await pileRun({pileLf: 1000, pileCut: '1'});
  assert.ok(c.PL.cut && Math.abs(c.PL.zb - (c.L - c.PL.Lmin)) < 1e-9, '終止段');
  assert.ok(c.pos0 >= c.PL.zb - 1e-6, '箍筋只配置於配筋段');
  assert.ok(c.total < a.total, `箍筋數減少（${a.total} → ${c.total}）`);
  assert.ok(c.A0 >= c.PL.zb - 1e-9 && c.spl.every(z => z >= c.PL.zb - 1e-9), '續接位於配筋段內');
  const s3 = await pileRun({pileSite: '3', pileRock: 600, pileCut: '1'});
  assert.equal(s3.PL.Lmin, Math.ceil(s3.L - Math.min(0.05*s3.L, 0.33*600) - 1e-9), '第三類註[1]');
  const s30 = await pileRun({pileSite: '3'});
  assert.equal(s30.PL.Lmin, s30.L, '第三類全長');
  await page.evaluate(() => applyPreset('bldg60'));
});

test('基樁：Excel 最小配置鋼筋樁長、箍筋與續接同網頁（LibreOffice 重算）', {skip: !hasSoffice && '未安裝 LibreOffice（soffice）'}, async () => {
  const cases = [{}, {pileLf: 1000, pileCut: '1'}, {pileSite: '3', pileRock: 600, pileCut: '1'}];
  const web = [];
  for(let i=0;i<cases.length;i++){ web.push(await pileRun(cases[i])); await exportAs('pl' + i); }
  const load = recalc(cases.map((_, i) => 'pl' + i), 'outpl');
  for(let i=0;i<cases.length;i++){
    const wb = await load('pl' + i), r = web[i], tag = JSON.stringify(cases[i]);
    assert.equal(Number(row(wb, '最小配置鋼筋樁長').v), r.PL.Lmin, `${tag}：Lmin`);
    assert.equal(Number(row(wb, '全長合計數量').v), r.total, `${tag}：箍筋數`);
    assert.equal(Number(row(wb, '續接處數（含增設）').v), r.spl.length, `${tag}：續接處數`);
  }
  await page.evaluate(() => applyPreset('bldg60'));
});

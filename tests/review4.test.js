/* 回歸測試（第四次全程式檢核 R-01～R-50）：
     R-01／R-02 √f'c ≤ 26.5（柱 V_c、柱與梁 T_th）；R-03 扭力縱筋 f_y ≤ 4200；R-04 M_2,min 兩軸分別檢核；
     R-06 θ 須 30°～60°；R-07 長期撓度 ρ′ 以受壓面寬；R-10 橋梁兩端無彎矩 C_m = 1；R-11 圓柱 h_x 取弦長；
     R-16 柱縱筋 f_y ≥ 5600 之 K_tr；R-17 擴頭 ψ_p；R-20 版 ρ_min ≥ 0.0018；R-21／R-22 耐震梁兩支連續筋與側向支撐；
     R-23 受壓筋圍束間距；R-27 實心斷面剪扭平方和開根號；R-32 受壓搭接 1/3 增量順序；Excel 新增判定列 */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs'), os = require('os'), path = require('path');
const ExcelJS = require('exceljs');
const core = require('../core/rc-core.js');
const {openApp} = require('./helpers');

let app, page;
test.before(async () => { app = await openApp(); page = app.page; });
test.after(async () => { if(app){ assert.deepEqual(app.errors, [], '頁面不得有 JavaScript 錯誤'); await app.close(); } });

test('R-01、R-02：柱 V_c 與柱、梁扭矩門檻之 √f\'c 取上限 26.5', async () => {
  assert.equal(core.sqfc(1050), 26.5); assert.equal(core.sqfc(400), 20);
  const r = await page.evaluate(() => { document.getElementById('tabPylon').click(); applyPreset('bldg60');
    document.getElementById('fc').value = 1000; const M = render(); const sh = M.shX;
    const out = {base: sh.V.base, bw: sh.bw, d: sh.d, lam: M.m.lam || 1, Tthr: sh.Tthr, phiv: M.m.phiv, Ag: M.Ag, B: M.s.B, H: M.s.H};
    document.getElementById('tabBeam').click(); applyBPreset('bldgBeam'); document.getElementById('bfc').value = 1000; const B = renderBeam();
    out.bT = B.sh.Tthr; out.bAcp = B.sh.Acp; out.bpcp = B.sh.pcp; out.bphiv = B.m.phiv;
    applyBPreset('bldgBeam'); document.getElementById('tabPylon').click(); applyPreset('bldg60'); return out; });
  assert.ok(Math.abs(r.base - 0.53*r.lam*26.5*r.bw*r.d) < 1e-6, '柱 V_c 基本項以 26.5 計');
  assert.ok(Math.abs(r.Tthr - r.phiv*0.265*r.lam*26.5*r.Ag*r.Ag/(2*(r.B+r.H))) < 1e-6, '柱 T_th 以 26.5 計');
  assert.ok(Math.abs(r.bT - r.bphiv*0.265*26.5*r.bAcp*r.bAcp/r.bpcp) < 1e-6, '梁 T_th 以 26.5 計');
});

test('R-03：扭力縱筋 A_ℓ 之 f_y 取 ≤ 4200', async () => {
  const r = await page.evaluate(() => { document.getElementById('tabBeam').click(); applyBPreset('bldgBeam');
    BLOADS = [{name:'T', Mpos:30, Mneg:20, Vu:25, Tu:6}]; drawBLoads(); BSIGN = 'pos';
    document.getElementById('bfy').value = 5600; const B = renderBeam(); const c = B.sh.all.find(q => q.tors);
    const out = {Al: c.Al, AtS: c.AtS, ph: B.sh.ph, fyt: B.m.fyt, th: B.S.theta};
    document.getElementById('bfy').value = 4200; applyBPreset('bldgBeam'); return out; });
  const cot = 1/Math.tan(r.th*Math.PI/180);
  assert.ok(Math.abs(r.Al - r.AtS*r.ph*(r.fyt/4200)*cot*cot) < 1e-9);
});

test('R-04：M_2,min 兩軸分別檢核——單向彎矩之組合另檢核弱軸最小偏心', async () => {
  const L = {Pu:300000, Mux:500000, Muy:0};
  const SL = {needX:true, needY:true, treat:'mag', bridge:false, sway:false, kX:1, kY:1, luX:600, luY:600, hX:60, hY:40,
    IgX:40*60**3/12, IgY:60*40**3/12, Ec:250000, beta:0.6, m12in:-1};
  core.slenderApply(SL, L, null);
  assert.ok(L.slAlt && L.slAlt.ax === 'y' && L.slAlt.Muy > 0, '另一情況於 Y 向施加 δ·M_2,min');
  assert.equal(L.Muy, 0, '基本情況 Y 向為 0');
  const pick = core.chooseSlender(L, K => K.Muy > 0 ? 2 : 1);
  assert.equal(pick, 2); assert.equal(L.minAx, 'y'); assert.ok(L.Muy > 0);
  const r = await page.evaluate(() => { document.getElementById('tabPylon').click(); applyPreset('bldg60');
    document.getElementById('lu').value = 1000; LOADS = [{name:'X', Pu:120, Mux:20, Muy:0, Vux:0, Vuy:10, Tu:0}]; drawLoads(); const M = render();
    const q = M.res[0], out = {minAx: q.Lk.minAx, Muy: q.Lk.Muy, alt: !!q.Lk.slAlt};
    applyPreset('bldg60'); return out; });
  assert.ok(r.alt, '網頁產生另一情況');
  if(r.minAx) assert.ok(r.Muy > 0);
});

test('R-06：扭矩桁架角 θ 超出 30°～60° 時停止計算', async () => {
  const r = await page.evaluate(() => { document.getElementById('tabPylon').click(); applyPreset('bldg60');
    document.getElementById('theta').value = 25; const M = render(); const a = (M.badIn||[]).some(b => b.id==='theta');
    document.getElementById('tabBeam').click(); applyBPreset('bldgBeam'); document.getElementById('bTheta').value = 65; const B = renderBeam();
    const b = (B.badIn||[]).some(x => x.id==='bTheta');
    document.getElementById('bTheta').value = 45; document.getElementById('theta').value = 45; applyBPreset('bldgBeam'); document.getElementById('tabPylon').click(); applyPreset('bldg60'); return {a, b}; });
  assert.ok(r.a && r.b);
});

test('R-07：T 梁正彎矩長期撓度之 ρ′ 以受壓面寬 b_e 計', async () => {
  const r = await page.evaluate(() => { document.getElementById('tabBeam').click(); applyBPreset('bldgBeam'); BSIGN = 'pos'; const B = renderBeam();
    const top = B.R.layers[B.R.layers.length-1], d = ctxDepth({g:beamShape(B.S,'pos'), layers:B.R.layers, m:B.m, sign:'pos'});
    return {rhoP: B.defl.rhoP, exp: top.A/(B.S.be*d), be: B.S.be, bw: B.S.bw}; });
  assert.ok(r.be > r.bw); assert.ok(Math.abs(r.rhoP - r.exp) < 1e-9);
});

test('R-10、R-11：橋梁兩端無彎矩 C_m = 1.0；圓柱 h_x 取相鄰主筋弦長', async () => {
  const L = {Pu:300000, Mux:0, Muy:0};
  const SL = {needX:true, needY:false, treat:'mag', bridge:true, sway:false, kX:1, kY:1, luX:800, luY:800, hX:100, hY:100,
    IgX:Math.PI*100**4/64, IgY:Math.PI*100**4/64, Ec:250000, beta:0.6, m12in:1, phic:0.65};
  core.slenderApply(SL, L, null);
  assert.equal(L.slx.Cm, 1, 'C_m = 1.0');
  const r = await page.evaluate(() => { document.getElementById('tabPylon').click(); applyPreset('circ80'); const M = render();
    const out = {hx: M.conf.eB.hx, rb: M.R.rb, n: M.R.bars.length, chk: M.chk.hxMaxB, on: M.chk.hxOn};
    applyPreset('bldg60'); return out; });
  assert.ok(Math.abs(r.hx - 2*r.rb*Math.sin(Math.PI/r.n)) < 1e-9); assert.equal(r.chk, r.hx); assert.equal(r.on, 'bldg');
});

test('R-16、R-17、R-32：柱縱筋 K_tr、擴頭 ψ_p、受壓搭接', async () => {
  const r = await page.evaluate(() => {
    const base = {fy:5600, fc:280, db:2.54, epoxy:false, lambda:1, excess:1, psiR:1, psiO:1, top:false, ccov:5, cSp:20, Atr:0, s:15, n:4, compConf:false, seis:false, joint:false};
    const a = devLength({...base, colMain:true}), b = devLength({...base});
    const h1 = devLength({...base, fy:4200, cSp:10, inJoint:false}), h2 = devLength({...base, fy:4200, cSp:10, inJoint:true}), h3 = devLength({...base, fy:4200, cSp:16, inJoint:false});
    const c = devLength({...base, fy:2800, fc:200, db:1.27});
    return {aBad:a.ktrBad, bBad:b.ktrBad, p1:h1.psiP, p2:h2.psiP, p3:h3.psiP, lapC:c.lapC}; });
  assert.ok(r.aBad && !r.bBad, '柱縱筋不論間距須 K_tr ≥ 0.5d_b');
  assert.equal(r.p1, 1.6); assert.equal(r.p2, 1); assert.equal(r.p3, 1);
  assert.ok(Math.abs(r.lapC - 40) < 1e-9, 'max(0.0073f_yd_b, 30)·4/3 = 40');
});

test('R-20、R-21、R-22、R-23：版 ρ_min、耐震梁兩支連續筋與側撐、受壓筋圍束間距', async () => {
  const r = await page.evaluate(() => { document.getElementById('tabBeam').click();
    applyBPreset('pondTop'); document.getElementById('bRhoTs').value = 0.0015; let B = renderBeam(); const slabBad = (B.badIn||[]).some(x => x.id==='bRhoTs');
    applyBPreset('bldgBeam'); document.getElementById('bNtop').value = 1; B = renderBeam();
    const one = beamAlertItems(B).some(a => a.lv==='bad' && /頂筋僅 1 根/.test(a.t));
    applyBPreset('bldgBeam'); document.getElementById('bNlegs').value = 2; B = renderBeam();
    const lat = beamAlertItems(B).some(a => a.lv==='bad' && /側向支撐不足/.test(a.t));
    applyBPreset('bldgBeam'); BSIGN = 'pos'; document.getElementById('bTopSize').value = '#4'; B = renderBeam();
    const comp = B.fin.cands.some(c => /受壓筋圍束/.test(c.t) && Math.abs(c.v - Math.min(16*BAR['#4'].d, 48*BAR[B.S.tieSize].d, Math.min(B.S.bw, B.S.h))) < 1e-9);
    applyBPreset('bldgBeam'); return {slabBad, one, lat, comp}; });
  assert.ok(r.slabBad, '版 ρ_min < 0.0018 停止計算'); assert.ok(r.one, '耐震梁頂筋 1 根不合格');
  assert.ok(r.lat, '5 根頂筋僅 2 肢時側撐不足'); assert.ok(r.comp, '受壓筋圍束間距列入候選');
});

test('R-27：柱實心斷面剪扭上限採平方和開根號', async () => {
  const r = await page.evaluate(() => { document.getElementById('tabPylon').click(); applyPreset('bldg60');
    LOADS = [{name:'T', Pu:100, Mux:10, Muy:0, Vux:0, Vuy:30, Tu:15}]; drawLoads(); const M = render(); const sh = M.shY;
    const out = {tors: sh.torsion, lhs: sh.lhs, vV: sh.Vdes/(sh.bw*sh.d), vT: sh.Tu*sh.ph/(1.7*sh.Aoh*sh.Aoh)};
    applyPreset('bldg60'); return out; });
  assert.ok(r.tors); assert.ok(Math.abs(r.lhs - Math.hypot(r.vV, r.vT)) < 1e-9);
});

test('Excel：新增 f_yt 上限、K_tr、D36 搭接、單獨 T 梁 h_f、耐震梁兩支連續筋與側撐判定列', async () => {
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'rcsd-rv4-'));
  const labels = async (fn) => { const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(fn); const out = [];
    wb.getWorksheet('檢核表').eachRow(r => out.push(String((r.getCell(1).value && r.getCell(1).value.richText) ? r.getCell(1).value.richText.map(t=>t.text).join('') : r.getCell(1).value || '').trim())); return out; };
  const exp = async (name, prep) => { await page.evaluate(prep); await page.evaluate(() => confAll(TAB==='beam' ? 'beam' : 'pylon'));
    const [dl] = await Promise.all([page.waitForEvent('download', {timeout:120000}), page.click('#btnXlsx')]); const f = path.join(TMP, name); await dl.saveAs(f); return f; };
  const c = await labels(await exp('c.xlsx', () => { document.getElementById('tabPylon').click(); applyPreset('bldg60'); render(); }));
  for(const l of ['橫向筋 fyt ≦ 剪力／扭力設計上限（kgf/cm²）', 'fy ≧ 5600 柱縱筋 Ktr ≧ 0.5db', '大於 D36 之鋼筋不得搭接（主筋 db，cm）'])
    assert.ok(c.includes(l), `柱檢核表缺「${l}」`);
  const b = await labels(await exp('b.xlsx', () => { document.getElementById('tabBeam').click(); applyBPreset('bldgBeam'); renderBeam(); }));
  for(const l of ['箍筋 fyt ≦ 剪力／扭力設計上限（kgf/cm²）', '單獨 T 梁翼板厚 hf ≧ 0.5bw（cm）', '耐震：上下兩面各至少兩支連續鋼筋（根）', '耐震梁主筋側向支撐（網頁值）', '大於 D36 之鋼筋不得搭接（db，cm）'])
    assert.ok(b.includes(l), `梁檢核表缺「${l}」`);
  await page.evaluate(() => { document.getElementById('tabPylon').click(); applyPreset('bldg60'); });
});

test('R-24：梁自動分排之每排根數以淨距 max(2.5, d_b, 4/3·d_agg) 估算', async () => {
  const r = await page.evaluate(() => { document.getElementById('tabBeam').click(); applyBPreset('bldgBeam');
    const g = (id, v) => { document.getElementById(id).value = v; };
    g('bw', 40); g('bCover', 4); g('bTieSize', '#4'); g('bBotSize', '#8'); g('bNbot', 6); g('bDagg', 2.5); const B = renderBeam();
    const out = {rows: B.R.rows, clrOk: B.clr.bot === null || B.clr.bot >= Math.max(2.5, BAR['#8'].d, 4/3*2.5) - 1e-9};
    applyBPreset('bldgBeam'); return out; });
  assert.equal(r.rows, 2, '6-#8 於 b_w = 40 須分兩排'); assert.ok(r.clrOk, '分排後底筋淨距合格');
});

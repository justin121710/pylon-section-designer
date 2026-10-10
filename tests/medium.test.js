/* 回歸測試（M1～M8）：期望值以規範原式於測試端獨立計算（只讀取輸入欄位），或以 tests/oracle.js 之獨立求解器驗算：
     M1 裂縫控制 c_c 為拉力主筋淨保護層（§24.3.2）
     M2 T 梁有效翼寬依表 6.3.2.1（淨跨 ℓ_n）與單獨 T 梁 §6.3.2.2
     M3 P-M 掃描下限：寬扁斷面低軸壓之 D/C 不得為 ∞
     M4 結果摘要 V_e 標題依設計依據
     M5 範本載重：M_ux 配 V_uy；懸臂橋墩 V ≈ M/L_v
     M6 載重表頂／底配對：M_1/M_2 由兩端彎矩求得
     M7 橋梁容量設計選項（靜載重軸力、彈性剪力上限、P_e）
     M8 橋梁 φ 依公路橋梁耐震設計規範 §5.3.2；計算書列明 φ 與載重因數依據 */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const {openApp} = require('./helpers');
const {makeSection} = require('./oracle');

let app, page;
test.before(async () => { app = await openApp(); page = app.page; });
test.after(async () => { if(app){ assert.deepEqual(app.errors, [], '頁面不得有 JavaScript 錯誤'); await app.close(); } });

const near = (act, exp, rel, msg) =>
  assert.ok(Math.abs(act - exp) <= rel*Math.abs(exp), `${msg}：實際 ${act}，期望 ${exp}（容許 ${rel*100}%）`);
const BAR_D = {'#3':0.953, '#4':1.270, '#5':1.590, '#6':1.910, '#7':2.220, '#8':2.540, '#9':2.870, '#10':3.220, '#11':3.580};

const col = (preset, inputs, loads) => page.evaluate(([preset, inputs, loads]) => {
  document.getElementById('tabPylon').click(); applyPreset(preset);
  for(const [id, v] of Object.entries(inputs || {})) document.getElementById(id).value = v;
  if(loads){ LOADS = loads.map(L => ({Vux:0, Vuy:0, Tu:0, ...L})); drawLoads(); }
  render();
}, [preset, inputs, loads]);
const beam = (preset, inputs) => page.evaluate(([preset, inputs]) => {
  document.getElementById('tabBeam').click(); applyBPreset(preset);
  for(const [id, v] of Object.entries(inputs || {})) document.getElementById(id).value = v;
  renderBeam();
  const g = id => document.getElementById(id).value, B = BMODEL;
  return {bw:+g('bw'), h:+g('bh'), hf:+g('bhf'), ln:+g('bLn'), sp:+g('bSpacing'), sw:+g('bSclear'), bf:+g('bBf'), cover:+g('bCover'),
          fy:+g('bfy'), tie:g('bTieSize'), be:B.S.be, slab:B.S.slab, sLim:B.crack.sLim, cc:B.crack.cc,
          bad: beamAlertItems(B).filter(a => a.lv === 'bad').map(a => a.t.replace(/<[^>]+>/g, ''))};
}, [preset, inputs]);

/* ------------------------------------------------------------------ */
test('M1：裂縫控制 c_c = 保護層 + d_t（梁）、= 保護層（版）', async () => {
  for(const k of ['deckT', 'bldgBeam', 'pondTop']){
    const r = await beam(k);
    const cc = r.slab ? r.cover : r.cover + BAR_D[r.tie], fs = 2/3*r.fy;
    near(r.cc, cc, 1e-12, `${k} c_c`);
    near(r.sLim, Math.min(38*2800/fs - 2.5*cc, 30*2800/fs), 1e-12, `${k} 間距上限`);
  }
});

test('M2：T 梁有效翼寬依表 6.3.2.1（ℓ_n 為淨跨），單獨 T 梁依 §6.3.2.2', async () => {
  let r = await beam('bldgBeam', {bEdge:'0'});
  near(r.be, r.bw + 2*Math.min(8*r.hf, (r.sp - r.bw)/2, r.ln/8), 1e-12, '內梁 b_e');
  r = await beam('deckT', {bEdge:'1', bSclear:150});
  near(r.be, r.bw + Math.min(6*r.hf, r.sw/2, r.ln/12), 1e-12, '邊梁 b_e');
  r = await beam('deckT', {bEdge:'2', bBf:300});           // h_f = 22 ≥ 0.5b_w = 22.5？否 → 不合格
  near(r.be, Math.min(r.bf, 4*r.bw), 1e-12, '單獨 T 梁 b_e');
  assert.equal(r.hf >= 0.5*r.bw, false);
  assert.ok(r.bad.some(t => /單獨 T 梁翼板過薄/.test(t)), '單獨 T 梁 h_f < 0.5b_w 應判不合格');
  r = await beam('deckT', {bEdge:'2', bBf:300, bhf:25});
  assert.ok(!r.bad.some(t => /單獨 T 梁翼板過薄/.test(t)));
  await page.evaluate(() => { applyBPreset('deckT'); document.getElementById('tabPylon').click(); });
});

test('M3：寬扁斷面低軸壓（600×60）之 D/C 為有限值，且與獨立求解一致', async () => {
  const L = {name:'A', Pu:5, Mux:30, Muy:0};
  await col('bldg60', {B:600, H:60, nB:10, nH:2, slTreat:'second'}, [L]);
  const r = await page.evaluate(() => { const M = MODEL, m = M.m;
    return {dc: M.res[0].dc, low: M.cvX.pts[0].Pn, Pnt: M.cvX.Pnt, phiLow: phiMnAtP(M.cvX, m, 0.5*m.phit*M.cvX.Pnt),
      sec: {B: M.s.B, H: M.s.H, fc: m.fc, fy: m.fy, Es: m.Es, ecu: m.ecu, phic: m.phic, phit: m.phit, pmaxf: m.pmaxf, bars: M.R.bars.map(b => ({x:b.x, y:b.y, A:b.A}))}}; });
  assert.ok(Number.isFinite(r.dc), `D/C = ${r.dc}`);
  assert.ok(r.low < 0.9*r.Pnt, '最低取樣點應接近純拉');
  assert.ok(r.phiLow > 0, '最低取樣點與純拉之間 φM_n 應內插、不得為 0');
  // 受壓區僅數 cm：獨立求解器沿 H 方向以 0.05 cm 纖維、c 細掃（取最內側交點）
  const O = makeSection({...r.sec, nx: 8, ny: 1200});
  near(r.dc, O.dcUniX(L.Pu*1000, L.Mux*1e5), 0.01, 'D/C 與獨立求解');
});

test('M4：結果摘要 V_e 標題依設計依據（建築 2M_pr/l_u、橋梁單柱 φ_oM_n/L_v）', async () => {
  await col('bldg60');
  let t = await page.evaluate(() => document.getElementById('summary').innerHTML);
  assert.match(t, /2M_\{pr\}/); assert.doesNotMatch(t, /\\phi_o M_n\}\{L_v\}/);
  { const x = await page.evaluate(() => document.getElementById('summary').textContent.replace(/\s+/g, ''));
    assert.ok(/pr,max/.test(x) && /1\.25/.test(x) && /[φϕ]=1/.test(x), '建築柱應標示 M_pr（1.25f_y、φ = 1）'); }
  await col('pier100');
  t = await page.evaluate(() => document.getElementById('summary').innerHTML);
  assert.match(t, /\\phi_o M_n\}\{L_v\}/);
  await col('bldg60', {ph:'0'});
  t = await page.evaluate(() => document.getElementById('summary').innerHTML);
  assert.match(await page.evaluate(() => document.getElementById('summary').textContent), /容量設計剪力 Ve不適用/);
});

test('M5：範本載重 M_ux 配 V_uy、M_uy 配 V_ux；懸臂橋墩地震剪力 V ≈ M/L_v', async () => {
  const P = await page.evaluate(() => Object.entries(PRESETS).map(([k, p]) => ({k, code:p.code, Lv:p.Lv, ends:p.phEnds, loads:p.loads})));
  for(const p of P) for(const L of p.loads){
    const mx = Math.abs(L.Mux), my = Math.abs(L.Muy), vx = Math.abs(L.Vux), vy = Math.abs(L.Vuy);
    if(Math.abs(mx - my) > 0.2*Math.max(mx, my) && Math.max(vx, vy) > 0)
      assert.equal(mx > my, vy > vx, `${p.k}「${L.name}」：M_ux 大者應配 V_uy 大`);
    if(/\(X\)|地震 X/.test(L.name)) assert.ok(vx >= vy, `${p.k}「${L.name}」：X 向地震應以 V_ux 為主`);
    if(/\(Y\)|地震 Y/.test(L.name)) assert.ok(vy >= vx, `${p.k}「${L.name}」：Y 向地震應以 V_uy 為主`);
    if(p.code === 'bridge' && (p.ends || '1') === '1' && /地震|最小軸力/.test(L.name)){
      const M = Math.max(mx, my), V = mx > my ? vy : vx;
      near(V, M/(p.Lv/100), 0.05, `${p.k}「${L.name}」V = M/L_v`);
    }
  }
});

test('M6：同名組合分列頂／底時，M_1/M_2 由兩端彎矩求得並用於 C_m 與可忽略門檻', async () => {
  const loads = [{name:'E', pos:'top', Pu:200, Mux:20, Muy:6}, {name:'E', pos:'bot', Pu:205, Mux:-12, Muy:4}];
  const inp = await page.evaluate(() => ({fc: +document.getElementById('fc').value}));
  await col('bldg60', {lu:600, slFrame:'braced', slTreat:'mag', slK:1, slKy:1, slM12:-1, slBeta:0.6}, loads);
  const r = await page.evaluate(() => ({pairs: MODEL.opt.slPairs, sl: colSlender(MODEL),
    rows: MODEL.res.map(q => ({name: q.L.name, dy: q.Lk.sly.d, Cmy: q.Lk.sly.Cm, Muy: q.Lk.Muy})), signs: LOADS.map(L => L.Mux)}));
  // M_ux：20 與 −12 異號 → 雙曲率，M_1/M_2 = +12/20 = 0.6；M_uy：6 與 4 同號 → 單曲率，−4/6
  near(r.pairs.m12x, 0.6, 1e-12, 'X 向 M_1/M_2'); near(r.pairs.m12y, -4/6, 1e-12, 'Y 向 M_1/M_2');
  near(r.sl.limX, Math.min(34 + 12*0.6, 40), 1e-12, 'X 向門檻'); near(r.sl.limY, 34 + 12*(-4/6), 1e-12, 'Y 向門檻');
  assert.equal(r.sl.needX, false, 'k l_u/r = 33.3 ≤ 40：X 向可忽略'); assert.equal(r.sl.needY, true, '33.3 > 26：Y 向須考慮');
  assert.deepEqual(r.signs, [20, -12], '載重表保留正負號');
  assert.deepEqual(r.rows.map(q => q.name), ['E（頂）', 'E（底）']);
  // 獨立計算 δ_ns（Y 向）：C_m = 0.6 − 0.4(−4/6)
  const Cm = 0.6 - 0.4*(-4/6), Ec = 12000*Math.sqrt(inp.fc), Ig = 60**4/12, Pc = Math.PI**2*(0.4*Ec*Ig/1.6)/600**2;
  for(const [i, L] of loads.entries()){
    near(r.rows[i].Cmy, Cm, 1e-12, `${r.rows[i].name} C_m`);
    near(r.rows[i].dy, Math.max(1, Cm/(1 - L.Pu*1000/(0.75*Pc))), 1e-9, `${r.rows[i].name} δ_ns`);
  }
  // CSV 第 8 欄「位置」
  const c = await page.evaluate(() => parseLoadCsv('E,200,20,6,0,0,0,頂\nE,205,-12,4,0,0,0,bot\nF,100,5,5,0,0,0', ['Pu','Mux','Muy','Vux','Vuy','Tu'], {pos:true}).ok.map(o => o.pos || ''));
  assert.deepEqual(c, ['top', 'bot', '']);
});

test('M7：橋梁單柱以靜載重軸力計 M_n、彈性剪力上限、指定 P_e', async () => {
  const L = [{name:'地震 Y', Pu:280, Mux:196, Muy:0, Vux:0, Vuy:25}];
  await col('pier100', {slTreat:'second'}, L);
  const r0 = await page.evaluate(() => ({VeX: MODEL.shX.Ve, VeY: MODEL.shY.Ve}));
  await col('pier100', {slTreat:'second', brMnP:'dead', brPD:150, brVelX:30, brPe:250}, L);
  const r = await page.evaluate(() => { const M = MODEL, m = M.m;
    return {sh: {Mn: M.shY.Mn, Ve: M.shY.Ve, VeX: M.shX.Ve, VeRawX: M.shX.VeRaw, capX: M.shX.VeCapped}, pe: M.conf.peTerm, fc: m.fc, Ag: m.Ag, Lv: M.opt.Lv, phio: M.opt.phio,
      sec: {B: M.s.B, H: M.s.H, fc: m.fc, fy: m.fy, Es: m.Es, ecu: m.ecu, phic: m.phic, phit: m.phit, pmaxf: m.pmaxf, bars: M.R.bars.map(b => ({x:b.x, y:b.y, A:b.A}))}}; });
  // 獨立計算：P_D = 150 tf 下之標稱 M_n（繞 X 軸，Y 向剪力）→ V_e = φ_oM_n/L_v
  const O = makeSection({...r.sec, nf: 120}), Mn = O.MnAt(150e3, Math.PI/2);
  near(r.sh.Mn, Mn, 0.01, 'P_D 下之 M_n');
  near(r.sh.Ve, r.phio*Mn/r.Lv, 0.01, 'Y 向 V_e = φ_oM_n(P_D)/L_v');
  assert.ok(r.sh.Ve < r0.VeY, '以靜載重軸力（較小）計之 V_e 應小於各組合取大');
  assert.equal(r.sh.capX, true); near(r.sh.VeX, 30e3, 1e-12, 'X 向 V_e 受彈性剪力上限');
  near(r.pe, 0.5 + 1.25*250e3/(r.fc*r.Ag), 1e-12, '0.5 + 1.25P_e/(f′cA_g)');
});

test('M8：橋梁 φ 依 §5.3.2（軸壓應力）時之 D/C 與獨立求解一致；計算書列明 φ 與載重因數依據', async () => {
  const L = [{name:'A', Pu:150, Mux:200, Muy:0}, {name:'B', Pu:600, Mux:180, Muy:0}];
  await col('pier100', {slTreat:'second', brPhiMode:'532'}, L);
  const r = await page.evaluate(() => { const M = MODEL, m = M.m;
    return {dc: M.res.map(q => q.dc), mode: m.phiMode, min: m.phiBrMin, cap: M.pX.cap,
      sec: {B: M.s.B, H: M.s.H, fc: m.fc, fy: m.fy, Es: m.Es, ecu: m.ecu, phic: m.phic, phit: m.phit, pmaxf: m.pmaxf, bars: M.R.bars.map(b => ({x:b.x, y:b.y, A:b.A}))}}; });
  assert.equal(r.mode, 'br532'); assert.equal(r.min, 0.70);
  const O0 = makeSection({...r.sec, nf: 120}), thr = Math.min(0.1*r.sec.fc*O0.Ag, O0.Pb(Math.PI/2));
  const O = makeSection({...r.sec, nf: 120, phiAx: {min: 0.70, thr}});
  near(r.cap, O.cap, 1e-3, 'φP_n,max = 0.70 × 0.80P_o');
  L.forEach((q, i) => near(r.dc[i], O.dc(q.Pu*1000, q.Mux*1e5, 1e-9), 0.012, `組合 ${q.name} D/C`));
  const t = await page.evaluate(() => { buildReport(MODEL); return document.getElementById('report').textContent; });
  assert.match(t, /強度折減因數與載重因數之依據/); assert.match(t, /公路橋梁耐震設計規範 §5\.3\.2/); assert.match(t, /公路橋梁設計規範第三章/);
  await col('bldg60');
  const t2 = await page.evaluate(() => { buildReport(MODEL); return document.getElementById('report').textContent; });
  assert.match(t2, /土木401-112 第五章/);
});

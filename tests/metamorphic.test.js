/* 變形測試（metamorphic testing）：不需「正確答案」，只檢查物理上必然成立的關係。
   獨立求解器與手算只能比對它們涵蓋的情形；變形關係可對任意輸入成立，用以抓出公式、分支、方向配對或內插之系統性錯誤。
     MR1～MR3  單調性：主筋根數、號數增加 → D/C 不增；保護層增加 → D/C 不減
     MR4       單調性：彎矩等比放大 → D/C 不減（比例接近線性）
     MR5       旋轉不變：B ↔ H、M_ux ↔ M_uy、V_ux ↔ V_uy 對調 → D/C、剪力需求、箍筋間距相同
     MR6、MR7  彎矩正負號、載重組合順序不影響結果
     MR8       連續性：軸力細掃描，D/C 無跳躍（含 0.1f'cA_g、平衡點、φ 過渡）
     MR9、MR10 剪力：V_u 增加 → V_s 需求不減；箍筋號數增加 → 強度需求間距增加
     MR11      梁：腹寬增加 → φM_n 不減；拉力筋增加（拉力控制範圍內）→ φM_n 增加
   各關係之容許誤差寫在測試內；違反即代表計算流程有錯（非數值噪音）。 */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const {openApp} = require('./helpers');

let app, page;
test.before(async () => { app = await openApp(); page = app.page; });
test.after(async () => { if(app){ assert.deepEqual(app.errors, [], '頁面不得有 JavaScript 錯誤'); await app.close(); } });

/* 柱：套用範本後改輸入、改載重，回傳控制 D/C 與剪力結果 */
const col = (preset, inputs = {}, loads = null) => page.evaluate(([preset, inputs, loads]) => {
  document.getElementById('tabPylon').click();
  applyPreset(preset);
  for(const [k, v] of Object.entries(inputs)) document.getElementById(k).value = v;
  if(loads) LOADS = loads.map(l => ({...l}));
  drawLoads(); render();
  const M = MODEL;
  return {dc: M.ctrlFlex.dc, dcs: M.res.map(q => q.dc), bad: !!(M.badIn && M.badIn.length),
    shX: {VsReq: M.shX.VsReq, gov: M.shX.s_gov}, shY: {VsReq: M.shY.VsReq, gov: M.shY.s_gov, sStr: M.shY.s_str},
    sReq: M.fin.gov.v, sUse: M.fin.sUse, Po: M.cvX.Po, cap: M.m.phic*M.m.pmaxf*M.cvX.Po, fcAg: M.m.fc*M.m.Ag};
}, [preset, inputs, loads]);

const BASE = [{name:'A', Pu:200, Mux:30, Muy:12, Vux:6, Vuy:14, Tu:0},
              {name:'B', Pu:90,  Mux:14, Muy:34, Vux:15, Vuy:5, Tu:0}];

test('MR1：主筋根數增加，D/C 不增（60×60、#8）', async () => {
  const r = [];
  for(const n of [3, 4, 5, 6]) r.push((await col('bldg60', {nB: n, nH: n}, BASE)).dc);
  for(let i = 1; i < r.length; i++) assert.ok(r[i] < r[i-1], `每面 ${i+3} 根之 D/C ${r[i]} 應小於 ${i+2} 根之 ${r[i-1]}`);
});

test('MR2：主筋號數增加，D/C 不增', async () => {
  const r = [];
  for(const b of ['#6', '#7', '#8', '#9', '#10']) r.push({b, dc: (await col('bldg60', {barSize: b}, BASE)).dc});
  for(let i = 1; i < r.length; i++) assert.ok(r[i].dc < r[i-1].dc, `${r[i].b} 之 D/C ${r[i].dc} 應小於 ${r[i-1].b} 之 ${r[i-1].dc}`);
});

test('MR3：保護層增加，D/C 不減（有效深度變小）', async () => {
  const r = [];
  for(const c of [3, 4, 5, 6, 8]) r.push({c, dc: (await col('bldg60', {covO: c}, BASE)).dc});
  for(let i = 1; i < r.length; i++) assert.ok(r[i].dc >= r[i-1].dc - 1e-9, `保護層 ${r[i].c} cm 之 D/C ${r[i].dc} 不應小於 ${r[i-1].c} cm 之 ${r[i-1].dc}`);
  assert.ok(r[r.length-1].dc > r[0].dc, '保護層 3 → 8 cm，D/C 應確實增加');
});

test('MR4：彎矩等比放大，D/C 不減', async () => {
  let prev = 0;
  for(const k of [0.25, 0.5, 0.75, 1, 1.25, 1.5]){
    const dc = (await col('bldg60', {}, BASE.map(l => ({...l, Mux: l.Mux*k, Muy: l.Muy*k})))).dc;
    assert.ok(dc >= prev - 1e-9, `彎矩 × ${k} 之 D/C ${dc} 不應小於前一級 ${prev}`);
    prev = dc;
  }
});

test('MR5：斷面旋轉 90°（B ↔ H、兩向載重對調）結果相同', async () => {
  const L = [{name:'A', Pu:180, Mux:45, Muy:15, Vux:7, Vuy:30, Tu:0}, {name:'B', Pu:80, Mux:20, Muy:28, Vux:18, Vuy:9, Tu:0}];
  const a = await col('bldg60', {B: 40, H: 90, nB: 3, nH: 6, ph: '0'}, L);
  const b = await col('bldg60', {B: 90, H: 40, nB: 6, nH: 3, ph: '0'}, L.map(l => ({...l, Mux: l.Muy, Muy: l.Mux, Vux: l.Vuy, Vuy: l.Vux})));
  for(let i = 0; i < a.dcs.length; i++)
    assert.ok(Math.abs(a.dcs[i] - b.dcs[i]) <= 2e-3*a.dcs[i], `組合 ${i+1}：D/C ${a.dcs[i]} 與旋轉後 ${b.dcs[i]} 應相同`);
  assert.ok(Math.abs(a.shY.VsReq - b.shX.VsReq) <= 1e-6*Math.max(1, a.shY.VsReq), 'Y 向剪力需求 ↔ 旋轉後 X 向');
  assert.ok(Math.abs(a.shX.VsReq - b.shY.VsReq) <= 1e-6*Math.max(1, a.shX.VsReq), 'X 向剪力需求 ↔ 旋轉後 Y 向');
  assert.ok(Math.abs(a.sReq - b.sReq) <= 1e-6*a.sReq, `箍筋需求間距 ${a.sReq} 與旋轉後 ${b.sReq} 應相同`);
});

test('MR6、MR7：彎矩正負號與載重組合順序不影響結果', async () => {
  const a = await col('bldg60', {}, BASE);
  const neg = await col('bldg60', {}, BASE.map(l => ({...l, Mux: -l.Mux, Muy: -l.Muy, Vux: -l.Vux, Vuy: -l.Vuy})));
  assert.ok(Math.abs(a.dc - neg.dc) < 1e-12, '彎矩反號 D/C 相同（雙軸對稱配筋）');
  const rev = await col('bldg60', {}, [...BASE].reverse());
  assert.ok(Math.abs(a.dc - rev.dc) < 1e-12, '組合順序反轉 D/C 相同');
  assert.ok(Math.abs(a.sUse - rev.sUse) < 1e-12, '組合順序反轉箍筋間距相同');
});

test('MR8：軸力細掃描 D/C 連續（含 0.1f′cA_g、平衡點、φ 過渡區）', async () => {
  const ref = await col('bldg60', {ph: '0'}, [{name:'0', Pu:0, Mux:20, Muy:0, Vux:0, Vuy:0, Tu:0}]);
  const P0 = -0.25*ref.Po/1000, P1 = 0.95*ref.cap/1000, n = 120, step = (P1 - P0)/n;
  const loads = Array.from({length: n + 1}, (_, i) => ({name: String(i), Pu: P0 + i*step, Mux: 20, Muy: 8, Vux: 0, Vuy: 0, Tu: 0}));
  const r = await col('bldg60', {ph: '0'}, loads);
  // 相鄰兩點差不得超過「該處斜率之合理上限」：以全域中位數差之 6 倍 + 0.01 為門檻（跳躍才會超過）
  const diffs = r.dcs.slice(1).map((v, i) => Math.abs(v - r.dcs[i]));
  const med = [...diffs].sort((a, b) => a - b)[Math.floor(diffs.length/2)];
  const lim = 6*med + 0.01, worst = diffs.reduce((a, d, i) => d > a.d ? {d, i} : a, {d: 0, i: 0});
  assert.ok(worst.d <= lim, `P_u = ${(P0 + worst.i*step).toFixed(1)} → ${(P0 + (worst.i+1)*step).toFixed(1)} tf 之 D/C 跳 ${worst.d.toFixed(4)}（門檻 ${lim.toFixed(4)}）`);
  assert.ok(r.dcs.every(Number.isFinite), '掃描範圍內 D/C 皆為有限值');
  // 0.1f′cA_g 兩側（±0.5%）
  const p = 0.1*ref.fcAg/1000;
  const s = await col('bldg60', {ph: '0'}, [0.995, 1.005].map(k => ({name: String(k), Pu: p*k, Mux: 20, Muy: 8, Vux: 0, Vuy: 0, Tu: 0})));
  assert.ok(Math.abs(s.dcs[0] - s.dcs[1]) < 0.01, `0.1f′cA_g 兩側 D/C ${s.dcs[0]}、${s.dcs[1]} 應連續`);
});

test('MR9：剪力需求隨 V_u 單調（非容量設計）', async () => {
  let prev = -1;
  for(const v of [5, 15, 25, 35, 45, 55]){
    const r = await col('bldg60', {ph: '0'}, [{name:'V', Pu:120, Mux:10, Muy:0, Vux:0, Vuy:v, Tu:0}]);
    assert.ok(r.shY.VsReq >= prev - 1e-9, `V_u = ${v} tf 之 V_s 需求 ${r.shY.VsReq} 不應小於前一級 ${prev}`);
    prev = r.shY.VsReq;
  }
  assert.ok(prev > 0, '大剪力時須有 V_s 需求');
});

test('MR10：箍筋號數增加，強度需求間距增加', async () => {
  // 比較強度需求間距 s = A_v f_yt d/V_s（不含 d/2 等幾何上限：箍筋加粗使 d 略減，d/2 上限會略降，屬正確行為）
  const L = [{name:'V', Pu:120, Mux:10, Muy:0, Vux:0, Vuy:45, Tu:0}];
  let prev = 0;
  for(const t of ['#3', '#4', '#5', '#6']){
    const r = await col('bldg60', {ph: '0', tieSize: t}, L);
    assert.ok(r.shY.sStr > prev, `${t} 之強度需求間距 ${r.shY.sStr} 應大於前一號數 ${prev}`);
    prev = r.shY.sStr;
  }
});

test('MR11：梁腹寬增加 φM_n 不減；拉力控制內拉力筋增加 φM_n 增加', async () => {
  const beam = inputs => page.evaluate(inputs => {
    document.getElementById('tabBeam').click(); applyBPreset('bldgJoist'); BSIGN = 'pos';
    for(const [k, v] of Object.entries(inputs)) document.getElementById(k).value = v;
    renderBeam(); return {phiMn: BMODEL.phiMn, et: BMODEL.sol.eps_t, ety: BMODEL.m.ety};
  }, inputs);
  let prev = 0;
  for(const bw of [25, 30, 40, 50, 60]){
    const r = await beam({bw});
    assert.ok(r.phiMn >= prev - 1e-6, `b_w = ${bw} 之 φM_n ${r.phiMn} 不應小於前一級 ${prev}`);
    prev = r.phiMn;
  }
  prev = 0;
  for(const n of [2, 3, 4]){
    const r = await beam({bNbot: n});
    assert.ok(r.et >= r.ety + 0.003, `${n} 根仍應為拉力控制（ε_t = ${r.et}）`);
    assert.ok(r.phiMn > prev, `${n} 根之 φM_n ${r.phiMn} 應大於前一級 ${prev}`);
    prev = r.phiMn;
  }
  await page.evaluate(() => document.getElementById('tabPylon').click());
});

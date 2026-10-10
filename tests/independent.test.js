/* 獨立驗證（H7）：以 tests/oracle.js 之獨立求解器（纖維離散、逐點二分，不引用工具計算函式）
   比對工具之雙軸 D/C（旋轉中性軸網格解）。斷面涵蓋實心方形、矩形（B ≠ H）、大型矩形與中空箱型；
   載重為可重現之亂數（軸壓、軸拉、任意彎矩方向，D/C 約 0.5～1.1）。 */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const {openApp} = require('./helpers');
const {makeSection, rng} = require('./oracle');

let app, page;
test.before(async () => { app = await openApp(); page = app.page; });
test.after(async () => { if(app){ assert.deepEqual(app.errors, [], '頁面不得有 JavaScript 錯誤'); await app.close(); } });

const CASES = [
  {name: '建築 RC 方柱 60×60', setup: () => applyPreset('bldg60')},
  {name: '矩形柱 40×90（B ≠ H）', setup: () => { applyPreset('bldg60'); const s = (i, v) => { document.getElementById(i).value = v; };
    s('B', 40); s('H', 90); s('nB', 3); s('nH', 6); render(); }},
  {name: '大型橋墩 200×300', setup: () => applyPreset('pier200')},
  {name: '中跨徑橋塔 箱型 300×400×40', setup: () => applyPreset('box300')}
];

for(const [ci, C] of CASES.entries()){
  test(`雙軸 D/C 與獨立求解一致：${C.name}`, async () => {
    const sec = await page.evaluate(src => { (0, eval)('(' + src + ')')();
      const M = MODEL;
      return {B: M.s.B, H: M.s.H, tw: M.s.tw, box: M.s.type === 'box', fc: M.m.fc, fy: M.m.fy, Es: M.m.Es, ecu: M.m.ecu,
              phic: M.m.phic, phit: M.m.phit, pmaxf: M.m.pmaxf, bars: M.R.bars.map(b => ({x: b.x, y: b.y, A: b.A}))};
    }, C.setup.toString());
    const O = makeSection({...sec, nf: 80}), R = rng(1234 + ci), loads = [];
    for(let k=0; k<3; k++){
      const Pu = k === 2 ? -0.15*Math.abs(O.Pt) : (0.05 + 0.6*R())*O.cap;           // 第 3 組為軸拉
      const psi = (10 + 70*R())*Math.PI/180, ux = Math.cos(psi), uy = Math.sin(psi);
      const Mr = (0.5 + 0.6*R())*O.Mcap(Pu, ux, uy);
      loads.push({Pu, Mux: Mr*ux, Muy: Mr*uy});
    }
    const tool = await page.evaluate(loads => loads.map(L => {
      const M = MODEL; return biaxialDC(M.cvX, M.cvY, M.pX, M.pY, M.m, L, M.S3);
    }).map(b => ({dc: b.dc, method: b.method})), loads);
    loads.forEach((L, i) => {
      const ex = O.dc(L.Pu, L.Mux, L.Muy);
      assert.equal(tool[i].method, '旋轉中性軸（精確解）');
      assert.ok(Math.abs(tool[i].dc/ex - 1) <= 0.015,
        `第 ${i+1} 組（P_u = ${(L.Pu/1000).toFixed(0)} tf，M_ux = ${(L.Mux/1e5).toFixed(0)}、M_uy = ${(L.Muy/1e5).toFixed(0)} tf·m）：工具 ${tool[i].dc.toFixed(4)}，獨立求解 ${ex.toFixed(4)}`);
    });
  });
}

test('單軸 P-M 曲線與獨立求解一致（θ = 0°、90°，40×90）', async () => {
  const sec = await page.evaluate(() => { applyPreset('bldg60'); const s = (i, v) => { document.getElementById(i).value = v; };
    s('B', 40); s('H', 90); s('nB', 3); s('nH', 6); render();
    const M = MODEL; return {B: M.s.B, H: M.s.H, fc: M.m.fc, fy: M.m.fy, phic: M.m.phic, phit: M.m.phit, pmaxf: M.m.pmaxf,
      bars: M.R.bars.map(b => ({x: b.x, y: b.y, A: b.A}))}; });
  const O = makeSection({...sec, nf: 120});
  for(const fr of [0.1, 0.3, 0.5]){
    const P = fr*O.cap;
    const t = await page.evaluate(P => ({x: phiMnAtP(MODEL.cvX, MODEL.m, P), y: phiMnAtP(MODEL.cvY, MODEL.m, P)}), P);
    const ox = O.Mcap(P, 1, 0), oy = O.Mcap(P, 0, 1);   // (M_x, M_y) 方向：繞 X 軸、繞 Y 軸
    assert.ok(Math.abs(t.x/ox - 1) < 0.01, `P = ${(P/1000).toFixed(0)} tf 繞 X：工具 ${(t.x/1e5).toFixed(2)}，獨立 ${(ox/1e5).toFixed(2)} tf·m`);
    assert.ok(Math.abs(t.y/oy - 1) < 0.01, `P = ${(P/1000).toFixed(0)} tf 繞 Y：工具 ${(t.y/1e5).toFixed(2)}，獨立 ${(oy/1e5).toFixed(2)} tf·m`);
  }
});

test('φ 過渡凹口：長跨徑橋塔箱型 600×800 近純繞 Y 軸，雙軸精確解不得偏不保守', async () => {
  /* φP_n 於 φ 過渡區隨 c 回降（re-entrant notch），同一軸力有多個分支；須取最內側（彎矩最小）者。
     修正前曲面網格剔除回降點後跨越缺口內插，於 P ≈ 0.5φP_n,max 高估容量約 15%，D/C 偏低約 4%。 */
  const r = await page.evaluate(() => { applyPreset('box600'); const M = MODEL, m = M.m;
    const P0 = 0.25*M.pX.cap, out = [];
    for(const [fx, fy] of [[1e-3, 1], [1, 1e-3]]){
      const Mx = fx*0.6*phiMnAtP(M.cvX, m, P0), My = fy*0.6*phiMnAtP(M.cvY, m, P0);
      out.push({Pu: P0, Mux: Mx, Muy: My, dc: biaxialDC(M.cvX, M.cvY, M.pX, M.pY, m, {Pu: P0, Mux: Mx, Muy: My}, M.S3).dc});
    }
    return {sec: {B: M.s.B, H: M.s.H, tw: M.s.tw, box: true, fc: m.fc, fy: m.fy, Es: m.Es, ecu: m.ecu, phic: m.phic, phit: m.phit, pmaxf: m.pmaxf,
      bars: M.R.bars.map(b => ({x: b.x, y: b.y, A: b.A}))}, out}; });
  const O = makeSection({...r.sec, nf: 120});
  for(const q of r.out){
    const ex = O.dc(q.Pu, q.Mux, q.Muy);
    assert.ok(q.dc >= 0.995*ex && q.dc <= 1.02*ex, `工具 ${q.dc.toFixed(4)}，獨立求解 ${ex.toFixed(4)}`);
  }
});

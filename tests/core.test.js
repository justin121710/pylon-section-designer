/* 核心純函式單元測試（U9）：直接 require core/rc-core.js，不開瀏覽器。
   期望值以規範原式於測試端獨立計算。 */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('crypto');
const C = require('../core/rc-core.js');

const near = (a, b, rel, msg) => assert.ok(Math.abs(a - b) <= rel*Math.abs(b) + 1e-12, `${msg}：實際 ${a}，期望 ${b}`);

test('β₁（§22.2.2.4.3）與 E_c（§19.2.2.1(b)）', () => {
  for(const [fc, b] of [[210, 0.85], [280, 0.85], [350, 0.80], [420, 0.75], [560, 0.65], [700, 0.65]]) near(C.beta1(fc), b, 1e-12, `β₁(${fc})`);
  near(C.Ecof(280), 12000*Math.sqrt(280), 1e-12, 'E_c');
});

test('φ 依 ε_t（表 21.2.2）與依軸壓（公路橋梁耐震設計規範 §5.3.2）', () => {
  const m = {ety: 4200/2.04e6, phic: 0.65, phit: 0.9};
  assert.equal(C.phiOf(m.ety, m), 0.65); assert.equal(C.phiOf(m.ety + 0.003, m), 0.9);
  near(C.phiOf(m.ety + 0.0015, m), 0.775, 1e-12, '過渡區中點');
  const mb = {phiBrMin: 0.70}, thr = 280e3;
  assert.equal(C.phiAxialBr(-10, thr, mb), 0.9); assert.equal(C.phiAxialBr(0, thr, mb), 0.9);
  // φ = 0.9 − 0.2·(φP_n)/thr 之解：φP_n = thr/2 時 φ = 0.8
  near(C.phiAxialBr(thr/2/0.8, thr, mb), 0.8, 1e-12, 'φP_n = thr/2');
  assert.equal(C.phiAxialBr(10*thr, thr, mb), 0.70);
  assert.equal(C.capPhi({phic: 0.65}), 0.65); assert.equal(C.capPhi({phic: 0.65, phiMode: 'br532', phiBrMin: 0.75}), 0.75);
  assert.equal(C.capPhi({phic: 0.65, phiAx: 0.55}), 0.55);
});

test('平衡軸力 P_b 於 ε_t = ε_ty 處內插', () => {
  const m = {ety: 0.002, fc: 280, Ag: 1e4};
  const pts = [{eps_t: 0.004, Pn: 100}, {eps_t: 0.003, Pn: 200}, {eps_t: 0.001, Pn: 400}];
  near(C.balancedPn(pts, m), 300, 1e-12, 'P_b');
  near(C.phiThrBr(pts, m), Math.min(0.1*280*1e4, 300), 1e-12, 'min(0.1f′cA_g, P_b)');
});

test('SHA-256 與 Node crypto 一致', () => {
  for(const t of ['', 'abc', 'RC 斷面設計'.repeat(50)]) assert.equal(C.sha256(t), crypto.createHash('sha256').update(t).digest('hex'));
});

test('頂／底配對之 M₁/M₂（§6.6.4.5.3：單曲率為負）', () => {
  const L = [{name:'E', pos:'top', Mux: 20, Muy: 6}, {name:'E', pos:'bot', Mux: -12, Muy: 4}, {name:'F', Mux: 5, Muy: 5}];
  const r = C.loadPairs(L, -0.5);
  near(r.rows[0].m12x, 0.6, 1e-12, '異號＝雙曲率'); near(r.rows[0].m12y, -4/6, 1e-12, '同號＝單曲率');
  assert.equal(r.rows[2].paired, false); assert.equal(r.rows[2].m12x, -0.5);
  near(r.m12y, -4/6, 1e-12, '各軸最不利'); near(r.m12x, -0.5, 1e-12, '含未配對之輸入值');
});

test('長細效應：建築無側移 δ_ns 與橋梁 δ_b', () => {
  const s = {type:'rect', B:60, H:60}, m = {Ec: 12000*Math.sqrt(280), phic: 0.65};
  const o = {code:'bldg', slSway:false, slK:1, slKy:1, lu:600, slLuY:0, slM12:-1, slTreat:'mag', slBeta:0.6, slQ:0};
  const SL = C.colSlenderCore(s, m, o);
  assert.equal(SL.needX, true);                                   // 600/18 = 33.3 > 34 − 12 = 22
  const Pu = 200e3, Ig = 60**4/12, Pc = Math.PI**2*(0.4*m.Ec*Ig/1.6)/600**2;
  const d = C.slenderDelta(SL, 'x', Pu, 10e5);
  near(d.d, 1/(1 - Pu/(0.75*Pc)), 1e-12, 'δ_ns'); near(d.Mc, d.d*Math.max(10e5, Pu*(1.5 + 0.03*60)), 1e-12, 'M_c');
  const SB = C.colSlenderCore(s, m, {...o, code:'bridge'});
  const db = C.slenderDelta(SB, 'x', Pu, 10e5), PcB = Math.PI**2*(m.Ec*Ig/2.5/1.6)/600**2;
  near(db.d, 1/(1 - Pu/(0.65*PcB)), 1e-12, 'δ_b');
});

test('T 梁有效翼寬（表 6.3.2.1、§6.3.2.2）與裂縫控制間距（§24.3.2）', () => {
  const S = {type:'T', bw:40, h:75, hf:15, ln:740, sSpacing:400, sClear:150, flPos:'int'};
  near(C.effFlange(S).be, 40 + 2*Math.min(120, 180, 92.5), 1e-12, '內梁');
  near(C.effFlange({...S, flPos:'edge'}).be, 40 + Math.min(90, 75, 740/12), 1e-12, '邊梁');
  near(C.effFlange({...S, flPos:'iso', bf:200}).be, 160, 1e-12, '單獨 T 梁 min(b_f, 4b_w)');
  assert.equal(C.effFlange({...S, flPos:'iso', bf:200}).hfOk, false);
  const B = {slab:false, cover:4, sign:'pos', bw:40, nTop:3, spBot:0}, R = {dt:1.27, r1:3, dbB:2.54, dbT:2.22};
  const r = C.beamCrack(B, R, {fy:4200}, {}), fs = 2800, cc = 5.27;
  near(r.cc, cc, 1e-12, 'c_c'); near(r.sLim, Math.min(38*2800/fs - 2.5*cc, 30*2800/fs), 1e-12, '間距上限');
});

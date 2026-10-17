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

/* ---------- 突變測試補強（tests/mutation.js）：逐一鎖定係數、門檻與分支，使計算核心之單一錯誤可被抓到 ---------- */
test('E_c 式 (a)(b) 之適用範圍（§19.2.2.1）、√f′c 上限、Mander 模數', () => {
  const Ea = (fc, wc) => Math.pow(wc, 1.5)*0.11*Math.sqrt(fc);
  near(C.Ecw(280, 2000, 1), Ea(280, 2000), 1e-12, 'w_c < 2160 用式 (a)');
  near(C.Ecw(280, 2160, 1), 12000*Math.sqrt(280), 1e-12, 'w_c = 2160 常重用式 (b)');
  near(C.Ecw(280, 2300, 0.85), Ea(280, 2300), 1e-12, 'λ < 1 用式 (a)');
  near(C.Ecw(280, 2300, 1), 12000*Math.sqrt(280), 1e-12, '常重');
  near(C.Ecw(280, 1440, 1), Ea(280, 1440), 1e-12, 'w_c = 1440 下限含');
  near(C.Ecw(280, 1430, 1), 12000*Math.sqrt(280), 1e-12, 'w_c < 1440 超出式 (a) 範圍');
  near(C.Ecw(280, 2560, 0.85), Ea(280, 2560), 1e-12, 'w_c = 2560 上限含');
  near(C.Ecw(280, 2570, 0.85), 12000*Math.sqrt(280), 1e-12, 'w_c > 2560 超出式 (a) 範圍');
  near(C.Ecw(350, 2380, 1), 12000*Math.sqrt(350), 1e-12, '2376 < w_c < 2560 常重仍用式 (b)');
  near(C.sqfc(900), 26.5, 1e-12, '√900 = 30 → 26.5'); near(C.sqfc(280), Math.sqrt(280), 1e-12, '√280');
  near(C.EcMander(280), 15000*Math.sqrt(280), 1e-12, 'Mander E_c');
  near(C.etTC({ety: 0.002}), 0.005, 1e-12, 'ε_ty + 0.003');
  assert.equal(C.vcZeroRatio(), 0.05);
});

test('耐震箍筋間距之 d_b 倍數（§18.4.5.3）', () => {
  assert.equal(C.seisKdb(2800), 6); assert.equal(C.seisKdb(4200), 6); assert.equal(C.seisKdb(4300), 5.5);
  assert.equal(C.seisKdb(5000), 5.5); assert.equal(C.seisKdb(5100), 5); assert.equal(C.seisKdb(5600), 5);
  assert.equal(C.seisKdb(4500), 5.5);
});

test('橋梁 φ 依軸壓之邊界（§5.3.2）與 P_b、門檻', () => {
  const mb = {phiBrMin: 0.70};
  near(C.phiAxialBr(0.5, 1, mb), 0.9/(1 + 0.2*0.5), 1e-12, '小軸壓');
  assert.equal(C.phiAxialBr(5, 0, mb), 0.9, 'P_thr = 0 視為無門檻');
  near(C.phiAxialBr(1, 0.5, mb), Math.max(0.7, 0.9/(1 + 0.2*2)), 1e-12, 'P_thr < 1');
  const m = {ety: 0.002, fc: 280, Ag: 1e4};
  near(C.balancedPn([{eps_t: 0.003, Pn: 100}, {eps_t: 0.001, Pn: 300}, {eps_t: 0.0005, Pn: 400}], m), 200, 1e-12, '首段跨越');
  near(C.balancedPn([{eps_t: 0.002, Pn: 100}, {eps_t: 0.001, Pn: 300}], m), 100, 1e-12, 'ε_t = ε_ty 恰在取樣點');
  near(C.balancedPn([{eps_t: 0.004, Pn: 100}, {eps_t: 0.002, Pn: 300}, {eps_t: 0.001, Pn: 500}], m), 300, 1e-12, 'ε_t = ε_ty 在後一點');
  assert.equal(C.balancedPn([{eps_t: 0.004, Pn: 100}, {eps_t: 0.003, Pn: 200}], m), null, '未跨越');
  assert.equal(C.balancedPn([{eps_t: 0.001, Pn: 100}, {eps_t: 0.0005, Pn: 200}], m), null, '全在壓力控制');
  near(C.phiThrBr([{eps_t: 0.003, P: 1e6}, {eps_t: 0.001, P: 1e6}], m), 0.1*280*1e4, 1e-12, '0.1f′cA_g 控制');
  near(C.phiThrBr([{eps_t: 0.003, P: 0}, {eps_t: 0.001, P: 0}], m), 0.1*280*1e4, 1e-12, 'P_b = 0 不採用');
  near(C.phiThrBr([{eps_t: 0.003, P: 0.5}, {eps_t: 0.001, P: 0.5}], m), 0.5, 1e-12, 'P_b 很小仍採用');
});

test('頂／底配對：一端彎矩為零時 M₁/M₂ = 0', () => {
  const r = C.loadPairs([{name:'E', pos:'top', Mux: 0, Muy: 8}, {name:'E', pos:'bot', Mux: 5, Muy: 0}], -1);
  assert.equal(Math.abs(r.rows[0].m12x), 0, '頂端 M_ux = 0'); assert.equal(Math.abs(r.rows[1].m12y), 0, '底端 M_uy = 0');
  assert.equal(Math.abs(r.rows[1].m12x), 0); assert.equal(Math.abs(r.rows[0].m12y), 0);
  const u = C.loadPairs([{name:'E', pos:'top', Muy: 8}, {name:'E', pos:'bot', Mux: 5}], -1);
  assert.equal(Math.abs(u.rows[0].m12x), 0, '未填視為 0'); assert.equal(Math.abs(u.rows[1].m12y), 0);
});

test('長細效應判定：迴轉半徑、門檻與旗標（§6.2.5、橋梁 §7.3.5.2）', () => {
  const m = {Ec: 250000, phic: 0.65};
  const o = {code:'bldg', slSway:false, slK:1, slKy:1, lu:600, slLuY:0, slM12:-1, slTreat:'mag', slBeta:0.6, slQ:0};
  const R = C.colSlenderCore({type:'rect', B:40, H:60}, m, o);
  near(R.rX, 18, 1e-12, 'r_x = 0.3H'); near(R.rY, 12, 1e-12, 'r_y = 0.3B');
  near(R.IgX, 40*60**3/12, 1e-12, 'I_gx'); near(R.IgY, 60*40**3/12, 1e-12, 'I_gy');
  near(R.lamX, 600/18, 1e-12, 'λ_x'); near(R.lamY, 600/12, 1e-12, 'λ_y'); near(R.lam, 50, 1e-12, 'λ = max');
  near(R.limX, 22, 1e-12, '34 + 12(−1)'); assert.equal(R.hX, 60); assert.equal(R.hY, 40);
  assert.equal(R.betaS, 0); assert.equal(R.kLow, false); assert.equal(R.kSwayLow, false); assert.equal(R.over100, false);
  const Ry = C.colSlenderCore({type:'rect', B:40, H:60}, m, {...o, slLuY: 300, slKy: 0.9});
  near(Ry.luY, 300, 1e-12, 'l_u,y 另給'); near(Ry.lamY, 0.9*300/12, 1e-12, 'λ_y 用 k_y、l_u,y'); assert.equal(Ry.kLow, true);
  const Rc = C.colSlenderCore({type:'circle', D:80}, m, o);
  near(Rc.rX, 20, 1e-12, 'r = 0.25D'); near(Rc.IgX, Math.PI*80**4/64, 1e-12, '圓 I_g'); assert.equal(Rc.hX, 80); assert.equal(Rc.hY, 80);
  const Rb = C.colSlenderCore({type:'box', B:300, H:400, tw:40}, m, o), A = 300*400 - 220*320;
  near(Rb.IgX, (300*400**3 - 220*320**3)/12, 1e-12, '箱型 I_gx'); near(Rb.IgY, (400*300**3 - 320*220**3)/12, 1e-12, '箱型 I_gy');
  near(Rb.rX, Math.sqrt(Rb.IgX/A), 1e-12, '箱型 r_x'); near(Rb.rY, Math.sqrt(Rb.IgY/A), 1e-12, '箱型 r_y');
  const Rs = C.colSlenderCore({type:'box', B:60, H:60, tw:35}, m, o);
  near(Rs.IgX, 60**4/12, 1e-12, '壁厚過大時內孔為 0');
  // 門檻：雙曲率 M1/M2 = 0.6 → 34 + 7.2 = 41.2 → 上限 40；橋梁無上限
  const Rd = C.colSlenderCore({type:'rect', B:60, H:60}, m, {...o, slM12: 0.6, lu: 1200});
  near(Rd.limX, 40, 1e-12, '建築上限 40'); near(Rd.limB, 34 + 12*0.6, 1e-12, 'limB');
  const RdB = C.colSlenderCore({type:'rect', B:60, H:60}, m, {...o, code:'bridge', slM12: 0.6, lu: 1200});
  near(RdB.limX, 41.2, 1e-12, '橋梁 34 − 12(M1b/M2b)，無上限');
  near(C.colSlenderCore({type:'rect', B:60, H:60}, m, {...o, slSway: true, slK: 1.5}).lim, 22, 1e-12, '有側移 22');
  // 兩軸 M1/M2 不同：lim 取兩軸較小、m12 取較小
  const Rm = C.colSlenderCore({type:'rect', B:60, H:60}, m, {...o, slM12x: -0.5, slM12y: 0.2});
  near(Rm.limX, 28, 1e-12, 'X'); near(Rm.limY, 36.4, 1e-12, 'Y'); near(Rm.lim, 28, 1e-12, 'lim = min'); near(Rm.m12, -0.5, 1e-12, 'm12 = min');
  // 等於門檻：建築 ≤ 可忽略；橋梁須 <
  const eq = {...o, slM12: 0, lu: 34*18};
  assert.equal(C.colSlenderCore({type:'rect', B:60, H:60}, m, eq).needX, false, '建築 λ = 34 可忽略');
  assert.equal(C.colSlenderCore({type:'rect', B:60, H:60}, m, {...eq, code:'bridge'}).needX, true, '橋梁 λ = 34 不可忽略');
  assert.equal(C.colSlenderCore({type:'rect', B:60, H:60}, m, {...eq, lu: 34*18 + 0.01}).needX, true, '建築 λ 略大於 34');
  const Ro = C.colSlenderCore({type:'rect', B:60, H:60}, m, {...o, code:'bridge', lu: 100*18});
  assert.equal(Ro.over100, false, 'λ = 100 不超過'); assert.equal(C.colSlenderCore({type:'rect', B:60, H:60}, m, {...o, code:'bridge', lu: 101*18}).over100, true);
  assert.equal(C.colSlenderCore({type:'rect', B:60, H:60}, m, {...o, lu: 101*18}).over100, false, '建築無 100 上限旗標');
  assert.equal(C.colSlenderCore({type:'rect', B:60, H:60}, m, {...o, slSway: true, slK: 1, slKy: 1.5}).kSwayLow, true, '有側移 k = 1');
  assert.equal(C.colSlenderCore({type:'rect', B:60, H:60}, m, {...o, slSway: true, slK: 1.2, slKy: 1.5}).kSwayLow, false);
  assert.equal(C.colSlenderCore({type:'rect', B:60, H:60}, m, {...o, slSway: false, slK: 1, slKy: 1}).kLow, false, 'k = 1 不算偏小');
  const okBoth = C.colSlenderCore({type:'rect', B:60, H:60}, m, {...o, lu: 300});
  assert.equal(okBoth.ok, true); assert.equal(C.colSlenderCore({type:'rect', B:40, H:60}, m, {...o, lu: 300}).ok, false, '一向不可忽略即 ok = false');
  assert.equal(C.colSlenderCore({type:'rect', B:60, H:60}, m, {...o, slBetaS: 0.3}).betaS, 0.3);
});

test('彎矩放大細節：Q 法、β_ds、C_m、最小彎矩與挫屈（§6.6.4、橋梁 §7.3.5.2）', () => {
  const m = {Ec: 12000*Math.sqrt(280), phic: 0.65}, Ig = 60**4/12;
  const base = {code:'bldg', slSway:true, slK:1.5, slKy:1.5, lu:600, slLuY:0, slM12:-1, slTreat:'mag', slBeta:0.6, slQ:0};
  const s = {type:'rect', B:60, H:60}, Pu = 150e3;
  const PcN = Math.PI**2*(0.4*m.Ec*Ig/1.6)/600**2, Cm = 0.6 - 0.4*0.5;
  // (b) 法：β_ds = 0
  const SL = C.colSlenderCore(s, m, base), PcS = Math.PI**2*(0.4*m.Ec*Ig)/(1.5*600)**2;
  const d = C.slenderDelta(SL, 'x', Pu, 20e5, 0.5);
  const dn = Math.max(1, Cm/(1 - Pu/(0.75*PcN))), ds = Math.max(1, 1/(1 - Pu/(0.75*PcS)));
  near(d.Cm, Cm, 1e-12, 'C_m = 0.6 − 0.4M1/M2'); near(d.dn, dn, 1e-12, 'δ_ns'); near(d.ds, ds, 1e-12, 'δ_s (b)'); near(d.d, ds*dn, 1e-12, 'δ = δ_s·δ_ns');
  // β_ds
  const d2 = C.slenderDelta(C.colSlenderCore(s, m, {...base, slBetaS: 0.5}), 'x', Pu, 20e5, 0.5);
  near(d2.PcS, Math.PI**2*(0.4*m.Ec*Ig/1.5)/(1.5*600)**2, 1e-12, 'β_ds');
  // Q 法：Q = 0.2 → 1.25；Q = 0.4 → 1.667 > 1.5 改用 (b)
  const dq = C.slenderDelta(C.colSlenderCore(s, m, {...base, slQ: 0.2}), 'x', Pu, 20e5, 0.5);
  near(dq.ds, 1.25, 1e-12, 'Q 法'); assert.equal(dq.qOver, false);
  const dq3 = C.slenderDelta(C.colSlenderCore(s, m, {...base, slQ: 1/3}), 'x', Pu, 20e5, 0.5);
  near(dq3.ds, 1.5, 1e-9, 'δ_s = 1.5 恰可用 Q 法'); assert.equal(dq3.qOver, false);
  const dq2 = C.slenderDelta(C.colSlenderCore(s, m, {...base, slQ: 0.4}), 'x', Pu, 20e5, 0.5);
  near(dq2.ds, ds, 1e-12, 'Q 法超過 1.5 改用 (b)'); assert.equal(dq2.qOver, true);
  near(C.slenderDelta(C.colSlenderCore(s, m, {...base, slQ: 0.01}), 'x', 1, 20e5, 0.5).ds, 1/(1 - 0.01), 1e-12, 'Q 小');
  // M2,min 控制
  const dm = C.slenderDelta(SL, 'x', Pu, 1e3, 0.5);
  near(dm.M2min, Pu*(1.5 + 0.03*60), 1e-12, 'M2,min'); assert.equal(dm.minGov, true); near(dm.Mc, dm.d*dm.M2min, 1e-12, 'M_c');
  assert.equal(C.slenderDelta(SL, 'x', Pu, 50e5, 0.5).minGov, false);
  // 軸拉、二階已含 → 不放大
  assert.equal(C.slenderDelta(SL, 'x', -10, 20e5).d, 1); assert.equal(C.slenderDelta(SL, 'x', 0, 20e5).d, 1);
  assert.equal(C.slenderDelta(C.colSlenderCore(s, m, {...base, slTreat:'second'}), 'x', Pu, 20e5).d, 1);
  assert.equal(C.slenderDelta(SL, 'x', 5, 20e5, 0.5).d > 1, true, '小軸壓仍放大（P_u > 0）');
  // 挫屈：P_u ≥ 0.75P_c → ∞；恰等於亦為 ∞
  assert.equal(C.slenderDelta(SL, 'x', 0.75*PcS, 20e5, 0.5).d, Infinity);
  // 橋梁：C_m 下限 0.4、M_u = 0 時 C_m = 1、有側移取 max(δ_b, δ_s)
  const B = C.colSlenderCore(s, m, {...base, code:'bridge', slSway:false, slK:1}), EI = m.Ec*Ig/2.5/1.6;
  near(C.slenderDelta(B, 'x', Pu, 20e5, 1).Cm, 0.4, 1e-12, 'C_m ≥ 0.4（雙曲率）');
  near(C.slenderDelta(B, 'x', Pu, 20e5, 0.5).Cm, 0.4, 1e-12, 'C_m = 0.6 − 0.4·0.5 = 0.4');
  near(C.slenderDelta(B, 'x', Pu, 20e5, -0.5).Cm, 0.8, 1e-12, 'C_m = 0.6 + 0.4·0.5');
  near(C.slenderDelta(B, 'x', Pu, 0, -0.5).Cm, 1, 1e-12, '兩端無彎矩 C_m = 1.0');
  const dB = C.slenderDelta(B, 'x', Pu, 20e5, -1), PcB = Math.PI**2*EI/600**2;
  near(dB.db, Math.max(1, 1/(1 - Pu/(0.65*PcB))), 1e-12, 'δ_b');
  const BS = C.colSlenderCore(s, m, {...base, code:'bridge', slK:2}), dS = C.slenderDelta(BS, 'x', Pu, 20e5, -1);
  near(dS.PcB, PcB, 1e-12, '有側移之 δ_b 以 k = 1 計'); near(dS.PcS, Math.PI**2*EI/(2*600)**2, 1e-12, 'δ_s 以 k 計');
  near(dS.d, Math.max(dS.db, dS.ds), 1e-12, 'max(δ_b, δ_s)'); assert.ok(dS.ds > dS.db);
  const BS2 = C.colSlenderCore(s, m, {...base, code:'bridge', slK:1.01}), dS2 = C.slenderDelta(BS2, 'x', Pu, 20e5, 1);
  near(dS2.d, Math.max(dS2.db, dS2.ds), 1e-12, '有側移 δ_b 控制時亦取 max');
  // 無側移建築 k 用輸入
  const NB = C.colSlenderCore(s, m, {...base, slSway:false, slK:0.8, slKy:0.8});
  near(C.slenderDelta(NB, 'x', Pu, 20e5, 0.5).PcN, Math.PI**2*(0.4*m.Ec*Ig/1.6)/(0.8*600)**2, 1e-12, '無側移 P_c 以 k 計');
  // 預設端彎矩比引數
  near(C.slenderDelta(NB, 'x', Pu, 20e5).Cm, 0.6 - 0.4*(-1), 1e-12, 'm12 預設用輸入值');
});

test('彎矩放大套用於兩軸（§6.6.4.5.4 M2,min 兩軸分別）與擇大', () => {
  const m = {Ec: 12000*Math.sqrt(280), phic: 0.65};
  const SL = C.colSlenderCore({type:'rect', B:60, H:60}, m, {code:'bldg', slSway:false, slK:1, slKy:1, lu:900, slLuY:0, slM12:-1, slTreat:'mag', slBeta:0.6, slQ:0});
  const Pu = 150e3, M2min = Pu*(1.5 + 0.03*60);
  // 兩向皆非零
  const L1 = {Pu, Mux: 20e5, Muy: 10e5}; C.slenderApply(SL, L1);
  near(L1.Mux, L1.slx.d*20e5, 1e-12, 'M_cx'); near(L1.Muy, L1.sly.d*Math.max(10e5, M2min), 1e-12, 'M_cy'); assert.equal(L1.slAlt, null);
  near(L1.Mux1, 20e5, 1e-12, '保留一階 M_ux'); near(L1.Muy1, 10e5, 1e-12, '保留一階 M_uy');
  // Muy = 0 → Y 向取 0，另一情況於 Y 施加 M2,min
  const L2 = {Pu, Mux: 20e5, Muy: 0}; C.slenderApply(SL, L2);
  assert.equal(L2.Muy, 0); assert.equal(L2.sly.minGov, false); near(L2.Mux, L2.slx.d*20e5, 1e-12, 'X');
  assert.equal(L2.slAlt.ax, 'y'); near(L2.slAlt.Muy, L2.sly.d*M2min, 1e-12, '另一情況 Y = δM2,min'); near(L2.slAlt.Mux, L2.slx.d*20e5, 1e-12, '另一情況 X 不加最小值');
  // Mux = 0
  const L3 = {Pu, Mux: 0, Muy: 20e5}; C.slenderApply(SL, L3);
  assert.equal(L3.Mux, 0); assert.equal(L3.slx.minGov, false); assert.equal(L3.slAlt.ax, 'x'); near(L3.slAlt.Mux, L3.slx.d*M2min, 1e-12, 'X = δM2,min');
  near(L3.slAlt.Muy, L3.sly.d*20e5, 1e-12, 'Y');
  // 兩向皆零 → 施加於 X、Y 取 0；另一情況施加於 Y
  const L4 = {Pu, Mux: 0, Muy: 0}; C.slenderApply(SL, L4);
  near(L4.Mux, L4.slx.d*M2min, 1e-12, '兩向皆零施加於 X'); assert.equal(L4.Muy, 0); assert.equal(L4.slAlt.ax, 'y'); assert.equal(L4.slAlt.Mux, 0);
  // 配對端彎矩比
  const L5 = {Pu, Mux: 20e5, Muy: 10e5}; C.slenderApply(SL, L5, {m12x: 0.5, m12y: -0.5});
  near(L5.slx.Cm, 0.4, 1e-12, '配對 C_m x'); near(L5.sly.Cm, 0.8, 1e-12, '配對 C_m y');
  // 不需放大之軸
  const SLx = C.colSlenderCore({type:'rect', B:60, H:60}, m, {code:'bldg', slSway:false, slK:1, slKy:1, lu:900, slLuY:300, slM12:-1, slTreat:'mag', slBeta:0.6, slQ:0});
  const L6 = {Pu, Mux: 20e5, Muy: 0}; C.slenderApply(SLx, L6); assert.equal(L6.slAlt, null, 'Y 向可忽略時不另施加');
  const L7 = {Pu, Mux: 0, Muy: 20e5}; C.slenderApply(SLx, L7); assert.equal(L7.slAlt.ax, 'x');
  // 挫屈之軸不產生另一情況
  const L8 = {Pu: 1e9, Mux: 20e5, Muy: 0}; C.slenderApply(SL, L8); assert.equal(L8.slAlt, null);
  // chooseSlender：另一情況 D/C 較大時採用並標 minAx
  const L9 = {Pu, Mux: 20e5, Muy: 0}; C.slenderApply(SL, L9);
  const r = C.chooseSlender(L9, K => ({dc: K.Muy > 0 ? 0.9 : 0.5}));
  near(r.dc, 0.9, 1e-12, '擇大'); assert.equal(L9.minAx, 'y'); near(L9.Muy, L9.slAlt.Muy, 1e-12, '採用另一情況之彎矩');
  const L10 = {Pu, Mux: 20e5, Muy: 0}; C.slenderApply(SL, L10);
  assert.equal(C.chooseSlender(L10, K => ({dc: K.Muy > 0 ? 0.5 : 0.5})).dc, 0.5); assert.equal(L10.minAx, null, '相等時維持基本情況');
  assert.equal(C.chooseSlender(L10, K => (K.Muy > 0 ? 0.7 : 0.6)), 0.7, '回傳數值亦可');
  const L11 = {Pu, Mux: 20e5, Muy: 10e5}; C.slenderApply(SL, L11); assert.equal(C.chooseSlender(L11, () => 0.3), 0.3, '無另一情況');
});

test('有效翼寬與裂縫控制之分支', () => {
  assert.equal(C.effFlange({type:'rect', bw:40}).be, 40, '矩形梁 b_e = b_w');
  const S = {type:'T', bw:40, hf:15, ln:740, sSpacing:30, sClear:0, flPos:'int'};
  near(C.effFlange(S).be, 40, 1e-12, '腹板淨距為 0（間距 < b_w）'); assert.equal(C.effFlange(S).sw, 0);
  near(C.effFlange({...S, sSpacing: 400}).cand[0].ov, 120, 1e-12, '8h_f'); near(C.effFlange({...S, sSpacing: 400}).cand[1].ov, 180, 1e-12, 's_w/2');
  near(C.effFlange({...S, flPos:'edge', sClear: 400}).cand[0].ov, 90, 1e-12, '6h_f');
  near(C.effFlange({...S, flPos:'edge', sClear: 400}).cand[1].ov, 200, 1e-12, '單側 s_w/2');
  near(C.effFlange({...S, flPos:'edge', sClear: 400, hf: 30}).be, 40 + 740/12, 1e-12, 'ℓ_n/12 控制');
  near(C.effFlange({...S, flPos:'int', sSpacing: 1000, hf: 30}).be, 40 + 2*740/8, 1e-12, 'ℓ_n/8 控制');
  near(C.effFlange({...S, flPos:'iso', bf: 120}).be, 120, 1e-12, '單獨 T 梁 b_f 控制');
  assert.equal(C.effFlange({...S, flPos:'iso', bf: 120, hf: 20}).hfOk, true, 'h_f = 0.5b_w');
  assert.equal(C.effFlange({...S, flPos:'iso', bf: 120, hf: 19}).hfOk, false);
  // 裂縫：版 c_c 不加箍筋；s2 = 30(2800/f_s)；負彎矩用頂筋
  const R = {dt:1.27, r1:3, dbB:2.54, dbT:2.22};
  const slab = C.beamCrack({slab:true, cover:3, sign:'pos', spBot: 20}, R, {fy:4200}, {});
  near(slab.cc, 3, 1e-12, '版 c_c = 保護層'); near(slab.s2, 30, 1e-12, 's2'); near(slab.s1, 38 - 7.5, 1e-12, 's1'); near(slab.sAct, 20, 1e-12, '版底筋間距');
  near(C.beamCrack({slab:true, cover:3, sign:'neg', spBot: 20, spTop: 15}, R, {fy:4200}, {}).sAct, 15, 1e-12, '版頂筋間距');
  near(C.beamCrack({slab:true, cover:3, sign:'neg', spBot: 20, spTop: 0}, R, {fy:4200}, {}).sAct, 20, 1e-12, '版未給頂筋間距');
  near(C.beamCrack({slab:true, cover:3, sign:'pos', spBot: 20}, R, {fy:2800}, {}).s2, 45, 1e-12, 'f_y 2800 之 s2');
  const bm = C.beamCrack({slab:false, cover:4, sign:'pos', bw:40, nTop:2}, R, {fy:4200}, {});
  near(bm.sAct, (40 - 2*5.27 - 3*2.54)/2 + 2.54, 1e-12, '梁底筋中心距');
  const ng = C.beamCrack({slab:false, cover:4, sign:'neg', bw:40, nTop:4}, R, {fy:4200}, {});
  near(ng.sAct, (40 - 2*5.27 - 4*2.22)/3 + 2.22, 1e-12, '梁頂筋中心距'); assert.equal(ng.nRow, 4);
  near(C.beamCrack({slab:false, cover:4, sign:'neg', bw:40, nTop:1}, R, {fy:4200}, {}).sAct, 40, 1e-12, '單根取 b_w');
  assert.equal(C.beamCrack({slab:true, cover:3, sign:'pos', spBot: 30}, R, {fy:4200}, {}).ok, true, 's = s_lim 合格');
  assert.equal(C.beamCrack({slab:true, cover:3, sign:'pos', spBot: 30.5}, R, {fy:4200}, {}).ok, false);
});

test('突變補強：箱型單向內孔、l_u,y、挫屈門檻、δ_s 與 Q 法上限、梁兩根筋間距', () => {
  const m = {Ec: 12000*Math.sqrt(280), phic: 0.65};
  const o = {code:'bldg', slSway:false, slK:1, slKy:1, lu:600, slLuY:0, slM12:-1, slTreat:'mag', slBeta:0.6, slQ:0};
  const b1 = C.colSlenderCore({type:'box', B:60, H:200, tw:35}, m, o);   // B − 2t_w < 0：無內孔
  near(b1.IgX, 60*200**3/12, 1e-12, 'H 向'); near(b1.rX, Math.sqrt(b1.IgX/12000), 1e-12, 'A');
  const b2 = C.colSlenderCore({type:'box', B:200, H:60, tw:35}, m, o);
  near(b2.IgY, 60*200**3/12, 1e-12, 'B 向'); near(b2.rY, Math.sqrt(b2.IgY/12000), 1e-12, 'A');
  near(C.colSlenderCore({type:'rect', B:60, H:60}, m, {...o, slLuY: 0.5}).luY, 0.5, 1e-12, 'l_u,y > 0 即採用');
  // P_u > 0.75P_c（超過 1 倍以上）仍為 ∞
  const base = {...o, slSway:true, slK:1.5, slKy:1.5}, SL = C.colSlenderCore({type:'rect', B:60, H:60}, m, base), Ig = 60**4/12;
  const PcS = Math.PI**2*(0.4*m.Ec*Ig)/(1.5*600)**2;
  assert.equal(C.slenderDelta(SL, 'x', 1.5*0.75*PcS, 20e5, 0.5).d, Infinity, 'P_u = 1.5 × 0.75P_c');
  // 建築有側移之 δ_ns 以 k = 1 計（C_m = 1 使 δ_ns > 1）
  const PcN = Math.PI**2*(0.4*m.Ec*Ig/1.6)/600**2, Pu = 150e3;
  near(C.slenderDelta(SL, 'x', Pu, 20e5, -1).dn, 1/(1 - Pu/(0.75*PcN)), 1e-12, 'δ_ns（k = 1）');
  // 橋梁有側移 δ_s = 1/(1 − P_u/φP_c)
  const BS = C.colSlenderCore({type:'rect', B:60, H:60}, m, {...base, code:'bridge', slK:2}), EI = m.Ec*Ig/2.5/1.6;
  near(C.slenderDelta(BS, 'x', Pu, 20e5, -1).ds, Math.max(1, 1/(1 - Pu/(0.65*Math.PI**2*EI/(2*600)**2))), 1e-12, '橋梁 δ_s');
  // Q ≥ 1（輸入已擋）與 δ_s,Q 介於 1.5～1.65 改用 (b) 法
  const dB = C.slenderDelta(SL, 'x', Pu, 20e5, 0.5).ds;
  const q12 = C.slenderDelta(C.colSlenderCore({type:'rect', B:60, H:60}, m, {...base, slQ: 1.2}), 'x', Pu, 20e5, 0.5);
  near(q12.ds, dB, 1e-12, 'Q ≥ 1 改用 (b)'); assert.equal(q12.qOver, true);
  const q36 = C.slenderDelta(C.colSlenderCore({type:'rect', B:60, H:60}, m, {...base, slQ: 0.36}), 'x', Pu, 20e5, 0.5);
  near(q36.ds, dB, 1e-12, 'δ_s,Q = 1.5625 > 1.5 改用 (b)'); assert.equal(q36.qOver, true);
  // 梁兩根拉力筋之中心距
  const R = {dt:1.27, r1:2, dbB:2.54, dbT:2.22};
  near(C.beamCrack({slab:false, cover:4, sign:'neg', bw:40, nTop:2}, R, {fy:4200}, {}).sAct, 40 - 2*5.27 - 2*2.22 + 2.22, 1e-12, '兩根頂筋');
});

/* 驗證案例（V&V benchmark）：手算期望值與工具讀值之對照表定義。
   - hand()：只用斷面幾何、CNS 560 鋼筋表與規範原式逐步計算（應變諧和以試算中性軸深度 c 求解，等同手算表格試算），
     不引用 index.html 任何計算函式；steps 為計算書式之代入過程（LaTeX），供驗證手冊（docs/tutorial/verify.html）列印。
   - tool：於頁面中執行之讀值運算式（字串），參數 H 為本案手算結果（例如在手算之 c_b 處讀取工具之斷面合力）。
   - tol：相對容許誤差。閉合式（面積、P_o、V_c 等）1e-4；應變諧和逐點 1e-4；工具曲線取樣內插 0.5%；雙軸數值解 1.5%。
   由 tests/verification.test.js（回歸測試）與 docs/tutorial/build.js（驗證手冊）共用；不需安裝套件。
   單位：kgf、cm（顯示時換算 tf、tf·m）。 */
'use strict';

const TF = 1000, TFM = 1e5;
const BAR = {'#3': {d: 0.953, A: 0.7133}, '#4': {d: 1.270, A: 1.267}, '#7': {d: 2.220, A: 3.871}, '#8': {d: 2.540, A: 5.067}};   // CNS 560 表 3
const Es = 2.04e6, ECU = 0.003;
const beta1 = fc => Math.max(0.65, Math.min(0.85, 0.85 - 0.05*(fc - 280)/70));       // §22.2.2.4.3
const n4 = x => Number(x).toLocaleString('en-US', {maximumFractionDigits: 4});
const r = (x, k = 2) => (Math.round(x*10**k)/10**k).toFixed(k);
const sci = (x, k = 4) => { const e = Math.floor(Math.log10(Math.abs(x))); return `${(x/10**e).toFixed(k)}\\times10^{${e}}`; };

/* 矩形柱周邊配筋座標（每面 nB、nH 根含角隅），dOut = 保護層 + 箍筋 + d_b/2 */
function rectBars(B, H, cov, tie, bar, nB, nH){
  const dOut = cov + BAR[tie].d + BAR[bar].d/2, hx = B/2 - dOut, hy = H/2 - dOut, A = BAR[bar].A, out = [];
  for(let i = 0; i < nB; i++){ const x = -hx + 2*hx*i/(nB - 1); out.push({x, y: hy, A}, {x, y: -hy, A}); }
  for(let j = 1; j < nH - 1; j++){ const y = -hy + 2*hy*j/(nH - 1); out.push({x: hx, y, A}, {x: -hx, y, A}); }
  return {bars: out, dOut, hx, hy};
}
/* 矩形斷面繞 X 軸（受壓緣 y = +H/2）之應變諧和：c → P_n、M_n、ε_t（受壓區內鋼筋扣 0.85f′c） */
function strain(sec, c, fy = sec.fy){
  const {B, H, fc, bars} = sec, b1 = beta1(fc), a = Math.min(b1*c, H), Cc = 0.85*fc*B*a;
  let P = Cc, M = Cc*(H/2 - a/2), dmax = -Infinity;
  for(const s of bars){
    const d = H/2 - s.y; let fs = Math.max(-fy, Math.min(fy, Es*ECU*(c - d)/c));
    if(d <= a) fs -= 0.85*fc;
    P += fs*s.A; M += fs*s.A*s.y; dmax = Math.max(dmax, d);
  }
  const et = ECU*(dmax - c)/c, ety = sec.fy/Es;
  const phi = et <= ety ? 0.65 : et >= ety + 0.003 ? 0.9 : 0.65 + 0.25*(et - ety)/0.003;   // 表 21.2.2（橫箍）
  return {c, a, Cc, P, M, et, phi, dt: dmax};
}
const bisect = (f, lo, hi, n = 200) => { for(let k = 0; k < n; k++){ const m = (lo + hi)/2; if(f(m) > 0) hi = m; else lo = m; } return (lo + hi)/2; };
const cAtP = (sec, P, fy) => bisect(c => strain(sec, c, fy).P - P, 1e-3, 20*sec.H);
const cAtE = (sec, e) => bisect(c => { const s = strain(sec, c); return e*s.P - s.M; }, 0.5, 20*sec.H);   // M/P = e（P > 0）
/* 單軸 D/C（定偏心射線，土木401-112 §22.4）：交點於 φP_n,max 截斷線上時取 P_u/φP_n,max */
function uniDC(sec, Pu, Mu){
  const Po = 0.85*sec.fc*(sec.B*sec.H - sec.Ast) + sec.fy*sec.Ast, cap = 0.65*0.8*Po;
  const c = cAtE(sec, Mu/Pu), s = strain(sec, c), phiP = s.phi*s.P;
  return phiP > cap ? {dc: Pu/cap, c, s, capped: true, cap} : {dc: Pu/phiP, c, s, capped: false, cap};
}

/* ======================= 案例 ======================= */
const COL60 = {B: 60, H: 60, fc: 280, fy: 4200, cov: 4, tie: '#4', bar: '#8', nB: 4, nH: 4};
function col60(over = {}){
  const g = {...COL60, ...over}, rb = rectBars(g.B, g.H, g.cov, g.tie, g.bar, g.nB, g.nH);
  return {...g, ...rb, Ast: rb.bars.reduce((a, b) => a + b.A, 0), Ag: g.B*g.H};
}

const CASES = [
/* ---------------------------------------------------------------- */
{id: 'V1', title: '建築矩形柱（非耐震）：P-M 特徵點、單軸 D/C、剪力',
 desc: '60×60 cm，f′c = 280、f_y = f_yt = 4200 kgf/cm²，12-D25（#8，每面 4 根），D13（#4）箍筋每根主筋皆設繫筋（每向 4 肢），保護層 4 cm（至箍筋外緣）。無側移、k = 1、l_u = 300 cm（長細效應可忽略），非塑鉸區。載重：P_u = 200 tf、M_ux = 30 tf·m、V_uy = 60 tf。',
 setup: {preset: 'bldg60', inputs: {ph: '0', slFrame: 'braced', slK: 1, slKy: 1, lu: 300, slM12: -1, tieLayout: 'all'},
         loads: [{name: 'V1', Pu: 200, Mux: 30, Muy: 0, Vux: 0, Vuy: 60, Tu: 0}]},
 hand(){
   const S = col60(), fc = S.fc, fy = S.fy, sq = Math.sqrt(fc);
   const Po = 0.85*fc*(S.Ag - S.Ast) + fy*S.Ast, cap = 0.65*0.80*Po;
   const dt = S.H/2 + S.hy, ety = fy/Es, cb = ECU*dt/(ECU + ety), bal = strain(S, cb);
   const c0 = cAtP(S, 0), pb = strain(S, c0);
   const U = uniDC(S, 200*TF, 30*TFM);
   // 剪力（沿 y，d = 受壓緣至受拉側鋼筋群形心）
   const tens = S.bars.filter(b => b.y < 0), ub = tens.reduce((a, b) => a + b.y*b.A, 0)/tens.reduce((a, b) => a + b.A, 0), d = S.H/2 - ub;
   const Nu = 200*TF, ax = Math.min(Nu/(6*S.Ag), 0.05*fc), Vc = Math.min((0.53*sq + ax)*S.B*d, 1.33*sq*S.B*d);
   const Vu = 60*TF, Vs = Math.max(0, Vu/0.75 - Vc), Av = 4*BAR['#4'].A, sStr = Av*4200*d/Vs;
   const sMin = Math.min(Av*4200/(0.2*sq*S.B), Av*4200/(3.5*S.B));
   const H = {cb, c0, Po, cap, bal, pb, U, d, Vc, Vs, sStr, sMin};
   H.steps = [
     ['鋼筋', `A_{st}=12\\times 5.067=${r(S.Ast,3)}\\ \\text{cm}^2,\\quad d'=4+1.27+\\tfrac{2.54}{2}=${r(S.dOut,2)}\\ \\text{cm}`],
     ['純壓（§22.4.2.2）', `P_o=0.85f'_c(A_g-A_{st})+f_yA_{st}=0.85(280)(3600-${r(S.Ast,2)})+4200(${r(S.Ast,2)})=${r(Po/TF,2)}\\ \\text{tf}`],
     ['軸力上限（表 22.4.2.1）', `\\phi P_{n,max}=0.65\\times0.80\\,P_o=${r(cap/TF,2)}\\ \\text{tf}`],
     ['平衡點', `c_b=\\dfrac{0.003\\,d_t}{0.003+f_y/E_s}=\\dfrac{0.003(${r(dt,2)})}{0.003+${r(ety,5)}}=${r(cb,3)}\\ \\text{cm}`],
     ['', `a=\\beta_1c_b=${r(bal.a,3)},\\ P_b=${r(bal.P/TF,2)}\\ \\text{tf},\\ M_b=${r(bal.M/TFM,2)}\\ \\text{tf}\\cdot\\text{m}`],
     ['純彎（試算 c 使 P_n = 0）', `c=${r(c0,3)}\\ \\text{cm},\\ M_n=${r(pb.M/TFM,2)}\\ \\text{tf}\\cdot\\text{m},\\ \\varepsilon_t=${r(pb.et,5)},\\ \\phi=${r(pb.phi,3)},\\ \\phi M_n=${r(pb.phi*pb.M/TFM,2)}`],
     ['單軸 D/C（e = M_u/P_u）', `e=\\dfrac{30}{200}=0.15\\ \\text{m}\\Rightarrow c=${r(U.c,3)}\\ \\text{cm},\\ P_n=${r(U.s.P/TF,2)},\\ M_n=${r(U.s.M/TFM,2)},\\ \\phi=${r(U.s.phi,3)}`],
     ['', `D/C=\\dfrac{P_u}{\\phi P_n}=\\dfrac{200}{${r(U.s.phi*U.s.P/TF,2)}}=${r(U.dc,4)}${U.capped ? '\\ (\\phi P_{n,max}\\text{ 截斷})' : ''}`],
     ['剪力有效深度', `d=\\tfrac{H}{2}+\\bar y_{t}=30+${r(-ub,3)}=${r(d,3)}\\ \\text{cm}\\ (\\text{受拉側 6 根形心})`],
     ['V_c（表 22.5.5.1）', `V_c=\\left(0.53\\sqrt{f'_c}+\\dfrac{N_u}{6A_g}\\right)b_wd=(0.53\\sqrt{280}+${r(ax,3)})(60)(${r(d,2)})=${r(Vc/TF,3)}\\ \\text{tf}`],
     ['V_s 需求', `V_s=\\dfrac{V_u}{\\phi}-V_c=\\dfrac{60}{0.75}-${r(Vc/TF,3)}=${r(Vs/TF,3)}\\ \\text{tf}`],
     ['強度需求間距', `s=\\dfrac{A_vf_{yt}d}{V_s}=\\dfrac{4(1.267)(4200)(${r(d,2)})}{${r(Vs,0)}}=${r(sStr,2)}\\ \\text{cm}`],
     ['最少剪力筋（§10.6.2.2）', `s\\le\\min\\left(\\dfrac{A_vf_{yt}}{0.2\\sqrt{f'_c}b_w},\\dfrac{A_vf_{yt}}{3.5b_w}\\right)=${r(sMin,2)}\\ \\text{cm}`]
   ];
   return H;
 },
 items: [
   {k: 'Ast', label: 'A_st', unit: 'cm²', hand: H => col60().Ast, tool: 'MODEL.Ast', tol: 1e-4},
   {k: 'Po', label: 'P_o', unit: 'tf', s: TF, hand: H => H.Po, tool: 'MODEL.cvX.Po', tol: 1e-4},
   {k: 'cap', label: 'φP_n,max', unit: 'tf', s: TF, hand: H => H.cap, tool: 'MODEL.pX.cap', tol: 1e-4},
   {k: 'Pb', label: 'P_b（c = c_b）', unit: 'tf', s: TF, hand: H => H.bal.P, tool: 'sectionAt(H.cb, MODEL.ctxX).Pn', tol: 1e-4},
   {k: 'Mb', label: 'M_b（c = c_b）', unit: 'tf·m', s: TFM, hand: H => H.bal.M, tool: 'sectionAt(H.cb, MODEL.ctxX).Mn', tol: 1e-4},
   {k: 'phiMn0', label: 'φM_n（P = 0）', unit: 'tf·m', s: TFM, hand: H => H.pb.phi*H.pb.M, tool: 'phiMnAtP(MODEL.cvX, MODEL.m, 0)', tol: 5e-3},
   {k: 'dc', label: 'D/C（單軸）', unit: '', hand: H => H.U.dc, tool: 'MODEL.res[0].dc', tol: 5e-3},
   {k: 'd', label: 'd（剪力）', unit: 'cm', hand: H => H.d, tool: 'MODEL.shY.d', tol: 1e-4},
   {k: 'Vc', label: 'V_c', unit: 'tf', s: TF, hand: H => H.Vc, tool: 'MODEL.shY.V.vc', tol: 1e-4},
   {k: 'Vs', label: 'V_s 需求', unit: 'tf', s: TF, hand: H => H.Vs, tool: 'MODEL.shY.VsReq', tol: 1e-4},
   {k: 'sStr', label: '強度需求間距', unit: 'cm', hand: H => H.sStr, tool: 'MODEL.shY.s_str', tol: 1e-4},
   {k: 'sMin', label: '最少剪力筋間距', unit: 'cm', hand: H => H.sMin, tool: 'MODEL.shY.s_minAv', tol: 1e-4}
 ]},
/* ---------------------------------------------------------------- */
{id: 'V2', title: '建築耐震柱（特殊抗彎矩構架）：容量設計剪力、圍束筋、加密區間距',
 desc: '同 V1 斷面，柱端加密區（兩端塑鉸），l_u = 320 cm；四組載重 P_u = 95～260 tf（範本「建築 RC 方柱」）。M_pr 取 1.25f_y、φ = 1，於各組合軸力區間內之最大值（§18.4.6.1.1）。',
 setup: {preset: 'bldg60', inputs: {}, loads: null},
 hand(){
   const S = col60(), fc = S.fc, sq = Math.sqrt(fc), lu = 320;
   const L = [{Pu: 260, Vu: 8}, {Pu: 205, Vu: 6}, {Pu: 210, Vu: 24}, {Pu: 95, Vu: 22}];   // 範本載重（Y 向剪力 V_uy）
   const Pmin = Math.min(...L.map(l => l.Pu))*TF, Pmax = Math.max(...L.map(l => l.Pu))*TF;
   let best = {M: -1};
   for(let k = 0; k <= 330; k++){ const P = Pmin + (Pmax - Pmin)*k/330, s = strain(S, cAtP(S, P, 1.25*S.fy), 1.25*S.fy); if(s.M > best.M) best = {M: s.M, P}; }
   const Mpr = best.M, Ve = 2*Mpr/lu;
   const tens = S.bars.filter(b => b.y < 0), d = S.H/2 - tens.reduce((a, b) => a + b.y*b.A, 0)/tens.reduce((a, b) => a + b.A, 0);
   const Av = 4*BAR['#4'].A;
   const rows = L.map(l => { const Nu = l.Pu*TF, Vdes = Math.max(Ve, l.Vu*TF), zero = Ve >= 0.5*Vdes && Nu < S.Ag*fc/20;
     const Vc = zero ? 0 : Math.min((0.53*sq + Math.min(Nu/(6*S.Ag), 0.05*fc))*S.B*d, 1.33*sq*S.B*d), Vs = Math.max(0, Vdes/0.75 - Vc);
     return {Nu, Vc, Vs, s: Av*4200*d/Vs}; });
   const g = rows.reduce((a, b) => b.s < a.s ? b : a);
   const bc = S.B - 2*S.cov, Ach = bc*bc, q1 = 0.3*(S.Ag/Ach - 1)*fc/4200, q2 = 0.09*fc/4200, sAsh = Av/(bc*Math.max(q1, q2));
   const hx = 2*S.hy/(S.nH - 1), so = Math.max(10, Math.min(15, 10 + (35 - hx)/3)), sLim = Math.min(S.B/4, 6*BAR['#8'].d, so);
   const H = {Mpr, PMpr: best.P, Ve, d, g, sAsh, q1, q2, hx, so, sLim, need: Math.min(g.s, sAsh, sLim)};
   H.steps = [
     ['M_pr（1.25f_y、φ = 1）', `\\text{軸力區間 }[${Pmin/TF},\\ ${Pmax/TF}]\\ \\text{tf 逐點試算 }c\\Rightarrow M_{pr,max}=${r(Mpr/TFM,3)}\\ \\text{tf}\\cdot\\text{m}\\ (P=${r(best.P/TF,1)}\\ \\text{tf})`],
     ['容量設計剪力（§18.4.6.1）', `V_e=\\dfrac{2M_{pr}}{l_u}=\\dfrac{2(${r(Mpr/TFM,3)})}{3.20}=${r(Ve/TF,3)}\\ \\text{tf}`],
     ['V_c 保留（§18.4.6.2.1）', `N_{u,min}=95\\ \\text{tf}\\ \\ge A_gf'_c/20=${r(S.Ag*fc/20/TF,1)}\\ \\text{tf}\\ \\Rightarrow\\ V_c\\text{ 不取 0}`],
     ['控制組合（N_u 最小）', `V_c=${r(g.Vc/TF,3)}\\ \\text{tf},\\ V_s=\\dfrac{V_e}{0.75}-V_c=${r(g.Vs/TF,3)}\\ \\text{tf},\\ s=\\dfrac{A_vf_{yt}d}{V_s}=${r(g.s,2)}\\ \\text{cm}`],
     ['圍束筋（表 18.4.5.4）', `b_c=60-2(4)=${bc},\\ A_{ch}=${Ach},\\ \\max\\left(0.3\\left(\\tfrac{A_g}{A_{ch}}-1\\right)\\tfrac{f'_c}{f_{yt}},\\ 0.09\\tfrac{f'_c}{f_{yt}}\\right)=\\max(${r(q1,5)},\\ ${r(q2,5)})`],
     ['', `s\\le\\dfrac{A_{sh}}{b_c\\,(\\cdot)}=\\dfrac{4(1.267)}{${bc}(${r(Math.max(q1,q2),5)})}=${r(sAsh,2)}\\ \\text{cm}`],
     ['間距上限（§18.4.5.3）', `h_x=${r(hx,2)},\\ s_o=10+\\tfrac{35-h_x}{3}\\ (10\\sim15)=${r(so,2)},\\ s\\le\\min(b/4,\\ 6d_b,\\ s_o)=${r(sLim,2)}\\ \\text{cm}`],
     ['需求間距', `s_{req}=\\min(${r(g.s,2)},\\ ${r(sAsh,2)},\\ ${r(sLim,2)})=${r(H.need,2)}\\ \\text{cm}`]
   ];
   return H;
 },
 items: [
   {k: 'Mpr', label: 'M_pr（軸力區間最大）', unit: 'tf·m', s: TFM, hand: H => H.Mpr, tool: 'MODEL.shY.Mn', tol: 5e-3},
   {k: 'Ve', label: 'V_e = 2M_pr/l_u', unit: 'tf', s: TF, hand: H => H.Ve, tool: 'MODEL.shY.Ve', tol: 5e-3},
   {k: 'Vc', label: 'V_c（控制組合）', unit: 'tf', s: TF, hand: H => H.g.Vc, tool: 'MODEL.shY.V.vc', tol: 1e-4},
   {k: 'sStr', label: '剪力強度需求間距', unit: 'cm', hand: H => H.g.s, tool: 'MODEL.shY.s_str', tol: 1e-2},
   {k: 'sAsh', label: 'A_sh 需求間距', unit: 'cm', hand: H => H.sAsh, tool: 'MODEL.conf.sAsh', tol: 1e-4},
   {k: 'so', label: 's_o', unit: 'cm', hand: H => H.so, tool: 'MODEL.opt.so', tol: 1e-4},
   {k: 'need', label: '加密區需求間距', unit: 'cm', hand: H => H.need, tool: 'MODEL.fin.gov.v', tol: 1e-2}
 ]},
/* ---------------------------------------------------------------- */
{id: 'V3', title: '矩形梁（雙筋、非耐震）：撓曲強度、最少鋼筋、剪力',
 desc: '30×60 cm，f′c = 280、f_y = f_yt = 4200，底筋 3-D22（#7）、頂筋 3-D22，D10（#3）肋筋 2 肢，保護層 4 cm。M_u⁺ = 14 tf·m、V_u = 15 tf（範本「建築 RC 小梁」）。',
 setup: {beam: 'bldgJoist', inputs: {}, sign: 'pos'},
 hand(){
   const b = 30, h = 60, fc = 280, fy = 4200, cov = 4, st = BAR['#3'].d, db = BAR['#7'].d, A = 3*BAR['#7'].A, b1 = beta1(fc);
   const d = h - cov - st - db/2, dp = cov + st + db/2;
   const F = c => { const a = Math.min(b1*c, h); let fsT = Math.max(-fy, Math.min(fy, Es*ECU*(c - dp)/c)); if(dp <= a) fsT -= 0.85*fc;
     const fsB = Math.max(-fy, Math.min(fy, Es*ECU*(c - d)/c)); return {a, Cc: 0.85*fc*b*a, fsT, fsB, P: 0.85*fc*b*a + fsT*A + fsB*A}; };
   const c = bisect(x => F(x).P, 1e-3, 5*h), q = F(c), Mn = q.Cc*(d - q.a/2) + q.fsT*A*(d - dp), et = ECU*(d - c)/c;
   const phi = et >= fy/Es + 0.003 ? 0.9 : 0.65 + 0.25*(et - fy/Es)/0.003;
   const a1 = A*fy/(0.85*fc*b), Mn1 = A*fy*(d - a1/2);   // 單筋簡算（對照）
   const AsMin = Math.max(0.8*Math.sqrt(fc)/fy, 14/fy)*b*d;
   const Vc = 0.53*Math.sqrt(fc)*b*d, Vs = Math.max(0, 15*TF/0.75 - Vc), sStr = 2*BAR['#3'].A*4200*d/Vs;
   const H = {d, c, a: q.a, fsT: q.fsT, et, phi, Mn, phiMn: phi*Mn, a1, Mn1, AsMin, Vc, Vs, sStr};
   H.steps = [
     ['有效深度', `d=60-4-0.953-\\tfrac{2.22}{2}=${r(d,3)}\\ \\text{cm},\\quad d'=${r(dp,3)}\\ \\text{cm},\\quad A_s=A'_s=3(3.871)=${r(A,3)}\\ \\text{cm}^2`],
     ['試算 c 使 ΣF = 0', `c=${r(c,3)}\\ \\text{cm},\\ a=\\beta_1c=${r(q.a,3)},\\ f'_s=${r(q.fsT,1)}\\ \\text{kgf/cm}^2\\ (\\text{受壓筋未降伏、已扣 }0.85f'_c)`],
     ['標稱彎矩', `M_n=0.85f'_cba\\left(d-\\tfrac a2\\right)+A'_sf'_s(d-d')=${r(Mn/TFM,3)}\\ \\text{tf}\\cdot\\text{m}`],
     ['φ（表 21.2.2）', `\\varepsilon_t=0.003\\,\\dfrac{d-c}{c}=${r(et,5)}\\ \\ge\\ \\varepsilon_{ty}+0.003\\Rightarrow\\phi=${r(phi,2)},\\ \\phi M_n=${r(phi*Mn/TFM,3)}\\ \\text{tf}\\cdot\\text{m}`],
     ['（對照）單筋簡算', `a=\\dfrac{A_sf_y}{0.85f'_cb}=${r(a1,3)},\\ M_n=A_sf_y\\left(d-\\tfrac a2\\right)=${r(Mn1/TFM,3)}\\ \\text{tf}\\cdot\\text{m（差 ${r((Mn/Mn1-1)*100,2)}\\%）}`],
     ['最少鋼筋（§9.6.1.2）', `A_{s,min}=\\max\\left(\\dfrac{0.8\\sqrt{f'_c}}{f_y},\\dfrac{14}{f_y}\\right)b_wd=${r(AsMin,3)}\\ \\text{cm}^2`],
     ['剪力', `V_c=0.53\\sqrt{f'_c}b_wd=${r(Vc/TF,3)}\\ \\text{tf},\\ V_s=\\tfrac{15}{0.75}-V_c=${r(Vs/TF,3)}\\ \\text{tf},\\ s=\\dfrac{2(0.7133)(4200)d}{V_s}=${r(sStr,2)}\\ \\text{cm}`]
   ];
   return H;
 },
 items: [
   {k: 'd', label: 'd', unit: 'cm', hand: H => H.d, tool: 'BMODEL.d', tol: 1e-4},
   {k: 'c', label: '中性軸 c', unit: 'cm', hand: H => H.c, tool: 'BMODEL.sol.c', tol: 1e-4},
   {k: 'et', label: 'ε_t', unit: '', hand: H => H.et, tool: 'BMODEL.sol.eps_t', tol: 1e-4},
   {k: 'Mn', label: 'M_n', unit: 'tf·m', s: TFM, hand: H => H.Mn, tool: 'BMODEL.sol.Mn', tol: 1e-4},
   {k: 'phiMn', label: 'φM_n', unit: 'tf·m', s: TFM, hand: H => H.phiMn, tool: 'BMODEL.phiMn', tol: 1e-4},
   {k: 'AsMin', label: 'A_s,min', unit: 'cm²', hand: H => H.AsMin, tool: 'BMODEL.AsMin', tol: 1e-4},
   {k: 'Vc', label: 'V_c', unit: 'tf', s: TF, hand: H => H.Vc, tool: 'BMODEL.sh.Vc', tol: 1e-4},
   {k: 'sStr', label: '強度需求間距', unit: 'cm', hand: H => H.sStr, tool: 'BMODEL.sh.sStr', tol: 1e-4}
 ]},
/* ---------------------------------------------------------------- */
{id: 'V4', title: '雙軸彎矩：旋轉中性軸精確解與獨立纖維求解器、Bresler 倒數式',
 desc: '同 V1 斷面（非耐震），載重 P_u = 150 tf、M_ux = 25 tf·m、M_uy = 20 tf·m。比較工具（旋轉中性軸、等效應力塊多邊形裁切）、獨立纖維求解器（tests/oracle.js，80×80 纖維、逐方向角掃描）與 Bresler 倒數式（手算，僅參考）。',
 setup: {preset: 'bldg60', inputs: {ph: '0', slFrame: 'braced', slK: 1, slKy: 1, lu: 300, slM12: -1, tieLayout: 'all'},
         loads: [{name: 'V4', Pu: 150, Mux: 25, Muy: 20, Vux: 0, Vuy: 0, Tu: 0}]},
 hand(){
   const S = col60(), Pu = 150*TF;
   const {makeSection} = require('./oracle');
   const O = makeSection({B: S.B, H: S.H, bars: S.bars, fc: S.fc, fy: S.fy, Es, nf: 80});
   const dcO = O.dc(Pu, 25*TFM, 20*TFM);
   const ux = uniDC(S, Pu, 25*TFM), uy = uniDC(S, Pu, 20*TFM);   // 方形對稱：繞 Y 軸同一曲線
   const Po = 0.85*S.fc*(S.Ag - S.Ast) + S.fy*S.Ast, phiPo = 0.65*Po;
   const Pnx = Pu/ux.dc, Pny = Pu/uy.dc, Pi = 1/(1/Pnx + 1/Pny - 1/phiPo), dcB = Pu/Pi;
   const H = {dcO, dcB, Pnx, Pny, phiPo, Pi};
   H.steps = [
     ['獨立纖維求解器', `80\\times80\\ \\text{纖維、方向角逐點掃描、等軸力輪廓二分}\\Rightarrow D/C=${r(dcO,4)}`],
     ['Bresler（§R22.4，參考）', `\\phi P_{nx}=${r(Pnx/TF,2)},\\ \\phi P_{ny}=${r(Pny/TF,2)},\\ \\phi P_o=0.65P_o=${r(phiPo/TF,2)}\\ \\text{tf}`],
     ['', `\\dfrac1{P_i}=\\dfrac1{\\phi P_{nx}}+\\dfrac1{\\phi P_{ny}}-\\dfrac1{\\phi P_o}\\Rightarrow P_i=${r(Pi/TF,2)}\\ \\text{tf},\\ D/C=${r(dcB,4)}`]
   ];
   return H;
 },
 items: [
   {k: 'dcO', label: 'D/C（工具 vs 纖維求解器）', unit: '', hand: H => H.dcO, tool: 'MODEL.res[0].dc', tol: 1.5e-2},
   {k: 'dcB', label: 'Bresler 參考值', unit: '', hand: H => H.dcB, tool: 'MODEL.res[0].bi.ref.dc', tol: 5e-3}
 ]},
/* ---------------------------------------------------------------- */
{id: 'V5', title: '長細效應：無側移構架彎矩放大（§6.6.4.5）',
 desc: '同 V1 斷面，無側移、k = 1、l_u = 600 cm、M_1/M_2 = −0.8（單曲率）、β_dns = 0.6。P_u = 300 tf、M_ux = M_2 = 20 tf·m。',
 setup: {preset: 'bldg60', inputs: {ph: '0', slFrame: 'braced', slK: 1, slKy: 1, lu: 600, slM12: -0.8, slBeta: 0.6, slTreat: 'mag', tieLayout: 'all'},
         loads: [{name: 'V5', Pu: 300, Mux: 20, Muy: 0, Vux: 0, Vuy: 0, Tu: 0}]},
 hand(){
   const fc = 280, Ec = 12000*Math.sqrt(fc), Ig = 60**4/12, lam = 600/(0.3*60), lim = Math.min(34 + 12*(-0.8), 40);
   const EI = 0.4*Ec*Ig/(1 + 0.6), Pc = Math.PI**2*EI/600**2, Cm = 0.6 - 0.4*(-0.8), Pu = 300*TF;
   const dl = Math.max(1, Cm/(1 - Pu/(0.75*Pc))), M2min = Pu*(1.5 + 0.03*60), M2 = Math.max(20*TFM, M2min), Mc = dl*M2;
   const H = {lam, lim, EI, Pc, Cm, dl, M2min, Mc};
   H.steps = [
     ['可忽略門檻（§6.2.5.1）', `\\dfrac{kl_u}{r}=\\dfrac{600}{0.3(60)}=${r(lam,2)}\\ >\\ 34+12\\dfrac{M_1}{M_2}=${r(lim,1)}\\Rightarrow\\text{須計長細效應}`],
     ['(EI)_eff（§6.6.4.4.4(a)）', `E_c=12000\\sqrt{280}=${n4(Math.round(Ec))},\\ (EI)_{eff}=\\dfrac{0.4E_cI_g}{1+\\beta_{dns}}=${sci(EI)}\\ \\text{kgf}\\cdot\\text{cm}^2`],
     ['P_c', `P_c=\\dfrac{\\pi^2(EI)_{eff}}{(kl_u)^2}=${r(Pc/TF,2)}\\ \\text{tf}`],
     ['C_m、δ（§6.6.4.5）', `C_m=0.6-0.4\\dfrac{M_1}{M_2}=${r(Cm,2)},\\ \\delta=\\dfrac{C_m}{1-P_u/0.75P_c}=${r(dl,4)}\\ \\ge1`],
     ['M_c', `M_{2,min}=P_u(1.5+0.03h)=${r(M2min/TFM,2)}\\ \\text{tf}\\cdot\\text{m},\\ M_c=\\delta\\max(M_2,M_{2,min})=${r(Mc/TFM,3)}\\ \\text{tf}\\cdot\\text{m}`]
   ];
   return H;
 },
 items: [
   {k: 'Pc', label: 'P_c', unit: 'tf', s: TF, hand: H => H.Pc, tool: 'MODEL.res[0].Lk.slx.PcN', tol: 1e-4},
   {k: 'dl', label: 'δ_ns', unit: '', hand: H => H.dl, tool: 'MODEL.res[0].Lk.slx.d', tol: 1e-4},
   {k: 'Mc', label: 'M_c', unit: 'tf·m', s: TFM, hand: H => H.Mc, tool: 'MODEL.res[0].Lk.Mux', tol: 1e-4}
 ]},
/* ---------------------------------------------------------------- */
{id: 'V6', title: '水工單位寬版（滯洪池側牆）：撓曲、剪力、裂縫間距與裂縫寬度',
 desc: '每公尺寬，厚 60 cm，f′c = 280、f_y = 4200，受拉面 D22@12.5、另一面 D22@15（版依間距換算每公尺根數 100/s：8 根、6.67 根），保護層 5 cm（版不配箍筋），環境「水工」。M_u = 38 tf·m、V_u = 17.5 tf；使用彎矩 M_D + M_L = 18 + 4 = 22 tf·m，容許裂縫寬度 0.20 mm（範本「滯洪池側牆」）。',
 setup: {beam: 'pondWall', inputs: {}, sign: 'pos'},
 hand(){
   const b = 100, h = 60, fc = 280, fy = 4200, cov = 5, db = BAR['#7'].d, As = 100/12.5*BAR['#7'].A, Asp = 100/15*BAR['#7'].A, b1 = beta1(fc);
   const d = h - cov - db/2, dp = cov + db/2;
   const F = c => { const a = Math.min(b1*c, h); let fsT = Math.max(-fy, Math.min(fy, Es*ECU*(c - dp)/c)); if(dp <= a) fsT -= 0.85*fc;
     const fsB = Math.max(-fy, Math.min(fy, Es*ECU*(c - d)/c)); return {a, Cc: 0.85*fc*b*a, fsT, P: 0.85*fc*b*a + fsT*Asp + fsB*As}; };
   const c = bisect(x => F(x).P, 1e-3, 5*h), q = F(c), Mn = q.Cc*(d - q.a/2) + q.fsT*Asp*(d - dp), et = ECU*(d - c)/c;
   const phi = et >= fy/Es + 0.003 ? 0.9 : 0.65 + 0.25*(et - fy/Es)/0.003;
   // 剪力：版不配剪力筋 → 表 22.5.5.1 式 (c)，λ_s = √(2/(1 + d/25)) ≤ 1
   const rw = As/(b*d), lamS = Math.min(1, Math.sqrt(2/(1 + d/25))), Vc = Math.min(2.12*lamS*Math.cbrt(rw)*Math.sqrt(fc)*b*d, 1.33*Math.sqrt(fc)*b*d);
   // 裂縫間距（表 24.3.2，f_s = 2/3 f_y，c_c = 保護層）
   const fs23 = 2/3*fy, sLim = Math.min(38*2800/fs23 - 2.5*cov, 30*2800/fs23);
   // 開裂轉換斷面（受壓筋以 (n − 1)A′_s 計）與 Gergely-Lutz
   const Ec = 12000*Math.sqrt(fc), n = Es/Ec;
   const x = bisect(y => b*y*y/2 + (n - 1)*Asp*(y - dp) - n*As*(d - y), 1e-3, h);
   const Icr = b*x**3/3 + (n - 1)*Asp*(x - dp)**2 + n*As*(d - x)**2, Ms = 22*TFM, fs = n*Ms*(d - x)/Icr;
   const dc = cov + db/2, A = 2*dc*12.5, beta = (h - x)/(d - x), w = 1.1e-5*beta*(fs*0.0980665)*Math.cbrt(dc*10*A*100);
   const H = {d, c, phiMn: phi*Mn, et, Vc, lamS, sLim, n, x, Icr, fs, w, beta, dc, A};
   H.steps = [
     ['有效深度', `d=60-5-\\tfrac{2.22}{2}=${r(d,3)}\\ \\text{cm},\\ A_s=\\tfrac{100}{12.5}(3.871)=${r(As,3)},\\ A'_s=\\tfrac{100}{15}(3.871)=${r(Asp,3)}\\ \\text{cm}^2/\\text{m}`],
     ['撓曲', `c=${r(c,3)},\\ \\varepsilon_t=${r(et,5)},\\ \\phi=${r(phi,2)},\\ \\phi M_n=${r(phi*Mn/TFM,3)}\\ \\text{tf}\\cdot\\text{m/m}\\ \\ge M_u=38`],
     ['V_c（表 22.5.5.1 式 (c)）', `\\lambda_s=\\sqrt{\\tfrac{2}{1+d/25}}=${r(lamS,4)},\\ \\rho_w=${r(rw,5)},\\ V_c=2.12\\lambda_s\\rho_w^{1/3}\\sqrt{f'_c}\\,bd=${r(Vc/TF,3)}\\ \\text{tf/m}`],
     ['裂縫間距（表 24.3.2）', `s\\le\\min\\left(38\\tfrac{2800}{f_s}-2.5c_c,\\ 30\\tfrac{2800}{f_s}\\right)=${r(sLim,2)}\\ \\text{cm}\\ (f_s=\\tfrac23f_y)`],
     ['開裂斷面', `n=\\dfrac{E_s}{E_c}=${r(n,3)},\\ x=${r(x,3)}\\ \\text{cm},\\ I_{cr}=${n4(Math.round(Icr))}\\ \\text{cm}^4,\\ f_s=\\dfrac{nM_s(d-x)}{I_{cr}}=${r(fs,1)}\\ \\text{kgf/cm}^2`],
     ['裂縫寬度（Gergely-Lutz）', `\\beta=\\dfrac{h-x}{d-x}=${r(beta,3)},\\ d_c=${r(dc,2)}\\ \\text{cm},\\ A=2d_cs=${r(A,2)}\\ \\text{cm}^2,\\ w=1.1\\times10^{-5}\\beta f_s\\sqrt[3]{d_cA}=${r(w,3)}\\ \\text{mm}`]
   ];
   return H;
 },
 items: [
   {k: 'd', label: 'd', unit: 'cm', hand: H => H.d, tool: 'BMODEL.d', tol: 1e-4},
   {k: 'phiMn', label: 'φM_n', unit: 'tf·m', s: TFM, hand: H => H.phiMn, tool: 'BMODEL.phiMn', tol: 1e-4},
   {k: 'Vc', label: 'V_c（式 (c)）', unit: 'tf', s: TF, hand: H => H.Vc, tool: 'BMODEL.sh.Vc', tol: 1e-4},
   {k: 'sLim', label: '裂縫控制間距上限', unit: 'cm', hand: H => H.sLim, tool: 'BMODEL.crack.sLim', tol: 1e-4},
   {k: 'x', label: '開裂中性軸 x', unit: 'cm', hand: H => H.x, tool: 'BMODEL.icr.x', tol: 1e-4},
   {k: 'fs', label: '使用鋼筋應力 f_s', unit: 'kgf/cm²', hand: H => H.fs, tool: 'BMODEL.cw.fs', tol: 1e-4},
   {k: 'w', label: '裂縫寬度 w', unit: 'mm', hand: H => H.w, tool: 'BMODEL.cw.w', tol: 1e-4}
 ]}
];

/* 驗證範圍端點（f'c 210、420）：同 V1 斷面之 P-M 特徵點與單軸 D/C（β₁ = 0.85、0.75） */
function pmCase(id, fc){
  return {id, title: `驗證範圍端點 f′c = ${fc}：P-M 特徵點與單軸 D/C（β₁ = ${beta1(fc).toFixed(2)}）`,
   desc: `同 V1 斷面與載重（P_u = 200 tf、M_ux = 30 tf·m），f′c = ${fc} kgf/cm²。`,
   setup: {preset: 'bldg60', inputs: {fc, ph: '0', slFrame: 'braced', slK: 1, slKy: 1, lu: 300, slM12: -1, tieLayout: 'all'},
           loads: [{name: id, Pu: 200, Mux: 30, Muy: 0, Vux: 0, Vuy: 0, Tu: 0}]},
   hand(){
     const S = col60({fc}), Po = 0.85*fc*(S.Ag - S.Ast) + S.fy*S.Ast, dt = S.H/2 + S.hy, cb = ECU*dt/(ECU + S.fy/Es), bal = strain(S, cb);
     const c0 = cAtP(S, 0), pb = strain(S, c0), U = uniDC(S, 200*TF, 30*TFM);
     const H = {Po, cb, bal, pb, U};
     H.steps = [
       ['β₁ 與 P_o', `\\beta_1=${beta1(fc).toFixed(2)},\\ P_o=0.85(${fc})(3600-${r(S.Ast,2)})+4200(${r(S.Ast,2)})=${r(Po/TF,2)}\\ \\text{tf}`],
       ['平衡點', `c_b=${r(cb,3)}\\ \\text{cm},\\ P_b=${r(bal.P/TF,2)}\\ \\text{tf},\\ M_b=${r(bal.M/TFM,2)}\\ \\text{tf}\\cdot\\text{m}`],
       ['純彎', `c=${r(c0,3)},\\ \\phi M_n=${r(pb.phi*pb.M/TFM,2)}\\ \\text{tf}\\cdot\\text{m}`],
       ['單軸 D/C', `c=${r(U.c,3)},\\ \\phi P_n=${r(U.s.phi*U.s.P/TF,2)}\\ \\text{tf},\\ D/C=${r(U.dc,4)}`]
     ];
     return H;
   },
   items: [
     {k: 'Po', label: 'P_o', unit: 'tf', s: TF, hand: H => H.Po, tool: 'MODEL.cvX.Po', tol: 1e-4},
     {k: 'Pb', label: 'P_b', unit: 'tf', s: TF, hand: H => H.bal.P, tool: 'sectionAt(H.cb, MODEL.ctxX).Pn', tol: 1e-4},
     {k: 'Mb', label: 'M_b', unit: 'tf·m', s: TFM, hand: H => H.bal.M, tool: 'sectionAt(H.cb, MODEL.ctxX).Mn', tol: 1e-4},
     {k: 'phiMn0', label: 'φM_n（P = 0）', unit: 'tf·m', s: TFM, hand: H => H.pb.phi*H.pb.M, tool: 'phiMnAtP(MODEL.cvX, MODEL.m, 0)', tol: 5e-3},
     {k: 'dc', label: 'D/C（單軸）', unit: '', hand: H => H.U.dc, tool: 'MODEL.res[0].dc', tol: 5e-3}
   ]};
}
CASES.push(pmCase('V7', 210), pmCase('V8', 420));

/* 於 Playwright page 執行一個案例：套用範本與輸入 → 讀取各項工具值。回傳 [{k, label, unit, hand, tool, rel, tol, ok}] */
async function runCase(page, C){
  const H = C.hand();
  const hJSON = JSON.parse(JSON.stringify(H, (k, v) => k === 'steps' ? undefined : v));
  const tool = await page.evaluate(([setup, exprs, H]) => {
    if(setup.beam){
      document.getElementById('tabBeam').click(); applyBPreset(setup.beam);
      for(const [k, v] of Object.entries(setup.inputs)) document.getElementById(k).value = v;
      BSIGN = setup.sign || 'pos'; renderBeam();
    }else{
      document.getElementById('tabPylon').click(); applyPreset(setup.preset);
      for(const [k, v] of Object.entries(setup.inputs)) document.getElementById(k).value = v;
      if(setup.loads) LOADS = setup.loads.map(l => ({...l}));
      drawLoads(); render();
    }
    return exprs.map(e => { try{ return (0, eval)(`(H => ${e})`)(H); }catch(err){ return 'ERR ' + err.message; } });
  }, [C.setup, C.items.map(i => i.tool), hJSON]);
  return {H, rows: C.items.map((it, i) => {
    const h = it.hand(H), t = tool[i], rel = typeof t === 'number' ? Math.abs(t - h)/Math.max(Math.abs(h), 1e-12) : NaN;
    return {k: it.k, label: it.label, unit: it.unit, s: it.s || 1, hand: h, tool: t, rel, tol: it.tol, ok: rel <= it.tol};
  })};
}

module.exports = {CASES, runCase, strain, col60, rectBars, uniDC, BAR};

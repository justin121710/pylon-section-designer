/* =========================================================================
   rc-core.js — RC 斷面設計工具之純計算函式（U9）
   不讀寫 DOM、不依賴全域狀態，網頁（<script src>，宣告為全域）與 Node（require）共用，
   以 tests/core.test.js 直接做單元測試。單位：kgf、cm。
   本檔為計算核心之一部分，修改後須同步重建教學 PDF（見 docs/tutorial/inputs.txt）。
   ========================================================================= */
'use strict';
const RC_CORE_VERSION = '2026.10-r4';

/* ---------- 基本力學 ---------- */
const beta1 = fc => Math.max(0.65, Math.min(0.85, 0.85 - 0.05*(fc-280)/70));

const Ecof  = fc => 12000*Math.sqrt(fc);      // 常重混凝土 E_c = 12,000√f'c（土木401-112 §19.2.2.1(b)）

const EcMander = fc => 15000*Math.sqrt(fc);  // Mander 模型之初始切線模數（≈ 4,700√f'c MPa），屬該模型參數、非規範分析用 E_c

/* ---------- P-M 互制（單軸） ---------- */
/* 土木401-112 表 21.2.2：ε_t ≤ ε_ty 壓力控制；ε_ty < ε_t < ε_ty + 0.003 過渡區線性內插；ε_t ≥ ε_ty + 0.003 拉力控制 */
const etTC = m => m.ety + 0.003;                 // 拉力控制應變下限

function phiOf(eps_t, m){
  if(eps_t<=m.ety) return m.phic;
  if(eps_t>=etTC(m)) return m.phit;
  return m.phic + (m.phit-m.phic)*(eps_t-m.ety)/0.003;
}

/* ---------- 剪力 / 扭矩 / 圍束筋 ---------- */
/* 塑鉸區 V_c 歸零門檻：橋梁 N_u < 0.1f'c·A_g；建築柱 N_u < A_g·f'c/20（土木401 §18.4.6.2.1，
   另一條件「地震引致剪力 ≥ 設計剪力之半」在耐震設計柱端視為成立，保守取之） */
const vcZeroRatio = code => code==='bldg' ? 0.05 : 0.10;

/* 耐震箍筋間距之主筋直徑倍數（土木401 §18.3.4.4、§18.4.5.3、§18.4.5.5）：f_y 4200 → 6、5000 → 5.5、5600 → 5 */
const seisKdb = fy => fy <= 4200+1e-6 ? 6 : fy <= 5000+1e-6 ? 5.5 : 5;

const capPhi = m => m.phiAx ?? (m.phiMode==='br532' ? m.phiBrMin : m.phic);   // φP_n,max 之 φ：基樁 0.55；橋梁 §5.3.2 為 0.70／0.75；其餘 φ_c

/* 公路橋梁耐震設計規範 §5.3.2：設計軸力之壓應力 ≥ min(0.1f'c, P_b/A_g) 時 φ = φ_min（橫箍 0.70、螺箍 0.75），
   降至 0 時線性增至 0.90。以設計軸力 φP_n 為變數：φ = 0.9 − (0.9 − φ_min)·φP_n/P_thr → φ = 0.9/(1 + (0.9 − φ_min)P_n/P_thr)；軸拉取 0.90 */
function phiAxialBr(Pn, Pthr, m){
  if(Pn <= 0 || !(Pthr > 0)) return 0.9;
  return Math.max(m.phiBrMin, 0.9/(1 + (0.9 - m.phiBrMin)*Pn/Pthr));
}

/* 平衡軸力 P_b（最外拉力筋 ε_t = ε_ty）：pts 依 c 遞增（ε_t 遞減），於跨越處內插 */
function balancedPn(pts, m){
  for(let i=0;i<pts.length-1;i++){
    const a = pts[i], b = pts[i+1];
    if(a.eps_t >= m.ety && b.eps_t < m.ety){ const t = (a.eps_t - m.ety)/((a.eps_t - b.eps_t) || 1); return (a.Pn ?? a.P) + t*((b.Pn ?? b.P) - (a.Pn ?? a.P)); }
  }
  return null;
}

const phiThrBr = (pts, m) => { const Pb = balancedPn(pts, m); return Math.min(0.1*m.fc*m.Ag, Pb > 0 ? Pb : Infinity); };

function sha256(str){
  const K = [0x428a2f98,0x71374491,0xb5c0fbcf,0xe9b5dba5,0x3956c25b,0x59f111f1,0x923f82a4,0xab1c5ed5,0xd807aa98,0x12835b01,0x243185be,0x550c7dc3,0x72be5d74,0x80deb1fe,0x9bdc06a7,0xc19bf174,
    0xe49b69c1,0xefbe4786,0x0fc19dc6,0x240ca1cc,0x2de92c6f,0x4a7484aa,0x5cb0a9dc,0x76f988da,0x983e5152,0xa831c66d,0xb00327c8,0xbf597fc7,0xc6e00bf3,0xd5a79147,0x06ca6351,0x14292967,
    0x27b70a85,0x2e1b2138,0x4d2c6dfc,0x53380d13,0x650a7354,0x766a0abb,0x81c2c92e,0x92722c85,0xa2bfe8a1,0xa81a664b,0xc24b8b70,0xc76c51a3,0xd192e819,0xd6990624,0xf40e3585,0x106aa070,
    0x19a4c116,0x1e376c08,0x2748774c,0x34b0bcb5,0x391c0cb3,0x4ed8aa4a,0x5b9cca4f,0x682e6ff3,0x748f82ee,0x78a5636f,0x84c87814,0x8cc70208,0x90befffa,0xa4506ceb,0xbef9a3f7,0xc67178f2];
  const b = new TextEncoder().encode(str), l = b.length, n = ((l + 9 + 63) >> 6) << 6, m = new Uint8Array(n);
  m.set(b); m[l] = 0x80; const dv = new DataView(m.buffer); dv.setUint32(n-4, (l*8)>>>0); dv.setUint32(n-8, Math.floor(l/0x20000000));
  const H = [0x6a09e667,0xbb67ae85,0x3c6ef372,0xa54ff53a,0x510e527f,0x9b05688c,0x1f83d9ab,0x5be0cd19], W = new Uint32Array(64), R = (x,k) => (x>>>k)|(x<<(32-k));
  for(let o=0;o<n;o+=64){
    for(let i=0;i<16;i++) W[i] = dv.getUint32(o+4*i);
    for(let i=16;i<64;i++){ const s0 = R(W[i-15],7)^R(W[i-15],18)^(W[i-15]>>>3), s1 = R(W[i-2],17)^R(W[i-2],19)^(W[i-2]>>>10); W[i] = (W[i-16]+s0+W[i-7]+s1)>>>0; }
    let [a,bb,c,d,e,f,g,h] = H;
    for(let i=0;i<64;i++){ const t1 = (h + (R(e,6)^R(e,11)^R(e,25)) + ((e&f)^(~e&g)) + K[i] + W[i])>>>0, t2 = ((R(a,2)^R(a,13)^R(a,22)) + ((a&bb)^(a&c)^(bb&c)))>>>0;
      h=g; g=f; f=e; e=(d+t1)>>>0; d=c; c=bb; bb=a; a=(t1+t2)>>>0; }
    H[0]=(H[0]+a)>>>0; H[1]=(H[1]+bb)>>>0; H[2]=(H[2]+c)>>>0; H[3]=(H[3]+d)>>>0; H[4]=(H[4]+e)>>>0; H[5]=(H[5]+f)>>>0; H[6]=(H[6]+g)>>>0; H[7]=(H[7]+h)>>>0;
  }
  return H.map(x => x.toString(16).padStart(8,'0')).join('');
}

/* ---------- 載重組合之頂／底配對（M6） ----------
   同名組合分列「頂」「底」兩列時視為同一構材之兩端內力：M_1/M_2 = −M_a·M_b／max(M_a², M_b²)
   （構材內力符號慣例，例如 ETABS／SAP2000 之 M2、M3 測站值：兩端同號＝單曲率 → 取負；異號＝雙曲率 → 取正；
   絕對值為小端／大端）。兩端皆為零或未配對者用參數表之 M_1/M_2。opt 層級各軸取最不利（最小）值供可忽略門檻使用。 */
function loadPairs(loads, m12Global){
  const key = L => String(L.name ?? '').trim();
  const ratio = (a, b) => { const mx = Math.max(a*a, b*b); return mx < 1e-18 ? null : Math.max(-1, Math.min(1, -a*b/mx)); };
  const rows = (loads || []).map(L => {
    const other = L.pos==='top' ? 'bot' : L.pos==='bot' ? 'top' : null;
    const mates = other ? loads.filter(Q => key(Q)===key(L) && Q.pos===other) : [];
    const mate = mates.length===1 ? mates[0] : null;
    const mx = mate ? ratio(+L.Mux||0, +mate.Mux||0) : null, my = mate ? ratio(+L.Muy||0, +mate.Muy||0) : null;
    return {paired: !!mate, m12x: mx ?? m12Global, m12y: my ?? m12Global};
  });
  const act = rows.length ? rows : [{m12x:m12Global, m12y:m12Global}];
  return {rows, any: rows.some(r => r.paired), m12x: Math.min(...act.map(r => r.m12x)), m12y: Math.min(...act.map(r => r.m12y))};
}

const POS_TXT = {top:'頂', bot:'底'};

/* ---------- 長細效應判定 ----------
   建築（土木401-112 §6.2.5）：
     §6.2.5.1 無側向位移支撐 kl_u/r ≤ 22；有側向位移支撐 kl_u/r ≤ 34 + 12(M_1/M_2) 且 ≤ 40
              （單曲率 M_1/M_2 為負、雙曲率為正）時得忽略長細效應；第一次近似分析時 k 可取 1.0（解說 R6.2.5）。
     §6.2.5.2 r = 0.30 × 考量方向之尺度（矩形）、0.25D（圓形），或全斷面 √(I_g/A_g)（箱型採此式）。
     §6.2.5.3 不可忽略時須依 §6.6.4／6.7／6.8 計二階效應，且二階 M_u ≤ 1.4 倍一階 M_u。
   橋梁（公路橋梁設計規範 §7.3.5.2）：
     (4) 有支撐防止側移 kl_u/r < 34 − 12(M_1b/M_2b)（M_1b/M_2b 單曲度為正，與上式正負相反；無 40 上限）；
         無支撐 kl_u/r < 22 時可不計；所有 kl_u/r > 100 者應依 §7.3.5.1 分析。(3) 有支撐構材 k 除經分析外應為 1.0。
   輸入之 M_1/M_2 一律採土木401-112 慣例（單曲率為負）；橋梁門檻 34 − 12(M_1b/M_2b) = 34 + 12(M_1/M_2)。
   X 向＝繞 X 軸彎曲（深度沿 H），Y 向＝繞 Y 軸（深度沿 B）；兩向共用同一 k 與 l_u。只判定，不計算彎矩放大。 */
/* ---------- 長細效應（建築 土木401-112 §6.2.5、§6.6.4；橋梁 公路橋梁設計規範 §7.3.5） ----------
   兩向分別判定：「繞 X 軸」（M_ux，迴轉半徑沿 H）用 k_x、l_u,x；「繞 Y 軸」（M_uy，沿 B）用 k_y、l_u,y。
   不可忽略時依「二階效應處理方式」：
     mag    — 輸入為一階分析彎矩，本工具以彎矩放大法放大後檢核（建築 §6.6.4.5／§6.6.4.6；橋梁 §7.3.5.2(5)）；
     second — 輸入已含二階效應（P-Δ 與構材 P-δ）之分析結果，不再放大（判定列與計算書註明）。 */
function colSlenderCore(s, m, o){
  const bridge = o.code==='bridge';
  let rX, rY, how, IgX, IgY;
  if(s.type==='circle'){ rX = rY = 0.25*s.D; how = 'circ'; IgX = IgY = Math.PI*s.D**4/64; }
  else if(s.type==='box'){
    const bi = Math.max(0, s.B-2*s.tw), hi = Math.max(0, s.H-2*s.tw), A = s.B*s.H - bi*hi;
    IgX = (s.B*s.H**3 - bi*hi**3)/12; IgY = (s.H*s.B**3 - hi*bi**3)/12;
    rX = Math.sqrt(IgX/A); rY = Math.sqrt(IgY/A); how = 'box';
  } else { rX = 0.3*s.H; rY = 0.3*s.B; how = 'rect'; IgX = s.B*s.H**3/12; IgY = s.H*s.B**3/12; }
  const kX = o.slK, kY = o.slKy, luX = o.lu, luY = o.slLuY > 0 ? o.slLuY : o.lu;
  const lamX = kX*luX/rX, lamY = kY*luY/rY, lam = Math.max(lamX, lamY);
  // M_1/M_2：載重表有頂／底配對時由兩端彎矩求得（各軸取最不利＝最小值），否則用輸入值
  const m12x = o.slM12x ?? o.slM12, m12y = o.slM12y ?? o.slM12;
  const limOf = m12 => o.slSway ? 22 : bridge ? 34 + 12*m12 : Math.min(34 + 12*m12, 40);
  const limX = limOf(m12x), limY = limOf(m12y), lim = Math.min(limX, limY), m12 = Math.min(m12x, m12y), limB = 34 + 12*m12;
  const okAx = (l, L) => bridge ? l < L : l <= L + 1e-9;
  const needX = !okAx(lamX, limX), needY = !okAx(lamY, limY);
  return {rX, rY, how, IgX, IgY, kX, kY, luX, luY, lamX, lamY, lam, lim, limX, limY, limB, needX, needY, m12x, m12y,
          pairs: !!(o.slPairs && o.slPairs.any), m12in: o.slM12,
          sway:o.slSway, k:kX, lu:luX, m12, ok: !needX && !needY, bridge,
          treat: o.slTreat, beta: o.slBeta, Q: o.slQ, Ec: m.Ec, phic: m.phic,
          hX: s.type==='circle' ? s.D : s.H, hY: s.type==='circle' ? s.D : s.B,
          over100: bridge && lam > 100, kLow: !o.slSway && Math.min(kX, kY) < 1 - 1e-9,
          ref: bridge ? '公路橋梁設計規範 §7.3.5.2' : '土木401-112 §6.2.5.1'};
}

/* 對一組載重套用彎矩放大（M_ux、M_uy 分別計算；最小偏心距不需兩軸同時施加，兩向彎矩皆為零時只施加於 X 向） */
function slenderApply(SL, Lk, pr){
  const both0 = Lk.Mux <= 1e-9 && Lk.Muy <= 1e-9;
  // C_m 之 M_1/M_2：頂／底配對之組合由兩端彎矩求得，其餘用輸入值
  const dx = slenderDelta(SL, 'x', Lk.Pu, Lk.Mux, pr ? pr.m12x : SL.m12in), dy = slenderDelta(SL, 'y', Lk.Pu, Lk.Muy, pr ? pr.m12y : SL.m12in);
  if(Lk.Muy <= 1e-9 && !both0){ dy.Mc = 0; dy.minGov = false; }
  if(Lk.Mux <= 1e-9 && !both0){ dx.Mc = 0; dx.minGov = false; }
  if(both0) dy.Mc = 0;
  Lk.Mux1 = Lk.Mux; Lk.Muy1 = Lk.Muy; Lk.slx = dx; Lk.sly = dy;
  Lk.Mux = dx.Mc; Lk.Muy = dy.Mc;
}

/* 單一載重組合、單一軸之彎矩放大。ax 'x'：M_ux（I_gx、k_x、l_u,x、h = H）；'y'：M_uy。回傳設計彎矩 Mc 與各係數（kgf、cm） */
function slenderDelta(SL, ax, Pu, Mu, m12 = SL.m12in ?? SL.m12){
  const need = ax==='x' ? SL.needX : SL.needY;
  const out = {need, d:1, Mc:Mu, Mu};
  if(!need || SL.treat!=='mag' || Pu <= 0) return out;
  const Ig = ax==='x' ? SL.IgX : SL.IgY, k = ax==='x' ? SL.kX : SL.kY, lu = ax==='x' ? SL.luX : SL.luY, h = ax==='x' ? SL.hX : SL.hY;
  const M2min = Pu*(1.5 + 0.03*h), M2 = Math.max(Mu, M2min);
  const amp = (Cm, Pc, K) => { const r = Pu/(K*Pc); return r >= 1 ? Infinity : Math.max(1, Cm/(1-r)); };
  let d, parts;
  if(SL.bridge){
    // §7.3.5.2(5)：EI = (E_cI_g/2.5)/(1+β_d)；δ_b = C_m/(1 − P_u/φP_c)，δ_s = 1/(1 − ΣP_u/φΣP_c)（以本柱 P_u/P_c 代替）
    const EI = SL.Ec*Ig/2.5/(1+SL.beta), phi = SL.phic;
    const Cm = Math.max(0.4, 0.6 + 0.4*(-m12));               // M_1b/M_2b 單曲度為正
    const PcB = Math.PI**2*EI/((SL.sway ? 1 : k)*lu)**2, db = amp(Cm, PcB, phi);
    if(SL.sway){ const PcS = Math.PI**2*EI/(k*lu)**2, ds = amp(1, PcS, phi); d = Math.max(db, ds); parts = {EI, Cm, PcB, db, PcS, ds, phi}; }
    else { d = db; parts = {EI, Cm, PcB, db, phi}; }
  }else{
    // §6.6.4.4.4(a) (EI)eff = 0.4E_cI_g/(1+β_dns)；§6.6.4.5 δ = C_m/(1 − P_u/0.75P_c) ≥ 1，C_m = 0.6 − 0.4M_1/M_2
    const Cm = 0.6 - 0.4*m12, EIns = 0.4*SL.Ec*Ig/(1+SL.beta);
    const PcN = Math.PI**2*EIns/((SL.sway ? 1 : k)*lu)**2, dn = amp(Cm, PcN, 0.75);
    if(SL.sway){
      // §6.6.4.6.2：(a) δ_s = 1/(1 − Q)（δ_s > 1.5 時不得使用）或 (b) 1/(1 − ΣP_u/0.75ΣP_c)，(EI)eff 以 β_ds = 0；本工具以本柱 P_u/P_c 代替 Σ
      const EIs = 0.4*SL.Ec*Ig, PcS = Math.PI**2*EIs/(k*lu)**2;
      const dsQ = SL.Q > 0 ? (SL.Q < 1 ? 1/(1-SL.Q) : Infinity) : null, dsB = amp(1, PcS, 0.75);
      const ds = dsQ !== null && dsQ <= 1.5 ? Math.max(1, dsQ) : dsB;
      // §6.6.4.6.4：沿柱長之二階效應依 §6.6.4.5（無側移，k = 1.0）
      d = ds*dn; parts = {Cm, EIns, PcN, dn, EIs, PcS, ds, dsQ, dsB, qOver: dsQ !== null && dsQ > 1.5};
    }else{ d = dn; parts = {Cm, EIns, PcN, dn}; }
  }
  return {...out, d, Mc: d*M2, M2, M2min, minGov: M2min > Mu, ...parts};
}

/* T 梁有效翼緣寬（土木401 / ACI 6.3.2） */
/* T 梁有效翼板寬（土木401-112 表 6.3.2.1，h 為板厚 h_f、s_w 為相鄰腹板淨間距、ℓ_n 為淨跨）：
   腹板兩側 b_e = b_w + 2·min(8h_f, s_w/2, ℓ_n/8)；單側 b_e = b_w + min(6h_f, s_w/2, ℓ_n/12)。
   單獨 T 梁（§6.3.2.2）：h_f ≥ 0.5b_w 且 b_e ≤ 4b_w，b_e = min(b_f, 4b_w)。
   舊版沿用 L/4、b_w + 16h_f 與全跨 L/12（邊梁偏大、不保守）。cand.t 為 LaTeX，g 為畫面文字 */
function effFlange(S){
  if(S.type!=='T') return {be:S.bw, gov:'—', cand:[]};
  if(S.flPos==='iso'){
    const c = [{v:S.bf, t:'b_f', g:'翼板全寬 b_f'}, {v:4*S.bw, t:'4b_w', g:'4b_w（§6.3.2.2）'}];
    const m = c.reduce((a,b)=> b.v<a.v?b:a, c[0]);
    return {be:m.v, gov:m.g, cand:c, iso:true, hfOk: S.hf >= 0.5*S.bw - 1e-9};
  }
  const ln = S.ln, edge = S.flPos==='edge', sw = edge ? S.sClear : Math.max(0, S.sSpacing - S.bw), k = edge ? 1 : 2;
  const ov = edge
    ? [{v:6*S.hf, t:'6h_f', g:'6h_f'}, {v:sw/2, t:'s_w/2', g:'s_w/2（淨距之半）'}, {v:ln/12, t:'\\ell_n/12', g:'ℓ_n/12'}]
    : [{v:8*S.hf, t:'8h_f', g:'8h_f'}, {v:sw/2, t:'s_w/2', g:'s_w/2（腹板淨距之半）'}, {v:ln/8, t:'\\ell_n/8', g:'ℓ_n/8'}];
  const c = ov.map(o => ({v: S.bw + k*o.v, t: `b_w+${k===2?'2\\times ':''}${o.t}`, g: `b_w + ${k===2?'2×':''}${o.g}`, ov:o.v}));
  const m = c.reduce((a,b)=> b.v<a.v?b:a, c[0]);
  return {be:m.v, gov:m.g, cand:c, sw};
}

/* ---------- 裂縫控制（土木401-112 表 24.3.2） ---------- */
function beamCrack(S, R, m, ctx){
  const fs = (2/3)*m.fy;         // ACI 24.3.2 允許以 (2/3)f_y 代替由使用載重算得之鋼筋應力
  // c_c：最接近受拉面之拉力鋼筋至受拉面之淨保護層（§24.3.2）；梁之保護層量至箍筋，須加 d_t；版無箍筋
  const cc = S.cover + (S.slab ? 0 : R.dt);
  const s1 = 38*(2800/fs) - 2.5*cc;
  const s2 = 30*(2800/fs);
  const sLim = Math.min(s1, s2);
  // 受拉側鋼筋中心距：正彎矩取底筋第 1 排、負彎矩取頂筋
  const pos = S.sign!=='neg';
  const nRow = pos ? R.r1 : Math.round(S.nTop), dbT = pos ? R.dbB : R.dbT;
  const sAct = S.slab ? (pos ? S.spBot : (S.spTop || S.spBot))
             : nRow>1 ? (S.bw - 2*(S.cover+R.dt) - nRow*dbT)/(nRow-1) + dbT : S.bw;
  return {fs, cc, s1, s2, sLim, sAct, nRow, ok:sAct<=sLim};
}

if(typeof module !== 'undefined' && module.exports)
  module.exports = {RC_CORE_VERSION, beta1, Ecof, EcMander, etTC, phiOf, capPhi, vcZeroRatio, seisKdb, phiAxialBr, balancedPn, phiThrBr,
    sha256, loadPairs, POS_TXT, colSlenderCore, slenderApply, slenderDelta, effFlange, beamCrack};

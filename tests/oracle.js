/* 獨立求解器（測試預言機）：不引用 index.html 之任何計算函式，僅使用斷面幾何、鋼筋座標與材料參數。
   作法與工具不同以利交叉驗證：混凝土以纖維離散（非多邊形裁切）、中性軸方向角逐點掃描、
   等軸力輪廓以 c 二分求得（非網格內插）。
   假設（土木401-112 §22.2）：平面保持平面、ε_cu、等效應力塊 0.85f'c × β₁c（距最外受壓纖維量測）、
   鋼筋彈塑性、受壓區內鋼筋扣除 0.85f'c；φ 依最外拉力筋 ε_t（表 21.2.2）；φP_n,max = φ_c·α·P_o 截斷。 */
'use strict';

/* phiAx（選用）：公路橋梁耐震設計規範 §5.3.2 之 φ 依軸壓（{min, thr}：φ = 0.9 − (0.9 − min)·φP_n/thr，下限 min，軸拉 0.9）；
   不給時 φ 依 ε_t（土木401-112 表 21.2.2） */
function makeSection({B, H, tw = 0, box = false, bars, fc, fy, Es = 2.04e6, ecu = 0.003, phic = 0.65, phit = 0.9, pmaxf = 0.8, nf = 80, nx = nf, ny = nf, phiAx = null}){
  const b1 = Math.max(0.65, Math.min(0.85, 0.85 - 0.05*(fc - 280)/70)), ety = fy/Es;
  const fib = [], dA = (B/nx)*(H/ny);                      // nx × ny 纖維（薄受壓區可加密 ny）
  for(let i=0;i<nx;i++) for(let j=0;j<ny;j++){
    const x = -B/2 + (i+0.5)*B/nx, y = -H/2 + (j+0.5)*H/ny;
    if(box && Math.abs(x) < B/2 - tw && Math.abs(y) < H/2 - tw) continue;
    fib.push([x, y]);
  }
  const Ag = fib.length*dA, Ast = bars.reduce((a, b) => a + b.A, 0);
  const Po = 0.85*fc*(Ag - Ast) + fy*Ast, cap = (phiAx ? phiAx.min : phic)*pmaxf*Po, Pt = -(phiAx ? 0.9 : phit)*fy*Ast;
  const corners = [[B/2,H/2],[-B/2,H/2],[B/2,-H/2],[-B/2,-H/2]];
  const phiOf = et => et <= ety ? phic : et >= ety + 0.003 ? phit : phic + (phit - phic)*(et - ety)/0.003;
  /* φ 依軸壓（§5.3.2）：解 φ = 0.9 − (0.9 − min)·φP/thr 之 φ（以二分求，不用封閉式） */
  const phiAxOf = Pn => { if(Pn <= 0) return 0.9; let lo = phiAx.min, hi = 0.9;
    for(let k=0;k<60;k++){ const mid = (lo+hi)/2; if(mid > 0.9 - (0.9 - phiAx.min)*Math.min(1, mid*Pn/phiAx.thr)) hi = mid; else lo = mid; }
    return (lo+hi)/2; };
  /* 方向角 th（受壓側法向量 (cos th, sin th)）、中性軸深度 c 之設計值；nominal = true 回傳標稱值 */
  function point(th, c, nominal = false){
    const nx = Math.cos(th), ny = Math.sin(th), top = Math.max(...corners.map(q => q[0]*nx + q[1]*ny));
    const a = b1*c;
    let P = 0, Mx = 0, My = 0, dmax = -Infinity;
    for(const [x, y] of fib) if(top - (x*nx + y*ny) <= a){ const F = 0.85*fc*dA; P += F; Mx += F*y; My += F*x; }
    for(const b of bars){
      const d = top - (b.x*nx + b.y*ny);
      let fs = Math.max(-fy, Math.min(fy, Es*ecu*(c - d)/c));
      if(d <= a) fs -= 0.85*fc;
      P += fs*b.A; Mx += fs*b.A*b.y; My += fs*b.A*b.x;
      dmax = Math.max(dmax, d);
    }
    const phi = nominal ? 1 : phiAx ? phiAxOf(P) : phiOf(ecu*(dmax - c)/c);
    return {P: phi*P, Mx: Math.abs(phi*Mx), My: Math.abs(phi*My), dmax};
  }
  /* 軸力 P 下、彎矩方向 (ux, uy) 之設計彎矩容量（0 表示 P 超出範圍） */
  function Mcap(P, ux, uy, nth = 46){
    if(P > cap || P < Pt) return 0;
    const D = Math.hypot(B, H), pts = [];
    for(let k=0;k<nth;k++){
      const th = Math.PI/2*k/(nth-1);
      let lo = 1e-4*D, hi = 20*D;
      if(point(th, lo).P > P) return 0;
      for(let it=0; it<40; it++){ const mid = Math.sqrt(lo*hi); if(point(th, mid).P < P) lo = mid; else hi = mid; }
      const q = point(th, Math.sqrt(lo*hi)); pts.push([q.Mx, q.My]);
    }
    let best = 0;
    for(let i=0;i<pts.length-1;i++){
      const [x1,y1] = pts[i], [x2,y2] = pts[i+1], dx = x2-x1, dy = y2-y1, den = ux*dy - uy*dx;
      if(Math.abs(den) < 1e-12) continue;
      const r = (x1*dy - y1*dx)/den, w = Math.abs(dx) > Math.abs(dy) ? (r*ux - x1)/dx : (r*uy - y1)/dy;
      if(r > 0 && w >= -1e-9 && w <= 1 + 1e-9) best = Math.max(best, r);
    }
    return best;
  }
  /* 定偏心射線 D/C */
  function dc(Pu, Mux, Muy){
    const Mr = Math.hypot(Mux, Muy), ux = Mux/Mr, uy = Muy/Mr;
    const inside = t => Mcap(t*Pu, ux, uy) >= t*Mr;
    let lo = 0, hi = 1;
    while(inside(hi) && hi < 1e4){ lo = hi; hi *= 2; }
    for(let it=0; it<30; it++){ const mid = (lo+hi)/2; if(inside(mid)) lo = mid; else hi = mid; }
    return 1/lo;
  }
  /* 標稱彎矩 M_n（繞 X：th = 90°；繞 Y：th = 0°）於標稱軸力 P；平衡軸力 P_b（最外拉力筋 ε_t = ε_ty） */
  function MnAt(P, th){
    const D = Math.hypot(B, H); let lo = 1e-4*D, hi = 20*D;
    for(let it=0; it<60; it++){ const mid = Math.sqrt(lo*hi); if(point(th, mid, true).P < P) lo = mid; else hi = mid; }
    const q = point(th, Math.sqrt(lo*hi), true); return Math.hypot(q.Mx, q.My);
  }
  function Pb(th){
    const D = Math.hypot(B, H); let lo = 1e-3*D, hi = 5*D;   // ε_t = ε_cu(d_max − c)/c 隨 c 遞減
    for(let it=0; it<60; it++){ const c = (lo+hi)/2, q = point(th, c, true); if(ecu*(q.dmax - c)/c > ety) lo = c; else hi = c; }
    return point(th, (lo+hi)/2, true).P;
  }
  /* 單軸（繞 X 軸）定偏心射線 D/C：沿 c 細掃 (φP, φM_x)，找射線 P/M = P_u/M_u 之所有交點，取最內側（容量最小）者 */
  function dcUniX(Pu, Mu, n = 3000){
    const th = Math.PI/2, D = H; let prev = null, tMin = Infinity;
    for(let k=0;k<=n;k++){
      const c = 1e-3*D*Math.pow(5000, k/n), q = point(th, c), g = q.P*Mu - Pu*q.Mx;
      if(prev && (prev.g <= 0) !== (g <= 0)){
        const w = prev.g/(prev.g - g), Mx = prev.Mx + w*(q.Mx - prev.Mx), P = prev.P + w*(q.P - prev.P);
        if(P <= cap) tMin = Math.min(tMin, Mx/Mu);
      }
      prev = {g, Mx: q.Mx, P: q.P};
    }
    if(Pu > 0 && tMin*Pu > cap) tMin = cap/Pu;                 // φP_n,max 截斷
    return 1/tMin;
  }
  return {point, Mcap, dc, dcUniX, cap, Pt, Po, Ag, b1, MnAt, Pb};
}

/* 可重現之亂數（mulberry32） */
function rng(seed){
  return () => { seed |= 0; seed = seed + 0x6D2B79F5 | 0; let t = Math.imul(seed ^ seed >>> 15, 1 | seed);
    t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0)/4294967296; };
}

module.exports = {makeSection, rng};

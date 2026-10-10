/* 獨立求解器（測試預言機）：不引用 index.html 之任何計算函式，僅使用斷面幾何、鋼筋座標與材料參數。
   作法與工具不同以利交叉驗證：混凝土以纖維離散（非多邊形裁切）、中性軸方向角逐點掃描、
   等軸力輪廓以 c 二分求得（非網格內插）。
   假設（土木401-112 §22.2）：平面保持平面、ε_cu、等效應力塊 0.85f'c × β₁c（距最外受壓纖維量測）、
   鋼筋彈塑性、受壓區內鋼筋扣除 0.85f'c；φ 依最外拉力筋 ε_t（表 21.2.2）；φP_n,max = φ_c·α·P_o 截斷。 */
'use strict';

function makeSection({B, H, tw = 0, box = false, bars, fc, fy, Es = 2.04e6, ecu = 0.003, phic = 0.65, phit = 0.9, pmaxf = 0.8, nf = 80}){
  const b1 = Math.max(0.65, Math.min(0.85, 0.85 - 0.05*(fc - 280)/70)), ety = fy/Es;
  const fib = [], dA = (B/nf)*(H/nf);
  for(let i=0;i<nf;i++) for(let j=0;j<nf;j++){
    const x = -B/2 + (i+0.5)*B/nf, y = -H/2 + (j+0.5)*H/nf;
    if(box && Math.abs(x) < B/2 - tw && Math.abs(y) < H/2 - tw) continue;
    fib.push([x, y]);
  }
  const Ag = fib.length*dA, Ast = bars.reduce((a, b) => a + b.A, 0);
  const Po = 0.85*fc*(Ag - Ast) + fy*Ast, cap = phic*pmaxf*Po, Pt = -phit*fy*Ast;
  const corners = [[B/2,H/2],[-B/2,H/2],[B/2,-H/2],[-B/2,-H/2]];
  const phiOf = et => et <= ety ? phic : et >= ety + 0.003 ? phit : phic + (phit - phic)*(et - ety)/0.003;
  /* 方向角 th（受壓側法向量 (cos th, sin th)）、中性軸深度 c 之設計值 */
  function point(th, c){
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
    const phi = phiOf(ecu*(dmax - c)/c);
    return {P: phi*P, Mx: Math.abs(phi*Mx), My: Math.abs(phi*My)};
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
  return {point, Mcap, dc, cap, Pt, Po, Ag, b1};
}

/* 可重現之亂數（mulberry32） */
function rng(seed){
  return () => { seed |= 0; seed = seed + 0x6D2B79F5 | 0; let t = Math.imul(seed ^ seed >>> 15, 1 | seed);
    t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0)/4294967296; };
}

module.exports = {makeSection, rng};

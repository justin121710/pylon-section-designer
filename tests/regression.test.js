/* 數值回歸測試（node --test）
   每一項的期望值都來自「不呼叫工具計算函式」的獨立計算或規範原式，用以防止已修正的錯誤再出現：
     C1 柱容量設計剪力之剪力方向與彎矩平面配對
     C2 雙軸 Bresler 倒數式之 P_o（未截斷 φ_c·P_o）
     C3 剪扭合併時扭矩只由外圍閉合肢承擔（土木401-112 §9.5.4.3 解說）
     C4 梁剪力／扭矩逐組合檢核（扭矩不一定與最大剪力同組）
     H1 雙軸 D/C 採旋轉中性軸精確解（與獨立程式一致、0.1f'cA_g 前後連續、退化為單軸）
   另含單軸 P-M 逐點驗算（核心基準）與全部範本不得出現不合格。H2～H8 見 high.test.js，獨立求解器比對見 independent.test.js。 */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const {openApp} = require('./helpers');

let app, page;
test.before(async () => { app = await openApp(); page = app.page; });
test.after(async () => { if(app){ assert.deepEqual(app.errors, [], '頁面不得有 JavaScript 錯誤'); await app.close(); } });

const near = (act, exp, rel, msg) =>
  assert.ok(Math.abs(act - exp) <= rel*Math.abs(exp), `${msg}：實際 ${act}，期望 ${exp}（容許 ${rel*100}%）`);

/* 建築 RC 方柱 60×60 改為 B×H 與每面根數（其餘沿用範本） */
const rectColumn = (B, H, nB, nH) => page.evaluate(([B, H, nB, nH]) => {
  applyPreset('bldg60');
  const set = (id, v) => { document.getElementById(id).value = v; };
  set('B', B); set('H', H); set('nB', nB); set('nH', nH);
  render();
}, [B, H, nB, nH]);

/* ------------------------------------------------------------------ */
test('單軸 P-M：與獨立計算逐點一致（60×60、12-#8、繞 X 軸）', async () => {
  // 獨立計算：等效應力塊 + 離散鋼筋（受壓區內扣 0.85f'c），φ 依最外拉力筋 ε_t（表 21.2.2）
  const fc = 280, fy = 4200, Es = 2.04e6, ecu = 0.003, b1 = 0.85, B = 60, H = 60, Ab = 5.067;
  const dOut = 4 + 1.27 + 2.54/2, y0 = H/2 - dOut;
  const rows = [[y0, 4], [y0/3, 2], [-y0/3, 2], [-y0, 4]];
  const point = c => {
    const a = Math.min(b1*c, H), Cc = 0.85*fc*B*a;
    let P = Cc, M = Cc*(H/2 - a/2);
    for(const [y, k] of rows){
      const d = H/2 - y; let fs = Math.max(-fy, Math.min(fy, Es*ecu*(c - d)/c));
      if(d <= a) fs -= 0.85*fc;
      P += fs*Ab*k; M += fs*Ab*k*y;
    }
    const et = ecu*((H/2 + y0) - c)/c, ety = fy/Es;
    const phi = et <= ety ? 0.65 : et >= ety + 0.003 ? 0.9 : 0.65 + 0.25*(et - ety)/0.003;
    return {P, M, phi};
  };
  await page.evaluate(() => applyPreset('bldg60'));
  for(const c of [10, 20, 31.7, 45, 60]){
    const t = await page.evaluate(c => { const s = sectionAt(c, MODEL.ctxX); return {P:s.Pn, M:s.Mn, phi:s.phi}; }, c);
    const e = point(c);
    near(t.P, e.P, 1e-6, `c = ${c} 之 Pn`); near(t.M, e.M, 1e-6, `c = ${c} 之 Mn`); near(t.phi, e.phi, 1e-9, `c = ${c} 之 φ`);
  }
});

/* ------------------------------------------------------------------ */
test('C1：矩形柱容量設計剪力取同一受力平面之 M_pr（40×90 建築柱）', async () => {
  await rectColumn(40, 90, 3, 6);
  const r = await page.evaluate(() => {
    const M = MODEL, mx = c => Math.max(...M.res.map(q => MnAtP(c, q.Lk.Pu)));
    return {
      MprX: mx(M.cvXpr), MprY: mx(M.cvYpr), lu: M.opt.lu,
      shX: {dir:M.shX.dir, d:M.shX.d, Mn:M.shX.Mn, Ve:M.shX.Ve},
      shY: {dir:M.shY.dir, d:M.shY.d, Mn:M.shY.Mn, Ve:M.shY.Ve, VsReq:M.shY.VsReq, VsMax:M.shY.VsMax, fail:M.shY.sectionFail},
      shX2Mn: M.shX2.Mn, shY2Mn: M.shY2.Mn,
      verdictShearBad: [...document.querySelectorAll('#verdict .vd.bad')].some(b => b.textContent.includes('剪力'))
    };
  });
  // M_pr 繞 X 軸（深度沿 H = 90）為強軸，必大於繞 Y 軸
  assert.ok(r.MprX > 2*r.MprY, '40×90 柱繞 X 軸之 M_pr 應約為繞 Y 軸之 2 倍以上');
  // y 向剪力（d 沿 H）↔ 繞 X 軸彎矩；x 向剪力（d 沿 B）↔ 繞 Y 軸彎矩
  near(r.shY.Mn, r.MprX, 1e-9, 'Y 向剪力之容量設計彎矩');
  near(r.shX.Mn, r.MprY, 1e-9, 'X 向剪力之容量設計彎矩');
  near(r.shY2Mn, r.MprX, 1e-9, '加密區外 Y 向剪力之容量設計彎矩');
  near(r.shX2Mn, r.MprY, 1e-9, '加密區外 X 向剪力之容量設計彎矩');
  // V_e = 2M_pr/l_u（土木401-112 §18.4.6.1）
  near(r.shY.Ve, 2*r.MprX/r.lu, 1e-9, 'Y 向 V_e = 2M_pr/l_u');
  near(r.shY.Ve/1000, 104.2, 0.01, 'Y 向 V_e（tf）');
  // 此斷面強軸受剪 V_s 需求 > 2.12√f'c·b_w·d：必須判斷面不足
  assert.ok(r.shY.fail && r.shY.VsReq > r.shY.VsMax, 'Y 向應判定斷面剪力不足');
  assert.ok(r.verdictShearBad, '判定列之剪力應為不合格');
});

test('C1：橋梁矩形墩（200×300）V_e = φ_o·M_n/L_v 取同一平面之 M_n', async () => {
  await page.evaluate(() => applyPreset('pier200'));
  const r = await page.evaluate(() => {
    const M = MODEL, mx = c => Math.max(...M.res.map(q => MnAtP(c, q.Lk.Pu)));
    return {MnX: mx(M.cvX), MnY: mx(M.cvY), shX: M.shX.Mn, shY: M.shY.Mn, VeY: M.shY.Ve, phio: M.opt.phio, Lv: M.opt.Lv};
  });
  near(r.shY, r.MnX, 1e-9, 'Y 向剪力取繞 X 軸 M_n'); near(r.shX, r.MnY, 1e-9, 'X 向剪力取繞 Y 軸 M_n');
  near(r.VeY, r.phio*r.MnX/r.Lv, 1e-9, 'Y 向 V_e = φ_o·M_n/L_v');
  near(r.VeY/1000, 438, 0.01, 'Y 向 V_e（tf）');
});

/* ------------------------------------------------------------------ */
test('C2：Bresler 參考值使用未截斷之 φ_c·P_o', async () => {
  for(const k of ['bldg60', 'pond80', 'pier100', 'box300']){
    const rows = await page.evaluate(k => {
      applyPreset(k);
      return MODEL.res.filter(q => q.bi.ref && q.bi.ref.method.startsWith('Bresler')).map(q => ({
        name: q.L.name, Pu: q.Lk.Pu, dc: q.bi.ref.dc, Pnx: q.bi.ref.Pnx, Pny: q.bi.ref.Pny, Po: q.bi.ref.Po,
        phiPo: MODEL.m.phic*MODEL.cvX.Po, cap: MODEL.pX.cap}));
    }, k);
    assert.ok(rows.length, `${k} 應有 Bresler 參考值`);
    for(const q of rows){
      near(q.Po, q.phiPo, 1e-12, `${k}「${q.name}」之 φP_o`);
      assert.ok(q.Po > q.cap, 'φP_o 應大於截斷後之 φP_n,max');
      near(q.dc, q.Pu*(1/q.Pnx + 1/q.Pny - 1/q.phiPo), 1e-9, `${k}「${q.name}」D/C = P_u(1/P_nx + 1/P_ny − 1/φP_o)`);
    }
  }
});

test('H1：雙軸 D/C 採旋轉中性軸精確解，與獨立求解一致', async () => {
  /* 期望值：審查時以獨立程式（90×90 混凝土纖維、91 個中性軸方向角、離散鋼筋、φ 依 ε_t、沿定偏心射線二分）求得，
     未呼叫工具任何計算函式。工具以多邊形裁切（解析）＋ c 對數網格內插，容許 ±0.006。 */
  const EXACT = {
    bldg60: {'1.2D+1.6L': 0.504, '1.2D+L+E (X)': 0.718, '1.2D+L+E (Y)': 0.692, '0.9D+E': 0.577},
    pond80: {'1.2D+1.6L（頂版覆土）': 0.542, '1.2D+1.0L+E (X)': 0.553, '1.2D+1.0L+E (Y)': 0.541, '0.9D+E（滿水浮力）': 0.376},
    pier100: {'D+L': 0.295, '地震 X': 0.678, '地震 Y': 0.678, '最小軸力': 0.672},
    box300: {'D+L': 0.239, '地震 X': 0.652, '地震 Y': 0.618, '最小軸力': 0.762}
  };
  for(const [k, ex] of Object.entries(EXACT)){
    // 期望值依範本原始配筋（橋墩柱 6×6-#8、橋塔 19×25-#8）與未放大之輸入彎矩求得 → 聲明「輸入已含二階效應」以免放大；
    // 比對撓曲本身之 D/C（q.bi.dc，未扣除扭力縱筋）
    const r = await page.evaluate(k => { applyPreset(k); const set = (id, v) => { document.getElementById(id).value = v; };
      set('slTreat', 'second');
      if(k === 'pier100') set('barSize', '#8');
      if(k === 'box300'){ set('barSize', '#8'); set('nB', 19); set('nH', 25); }
      render(); return Object.fromEntries(MODEL.res.map(q => [q.L.name, {dc: q.bi.dc, method: q.bi.method}])); }, k);
    for(const [name, v] of Object.entries(ex)){
      assert.equal(r[name].method, '旋轉中性軸（精確解）', `${k}「${name}」之判定方法`);
      assert.ok(Math.abs(r[name].dc - v) <= 0.006, `${k}「${name}」D/C ${r[name].dc.toFixed(4)} 與獨立精確解 ${v} 不符`);
    }
  }
});

test('H1：雙軸 D/C 在 P_u = 0.1f′cA_g 前後連續，且退化為單軸時與 P-M 曲線一致', async () => {
  const r = await page.evaluate(() => {
    applyPreset('pier100');
    const M = MODEL, m = M.m, L = M.res[1].Lk, thr = 0.1*m.fc*m.Ag;
    const dc = P => biaxialDC(M.cvX, M.cvY, M.pX, M.pY, m, {Pu: P, Mux: L.Mux, Muy: L.Muy}, M.S3).dc;
    const P0 = 0.3*M.pX.cap, Mx = 0.5*phiMnAtP(M.cvX, m, P0);
    return {below: dc(0.999*thr), above: dc(1.001*thr),
            ex: biaxExactDC(M.S3, P0, Mx, 1e-7*Mx).dc, uni: radialDC(M.pX.poly, Mx, P0, M.pX.cap, m.phit*M.cvX.Pnt).dc};
  });
  assert.ok(Math.abs(r.below - r.above) < 0.003, `切換點前後 D/C ${r.below} → ${r.above} 不連續`);
  near(r.ex, r.uni, 0.005, '退化為繞 X 軸單軸');
});

/* ------------------------------------------------------------------ */
test('C3：柱剪扭合併時扭矩只由外圍閉合肢承擔（4 肢，T_u = 6 tf·m）', async () => {
  const r = await page.evaluate(() => {
    applyPreset('bldg60'); LOADS.forEach(L => { L.Tu = 6; }); drawLoads(); render();
    return [MODEL.shX, MODEL.shY].map(sh => ({nlegs: sh.nlegs, At1: sh.At1, AtS: sh.AtS, AvS: sh.AvS_shear, s: sh.s_str, torsion: sh.torsion,
      all: sh.all.map(c => ({AtS: c.AtS, AvS: c.AvS_shear, s: c.s_str}))})).concat([{sUse: MODEL.fin.sUse, s2: MODEL.fin2.sUse}]);
  });
  const fin = r.pop();
  for(const sh of r){
    assert.equal(sh.nlegs, 4); assert.ok(sh.torsion);
    // 外圍單肢：A_b/s ≥ A_t/s + (A_v/s)/n_legs
    for(const c of sh.all) near(c.s, sh.At1/(c.AtS + c.AvS/sh.nlegs), 1e-9, '各組合強度需求間距');
    near(sh.s, Math.min(...sh.all.map(c => c.s)), 1e-9, '控制組合為強度需求間距最小者');
    near(sh.s, 11.91, 0.005, '強度需求間距（cm）');
    // 舊算法（全部肢分攤扭矩）會得到較大間距
    assert.ok(sh.nlegs*sh.At1/(sh.AvS + 2*sh.AtS) > sh.s + 2, '外圍肢檢核應比全部肢平均分攤嚴格');
    assert.ok(fin.sUse <= sh.s, `採用間距 ${fin.sUse} 不得大於需求 ${sh.s.toFixed(2)}`);
    assert.ok(fin.s2 <= sh.s, `加密區外採用間距 ${fin.s2} 不得大於需求 ${sh.s.toFixed(2)}`);
  }
});

test('C3：2 肢時與 A_v/(A_v/s + 2A_t/s) 相同（圓柱）', async () => {
  const r = await page.evaluate(() => {
    applyPreset('circ80'); LOADS.forEach(L => { L.Tu = 8; }); drawLoads(); render();
    const sh = MODEL.shX; return {nlegs: sh.nlegs, Av: sh.Av, AvS: sh.AvS_shear, AtS: sh.AtS, s: sh.s_str, torsion: sh.torsion};
  });
  assert.equal(r.nlegs, 2); assert.ok(r.torsion);
  near(r.s, r.Av/(r.AvS + 2*r.AtS), 1e-9, '2 肢強度需求間距');
});

/* ------------------------------------------------------------------ */
const beamCase = loads => page.evaluate(loads => {
  document.getElementById('tabBeam').click();
  applyBPreset('bldgBeam'); BLOADS = loads; drawBLoads(); renderBeam();
  const B = BMODEL;
  return {ctrl: B.sh.ctrl.L.name, Tu: B.sh.Tu, tors: B.sh.tors, sStr: B.sh.sStr, sUse: B.fin.sUse, s2: B.fin2 ? B.fin2.sUse : null,
          nlegs: B.sh.nlegs, At1: B.sh.At1, AtS: B.sh.AtS, AvS: B.sh.AvS, all: B.sh.all.length};
}, loads);

test('C4：梁剪力／扭矩逐組合檢核，最大剪力組合以外之扭矩不得被忽略', async () => {
  const A = {name:'A：最大剪力', Mpos:20, Mneg:35, Vu:30, Tu:0}, B = {name:'B：伴隨扭矩', Mpos:20, Mneg:35, Vu:28, Tu:4};
  const onlyB = await beamCase([B]);
  const both = await beamCase([A, B]);
  const both2 = await beamCase([B, A]);
  assert.ok(onlyB.tors, 'B 組合須設計扭矩（T_u = 4 > 門檻）');
  assert.equal(both.all, 2);
  assert.equal(both.ctrl, B.name, '控制組合應為帶扭矩之 B');
  assert.ok(both.tors, 'A＋B 時仍須設計扭矩');
  assert.equal(both.sUse, onlyB.sUse, 'A＋B 之採用間距須與只有 B 相同');
  assert.equal(both.s2, onlyB.s2, 'A＋B 之加密區外採用間距須與只有 B 相同');
  assert.equal(both2.sUse, onlyB.sUse, '組合順序不得影響結果');
  assert.equal(onlyB.sUse, 10);
  await page.evaluate(() => document.getElementById('tabPylon').click());
});

test('C3：梁肢數 > 2 時扭矩只由外圍閉合肢承擔', async () => {
  const r = await page.evaluate(() => {
    document.getElementById('tabBeam').click();
    applyBPreset('bldgBeam'); document.getElementById('bNlegs').value = 4;
    BLOADS = [{name:'T', Mpos:20, Mneg:35, Vu:28, Tu:4}]; drawBLoads(); renderBeam();
    const sh = BMODEL.sh; return {nlegs: sh.nlegs, At1: sh.At1, AtS: sh.AtS, AvS: sh.AvS, s: sh.sStr};
  });
  assert.equal(r.nlegs, 4);
  near(r.s, r.At1/(r.AtS + r.AvS/4), 1e-9, '4 肢梁強度需求間距');
  await page.evaluate(() => document.getElementById('tabPylon').click());
});

/* ------------------------------------------------------------------ */
test('全部範本：不得出現不合格項目', async () => {
  const bad = await page.evaluate(() => {
    const out = [];
    for(const k of Object.keys(PRESETS)){ applyPreset(k);
      for(const a of colAlertItems(MODEL)) if(a.lv === 'bad') out.push(`柱 ${k}：${a.t.replace(/<[^>]+>/g, '').slice(0, 60)}`); }
    document.getElementById('tabBeam').click();
    for(const k of Object.keys(BPRESETS)){ applyBPreset(k);
      for(const a of beamAlertItems(BMODEL)) if(a.lv === 'bad') out.push(`梁 ${k}：${a.t.replace(/<[^>]+>/g, '').slice(0, 60)}`); }
    document.getElementById('tabPylon').click();
    return out;
  });
  assert.deepEqual(bad, []);
});

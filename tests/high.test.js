/* 回歸測試（H2～H8）：期望值以規範原式於測試端獨立計算（只讀取輸入欄位與鋼筋座標，不呼叫工具計算函式），
   或以 tests/oracle.js 之獨立求解器驗算：
     H2 長細效應彎矩放大（建築無側移／有側移 Q 法與 ΣP_c 法、橋梁、輸入已含二階、P_u ≥ 0.75P_c）
     H3/H4 無效輸入停止計算、禁止輸出（空白、箱型壁厚、非整數根數、梁翼緣厚）
     H5 扭力縱筋 A_ℓ、A_ℓ,min 與主筋扣除後之 D/C；T 梁 A_cp 依 §9.2.4.4 計入翼板
     H6 強柱弱梁、未涵蓋項目、橋梁塑鉸區續接位置
     H8 建議配筋之目標 D/C 上限與用鋼量（含搭接、彎鉤） */
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

/* 套用範本後改寫輸入欄位與載重（tf、tf·m），重算並回傳輸入值 */
const setup = (preset, inputs, loads) => page.evaluate(([preset, inputs, loads]) => {
  if(document.getElementById('tabPylon')) document.getElementById('tabPylon').click();
  applyPreset(preset);
  for(const [id, v] of Object.entries(inputs || {})) document.getElementById(id).value = v;
  if(loads){ LOADS = loads.map(L => ({Vux:0, Vuy:0, Tu:0, ...L})); drawLoads(); }
  render();
  const g = id => document.getElementById(id).value;
  return {fc: +g('fc'), fy: +g('fy'), fyt: +g('fyt'), B: +g('B'), H: +g('H'), lu: +g('lu'), covO: +g('covO'), tie: g('tieSize'),
          theta: +g('theta'), code: MODEL.opt.code};
}, [preset, inputs, loads]);

/* ------------------------------------------------------------------ */
test('H2：建築無側移柱彎矩放大 δ_ns = C_m/(1 − P_u/0.75P_c)（§6.6.4.5）', async () => {
  const L = {name:'A', Pu:200, Mux:10, Muy:0};
  const inp = await setup('bldg60', {lu:600, slFrame:'braced', slTreat:'mag', slK:1, slKy:1, slM12:-1, slBeta:0.6}, [L]);
  const r = await page.evaluate(() => { const q = MODEL.res[0]; return {d: q.Lk.slx.d, Mux: q.Lk.Mux, Muy: q.Lk.Muy, lam: MODEL.SL ? MODEL.SL.lamX : null}; });
  // 獨立計算：E_c = 12,000√f'c；(EI)_eff = 0.4E_cI_g/(1 + β_dns)；k l_u/r = 600/(0.3·60) = 33.3 > 34 + 12(−1) = 22 → 須考慮
  const Ec = 12000*Math.sqrt(inp.fc), Ig = inp.B*inp.H**3/12, EI = 0.4*Ec*Ig/1.6, Pc = Math.PI**2*EI/(1*600)**2;
  const Pu = L.Pu*1000, d = Math.max(1, 1.0/(1 - Pu/(0.75*Pc)));            // C_m = 0.6 − 0.4(−1) = 1.0
  const M2 = Math.max(L.Mux*1e5, Pu*(1.5 + 0.03*inp.H));
  near(r.d, d, 1e-9, 'δ_ns');
  near(r.Mux, d*M2, 1e-9, '放大後 M_ux');
  assert.equal(r.Muy, 0, '最小偏心距不需兩軸同時施加');
});

test('H2：建築有側移柱 δ_s 依 Q（§6.6.4.6.2(a)），Q 使 δ_s > 1.5 時改用 ΣP_c 法', async () => {
  const L = {name:'A', Pu:150, Mux:20, Muy:0};
  for(const Q of [0.2, 0.5]){
    const inp = await setup('bldg60', {lu:400, slFrame:'sway', slTreat:'mag', slK:1.6, slKy:1.6, slM12:-1, slBeta:0.6, slQ:Q}, [L]);
    const d = await page.evaluate(() => MODEL.res[0].Lk.slx.d);
    const Ec = 12000*Math.sqrt(inp.fc), Ig = inp.B*inp.H**3/12, Pu = L.Pu*1000;
    // 沿柱長（§6.6.4.6.4）：無側移公式、k = 1.0、β_dns = 0.6
    const PcN = Math.PI**2*(0.4*Ec*Ig/1.6)/(1*400)**2, dn = Math.max(1, 1/(1 - Pu/(0.75*PcN)));
    // 側移：(a) 1/(1 − Q) ≤ 1.5 才可用；否則 (b) 1/(1 − P_u/0.75P_c)，(EI)_eff 以 β_ds = 0、k = 1.6
    const PcS = Math.PI**2*(0.4*Ec*Ig)/(1.6*400)**2;
    const ds = 1/(1 - Q) <= 1.5 ? 1/(1 - Q) : 1/(1 - Pu/(0.75*PcS));
    near(d, ds*dn, 1e-9, `Q = ${Q} 之 δ`);
  }
});

test('H2：橋梁有支撐構材 δ_b = C_m/(1 − P_u/φP_c)，EI = (E_cI_g/2.5)/(1 + β_d)（公路橋梁設計規範 §7.3.5.2）', async () => {
  const L = {name:'A', Pu:300, Mux:150, Muy:0};
  const inp = await setup('pier100', {lu:1500, slFrame:'braced', slTreat:'mag', slK:1, slKy:1, slM12:-1, slBeta:0.6}, [L]);
  assert.equal(inp.code, 'bridge');
  const r = await page.evaluate(() => ({d: MODEL.res[0].Lk.slx.d, phi: +document.getElementById('phic').value}));
  const Ec = 12000*Math.sqrt(inp.fc), Ig = inp.B*inp.H**3/12, EI = Ec*Ig/2.5/1.6, Pc = Math.PI**2*EI/1500**2;
  const Cm = Math.max(0.4, 0.6 + 0.4*1);                                   // M_1b/M_2b = +1（單曲度）
  near(r.d, Math.max(1, Cm/(1 - L.Pu*1000/(r.phi*Pc))), 1e-9, 'δ_b');
});

test('H2：聲明「輸入已含二階效應」時不放大；δ > 1.4（§6.2.5.3）與 P_u ≥ 0.75P_c 判為不合格', async () => {
  const L = {name:'A', Pu:200, Mux:10, Muy:0};
  await setup('bldg60', {lu:600, slFrame:'braced', slTreat:'second'}, [L]);
  const r1 = await page.evaluate(() => ({d: MODEL.res[0].Lk.slx.d, Mux: MODEL.res[0].Lk.Mux, txt: document.getElementById('verdict').innerText}));
  assert.equal(r1.d, 1); near(r1.Mux, 10e5, 1e-12, '未放大之 M_ux');
  assert.match(r1.txt, /輸入含二階/);
  // 二階彎矩超過一階 1.4 倍（§6.2.5.3）→ 不合格：l_u = 900 cm 時 δ = 1/(1 − P_u/0.75P_c) ≈ 1.68
  await setup('bldg60', {lu:900, slFrame:'braced', slTreat:'mag', slK:1, slKy:1, slM12:-1, slBeta:0.6}, [L]);
  const r3 = await page.evaluate(() => ({d: MODEL.res[0].Lk.slx.d, bad: colAlertItems(MODEL).some(a => a.lv === 'bad' && /1\.4 倍/.test(a.t))}));
  assert.ok(r3.d > 1.4 && Number.isFinite(r3.d), `δ = ${r3.d}`); assert.ok(r3.bad, 'δ > 1.4 應判不合格');
  // 極長柱：P_u ≥ 0.75P_c → δ = ∞（構材挫屈）
  await setup('bldg60', {lu:3000, slFrame:'braced', slTreat:'mag'}, [L]);
  const r2 = await page.evaluate(() => ({dc: MODEL.ctrlFlex.dc, bad: colAlertItems(MODEL).filter(a => a.lv === 'bad').length}));
  assert.equal(r2.dc, Infinity); assert.ok(r2.bad > 0, '應判為不合格');
});

/* ------------------------------------------------------------------ */
test('H3/H4：空白、箱型壁厚過大、非整數根數皆停止計算並禁止輸出', async () => {
  for(const [preset, id, v] of [['bldg60', 'B', ''], ['box300', 'tw', '160'], ['bldg60', 'nB', '3.5']]){
    await setup(preset, {[id]: v});
    const r = await page.evaluate(id => ({ids: MODEL.badIn.map(b => b.id), blocked: isBlocked('pylon'),
      ws: document.getElementById('wsP').classList.contains('blocked'), first: colAlertItems(MODEL)[0],
      n: colAlertItems(MODEL).length, el: document.getElementById(id).classList.contains('inp-bad'),
      derived: getComputedStyle(document.getElementById('derivedP')).display}), id);
    assert.ok(r.ids.includes(id), `${preset}「${id}」= ${JSON.stringify(v)} 應標記為無效輸入`);
    assert.ok(r.blocked && r.ws && r.el, '結果區應封鎖、欄位應標紅');
    assert.equal(r.derived, 'none', '以佔位值算出之導出值（β₁、E_c、A_st…）不得顯示');
    assert.equal(r.first.lv, 'bad'); assert.equal(r.n, 1, '停止計算時不得列出以佔位值算出之其他警示');
    // Excel 匯出被擋下（彈出說明、不產生下載）
    let msg = null; const onDlg = d => { msg = d.message(); };
    page.on('dialog', onDlg);
    let downloaded = false; const onDl = () => { downloaded = true; };
    page.on('download', onDl);
    await page.click('#btnXlsx'); await page.waitForTimeout(500);
    page.off('dialog', onDlg); page.off('download', onDl);
    assert.ok(msg && msg.includes('輸入不完整'), 'Excel 匯出應被擋下並說明原因');
    assert.equal(downloaded, false);
    const opt = await page.evaluate(() => { optRun('pylon'); return document.getElementById('optOutP').innerText; });
    assert.match(opt, /輸入不完整/);
  }
  // 修正後恢復
  await setup('bldg60');
  assert.equal(await page.evaluate(() => isBlocked('pylon') || getComputedStyle(document.getElementById('derivedP')).display === 'none'), false);
});

test('H3/H4：梁翼緣厚 ≥ 全深時停止計算（不再默默改為 h/3）', async () => {
  const r = await page.evaluate(() => { document.getElementById('tabBeam').click(); applyBPreset('deckT');
    document.getElementById('bhf').value = 120; renderBeam();
    const out = {ids: BMODEL.badIn.map(b => b.id), blocked: isBlocked('beam')};
    applyBPreset('deckT'); document.getElementById('tabPylon').click(); return out; });
  assert.ok(r.ids.includes('bhf') && r.blocked);
});

/* ------------------------------------------------------------------ */
test('H5：柱扭力縱筋 A_ℓ、A_ℓ,min（§9.6.4.3）與扣除後 D/C（獨立求解）', async () => {
  const loads = [{name:'T', Pu:180, Mux:30, Muy:12, Vux:10, Vuy:6, Tu:8}];
  const inp = await setup('bldg60', {slTreat:'second'}, loads);
  const r = await page.evaluate(() => { const M = MODEL, q = M.res[0], m = M.m;
    return {AlReq: q.AlReq, dc0: q.dc0, dc: q.dc, phiv: m.phiv, Ast: M.R.bars.reduce((a, b) => a + b.A, 0), tors: M.shX.all[0].torsion,
      sec: {B: M.s.B, H: M.s.H, fc: m.fc, fy: m.fy, Es: m.Es, ecu: m.ecu, phic: m.phic, phit: m.phit, pmaxf: m.pmaxf},
      bars: M.R.bars.map(b => ({x: b.x, y: b.y, A: b.A})), Lk: {Pu: q.Lk.Pu, Mux: q.Lk.Mux, Muy: q.Lk.Muy}}; });
  assert.ok(r.tors, 'T_u = 8 tf·m 應超過 φT_th');
  // 獨立計算（kgf、cm）
  const dt = BAR_D[inp.tie], ch = inp.covO + dt/2, x1 = inp.B - 2*ch, y1 = inp.H - 2*ch;
  const Aoh = x1*y1, ph = 2*(x1 + y1), Ao = 0.85*Aoh, cot = 1/Math.tan(inp.theta*Math.PI/180);
  const AtS = 8e5/(2*r.phiv*Ao*inp.fyt*cot), Al = AtS*ph*(inp.fyt/inp.fy)*cot*cot;
  const Acp = inp.B*inp.H, base = 1.33*Math.sqrt(inp.fc)*Acp/inp.fy;
  const AlMin = Math.max(0, Math.min(base - AtS*ph*inp.fyt/inp.fy, base - (1.75*inp.B/inp.fyt)*ph*inp.fyt/inp.fy));
  const AlReq = Math.max(Al, AlMin);
  near(r.AlReq, AlReq, 1e-9, 'A_ℓ,req = max(A_ℓ, A_ℓ,min)');
  // 主筋等比例扣除 A_ℓ,req 後之 D/C：獨立求解器
  const kf = 1 - AlReq/r.Ast;
  const O = makeSection({...r.sec, bars: r.bars.map(b => ({...b, A: b.A*kf})), nf: 80});
  const ex = O.dc(r.Lk.Pu, r.Lk.Mux, r.Lk.Muy);
  assert.ok(r.dc > r.dc0, '扣除扭力縱筋後 D/C 應增加');
  near(r.dc, ex, 0.015, '扣除 A_ℓ 後之 D/C 與獨立求解');
});

test('H5：箱型 A_cp 取外周所圍面積（含中空，§9.6.4.3 原文；中空以 A_g 取代僅適用 T_th）', async () => {
  const inp = await setup('box300');
  const r = await page.evaluate(() => ({Acp: MODEL.torsL.Acp, any: MODEL.torsL.any, bad: colAlertItems(MODEL).filter(a => a.lv === 'bad').map(a => a.t)}));
  assert.ok(r.any); assert.equal(r.Acp, inp.B*inp.H);
  assert.deepEqual(r.bad, [], '範本 box300 已依 A_ℓ,min 配足主筋');
});

test('H5：T 梁扭矩 A_cp、p_cp 依 §9.2.4.4 計入翼板外懸 min(h − h_f, 4h_f)', async () => {
  const r = await page.evaluate(() => { document.getElementById('tabBeam').click(); applyBPreset('deckT');
    const g = id => +document.getElementById(id).value, B = BMODEL;
    const out = {bw: g('bw'), h: g('bh'), hf: g('bhf'), be: B.S.be, fc: g('bfc'), phiv: B.m.phiv, Acp: B.sh.Acp, pcp: B.sh.pcp, Tthr: B.sh.Tthr, tors: B.torsL.any};
    applyBPreset('deckT'); document.getElementById('tabPylon').click(); return out; });
  const ov = Math.min(r.h - r.hf, 4*r.hf, (r.be - r.bw)/2);
  const A1 = r.bw*r.h, p1 = 2*(r.bw + r.h), A2 = A1 + 2*ov*r.hf, p2 = p1 + 4*ov;
  assert.ok(A2*A2/p2 > A1*A1/p1, '本例計入翼板後 A_cp²/p_cp 較大');
  near(r.Acp, A2, 1e-12, 'A_cp'); near(r.pcp, p2, 1e-12, 'p_cp');
  near(r.Tthr, r.phiv*0.265*Math.sqrt(r.fc)*A2*A2/p2, 1e-9, 'φT_th');
  assert.equal(r.tors, false, 'T_u = 3 tf·m < φT_th：扭矩可忽略');
});

/* ------------------------------------------------------------------ */
test('H6：強柱弱梁 ΣM_nc ≥ 1.2ΣM_nb（§18.4.3.2）合格／不合格／得免', async () => {
  await setup('bldg60');
  const W0 = await page.evaluate(() => ({X: MODEL.scwb.X.Mnc, Y: MODEL.scwb.Y.Mnc, nc: MODEL.scwb.nc, exempt: MODEL.scwb.exempt}));
  assert.equal(W0.exempt, false);
  for(const [k, ok] of [[1.21, true], [1.19, false]]){
    const MbX = W0.nc*W0.X/k/1e5, MbY = W0.nc*W0.Y/k/1e5;
    await setup('bldg60', {scwbMbX: MbX, scwbMbY: MbY});
    const r = await page.evaluate(() => ({ok: MODEL.scwb.ok, bad: colAlertItems(MODEL).some(a => a.lv === 'bad' && /強柱弱梁/.test(a.t))}));
    assert.equal(r.ok, ok, `ΣM_nc/ΣM_nb = ${k}`); assert.equal(r.bad, !ok);
  }
  // 軸壓 ≤ A_gf'c/10 得免
  const inp = await setup('bldg60', {scwbMbX: 999, scwbMbY: 999}, [{name:'低軸力', Pu:50, Mux:20, Muy:5}]);
  assert.ok(50*1000 <= inp.B*inp.H*inp.fc/10);
  assert.equal(await page.evaluate(() => MODEL.scwb.exempt && MODEL.scwb.ok), true);
});

test('H6：未涵蓋項目以「未涵蓋」顯示，不得呈現為全數合格', async () => {
  await setup('bldg60');
  const t1 = await page.evaluate(() => document.getElementById('verdict').innerText);
  assert.match(t1, /未涵蓋/);
  await setup('pier100');
  const r = await page.evaluate(() => ({t: document.getElementById('verdict').innerText, list: colUncovered(MODEL)}));
  assert.match(r.t, /未涵蓋/);
  assert.ok(r.list.some(s => /§5\.3\.5/.test(s)), '橋梁應列出圍束鋼筋延伸入接頭（§5.3.5）');
});

test('H6：橋梁塑鉸區內不續接（§5.3.6），第三類續接器亦同；建築第三類可設於任何位置', async () => {
  const where = async (preset, sp) => { await setup(preset, {dSplice: sp});
    return page.evaluate(() => { const c = MODEL.devPlan.checks.find(q => /續接器位置|搭接位置/.test(q.item)); return c ? c.msg : ''; }); };
  assert.match(await where('pier100', 'm3'), /塑鉸區外（§5\.3\.6）/);
  assert.match(await where('pier100', 'B'), /塑鉸區外（§5\.3\.6）/);
  assert.match(await where('bldg60', 'm3'), /第三類可設於任何位置/);
});

/* ------------------------------------------------------------------ */
test('H8：建議配筋遵守目標 D/C 上限；用鋼量含搭接、錨定與彎鉤', async () => {
  for(const DCT of [0.95, 0.8]){
    const r = await page.evaluate(DCT => { applyPreset('bldg60'); document.getElementById('optDc').value = DCT; render();
      const R = colOptimize();
      return R.main.map(g => ({dc: g.M.ctrlFlex.dc, kg: g.kgMain, Ast: g.M.Ast, total: g.total, tie: g.tie ? g.tie.kg : null})); }, DCT);
    assert.ok(r.length > 0, `目標 D/C ${DCT} 應有候選`);
    for(const g of r){
      assert.ok(g.dc <= DCT + 1e-9, `候選 D/C ${g.dc} 超過目標 ${DCT}`);
      // 直段每公尺重量 A_st × 100 cm × 0.00785 kg/cm³；下料長度含搭接與錨定，必大於直段
      assert.ok(g.kg > g.Ast*100*0.00785, `主筋用量 ${g.kg} kg/m 應大於直段 ${g.Ast*0.785} kg/m`);
    }
    for(let i=1; i<r.length; i++) assert.ok(r[i-1].total <= r[i].total + 1e-9, '依合計用鋼量排序');
  }
  const b = await page.evaluate(() => { document.getElementById('tabBeam').click(); applyBPreset('bldgBeam');
    document.getElementById('bOptDc').value = 0.85; renderBeam(); const R = beamOptimize();
    const out = [...R.bot, ...R.top].map(g => g.B.ctrl ? g.B.ctrl.dc : 0);
    document.getElementById('bOptDc').value = 0.95; applyBPreset('bldgBeam'); document.getElementById('tabPylon').click(); return out; });
  assert.ok(b.length > 0); for(const dc of b) assert.ok(dc <= 0.85 + 1e-9, `梁候選 D/C ${dc} 超過 0.85`);
  await page.evaluate(() => { document.getElementById('optDc').value = 0.95; applyPreset('bldg60'); });
});

test('H8：橫向筋每組長度含 135° 彎鉤延伸 max(6d_t, 7.5 cm)', async () => {
  const inp = await setup('bldg60');
  const r = await page.evaluate(() => ({L: colTieLen(MODEL), nx: MODEL.conf.tieX.length, ny: MODEL.conf.tieY.length}));
  const dt = BAR_D[inp.tie], hk = Math.max(6*dt, 7.5), cx = inp.B - 2*inp.covO - dt, cy = inp.H - 2*inp.covO - dt;
  near(r.L, 2*(cx + cy) + 2*hk + r.nx*(cy + 2*hk) + r.ny*(cx + 2*hk), 1e-12, '每組橫向筋長度');
  const html = await page.evaluate(() => { optRun('pylon'); return document.getElementById('optOutP').parentElement.textContent; });
  assert.match(html, /僅供初步配筋/);
});

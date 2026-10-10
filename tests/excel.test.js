/* Excel 計算書回歸：匯出 .xlsx → LibreOffice 重算 → 與網頁值比較（C1～C4、H2、H5、M1～M8 之 Excel 公式）。
   需要 LibreOffice（soffice）；未安裝時略過。 */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs'), os = require('os'), path = require('path'), {execFileSync} = require('child_process');
const ExcelJS = require('exceljs');
const {openApp} = require('./helpers');

const hasSoffice = (() => { try{ execFileSync('soffice', ['--version'], {stdio:'ignore', timeout:60000}); return true; }catch(e){ return false; } })();
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'rcsd-xl-'));

const val = v => (v && typeof v === 'object') ? ('result' in v ? v.result : v.richText ? v.richText.map(t => t.text).join('') : v) : v;
/* 「檢核表」：A 欄項目名稱、C 欄數值 */
function item(wb, label){
  const ws = wb.getWorksheet('檢核表'); let out;
  ws.eachRow(row => { if(out === undefined && String(val(row.getCell(1).value) || '').trim() === label) out = val(row.getCell(3).value); });
  assert.notEqual(out, undefined, `檢核表找不到「${label}」`);
  return out;
}

test('Excel 計算書與網頁一致（C1～C4、H2、H5、M1～M8）', {skip: !hasSoffice && '未安裝 LibreOffice（soffice）'}, async () => {
  const app = await openApp(), page = app.page;
  const exportAs = async name => {
    const [dl] = await Promise.all([page.waitForEvent('download', {timeout:120000}), page.click('#btnXlsx')]);
    await dl.saveAs(path.join(TMP, name + '.xlsx'));
  };
  const web = {};
  try{
    // C1：40×90 建築柱
    await page.evaluate(() => { applyPreset('bldg60'); const set = (id, v) => { document.getElementById(id).value = v; };
      set('B', 40); set('H', 90); set('nB', 3); set('nH', 6); render(); });
    web.rect = await page.evaluate(() => ({VeX: MODEL.shX.Ve/1000, VeY: MODEL.shY.Ve/1000, sX: MODEL.shX.s_str, sY: MODEL.shY.s_str}));
    await exportAs('rect');
    // C2、C3：建築方柱 + T_u = 6
    await page.evaluate(() => { applyPreset('bldg60'); LOADS.forEach(L => { L.Tu = 6; }); drawLoads(); render(); });
    web.tors = await page.evaluate(() => ({sX: MODEL.shX.s_str, sY: MODEL.shY.s_str, dc: MODEL.res.map(q => q.dc)}));
    await exportAs('tors');
    // H2：長細效應彎矩放大（單軸載重 → D/C 由 Excel 公式計算，可比對 δ 與 D/C）
    for(const [nm, k, loads] of [['slBr', 'pier100', [{name:'X', Pu:280, Mux:196, Muy:0, Vux:52, Vuy:0, Tu:0}, {name:'Y', Pu:120, Mux:0, Muy:152, Vux:0, Vuy:44, Tu:0}]],
                                 ['slBd', 'pond80', [{name:'X', Pu:430, Mux:58, Muy:0, Vux:19, Vuy:0, Tu:0}, {name:'0', Pu:520, Mux:0, Muy:0, Vux:0, Vuy:0, Tu:0}]]]){
      await page.evaluate(([k, loads]) => { applyPreset(k); LOADS = loads; drawLoads(); render(); }, [k, loads]);
      web[nm] = await page.evaluate(() => MODEL.res.map(q => ({dc: q.dc, dx: q.Lk.slx.d, dy: q.Lk.sly.d})));
      await exportAs(nm);
    }
    // C4：梁兩組合
    await page.evaluate(() => { document.getElementById('tabBeam').click(); applyBPreset('bldgBeam');
      BLOADS = [{name:'A', Mpos:20, Mneg:35, Vu:30, Tu:0}, {name:'B', Mpos:20, Mneg:35, Vu:28, Tu:4}]; drawBLoads(); renderBeam(); });
    web.beam = await page.evaluate(() => ({ctrl: BMODEL.sh.ctrl.L.name, sStr: BMODEL.sh.sStr, sUse: BMODEL.fin.sUse,
      s2gov: BMODEL.fin2.gov.v, s2: BMODEL.fin2.sUse}));
    await exportAs('beam');
    // H5：T 梁 A_cp 依 §9.2.4.4 計入翼板（T_u 提高至須設計扭矩，D/C 含扣除扭力縱筋）
    await page.evaluate(() => { applyBPreset('deckT'); BLOADS.forEach(L => { L.Tu = 6; }); drawBLoads(); renderBeam(); });
    web.deck = await page.evaluate(() => ({Acp: BMODEL.sh.Acp, Tth: BMODEL.sh.Tthr/1e5, dc: BMODEL.ctrl.dc, tors: BMODEL.torsL.any}));
    await exportAs('deck');
    // M1、M2：邊梁有效翼寬（表 6.3.2.1）與裂縫控制 c_c
    await page.evaluate(() => { applyBPreset('deckT'); document.getElementById('bEdge').value = '1'; document.getElementById('bSclear').value = 150; renderBeam(); });
    web.edge = await page.evaluate(() => ({be: BMODEL.S.be, sLim: BMODEL.crack.sLim}));
    await exportAs('edge');
    // M3、M7、M8：橋梁 φ 依 §5.3.2、靜載重軸力、彈性剪力上限、指定 P_e；寬扁斷面低軸壓
    await page.evaluate(() => { document.getElementById('tabPylon').click(); applyPreset('pier100'); const s = (i, v) => { document.getElementById(i).value = v; };
      s('slTreat', 'second'); s('brPhiMode', '532'); s('brMnP', 'dead'); s('brPD', 150); s('brVelX', 30); s('brPe', 250);
      LOADS = [{name:'A', Pu:150, Mux:200, Muy:0, Vux:0, Vuy:20, Tu:0}, {name:'B', Pu:600, Mux:180, Muy:0, Vux:10, Vuy:22, Tu:0}]; drawLoads(); render(); });
    web.br = await page.evaluate(() => ({dc: MODEL.res.map(q => q.dc), VeX: MODEL.shX.Ve/1000, VeY: MODEL.shY.Ve/1000, pe: MODEL.conf.peTerm}));
    await exportAs('br');
    await page.evaluate(() => { applyPreset('bldg60'); const s = (i, v) => { document.getElementById(i).value = v; };
      s('B', 600); s('H', 60); s('nB', 10); s('nH', 2); s('slTreat', 'second');
      LOADS = [{name:'A', Pu:5, Mux:30, Muy:0, Vux:0, Vuy:5, Tu:0}]; drawLoads(); render(); });
    web.flat = await page.evaluate(() => MODEL.res[0].dc);
    await exportAs('flat');
    // M6：頂／底配對之 M1/M2 與 Cm
    await page.evaluate(() => { applyPreset('bldg60'); const s = (i, v) => { document.getElementById(i).value = v; };
      s('lu', 600); s('slFrame', 'braced'); s('slTreat', 'mag');
      LOADS = [{name:'E', pos:'top', Pu:200, Mux:20, Muy:6, Vux:3, Vuy:10, Tu:0}, {name:'E', pos:'bot', Pu:205, Mux:-12, Muy:4, Vux:3, Vuy:10, Tu:0}];
      drawLoads(); render(); });
    web.pair = await page.evaluate(() => MODEL.res.map(q => ({dc: q.dc, dx: q.Lk.slx.d, dy: q.Lk.sly.d})));
    await exportAs('pair');
  }finally{ await app.close(); }

  const out = path.join(TMP, 'out');
  execFileSync('soffice', ['--headless', '--convert-to', 'xlsx', '--outdir', out, ...['rect', 'tors', 'beam', 'slBr', 'slBd', 'deck', 'edge', 'br', 'flat', 'pair'].map(n => path.join(TMP, n + '.xlsx'))],
    {stdio:'ignore', timeout:300000});
  const load = async n => { const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(path.join(out, n + '.xlsx')); return wb; };
  const near = (a, b, rel, msg) => assert.ok(Math.abs(a - b) <= rel*Math.abs(b), `${msg}：Excel ${a}，網頁 ${b}`);

  const rect = await load('rect');
  near(item(rect, 'X 向容量設計剪力 Ve'), web.rect.VeX, 0.005, 'X 向 Ve');
  near(item(rect, 'Y 向容量設計剪力 Ve'), web.rect.VeY, 0.005, 'Y 向 Ve');
  near(item(rect, 'Y 向強度需求間距（各組合最小）'), web.rect.sY, 0.005, 'Y 向強度需求間距');

  const tors = await load('tors');
  near(item(tors, 'X 向強度需求間距（各組合最小）'), web.tors.sX, 0.005, 'X 向強度需求間距（外圍閉合肢）');
  const L = tors.getWorksheet('載重組合');
  web.tors.dc.forEach((dc, i) => near(val(L.getCell('M' + (4 + i)).value), dc, 0.005, `第 ${i+1} 組 D/C`));

  for(const nm of ['slBr', 'slBd']){
    const L2 = (await load(nm)).getWorksheet('載重組合');
    web[nm].forEach((w, i) => {
      assert.ok(w.dx > 1 || w.dy > 1 || i > 0, `${nm} 第 1 組應有彎矩放大`);
      near(val(L2.getCell('AS' + (4 + i)).value), w.dx, 0.002, `${nm} 第 ${i+1} 組 δx`);
      near(val(L2.getCell('AT' + (4 + i)).value), w.dy, 0.002, `${nm} 第 ${i+1} 組 δy`);
      near(val(L2.getCell('M' + (4 + i)).value), w.dc, 0.005, `${nm} 第 ${i+1} 組 D/C（放大後）`);
    });
  }

  const beam = await load('beam');
  assert.equal(item(beam, '剪扭控制組合'), web.beam.ctrl);
  near(item(beam, '剪扭強度需求間距'), web.beam.sStr, 0.001, '梁剪扭強度需求間距');
  assert.equal(item(beam, '採用箍筋間距（梁）'), web.beam.sUse);
  near(item(beam, '加密區外控制需求間距'), web.beam.s2gov, 0.001, '梁加密區外控制需求間距');

  const deck = await load('deck');
  assert.ok(web.deck.tors, 'deckT（T_u = 6）應須設計扭矩');
  near(item(deck, '扭矩：外周包圍面積'), web.deck.Acp, 1e-9, 'T 梁 A_cp（§9.2.4.4）');
  near(item(deck, '可忽略扭矩門檻 φTth'), web.deck.Tth, 1e-6, 'T 梁 φT_th');
  near(item(deck, '撓曲 D/C'), web.deck.dc, 0.002, 'T 梁 D/C（含扣除扭力縱筋）');

  const edge = await load('edge');
  near(item(edge, '有效翼緣寬'), web.edge.be, 1e-9, '邊梁 b_e（表 6.3.2.1）');
  near(item(edge, '鋼筋中心距上限'), web.edge.sLim, 1e-6, '裂縫控制間距上限（c_c = 保護層 + d_t）');

  const br = await load('br'), Lb = br.getWorksheet('載重組合');
  web.br.dc.forEach((dc, i) => near(val(Lb.getCell('M' + (4 + i)).value), dc, 0.006, `§5.3.2 φ 第 ${i+1} 組 D/C`));
  near(item(br, 'X 向容量設計剪力 Ve'), web.br.VeX, 0.002, 'X 向 V_e（彈性剪力上限）');
  near(item(br, 'Y 向容量設計剪力 Ve'), web.br.VeY, 0.005, 'Y 向 V_e（靜載重軸力 P_D）');

  const flat = await load('flat');
  near(val(flat.getWorksheet('載重組合').getCell('M4').value), web.flat, 0.01, '寬扁斷面低軸壓 D/C（有限值）');

  const pr = (await load('pair')).getWorksheet('載重組合');
  web.pair.forEach((w, i) => {
    near(val(pr.getCell('AS' + (4 + i)).value), w.dx, 0.002, `頂／底配對第 ${i+1} 列 δx`);
    near(val(pr.getCell('AT' + (4 + i)).value), w.dy, 0.002, `頂／底配對第 ${i+1} 列 δy`);
    near(val(pr.getCell('M' + (4 + i)).value), w.dc, 0.005, `頂／底配對第 ${i+1} 列 D/C`);
  });
});

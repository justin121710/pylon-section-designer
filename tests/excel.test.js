/* Excel 計算書回歸：匯出 .xlsx → LibreOffice 重算 → 與網頁值比較（C1～C4 之 Excel 公式）。
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

test('Excel 計算書與網頁一致（C1～C4）', {skip: !hasSoffice && '未安裝 LibreOffice（soffice）'}, async () => {
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
    // C4：梁兩組合
    await page.evaluate(() => { document.getElementById('tabBeam').click(); applyBPreset('bldgBeam');
      BLOADS = [{name:'A', Mpos:20, Mneg:35, Vu:30, Tu:0}, {name:'B', Mpos:20, Mneg:35, Vu:28, Tu:4}]; drawBLoads(); renderBeam(); });
    web.beam = await page.evaluate(() => ({ctrl: BMODEL.sh.ctrl.L.name, sStr: BMODEL.sh.sStr, sUse: BMODEL.fin.sUse,
      s2gov: BMODEL.fin2.gov.v, s2: BMODEL.fin2.sUse}));
    await exportAs('beam');
  }finally{ await app.close(); }

  const out = path.join(TMP, 'out');
  execFileSync('soffice', ['--headless', '--convert-to', 'xlsx', '--outdir', out, ...['rect', 'tors', 'beam'].map(n => path.join(TMP, n + '.xlsx'))],
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

  const beam = await load('beam');
  assert.equal(item(beam, '剪扭控制組合'), web.beam.ctrl);
  near(item(beam, '剪扭強度需求間距'), web.beam.sStr, 0.001, '梁剪扭強度需求間距');
  assert.equal(item(beam, '採用箍筋間距（梁）'), web.beam.sUse);
  near(item(beam, '加密區外控制需求間距'), web.beam.s2gov, 0.001, '梁加密區外控制需求間距');
});

/* 回歸測試（介面修正第二輪）：
     需確認區塊：勾選框不撐滿、文字欄有足夠寬度
     建議配筋：所列組合套用後無任何不合格（含目前配筋本身不合格時）
     計算書不含 emoji 或打勾／打叉圖示
     參數表：值（含下拉選項全文）、單位、狀態於各視窗寬度不被截斷
     設計對象視窗：「取消」與 Esc 還原開啟前內容，「直接進入主畫面」保留變更 */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const {openApp} = require('./helpers');

let app, page;
test.before(async () => { app = await openApp(); page = app.page; });
test.after(async () => { if(app){ assert.deepEqual(app.errors, [], '頁面不得有 JavaScript 錯誤'); await app.close(); } });

test('需確認區塊：勾選框不撐滿，項目文字完整可讀', async () => {
  for(const w of [1440, 1024]){
    await page.setViewportSize({width:w, height:900});
    await page.evaluate(() => { document.getElementById('tabPylon').click(); applyPreset('bldg60'); render(); });
    await page.click('#wsP button[role=tab][data-pane="warn"], button[role=tab][data-pane="warn"] >> visible=true');
    const r = await page.evaluate(() => [...document.querySelectorAll('#confirmP .conf-row')].map(row => ({
      box: row.querySelector('input').getBoundingClientRect().width, txt: row.querySelector('.conf-t').getBoundingClientRect().width,
      row: row.getBoundingClientRect().width, h: row.getBoundingClientRect().height})));
    assert.ok(r.length >= 1, '應有需確認項目');
    for(const x of r){ assert.ok(x.box <= 20, `勾選框寬 ${x.box}`); assert.ok(x.txt >= 0.6*x.row, `文字欄 ${x.txt} / 列寬 ${x.row}`); assert.ok(x.h < 60, `列高 ${x.h}`); }
  }
  await page.setViewportSize({width:1440, height:900});
});

test('建議配筋：所列組合套用後無任何不合格', async () => {
  const r = await page.evaluate(() => {
    const out = {};
    document.getElementById('tabPylon').click(); applyPreset('bldg60'); render();
    const badCol = ovr => withInputs(ovr, () => colAlertItems(colPipeline()).filter(a => a.lv==='bad').length);
    let R = colOptimize();
    out.col = {n: R.main.length + R.ties.length, bad: [...R.main.map(g => badCol(g.ovrAll)), ...R.ties.map(g => badCol(g.ovr))].filter(Boolean).length};
    // 目前主筋不足（本身不合格）：只調橫向筋無法合格 → 不列橫向筋；主筋＋橫向筋組合仍須全數合格
    document.getElementById('nB').value = 2; document.getElementById('nH').value = 2; render();
    R = colOptimize();
    out.colWeak = {cur: colAlertItems(MODEL).filter(a => a.lv==='bad').length, main: R.main.length, ties: R.ties.length,
      bad: R.main.map(g => badCol(g.ovrAll)).filter(Boolean).length, note: R.tieNote.join('')};
    document.getElementById('tabBeam').click(); applyBPreset('bldgBeam'); document.getElementById('bNtop').value = 2; renderBeam();
    R = beamOptimize();
    const rows = [...R.bot, ...R.top];
    out.beam = {cur: optAllBadBeam({}).length, n: rows.length, bad: rows.map(g => optAllBadBeam(g.ovrAll).length).filter(Boolean).length,
      ties: R.ties.length, other: rows.filter(g => g.withOther).length};
    applyBPreset('bldgBeam'); renderBeam(); document.getElementById('tabPylon').click(); applyPreset('bldg60'); render();
    return out;
  });
  assert.ok(r.col.n > 0); assert.equal(r.col.bad, 0, '柱建議皆無不合格');
  assert.ok(r.colWeak.cur > 0 && r.colWeak.main > 0); assert.equal(r.colWeak.bad, 0); assert.equal(r.colWeak.ties, 0); assert.match(r.colWeak.note, /無法使全部檢核合格|皆不合格/);
  assert.ok(r.beam.cur > 0 && r.beam.n > 0); assert.equal(r.beam.bad, 0, '梁建議皆無不合格'); assert.equal(r.beam.ties, 0); assert.ok(r.beam.other > 0, '底筋建議搭配頂筋一併套用');
});

test('計算書不含 emoji 或打勾／打叉圖示', async () => {
  const hits = await page.evaluate(() => {
    const RX = /[\u{1F000}-\u{1FFFF}\u{2600}-\u{27BF}\u{2B00}-\u{2BFF}\u{FE0F}\u{2139}]/u, out = [];
    const chk = tag => { const t = document.getElementById('report').textContent; const i = t.search(RX); if(i >= 0) out.push(tag + '：' + t.slice(Math.max(0, i-30), i+5)); };
    document.getElementById('tabPylon').click();
    for(const k of ['bldg60', 'pier100', 'box300']){ applyPreset(k); LOADS.forEach(L => { L.Tu = 40; L.Vux *= 4; L.Vuy *= 4; }); drawLoads(); buildReport(render()); chk('柱 ' + k); }
    document.getElementById('tabBeam').click();
    for(const k of ['bldgBeam', 'deckT', 'pondWall']){ applyBPreset(k); BLOADS.forEach(L => { L.Vu = (L.Vu||0)*4; }); drawBLoads(); buildBeamReport(renderBeam()); chk('梁 ' + k); }
    applyBPreset('bldgBeam'); renderBeam(); document.getElementById('tabPylon').click(); applyPreset('bldg60'); render();
    return out;
  });
  assert.deepEqual(hits, []);
});

test('參數表：值、下拉選項全文、單位與狀態於各寬度皆不被截斷', async () => {
  const probe = () => {
    const cv = document.createElement('canvas').getContext('2d'), out = [];
    for(const row of document.querySelectorAll('.pin-row')){
      if(!row.offsetParent) continue;
      const body = row.closest('.pin-body'), right = body.getBoundingClientRect().left + body.clientWidth;
      const el = row.querySelector('.pin-v input, .pin-v select'), id = el ? el.id : '?';
      if(el && el.tagName==='SELECT'){ const cs = getComputedStyle(el); cv.font = `${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;
        const t = el.options[el.selectedIndex] ? el.options[el.selectedIndex].text : '';
        if(cv.measureText(t).width + parseFloat(cs.paddingLeft) + parseFloat(cs.paddingRight) > el.clientWidth + 1) out.push(`${id} 選項「${t}」被截`); }
      else if(el && el.type!=='checkbox' && el.scrollWidth > el.clientWidth + 1) out.push(`${id} 值被截`);
      if(el && el.getBoundingClientRect().right > right + 0.5) out.push(`${id} 值欄超出`);
      const u = row.querySelector('.pin-u'); if(u && u.offsetParent && (u.scrollWidth > u.clientWidth + 1 || u.getBoundingClientRect().right > right + 0.5)) out.push(`${id} 單位被截`);
      const c = row.querySelector('.pin-chip'); if(c && c.offsetParent){ const r = c.getBoundingClientRect(), cell = c.parentElement.getBoundingClientRect();
        if(r.right > Math.min(right, cell.right) + 0.5 || c.scrollWidth > c.clientWidth + 1) out.push(`${id} 狀態「${c.textContent}」被截`); }
    }
    return out;
  };
  const bad = new Set();
  for(const [w, h] of [[1920,1080],[1366,768],[1280,720],[1100,800],[1024,768],[900,800],[390,844]]){
    await page.setViewportSize({width:w, height:h}); await page.waitForTimeout(120);
    for(const k of ['bldg60','circ80','box300','pier100','pileC120']){
      await page.evaluate(k => { document.getElementById('tabPylon').click(); applyPreset(k); render(); }, k); await page.waitForTimeout(30);
      for(const x of await page.evaluate(probe)) bad.add(`[${w}] 柱 ${k} ${x}`);
    }
    for(const k of ['bldgBeam','deckT','pondWall']){
      await page.evaluate(k => { document.getElementById('tabBeam').click(); applyBPreset(k); renderBeam(); }, k); await page.waitForTimeout(30);
      for(const x of await page.evaluate(probe)) bad.add(`[${w}] 梁 ${k} ${x}`);
    }
  }
  await page.setViewportSize({width:1440, height:900});
  await page.evaluate(() => { applyBPreset('bldgBeam'); renderBeam(); document.getElementById('tabPylon').click(); applyPreset('bldg60'); render(); });
  assert.deepEqual([...bad], []);
});

test('設計對象視窗：取消／Esc 還原開啟前內容，直接進入主畫面保留變更', async () => {
  const st = () => page.evaluate(() => ({B: document.getElementById('B').value, n: LOADS.length, tab: TAB, open: setupOpen(), pre: document.getElementById('preset').value}));
  await page.evaluate(() => { document.getElementById('tabPylon').click(); applyPreset('bldg60'); render(); });
  const s0 = await st();
  await page.click('#btnSetup');
  assert.ok(await page.isVisible('#setupCancel'), '應有取消按鈕');
  await page.fill('#setupDims input[data-for="B"]', '85');
  await page.click('#setupGo');
  await page.evaluate(() => { LOADS.push({name:'X', Pu:1, Mux:0, Muy:0, Vux:0, Vuy:0, Tu:0}); drawLoads(); render(); });
  await page.click('#setupCancel');
  assert.deepEqual(await st(), s0, '取消：尺寸與載重還原');
  await page.click('#btnSetup'); await page.click('#tabBeam'); await page.keyboard.press('Escape');
  assert.deepEqual(await st(), s0, 'Esc：構件類型還原');
  await page.click('#btnSetup'); await page.fill('#setupDims input[data-for="B"]', '70'); await page.click('#setupSkip');
  assert.equal((await st()).B, '70', '直接進入主畫面：保留變更');
  await page.evaluate(() => { applyPreset('bldg60'); render(); });
});

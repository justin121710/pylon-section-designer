/* 建置操作教學 PDF：docs/RC斷面設計工具_操作教學.pdf
   1. 以本機 HTTP 伺服器開啟 index.html（CDN 改由本目錄 node_modules 提供），依教學步驟操作並截圖 → img/
   2. 匯出 Excel，以 LibreOffice 轉 PDF、pdftoppm 轉圖並裁白邊 → img/26、27
   3. 以 Chromium 列印 tutorial.html（KaTeX 排版公式與符號）為 A4 PDF
   4. 寫入 stamp.txt（輸入檔的 git blob 雜湊），供 check.sh／pre-commit／CI 判斷 PDF 是否為最新

   需求：Node 18+、本目錄 npm install、Chromium（npx playwright install chromium）、
         LibreOffice（soffice）、poppler-utils（pdftoppm）、python3 + Pillow。
   用法：cd docs/tutorial && npm install && npm run build */
'use strict';
const { chromium } = require('playwright');
const http = require('http'), path = require('path'), fs = require('fs'), { execFileSync } = require('child_process');

const HERE = __dirname, ROOT = path.resolve(HERE, '../..'), NM = path.join(HERE, 'node_modules');
const IMG = path.join(HERE, 'img'), XL = path.join(HERE, 'xl');
const OUT = path.join(ROOT, 'docs', 'RC斷面設計工具_操作教學.pdf');
const INPUTS = fs.readFileSync(path.join(HERE, 'inputs.txt'), 'utf8').split('\n').map(s => s.trim()).filter(s => s && !s.startsWith('#'));   // 影響 PDF 內容的檔案
const W = 1440, H = 900;

/* ---------- 靜態伺服器（repo 根目錄） ---------- */
const MIME = {'.html':'text/html; charset=utf-8', '.js':'text/javascript', '.css':'text/css', '.png':'image/png', '.json':'application/json',
  '.woff2':'font/woff2', '.woff':'font/woff', '.ttf':'font/ttf', '.svg':'image/svg+xml', '.pdf':'application/pdf'};
function serve(){
  const srv = http.createServer((req, res) => {
    const u = decodeURIComponent(new URL(req.url, 'http://x').pathname);
    const f = path.join(ROOT, path.normalize(u).replace(/^([/\\])+/, ''));
    if(!f.startsWith(ROOT) || !fs.existsSync(f) || fs.statSync(f).isDirectory()){ res.writeHead(404); return res.end(); }
    res.writeHead(200, {'content-type': MIME[path.extname(f)] || 'application/octet-stream'});
    fs.createReadStream(f).pipe(res);
  });
  return new Promise(ok => srv.listen(0, '127.0.0.1', () => ok(srv)));
}

/* ---------- 截圖 ---------- */
async function screenshots(browser, base){
  fs.rmSync(IMG, {recursive:true, force:true}); fs.mkdirSync(IMG, {recursive:true});
  const ctx = await browser.newContext({viewport:{width:W, height:H}, deviceScaleFactor:1.5, acceptDownloads:true});
  // CDN → 本機套件；其餘外部資源（Google Fonts 等）略過，與離線環境相同
  await ctx.route(/^https?:\/\/(?!127\.0\.0\.1)/, r => {
    const m = r.request().url().match(/cdn\.jsdelivr\.net\/npm\/([^@]+)@[^/]+\/(.*)$/);
    if(m && fs.existsSync(path.join(NM, m[1], m[2]))) return r.fulfill({path: path.join(NM, m[1], m[2])});
    return r.abort();
  });
  const page = await ctx.newPage();
  const errs = []; page.on('pageerror', e => errs.push(e.message));
  await page.goto(base + '/index.html'); await page.waitForFunction(() => window.katex && window.renderMathInElement); await page.waitForTimeout(1500);
  const W8 = ms => page.waitForTimeout(ms);
  // 紅框＋編號：[selector | rect, 編號, 外擴 px]
  const mark = items => page.evaluate(items => {
    const L = document.createElement('div'); L.id = '__mk'; L.style.cssText = 'position:fixed;inset:0;pointer-events:none;z-index:99999';
    for(const [sel, n, pad] of items){
      const e = document.querySelector(sel); if(!e) throw new Error('找不到標註目標 ' + sel);
      const r = e.getBoundingClientRect(), p = pad ?? 3;
      const b = document.createElement('div');
      b.style.cssText = `position:fixed;left:${r.left-p}px;top:${r.top-p}px;width:${r.width+2*p}px;height:${r.height+2*p}px;border:2.5px solid #e11d48;border-radius:7px;box-shadow:0 0 0 2px rgba(255,255,255,.7)`;
      const t = document.createElement('div'); t.textContent = n;
      const tx = Math.max(2, Math.min(window.innerWidth-28, r.left-p-11)), ty = Math.max(2, r.top-p-11);
      t.style.cssText = `position:fixed;left:${tx}px;top:${ty}px;width:24px;height:24px;border-radius:50%;background:#e11d48;color:#fff;font:700 14px/24px Arial;text-align:center;box-shadow:0 1px 3px rgba(0,0,0,.4)`;
      L.appendChild(b); L.appendChild(t);
    }
    document.body.appendChild(L);
  }, items);
  const unmark = () => page.evaluate(() => { const e = document.getElementById('__mk'); if(e) e.remove(); });
  const shot = async (name, items = [], clip) => {
    if(items.length) await mark(items);
    await page.screenshot({path: path.join(IMG, name + '.png'), clip});
    await unmark();
    console.log('  截圖', name);
  };
  const ws = pane => page.evaluate(p => document.querySelector((TAB==='beam'?'#wsB':'#wsP')+' .ws-tabs [data-pane="'+p+'"]').click(), pane);
  const box = sel => page.evaluate(s => { const r = document.querySelector(s).getBoundingClientRect(); return {x:r.left, y:r.top, width:r.width, height:r.height}; }, sel);
  const pane = async () => { const a = await box(TABSEL()); return {x:a.x-8, y:a.y-8, width:a.width+16, height:H-a.y}; };
  let TABSEL = () => '#wsP';
  const group = name => page.evaluate(n => [...document.querySelectorAll('#railP .pin-g')].find(d => d.textContent.includes(n)).click(), name);

  /* 1 設計對象 */
  await page.fill('#memName', 'C1'); await page.dispatchEvent('#memName', 'input');
  await shot('01-setup', [['#tabPylon',1],['#setupSecP',2],['#setupPreP',3],['#setupMemP',4],['#setupCode',5],['#setupEnv',6],['#setupGo',7]]);
  await page.click('#setupGo'); await W8(500);
  /* 2 畫面配置 */
  await shot('02-layout', [['#topbar .tb-actions',1],['#verdict',2],['#railP',3,2],['#sheetP',4,2],['#detailP',5,2],['#wsP .ws-tabs',6],['#btnSetup',7]]);
  /* 3 參數表 */
  await group('主筋'); await W8(300); await page.hover('.pin-row[data-pid="nB"]'); await W8(300);
  { const a = await box('#railP'), d = await box('#detailP');
    await shot('03-params', [['.pin-g.on',1],['.pin-row[data-pid="nB"] .pin-v',2],['.pin-row[data-pid="nB"] .pin-s',3],['#detailP',4,2]], {x:0, y:a.y-10, width:d.x+d.width+12, height:Math.min(H-a.y+10, 700)}); }
  /* 4 細長效應 */
  await group('細長效應'); await W8(300); await page.hover('.pin-row[data-pid="slK"]'); await W8(300);
  { const a = await box('#railP'), d = await box('#detailP');
    await shot('04-slender', [['.pin-row[data-pid="slFrame"] .pin-v',1],['.pin-row[data-pid="slK"] .pin-v',2],['.pin-row[data-pid="slK"] .pin-s',3],['#detailP',4,2]], {x:0, y:a.y-10, width:d.x+d.width+12, height:Math.min(H-a.y+10, 700)}); }
  await page.click('#railP .pin-all');
  /* 5 載重組合與 CSV */
  await page.click('#btnLoads'); await W8(400);
  await shot('05-loads', [['#loadTbl',1],['#btnAddRow',2],['#btnPaste',3]]);
  await page.click('#btnPaste'); await W8(200);
  await page.fill('#csvBox', '組合,Pu,Mux,Muy,Vux,Vuy,Tu\n1.2D+1.6L,260,18,10,8,5,0\n1.2D+L+Ex,210,42,12,24,6,0\n1.2D+L+Ey,205,12,40,6,23,0\n0.9D+Ex,95,38,11,22,6,0\n0.9D+Ey,92,11,36,6,21,0');
  await shot('06-csv', [['#csvBox',1],['#csvApply',2]]);
  await page.click('#csvApply'); await W8(300);
  await shot('07-csv-ask', [['#askAct',1]]);
  await page.click('#askAct button:has-text("取代全部")'); await W8(500);
  await shot('08-csv-done', [['#csvMsg',1],['#loadTbl tbody',2]]);
  await page.click('#loadClose'); await W8(300);
  /* 6 判定列、警示 */
  { const v = await box('#verdict'); await shot('09-verdict', [['#verdict',1,2]], {x:0, y:0, width:W, height:v.y+v.height+12}); }
  await ws('warn'); await W8(400);
  await shot('10-alerts', [['#alerts .alert',1],['#alerts .al-btn',2]], await pane());
  /* 7 斷面與 P-M */
  await ws('fig'); await W8(500);
  await shot('11-figs', [['#secWrap',1,0],['#axX',2],['#cTieBar .adv-tog',3]], await pane());
  await page.evaluate(() => document.querySelector('#secWrap').closest('.card').nextElementSibling.scrollIntoView({block:'start'})); await W8(300);
  { const a = await box('#wsP'), y = Math.max(0, a.y-8); await shot('12-pm', [], {x:a.x-8, y, width:a.width+16, height:H-y}); }
  await page.evaluate(() => { window.scrollTo(0,0); document.querySelectorAll('.figs, #wsP, .ws').forEach(e => e.scrollTop = 0); });
  /* 8–10 結果摘要、軸向配置、伸展長度 */
  for(const [p, n, items] of [['sum','13-summary',[]],['elev','14-elev',[['#elevInfo',1]]],['dev','15-dev',[['#devExp',1]]]]){
    await ws(p); await W8(500); await shot(n, items, await pane());
  }
  /* 11 計算書預覽 */
  await ws('fig');
  await page.click('#btnPreview'); await W8(2000);
  await shot('16-report', [['#rpPrint',1],['#rpBack',2]]);
  await page.evaluate(() => { [...document.querySelectorAll('#report .rp-h2')].find(x => /細長效應/.test(x.textContent)).scrollIntoView({block:'start'}); window.scrollBy(0,-20); }); await W8(500);
  await shot('17-report-sl');
  await page.click('#rpBack'); await W8(500); await page.evaluate(() => window.scrollTo(0,0));
  /* 頂列 */
  { const t = await box('#topbar');
    await shot('18-topbar', [['#btnImport',1],['#btnCaseOpen',2],['#btnCaseSave',3],['#btnXlsx',4],['#btnPreview',5],['#btnPrint',6],['#btnTest',7]], {x:0, y:0, width:W, height:t.height+4}); }
  /* 12 匯出 Excel */
  const [dl] = await Promise.all([page.waitForEvent('download', {timeout:90000}), page.click('#btnXlsx')]);
  fs.rmSync(XL, {recursive:true, force:true}); fs.mkdirSync(XL, {recursive:true});
  await dl.saveAs(path.join(XL, 'C1.xlsx'));
  /* 13 範本覆寫確認 */
  await page.click('#btnSetup'); await W8(300);
  await page.selectOption('#preset', 'pond80'); await W8(400);
  await shot('20-preset-ask', [['#askAct',1]]);
  await page.click('#askAct button:has-text("取消")'); await W8(300);
  await page.click('#setupGo'); await W8(300);
  /* 14 從圖說匯入 */
  await page.click('#btnImport'); await W8(500);
  await shot('21-import', [['#impFile',1],['#impKey',2],['#impModel',3],['#impGo',4]]);
  await page.click('#impClose'); await W8(300);
  /* 16 自測 */
  await page.click('#btnTest'); await W8(900);
  await shot('22-test', [], await pane());
  /* 15 梁模組 */
  await page.click('#btnSetup'); await W8(300);
  await page.click('#tabBeam'); await W8(500);
  await page.fill('#bMemName', 'G1');
  await shot('23-beam-setup', [['#tabBeam',1],['#setupSecB',2],['#setupPreB',3]]);
  await page.click('#setupGo'); await W8(600);
  TABSEL = () => '#wsB';
  await ws('fig'); await W8(500);
  await shot('24-beam', [['#bSignPos',1],['#railB',2,2],['#wsB .ws-tabs',3]]);
  await ws('load'); await W8(500);
  { const a = await pane(); a.height = Math.min(520, a.height); await shot('25-beam-load', [['#bLoadTbl',1],['#bPaste',2]], a); }
  await ctx.close();
  if(errs.length) throw new Error('網頁執行錯誤：' + errs.join(' | '));
}

/* ---------- Excel 頁面圖：檢核表含細長效應之頁、A4 計算書第一頁 ---------- */
function excelImages(){
  execFileSync('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', XL, path.join(XL, 'C1.xlsx')], {stdio:'ignore', timeout:240000});
  const pdf = path.join(XL, 'C1.pdf');
  const n = +execFileSync('pdfinfo', [pdf]).toString().match(/Pages:\s+(\d+)/)[1];
  const text = i => execFileSync('pdftotext', ['-f', String(i), '-l', String(i), pdf, '-']).toString();
  let pCheck = 0, pA4 = 0;
  for(let i = 1; i <= n && !(pCheck && pA4); i++){
    const t = text(i);
    if(!pCheck && /檢核表/.test(t) && /側移支撐/.test(t)) pCheck = i;
    if(!pA4 && /RC 柱斷面設計檢核計算書/.test(t)) pA4 = i;
  }
  if(!pCheck || !pA4) throw new Error('Excel PDF 找不到檢核表或 A4 計算書頁');
  for(const [pg, name, dpi, keep] of [[pCheck, '26-xl-check', 110, 0.97], [pA4, '27-xl-a4', 220, 0.95]]){
    execFileSync('pdftoppm', ['-r', String(dpi), '-png', '-singlefile', '-f', String(pg), '-l', String(pg), pdf, path.join(IMG, name)]);
    // 裁掉頁尾與白邊
    execFileSync('python3', ['-c', `
from PIL import Image, ImageOps
f=${JSON.stringify(path.join(IMG, name + '.png'))}
im=Image.open(f).convert('RGB'); w,h=im.size; im=im.crop((0,0,w,int(h*${keep})))
bb=ImageOps.invert(im).getbbox(); im.crop((max(0,bb[0]-16),max(0,bb[1]-16),min(w,bb[2]+16),min(h,bb[3]+16))).save(f)`]);
    console.log('  Excel 頁面', name, '（第', pg, '頁）');
  }
}

/* ---------- 輸入檔雜湊（git blob） ---------- */
function stampText(){
  return INPUTS.map(f => execFileSync('git', ['hash-object', path.join(ROOT, f)]).toString().trim() + '  ' + f).join('\n') + '\n';
}

/* ---------- 教學 PDF ---------- */
async function tutorialPdf(browser, base){
  const page = await browser.newPage();
  await page.goto(base + '/docs/tutorial/tutorial.html');
  await page.waitForFunction(() => window.__ready === true, null, {timeout:30000});
  const err = await page.evaluate(() => document.querySelectorAll('.katex-error').length);
  if(err) throw new Error(`教學內有 ${err} 個 KaTeX 排版錯誤`);
  const bad = await page.evaluate(() => [...document.images].filter(i => !i.complete || !i.naturalWidth).map(i => i.getAttribute('src')));
  if(bad.length) throw new Error('教學內圖片載入失敗：' + bad.join(', '));
  const d = new Date(), p = n => String(n).padStart(2, '0');
  const idx = execFileSync('git', ['hash-object', path.join(ROOT, 'index.html')]).toString().trim().slice(0, 8);
  await page.evaluate(t => { document.getElementById('ver').textContent = t; },
    `建置日期：${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}　對應程式：index.html ${idx}`);
  await page.pdf({path: OUT, format:'A4', printBackground:true, displayHeaderFooter:true, headerTemplate:'<div></div>',
    footerTemplate:'<div style="width:100%;font-size:8px;color:#9ca3af;text-align:center;font-family:sans-serif">RC 斷面設計工具 操作教學　·　<span class="pageNumber"></span> / <span class="totalPages"></span></div>',
    margin:{top:'16mm', bottom:'16mm', left:'15mm', right:'15mm'}});
  await page.close();
}

(async () => {
  const srv = await serve(), base = `http://127.0.0.1:${srv.address().port}`;
  const browser = await chromium.launch(process.env.CHROMIUM_PATH ? {executablePath: process.env.CHROMIUM_PATH} : {});
  try{
    console.log('1/3 操作截圖'); await screenshots(browser, base);
    console.log('2/3 Excel 頁面'); excelImages();
    console.log('3/3 產生 PDF'); await tutorialPdf(browser, base);
    fs.writeFileSync(path.join(HERE, 'stamp.txt'), stampText());
    console.log('完成：', path.relative(ROOT, OUT));
  } finally { await browser.close(); srv.close(); }
})().catch(e => { console.error('建置失敗：', e.message); process.exit(1); });

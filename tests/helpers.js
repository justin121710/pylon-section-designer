/* 回歸測試共用：以本機 HTTP 伺服器開啟 index.html（CDN 改由 tests/node_modules 提供），回傳 Playwright page。
   與 docs/tutorial/build.js 相同作法；Chromium 可用環境變數 CHROMIUM_PATH 指定。 */
'use strict';
const { chromium } = require('playwright');
const http = require('http'), path = require('path'), fs = require('fs');

const ROOT = path.resolve(__dirname, '..'), NM = path.join(__dirname, 'node_modules');
const MIME = {'.html':'text/html; charset=utf-8', '.js':'text/javascript', '.css':'text/css', '.json':'application/json',
  '.woff2':'font/woff2', '.woff':'font/woff', '.ttf':'font/ttf', '.png':'image/png', '.svg':'image/svg+xml'};

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

async function openApp(){
  const srv = await serve();
  const browser = await chromium.launch(process.env.CHROMIUM_PATH ? {executablePath: process.env.CHROMIUM_PATH} : {});
  const ctx = await browser.newContext({viewport:{width:1440, height:900}, acceptDownloads:true});
  // CDN → 本機套件；其餘外部資源（Google Fonts 等）略過
  await ctx.route(/^https?:\/\/(?!127\.0\.0\.1)/, r => {
    const m = r.request().url().match(/cdn\.jsdelivr\.net\/npm\/([^@]+)@[^/]+\/(.*)$/);
    if(m && fs.existsSync(path.join(NM, m[1], m[2]))) return r.fulfill({path: path.join(NM, m[1], m[2])});
    return r.abort();
  });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  page.on('dialog', d => d.dismiss().catch(() => {}));
  await page.goto(`http://127.0.0.1:${srv.address().port}/index.html`);
  await page.waitForFunction(() => typeof MODEL !== 'undefined' && MODEL && typeof applyPreset === 'function');
  // 關閉開場精靈（不影響計算）
  await page.evaluate(() => { const b = document.getElementById('setupSkip'); if(b) b.click(); });
  return {page, errors, close: async () => { await browser.close(); srv.close(); }};
}

module.exports = {openApp, ROOT};

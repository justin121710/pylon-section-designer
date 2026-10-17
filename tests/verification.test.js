/* 驗證案例回歸測試（V&V）：tests/vv-cases.js 之全部案例（V1～V8），工具讀值須與手算（規範原式＋應變諧和試算）一致。
   同一組案例由 docs/tutorial/build.js 列印為《驗證與安全性流程》手冊之對照表，手冊數字即本測試之結果。 */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const {openApp} = require('./helpers');
const {CASES, runCase} = require('./vv-cases');
const katex = require('katex');

let app, page;
test.before(async () => { app = await openApp(); page = app.page; });
test.after(async () => { if(app){ assert.deepEqual(app.errors, [], '頁面不得有 JavaScript 錯誤'); await app.close(); } });

for(const C of CASES){
  test(`${C.id}：${C.title}`, async () => {
    const {rows} = await runCase(page, C);
    for(const r of rows){
      assert.equal(typeof r.tool, 'number', `${r.label}：工具讀值 ${r.tool}`);
      assert.ok(r.ok, `${r.label}：工具 ${r.tool/r.s}、手算 ${r.hand/r.s} ${r.unit}，相對差 ${(r.rel*100).toFixed(4)}% ＞ 容許 ${(r.tol*100).toFixed(2)}%`);
    }
    // 數值近似（曲線取樣、雙軸網格）之偏差方向：工具 D/C 不得比基準偏低（偏不保守）超過 0.1%
    for(const r of rows.filter(q => /^dc/.test(q.k) && q.k !== 'dcB'))
      assert.ok(r.tool >= r.hand*(1 - 1e-3), `${r.label}：工具 D/C ${r.tool} 比基準 ${r.hand} 偏低（不保守）`);
  });
}

test('驗證手冊之手算步驟皆可由 KaTeX 排版', () => {
  for(const C of CASES) for(const [label, tex] of C.hand().steps)
    assert.doesNotThrow(() => katex.renderToString('\\displaystyle ' + tex, {throwOnError: true, strict: 'ignore'}), `${C.id}「${label}」`);
});

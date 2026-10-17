/* 突變測試（mutation testing）：檢驗 tests/core.test.js 能否抓到計算核心 core/rc-core.js 之錯誤。
   作法：逐一對 rc-core.js 注入一個小錯誤（突變），以暫存目錄執行 core.test.js；測試失敗＝突變被「殺死」（測試有效），
   測試仍通過＝突變「存活」（驗證盲區，須補測試或確認為等價突變）。
   突變運算子（只作用於程式碼，略過註解與字串）：
     NUM  數值常數 × 1.1（0 → 1、整數 1 → 2）     CMP  < ↔ <=、> ↔ >=
     MM   Math.min ↔ Math.max                      LOG  && ↔ ||
   sha256（雜湊常數，另由對照 Node crypto 之測試把關）不突變。
   用法：node tests/mutation.js [--json 輸出檔] [--verbose] [--check]；不需安裝套件，約 1 分鐘。
   --check：有非等價突變存活時以失敗結束（CI 使用）。等價突變列於 EQUIV（附理由），不計入分母。
   結果寫入 stdout；--json 時另輸出 {total, killed, equivalent, score, survived, equiv}，供驗證手冊（docs/tutorial/verify.html）引用。 */
'use strict';
const fs = require('fs'), os = require('os'), path = require('path'), { spawnSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..');
const SRC = fs.readFileSync(path.join(ROOT, 'core', 'rc-core.js'), 'utf8');
const TEST = fs.readFileSync(path.join(__dirname, 'core.test.js'), 'utf8');

/* 掃描程式碼區段（略過 // 與 /* *\/ 註解、'…'、"…"、`…` 字串），回傳 [起, 迄) 區間 */
function codeSpans(src){
  const spans = []; let i = 0, start = 0;
  const push = e => { if(e > start) spans.push([start, e]); };
  while(i < src.length){
    const c = src[i], d = src[i+1];
    if(c === '/' && d === '/'){ push(i); i = src.indexOf('\n', i); if(i < 0) i = src.length; start = i; continue; }
    if(c === '/' && d === '*'){ push(i); i = src.indexOf('*/', i + 2) + 2; start = i; continue; }
    if(c === "'" || c === '"' || c === '`'){
      push(i); let j = i + 1;
      while(j < src.length && src[j] !== c){ if(src[j] === '\\') j++; j++; }
      i = j + 1; start = i; continue;
    }
    i++;
  }
  push(src.length);
  return spans;
}

/* 產生突變清單 {pos, len, to, op} */
function mutants(src){
  const out = [], skipA = src.indexOf('function sha256'), skipB = src.indexOf('\n}\n', skipA) + 2;
  for(const [a, b] of codeSpans(src)){
    const seg = src.slice(a, b);
    const scan = (re, fn) => { for(const m of seg.matchAll(re)){ const pos = a + m.index + (m[1] ? m[0].indexOf(m[1]) : 0), tok = m[1] ?? m[0];
      if(pos >= skipA && pos < skipB) continue; const r = fn(tok); if(r != null) out.push({pos, len: tok.length, from: tok, ...r}); } };
    // 數值常數（不含識別字中之數字、十六進位）
    scan(/(?<![\w.$])(\d+(?:\.\d+)?(?:e[+-]?\d+)?)(?![\w.])/gi, t => {
      const v = Number(t); if(!Number.isFinite(v)) return null;
      const to = v === 0 ? '1' : v === 1 ? '2' : String(+(v*1.1).toPrecision(6));
      return {to, op: 'NUM'};
    });
    scan(/<=|>=|<(?![<=])|>(?![>=])/g, t => {
      const i0 = 0; void i0;
      return {to: {'<': '<=', '<=': '<', '>': '>=', '>=': '>'}[t], op: 'CMP'};
    });
    scan(/Math\.(min|max)\b/g, t => ({to: t === 'min' ? 'max' : 'min', op: 'MM'}));
    scan(/&&|\|\|/g, t => ({to: t === '&&' ? '||' : '&&', op: 'LOG'}));
  }
  // 排除箭頭函式之 => 與位元移位 >>、>>>（上面的 regex 已避開 >>；=> 之 > 需另判）
  return out.filter(m => !(m.op === 'CMP' && (src[m.pos - 1] === '=' && m.from === '>'))).sort((p, q) => p.pos - q.pos);
}

const lineOf = (src, pos) => src.slice(0, pos).split('\n').length;

/* 等價突變（不改變任何可能之輸出）：逐項說明理由；存活者若符合下列條件即不計入分母。
   code 為該行程式碼之片段，op／from 須相符；新增時須寫明為何等價。 */
const EQUIV = [
  {op:'NUM', from:/^1e-\d+$/, why:'數值比較容差（1e-6、1e-9、1e-18）改為 1.1 倍：只影響與門檻相差小於容差之輸入'},
  {op:'CMP', code:'if(eps_t<=m.ety) return m.phic', why:'φ 於 ε_t = ε_ty 連續：兩分支皆得 φ_c'},
  {op:'CMP', code:'if(eps_t>=etTC(m)) return m.phit', why:'φ 於 ε_t = ε_ty + 0.003 連續：內插式亦得 φ_t'},
  {op:'CMP', code:'const seisKdb', why:'門檻已含 +1e-6 容差，< 與 <= 只差在恰等於 4200.000001 之輸入'},
  {op:'CMP', code:'if(Pn <= 0 || !(Pthr > 0)) return 0.9', from:'<=', why:'P_n = 0 時公式亦得 0.9'},
  {op:'CMP', code:'b.eps_t < m.ety', from:'<', why:'b 恰為 ε_ty 時下一段以 a = b 命中，內插結果相同'},
  {op:'NUM', code:'((a.eps_t - b.eps_t) || 1)', why:'跨越條件下 a.ε_t > b.ε_t，分母不會為 0'},
  {op:'CMP', code:'const ratio = (a, b)', why:'僅容差比較'},
  {op:'NUM', code:'Math.max(-1, Math.min(1, -a*b/mx))', from:'1', why:'|ab| ≤ max(a², b²)，夾限不會作用'},
  {op:'CMP', code:'const okAx = (l, L) => bridge ? l < L : l <= L + 1e-9', from:'<=', why:'建築門檻已含 +1e-9 容差'},
  {op:'CMP', code:'kLow:', from:'<', why:'門檻含 −1e-9 容差'},
  {op:'CMP', code:'kSwayLow:', from:'<=', why:'門檻含 +1e-9 容差'},
  {op:'CMP', code:'<= 1e-9', why:'與 1e-9 之比較為「視為零」：< 與 <= 只差在恰等於 1e-9'},
  {op:'CMP', code:'act(dy) && dy.Mc > 0', from:'>', why:'act() 成立時 P_u > 0，M_c = δ·max(M_2, M_2,min) ≥ P_u(1.5 + 0.03h) ≫ 1 kgf·cm，恆 > 0'},
  {op:'NUM', code:'act(dy) && dy.Mc > 0', from:'0', why:'act() 成立時 P_u > 0，M_c = δ·max(M_2, M_2,min) ≥ P_u(1.5 + 0.03h) ≫ 1 kgf·cm，恆 > 0'},
  {op:'CMP', code:'act(dx) && dx.Mc > 0', from:'>', why:'act() 成立時 P_u > 0，M_c = δ·max(M_2, M_2,min) ≥ P_u(1.5 + 0.03h) ≫ 1 kgf·cm，恆 > 0'},
  {op:'NUM', code:'act(dx) && dx.Mc > 0', from:'0', why:'act() 成立時 P_u > 0，M_c = δ·max(M_2, M_2,min) ≥ P_u(1.5 + 0.03h) ≫ 1 kgf·cm，恆 > 0'},
  {op:'CMP', code:"SL.treat!=='mag' || Pu <= 0", from:'<=', why:'P_u = 0 時 M_2,min = 0、P_u/P_c = 0，δ = max(1, C_m) = 1（C_m ≤ 1）'},
  {op:'CMP', code:'return r >= 1 ? Infinity', from:'>=', why:'r = 1 時 C_m/(1 − r) 亦為 ∞'},
  {op:'CMP', code:'(SL.Q < 1 ? 1/(1-SL.Q) : Infinity)', from:'<', why:'Q = 1 時 1/(1 − Q) 亦為 ∞（且輸入端擋 Q ≥ 1）'},
  {op:'CMP', code:'dsQ !== null && dsQ <= 1.5 ?', from:'<=', why:'δ_s = 1.5 恰為邊界：兩法於此相等之情形不影響判定方向（Q 法 1.5 仍可用）'},
  {op:'CMP', code:'qOver: dsQ !== null && dsQ > 1.5', from:'>', why:'δ_s,Q 恰為 1.5 之邊界；qOver 旗標僅供說明文字'},
  {op:'CMP', code:'minGov: M2min > Mu', from:'>', why:'M_2,min = M_u 時兩者相同，旗標僅供說明文字'},
  {op:'CMP', code:'const m = c.reduce((a,b)=> b.v<a.v?b:a, c[0])', why:'相等候選之取捨：b_e 值相同，僅控制項名稱不同'},
  {op:'NUM', code:'const m = c.reduce((a,b)=> b.v<a.v?b:a, c[0])', why:'reduce 起始值改為 c[1]：仍走訪全部候選，最小值相同'},
  {op:'CMP', code:'hfOk: S.hf >= 0.5*S.bw - 1e-9', why:'門檻含 −1e-9 容差'},
  {op:'LOG', code:"if(typeof module !== 'undefined' && module.exports)", why:'僅決定是否匯出（瀏覽器與 Node 皆不影響計算）'}
];
const isEquiv = (m, code) => EQUIV.find(e => e.op === m.op && (!e.code || code.includes(e.code))
  && (!e.from || (e.from instanceof RegExp ? e.from.test(m.from) : e.from === m.from)));

function run(){
  const jsonOut = process.argv.includes('--json') ? process.argv[process.argv.indexOf('--json') + 1] : null;
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'rcsd-mut-'));
  fs.mkdirSync(path.join(tmp, 'core')); fs.mkdirSync(path.join(tmp, 'tests'));
  fs.writeFileSync(path.join(tmp, 'tests', 'core.test.js'), TEST);
  const exec = () => spawnSync(process.execPath, ['--test', path.join(tmp, 'tests', 'core.test.js')], {encoding: 'utf8', timeout: 30000});
  // 未突變之原始碼必須通過
  fs.writeFileSync(path.join(tmp, 'core', 'rc-core.js'), SRC);
  if(exec().status !== 0){ console.error('原始 core.test.js 未通過，無法進行突變測試'); process.exit(2); }
  const list = mutants(SRC), survived = [], equiv = []; let killed = 0;
  const lines = SRC.split('\n');
  for(const [k, m] of list.entries()){
    const mut = SRC.slice(0, m.pos) + m.to + SRC.slice(m.pos + m.len);
    fs.writeFileSync(path.join(tmp, 'core', 'rc-core.js'), mut);
    const r = exec();
    if(r.status !== 0) killed++;
    else { const ln = lineOf(SRC, m.pos), code = lines[ln - 1].trim(), eq = isEquiv(m, code);
      (eq ? equiv : survived).push({line: ln, op: m.op, from: m.from, to: m.to, code: code.slice(0, 160), ...(eq ? {why: eq.why} : {})}); }
    if((k + 1) % 50 === 0) process.stderr.write(`  ${k + 1}/${list.length}\n`);
  }
  fs.rmSync(tmp, {recursive: true, force: true});
  const total = list.length, eff = total - equiv.length, score = eff ? killed/eff : 1;
  const byOp = {}; for(const m of list) byOp[m.op] = (byOp[m.op] || 0) + 1;
  console.log(`突變總數 ${total}（${Object.entries(byOp).map(([k, v]) => `${k} ${v}`).join('、')}），殺死 ${killed}，等價 ${equiv.length}，存活 ${survived.length}，突變分數 ${(score*100).toFixed(1)}%（殺死／非等價）`);
  for(const s of survived) console.log(`  存活 L${s.line} [${s.op}] ${s.from} → ${s.to}：${s.code}`);
  if(process.argv.includes('--verbose')) for(const s of equiv) console.log(`  等價 L${s.line} [${s.op}] ${s.from} → ${s.to}：${s.why}`);
  const out = {total, byOp, killed, equivalent: equiv.length, score, survived, equiv};
  if(jsonOut) fs.writeFileSync(jsonOut, JSON.stringify(out, null, 1));
  if(process.argv.includes('--check') && survived.length){ console.error(`✗ 有 ${survived.length} 個非等價突變存活：請補 core.test.js，或確認等價後加入 EQUIV 並寫明理由`); process.exitCode = 1; }
  return out;
}

if(require.main === module) run();
module.exports = {mutants, codeSpans, run};

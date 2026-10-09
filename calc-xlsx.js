/* =========================================================================
   calc-xlsx.js — 匯出「附公式」Excel 計算書（檢核表 ＋ 結構計算書(A4)）
   -------------------------------------------------------------------------
   · 只在使用者按「匯出 Excel」時才載入（連同 ExcelJS），不影響主頁效能。
   · 所有導出數值皆為 Excel 公式；輸入值取自網頁目前的參數與載重。
   · 格式依「結構設計檢核表與 A4 計算書（中華民國規範）v2」：
       檢核表 A~G 七欄（項目｜符號｜數值｜單位｜計算式｜判定標準／備註｜規範依據）、
       輸入格淺藍底＋藍框、下拉格黃底＋金框、公式格白底細灰框、總判定淺琥珀底。
   · 同一份程式可在 Node 執行（module.exports），供以本機 Excel 重算比對。
   ========================================================================= */
(function(root){
'use strict';

/* ---------- 色碼（v2 唯一標準） ---------- */
const K = {
  inFill:'FFDDEBF7', inBorder:'FF2E75B6', inFont:'FF0000FF',
  listFill:'FFFFFF00', listBorder:'FFBF8F00',
  grid:'FF808080', hdr:'FFD9E1F2', sec:'FFF2F2F2', amber:'FFFFF2CC', link:'FF008000', black:'FF000000'
};
const FONT = 'Arial';
const BIG = 1e9;            // 「無限大／不適用」的數值代表（顯示格式會轉成 ∞ 或 —）
const FMT_INF = '[>=1E+9]"∞";0.000';
const FMT_SINF = '[>=1E+9]"—";0.00';

/* 鋼筋規格（CNS 560）— 與主頁 BAR 相同 */
const BARS = [['#3',0.953,0.7133],['#4',1.270,1.267],['#5',1.588,1.979],['#6',1.910,2.865],['#7',2.222,3.871],
              ['#8',2.540,5.067],['#9',2.865,6.469],['#10',3.226,8.143],['#11',3.581,10.07],['#14',4.300,14.52],['#18',5.733,25.79]];
const PRACTICAL_S = [30,25,20,15,12.5,10,7.5];

const NMAX = 80;            // 每邊主筋／圓周主筋最多根數（座標表列數）
const NPM  = 120;           // P-M 掃描點數（中性軸深度 c 對數取樣）
const NCB  = 12;            // 載重組合列數（可在 Excel 內增填）
const NIT  = 40;            // 雙軸載重輪廓法二分迭代次數

/* ======================================================================
   版面引擎：先收集列，再配置列號，最後解析 {key} 佔位並寫入
   ====================================================================== */
class CalcSheet {
  constructor(name){ this.name=name; this.rows=[]; this.keys={}; this.ranges={}; }
  add(r){ this.rows.push(r); return r; }
  title(t){ return this.add({t:'title',text:t}); }
  header(){ return this.add({t:'header'}); }
  section(t){ this.add({t:'blank'}); return this.add({t:'section',text:t}); }
  blank(){ return this.add({t:'blank'}); }
  /* item：key、label(A)、sym(B)、v 數值或 f 公式(C)、unit(D)、expr(E)、crit(F)、ref(G)、
           kind 'in'|'list'|'calc'|'link'、list（下拉選項）、note（備註）、fmt（數字格式） */
  item(o){ return this.add(Object.assign({t:'item'},o)); }
  /* 彙總列：A 項目、B 需求、C 容量、D 比值、E 判定 */
  sum(o){ return this.add(Object.assign({t:'sum'},o)); }
  table(o){ return this.add(Object.assign({t:'table'},o)); }   // 對照表：多列多欄
  text(t,o){ return this.add(Object.assign({t:'text',text:t},o||{})); }
  layout(){
    let r=0;
    for(const row of this.rows){
      if(row.t==='table'){
        row.r0 = r+1; r += 1 + row.data.length;          // 表頭一列 ＋ 資料列
        row.cols.forEach((c,j)=>{ if(c.key) this.ranges[c.key] = {sheet:this.name, c:String.fromCharCode(65+j), r1:row.r0+1, r2:row.r0+row.data.length}; });
      }else{ r++; row.r=r; if(row.key) this.keys[row.key]=r; }
    }
  }
}

/* 解析公式中的 {key} 與 {rng:key}；from 為公式所在工作表名稱 */
function makeResolver(sheets){
  const find = k => { for(const s of sheets) if(s.keys && k in s.keys) return {s:s.name, r:s.keys[k]}; return null; };
  const findR = k => { for(const s of sheets) if(s.ranges && k in s.ranges) return s.ranges[k]; return null; };
  const q = n => /^[A-Za-z0-9_]+$/.test(n) ? n : `'${n}'`;
  return (f, from) => f.replace(/\{(rng:)?([A-Za-z0-9_]+)\}/g, (m, isR, k) => {
    if(isR){ const R=findR(k); if(!R) throw new Error('未定義範圍 '+k);
      const a=`$${R.c}$${R.r1}:$${R.c}$${R.r2}`; return R.sheet===from ? a : `${q(R.sheet)}!${a}`; }
    const c=find(k); if(!c) throw new Error('未定義鍵 '+k);
    return c.s===from ? `$C$${c.r}` : `${q(c.s)}!$C$${c.r}`;
  });
}

/* ---------- 樣式工具 ---------- */
const side = (rgb, style='thin') => ({style, color:{argb:rgb}});
const box  = (rgb, style='thin') => ({top:side(rgb,style), left:side(rgb,style), bottom:side(rgb,style), right:side(rgb,style)});
function styleInput(cell, list){
  cell.fill = {type:'pattern', pattern:'solid', fgColor:{argb: list?K.listFill:K.inFill}};
  cell.font = {name:FONT, size:10, bold:true, color:{argb:K.inFont}};
  cell.border = box(list?K.listBorder:K.inBorder, 'medium');
}
function styleCalc(cell, link){
  cell.font = {name:FONT, size:10, color:{argb: link?K.link:K.black}};
  cell.border = box(K.grid);
}
const plain = (cell, o={}) => { cell.font = Object.assign({name:FONT, size:10}, o.font||{}); cell.border = box(K.grid);
  cell.alignment = Object.assign({vertical:'middle', wrapText:true}, o.align||{}); if(o.fill) cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:o.fill}}; };

/* 把 CalcSheet 寫成 ExcelJS 工作表（A~G 七欄規格） */
function writeCalcSheet(ws, cs, resolve){
  ws.columns = [230,70,80,60,260,260,300].map(pt=>({width: pt/5.6}));
  ws.pageSetup = {paperSize:8, orientation:'landscape', fitToPage:true, fitToWidth:1, fitToHeight:0,
    margins:{left:0.4,right:0.4,top:0.5,bottom:0.5,header:0.25,footer:0.25}, printTitlesRow:'1:2'};
  ws.headerFooter = {oddFooter:'&L&D&R第 &P 頁 / 共 &N 頁'};
  for(const row of cs.rows){
    if(row.t==='table'){ writeTable(ws, row, resolve, cs.name); continue; }
    const R = ws.getRow(row.r);
    if(row.t==='title'){
      ws.mergeCells(row.r,1,row.r,7); const c=R.getCell(1); c.value=row.text;
      c.font={name:FONT,size:14,bold:true}; c.alignment={vertical:'middle'}; R.height=24; continue;
    }
    if(row.t==='header'){
      ['項目','符號','數值','單位','計算式／來源說明','判定標準／備註','規範依據（條文出處）'].forEach((h,j)=>{
        const c=R.getCell(j+1); c.value=h; plain(c,{font:{bold:true},align:{horizontal:'center'},fill:K.hdr}); });
      continue;
    }
    if(row.t==='section'){
      ws.mergeCells(row.r,1,row.r,7);
      const c=R.getCell(1); c.value=row.text; plain(c,{font:{bold:true},fill:K.sec}); continue;
    }
    if(row.t==='blank') continue;
    if(row.t==='text'){
      ws.mergeCells(row.r,1,row.r,7); const c=R.getCell(1); c.value=row.text;
      plain(c,{font:row.bold?{bold:true}:{} , align:{vertical:'top'}}); R.height=row.h||30; continue;
    }
    if(row.t==='sum'){
      const cells=[row.label,row.need,row.cap,row.ratio,row.judge];
      cells.forEach((v,j)=>{
        const c=R.getCell(j+1);
        if(v && typeof v==='object' && v.f) c.value={formula: resolve(j===3?`IFERROR(${v.f},0)`:v.f, cs.name)}; else if(v!==undefined) c.value=v;
        plain(c,{font:row.head||row.total?{bold:true}:{}, align:{horizontal:j?'center':'left'},
                 fill: row.head?K.hdr : row.total?K.amber : undefined});
        if(j>0 && j<4 && !row.head) c.numFmt = j===3?FMT_INF:(row.fmt||FMT_SINF);
      });
      if(row.note){ const c=R.getCell(6); c.value=(typeof row.note==='object')?{formula:resolve(row.note.f,cs.name)}:row.note; plain(c,{fill:row.total?K.amber:undefined}); }
      if(row.ref){ const c=R.getCell(7); c.value=row.ref; plain(c,{fill:row.total?K.amber:undefined}); }
      if(row.head){ R.getCell(6).value='說明'; plain(R.getCell(6),{font:{bold:true},fill:K.hdr}); R.getCell(7).value='規範依據'; plain(R.getCell(7),{font:{bold:true},fill:K.hdr}); }
      continue;
    }
    // item
    const vals=[row.label,row.sym||'',null,row.unit||'',row.expr||'',row.crit||'',row.ref||''];
    vals.forEach((v,j)=>{ if(j!==2){ const c=R.getCell(j+1); c.value=v; plain(c,{align:{horizontal:j===1||j===3?'center':'left'}}); }});
    const c=R.getCell(3);
    if(row.f!==undefined) c.value={formula: resolve(row.f, cs.name)};
    else c.value=row.v;
    c.alignment={horizontal:'center', vertical:'middle', shrinkToFit:true};   // 文字值過長時縮小字級而非裁切
    if(row.kind==='in' || row.kind==='list') styleInput(c, row.kind==='list'); else styleCalc(c, row.kind==='link');
    if(row.fmt) c.numFmt=row.fmt;
    if(row.list){
      const L = Array.isArray(row.list) ? `"${row.list.join(',')}"` : resolve(row.list, cs.name);
      c.dataValidation={type:'list', allowBlank:true, formulae:[L], showErrorMessage:false};
    }
    if(row.note) c.note={texts:[{text:row.note}]};
  }
}
function writeTable(ws, row, resolve, from){
  const R0=ws.getRow(row.r0);
  row.cols.forEach((col,j)=>{ const c=R0.getCell(j+1); c.value=col.h; plain(c,{font:{bold:true},align:{horizontal:'center'},fill:K.hdr}); });
  row.data.forEach((d,i)=>{
    const R=ws.getRow(row.r0+1+i);
    d.forEach((v,j)=>{
      const c=R.getCell(j+1);
      if(v && typeof v==='object' && v.f) c.value={formula: resolve(v.f, from)}; else c.value=v;
      if(row.cols[j].input){ styleInput(c,false); } else plain(c,{align:{horizontal:'center'}});
      if(row.cols[j].fmt) c.numFmt=row.cols[j].fmt;
    });
  });
}

/* 一般儲存格寫入（輔助工作表用） */
function put(ws, addr, v, o={}){
  const c=ws.getCell(addr);
  if(v && typeof v==='object' && v.f!==undefined) c.value={formula:v.f}; else c.value=v;
  if(o.input) styleInput(c, !!o.list); else if(o.head) plain(c,{font:{bold:true},align:{horizontal:'center'},fill:K.hdr});
  else if(o.sec) plain(c,{font:{bold:true},fill:K.sec}); else styleCalc(c, !!o.link);
  if(o.fmt) c.numFmt=o.fmt;
  if(o.list) c.dataValidation={type:'list', allowBlank:true, formulae:[o.list], showErrorMessage:false};
  if(o.note) c.note={texts:[{text:o.note}]};
  if(!o.input && !o.head && !o.sec) c.alignment={horizontal:'center', vertical:'middle'};
  return c;
}
const colL = n => { let s=''; n++; while(n>0){ const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26); } return s; }; // 0→A

/* ======================================================================
   柱斷面活頁簿
   inp：主頁目前的參數（kgf-cm；載重 tf、tf·m）
   ====================================================================== */
function buildColumn(ExcelJS, inp){
  const wb = new ExcelJS.Workbook();
  wb.creator = 'RC 斷面設計工具'; wb.created = new Date();
  wb.calcProperties = {fullCalcOnLoad:true};
  const S = new CalcSheet('檢核表');
  const ref401 = s => `土木401 §${s}（ACI 318 同條號）`;

  S.title(`RC 柱斷面設計檢核表（矩形／中空箱型／圓形，kgf-cm 制）`);
  S.header();

  /* ---------------- 一、設計依據與斷面 ---------------- */
  S.section('【一、設計依據與斷面幾何】');
  S.item({key:'code', label:'設計依據', v:inp.code==='bridge'?'橋梁':'建築物', kind:'list', list:['建築物','橋梁'],
    expr:'下拉選擇', crit:'建築物：柱端加密區依特殊抗彎構材；橋梁：塑鉸區依超強係數 φo 容量設計',
    ref:'建築物耐震設計規範及解說；公路橋梁耐震設計規範；斷面強度一律依土木401',
    note:'建築物＝土木401＋建築物耐震設計規範（柱端加密區 Ve = 2Mpr/lu）。\n橋梁＝土木401 斷面強度＋公路橋梁耐震設計規範（Ve = φo·Mn/Lv）。'});
  S.item({key:'isBldg', label:'　旗標：建築物', f:'--({code}="建築物")', expr:'=1 表建築物', crit:'內部判別用', ref:'非規範明列條文，係本表判別用'});
  S.item({key:'type', label:'斷面型式', v:{solid:'實心矩形',box:'中空箱型',circle:'圓形'}[inp.type], kind:'list', list:['實心矩形','中空箱型','圓形'],
    expr:'下拉選擇', crit:'中空箱型以受壓塊∩中空區交集扣除；圓形以弓形面積積分', ref:'土木401 §22.2（撓曲設計假設）',
    note:'實心矩形、中空箱型、圓形三種。圓形時 B、H、tw、內側保護層與每面根數不使用。'});
  S.item({key:'isBox', label:'　旗標：中空箱型', f:'--({type}="中空箱型")', ref:'非規範明列條文，係本表判別用'});
  S.item({key:'isCirc', label:'　旗標：圓形', f:'--({type}="圓形")', ref:'非規範明列條文，係本表判別用'});
  S.item({key:'Bin', label:'外寬 B（cm）', sym:'B', v:inp.B, unit:'cm', kind:'in', crit:'矩形／箱型使用', ref:'—（幾何輸入）', fmt:'0.0'});
  S.item({key:'Hin', label:'外高 H（cm）', sym:'H', v:inp.H, unit:'cm', kind:'in', crit:'矩形／箱型使用', ref:'—（幾何輸入）', fmt:'0.0'});
  S.item({key:'Din', label:'直徑 D（cm）', sym:'D', v:inp.D, unit:'cm', kind:'in', crit:'圓形使用', ref:'—（幾何輸入）', fmt:'0.0'});
  S.item({key:'twin', label:'壁厚 tw（cm）', sym:'tw', v:inp.tw, unit:'cm', kind:'in', crit:'中空箱型使用', ref:'—（幾何輸入）', fmt:'0.0'});
  S.item({key:'Be', label:'計算用 B', sym:'B', f:'IF({isCirc}=1,{Din},{Bin})', unit:'cm', expr:'圓形取 D', ref:'—', fmt:'0.0'});
  S.item({key:'He', label:'計算用 H', sym:'H', f:'IF({isCirc}=1,{Din},{Hin})', unit:'cm', expr:'圓形取 D', ref:'—', fmt:'0.0'});
  S.item({key:'tw', label:'計算用 tw', sym:'tw', f:'IF({twin}>=MIN({Be},{He})/2,MIN({Be},{He})/2-1,{twin})', unit:'cm',
    expr:'tw ≧ min(B,H)/2 時夾為 min(B,H)/2 − 1', crit:'防止中空區消失', ref:'非規範明列條文，係幾何防呆', fmt:'0.0'});
  S.item({key:'covO', label:'外側保護層 co（量至橫向筋外緣）', sym:'co', v:inp.covO, unit:'cm', kind:'in',
    crit:'水工／環境結構建議 ≧ 5 cm', ref:'土木401 §20.5.1.3（保護層）', fmt:'0.0'});
  S.item({key:'covI', label:'內側保護層 ci（箱型）', sym:'ci', v:inp.covI, unit:'cm', kind:'in', crit:'中空箱型使用', ref:'土木401 §20.5.1.3', fmt:'0.0'});
  S.item({key:'dagg', label:'骨材最大粒徑', sym:'dagg', v:inp.dagg, unit:'cm', kind:'in', crit:'主筋淨距下限之一', ref:'土木401 §25.2.3', fmt:'0.0'});
  S.item({key:'Ag', label:'全斷面積', sym:'Ag', unit:'cm²', fmt:'#,##0',
    f:'IF({isCirc}=1,PI()*{Din}^2/4,IF({isBox}=1,{Be}*{He}-MAX(0,{Be}-2*{tw})*MAX(0,{He}-2*{tw}),{Be}*{He}))',
    expr:'矩形 B·H；箱型 B·H − (B−2tw)(H−2tw)；圓形 πD²/4', ref:'—（幾何）'});

  /* ---------------- 二、材料與強度折減因數 ---------------- */
  S.section('【二、材料與強度折減因數】');
  S.item({key:'fc', label:"混凝土抗壓強度 f'c", sym:"f'c", v:inp.fc, unit:'kgf/cm²', kind:'in', fmt:'#,##0', ref:ref401('19.2'), crit:'常用 280、350、420'});
  S.item({key:'fy', label:'主筋降伏強度 fy', sym:'fy', v:inp.fy, unit:'kgf/cm²', kind:'list', list:['2800','4200','5000'], fmt:'#,##0',
    ref:'CNS 560；'+ref401('20.2'), note:'SD280＝2800、SD420＝4200、SD490＝5000 kgf/cm²（CNS 560）。'});
  S.item({key:'fyt', label:'橫向筋降伏強度 fyt', sym:'fyt', v:inp.fyt, unit:'kgf/cm²', kind:'list', list:['2800','4200','5000'], fmt:'#,##0',
    ref:'CNS 560；'+ref401('20.2'), note:'SD280＝2800、SD420＝4200、SD490＝5000 kgf/cm²（CNS 560）。'});
  S.item({key:'Es', label:'鋼筋彈性模數 Es', sym:'Es', v:inp.Es, unit:'kgf/cm²', kind:'in', fmt:'#,##0', ref:ref401('20.2.2.2')});
  S.item({key:'ecu', label:'混凝土極限壓應變 εcu', sym:'εcu', v:inp.ecu, unit:'無因次', kind:'in', fmt:'0.0000', ref:ref401('22.2.2.1')});
  S.item({key:'b1', label:'等值應力塊係數 β₁', sym:'β₁', f:'MAX(0.65,MIN(0.85,0.85-0.05*({fc}-280)/70))', unit:'無因次', fmt:'0.000',
    expr:"β₁ = 0.85 − 0.05(f'c − 280)/70，介於 0.65～0.85", ref:ref401('22.2.2.4.3')});
  S.item({key:'Ec', label:'混凝土彈性模數 Ec', sym:'Ec', f:'15000*SQRT({fc})', unit:'kgf/cm²', fmt:'#,##0', expr:"Ec = 15,000√f'c", ref:ref401('19.2.2.1')});
  S.item({key:'ety', label:'主筋降伏應變 εty', sym:'εty', f:'{fy}/{Es}', unit:'無因次', fmt:'0.00000', expr:'εty = fy / Es', ref:ref401('21.2.2')});
  S.item({key:'phic', label:'壓力控制強度折減因數 φc', sym:'φc', v:inp.phic, unit:'無因次', kind:'in', fmt:'0.00',
    crit:'關鍵假設：橫箍（含圓箍）0.65、螺箍 0.75', ref:ref401('21.2.2'), note:'關鍵假設。橫箍柱（含圓形箍筋）0.65；螺箍柱 0.75（土木401 §21.2.2）。'});
  S.item({key:'phit', label:'拉力控制強度折減因數 φt', sym:'φt', v:inp.phit, unit:'無因次', kind:'in', fmt:'0.00', ref:ref401('21.2.2'),
    crit:'過渡區線性內插；本表拉控門檻取 εt ≧ 0.005（較 εty + 0.003 保守）'});
  S.item({key:'phiv', label:'剪力／扭矩強度折減因數 φv', sym:'φv', v:inp.phiv, unit:'無因次', kind:'in', fmt:'0.00', ref:ref401('21.2.1'),
    crit:"關鍵假設：0.75 搭配 Vc = 0.53√f'c·bw·d，不可與 AASHTO 0.90 混用", note:'關鍵假設。φv = 0.75 係與 ACI 型 Vc 公式配套校準。'});
  S.item({key:'pmaxf', label:'最大軸力截斷係數', sym:'—', v:inp.pmaxf, unit:'×Po', kind:'in', fmt:'0.00', ref:ref401('22.4.2.1'),
    crit:'橫箍 0.80、螺箍 0.85', note:'橫箍柱 Pn,max = 0.80Po；螺箍柱 0.85Po（土木401 §22.4.2.1）。'});
  S.item({key:'alpha', label:'雙軸載重輪廓指數 α', sym:'α', v:inp.alpha, unit:'無因次', kind:'in', fmt:'0.00',
    crit:"Pu < 0.1f'c·Ag 時採載重輪廓法", ref:'非規範明列條文，係 PCA 載重輪廓法（Bresler 載重輪廓）慣用作法；α = 1.0 為保守值'});

  /* ---------------- 三、主筋與橫向鋼筋 ---------------- */
  S.section('【三、主筋與橫向鋼筋配置】');
  S.item({key:'bar', label:'主筋號數', v:inp.barSize, kind:'list', list:'{rng:barName}', expr:'下拉選擇（對照表區）', ref:'CNS 560',
    note:'#4～#18（D13～D57），直徑與面積查本表對照表區（CNS 560）。'});
  S.item({key:'db', label:'主筋直徑', sym:'db', f:'INDEX({rng:barD},MATCH({bar},{rng:barName},0))', unit:'cm', fmt:'0.000', expr:'查對照表', ref:'CNS 560'});
  S.item({key:'Ab', label:'主筋單根面積', sym:'Ab', f:'INDEX({rng:barA},MATCH({bar},{rng:barName},0))', unit:'cm²', fmt:'0.000', expr:'查對照表', ref:'CNS 560'});
  S.item({key:'tie', label:'箍筋／繫筋／螺箍號數', v:inp.tieSize, kind:'list', list:'{rng:barName}', ref:'CNS 560', note:'常用 #3～#6。'});
  S.item({key:'dt', label:'橫向筋直徑', sym:'dt', f:'INDEX({rng:barD},MATCH({tie},{rng:barName},0))', unit:'cm', fmt:'0.000', ref:'CNS 560'});
  S.item({key:'At', label:'橫向筋單肢面積', sym:'At', f:'INDEX({rng:barA},MATCH({tie},{rng:barName},0))', unit:'cm²', fmt:'0.000', ref:'CNS 560'});
  S.item({key:'nBin', label:'B 向每面根數（含角隅）', sym:'nB', v:inp.nB, unit:'根', kind:'in', crit:'矩形／箱型使用', ref:'—（配筋輸入）'});
  S.item({key:'nHin', label:'H 向每面根數（含角隅）', sym:'nH', v:inp.nH, unit:'根', kind:'in', crit:'矩形／箱型使用', ref:'—（配筋輸入）'});
  S.item({key:'nB', label:'計算用 nB', sym:'nB', f:'MAX(2,ROUND({nBin},0))', unit:'根', ref:'—'});
  S.item({key:'nH', label:'計算用 nH', sym:'nH', f:'MAX(2,ROUND({nHin},0))', unit:'根', ref:'—'});
  S.item({key:'nCin', label:'圓形主筋總根數', sym:'n', v:inp.nCir, unit:'根', kind:'in', crit:'圓形使用；螺箍柱 ≧ 6', ref:ref401('10.7.3.1')});
  S.item({key:'nC', label:'計算用 n（圓形）', sym:'n', f:'MAX(4,ROUND({nCin},0))', unit:'根', ref:'—'});
  S.item({key:'dbl', label:'箱型內層同配置', v:inp.dbl?'是':'否', kind:'list', list:['是','否'], crit:'中空箱型使用；內層座標沿用外層網格',
    ref:'非規範明列條文，係本工具配筋模式', note:'是：內外雙層；否：僅外層。'});
  S.item({key:'spiral', label:'圓柱橫向筋型式', v:inp.spiral?'螺旋箍筋':'圓形箍筋', kind:'list', list:['螺旋箍筋','圓形箍筋'],
    crit:'螺箍 φc = 0.75、截斷 0.85；圓箍同橫箍柱', ref:ref401('25.7.3'), note:'螺旋箍筋須連續且滿足體積比 ρs；圓形箍筋為單圈搭接（135° 彎鉤）。'});
  S.item({key:'isSp', label:'　旗標：螺箍', f:'--(AND({isCirc}=1,{spiral}="螺旋箍筋"))', ref:'非規範明列條文，係本表判別用'});
  S.item({key:'tieAll', label:'繫筋配置', v:inp.tieAll?'每一根主筋皆設':'規範最低（每隔一根）', kind:'list', list:['規範最低（每隔一根）','每一根主筋皆設'],
    crit:'淨距 > 15 cm 時自動改為每根支撐', ref:ref401('25.7.2.3'), note:'規範：每隔一根縱筋須側撐，且未支撐筋距被支撐筋淨距 ≦ 15 cm。'});
  S.item({key:'dOut', label:'外層主筋中心至外緣', sym:'d′', f:'{covO}+{dt}+{db}/2', unit:'cm', fmt:'0.00', expr:'co + dt + db/2', ref:'—（幾何）'});
  S.item({key:'dIn', label:'內層主筋中心至內緣', sym:'d′i', f:'{covI}+{dt}+{db}/2', unit:'cm', fmt:'0.00', expr:'ci + dt + db/2', ref:'—（幾何）'});
  S.item({key:'hxO', label:'外層主筋 x 半寬', sym:'hxo', f:'{Be}/2-{dOut}', unit:'cm', fmt:'0.00', ref:'—（幾何）'});
  S.item({key:'hyO', label:'外層主筋 y 半高', sym:'hyo', f:'{He}/2-{dOut}', unit:'cm', fmt:'0.00', ref:'—（幾何）'});
  S.item({key:'pB', label:'B 邊主筋中心距', sym:'pB', f:'2*{hxO}/({nB}-1)', unit:'cm', fmt:'0.00', ref:'—（幾何）'});
  S.item({key:'pH', label:'H 邊主筋中心距', sym:'pH', f:'2*{hyO}/({nH}-1)', unit:'cm', fmt:'0.00', ref:'—（幾何）'});
  S.item({key:'rb', label:'圓形主筋中心半徑', sym:'rb', f:'{Din}/2-{dOut}', unit:'cm', fmt:'0.00', ref:'—（幾何）'});
  S.item({key:'hxI', label:'內層主筋 x 半寬', sym:'hxi', f:'{Be}/2-{tw}+{dIn}', unit:'cm', fmt:'0.00', ref:'—（幾何）'});
  S.item({key:'hyI', label:'內層主筋 y 半高', sym:'hyi', f:'{He}/2-{tw}+{dIn}', unit:'cm', fmt:'0.00', ref:'—（幾何）'});
  S.item({key:'inOK', label:'內層配筋有效', f:'--AND({isBox}=1,{dbl}="是",{hxI}>0,{hyI}>0,{hxI}<{hxO}-{db},{hyI}<{hyO}-{db})',
    expr:'箱型、選雙層且壁厚容得下', crit:'=0 時僅外層', ref:'非規範明列條文，係幾何判別'});
  S.item({key:'cntX', label:'內層：落在內緣內之外層 x 座標數', f:"SUM('配筋座標'!F3:F"+(2+NMAX)+')', unit:'根', expr:'詳「配筋座標」F 欄', ref:'—'});
  S.item({key:'cntY', label:'內層：落在內緣內之外層 y 座標數', f:"SUM('配筋座標'!G3:G"+(2+NMAX)+')', unit:'根', expr:'詳「配筋座標」G 欄', ref:'—'});
  S.item({key:'nBars', label:'主筋總根數', sym:'n', unit:'根',
    f:'IF({isCirc}=1,{nC},2*{nB}+2*({nH}-2)+{inOK}*(2*({cntX}+2)+2*({cntY})))',
    expr:'外層 2nB + 2(nH−2)；內層 2(nx,i) + 2(ny,i − 2)；圓形 n', ref:'—（配筋）'});
  S.item({key:'Ast', label:'主筋總面積', sym:'Ast', f:'{nBars}*{Ab}', unit:'cm²', fmt:'0.00', ref:'—'});
  S.item({key:'rho', label:'主筋比', sym:'ρg', f:'{Ast}/{Ag}*100', unit:'%', fmt:'0.000', expr:'Ast / Ag', crit:'規範 1%～8%；本表以 4% 為施工性上限警示', ref:ref401('10.6.1.1')});
  S.item({key:'need', label:'主筋淨距需求', sym:'smin', f:'MAX(4,1.5*{db},4/3*{dagg})', unit:'cm', fmt:'0.00', expr:'max(4.0, 1.5db, 4/3·dagg)', ref:ref401('25.2.3')});
  S.item({key:'clB', label:'主筋淨距（B 邊／圓周）', sym:'s', f:'IF({isCirc}=1,2*PI()*{rb}/{nC}-{db},{pB}-{db})', unit:'cm', fmt:'0.00',
    expr:'矩形 pB − db；圓形 2πrb/n − db', ref:'—'});
  S.item({key:'clH', label:'主筋淨距（H 邊）', sym:'s', f:'IF({isCirc}=1,{clB},{pH}-{db})', unit:'cm', fmt:'0.00', expr:'pH − db', ref:'—'});
  S.item({key:'everyB', label:'B 邊每根皆須側撐', f:'--OR({tieAll}="每一根主筋皆設",{clB}>15)', expr:'淨距 > 15 cm 或選「每一根」', crit:'1＝每根', ref:ref401('25.7.2.3')});
  S.item({key:'everyH', label:'H 邊每根皆須側撐', f:'--OR({tieAll}="每一根主筋皆設",{clH}>15)', ref:ref401('25.7.2.3')});
  S.item({key:'nTieX', label:'B 邊繫筋根數（每面）', f:"IF({isCirc}=1,0,SUM('配筋座標'!D3:D"+(2+NMAX)+'))', unit:'根', expr:'詳「配筋座標」D 欄', ref:ref401('25.7.2.3')});
  S.item({key:'nTieY', label:'H 邊繫筋根數（每面）', f:"IF({isCirc}=1,0,SUM('配筋座標'!E3:E"+(2+NMAX)+'))', unit:'根', expr:'詳「配筋座標」E 欄', ref:ref401('25.7.2.3')});
  S.item({key:'hxB', label:'被支撐筋中心距 hx（B 邊）', sym:'hx', f:'IF({isCirc}=1,0,IF({everyB}=1,{pB},2*{pB}))', unit:'cm', fmt:'0.00', crit:'耐震 ≦ 35 cm', ref:ref401('18.7.5.2')});
  S.item({key:'hxH', label:'被支撐筋中心距 hx（H 邊）', sym:'hx', f:'IF({isCirc}=1,0,IF({everyH}=1,{pH},2*{pH}))', unit:'cm', fmt:'0.00', crit:'耐震 ≦ 35 cm', ref:ref401('18.7.5.2')});

  /* ---------------- 四、耐震參數 ---------------- */
  S.section('【四、耐震參數與容量設計】');
  S.item({key:'ph', label:'塑鉸區／柱端加密區', v:inp.ph?'是':'否', kind:'list', list:['是','否'], ref:ref401('18.7.5、18.7.6'),
    crit:'是：啟動容量設計剪力、加密區間距與圍束鋼筋', note:'建築物：柱端加密區（特殊抗彎構材）；橋梁：塑鉸區。'});
  S.item({key:'isPH', label:'　旗標：加密區', f:'--({ph}="是")', ref:'非規範明列條文，係本表判別用'});
  S.item({key:'phio', label:'超強係數 φo（橋梁）', sym:'φo', v:inp.phio, unit:'無因次', kind:'in', fmt:'0.00', crit:'僅橋梁使用；Mo = φo·Mn', ref:'公路橋梁耐震設計規範（容量設計）'});
  S.item({key:'Lv', label:'剪力跨度 Lv（橋梁）', sym:'Lv', v:inp.Lv, unit:'cm', kind:'in', fmt:'#,##0', crit:'僅橋梁使用；Ve = Mo/Lv', ref:'公路橋梁耐震設計規範（容量設計）'});
  S.item({key:'lu', label:'柱淨高 lu（建築）', sym:'lu', v:inp.lu, unit:'cm', kind:'in', fmt:'#,##0', crit:'僅建築物使用；Ve = 2Mpr/lu', ref:ref401('18.7.6.1')});
  S.item({key:'theta', label:'扭矩桁架角 θ', sym:'θ', v:inp.theta, unit:'°', kind:'in', fmt:'0', crit:'非預力構材取 45°', ref:ref401('22.7.6.1.2')});
  S.item({key:'pair', label:'Ash 之 bc 配對', v:inp.ashPerp?'垂直（肢的跨距）':'平行（量至肢外緣）', kind:'list', list:['垂直（肢的跨距）','平行（量至肢外緣）'],
    crit:'關鍵假設：條文對 bc 方向敘述有歧義，箱型牆片兩讀法可差 7 倍以上', ref:ref401('18.7.5.4'),
    note:'關鍵假設。垂直：bc 取與該組 Ash 肢垂直之核心尺寸（肢分布跨距）；平行：取與肢平行之核心尺寸。請對照條文確認。'});
  S.item({key:'PuMax', label:'最大軸壓力 Pu,max', sym:'Pu', f:"MAX(0,MAX('載重組合'!B4:B"+(3+NCB)+'))*1000', unit:'kgf', fmt:'#,##0', expr:'各載重組合 Pu 之最大值', ref:'—'});
  S.item({key:'so', label:'加密區間距參數 so（建築）', sym:'so', f:'MAX(10,MIN(15,10+(35-MAX({hxB},{hxH}))/3))', unit:'cm', fmt:'0.00',
    expr:'so = 10 + (35 − hx)/3，介於 10～15 cm', ref:ref401('18.7.5.3')});
  S.item({key:'sPH', label:'加密區／塑鉸區間距上限', sym:'s', unit:'cm', fmt:FMT_SINF,
    f:`IF({isPH}=1,MIN(MIN({Be},{He})/4,6*{db},IF({isBldg}=1,{so},15)),${BIG})`,
    expr:'橋梁 min(b/4, 6db, 15)；建築 min(b/4, 6db, so)', ref:ref401('18.7.5.3')+'；公路橋梁耐震設計規範'});
  S.item({key:'sTie', label:'橫箍間距通則', sym:'s', unit:'cm', fmt:FMT_SINF,
    f:`IF({isSp}=1,${BIG},MIN(16*{db},48*{dt},MIN({Be},{He})))`, expr:'min(16db, 48dt, 最小邊)；螺箍不適用', ref:ref401('25.7.2.1')});

  /* ---------------- 五、軸力強度與 P-M ---------------- */
  S.section('【五、軸力強度與 P-M 互制（逐點詳「P-M_X」「P-M_Y」，射線交點詳「射線交點」）】');
  S.item({key:'Po', label:'純壓標稱強度', sym:'Po', f:'(0.85*{fc}*({Ag}-{Ast})+{fy}*{Ast})/1000', unit:'tf', fmt:'#,##0',
    expr:"Po = 0.85f'c(Ag − Ast) + fy·Ast", ref:ref401('22.4.2.2')});
  S.item({key:'cap', label:'設計最大軸力', sym:'φPn,max', f:'{phic}*{pmaxf}*{Po}', unit:'tf', fmt:'#,##0', expr:'φc × 截斷係數 × Po', ref:ref401('22.4.2.1')});
  S.item({key:'Pnt', label:'純拉標稱強度', sym:'Pnt', f:'-{fy}*{Ast}/1000', unit:'tf', fmt:'#,##0', expr:'−fy·Ast', ref:ref401('22.4.3')});
  S.item({key:'phiPnt', label:'設計純拉強度', sym:'φPnt', f:'{phit}*{Pnt}', unit:'tf', fmt:'#,##0', ref:ref401('21.2.2')});
  S.item({key:'Pbr', label:'Bresler／載重輪廓分界', sym:"0.1f'cAg", f:'0.1*{fc}*{Ag}/1000', unit:'tf', fmt:'#,##0',
    expr:"Pu ≧ 0.1f'c·Ag 用 Bresler 倒數式，否則用載重輪廓法", ref:'非規範明列條文，係雙軸彎曲慣用分界（Bresler 倒數式之適用範圍）'});
  S.item({key:'DCmax', label:'撓曲應力比 D/C（各組合最大）', sym:'D/C', f:"MAX('載重組合'!M4:M"+(3+NCB)+')', unit:'無因次', fmt:FMT_INF,
    expr:'定偏心射線法：載重點沿 (Mu, Pu) 射線與 φ 包絡線交點', crit:'≦ 1.0 合格', ref:ref401('22.4、10.5.1')});
  S.item({key:'DCctrl', label:'控制組合', f:"IFERROR(INDEX('載重組合'!A4:A"+(3+NCB)+",MATCH({DCmax},'載重組合'!M4:M"+(3+NCB)+',0)),"—")', expr:'D/C 最大者', ref:'—'});

  /* ---------------- 六、剪力與扭矩 ---------------- */
  S.section('【六、剪力與扭矩（逐組合詳「載重組合」工作表）】');
  for(const ax of ['x','y']){
    const X = ax.toUpperCase(), P = ax==='x'?'P-M_X':'P-M_Y', other = ax==='x'?'LY':'LX';   // dir x：剪力沿 x，d 量於 x（取 Y 軸層表）
    S.item({key:'bw'+ax, label:`${X} 向剪力：腹板寬 bw`, sym:'bw', unit:'cm', fmt:'0.0',
      f: ax==='x' ? 'IF({isBox}=1,2*{tw},IF({isCirc}=1,{Din},{He}))' : 'IF({isBox}=1,2*{tw},IF({isCirc}=1,{Din},{Be}))',
      expr:'箱型 2tw；圓形 D；實心取垂直於剪力方向之邊長', ref:ref401('22.5.2')});
    const L = ax==='x' ? "'鋼筋層'!$B$3:$B$" : "'鋼筋層'!$E$3:$E$";
    const A = ax==='x' ? "'鋼筋層'!$C$3:$C$" : "'鋼筋層'!$F$3:$F$";
    const end = 2+LAYROWS;
    S.item({key:'d'+ax, label:`${X} 向剪力：有效深度 d`, sym:'d', unit:'cm', fmt:'0.00',
      f:`IF({isCirc}=1,0.8*{Din},${ax==='x'?'{Be}':'{He}'}/2-IFERROR(SUMPRODUCT(${A}${end},--(${L}${end}<0),${L}${end})/SUMPRODUCT(${A}${end},--(${L}${end}<0)),-${ax==='x'?'{Be}':'{He}'}/2*0.8))`,
      expr:'受壓緣至受拉側鋼筋群形心；圓形 d = 0.8D', ref:ref401('22.5.2.2')});
    S.item({key:'nl'+ax, label:`${X} 向剪力肢數`, sym:'nlegs', unit:'肢',
      f: ax==='x' ? 'IF({isCirc}=1,2,IF({isBox}=1,IF({dbl}="是",4,2),2+IF({everyH}=1,{nH}-2,INT(({nH}-2)/2))))'
                  : 'IF({isCirc}=1,2,IF({isBox}=1,IF({dbl}="是",4,2),2+IF({everyB}=1,{nB}-2,INT(({nB}-2)/2))))',
      expr:'實心：閉合箍筋 2 肢＋該向繫筋；箱型 2 或 4；圓形 Av = 2Asp', ref:ref401('22.5.10.5')});
    S.item({key:'Av'+ax, label:`${X} 向 Av`, sym:'Av', f:`{nl${ax}}*{At}`, unit:'cm²', fmt:'0.000', ref:'—'});
    S.item({key:'Mcd'+ax, label:`${X} 向容量設計彎矩（各組合取大）`, sym:'Mo', unit:'tf·m', fmt:'#,##0.0',
      f:`IF({isBldg}=1,MAX('載重組合'!${ax==='x'?'AD':'AE'}4:${ax==='x'?'AD':'AE'}${3+NCB}),{phio}*MAX('載重組合'!${ax==='x'?'AB':'AC'}4:${ax==='x'?'AB':'AC'}${3+NCB}))`,
      expr:'建築：Mpr（1.25fy、φ = 1）；橋梁：φo·Mn', ref:ref401('18.7.6.1')+'；公路橋梁耐震設計規範'});
    S.item({key:'Ve'+ax, label:`${X} 向容量設計剪力 Ve`, sym:'Ve', unit:'tf', fmt:'#,##0.0',
      f:`IF({isPH}=1,IF({isBldg}=1,2*{Mcd${ax}}*100/{lu},{Mcd${ax}}*100/{Lv}),0)`, expr:'建築 2Mpr/lu；橋梁 Mo/Lv；非加密區 0', ref:ref401('18.7.6.1')});
    S.item({key:'VsMax'+ax, label:`${X} 向 Vs 上限`, sym:'Vs,max', f:`2.12*SQRT({fc})*{bw${ax}}*{d${ax}}/1000`, unit:'tf', fmt:'#,##0.0', expr:"2.12√f'c·bw·d", ref:ref401('22.5.1.2')});
    S.item({key:'sStr'+ax, label:`${X} 向強度需求間距（各組合最小）`, sym:'s', unit:'cm', fmt:FMT_SINF,
      f:`MIN('載重組合'!${ax==='x'?'X':'Y'}4:${ax==='x'?'X':'Y'}${3+NCB})`, expr:'s ≦ Av / (Vs/(fyt·d) + 2At/s)', ref:ref401('22.5.10.5.3、22.7.6.1')});
    S.item({key:'sMin'+ax, label:`${X} 向最小剪力鋼筋量間距`, sym:'s', unit:'cm', fmt:'0.00',
      f:`MIN({Av${ax}}*{fyt}/(0.2*SQRT({fc})*{bw${ax}}),{Av${ax}}*{fyt}/(3.5*{bw${ax}}))`, expr:"Av,min = max(0.2√f'c, 3.5)·bw·s/fyt", ref:ref401('10.6.2.2')});
    S.item({key:'sCode'+ax, label:`${X} 向規範間距上限（各組合最小）`, sym:'s', unit:'cm', fmt:'0.00',
      f:`MIN('載重組合'!${ax==='x'?'Z':'AA'}4:${ax==='x'?'Z':'AA'}${3+NCB})`, expr:"Vs ≦ 1.06√f'c·bw·d：min(d/2, 60)；否則 min(d/4, 30)", ref:ref401('10.7.6.5.2')});
    S.item({key:'fail'+ax, label:`${X} 向斷面不足組合數`, f:`SUM('載重組合'!${ax==='x'?'AF':'AG'}4:${ax==='x'?'AF':'AG'}${3+NCB})`, unit:'組',
      crit:'= 0 合格；Vs > Vs,max 或剪扭應力超限', ref:ref401('22.5.1.2、22.7.7.1')});
  }
  S.item({key:'cH', label:'扭矩：箍筋中心至外緣', f:'{covO}+{dt}/2', unit:'cm', fmt:'0.00', ref:'—'});
  S.item({key:'Dh', label:'扭矩：圓形箍筋中心直徑', f:'MAX(1,{Din}-2*{cH})', unit:'cm', fmt:'0.00', ref:'—'});
  S.item({key:'Aoh', label:'扭矩：箍筋中心線所圍面積', sym:'Aoh', unit:'cm²', fmt:'#,##0',
    f:'IF({isCirc}=1,PI()*{Dh}^2/4,MAX(1,{Be}-2*{cH})*MAX(1,{He}-2*{cH}))', ref:ref401('22.7.6.1')});
  S.item({key:'phh', label:'扭矩：箍筋中心線周長', sym:'ph', unit:'cm', fmt:'#,##0.0',
    f:'IF({isCirc}=1,PI()*{Dh},2*(({Be}-2*{cH})+({He}-2*{cH})))', ref:ref401('22.7.6.1')});
  S.item({key:'pcp', label:'扭矩：外周長', sym:'pcp', unit:'cm', fmt:'#,##0.0', f:'IF({isCirc}=1,PI()*{Din},2*({Be}+{He}))', ref:ref401('22.7.4.1')});
  S.item({key:'Tth', label:'可忽略扭矩門檻', sym:'φTth', unit:'tf·m', fmt:'#,##0.00',
    f:'{phiv}*0.265*SQRT({fc})*{Ag}^2/{pcp}/100000', expr:"φ·0.265√f'c·Acp²/pcp（未計軸壓增益，保守）",
    crit:'Tu < φTth 可忽略扭矩；箱型以 Ag 取代 Acp', ref:ref401('22.7.4.1')});

  /* ---------------- 七、圍束鋼筋 ---------------- */
  S.section('【七、耐震圍束鋼筋 Ash／螺箍體積比 ρs】');
  S.item({key:'kf', label:'kf', sym:'kf', f:'MAX(1,{fc}/1750+0.6)', unit:'無因次', fmt:'0.000', expr:"f'c/1750 + 0.6 ≧ 1.0", ref:ref401('18.7.5.4')});
  S.item({key:'nlS', label:'受側撐主筋數 nl', sym:'nl', f:'IF({isCirc}=1,{nC},4+2*({nTieX}+{nTieY}))', unit:'根', ref:ref401('18.7.5.4')});
  S.item({key:'kn', label:'kn', sym:'kn', f:'{nlS}/MAX(1,{nlS}-2)', unit:'無因次', fmt:'0.000', expr:'nl/(nl − 2)', ref:ref401('18.7.5.4')});
  S.item({key:'termC', label:'適用高軸力第三式', f:'--AND({isBldg}=1,{isPH}=1,{isBox}=0,OR({PuMax}>0.3*{Ag}*{fc},{fc}>700))',
    expr:"建築加密區且 Pu > 0.3Ag·f'c 或 f'c > 700", crit:'1＝適用；此時每根主筋皆須側撐', ref:ref401('18.7.5.4(c)、18.7.5.2(f)')});
  S.item({key:'cx', label:'核心尺寸 bc,x（矩形）', f:'MAX(1,{Be}-2*{covO})', unit:'cm', fmt:'0.0', expr:'B − 2co', ref:ref401('18.7.5.4')});
  S.item({key:'cy', label:'核心尺寸 bc,y（矩形）', f:'MAX(1,{He}-2*{covO})', unit:'cm', fmt:'0.0', expr:'H − 2co', ref:ref401('18.7.5.4')});
  S.item({key:'ct', label:'箱壁核心厚', f:'MAX(1,{tw}-{covO}-{covI})', unit:'cm', fmt:'0.0', expr:'tw − co − ci', ref:'非規範明列條文，係箱型牆片拆解慣用作法'});
  const perp = '({pair}="垂直（肢的跨距）")';
  const rq = (bc, Agw, Achw) => `MAX(0.3*(${bc})*((${Agw})/(${Achw})-1)*{fc}/{fyt},0.09*(${bc})*{fc}/{fyt},IF({termC}=1,0.2*{kf}*{kn}*{PuMax}/({fyt}*(${Achw}))*(${bc}),0))`;
  // 實心：方向1（平行 X 之肢）Ash1 = (2+tieY)At，方向2 Ash2 = (2+tieX)At
  S.item({key:'sA1', label:'實心：方向1 Ash 上限間距', sym:'s', unit:'cm', fmt:FMT_SINF,
    f:`({nTieY}+2)*{At}/${rq(`IF(${perp},{cy},{cx})`,'{Be}*{He}','{cx}*{cy}')}`,
    expr:'Ash1 = (2 + H 邊繫筋)·At；Ash/(s·bc) ≧ max{0.3(Ag/Ach − 1)fc/fyt, 0.09fc/fyt}', ref:ref401('18.7.5.4')});
  S.item({key:'sA2', label:'實心：方向2 Ash 上限間距', sym:'s', unit:'cm', fmt:FMT_SINF,
    f:`({nTieX}+2)*{At}/${rq(`IF(${perp},{cx},{cy})`,'{Be}*{He}','{cx}*{cy}')}`, expr:'Ash2 = (2 + B 邊繫筋)·At', ref:ref401('18.7.5.4')});
  // 箱型：頂/底壁（長 B）與左/右壁（長 H），方向1 沿壁長 2At、方向2 貫穿壁厚繫筋
  S.item({key:'sW1', label:'箱型頂底壁：Ash 上限間距', sym:'s', unit:'cm', fmt:FMT_SINF,
    f:`MIN(2*{At}/${rq(`IF(${perp},{ct},{cx})`,'{Be}*{tw}','{cx}*{ct}')},IF({nTieX}=0,0,{nTieX}*{At}/${rq(`IF(${perp},{cx},{ct})`,'{Be}*{tw}','{cx}*{ct}')}))`,
    expr:'牆片：沿壁長 2At、貫穿壁厚 nTie·At（無繫筋時為 0）', crit:'條文係針對實心柱；箱型拆四片牆片檢核', ref:ref401('18.7.5.4')+'；非規範明列，係牆片拆解作法'});
  S.item({key:'sW2', label:'箱型左右壁：Ash 上限間距', sym:'s', unit:'cm', fmt:FMT_SINF,
    f:`MIN(2*{At}/${rq(`IF(${perp},{ct},{cy})`,'{He}*{tw}','{cy}*{ct}')},IF({nTieY}=0,0,{nTieY}*{At}/${rq(`IF(${perp},{cy},{ct})`,'{He}*{tw}','{cy}*{ct}')}))`,
    expr:'同上（長 H）', ref:ref401('18.7.5.4')});
  S.item({key:'Dc', label:'圓形核心直徑 Dc（量至螺箍外緣）', sym:'Dc', f:'MAX(1,{Din}-2*{covO})', unit:'cm', fmt:'0.0', ref:ref401('25.7.3.3')});
  S.item({key:'ds', label:'螺箍中心線直徑 ds', sym:'ds', f:'MAX(1,{Dc}-{dt})', unit:'cm', fmt:'0.00', ref:'—'});
  S.item({key:'Ach', label:'圓形核心面積 Ach', sym:'Ach', f:'PI()*{Dc}^2/4', unit:'cm²', fmt:'#,##0', ref:ref401('25.7.3.3')});
  S.item({key:'rhoReq', label:'圓形：ρs 需求', sym:'ρs', unit:'無因次', fmt:'0.00000',
    f:'IF(OR({isSp}=1,{isPH}=1),MAX(0.45*({Ag}/{Ach}-1)*{fc}/{fyt},IF({isPH}=1,0.12*{fc}/{fyt},0),IF(AND({isBldg}=1,{isPH}=1,OR({PuMax}>0.3*{Ag}*{fc},{fc}>700)),0.35*{kf}*{PuMax}/({fyt}*{Ach}),0)),0)',
    expr:"max{0.45(Ag/Ach − 1)f'c/fyt（螺箍恆需）, 0.12f'c/fyt（耐震）, 0.35kf·Pu/(fyt·Ach)（建築高軸力）}", ref:ref401('25.7.3.3、18.7.5.4')});
  S.item({key:'sAsh', label:'圍束需求間距 s（依斷面型式）', sym:'s', unit:'cm', fmt:FMT_SINF,
    f:`IF({isCirc}=1,IF({rhoReq}>0,4*{At}*{ds}/({Dc}^2*{rhoReq}),${BIG}),IF({isBox}=1,MIN({sW1},{sW2}),MIN({sA1},{sA2})))`,
    expr:'圓形 s ≦ 4Asp·ds/(Dc²·ρs)；矩形取兩向最小；箱型取牆片最小', ref:ref401('18.7.5.4、25.7.3.3')});

  /* ---------------- 八、箍筋間距彙整 ---------------- */
  S.section('【八、箍筋間距彙整】');
  const cands = [
    ['sStrx','剪力＋扭矩強度需求（X 向）'],['sStry','剪力＋扭矩強度需求（Y 向）'],
    ['sMinx','最小剪力鋼筋量（X 向）'],['sMiny','最小剪力鋼筋量（Y 向）'],
    ['sCodex','規範間距上限（X 向）'],['sCodey','規範間距上限（Y 向）'],['sPH','加密區／塑鉸區上限'],['sTors','扭矩間距上限'],
    ['sAshUse','圍束需求（Ash／ρs）'],['sTie','橫箍間距通則'],['sSpMax','螺箍淨距上限 7.5 cm']];
  S.item({key:'sTors', label:'扭矩間距上限（各組合最小）', sym:'s', unit:'cm', fmt:FMT_SINF,
    f:`MIN('載重組合'!AH4:AH${3+NCB})`, expr:'min(ph/8, 30)（需設計扭矩時）', ref:ref401('25.7.6.1')});
  S.item({key:'sAshUse', label:'圍束需求納入間距控制', sym:'s', unit:'cm', fmt:FMT_SINF,
    f:`IF(OR({isCirc}=1,{isBldg}=0,{isPH}=1),{sAsh},${BIG})`, expr:'建築柱非加密區不以 Ash 控制；橋梁與圓柱恆檢核', ref:ref401('18.7.5.4')});
  S.item({key:'sSpMax', label:'螺箍淨距上限換算間距', sym:'s', unit:'cm', fmt:FMT_SINF, f:`IF({isSp}=1,7.5+{dt},${BIG})`, expr:'s − dt ≦ 7.5 cm', ref:ref401('25.7.3.1')});
  S.item({key:'sGov', label:'控制需求間距', sym:'s,req', unit:'cm', fmt:'0.00',
    f:'MIN('+cands.map(c=>'{'+c[0]+'}').join(',')+')', expr:'上列各上限之最小值', ref:'—'});
  S.item({key:'sGovTag', label:'控制項', f:'INDEX({rng:candName},MATCH({sGov},{rng:candVal},0))', expr:'對照表區「間距候選」', ref:'—'});
  S.item({key:'sUse', label:'採用間距（實務值）', sym:'s', unit:'cm', fmt:'0.0',
    f:'IFERROR(_xlfn.AGGREGATE(14,6,{rng:sList}/({rng:sList}<={sGov}),1),MIN({rng:sList}))', expr:'實務間距表中 ≦ 需求之最大值', ref:'非規範明列條文，係施工慣用間距'});
  S.item({key:'spClr', label:'螺箍淨距', sym:'s − dt', f:`IF({isSp}=1,{sUse}-{dt},${BIG})`, unit:'cm', fmt:FMT_SINF, crit:'≧ max(2.5, 4/3·dagg)', ref:ref401('25.7.3.1')});
  S.item({key:'spMin', label:'螺箍最小淨距', f:'MAX(2.5,4/3*{dagg})', unit:'cm', fmt:'0.00', ref:ref401('25.7.3.1')});

  /* ---------------- 九、檢核彙總 ---------------- */
  S.section('【九、檢核彙總】');
  S.sum({head:true, label:'檢核項目', need:'需求值', cap:'容量／限值', ratio:'比值', judge:'判定'});
  const J = (cond) => ({f:`IF(${cond},"PASS","FAIL")`});
  const sumRows = [];
  const addSum = (o) => { sumRows.push(S.sum(o)); };
  addSum({key:'jDC', label:'撓曲＋軸力 D/C', need:{f:'{DCmax}'}, cap:1, ratio:{f:'{DCmax}'}, judge:J('{DCmax}<=1'), ref:ref401('22.4、10.5.1'), fmt:'0.000'});
  addSum({label:'X 向剪力斷面（Vs ≦ Vs,max）', need:{f:'{failx}'}, cap:0, ratio:{f:'{failx}'}, judge:J('{failx}=0'), note:'不足組合數', ref:ref401('22.5.1.2')});
  addSum({label:'Y 向剪力斷面（Vs ≦ Vs,max）', need:{f:'{faily}'}, cap:0, ratio:{f:'{faily}'}, judge:J('{faily}=0'), note:'不足組合數', ref:ref401('22.5.1.2')});
  addSum({label:'橫向筋間距 s 採用／需求', need:{f:'{sUse}'}, cap:{f:'{sGov}'}, ratio:{f:'{sUse}/{sGov}'}, judge:J('{sUse}<={sGov}'), note:{f:'"控制："&{sGovTag}'}, ref:'—'});
  addSum({label:'主筋比 ρg（%）', need:{f:'{rho}'}, cap:4, ratio:{f:'{rho}/4'}, judge:J('AND({rho}>=1,{rho}<=4)'), note:'下限 1%；4% 為施工性上限（規範 8%）', ref:ref401('10.6.1.1')});
  addSum({label:'主筋淨距（cm）', need:{f:'{need}'}, cap:{f:'MIN({clB},{clH})'}, ratio:{f:'{need}/MIN({clB},{clH})'}, judge:J('MIN({clB},{clH})>={need}'), ref:ref401('25.2.3')});
  addSum({label:'耐震 hx（cm）', need:{f:'MAX({hxB},{hxH})'}, cap:35, ratio:{f:'MAX({hxB},{hxH})/35'}, judge:{f:'IF({isCirc}=1,"N/A",IF(MAX({hxB},{hxH})<=35,"PASS","FAIL"))'}, ref:ref401('18.7.5.2')});
  addSum({label:'螺箍淨距（cm）', need:{f:'{spMin}'}, cap:{f:'{spClr}'}, ratio:{f:`IF({isSp}=1,{spMin}/{spClr},0)`}, judge:{f:'IF({isSp}=0,"N/A",IF({spClr}>={spMin},"PASS","FAIL"))'}, ref:ref401('25.7.3.1')});
  addSum({label:'主筋最少根數', need:{f:'IF({isSp}=1,6,4)'}, cap:{f:'{nBars}'}, ratio:{f:'IF({isSp}=1,6,4)/{nBars}'}, judge:J('{nBars}>=IF({isSp}=1,6,4)'), ref:ref401('10.7.3.1')});
  S.sum({key:'jAll', total:true, label:'總判定', judge:{f:`IF(COUNTIF($E$${'{SUMR1}'}:$E$${'{SUMR2}'},"FAIL")=0,"PASS","NG")`}, note:'僅計 FAIL；N/A 不計', ref:'—'});

  /* ---------------- 十、注意事項 ---------------- */
  S.section('【十、注意事項與使用說明】');
  S.text('色碼圖例：淺藍底＋藍字粗體＋藍框＝手動輸入格；黃底＋藍字粗體＋金框＝附下拉選單之輸入格（可輸入表列外之值）；白底黑字細灰框＝公式；綠字＝連結其他工作表；淺琥珀底＝總判定。', {h:32});
  S.text('1. P-M 互制以中性軸深度 c 於 0.02D～5D 對數取樣 '+NPM+' 點，等值應力塊 a = min(β₁c, D)；與網頁（220 點）之 D/C 可能於小數第三位有差。雙軸載重輪廓法以 '+NIT+' 次二分迭代求解。', {h:30});
  S.text('2. 不計細長效應／彎矩放大：輸入之 Pu、Mu 須為已含 P-Δ 之分析結果。塑性形心僅對對稱配筋成立。', {h:24});
  S.text("3. 剪力 Vc 採 0.53(1 + Nu/140Ag)√f'c·bw·d（ACI 318-14 架構之 §22.5.6.1 形式）；未納入土木401-112（ACI 318-19）§22.5.5.1 新式 Vc 之 ρw 與尺寸效應 λs，d 較大且未達 Av,min 時偏不保守。", {h:32});
  S.text('4. Mander 圍束混凝土曲線為參考資訊，未列入本表（不影響設計判定）。扭矩門檻未計軸壓增益（保守）。箱型牆片之 Ash 拆解非規範明列條文，請自行確認。', {h:30});
  S.text('5. 載重組合請於「載重組合」工作表填入已乘載重因數之設計值（壓為正，tf、tf·m）；最多 '+NCB+' 組，空白列不計。', {h:24});

  /* ---------------- 附錄：對照表 ---------------- */
  S.section('【附錄、對照表區】');
  S.table({cols:[{h:'鋼筋號數',key:'barName',input:true},{h:'直徑 (cm)',key:'barD',input:true,fmt:'0.000'},{h:'面積 (cm²)',key:'barA',input:true,fmt:'0.000'}],
           data:BARS.map(b=>[b[0],b[1],b[2]])});
  S.blank();
  S.table({cols:[{h:'實務箍筋間距 (cm)',key:'sList',input:true,fmt:'0.0'}], data:PRACTICAL_S.map(v=>[v])});
  S.blank();
  S.table({cols:[{h:'間距候選項目',key:'candName'},{h:'值 (cm)',key:'candVal',fmt:FMT_SINF}],
           data:cands.map(c=>[c[1],{f:'{'+c[0]+'}'}])});

  S.layout();
  // 總判定的範圍：彙總列
  const r1 = sumRows[0].r, r2 = sumRows[sumRows.length-1].r;
  const jt = S.rows.find(r=>r.key==='jAll'); jt.judge = {f:`IF(COUNTIF($E$${r1}:$E$${r2},"FAIL")=0,"PASS","NG")`};
  S.keys.jAll = jt.r;   // 結論引用用（判定位於 E 欄）

  /* ---------------- 輔助工作表（先建立名稱以供解析） ---------------- */
  const resolve = makeResolver([S]);
  const order = ['檢核表','結構計算書(A4)','載重組合','P-M_X','P-M_Y','射線交點','雙軸迭代','鋼筋層','配筋座標','圖表資料'];
  const W = {}; for(const nm of order) W[nm] = wb.addWorksheet(nm, nm==='檢核表'?{views:[{state:'frozen', ySplit:2}]}:{});
  const ws = W['檢核表'];
  writeCalcSheet(ws, S, resolve);
  // 總判定 E 欄位址
  const allJudge = `'檢核表'!$E$${jt.r}`;

  buildCoordSheet(W['配筋座標'], resolve);
  buildLayerSheet(W['鋼筋層'], resolve);
  buildPMSheet(W['P-M_X'], resolve, 'x');
  buildPMSheet(W['P-M_Y'], resolve, 'y');
  buildLoadSheet(W['載重組合'], resolve, inp.loads);
  buildRaySheet(W['射線交點'], resolve);
  buildBisectSheet(W['雙軸迭代'], resolve);
  const cinfo = buildColumnChartData(W['圖表資料'], resolve);
  const a4Start = buildA4Column(W['結構計算書(A4)'], S, resolve, inp, jt.r, sumRows);
  return {wb, keys:S.keys, judgeRow:jt.r, charts: columnCharts(inp, cinfo, a4Start)};
}

/* ---------- 配筋座標：每邊主筋座標、繫筋位置旗標、圓周座標 ---------- */
function buildCoordSheet(ws, R){
  const r = f => ({f:R(f,'配筋座標')});
  ws.columns = [6,12,12,10,10,10,10,12,12,12].map(w=>({width:w}));
  put(ws,'A1','配筋座標（自動產生；i 超過根數時為 0／不計）',{sec:true});
  ['i','x 外層','y 外層','B 邊繫筋','H 邊繫筋','內層 x 旗標','內層 y 旗標','圓周角 (rad)','x 圓形','y 圓形']
    .forEach((h,j)=>put(ws,colL(j)+'2',h,{head:true}));
  for(let i=1;i<=NMAX;i++){
    const n=i+2;
    put(ws,'A'+n,i);
    put(ws,'B'+n,r(`IF(AND({isCirc}=0,A${n}<={nB}),-{hxO}+{pB}*(A${n}-1),0)`),{fmt:'0.00'});
    put(ws,'C'+n,r(`IF(AND({isCirc}=0,A${n}<={nH}),-{hyO}+{pH}*(A${n}-1),0)`),{fmt:'0.00'});
    put(ws,'D'+n,r(`IF(AND({isCirc}=0,A${n}>=2,A${n}<={nB}-1,OR({everyB}=1,MOD(A${n}-1,2)=0),OR({inOK}=0,ABS(B${n})<{hxI}-1E-6)),1,0)`));
    put(ws,'E'+n,r(`IF(AND({isCirc}=0,A${n}>=2,A${n}<={nH}-1,OR({everyH}=1,MOD(A${n}-1,2)=0),OR({inOK}=0,ABS(C${n})<{hyI}-1E-6)),1,0)`));
    put(ws,'F'+n,r(`IF(AND({inOK}=1,A${n}<={nB},ABS(B${n})<{hxI}-1E-6),1,0)`));
    put(ws,'G'+n,r(`IF(AND({inOK}=1,A${n}<={nH},ABS(C${n})<{hyI}-1E-6),1,0)`));
    put(ws,'H'+n,r(`IF(AND({isCirc}=1,A${n}<={nC}),PI()/2+2*PI()*(A${n}-1)/{nC},0)`),{fmt:'0.0000'});
    put(ws,'I'+n,r(`IF(AND({isCirc}=1,A${n}<={nC}),{rb}*COS(H${n}),0)`),{fmt:'0.00'});
    put(ws,'J'+n,r(`IF(AND({isCirc}=1,A${n}<={nC}),{rb}*SIN(H${n}),0)`),{fmt:'0.00'});
  }
}

/* ---------- 鋼筋層：兩軸各一張（座標 u、面積 A、至受壓緣深度 d） ----------
   列組：外層 NMAX 列 → 內層 ±h_i 兩列 → 內層側邊 NMAX 列 → 圓形 NMAX 列 */
const LAYROWS = 3*NMAX + 2;
function buildLayerSheet(ws, R){
  const r = f => ({f:R(f,'鋼筋層')});
  ws.columns = [16,12,12,4,12,12,4,12,12].map(w=>({width:w}));
  put(ws,'A1','鋼筋層（Y 軸彎曲：座標 x；X 軸彎曲：座標 y）',{sec:true});
  ['組別','x (cm)','A (cm²)','','y (cm)','A (cm²)'].forEach((h,j)=>{ if(h) put(ws,colL(j)+'2',h,{head:true}); });
  const C = "'配筋座標'!";
  let n=3;
  // 外層：繞 X（y 層）面積：首末列 nB 根、中間 2 根；繞 Y（x 層）：首末 nH 根、中間 2 根
  for(let i=1;i<=NMAX;i++,n++){
    const cr = i+2;
    put(ws,'A'+n,'外層 '+i);
    put(ws,'B'+n,r(`${C}B${cr}`),{fmt:'0.00'});
    put(ws,'C'+n,r(`IF(AND({isCirc}=0,${i}<={nB}),IF(OR(${i}=1,${i}={nB}),{nH},2)*{Ab},0)`),{fmt:'0.000'});
    put(ws,'E'+n,r(`${C}C${cr}`),{fmt:'0.00'});
    put(ws,'F'+n,r(`IF(AND({isCirc}=0,${i}<={nH}),IF(OR(${i}=1,${i}={nH}),{nB},2)*{Ab},0)`),{fmt:'0.000'});
  }
  // 內層 ±h_i
  for(const sg of [1,-1]){
    put(ws,'A'+n,'內層 '+(sg>0?'+':'−')+'h_i');
    put(ws,'B'+n,r(`${sg}*{hxI}*{inOK}`),{fmt:'0.00'}); put(ws,'C'+n,r(`{inOK}*({cntY}+2)*{Ab}`),{fmt:'0.000'});
    put(ws,'E'+n,r(`${sg}*{hyI}*{inOK}`),{fmt:'0.00'}); put(ws,'F'+n,r(`{inOK}*({cntX}+2)*{Ab}`),{fmt:'0.000'});
    n++;
  }
  // 內層側邊：沿用外層座標，落在內緣內者各 2 根
  for(let i=1;i<=NMAX;i++,n++){
    const cr=i+2;
    put(ws,'A'+n,'內層側 '+i);
    put(ws,'B'+n,r(`${C}B${cr}`),{fmt:'0.00'}); put(ws,'C'+n,r(`${C}F${cr}*2*{Ab}`),{fmt:'0.000'});
    put(ws,'E'+n,r(`${C}C${cr}`),{fmt:'0.00'}); put(ws,'F'+n,r(`${C}G${cr}*2*{Ab}`),{fmt:'0.000'});
  }
  // 圓形：每根一層
  for(let i=1;i<=NMAX;i++,n++){
    const cr=i+2;
    put(ws,'A'+n,'圓形 '+i);
    put(ws,'B'+n,r(`${C}I${cr}`),{fmt:'0.00'}); put(ws,'C'+n,r(`IF(AND({isCirc}=1,${i}<={nC}),{Ab},0)`),{fmt:'0.000'});
    put(ws,'E'+n,r(`${C}J${cr}`),{fmt:'0.00'}); put(ws,'F'+n,r(`IF(AND({isCirc}=1,${i}<={nC}),{Ab},0)`),{fmt:'0.000'});
  }
  put(ws,'H2','核對 ΣA',{head:true}); put(ws,'H3',{f:`SUM(C3:C${2+LAYROWS})`},{fmt:'0.00'}); put(ws,'I3',{f:`SUM(F3:F${2+LAYROWS})`},{fmt:'0.00'});
}

/* ---------- P-M 曲線（單軸） ---------- */
function buildPMSheet(ws, R, ax){
  const name = ax==='x'?'P-M_X':'P-M_Y';
  const r = f => ({f:R(f,name)});
  const Ucol = ax==='x' ? 'E' : 'B', Acol = ax==='x' ? 'F' : 'C';
  const end = 2+LAYROWS;
  const U = `'鋼筋層'!$${Ucol}$3:$${Ucol}$${end}`, A = `'鋼筋層'!$${Acol}$3:$${Acol}$${end}`;
  ws.columns = Array(18).fill(0).map((_,j)=>({width: j===0?6:11}));
  put(ws,'A1',`P-M 互制曲線（繞 ${ax.toUpperCase()} 軸；單位 tf、tf·m）`,{sec:true});
  // 參數區
  const P = [
    ['D','受壓深度方向尺寸 D', ax==='x'?'{He}':'{Be}'],
    ['W','寬度 W', ax==='x'?'{Be}':'{He}'],
    ['Dv','中空區深度', `IF({isBox}=1,MAX(0,${ax==='x'?'{He}':'{Be}'}-2*{tw}),0)`],
    ['Wv','中空區寬度', `IF({isBox}=1,MAX(0,${ax==='x'?'{Be}':'{He}'}-2*{tw}),0)`],
    ['Rr','圓半徑', '{Din}/2'],
    ['dmax','最外受拉筋深度', `$C$4/2-_xlfn.AGGREGATE(15,6,${U}/(${A}>0),1)`],
    ['c0','c 起點 0.02D', '0.02*$C$4'],['c1','c 終點 5D', '5*$C$4']
  ];
  put(ws,'A2','參數',{head:true}); put(ws,'B2','說明',{head:true}); put(ws,'C2','值',{head:true});
  P.forEach((p,i)=>{ put(ws,'A'+(4+i),p[0]); put(ws,'B'+(4+i),p[1]); put(ws,'C'+(4+i),r(p[2]),{fmt:'0.000'}); });
  // C4=D C5=W C6=Dv C7=Wv C8=Rr C9=dmax C10=c0 C11=c1
  const H = 14;
  const heads = ['k','c (cm)','a (cm)','A_comp (cm²)','ȳc (cm)','Pn (tf)','Mn (tf·m)','εt','φ','φPn (tf)','φMn (tf·m)','設計 P (tf)','設計 M (tf·m)','Pn,pr (tf)','Mn,pr (tf·m)'];
  heads.forEach((h,j)=>put(ws,colL(j)+H,h,{head:true}));
  const fsArr = (fy, c, a) => {
    const X = `{Es}*{ecu}*(${c}-($C$4/2-${U}))/${c}`;
    return `(${X}-(${X}>${fy})*(${X}-${fy})-(${X}<-${fy})*(${X}+${fy})-0.85*{fc}*(($C$4/2-${U})<=${a}))`;
  };
  for(let k=1;k<=NPM;k++){
    const n=H+k, c=`B${n}`, a=`C${n}`;
    put(ws,'A'+n,k);
    put(ws,'B'+n,{f:`$C$10*($C$11/$C$10)^((A${n}-1)/${NPM-1})`},{fmt:'0.000'});
    put(ws,'C'+n,r(`MIN({b1}*B${n},$C$4)`),{fmt:'0.000'});
    // 受壓面積：圓形弓形／矩形扣中空交集
    const th = `2*ACOS(MAX(-1,MIN(1,($C$8-C${n})/$C$8)))`;
    put(ws,'D'+n,r(`IF({isCirc}=1,$C$8^2*(${th}-SIN(${th}))/2,$C$5*C${n}-IF(AND({isBox}=1,$C$6>0,$C$7>0),$C$7*MAX(0,MIN($C$4/2,$C$6/2)-MAX($C$4/2-C${n},-$C$6/2)),0))`),{fmt:'#,##0.0'});
    put(ws,'E'+n,r(`IF(D${n}<=1E-9,$C$4/2,IF({isCirc}=1,4*$C$8*SIN(${th}/2)^3/(3*(${th}-SIN(${th}))),($C$5*C${n}*($C$4/2-C${n}/2)-IF(AND({isBox}=1,$C$6>0,$C$7>0),$C$7*MAX(0,MIN($C$4/2,$C$6/2)-MAX($C$4/2-C${n},-$C$6/2))*(MIN($C$4/2,$C$6/2)+MAX($C$4/2-C${n},-$C$6/2))/2,0))/D${n}))`),{fmt:'0.000'});
    put(ws,'F'+n,r(`(0.85*{fc}*D${n}+SUMPRODUCT(${A},${fsArr('{fy}',c,a)}))/1000`),{fmt:'#,##0.0'});
    put(ws,'G'+n,r(`ABS(0.85*{fc}*D${n}*E${n}+SUMPRODUCT(${A},${fsArr('{fy}',c,a)},${U}))/100000`),{fmt:'#,##0.0'});
    put(ws,'H'+n,r(`{ecu}*($C$9-B${n})/B${n}`),{fmt:'0.00000'});
    put(ws,'I'+n,r(`IF(H${n}<={ety},{phic},IF(H${n}>=0.005,{phit},{phic}+({phit}-{phic})*(H${n}-{ety})/(0.005-{ety})))`),{fmt:'0.000'});
    put(ws,'J'+n,{f:`I${n}*F${n}`},{fmt:'#,##0.0'});
    put(ws,'K'+n,{f:`I${n}*G${n}`},{fmt:'#,##0.0'});
    // 設計包絡：超過截斷者壓在 φPn,max 線上；首個越過點以內插求交點
    put(ws,'L'+n,r(`MIN(J${n},{cap})`),{fmt:'#,##0.0'});
    put(ws,'M'+n,k===1 ? r(`K${n}`) : r(`IF(AND(J${n}>{cap},J${n-1}<={cap}),K${n-1}+({cap}-J${n-1})/(J${n}-J${n-1})*(K${n}-K${n-1}),K${n})`),{fmt:'#,##0.0'});
    const fpr = '1.25*{fy}';
    put(ws,'N'+n,r(`(0.85*{fc}*D${n}+SUMPRODUCT(${A},${fsArr(fpr,c,a)}))/1000`),{fmt:'#,##0.0'});
    put(ws,'O'+n,r(`ABS(0.85*{fc}*D${n}*E${n}+SUMPRODUCT(${A},${fsArr(fpr,c,a)},${U}))/100000`),{fmt:'#,##0.0'});
  }
  // 設計多邊形（射線交點用）：起點 (0, φt·Pnt)、各點、終點 (0, cap)
  const Q = H, Qc = 17;          // Q 欄＝M、R 欄＝P
  put(ws,'Q'+Q,'多邊形 M',{head:true}); put(ws,'R'+Q,'多邊形 P',{head:true});
  put(ws,'Q'+(Q+1),0); put(ws,'R'+(Q+1),r('{phiPnt}'),{fmt:'#,##0.0'});
  for(let k=1;k<=NPM;k++){ put(ws,'Q'+(Q+1+k),{f:`M${H+k}`},{fmt:'#,##0.0'}); put(ws,'R'+(Q+1+k),{f:`L${H+k}`},{fmt:'#,##0.0'}); }
  put(ws,'Q'+(Q+2+NPM),0); put(ws,'R'+(Q+2+NPM),r('{cap}'),{fmt:'#,##0.0'});
  put(ws,'A12','純壓 Po、截斷與純拉見「檢核表」第五節；多邊形見 Q、R 欄。',{});
}
const PM_H = 14;            // P-M 表頭列
const POLY_N = NPM + 2;     // 多邊形點數

/* ---------- 載重組合：輸入＋逐組合 D/C、剪力 ---------- */
function buildLoadSheet(ws, R, loads){
  const name='載重組合', r = f => ({f:R(f,name)});
  ws.columns = Array(36).fill(0).map((_,j)=>({width: j===0?22:11}));
  put(ws,'A1','載重組合（已乘載重因數之設計值；Pu 壓為正）',{sec:true});
  const heads = {A:'組合名稱',B:'Pu (tf)',C:'Mux (tf·m)',D:'Muy (tf·m)',E:'Vux (tf)',F:'Vuy (tf)',G:'Tu (tf·m)',
    H:'有效：1',I:'X 射線 D/C',J:'Y 射線 D/C',K:'方法',L:'Bresler D/C',M:'D/C',
    N:'X:Vdes (tf)',O:'X:Vc (tf)',P:'X:Vs 需求 (tf)',Q:'Y:Vdes (tf)',R:'Y:Vc (tf)',S:'Y:Vs 需求 (tf)',
    T:'扭矩須設計',U:'At/s (cm²/cm)',V:'X:(Av/s)tot',W:'Y:(Av/s)tot',X:'X:s 強度 (cm)',Y:'Y:s 強度 (cm)',Z:'X:s 規範 (cm)',AA:'Y:s 規範 (cm)',
    AB:'X:Mn@Pu (tf·m)',AC:'Y:Mn@Pu (tf·m)',AD:'X:Mpr@Pu (tf·m)',AE:'Y:Mpr@Pu (tf·m)',AF:'X 斷面不足',AG:'Y 斷面不足',AH:'s 扭矩 (cm)'};
  Object.entries(heads).forEach(([c,h])=>put(ws,c+'3',h,{head:true}));
  const pmLook = (sheet, Pcol, Mcol, P) => {
    const rng = c => `'${sheet}'!$${c}$${PM_H+1}:$${c}$${PM_H+NPM}`;
    return `IF(${P}<=INDEX(${rng(Pcol)},1),INDEX(${rng(Mcol)},1),IF(${P}>=INDEX(${rng(Pcol)},${NPM}),INDEX(${rng(Mcol)},${NPM}),`
      + `INDEX(${rng(Mcol)},MATCH(${P},${rng(Pcol)},1))+(${P}-INDEX(${rng(Pcol)},MATCH(${P},${rng(Pcol)},1)))/(INDEX(${rng(Pcol)},MATCH(${P},${rng(Pcol)},1)+1)-INDEX(${rng(Pcol)},MATCH(${P},${rng(Pcol)},1)))*(INDEX(${rng(Mcol)},MATCH(${P},${rng(Pcol)},1)+1)-INDEX(${rng(Mcol)},MATCH(${P},${rng(Pcol)},1)))))`;
  };
  for(let i=0;i<NCB;i++){
    const n=4+i, L=loads[i];
    put(ws,'A'+n, L?L.name:'', {input:true});
    ['Pu','Mux','Muy','Vux','Vuy','Tu'].forEach((k,j)=> put(ws, colL(1+j)+n, L?(+L[k]||0):null, {input:true, fmt:'#,##0.0'}));
    put(ws,'H'+n,{f:`IF(A${n}="",0,1)`});
    const on = `H${n}=1`;
    const Mx=`ABS(C${n})`, My=`ABS(D${n})`;
    // 單軸射線 D/C（由「射線交點」取 1/min t）
    put(ws,'I'+n,{f:`IF(${on},'射線交點'!${colL(1+i)}${RAY_DC_ROW},"")`},{fmt:FMT_INF});
    put(ws,'J'+n,{f:`IF(${on},'射線交點'!${colL(1+NCB+i)}${RAY_DC_ROW},"")`},{fmt:FMT_INF});
    put(ws,'K'+n, r(`IF(NOT(${on}),"",IF({isCirc}=1,"圓形：合彎矩",IF(AND(${Mx}<1E-9,${My}<1E-9),"純軸力",IF(${My}<1E-9,"單軸（繞 X）",IF(${Mx}<1E-9,"單軸（繞 Y）",IF(B${n}>={Pbr},"Bresler 倒數式","載重輪廓法"))))))`));
    put(ws,'L'+n, r(`IF(K${n}="Bresler 倒數式",IFERROR(B${n}/(1/(1/(B${n}/I${n})+1/(B${n}/J${n})-1/{cap})),${BIG}),"")`),{fmt:FMT_INF});
    put(ws,'M'+n, {f:`IF(NOT(${on}),"",CHOOSE(MATCH(K${n},{"圓形：合彎矩","純軸力","單軸（繞 X）","單軸（繞 Y）","Bresler 倒數式","載重輪廓法"},0),ABS(I${n}),ABS(I${n}),I${n},J${n},L${n},'雙軸迭代'!${colL(1+i)}${BIS_DC_ROW}))`},{fmt:FMT_INF});
    // 剪力（X 向剪力用 X 軸曲線，與網頁相同）
    for(const [ax,V,cVd,cVc,cVs] of [['x','E','N','O','P'],['y','F','Q','R','S']]){
      const Nu=`B${n}*1000`, Vu=`ABS(${V}${n})*1000`;
      put(ws,cVd+n, r(`IF(${on},IF({isPH}=1,MAX({Ve${ax}}*1000,${Vu}),${Vu})/1000,"")`),{fmt:'#,##0.0'});
      put(ws,cVc+n, r(`IF(${on},IF(AND({isPH}=1,${Nu}<IF({isBldg}=1,0.05,0.1)*{fc}*{Ag}),0,MAX(0,0.53*SQRT({fc})*{bw${ax}}*{d${ax}}*IF(${Nu}>=0,1+${Nu}/(140*{Ag}),1+${Nu}/(35*{Ag}))))/1000,"")`),{fmt:'#,##0.0'});
      put(ws,cVs+n, r(`IF(${on},MAX(0,${cVd}${n}/{phiv}-${cVc}${n}),"")`),{fmt:'#,##0.0'});
    }
    put(ws,'T'+n, r(`IF(${on},--(ABS(G${n})>{Tth}),"")`));
    put(ws,'U'+n, r(`IF(${on},IF(T${n}=1,ABS(G${n})*100000/(2*{phiv}*0.85*{Aoh}*{fyt}/TAN({theta}*PI()/180)),0),"")`),{fmt:'0.00000'});
    put(ws,'V'+n, r(`IF(${on},P${n}*1000/({fyt}*{dx})+2*U${n},"")`),{fmt:'0.00000'});
    put(ws,'W'+n, r(`IF(${on},S${n}*1000/({fyt}*{dy})+2*U${n},"")`),{fmt:'0.00000'});
    put(ws,'X'+n, r(`IF(${on},IF(V${n}>1E-9,{Avx}/V${n},${BIG}),"")`),{fmt:FMT_SINF});
    put(ws,'Y'+n, r(`IF(${on},IF(W${n}>1E-9,{Avy}/W${n},${BIG}),"")`),{fmt:FMT_SINF});
    put(ws,'Z'+n, r(`IF(${on},IF(P${n}*1000>1.06*SQRT({fc})*{bwx}*{dx},MIN({dx}/4,30),MIN({dx}/2,60)),"")`),{fmt:'0.00'});
    put(ws,'AA'+n, r(`IF(${on},IF(S${n}*1000>1.06*SQRT({fc})*{bwy}*{dy},MIN({dy}/4,30),MIN({dy}/2,60)),"")`),{fmt:'0.00'});
    put(ws,'AB'+n, {f:`IF(${on},${pmLook('P-M_X','F','G',`B${n}`)},"")`},{fmt:'#,##0.0'});
    put(ws,'AC'+n, {f:`IF(${on},${pmLook('P-M_Y','F','G',`B${n}`)},"")`},{fmt:'#,##0.0'});
    put(ws,'AD'+n, {f:`IF(${on},${pmLook('P-M_X','N','O',`B${n}`)},"")`},{fmt:'#,##0.0'});
    put(ws,'AE'+n, {f:`IF(${on},${pmLook('P-M_Y','N','O',`B${n}`)},"")`},{fmt:'#,##0.0'});
    // 斷面不足：Vs > Vs,max 或 剪扭應力超限（直接相加）
    for(const [ax,cVd,cVc,cVs,col] of [['x','N','O','P','AF'],['y','Q','R','S','AG']]){
      put(ws,col+n, r(`IF(${on},--OR(${cVs}${n}>{VsMax${ax}},${cVd}${n}*1000/({bw${ax}}*{d${ax}})+T${n}*ABS(G${n})*100000*{phh}/(1.7*{Aoh}^2)>{phiv}*(${cVc}${n}*1000/({bw${ax}}*{d${ax}})+2.12*SQRT({fc}))),"")`));
    }
    put(ws,'AH'+n, r(`IF(${on},IF(T${n}=1,MIN({phh}/8,30),${BIG}),"")`),{fmt:FMT_SINF});
  }
}

/* ---------- 射線交點：每組合 × 每多邊形邊 ----------
   t = (M1·dP − P1·dM)/(Mu·dP − Pu·dM)，s = (M1·Pu − P1·Mu)/(Mu·dP − Pu·dM)，
   有效：t > 0、0 ≦ s ≦ 1；D/C = 1/min t（Mu = 0 時改以 Pu/φPn,max 或 Pu/φPnt） */
const RAY_DC_ROW = 4;
function buildRaySheet(ws, R){
  const name='射線交點', r = f=>({f:R(f,name)});
  ws.columns = Array(1+2*NCB).fill(0).map((_,j)=>({width: j===0?14:10}));
  put(ws,'A1','定偏心射線法：設計包絡線各邊之交點參數 t（空白＝無交點）',{sec:true});
  put(ws,'A2','組合',{head:true}); put(ws,'A3','Mu (tf·m)',{head:true}); put(ws,'A4','單軸 D/C',{head:true}); put(ws,'A5','Pu (tf)',{head:true});
  const base = 7;
  put(ws,'A'+(base-1),'邊 k',{head:true});
  for(let side=0; side<2; side++){
    const pm = side===0 ? 'P-M_X' : 'P-M_Y';
    for(let i=0;i<NCB;i++){
      const col = colL(1+side*NCB+i), lr = 4+i;
      put(ws,col+'2',{f:`'載重組合'!A${lr}&"${side===0?'（X）':'（Y）'}"`},{head:true});
      put(ws,col+'3', r(side===0 ? `IF({isCirc}=1,SQRT('載重組合'!C${lr}^2+'載重組合'!D${lr}^2),ABS('載重組合'!C${lr}))` : `ABS('載重組合'!D${lr})`),{fmt:'#,##0.0'});
      put(ws,col+'5',{f:`N('載重組合'!B${lr})`},{fmt:'#,##0.0'});
      const Mu=`${col}$3`, Pu=`${col}$5`;
      for(let k=1;k<POLY_N;k++){
        const n=base+k-1, p1=PM_H+k, p2=PM_H+k+1;
        const M1=`'${pm}'!$Q$${p1}`, P1=`'${pm}'!$R$${p1}`, M2=`'${pm}'!$Q$${p2}`, P2=`'${pm}'!$R$${p2}`;
        if(side===0 && i===0) put(ws,'A'+n,k);
        const den = `(${Mu}*(${P2}-${P1})-${Pu}*(${M2}-${M1}))`;
        const t = `((${M1}*(${P2}-${P1})-${P1}*(${M2}-${M1}))/${den})`;
        const s = `((${M1}*${Pu}-${P1}*${Mu})/${den})`;
        put(ws,col+n,{f:`IF(ABS(${den})<1E-12,"",IF(AND(${t}>1E-9,${s}>=-1E-6,${s}<=1+1E-6),${t},""))`},{fmt:'0.0000'});
      }
      const rng = `${col}${base}:${col}${base+POLY_N-2}`;
      put(ws,col+'4', r(`IF('載重組合'!H${lr}<>1,"",IF(${Mu}<1E-9,IF(${Pu}>0,${Pu}/{cap},IF(${Pu}<0,${Pu}/{phiPnt},0)),IF(COUNT(${rng})=0,${BIG},1/MIN(${rng}))))`),{fmt:FMT_INF});
    }
  }
}

/* ---------- 雙軸載重輪廓法：λ 二分迭代 ----------
   util(λ) = (λMux/φMnx(λPu))^α + (λMuy/φMny(λPu))^α ≦ 1，D/C = 1/λ* */
const BIS_DC_ROW = 4;
function buildBisectSheet(ws, R){
  const name='雙軸迭代', r=f=>({f:R(f,name)});
  ws.columns = Array(1+NCB).fill(0).map((_,j)=>({width: j===0?14:11}));
  put(ws,'A1','載重輪廓法：λ 二分迭代（每組合一欄，每列一次迭代；lo、hi、util 以 | 分欄於下方區塊）',{sec:true});
  put(ws,'A2','組合',{head:true}); put(ws,'A3','有效',{head:true}); put(ws,'A4','D/C = 1/λ*',{head:true});
  const phiMnAt = (pm, P) => {
    const rP = `'${pm}'!$J$${PM_H+1}:$J$${PM_H+NPM}`, rM = `'${pm}'!$K$${PM_H+1}:$K$${PM_H+NPM}`;
    const m = `MATCH(${P},${rP},1)`;
    return `IF(OR(${P}>{cap},${P}<=INDEX(${rP},1),${P}>=INDEX(${rP},${NPM})),0,INDEX(${rM},${m})+(${P}-INDEX(${rP},${m}))/(INDEX(${rP},${m}+1)-INDEX(${rP},${m}))*(INDEX(${rM},${m}+1)-INDEX(${rM},${m})))`;
  };
  // 每組合佔 4 列一組：lo、hi、mid、util；共 NIT 組
  const B0 = 7;
  put(ws,'A6','迭代 j：lo／hi／mid／util',{head:true});
  for(let i=0;i<NCB;i++){
    const col=colL(1+i), lr=4+i;
    put(ws,col+'2',{f:`'載重組合'!A${lr}`},{head:true});
    put(ws,col+'3',{f:`--AND('載重組合'!H${lr}=1,'載重組合'!K${lr}="載重輪廓法")`});
    const Pu=`'載重組合'!$B$${lr}`, Mx=`ABS('載重組合'!$C$${lr})`, My=`ABS('載重組合'!$D$${lr})`;
    const util = lam => `IF(OR(${phiMnAt('P-M_X',`${lam}*${Pu}`)}<=0,${phiMnAt('P-M_Y',`${lam}*${Pu}`)}<=0),1E9,(${lam}*${Mx}/(${phiMnAt('P-M_X',`${lam}*${Pu}`)}))^{alpha}+(${lam}*${My}/(${phiMnAt('P-M_Y',`${lam}*${Pu}`)}))^{alpha})`;
    for(let j=0;j<NIT;j++){
      const b=B0+4*j;
      if(i===0){ put(ws,'A'+b,`j=${j+1} lo`); put(ws,'A'+(b+1),'hi'); put(ws,'A'+(b+2),'mid'); put(ws,'A'+(b+3),'util(mid)'); }
      if(j===0){ put(ws,col+b,0.0001,{fmt:'0.000000'}); put(ws,col+(b+1),50,{fmt:'0.000000'}); }
      else{
        const pb=b-4;
        put(ws,col+b,{f:`IF(${col}${pb+3}<=1,${col}${pb+2},${col}${pb})`},{fmt:'0.000000'});
        put(ws,col+(b+1),{f:`IF(${col}${pb+3}<=1,${col}${pb+1},${col}${pb+2})`},{fmt:'0.000000'});
      }
      put(ws,col+(b+2),{f:`(${col}${b}+${col}${b+1})/2`},{fmt:'0.000000'});
      put(ws,col+(b+3), r(`IF(${col}$3=1,${util(`${col}${b+2}`)},"")`),{fmt:'0.0000'});
    }
    const last=B0+4*(NIT-1);
    put(ws,col+'4', r(`IF(${col}$3<>1,"",IF(${util('0.0001')}>1,${BIG},1/IF(${col}${last+3}<=1,${col}${last+2},${col}${last})))`),{fmt:FMT_INF});
  }
}

/* ======================================================================
   結構計算書 (A4)
   ====================================================================== */
function a4Base(ws){
  ws.columns = [18,158,62,68,44,189].map(pt=>({width:pt/5.6}));
  ws.pageSetup = {paperSize:9, orientation:'portrait', fitToPage:true, fitToWidth:1, fitToHeight:0, horizontalCentered:true,
    margins:{left:42/72, right:28/72, top:42/72, bottom:42/72, header:20/72, footer:20/72}};
  ws.headerFooter = {oddFooter:'&L&D&R第 &P 頁 / 共 &N 頁'};
}
class A4 {
  constructor(ws, R){ this.ws=ws; this.R=R; this.n=0; }
  row(){ return ++this.n; }
  chap(t){ const n=this.row(); this.ws.mergeCells(n,2,n,6); const c=this.ws.getCell(n,2); c.value=t;
    c.font={name:FONT,size:11,bold:true}; c.fill={type:'pattern',pattern:'solid',fgColor:{argb:K.hdr}}; c.alignment={vertical:'middle'}; this.ws.getRow(n).height=22; return n; }
  sub(t){ const n=this.row(); this.ws.mergeCells(n,2,n,6); const c=this.ws.getCell(n,2); c.value=t;
    c.font={name:FONT,size:9,bold:true}; c.fill={type:'pattern',pattern:'solid',fgColor:{argb:K.sec}}; c.alignment={vertical:'middle'}; this.ws.getRow(n).height=18; return n; }
  para(no, text, h){ const n=this.row(); const ws=this.ws; ws.getCell(n,2).value=no; ws.getCell(n,2).font={name:FONT,size:9}; ws.getCell(n,2).alignment={vertical:'top'};
    ws.mergeCells(n,3,n,6); const c=ws.getCell(n,3);
    c.value = (text && typeof text==='object') ? {formula:this.R(text.f,ws.name)} : text;
    c.font={name:FONT,size:9}; c.alignment={wrapText:true, vertical:'top'}; ws.getRow(n).height=h||26; return n; }
  thead(){ const n=this.row(); ['項目','符號','數值','單位','說明／規範出處'].forEach((h,j)=>{ const c=this.ws.getCell(n,2+j); c.value=h;
    c.font={name:FONT,size:9,bold:true}; c.fill={type:'pattern',pattern:'solid',fgColor:{argb:K.hdr}}; c.border=box(K.grid); c.alignment={horizontal:'center',vertical:'middle'}; }); return n; }
  data(label, sym, f, unit, note, fmt, textVal){
    const n=this.row(), ws=this.ws;
    const cells=[label, sym, null, unit, note];
    cells.forEach((v,j)=>{ if(j===2) return; const c=ws.getCell(n,2+j); c.value=v; c.font={name:FONT,size:9}; c.border=box(K.grid);
      c.alignment={horizontal: j===4||j===0?'left':'center', vertical:'middle', wrapText:true}; });
    if(textVal){ ws.getCell(n,5).value=null; ws.mergeCells(n,4,n,5); }
    const c=ws.getCell(n,4);
    c.value={formula:this.R(f, ws.name)}; c.font={name:FONT,size:textVal?8:9,color:{argb:K.link}}; c.border=box(K.grid);
    c.alignment={horizontal:'center',vertical:'middle',wrapText:true}; if(fmt) c.numFmt=fmt;
    ws.getRow(n).height = note && note.length>34 ? 28 : 16;
    return n;
  }
  input(label, v){ const n=this.row(), ws=this.ws; ws.getCell(n,2).value=label; ws.getCell(n,2).font={name:FONT,size:9,bold:true};
    ws.mergeCells(n,3,n,6); const c=ws.getCell(n,3); c.value=v; styleInput(c,false); c.font={name:FONT,size:9,bold:true,color:{argb:K.inFont}}; return n; }
  blank(){ return this.row(); }
}
function buildA4Column(ws, S, R, inp, jRow, sumRows){
  a4Base(ws);
  const a = new A4(ws, R);
  const t = a.row(); ws.mergeCells(t,2,t,6); const tc=ws.getCell(t,2);
  tc.value={formula:R(`IF({isBldg}=1,"RC 柱斷面設計檢核計算書","RC 橋墩／橋塔斷面設計檢核計算書")`,ws.name)};
  tc.font={name:FONT,size:14,bold:true}; tc.alignment={horizontal:'center',vertical:'middle'}; ws.getRow(t).height=26;
  const t2=a.row(); ws.mergeCells(t2,2,t2,6); const tc2=ws.getCell(t2,2); tc2.value='單位：kgf、cm（力顯示 tf、彎矩 tf·m）'; tc2.font={name:FONT,size:9}; tc2.alignment={horizontal:'center'};
  ws.pageSetup.printTitlesRow = `${t}:${t2}`;
  a.input('工程名稱', inp.meta && inp.meta.project || '');
  a.input('構件名稱／編號', inp.meta && inp.meta.member || '');
  a.input('設計者／檢核者', '');
  a.input('計算日期／版次', new Date().toLocaleDateString('zh-TW')+'／第 1 版');

  a.chap('一、設計依據');
  a.para('1.1','混凝土結構設計規範（土木401），內政部國土管理署；斷面強度、剪力、扭矩與鋼筋細則依該規範（以 ACI 318 為基礎）。',28);
  a.para('1.2',{f:'IF({isBldg}=1,"建築物耐震設計規範及解說（柱端加密區依特殊抗彎構材規定）。","公路橋梁耐震設計規範（塑鉸區容量設計，超強係數 φo）。")'},26);
  a.para('1.3','CNS 560 鋼筋混凝土用鋼筋（號數、直徑、面積）。',18);
  a.para('1.4','本計算書數值均連結「檢核表」工作表，修改輸入後自動更新。',18);

  a.chap('二、設計條件');
  a.sub('2.1 幾何形狀'); a.thead();
  a.data('斷面型式','—','{type}','—','實心矩形／中空箱型／圓形',null,true);
  a.data('外寬 × 外高（矩形、箱型）','B × H','TEXT({Be},"0")&" × "&TEXT({He},"0")','cm','圓形時為直徑',null,true);
  a.data('壁厚','tw','{tw}','cm','僅中空箱型','0.0');
  a.data('全斷面積','Ag','{Ag}','cm²','—','#,##0');
  a.data('外側保護層','co','{covO}','cm','土木401 §20.5.1.3','0.0');
  a.sub('2.2 材料與強度折減因數'); a.thead();
  a.data("混凝土抗壓強度","f'c",'{fc}','kgf/cm²','土木401 §19.2','#,##0');
  a.data('主筋／橫向筋降伏強度','fy／fyt','TEXT({fy},"#,##0")&"／"&TEXT({fyt},"#,##0")','kgf/cm²','CNS 560',null,true);
  a.data('等值應力塊係數','β₁','{b1}','—','土木401 §22.2.2.4.3','0.000');
  a.data('強度折減因數','φc／φt／φv','TEXT({phic},"0.00")&"／"&TEXT({phit},"0.00")&"／"&TEXT({phiv},"0.00")','—','土木401 §21.2',null,true);
  a.sub('2.3 載重條件'); a.thead();
  a.data('載重組合數','—',"COUNTIF('載重組合'!H4:H"+(3+NCB)+',1)','組','詳「載重組合」工作表','0');
  a.data('最大軸壓力','Pu,max','{PuMax}/1000','tf','—','#,##0.0');
  a.sub('2.4 配筋條件'); a.thead();
  a.data('主筋','—','{bar}&" × "&{nBars}&" 根"','—','CNS 560',null,true);
  a.data('主筋比','ρg','{rho}','%','土木401 §10.6.1.1','0.000');
  a.data('橫向筋號數','—','{tie}','—',null,null,true);
  a.data('採用間距','s','{sUse}','cm','詳四、4.6','0.0');

  a.chap('三、設計假設');
  a.para('3.1','P-M 互制採應變相容與等值矩形應力塊（土木401 §22.2），a = min(β₁c, D)；純壓、純拉點直接以 Po、−fy·Ast 接上。',28);
  a.para('3.2',{f:'"雙軸彎曲：Pu ≧ 0.1f\'c·Ag = "&TEXT({Pbr},"#,##0")&" tf 時採 Bresler 倒數式，否則採載重輪廓法（α = "&TEXT({alpha},"0.00")&"）；圓形斷面以合彎矩檢核。"'},28);
  a.para('3.3',{f:'IF({isBldg}=1,"容量設計剪力 Ve = 2Mpr/lu，Mpr 以 1.25fy、φ = 1.0 計，lu = "&TEXT({lu},"#,##0")&" cm（土木401 §18.7.6.1）。","容量設計剪力 Ve = φo·Mn/Lv，φo = "&TEXT({phio},"0.00")&"、Lv = "&TEXT({Lv},"#,##0")&" cm（公路橋梁耐震設計規範）。")'},28);
  a.para('3.4',"未納入：細長效應（輸入須已含 P-Δ）、土木401-112 §22.5.5.1 新式 Vc 之尺寸效應 λs、Mander 圍束混凝土強度（僅參考）、扭矩門檻之軸壓增益。",28);

  ws.getRow(a.n).addPageBreak();
  a.chap('四、計算過程');
  a.sub('4.1 軸力強度（土木401 §22.4）'); a.thead();
  a.data('純壓標稱強度','Po','{Po}','tf',"0.85f'c(Ag − Ast) + fy·Ast",'#,##0');
  a.data('設計最大軸力','φPn,max','{cap}','tf','φc × 截斷 × Po','#,##0');
  a.data('設計純拉強度','φPnt','{phiPnt}','tf','φt(−fy·Ast)','#,##0');
  a.sub('4.2 P-M 互制與雙軸（土木401 §22.4、§10.5）'); a.thead();
  a.data('撓曲＋軸力應力比','D/C','{DCmax}','—','定偏心射線法；逐點詳 P-M 工作表',FMT_INF);
  a.data('控制組合','—','{DCctrl}','—',null,null,true);
  a.sub('4.3 剪力（土木401 §22.5、§18.7.6）'); a.thead();
  a.data('X 向容量設計剪力','Ve','{Vex}','tf','建築 2Mpr/lu；橋梁 φo·Mn/Lv','#,##0.0');
  a.data('Y 向容量設計剪力','Ve','{Vey}','tf','同上','#,##0.0');
  a.data('X 向強度需求間距','s','{sStrx}','cm','Av/(Vs/(fyt·d) + 2At/s)',FMT_SINF);
  a.data('Y 向強度需求間距','s','{sStry}','cm','同上',FMT_SINF);
  a.sub('4.4 扭矩（土木401 §22.7）'); a.thead();
  a.data('可忽略扭矩門檻','φTth','{Tth}','tf·m',"φ·0.265√f'c·Acp²/pcp",'#,##0.00');
  a.sub('4.5 圍束鋼筋（土木401 §18.7.5.4、§25.7.3.3）'); a.thead();
  a.data('圍束需求間距','s','{sAsh}','cm','矩形 Ash；圓形 ρs',FMT_SINF);
  a.sub('4.6 橫向筋間距彙整（土木401 §25.7、§18.7.5.3）'); a.thead();
  a.data('控制需求間距','s,req','{sGov}','cm','各上限之最小值','0.00');
  a.data('控制項','—','{sGovTag}','—',null,null,true);
  a.data('採用間距','s','{sUse}','cm','實務間距','0.0');
  a.sub('4.7 施工性（土木401 §25.2.3、§10.6.1.1）'); a.thead();
  a.data('主筋淨距（最小）','s','MIN({clB},{clH})','cm',"需求 max(4.0, 1.5db, 4/3·dagg)",'0.00');
  a.data('主筋比','ρg','{rho}','%','1%～4%','0.000');

  ws.getRow(a.n).addPageBreak();
  a.chap('五、檢核彙總');
  const hn=a.row(); ['檢核項目','需求值','容量／限值','比值','判定'].forEach((h,j)=>{ const c=ws.getCell(hn,2+j); c.value=h;
    c.font={name:FONT,size:9,bold:true}; c.fill={type:'pattern',pattern:'solid',fgColor:{argb:K.hdr}}; c.border=box(K.grid); c.alignment={horizontal:'center'}; });
  for(const sr of sumRows){
    const n=a.row();
    ['A','B','C','D','E'].forEach((col,j)=>{ const c=ws.getCell(n,2+j); c.value={formula:`'檢核表'!${col}${sr.r}`};
      c.font={name:FONT,size:9,color:{argb:K.link}}; c.border=box(K.grid); c.alignment={horizontal:j?'center':'left',wrapText:true};
      if(j>0&&j<4) c.numFmt = j===3?FMT_INF:(sr.fmt||FMT_SINF); });
  }
  const tn=a.row(); ws.mergeCells(tn,2,tn,5); const tl=ws.getCell(tn,2); tl.value='整體結構判定';
  [2,3,4,5,6].forEach(j=>{ const c=ws.getCell(tn,j); c.fill={type:'pattern',pattern:'solid',fgColor:{argb:K.amber}}; c.border=box(K.grid); c.font={name:FONT,size:10,bold:true}; });
  const tj=ws.getCell(tn,6); tj.value={formula:`'檢核表'!E${jRow}`}; tj.alignment={horizontal:'center'};

  a.chap('六、結論');
  a.para('6.1',{f:'"斷面 "&{type}&"，"&IF({isCirc}=1,"D = "&TEXT({Din},"0")&" cm","B × H = "&TEXT({Be},"0")&" × "&TEXT({He},"0")&" cm")&"；主筋 "&{bar}&" × "&{nBars}&" 根（ρg = "&TEXT({rho},"0.00")&"%）；橫向筋 "&{tie}&" @ "&TEXT({sUse},"0.0")&" cm；撓曲 D/C = "&TEXT({DCmax},"0.000")&"；總判定 "&\'檢核表\'!E'+jRow},40);
  a.para('6.2',{f:'"控制項目：撓曲由「"&{DCctrl}&"」控制；橫向筋間距由「"&{sGovTag}&"」控制。"'},26);
  a.para('6.3',{f:'IF(\'檢核表\'!E'+jRow+'="PASS","本斷面各項檢核均符合規定，可供施工圖說使用。","本斷面有檢核項目不符規定，應調整斷面或配筋後重新計算。")'},26);

  a.chap('七、限制與注意事項');
  a.para('7.1','P-M 以 '+NPM+' 點離散；雙軸載重輪廓法以 '+NIT+' 次二分迭代；與網頁結果可能於小數第三位有差。',26);
  a.para('7.2','未計細長效應、未計剪力尺寸效應 λs、扭矩門檻未計軸壓增益；箱型牆片 Ash 拆解與 bc 配對為關鍵假設，須自行確認。',28);
  a.para('7.3','輸入資料（尺寸、材料、載重）須經設計者核實；版次修改應更新封面版次。',22);
  a.para('7.4','色碼：淺藍底藍框＝輸入；黃底金框＝下拉輸入；白底細灰框＝公式；綠字＝連結；淺琥珀底＝總判定。',26);

  a.blank();
  const sg=a.row(), sl=a.row(); ws.getRow(sl).height=45;
  ws.mergeCells(sg,3,sg,4); ws.mergeCells(sg,5,sg,6); ws.mergeCells(sl,3,sl,4); ws.mergeCells(sl,5,sl,6);
  [[2,'設計'],[3,'校核'],[5,'審核']].forEach(([j,h])=>{ const c=ws.getCell(sg,j); c.value=h; c.font={name:FONT,size:9,bold:true}; c.alignment={horizontal:'center'};
    ws.getCell(sl,j).border={bottom:side(K.black)}; });
  const fig = a4Figures(ws, a, '附圖　斷面配筋圖與 P-M 互制曲線（Excel 圖表，隨輸入自動更新）', inp.type==='circle'?1:2);
  ws.pageSetup.printArea = `A1:F${a.n}`;
  return fig;
}

/* ======================================================================
   梁／版活頁簿
   ====================================================================== */
const NBI = 50;     // 中性軸二分迭代次數
function buildBeam(ExcelJS, inp){
  const wb = new ExcelJS.Workbook();
  wb.creator='RC 斷面設計工具'; wb.created=new Date(); wb.calcProperties={fullCalcOnLoad:true};
  const S = new CalcSheet('檢核表');
  const ref401 = s => `土木401 §${s}（ACI 318 同條號）`;
  S.title('RC 梁／版斷面設計檢核表（矩形梁／T 型梁／單位寬版，kgf-cm 制）');
  S.header();

  /* ---------------- 一、斷面幾何 ---------------- */
  S.section('【一、設計依據與斷面幾何】');
  S.item({key:'type', label:'斷面型式', v:inp.slab?'單位寬版':(inp.type==='T'?'T 型梁':'矩形梁'), kind:'list', list:['矩形梁','T 型梁','單位寬版'],
    expr:'下拉選擇', crit:'T 型梁僅正彎矩時翼緣受壓；單位寬版 b = 100 cm、不配剪力筋', ref:ref401('6.3.2、7、9'),
    note:'矩形梁、T 型梁、單位寬版（每公尺寬，鋼筋以號數＠間距輸入）。'});
  S.item({key:'isT', label:'　旗標：T 型', f:'--({type}="T 型梁")', ref:'非規範明列條文，係本表判別用'});
  S.item({key:'isSlab', label:'　旗標：版', f:'--({type}="單位寬版")', ref:'非規範明列條文，係本表判別用'});
  S.item({key:'sign', label:'檢核斷面', v:inp.sign==='neg'?'負彎矩（頂層受拉）':'正彎矩（底層受拉）', kind:'list', list:['正彎矩（底層受拉）','負彎矩（頂層受拉）'],
    crit:'負彎矩時 T 梁翼緣受拉不計入受壓區', ref:ref401('22.2'), note:'正彎矩：跨中，頂面受壓；負彎矩：支承，底面受壓。'});
  S.item({key:'isPos', label:'　旗標：正彎矩', f:'--({sign}="正彎矩（底層受拉）")', ref:'非規範明列條文，係本表判別用'});
  S.item({key:'env', label:'環境', v:inp.env==='water'?'水工／環境結構':'一般環境', kind:'list', list:['一般環境','水工／環境結構'],
    crit:'水工：檢核裂縫寬度、保護層 ≧ 5 cm、版 ρmin 提高', ref:'非規範明列條文，係參考 ACI 350、ACI 224R 之加嚴作法',
    note:'水工／環境結構：本表另檢核使用載重裂縫寬度；未含 ACI 350 環境耐久係數 Sd。'});
  S.item({key:'isW', label:'　旗標：水工', f:'--({env}="水工／環境結構")', ref:'非規範明列條文，係本表判別用'});
  S.item({key:'bwIn', label:'腹板寬 bw（梁）', sym:'bw', v:inp.bw, unit:'cm', kind:'in', fmt:'0.0', crit:'版時不使用（b = 100）', ref:'—（幾何輸入）'});
  S.item({key:'bw', label:'計算用寬度', sym:'b', f:'IF({isSlab}=1,100,{bwIn})', unit:'cm', fmt:'0.0', expr:'版取 100 cm', ref:'—'});
  S.item({key:'h', label:'全深 h', sym:'h', v:inp.h, unit:'cm', kind:'in', fmt:'0.0', ref:'—（幾何輸入）'});
  S.item({key:'hfIn', label:'翼緣厚 hf（T 梁）', sym:'hf', v:inp.hf, unit:'cm', kind:'in', fmt:'0.0', crit:'T 型梁使用', ref:'—（幾何輸入）'});
  S.item({key:'hf', label:'計算用 hf', sym:'hf', f:'IF({isT}=1,IF({hfIn}>={h},{h}/3,MAX(0,{hfIn})),0)', unit:'cm', fmt:'0.0', expr:'非 T 型取 0；hf ≧ h 時取 h/3', ref:'—'});
  S.item({key:'L', label:'跨度 L', sym:'L', v:inp.L, unit:'cm', kind:'in', fmt:'#,##0', crit:'有效翼緣寬與撓度使用', ref:'—'});
  S.item({key:'spc', label:'梁中心距（T 梁內梁）', v:inp.sSpacing, unit:'cm', kind:'in', fmt:'#,##0', ref:ref401('6.3.2.1')});
  S.item({key:'edge', label:'梁位置', v:inp.edge?'邊梁（單側翼緣）':'內梁（雙側翼緣）', kind:'list', list:['內梁（雙側翼緣）','邊梁（單側翼緣）'], ref:ref401('6.3.2.1'),
    note:'內梁 be = min(L/4, bw+16hf, 梁中心距)；邊梁 be = bw + min(L/12, 6hf, 淨距/2)。'});
  S.item({key:'sClr', label:'邊梁淨距', v:inp.sClear, unit:'cm', kind:'in', fmt:'#,##0', ref:ref401('6.3.2.1')});
  S.item({key:'be', label:'有效翼緣寬', sym:'be', unit:'cm', fmt:'0.0',
    f:'IF({isT}=1,IF({edge}="邊梁（單側翼緣）",MIN({bw}+{L}/12,{bw}+6*{hf},{bw}+{sClr}/2),MIN({L}/4,{bw}+16*{hf},{spc})),{bw})',
    expr:'內梁 min(L/4, bw+16hf, 中心距)；邊梁 min(bw+L/12, bw+6hf, bw+淨距/2)', ref:ref401('6.3.2.1')});
  S.item({key:'cover', label:'保護層 cc', sym:'cc', v:inp.cover, unit:'cm', kind:'in', fmt:'0.0', crit:'水工建議 ≧ 5 cm', ref:ref401('20.5.1.3')});
  S.item({key:'dagg', label:'骨材最大粒徑', sym:'dagg', v:inp.dagg, unit:'cm', kind:'in', fmt:'0.0', ref:ref401('25.2.1')});

  /* ---------------- 二、材料 ---------------- */
  S.section('【二、材料與強度折減因數】');
  S.item({key:'fc', label:"混凝土抗壓強度 f'c", sym:"f'c", v:inp.fc, unit:'kgf/cm²', kind:'in', fmt:'#,##0', ref:ref401('19.2')});
  S.item({key:'fy', label:'主筋降伏強度 fy', sym:'fy', v:inp.fy, unit:'kgf/cm²', kind:'list', list:['2800','4200','5000'], fmt:'#,##0', ref:'CNS 560', note:'SD280＝2800、SD420＝4200、SD490＝5000（CNS 560）。'});
  S.item({key:'fyt', label:'箍筋降伏強度 fyt', sym:'fyt', v:inp.fyt, unit:'kgf/cm²', kind:'list', list:['2800','4200'], fmt:'#,##0', ref:'CNS 560', note:'SD280＝2800、SD420＝4200（CNS 560）。'});
  S.item({key:'Es', label:'鋼筋彈性模數 Es', sym:'Es', v:inp.Es, unit:'kgf/cm²', kind:'in', fmt:'#,##0', ref:ref401('20.2.2.2')});
  S.item({key:'ecu', label:'極限壓應變 εcu', sym:'εcu', v:inp.ecu, unit:'無因次', kind:'in', fmt:'0.0000', ref:ref401('22.2.2.1')});
  S.item({key:'b1', label:'等值應力塊係數 β₁', sym:'β₁', f:'MAX(0.65,MIN(0.85,0.85-0.05*({fc}-280)/70))', unit:'無因次', fmt:'0.000', ref:ref401('22.2.2.4.3')});
  S.item({key:'Ec', label:'混凝土彈性模數 Ec', sym:'Ec', f:'15000*SQRT({fc})', unit:'kgf/cm²', fmt:'#,##0', ref:ref401('19.2.2.1')});
  S.item({key:'ety', label:'降伏應變 εty', sym:'εty', f:'{fy}/{Es}', unit:'無因次', fmt:'0.00000', ref:ref401('21.2.2')});
  S.item({key:'n', label:'彈性模數比 n', sym:'n', f:'{Es}/{Ec}', unit:'無因次', fmt:'0.000', ref:ref401('24.2.3.5')});
  S.item({key:'phic', label:'壓力控制 φc', sym:'φc', v:inp.phic, unit:'無因次', kind:'in', fmt:'0.00', ref:ref401('21.2.2')});
  S.item({key:'phit', label:'拉力控制 φt', sym:'φt', v:inp.phit, unit:'無因次', kind:'in', fmt:'0.00', ref:ref401('21.2.2'), crit:'本表拉控門檻取 εt ≧ 0.005（保守）'});
  S.item({key:'phiv', label:'剪力／扭矩 φv', sym:'φv', v:inp.phiv, unit:'無因次', kind:'in', fmt:'0.00', ref:ref401('21.2.1'), crit:"關鍵假設：0.75 搭配 Vc = 0.53√f'c·bw·d"});

  /* ---------------- 三、配筋 ---------------- */
  S.section('【三、配筋】');
  S.item({key:'botS', label:'底筋號數', v:inp.botSize, kind:'list', list:'{rng:barName}', ref:'CNS 560'});
  S.item({key:'dbB', label:'底筋直徑', sym:'db', f:'INDEX({rng:barD},MATCH({botS},{rng:barName},0))', unit:'cm', fmt:'0.000', ref:'CNS 560'});
  S.item({key:'AbB', label:'底筋單根面積', sym:'Ab', f:'INDEX({rng:barA},MATCH({botS},{rng:barName},0))', unit:'cm²', fmt:'0.000', ref:'CNS 560'});
  S.item({key:'topS', label:'頂筋號數', v:inp.topSize, kind:'list', list:'{rng:barName}', ref:'CNS 560'});
  S.item({key:'dbT', label:'頂筋直徑', sym:'db', f:'INDEX({rng:barD},MATCH({topS},{rng:barName},0))', unit:'cm', fmt:'0.000', ref:'CNS 560'});
  S.item({key:'AbT', label:'頂筋單根面積', sym:'Ab', f:'INDEX({rng:barA},MATCH({topS},{rng:barName},0))', unit:'cm²', fmt:'0.000', ref:'CNS 560'});
  S.item({key:'tieS', label:'箍筋號數（梁）', v:inp.tieSize, kind:'list', list:'{rng:barName}', ref:'CNS 560'});
  S.item({key:'dt', label:'箍筋直徑（版取 0）', sym:'dt', f:'IF({isSlab}=1,0,INDEX({rng:barD},MATCH({tieS},{rng:barName},0)))', unit:'cm', fmt:'0.000', ref:'CNS 560'});
  S.item({key:'At', label:'箍筋單肢面積', sym:'At', f:'INDEX({rng:barA},MATCH({tieS},{rng:barName},0))', unit:'cm²', fmt:'0.000', ref:'CNS 560'});
  S.item({key:'nBotIn', label:'底筋根數（梁）', v:inp.nBot, unit:'根', kind:'in', ref:'—'});
  S.item({key:'nTopIn', label:'頂筋根數（梁）', v:inp.nTop, unit:'根', kind:'in', ref:'—'});
  S.item({key:'spB', label:'底層主筋間距（版）', v:inp.spBot, unit:'cm', kind:'in', fmt:'0.0', ref:ref401('7.7.2.3')});
  S.item({key:'spT', label:'頂層主筋間距（版；0＝不配置）', v:inp.spTop, unit:'cm', kind:'in', fmt:'0.0', ref:ref401('7.7.2.3')});
  S.item({key:'nBot', label:'計算用底筋根數（每公尺或每梁）', f:'IF({isSlab}=1,100/{spB},MAX(1,ROUND({nBotIn},0)))', unit:'根', fmt:'0.00', expr:'版 100/s', ref:'—'});
  S.item({key:'nTop', label:'計算用頂筋根數', f:'IF({isSlab}=1,IF({spT}>0,100/{spT},0),MAX(0,ROUND({nTopIn},0)))', unit:'根', fmt:'0.00', ref:'—'});
  S.item({key:'nLegs', label:'箍筋肢數', sym:'nlegs', v:inp.nLegs, unit:'肢', kind:'in', ref:ref401('9.7.6.2')});
  S.item({key:'perRow', label:'每排可排根數', f:'IF({isSlab}=1,{nBot},MAX(1,INT(({bw}-2*({cover}+{dt})+2.5)/({dbB}+2.5))))', unit:'根', expr:'淨距 2.5 cm', ref:ref401('25.2.1')});
  S.item({key:'rows', label:'底筋排數', f:'IF({isSlab}=1,1,MIN(2,ROUNDUP({nBot}/MAX(1,{perRow}),0)))', unit:'排', expr:'自動判斷，上限 2 排', ref:'—'});
  S.item({key:'r1', label:'第 1 排根數', f:'IF({rows}<=1,{nBot},ROUNDUP({nBot}/2,0))', unit:'根', fmt:'0.00', ref:'—'});
  S.item({key:'r2', label:'第 2 排根數', f:'{nBot}-{r1}', unit:'根', fmt:'0.00', ref:'—'});
  S.item({key:'y1', label:'底筋第 1 排中心高', f:'IF({isSlab}=1,{cover}+{dbB}/2,{cover}+{dt}+{dbB}/2)', unit:'cm', fmt:'0.00', ref:'—'});
  S.item({key:'y2', label:'底筋第 2 排中心高', f:'{y1}+{dbB}+2.5', unit:'cm', fmt:'0.00', expr:'排距 2.5 cm', ref:ref401('25.2.2')});
  S.item({key:'yT', label:'頂筋中心高', f:'{h}-({cover}+{dt}+{dbT}/2)', unit:'cm', fmt:'0.00', ref:'—'});
  S.item({key:'A1', label:'第 1 排面積', f:'{r1}*{AbB}', unit:'cm²', fmt:'0.00', ref:'—'});
  S.item({key:'A2', label:'第 2 排面積', f:'{r2}*{AbB}', unit:'cm²', fmt:'0.00', ref:'—'});
  S.item({key:'AT', label:'頂筋面積', f:'{nTop}*{AbT}', unit:'cm²', fmt:'0.00', ref:'—'});
  // 至受壓緣深度（依檢核斷面）
  S.item({key:'d1', label:'第 1 排至受壓緣', f:'IF({isPos}=1,{h}-{y1},{y1})', unit:'cm', fmt:'0.00', ref:'—'});
  S.item({key:'d2', label:'第 2 排至受壓緣', f:'IF({isPos}=1,{h}-{y2},{y2})', unit:'cm', fmt:'0.00', ref:'—'});
  S.item({key:'dT', label:'頂筋至受壓緣', f:'IF({isPos}=1,{h}-{yT},{yT})', unit:'cm', fmt:'0.00', ref:'—'});
  S.item({key:'d', label:'有效深度 d', sym:'d', unit:'cm', fmt:'0.00',
    f:'IFERROR(SUMPRODUCT(({d1}>{h}/2)*{A1}*{d1}+({d2}>{h}/2)*{A2}*{d2}+({dT}>{h}/2)*{AT}*{dT})/SUMPRODUCT(({d1}>{h}/2)*{A1}+({d2}>{h}/2)*{A2}+({dT}>{h}/2)*{AT}),0.8*{h})',
    expr:'受拉側鋼筋群形心', ref:ref401('22.2')});
  S.item({key:'AsT', label:'受拉鋼筋 As', sym:'As', f:'IF({isPos}=1,{A1}+{A2},{AT})', unit:'cm²', fmt:'0.00', ref:'—'});
  S.item({key:'rho', label:'受拉鋼筋比 ρ', sym:'ρ', f:'{AsT}/({bw}*{d})', unit:'無因次', fmt:'0.00000', ref:'—'});
  S.item({key:'rhoTs', label:'版最小鋼筋比 ρmin', sym:'ρmin', v:inp.rhoTs, unit:'無因次', kind:'in', fmt:'0.0000',
    crit:'關鍵假設：一般 0.0018；水工依伸縮縫間距約 0.003～0.005', ref:ref401('24.4.3.2')+'；水工參考 ACI 350 §7.12.2.1（非我國規範）',
    note:'關鍵假設。fy = 4200 時 0.0018（土木401 §24.4.3.2）；水工結構依伸縮縫間距約 0.003～0.005（ACI 350，我國規範未明列）。'});
  S.item({key:'AsMin', label:'最少受拉鋼筋 As,min', sym:'As,min', unit:'cm²', fmt:'0.00',
    f:'IF({isSlab}=1,{rhoTs}*{bw}*{h},MAX(0.8*SQRT({fc})/{fy},14/{fy})*{bw}*{d})',
    expr:"梁 max(0.8√f'c/fy, 14/fy)·bw·d；版 ρmin·b·h", ref:ref401('9.6.1.2、7.6.1.1')});

  /* ---------------- 四、撓曲 ---------------- */
  S.section('【四、撓曲強度（中性軸以 ΣF = 0 二分求解，詳「撓曲求解」）】');
  S.item({key:'c', label:'中性軸深度 c', sym:'c', f:"'撓曲求解'!B3", unit:'cm', fmt:'0.000', kind:'link', expr:'ΣF(c) = 0', ref:ref401('22.2.1')});
  S.item({key:'a', label:'應力塊深度 a', sym:'a', f:'MIN({b1}*{c},{h})', unit:'cm', fmt:'0.000', ref:ref401('22.2.2.4.1')});
  S.item({key:'b1c', label:'受壓翼緣寬（T 梁正彎矩為 be）', f:'IF(AND({isT}=1,{isPos}=1),{be},{bw})', unit:'cm', fmt:'0.0', ref:ref401('6.3.2')});
  S.item({key:'hfc', label:'受壓翼緣厚', f:'IF(AND({isT}=1,{isPos}=1),{hf},0)', unit:'cm', fmt:'0.0', ref:'—'});
  S.item({key:'Ac', label:'受壓區面積', sym:'Ac', unit:'cm²', fmt:'#,##0.0',
    f:'IF(OR({hfc}=0,{a}<={hfc}),{b1c}*{a},{b1c}*{hfc}+{bw}*({a}-{hfc}))', expr:'a ≦ hf：be·a；否則 be·hf + bw(a − hf)', ref:ref401('22.2.2.4')});
  S.item({key:'Sc', label:'受壓區對受壓緣一次矩', unit:'cm³', fmt:'#,##0',
    f:'IF(OR({hfc}=0,{a}<={hfc}),{b1c}*{a}^2/2,{b1c}*{hfc}^2/2+{bw}*({a}-{hfc})*({hfc}+({a}-{hfc})/2))', ref:'—'});
  const fsL = (dk) => `MAX(-{fy},MIN({fy},{Es}*{ecu}*({c}-${dk})/{c}))-IF(${dk}<={a},0.85*{fc},0)`;
  S.item({key:'fs1', label:'第 1 排應力', sym:'fs', f:fsL('{d1}'), unit:'kgf/cm²', fmt:'#,##0', expr:'Es·εcu(c − d)/c，夾於 ±fy；受壓區內扣 0.85f\'c', ref:ref401('22.2.3')});
  S.item({key:'fs2', label:'第 2 排應力', sym:'fs', f:fsL('{d2}'), unit:'kgf/cm²', fmt:'#,##0', ref:ref401('22.2.3')});
  S.item({key:'fsT', label:'頂筋應力', sym:'fs', f:fsL('{dT}'), unit:'kgf/cm²', fmt:'#,##0', ref:ref401('22.2.3')});
  S.item({key:'Cc', label:'混凝土壓力', sym:'Cc', f:'0.85*{fc}*{Ac}', unit:'kgf', fmt:'#,##0', ref:ref401('22.2.2.4.1')});
  S.item({key:'resid', label:'軸力平衡殘差 ΣF', f:'{Cc}+{fs1}*{A1}+{fs2}*{A2}+{fsT}*{AT}', unit:'kgf', fmt:'0.00', crit:'應趨近 0', ref:'—'});
  S.item({key:'Mn', label:'標稱彎矩強度 Mn', sym:'Mn', unit:'tf·m', fmt:'#,##0.00',
    f:'(-0.85*{fc}*{Sc}-{fs1}*{A1}*{d1}-{fs2}*{A2}*{d2}-{fsT}*{AT}*{dT})/100000', expr:'對受壓緣取矩', ref:ref401('22.3')});
  S.item({key:'dmax', label:'最外受拉筋深度', f:'MAX({d1},IF({r2}>0,{d2},-1E9),{dT})', unit:'cm', fmt:'0.00', ref:'—'});
  S.item({key:'et', label:'最外受拉筋應變 εt', sym:'εt', f:'{ecu}*({dmax}-{c})/{c}', unit:'無因次', fmt:'0.00000', crit:'≧ 0.004（受撓構材）', ref:ref401('9.3.3.1')});
  S.item({key:'phi', label:'強度折減因數 φ', sym:'φ', f:'IF({et}<={ety},{phic},IF({et}>=0.005,{phit},{phic}+({phit}-{phic})*({et}-{ety})/(0.005-{ety})))', unit:'無因次', fmt:'0.000', ref:ref401('21.2.2')});
  S.item({key:'phiMn', label:'設計彎矩強度 φMn', sym:'φMn', f:'{phi}*{Mn}', unit:'tf·m', fmt:'#,##0.00', ref:ref401('9.5.1.1')});
  S.item({key:'MuMax', label:'需求彎矩 Mu（各組合最大）', sym:'Mu', f:"MAX('載重組合'!G4:G"+(3+NCB)+')', unit:'tf·m', fmt:'#,##0.00', expr:'依檢核斷面取 Mu⁺ 或 Mu⁻', ref:'—'});
  S.item({key:'DC', label:'撓曲 D/C', sym:'D/C', f:'IF({phiMn}>0,{MuMax}/{phiMn},1E9)', unit:'無因次', fmt:FMT_INF, crit:'≦ 1.0', ref:ref401('9.5.1.1')});
  S.item({key:'Mpos', label:'φMn⁺（正彎矩，耐震用）', f:"'撓曲求解'!F3", unit:'tf·m', fmt:'#,##0.00', kind:'link', ref:ref401('18.6.3.2')});
  S.item({key:'Mneg', label:'φMn⁻（負彎矩，耐震用）', f:"'撓曲求解'!F4", unit:'tf·m', fmt:'#,##0.00', kind:'link', ref:ref401('18.6.3.2')});
  S.item({key:'MprP', label:'Mpr⁺（1.25fy、φ=1）', f:"'撓曲求解'!F5", unit:'tf·m', fmt:'#,##0.00', kind:'link', ref:ref401('18.6.5.1')});
  S.item({key:'MprN', label:'Mpr⁻（1.25fy、φ=1）', f:"'撓曲求解'!F6", unit:'tf·m', fmt:'#,##0.00', kind:'link', ref:ref401('18.6.5.1')});

  /* ---------------- 五、剪力與扭矩 ---------------- */
  S.section('【五、剪力與扭矩】');
  S.item({key:'seis', label:'耐震特殊抗彎構材（梁）', v:inp.seismic?'是':'否', kind:'list', list:['是','否'], ref:ref401('18.6'), note:'是：容量設計剪力、加密區、bw／ln／ρ 限制與接頭面彎矩比。版恆為否。'});
  S.item({key:'isS', label:'　旗標：耐震', f:'--AND({seis}="是",{isSlab}=0)', ref:'非規範明列條文，係本表判別用'});
  S.item({key:'ln', label:'淨跨 ln', sym:'ln', v:inp.ln, unit:'cm', kind:'in', fmt:'#,##0', ref:ref401('18.6.5.1')});
  S.item({key:'Vg', label:'重力剪力 Vg', sym:'Vg', v:inp.Vg, unit:'tf', kind:'in', fmt:'0.0', ref:ref401('18.6.5.1')});
  S.item({key:'theta', label:'扭矩桁架角 θ', sym:'θ', v:inp.theta, unit:'°', kind:'in', ref:ref401('22.7.6.1.2')});
  S.item({key:'Vu', label:'設計剪力 Vu（各組合最大）', sym:'Vu', f:"MAX('載重組合'!H4:H"+(3+NCB)+')', unit:'tf', fmt:'#,##0.0', ref:'—'});
  S.item({key:'Tu', label:'對應扭矩 Tu', sym:'Tu', f:"IFERROR(INDEX('載重組合'!E4:E"+(3+NCB)+",MATCH({Vu},'載重組合'!H4:H"+(3+NCB)+',0)),0)', unit:'tf·m', fmt:'#,##0.00', expr:'取 Vu 最大組合之扭矩', ref:'—'});
  S.item({key:'Ve', label:'容量設計剪力 Ve', sym:'Ve', f:'IF({isS}=1,IF({ln}>0,({MprP}+{MprN})*100/{ln}+{Vg},0),0)', unit:'tf', fmt:'#,##0.0', expr:'(Mpr⁺ + Mpr⁻)/ln + Vg', ref:ref401('18.6.5.1')});
  S.item({key:'Vdes', label:'設計剪力', sym:'Vdes', f:'IF({isS}=1,MAX({Ve},{Vu}),{Vu})', unit:'tf', fmt:'#,##0.0', ref:ref401('18.6.5.1')});
  S.item({key:'VcZ', label:'Vc 歸零', f:'--AND({isS}=1,{Vdes}>0,({Ve}-{Vg})>=0.5*{Vdes})', expr:'地震剪力佔比 ≧ 50%', crit:'1＝Vc 取 0', ref:ref401('18.6.5.2')});
  S.item({key:'Vc', label:'混凝土剪力強度 Vc', sym:'Vc', f:'IF({VcZ}=1,0,0.53*SQRT({fc})*{bw}*{d}/1000)', unit:'tf', fmt:'#,##0.00', expr:"0.53√f'c·bw·d", ref:ref401('22.5.5.1')+'（舊式；未含 λs）'});
  S.item({key:'phiVc', label:'φVc（版不配剪力筋時須 ≧ Vu）', sym:'φVc', f:'{phiv}*{Vc}', unit:'tf', fmt:'#,##0.00', ref:ref401('7.6.3.1')});
  S.item({key:'VsReq', label:'箍筋需求 Vs', sym:'Vs', f:'MAX(0,{Vdes}/{phiv}-{Vc})', unit:'tf', fmt:'#,##0.00', ref:ref401('22.5.1.1')});
  S.item({key:'VsMax', label:'Vs 上限', sym:'Vs,max', f:'2.12*SQRT({fc})*{bw}*{d}/1000', unit:'tf', fmt:'#,##0.00', ref:ref401('22.5.1.2')});
  S.item({key:'Av', label:'Av', sym:'Av', f:'{nLegs}*{At}', unit:'cm²', fmt:'0.000', ref:'—'});
  S.item({key:'cc', label:'扭矩：箍筋中心至外緣', f:'{cover}+{dt}/2', unit:'cm', fmt:'0.00', ref:'—'});
  S.item({key:'Aoh', label:'Aoh', sym:'Aoh', f:'MAX(1,{bw}-2*{cc})*MAX(1,{h}-2*{cc})', unit:'cm²', fmt:'#,##0', ref:ref401('22.7.6.1')});
  S.item({key:'ph', label:'ph', sym:'ph', f:'2*(({bw}-2*{cc})+({h}-2*{cc}))', unit:'cm', fmt:'#,##0.0', ref:ref401('22.7.6.1')});
  S.item({key:'Tth', label:'可忽略扭矩門檻 φTth', sym:'φTth', f:'{phiv}*0.265*SQRT({fc})*({bw}*{h})^2/(2*({bw}+{h}))/100000', unit:'tf·m', fmt:'#,##0.00',
    expr:"φ·0.265√f'c·Acp²/pcp（Acp 保守僅取腹板）", ref:ref401('22.7.4.1')});
  S.item({key:'tors', label:'須設計扭矩', f:'--({Tu}>{Tth})', ref:ref401('22.7.4.1')});
  S.item({key:'AtS', label:'扭矩 At/s', sym:'At/s', f:'IF({tors}=1,{Tu}*100000/(2*{phiv}*0.85*{Aoh}*{fyt}/TAN({theta}*PI()/180)),0)', unit:'cm²/cm', fmt:'0.00000', ref:ref401('22.7.6.1')});
  S.item({key:'Al', label:'扭矩縱筋 Al', sym:'Al', f:'{AtS}*{ph}*{fyt}/{fy}/TAN({theta}*PI()/180)^2', unit:'cm²', fmt:'0.00', ref:ref401('22.7.6.1')});
  S.item({key:'over', label:'剪扭斷面應力超限', unit:'', f:'--(SQRT(({Vdes}*1000/({bw}*{d}))^2+IF({tors}=1,({Tu}*100000*{ph}/(1.7*{Aoh}^2))^2,0))>{phiv}*({Vc}*1000/({bw}*{d})+2.12*SQRT({fc})))',
    expr:'實心斷面平方根合成', crit:'0＝合格', ref:ref401('22.7.7.1')});
  S.item({key:'sStr', label:'剪扭強度需求間距', sym:'s', f:`IF(({VsReq}*1000/({fyt}*{d})+2*{AtS})>1E-9,{Av}/({VsReq}*1000/({fyt}*{d})+2*{AtS}),${BIG})`, unit:'cm', fmt:FMT_SINF, ref:ref401('22.5.10.5.3')});
  S.item({key:'sMin', label:'最小剪力鋼筋量間距', sym:'s', f:'MIN({Av}*{fyt}/(0.2*SQRT({fc})*{bw}),{Av}*{fyt}/(3.5*{bw}))', unit:'cm', fmt:'0.00', ref:ref401('9.6.3.3')});
  S.item({key:'sCode', label:'規範間距上限', sym:'s', f:'IF({VsReq}*1000>1.06*SQRT({fc})*{bw}*{d},MIN({d}/4,30),MIN({d}/2,60))', unit:'cm', fmt:'0.00', ref:ref401('9.7.6.2.2')});
  S.item({key:'sConf', label:'耐震加密區上限', sym:'s', f:`IF({isS}=1,MIN({d}/4,6*MIN({dbB},{dbT}),15),${BIG})`, unit:'cm', fmt:FMT_SINF, ref:ref401('18.6.4.4')});
  S.item({key:'sTors', label:'扭矩間距上限', sym:'s', f:`IF({tors}=1,MIN({ph}/8,30),${BIG})`, unit:'cm', fmt:FMT_SINF, ref:ref401('9.7.6.3.3')});
  S.item({key:'sGov', label:'控制需求間距（梁）', sym:'s', f:'MIN({sStr},{sMin},{sCode},{sConf},{sTors})', unit:'cm', fmt:'0.00', ref:'—'});
  S.item({key:'sUse', label:'採用箍筋間距（梁）', sym:'s', f:'IFERROR(_xlfn.AGGREGATE(14,6,{rng:sList}/({rng:sList}<={sGov}),1),MIN({rng:sList}))', unit:'cm', fmt:'0.0', ref:'非規範明列條文，係施工慣用間距'});
  S.item({key:'sSlab', label:'版主筋最大間距', sym:'s', f:'MIN(3*{h},45)', unit:'cm', fmt:'0.0', ref:ref401('7.7.2.3')});
  S.item({key:'sSlabUse', label:'版主筋採用間距（取大）', sym:'s', f:'MAX({spB},{spT})', unit:'cm', fmt:'0.0', ref:'—'});

  /* ---------------- 六、使用性 ---------------- */
  S.section('【六、使用性：裂縫控制、裂縫寬度與撓度】');
  S.item({key:'fs23', label:'裂縫控制 fs = ⅔fy', sym:'fs', f:'2/3*{fy}', unit:'kgf/cm²', fmt:'#,##0', ref:ref401('24.3.2.1')});
  S.item({key:'sLim', label:'鋼筋中心距上限', sym:'s', f:'MIN(38*(2855/{fs23})-2.5*{cover},30*(2855/{fs23}))', unit:'cm', fmt:'0.0', expr:'min(38(2855/fs) − 2.5cc, 30(2855/fs))', ref:ref401('24.3.2')});
  S.item({key:'sAct', label:'實際鋼筋中心距', sym:'s', f:'IF({isSlab}=1,{spB},IF({r1}>1,({bw}-2*({cover}+{dt})-{r1}*{dbB})/({r1}-1)+{dbB},{bw}))', unit:'cm', fmt:'0.0', ref:'—'});
  S.item({key:'MD', label:'使用靜載彎矩 MD', sym:'MD', v:inp.MD, unit:'tf·m', kind:'in', fmt:'0.00', ref:ref401('24.2')});
  S.item({key:'ML', label:'使用活載彎矩 ML', sym:'ML', v:inp.ML, unit:'tf·m', kind:'in', fmt:'0.00', ref:ref401('24.2')});
  S.item({key:'x', label:'開裂斷面中性軸 x', sym:'x', f:"'開裂斷面'!B3", unit:'cm', fmt:'0.000', kind:'link', expr:'轉換斷面 Q(x) = 0 二分求解', ref:ref401('24.2.3.5')});
  S.item({key:'Acx', label:'x 深度受壓面積', f:'IF(OR({hfc}=0,{x}<={hfc}),{b1c}*{x},{b1c}*{hfc}+{bw}*({x}-{hfc}))', unit:'cm²', fmt:'#,##0.0', ref:'—'});
  S.item({key:'Scx', label:'x 深度一次矩', f:'IF(OR({hfc}=0,{x}<={hfc}),{b1c}*{x}^2/2,{b1c}*{hfc}^2/2+{bw}*({x}-{hfc})*({hfc}+({x}-{hfc})/2))', unit:'cm³', fmt:'#,##0', ref:'—'});
  S.item({key:'I0x', label:'x 深度對受壓緣二次矩', f:'IF(OR({hfc}=0,{x}<={hfc}),{b1c}*{x}^3/3,{b1c}*{hfc}^3/3+{bw}*({x}^3-{hfc}^3)/3)', unit:'cm⁴', fmt:'#,##0', ref:'—'});
  const tr = dk => `IF(${dk}<{x},{n}-1,{n})`;
  S.item({key:'Icr', label:'開裂慣性矩 Icr', sym:'Icr', unit:'cm⁴', fmt:'#,##0',
    f:`{I0x}-2*{x}*{Scx}+{x}^2*{Acx}+${tr('{d1}')}*{A1}*({d1}-{x})^2+${tr('{d2}')}*{A2}*({d2}-{x})^2+${tr('{dT}')}*{AT}*({dT}-{x})^2`, ref:ref401('24.2.3.5')});
  S.item({key:'Ag', label:'全斷面積', sym:'Ag', f:'IF({isT}=1,{bw}*({h}-{hf})+{be}*{hf},{bw}*{h})', unit:'cm²', fmt:'#,##0', ref:'—'});
  S.item({key:'yb', label:'形心距底面', sym:'yb', f:'IF({isT}=1,({bw}*({h}-{hf})*({h}-{hf})/2+{be}*{hf}*({h}-{hf}/2))/{Ag},{h}/2)', unit:'cm', fmt:'0.00', ref:'—'});
  S.item({key:'Ig', label:'全斷面慣性矩 Ig', sym:'Ig', unit:'cm⁴', fmt:'#,##0',
    f:'IF({isT}=1,{bw}*({h}-{hf})^3/12+{bw}*({h}-{hf})*(({h}-{hf})/2-{yb})^2+{be}*{hf}^3/12+{be}*{hf}*({h}-{hf}/2-{yb})^2,{bw}*{h}^3/12)', ref:'—'});
  S.item({key:'Mcr', label:'開裂彎矩 Mcr', sym:'Mcr', f:'2*SQRT({fc})*{Ig}/IF({isPos}=1,{yb},{h}-{yb})', unit:'kgf·cm', fmt:'#,##0', expr:"fr·Ig/yt，fr = 2√f'c", ref:ref401('24.2.3.5')});
  S.item({key:'defl', label:'支承／載重型式', v:{simpleUDL:'簡支・均佈',simplePt:'簡支・中央集中',fixedUDL:'兩端固定・均佈',cantUDL:'懸臂・均佈'}[inp.deflCase]||'簡支・均佈', kind:'list',
    list:['簡支・均佈','簡支・中央集中','兩端固定・均佈','懸臂・均佈'], ref:'非規範明列條文，係彈性力學撓度公式', note:'Δ = K·Ma·L²/(Ec·Ie)：簡支均佈 5/48、簡支集中 1/12、固定均佈 1/16、懸臂均佈 1/4。'});
  S.item({key:'K', label:'撓度係數 K', sym:'K', f:'IF({defl}="簡支・均佈",5/48,IF({defl}="簡支・中央集中",1/12,IF({defl}="兩端固定・均佈",1/16,1/4)))', unit:'無因次', fmt:'0.0000', ref:'非規範明列條文，係彈性力學推導'});
  S.item({key:'sus', label:'持續活載比例', v:inp.sustain, unit:'無因次', kind:'in', fmt:'0.00', ref:ref401('24.2.4')});
  S.item({key:'xi', label:'長期係數 ξ', sym:'ξ', v:inp.xi, unit:'無因次', kind:'in', fmt:'0.0', crit:'5 年以上 2.0', ref:ref401('24.2.4.1.3')});
  S.item({key:'limL', label:'活載撓度限值 L/', v:inp.limLive, unit:'—', kind:'in', ref:ref401('24.2.2')});
  S.item({key:'limT', label:'總撓度限值 L/', v:inp.limTotal, unit:'—', kind:'in', ref:ref401('24.2.2')});
  // M 一律加括號代入，避免 Mcr/(MD+ML)*1e5 之運算順序錯誤
  const Ie = M => `IF((${M})<={Mcr},{Ig},MIN({Ig},({Mcr}/(${M}))^3*{Ig}+(1-({Mcr}/(${M}))^3)*{Icr}))`;
  const dl = (M,I) => `{K}*(${M})*{L}^2/({Ec}*(${I}))`;
  const MD='{MD}*100000', ML='{ML}*100000', Mt='({MD}+{ML})*100000', Ms='({MD}+{sus}*{ML})*100000';
  S.item({key:'dL', label:'即時活載撓度 ΔL', sym:'ΔL', unit:'cm', fmt:'0.000',
    f:`MAX(0,${dl(Mt,Ie(Mt))}-${dl(MD,Ie(Mt))})`, expr:'Δ(MD+ML) − Δ(MD)，皆以 Ie(MD+ML)', ref:ref401('24.2.3')});
  S.item({key:'rhoP', label:"受壓鋼筋比 ρ'", sym:"ρ'", f:'(({d1}<{h}/2)*{A1}+({d2}<{h}/2)*{A2}+({dT}<{h}/2)*{AT})/({bw}*{d})', unit:'無因次', fmt:'0.00000', ref:ref401('24.2.4.1.1')});
  S.item({key:'lam', label:'長期乘數 λΔ', sym:'λΔ', f:"{xi}/(1+50*{rhoP})", unit:'無因次', fmt:'0.000', ref:ref401('24.2.4.1.1')});
  S.item({key:'dLT', label:'長期總撓度 ΔLT', sym:'ΔLT', unit:'cm', fmt:'0.000', f:`${dl(Ms,Ie(Ms))}*{lam}+{dL}`, expr:'Δsus·λΔ + ΔL', ref:ref401('24.2.4')});
  S.item({key:'Ms', label:'裂縫寬度：使用彎矩 Ms', sym:'Ms', f:'{MD}+{ML}', unit:'tf·m', fmt:'0.00', ref:'—'});
  S.item({key:'fsS', label:'裂縫寬度：使用鋼筋應力 fs', sym:'fs', f:'IF({d}>{x},{n}*{Ms}*100000*({d}-{x})/{Icr},0)', unit:'kgf/cm²', fmt:'#,##0', ref:'非規範明列條文，係開裂轉換斷面彈性分析'});
  S.item({key:'dc', label:'裂縫寬度：dc', sym:'dc', f:'{cover}+{dt}+IF({isPos}=1,{dbB},{dbT})/2', unit:'cm', fmt:'0.00', ref:'—'});
  S.item({key:'pit', label:'裂縫寬度：受拉筋間距', f:'IF({isSlab}=1,IF({isPos}=1,{spB},IF({spT}>0,{spT},{spB})),{bw}/MAX(1,IF({isPos}=1,{r1},{nTop})))', unit:'cm', fmt:'0.00', ref:'—'});
  S.item({key:'beta', label:'裂縫寬度：β', sym:'β', f:'({h}-{x})/({d}-{x})', unit:'無因次', fmt:'0.000', ref:'—'});
  S.item({key:'w', label:'裂縫寬度 w', sym:'w', unit:'mm', fmt:'0.000',
    f:'1.1E-5*{beta}*{fsS}*0.0980665*(({dc}*10)*(2*{dc}*{pit}*100))^(1/3)', expr:'w = 1.1×10⁻⁵ β fs ∛(dc·A)（SI：MPa、mm）', ref:'非我國規範明列；係 Gergely-Lutz 式（ACI 224R）'});
  S.item({key:'wLim', label:'容許裂縫寬度', sym:'wlim', v:+inp.wLim||0.2, unit:'mm', kind:'list', list:['0.10','0.15','0.20','0.25','0.30'], fmt:'0.00',
    crit:'關鍵假設：水密要求嚴格 0.10；一般戶外 0.15～0.30', ref:'非我國規範明列；參考 ACI 224R 表 4.1', note:'關鍵假設。ACI 224R 建議水密結構 0.10 mm；請依設計準則選定。'});

  /* ---------------- 七、檢核彙總 ---------------- */
  S.section('【七、檢核彙總】');
  S.sum({head:true, label:'檢核項目', need:'需求值', cap:'容量／限值', ratio:'比值', judge:'判定'});
  const sumRows=[]; const addSum=o=>sumRows.push(S.sum(o));
  const J = c => ({f:`IF(${c},"PASS","FAIL")`});
  addSum({label:'撓曲 Mu／φMn（tf·m）', need:{f:'{MuMax}'}, cap:{f:'{phiMn}'}, ratio:{f:'{DC}'}, judge:J('{DC}<=1'), ref:ref401('9.5.1.1')});
  addSum({label:'受拉鋼筋 As ≧ As,min（cm²）', need:{f:'{AsMin}'}, cap:{f:'{AsT}'}, ratio:{f:'{AsMin}/{AsT}'}, judge:{f:'IF({MuMax}<1E-6,"N/A",IF({AsT}>={AsMin},"PASS","FAIL"))'}, note:'Mu = 0 時不適用', ref:ref401('9.6.1.2、7.6.1.1')});
  addSum({label:'εt ≧ 0.004', need:0.004, cap:{f:'{et}'}, ratio:{f:'0.004/{et}'}, judge:J('{et}>=0.004'), fmt:'0.00000', ref:ref401('9.3.3.1')});
  addSum({label:'梁剪力：Vs ≦ Vs,max（tf）', need:{f:'{VsReq}'}, cap:{f:'{VsMax}'}, ratio:{f:'{VsReq}/{VsMax}'}, judge:{f:'IF({isSlab}=1,"N/A",IF(AND({VsReq}<={VsMax},{over}=0),"PASS","FAIL"))'}, ref:ref401('22.5.1.2')});
  addSum({label:'梁箍筋間距 採用／需求（cm）', need:{f:'{sUse}'}, cap:{f:'{sGov}'}, ratio:{f:'{sUse}/{sGov}'}, judge:{f:'IF({isSlab}=1,"N/A",IF({sUse}<={sGov},"PASS","FAIL"))'}, ref:'—'});
  addSum({label:'版剪力：Vu ≦ φVc（tf）', need:{f:'{Vdes}'}, cap:{f:'{phiVc}'}, ratio:{f:'{Vdes}/{phiVc}'}, judge:{f:'IF({isSlab}=0,"N/A",IF({Vdes}<={phiVc},"PASS","FAIL"))'}, ref:ref401('7.6.3.1')});
  addSum({label:'版主筋間距（cm）', need:{f:'{sSlabUse}'}, cap:{f:'{sSlab}'}, ratio:{f:'{sSlabUse}/{sSlab}'}, judge:{f:'IF({isSlab}=0,"N/A",IF({sSlabUse}<={sSlab},"PASS","FAIL"))'}, ref:ref401('7.7.2.3')});
  addSum({label:'裂縫控制 鋼筋中心距（cm）', need:{f:'{sAct}'}, cap:{f:'{sLim}'}, ratio:{f:'{sAct}/{sLim}'}, judge:J('{sAct}<={sLim}'), ref:ref401('24.3.2')});
  addSum({label:'水工裂縫寬度（mm）', need:{f:'{w}'}, cap:{f:'{wLim}'}, ratio:{f:'{w}/{wLim}'}, judge:{f:'IF({isW}=0,"N/A",IF({w}<={wLim},"PASS","FAIL"))'}, fmt:'0.000', ref:'ACI 224R（參考）'});
  addSum({label:'水工保護層（cm）', need:5, cap:{f:'{cover}'}, ratio:{f:'5/{cover}'}, judge:{f:'IF({isW}=0,"N/A",IF({cover}>=5,"PASS","FAIL"))'}, ref:'ACI 350（參考）'});
  addSum({label:'即時活載撓度（cm）', need:{f:'{dL}'}, cap:{f:'{L}/{limL}'}, ratio:{f:'{dL}/({L}/{limL})'}, judge:J('{dL}<={L}/{limL}'), fmt:'0.000', ref:ref401('24.2.2')});
  addSum({label:'長期總撓度（cm）', need:{f:'{dLT}'}, cap:{f:'{L}/{limT}'}, ratio:{f:'{dLT}/({L}/{limT})'}, judge:J('{dLT}<={L}/{limT}'), fmt:'0.000', ref:ref401('24.2.2')});
  addSum({label:'耐震：bw ≧ max(0.3h, 25)（cm）', need:{f:'MAX(0.3*{h},25)'}, cap:{f:'{bw}'}, ratio:{f:'MAX(0.3*{h},25)/{bw}'}, judge:{f:'IF({isS}=0,"N/A",IF({bw}>=MAX(0.3*{h},25),"PASS","FAIL"))'}, ref:ref401('18.6.2.1')});
  addSum({label:'耐震：ln ≧ 4d（cm）', need:{f:'4*{d}'}, cap:{f:'{ln}'}, ratio:{f:'4*{d}/{ln}'}, judge:{f:'IF({isS}=0,"N/A",IF({ln}>=4*{d},"PASS","FAIL"))'}, ref:ref401('18.6.2.1')});
  addSum({label:'耐震：ρ ≦ 0.025', need:{f:'{rho}'}, cap:0.025, ratio:{f:'{rho}/0.025'}, judge:{f:'IF({isS}=0,"N/A",IF({rho}<=0.025,"PASS","FAIL"))'}, fmt:'0.0000', ref:ref401('18.6.3.1')});
  addSum({label:'耐震：φMn⁺ ≧ 0.5φMn⁻', need:{f:'0.5*{Mneg}'}, cap:{f:'{Mpos}'}, ratio:{f:'0.5*{Mneg}/{Mpos}'}, judge:{f:'IF({isS}=0,"N/A",IF({Mpos}>=0.5*{Mneg},"PASS","FAIL"))'}, ref:ref401('18.6.3.2')});
  S.sum({key:'jAll', total:true, label:'總判定', judge:{f:'"PASS"'}, note:'僅計 FAIL；N/A 不計', ref:'—'});

  S.section('【八、注意事項與使用說明】');
  S.text('色碼圖例：淺藍底＋藍字粗體＋藍框＝手動輸入格；黃底＋藍字粗體＋金框＝附下拉選單之輸入格（可輸入表列外之值）；白底黑字細灰框＝公式；綠字＝連結其他工作表；淺琥珀底＝總判定。',{h:32});
  S.text('1. 中性軸以 ΣF = 0 二分 '+NBI+' 次求解（「撓曲求解」），開裂中性軸以轉換斷面 Q(x) = 0 二分求解（「開裂斷面」）。底筋最多 2 排、頂筋單排。',{h:28});
  S.text("2. Vc 採 0.53√f'c·bw·d；未納入土木401-112（ACI 318-19）§22.5.5.1 之 ρw 與尺寸效應 λs。版不配剪力筋，須 Vu ≦ φVc。",{h:26});
  S.text('3. 水工加嚴之裂縫寬度（Gergely-Lutz）與保護層、ρmin 均參考 ACI 350／ACI 224R，非我國規範明列；未含 ACI 350 環境耐久係數 Sd。「任一斷面 Mn ≧ 端部 25%」需整根梁包絡線，本表未檢核。建議配筋請於網頁查看。',{h:32});

  S.section('【附錄、對照表區】');
  S.table({cols:[{h:'鋼筋號數',key:'barName',input:true},{h:'直徑 (cm)',key:'barD',input:true,fmt:'0.000'},{h:'面積 (cm²)',key:'barA',input:true,fmt:'0.000'}], data:BARS.map(b=>[b[0],b[1],b[2]])});
  S.blank();
  S.table({cols:[{h:'實務箍筋間距 (cm)',key:'sList',input:true,fmt:'0.0'}], data:PRACTICAL_S.map(v=>[v])});
  S.layout();
  const r1=sumRows[0].r, r2=sumRows[sumRows.length-1].r;
  const jt=S.rows.find(r=>r.key==='jAll'); jt.judge={f:`IF(COUNTIF($E$${r1}:$E$${r2},"FAIL")=0,"PASS","NG")`};

  const R = makeResolver([S]);
  const order=['檢核表','結構計算書(A4)','載重組合','撓曲求解','開裂斷面','圖表資料'];
  const W={}; for(const nm of order) W[nm]=wb.addWorksheet(nm, nm==='檢核表'?{views:[{state:'frozen',ySplit:2}]}:{});
  writeCalcSheet(W['檢核表'], S, R);
  buildBeamLoads(W['載重組合'], R, inp.loads);
  buildBeamSolve(W['撓曲求解'], R);
  buildBeamIcr(W['開裂斷面'], R);
  const binfo = buildBeamChartData(W['圖表資料'], R);
  const a4Start = buildA4Beam(W['結構計算書(A4)'], R, inp, jt.r, sumRows);
  return {wb, keys:S.keys, judgeRow:jt.r, charts: beamCharts(inp, binfo, a4Start)};
}

function buildBeamLoads(ws, R, loads){
  const nm='載重組合', r=f=>({f:R(f,nm)});
  ws.columns=[24,12,12,12,12,4,14,12,12].map(w=>({width:w}));
  put(ws,'A1','梁／版載重組合（已乘載重因數；每公尺寬或每梁）',{sec:true});
  ['組合名稱','Mu⁺ (tf·m)','Mu⁻ (tf·m)','Vu (tf)','Tu (tf·m)','','檢核用 Mu (tf·m)','|Vu| (tf)','D/C'].forEach((h,j)=>{ if(h) put(ws,colL(j)+'3',h,{head:true}); });
  for(let i=0;i<NCB;i++){
    const n=4+i, L=loads[i];
    put(ws,'A'+n, L?L.name:'', {input:true});
    ['Mpos','Mneg','Vu','Tu'].forEach((k,j)=>put(ws,colL(1+j)+n, L?(+L[k]||0):null, {input:true, fmt:'#,##0.00'}));
    put(ws,'G'+n, r(`IF(A${n}="","",IF({isPos}=1,ABS(B${n}),ABS(C${n})))`),{fmt:'#,##0.00'});
    put(ws,'H'+n, {f:`IF(A${n}="","",ABS(D${n}))`},{fmt:'#,##0.00'});
    put(ws,'I'+n, r(`IF(A${n}="","",IF({phiMn}>0,G${n}/{phiMn},1E9))`),{fmt:FMT_INF});
  }
}

/* 撓曲求解：四個二分區塊（正／負彎矩 × fy／1.25fy），每區塊一欄 */
function buildBeamSolve(ws, R){
  const nm='撓曲求解', r=f=>({f:R(f,nm)});
  ws.columns=[22,12,12,12,12,12].map(w=>({width:w}));
  put(ws,'A1','中性軸二分求解：ΣF(c) = Cc + Σfs·As = 0',{sec:true});
  // 摘要列：B3 = 目前檢核斷面之 c；F3~F6 = φMn⁺、φMn⁻、Mpr⁺、Mpr⁻
  put(ws,'A2','項目',{head:true}); put(ws,'B2','c (cm)',{head:true}); put(ws,'E2','項目',{head:true}); put(ws,'F2','值 (tf·m)',{head:true});
  const blocks = [['正彎矩 fy','pos','{fy}'],['負彎矩 fy','neg','{fy}'],['正彎矩 1.25fy','pos','1.25*{fy}'],['負彎矩 1.25fy','neg','1.25*{fy}']];
  const B0 = 12, rowsPer = 4;          // 每次迭代 4 列：lo、hi、mid、F(mid)
  blocks.forEach((b,i)=>{ put(ws,colL(1+i)+(B0-1), b[0], {head:true}); });
  const geo = (sg) => ({
    b1: sg==='pos' ? 'IF({isT}=1,{be},{bw})' : '{bw}', hf: sg==='pos' ? 'IF({isT}=1,{hf},0)' : '0',
    d1: sg==='pos' ? '({h}-{y1})' : '{y1}', d2: sg==='pos' ? '({h}-{y2})' : '{y2}', dT: sg==='pos' ? '({h}-{yT})' : '{yT}'
  });
  const Fexpr = (c, sg, fy) => {
    const g=geo(sg), a=`MIN({b1}*${c},{h})`;
    const Ac = `IF(OR(${g.hf}=0,${a}<=${g.hf}),${g.b1}*${a},${g.b1}*${g.hf}+{bw}*(${a}-${g.hf}))`;
    const fs = dk => `(MAX(-${fy},MIN(${fy},{Es}*{ecu}*(${c}-${dk})/${c}))-IF(${dk}<=${a},0.85*{fc},0))`;
    return `0.85*{fc}*${Ac}+${fs(g.d1)}*{A1}+${fs(g.d2)}*{A2}+${fs(g.dT)}*{AT}`;
  };
  const Mexpr = (c, sg, fy) => {
    const g=geo(sg), a=`MIN({b1}*${c},{h})`;
    const Sc = `IF(OR(${g.hf}=0,${a}<=${g.hf}),${g.b1}*${a}^2/2,${g.b1}*${g.hf}^2/2+{bw}*(${a}-${g.hf})*(${g.hf}+(${a}-${g.hf})/2))`;
    const fs = dk => `(MAX(-${fy},MIN(${fy},{Es}*{ecu}*(${c}-${dk})/${c}))-IF(${dk}<=${a},0.85*{fc},0))`;
    return `(-0.85*{fc}*${Sc}-${fs(g.d1)}*{A1}*${g.d1}-${fs(g.d2)}*{A2}*${g.d2}-${fs(g.dT)}*{AT}*${g.dT})`;
  };
  const cFinal = [];
  blocks.forEach((b,i)=>{
    const col=colL(1+i), sg=b[1], fy=b[2];
    for(let j=0;j<NBI;j++){
      const rr=B0+rowsPer*j;
      if(i===0){ put(ws,'A'+rr,`j=${j+1} lo`); put(ws,'A'+(rr+1),'hi'); put(ws,'A'+(rr+2),'mid'); put(ws,'A'+(rr+3),'ΣF(mid)'); }
      if(j===0){ put(ws,col+rr, r('0.0001*{h}'),{fmt:'0.000000'}); put(ws,col+(rr+1), r('5*{h}'),{fmt:'0.000000'}); }
      else{ const p=rr-rowsPer;
        put(ws,col+rr,{f:`IF(${col}${p+3}>0,${col}${p},${col}${p+2})`},{fmt:'0.000000'});
        put(ws,col+(rr+1),{f:`IF(${col}${p+3}>0,${col}${p+2},${col}${p+1})`},{fmt:'0.000000'}); }
      put(ws,col+(rr+2),{f:`(${col}${rr}+${col}${rr+1})/2`},{fmt:'0.000000'});
      put(ws,col+(rr+3), r(Fexpr(`${col}${rr+2}`, sg, fy)),{fmt:'#,##0.00'});
    }
    const last=B0+rowsPer*(NBI-1);
    cFinal.push(`(${col}${last}+${col}${last+1})/2`);
    put(ws,col+'8', {f:cFinal[i]},{fmt:'0.0000'});
    put(ws,col+'9', r(Mexpr(`${col}8`, sg, fy)+'/100000'),{fmt:'#,##0.00'});
    // φ（僅 fy 區塊需要）
    const g=geo(sg), dmax=`MAX(${g.d1},IF({r2}>0,${g.d2},-1E9),${g.dT})`, et=`{ecu}*(${dmax}-${col}8)/${col}8`;
    put(ws,col+'10', r(`IF(${et}<={ety},{phic},IF(${et}>=0.005,{phit},{phic}+({phit}-{phic})*(${et}-{ety})/(0.005-{ety})))`),{fmt:'0.000'});
  });
  put(ws,'A8','c 收斂值 (cm)',{head:true}); put(ws,'A9','Mn (tf·m)',{head:true}); put(ws,'A10','φ',{head:true});
  put(ws,'A3','目前檢核斷面');
  put(ws,'B3', r('IF({isPos}=1,B8,C8)'),{fmt:'0.0000'});
  put(ws,'E3','φMn⁺'); put(ws,'F3',{f:'B10*B9'},{fmt:'#,##0.00'});
  put(ws,'E4','φMn⁻'); put(ws,'F4',{f:'C10*C9'},{fmt:'#,##0.00'});
  put(ws,'E5','Mpr⁺'); put(ws,'F5',{f:'D9'},{fmt:'#,##0.00'});
  put(ws,'E6','Mpr⁻'); put(ws,'F6',{f:'E9'},{fmt:'#,##0.00'});
}

/* 開裂斷面：Q(x) = A(x)·x − S(x) + Σ… = 0 二分 */
function buildBeamIcr(ws, R){
  const nm='開裂斷面', r=f=>({f:R(f,nm)});
  ws.columns=[22,14].map(w=>({width:w}));
  put(ws,'A1','開裂轉換斷面中性軸 x：Q(x) = 0',{sec:true});
  put(ws,'A3','x 收斂值 (cm)',{head:true});
  const Q = x => {
    const A=`IF(OR({hfc}=0,${x}<={hfc}),{b1c}*${x},{b1c}*{hfc}+{bw}*(${x}-{hfc}))`;
    const S=`IF(OR({hfc}=0,${x}<={hfc}),{b1c}*${x}^2/2,{b1c}*{hfc}^2/2+{bw}*(${x}-{hfc})*({hfc}+(${x}-{hfc})/2))`;
    const st = (dk,Ak) => `IF(${dk}<${x},({n}-1)*${Ak}*(${x}-${dk}),-{n}*${Ak}*(${dk}-${x}))`;
    return `${A}*${x}-${S}+${st('{d1}','{A1}')}+${st('{d2}','{A2}')}+${st('{dT}','{AT}')}`;
  };
  const B0=6;
  for(let j=0;j<NBI;j++){
    const rr=B0+4*j;
    put(ws,'A'+rr,`j=${j+1} lo`); put(ws,'A'+(rr+1),'hi'); put(ws,'A'+(rr+2),'mid'); put(ws,'A'+(rr+3),'Q(mid)');
    if(j===0){ put(ws,'B'+rr, r('1E-6*{h}'),{fmt:'0.000000'}); put(ws,'B'+(rr+1), r('{h}'),{fmt:'0.000000'}); }
    else{ const p=rr-4; put(ws,'B'+rr,{f:`IF(B${p+3}>0,B${p},B${p+2})`},{fmt:'0.000000'}); put(ws,'B'+(rr+1),{f:`IF(B${p+3}>0,B${p+2},B${p+1})`},{fmt:'0.000000'}); }
    put(ws,'B'+(rr+2),{f:`(B${rr}+B${rr+1})/2`},{fmt:'0.000000'});
    put(ws,'B'+(rr+3), r(Q(`B${rr+2}`)),{fmt:'#,##0.00'});
  }
  const last=B0+4*(NBI-1);
  put(ws,'B3', r(`IF(${Q('{h}')}<0,{h},(B${last}+B${last+1})/2)`),{fmt:'0.0000'});
}

function buildA4Beam(ws, R, inp, jRow, sumRows){
  a4Base(ws);
  const a=new A4(ws,R);
  const t=a.row(); ws.mergeCells(t,2,t,6); const tc=ws.getCell(t,2);
  tc.value={formula:R('IF({isSlab}=1,"RC 版（單位寬）設計檢核計算書","RC 梁斷面設計檢核計算書")',ws.name)};
  tc.font={name:FONT,size:14,bold:true}; tc.alignment={horizontal:'center',vertical:'middle'}; ws.getRow(t).height=26;
  const t2=a.row(); ws.mergeCells(t2,2,t2,6); const tc2=ws.getCell(t2,2); tc2.value='單位：kgf、cm（力顯示 tf、彎矩 tf·m）'; tc2.font={name:FONT,size:9}; tc2.alignment={horizontal:'center'};
  ws.pageSetup.printTitlesRow=`${t}:${t2}`;
  a.input('工程名稱', inp.meta&&inp.meta.project||''); a.input('構件名稱／編號', inp.meta&&inp.meta.member||'');
  a.input('設計者／檢核者',''); a.input('計算日期／版次', new Date().toLocaleDateString('zh-TW')+'／第 1 版');
  a.chap('一、設計依據');
  a.para('1.1','混凝土結構設計規範（土木401），內政部國土管理署；撓曲、剪力、扭矩、裂縫控制與撓度依該規範（以 ACI 318 為基礎）。',28);
  a.para('1.2',{f:'IF({isS}=1,"建築物耐震設計規範及解說（耐震特殊抗彎構材）。","非耐震特殊抗彎構材。")'},18);
  a.para('1.3',{f:'IF({isW}=1,"水工／環境結構加嚴：裂縫寬度、保護層與最小鋼筋比參考 ACI 350、ACI 224R（我國規範未明列）。","一般環境。")'},26);
  a.para('1.4','本計算書數值均連結「檢核表」工作表，修改輸入後自動更新。',18);
  a.chap('二、設計條件');
  a.sub('2.1 幾何形狀'); a.thead();
  a.data('斷面型式','—','{type}','—',null,null,true);
  a.data('寬 × 深','b × h','TEXT({bw},"0")&" × "&TEXT({h},"0")','cm','版為每公尺寬',null,true);
  a.data('有效翼緣寬','be','{be}','cm','土木401 §6.3.2.1','0.0');
  a.data('保護層','cc','{cover}','cm','土木401 §20.5.1.3','0.0');
  a.sub('2.2 材料與強度折減因數'); a.thead();
  a.data("混凝土抗壓強度","f'c",'{fc}','kgf/cm²','土木401 §19.2','#,##0');
  a.data('主筋／箍筋降伏強度','fy／fyt','TEXT({fy},"#,##0")&"／"&TEXT({fyt},"#,##0")','kgf/cm²','CNS 560',null,true);
  a.data('強度折減因數','φt／φv','TEXT({phit},"0.00")&"／"&TEXT({phiv},"0.00")','—','土木401 §21.2',null,true);
  a.sub('2.3 載重條件'); a.thead();
  a.data('檢核斷面','—','{sign}','—',null,null,true);
  a.data('需求彎矩','Mu','{MuMax}','tf·m','各組合最大','#,##0.00');
  a.data('需求剪力','Vu','{Vu}','tf','各組合最大','#,##0.00');
  a.data('使用彎矩','MD + ML','{Ms}','tf·m','撓度與裂縫寬度用','#,##0.00');
  a.sub('2.4 配筋條件'); a.thead();
  a.data('底筋','—','IF({isSlab}=1,{botS}&" @ "&TEXT({spB},"0.0")&" cm",ROUND({nBot},0)&"-"&{botS})','—',null,null,true);
  a.data('頂筋','—','IF({isSlab}=1,IF({spT}>0,{topS}&" @ "&TEXT({spT},"0.0")&" cm","無"),ROUND({nTop},0)&"-"&{topS})','—',null,null,true);
  a.data('有效深度','d','{d}','cm','受拉鋼筋群形心','0.0');
  a.chap('三、設計假設');
  a.para('3.1','撓曲採應變相容與等值矩形應力塊（土木401 §22.2），中性軸以 ΣF = 0 求解；T 梁負彎矩時翼緣受拉不計入受壓區。',28);
  a.para('3.2',"剪力 Vc = 0.53√f'c·bw·d，未納入土木401-112 §22.5.5.1 之尺寸效應 λs；版不配剪力筋，須 Vu ≦ φVc。",26);
  a.para('3.3','撓度採 Branson 有效慣性矩 Ie 與長期乘數 λΔ = ξ/(1 + 50ρ\')；裂縫控制依土木401 §24.3.2（fs = ⅔fy）。',26);
  ws.getRow(a.n).addPageBreak();
  a.chap('四、計算過程');
  a.sub('4.1 撓曲（土木401 §22.2、§22.3）'); a.thead();
  a.data('中性軸深度','c','{c}','cm','ΣF = 0','0.00');
  a.data('標稱彎矩強度','Mn','{Mn}','tf·m','對受壓緣取矩','#,##0.00');
  a.data('最外受拉筋應變','εt','{et}','—','≧ 0.004','0.00000');
  a.data('設計彎矩強度','φMn','{phiMn}','tf·m','—','#,##0.00');
  a.data('撓曲應力比','D/C','{DC}','—','Mu/φMn',FMT_INF);
  a.sub('4.2 最小鋼筋（土木401 §9.6.1.2、§7.6.1.1）'); a.thead();
  a.data('受拉鋼筋','As','{AsT}','cm²','—','0.00');
  a.data('最少鋼筋','As,min','{AsMin}','cm²','梁 max(0.8√f\'c/fy, 14/fy)·bw·d；版 ρmin·b·h','0.00');
  a.sub('4.3 剪力與扭矩（土木401 §22.5、§22.7）'); a.thead();
  a.data('設計剪力','Vdes','{Vdes}','tf','耐震取 max(Ve, Vu)','#,##0.00');
  a.data('混凝土剪力強度','Vc','{Vc}','tf',"0.53√f'c·bw·d",'#,##0.00');
  a.data('箍筋採用間距（梁）','s','IF({isSlab}=1,"不配剪力筋",{tieS}&" "&{nLegs}&" 肢 @ "&TEXT({sUse},"0.0")&" cm")','—',null,null,true);
  a.data('可忽略扭矩門檻','φTth','{Tth}','tf·m',"φ·0.265√f'c·Acp²/pcp",'#,##0.00');
  a.sub('4.4 使用性（土木401 §24.2、§24.3）'); a.thead();
  a.data('開裂慣性矩','Icr','{Icr}','cm⁴','轉換斷面','#,##0');
  a.data('即時活載撓度','ΔL','{dL}','cm','—','0.000');
  a.data('長期總撓度','ΔLT','{dLT}','cm','—','0.000');
  a.data('裂縫寬度（水工）','w','{w}','mm','Gergely-Lutz（ACI 224R）','0.000');
  ws.getRow(a.n).addPageBreak();
  a.chap('五、檢核彙總');
  const hn=a.row(); ['檢核項目','需求值','容量／限值','比值','判定'].forEach((h,j)=>{ const c=ws.getCell(hn,2+j); c.value=h; c.font={name:FONT,size:9,bold:true}; c.fill={type:'pattern',pattern:'solid',fgColor:{argb:K.hdr}}; c.border=box(K.grid); c.alignment={horizontal:'center'}; });
  for(const sr of sumRows){ const n=a.row(); ['A','B','C','D','E'].forEach((col,j)=>{ const c=ws.getCell(n,2+j); c.value={formula:`'檢核表'!${col}${sr.r}`}; c.font={name:FONT,size:9,color:{argb:K.link}}; c.border=box(K.grid); c.alignment={horizontal:j?'center':'left',wrapText:true}; if(j>0&&j<4) c.numFmt=j===3?FMT_INF:(sr.fmt||FMT_SINF); }); }
  const tn=a.row(); ws.mergeCells(tn,2,tn,5); ws.getCell(tn,2).value='整體結構判定';
  [2,3,4,5,6].forEach(j=>{ const c=ws.getCell(tn,j); c.fill={type:'pattern',pattern:'solid',fgColor:{argb:K.amber}}; c.border=box(K.grid); c.font={name:FONT,size:10,bold:true}; });
  const tj=ws.getCell(tn,6); tj.value={formula:`'檢核表'!E${jRow}`}; tj.alignment={horizontal:'center'};
  a.chap('六、結論');
  a.para('6.1',{f:'{type}&" b × h = "&TEXT({bw},"0")&" × "&TEXT({h},"0")&" cm；撓曲 D/C = "&TEXT({DC},"0.000")&"；總判定 "&\'檢核表\'!E'+jRow},28);
  a.para('6.2',{f:'IF(\'檢核表\'!E'+jRow+'="PASS","本斷面各項檢核均符合規定，可供施工圖說使用。","本斷面有檢核項目不符規定，應調整斷面或配筋後重新計算。")'},22);
  a.chap('七、限制與注意事項');
  a.para('7.1','未檢核：整根梁彎矩包絡線（任一斷面 Mn ≧ 端部 25%）、鋼筋延伸與搭接、ACI 350 環境耐久係數 Sd、剪力尺寸效應 λs。',28);
  a.para('7.2','輸入資料（尺寸、材料、載重、使用彎矩）須經設計者核實；版次修改應更新封面版次。',22);
  a.para('7.3','色碼：淺藍底藍框＝輸入；黃底金框＝下拉輸入；白底細灰框＝公式；綠字＝連結；淺琥珀底＝總判定。',26);
  a.blank();
  const sg=a.row(), sl=a.row(); ws.getRow(sl).height=45;
  ws.mergeCells(sg,3,sg,4); ws.mergeCells(sg,5,sg,6); ws.mergeCells(sl,3,sl,4); ws.mergeCells(sl,5,sl,6);
  [[2,'設計'],[3,'校核'],[5,'審核']].forEach(([j,h])=>{ const c=ws.getCell(sg,j); c.value=h; c.font={name:FONT,size:9,bold:true}; c.alignment={horizontal:'center'}; ws.getCell(sl,j).border={bottom:side(K.black)}; });
  const fig = a4Figures(ws, a, '附圖　斷面圖（Excel 圖表，隨輸入自動更新）', 0);
  ws.pageSetup.printArea=`A1:F${a.n}`;
  return fig;
}

/* ======================================================================
   Excel 原生圖表（散佈圖）
   ExcelJS 不支援建立圖表，故活頁簿寫出後再以 JSZip 注入 DrawingML 圖表：
     xl/charts/chartN.xml、xl/drawings/drawingK.xml（+rels）、工作表 <drawing>、[Content_Types]。
   數列一律以 numRef 參照「圖表資料」等工作表的公式格 → 在 Excel 內改輸入，圖會跟著重畫。
   #N/A 點在 Excel 365 以「顯示為空白」處理（dispNaAsBlank），用來分隔繫筋線段。
   ====================================================================== */
const NA = 'NA()';
const xmlEsc = s => String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
const qs = n => `'${n.replace(/'/g,"''")}'`;

/* 散佈圖 XML。series：{name, x:'工作表!$A$3:$A$9', y:..., line:{color,w,dash}|null, marker:{symbol,size,color,fill}|null} */
function scatterChartXml(o){
  const drop = new Set(o.legendDel||[]);
  o = Object.assign({}, o, {series: o.series.filter((_,i)=>!drop.has(i)), legendDel: []});
  const ser = o.series.map((s,i)=>{
    const ln = s.line
      ? `<a:ln w="${Math.round((s.line.w||1.5)*12700)}" cap="rnd"><a:solidFill><a:srgbClr val="${s.line.color}"/></a:solidFill>${s.line.dash?`<a:prstDash val="${s.line.dash}"/>`:''}<a:round/></a:ln>`
      : `<a:ln w="12700"><a:noFill/></a:ln>`;
    const mk = s.marker
      ? `<c:marker><c:symbol val="${s.marker.symbol||'circle'}"/><c:size val="${Math.max(2,Math.min(72,Math.round(s.marker.size||5)))}"/><c:spPr>${s.marker.fill===false?'<a:noFill/>':`<a:solidFill><a:srgbClr val="${s.marker.fill||s.marker.color}"/></a:solidFill>`}<a:ln w="9525"><a:solidFill><a:srgbClr val="${s.marker.color}"/></a:solidFill></a:ln></c:spPr></c:marker>`
      : `<c:marker><c:symbol val="none"/></c:marker>`;
    return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:v>${xmlEsc(s.name)}</c:v></c:tx>`
      + `<c:spPr>${ln}</c:spPr>${mk}`
      + `<c:xVal><c:numRef><c:f>${xmlEsc(s.x)}</c:f></c:numRef></c:xVal>`
      + `<c:yVal><c:numRef><c:f>${xmlEsc(s.y)}</c:f></c:numRef></c:yVal><c:smooth val="0"/></c:ser>`;
  }).join('');
  const axis = (id, cross, pos, title, a) => {
    const sc = (a && a.min!==undefined) ? `<c:max val="${a.max}"/><c:min val="${a.min}"/>` : '';
    const del = a && a.hidden ? 1 : 0;
    return `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/>${sc}</c:scaling><c:delete val="${del}"/>`
      + `<c:axPos val="${pos}"/>${a&&(a.grid===false||a.hidden)?'':'<c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="E2E7EC"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>'}`
      + (title?`<c:title><c:tx><c:rich><a:bodyPr${pos==='l'?' rot="-5400000" vert="horz"':''}/><a:p><a:pPr><a:defRPr sz="900" b="0"/></a:pPr><a:r><a:rPr lang="zh-TW" sz="900" b="0"/><a:t>${xmlEsc(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`:'')
      + `<c:numFmt formatCode="${a&&a.fmt||'#,##0'}" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="low"/>`
      + `<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="3A4652"/></a:solidFill></a:ln></c:spPr>`
      + `<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="800"/></a:pPr><a:endParaRPr lang="zh-TW"/></a:p></c:txPr>`
      + `<c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`;
  };
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
    + `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart">`
    + `<c:roundedCorners val="0"/><c:chart>`
    + (o.titleRef
        ? `<c:title><c:tx><c:strRef><c:f>${xmlEsc(o.titleRef)}</c:f></c:strRef></c:tx><c:overlay val="0"/><c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1100" b="1"/></a:pPr><a:endParaRPr lang="zh-TW"/></a:p></c:txPr></c:title>`
        : `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1100" b="1"/></a:pPr><a:r><a:rPr lang="zh-TW" sz="1100" b="1"/><a:t>${xmlEsc(o.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`)
    + `<c:autoTitleDeleted val="0"/><c:plotArea>`
    + (o.plot ? `<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="${o.plot.x}"/><c:y val="${o.plot.y}"/><c:w val="${o.plot.w}"/><c:h val="${o.plot.h}"/></c:manualLayout></c:layout>` : '<c:layout/>')
    + `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${ser}<c:axId val="5001"/><c:axId val="5002"/></c:scatterChart>`
    + axis(5001,5002,'b',o.xTitle,o.xAxis) + axis(5002,5001,'l',o.yTitle,o.yAxis)
    + `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>`
    + (o.legend===false ? '' : `<c:legend><c:legendPos val="b"/>${(o.legendDel||[]).map(i=>`<c:legendEntry><c:idx val="${i}"/><c:delete val="1"/></c:legendEntry>`).join('')}<c:overlay val="0"/><c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="800"/></a:pPr><a:endParaRPr lang="zh-TW"/></a:p></c:txPr></c:legend>`)
    + `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>`
    + `<c:extLst><c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}"><c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="1"/></c16r3:dataDisplayOptions16></c:ext></c:extLst>`
    + `</c:chart><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="6350"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln></c:spPr>`
    + `<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr><a:latin typeface="Arial"/><a:ea typeface="Microsoft JhengHei"/></a:defRPr></a:pPr><a:endParaRPr lang="zh-TW"/></a:p></c:txPr>`
    + `<c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings>`
    + `</c:chartSpace>`;
}

/* 把圖表注入 xlsx（specs：[{sheet, from:[col,row], to:[col,row], xml}]，列欄 0 起算） */
async function injectCharts(buf, specs, JSZip){
  if(!specs.length) return buf;
  const zip = await JSZip.loadAsync(buf);
  const wbXml = await zip.file('xl/workbook.xml').async('string');
  const wbRels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const sheetPath = name => {
    const m = wbXml.match(new RegExp(`<sheet[^>]*name="${xmlEsc(name).replace(/[()]/g,'\\$&')}"[^>]*r:id="(rId\\d+)"`));
    if(!m) throw new Error('找不到工作表 '+name);
    const t = wbRels.match(new RegExp(`Id="${m[1]}"[^>]*Target="([^"]+)"`)) || wbRels.match(new RegExp(`Target="([^"]+)"[^>]*Id="${m[1]}"`));
    return 'xl/' + t[1].replace(/^\/?xl\//,'');
  };
  let ct = await zip.file('[Content_Types].xml').async('string');
  const bySheet = {};
  specs.forEach(s => (bySheet[s.sheet] = bySheet[s.sheet] || []).push(s));
  let chartNo = 0, drawNo = 0;
  for(const [sheet, list] of Object.entries(bySheet)){
    drawNo++;
    const sp = sheetPath(sheet), sf = sp.split('/').pop();
    const relPath = `xl/worksheets/_rels/${sf}.rels`;
    let rels = zip.file(relPath) ? await zip.file(relPath).async('string')
      : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    const rid = 'rIdDraw' + drawNo;
    rels = rels.replace('</Relationships>', `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawNo}.xml"/></Relationships>`);
    zip.file(relPath, rels);
    let sx = await zip.file(sp).async('string');
    const tag = `<drawing r:id="${rid}"/>`;
    if(/<legacyDrawing/.test(sx)) sx = sx.replace(/<legacyDrawing/, tag+'<legacyDrawing');
    else if(/<tableParts/.test(sx)) sx = sx.replace(/<tableParts/, tag+'<tableParts');
    else if(/<extLst>/.test(sx)) sx = sx.replace(/<extLst>/, tag+'<extLst>');
    else sx = sx.replace('</worksheet>', tag+'</worksheet>');
    if(!/xmlns:r=/.test(sx.slice(0,600))) sx = sx.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
    zip.file(sp, sx);
    const anchors=[], drels=[];
    list.forEach((s,i)=>{
      chartNo++;
      zip.file(`xl/charts/chart${chartNo}.xml`, s.xml);
      ct = ct.replace('</Types>', `<Override PartName="/xl/charts/chart${chartNo}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>`);
      drels.push(`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chartNo}.xml"/>`);
      anchors.push(`<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${s.from[0]}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${s.from[1]}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>`
        + `<xdr:to><xdr:col>${s.to[0]}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${s.to[1]}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>`
        + `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${i+2}" name="${xmlEsc(s.name||('圖表 '+(i+1)))}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>`
        + `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">`
        + `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId${i+1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`);
    });
    zip.file(`xl/drawings/drawing${drawNo}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${anchors.join('')}</xdr:wsDr>`);
    zip.file(`xl/drawings/_rels/drawing${drawNo}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${drels.join('')}</Relationships>`);
    ct = ct.replace('</Types>', `<Override PartName="/xl/drawings/drawing${drawNo}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`);
  }
  zip.file('[Content_Types].xml', ct);
  return zip.generateAsync({type: typeof window!=='undefined' ? 'arraybuffer' : 'nodebuffer', compression:'DEFLATE'});
}

/* 圖面尺寸（pt）：A4 計算書 B~F 欄寬合計 521pt；斷面圖 34 列、P-M 圖 26 列，列高 14pt */
const FIG_SEC_ROWS = 34, FIG_PM_ROWS = 23, ROW_PT = 14;
const SEC_W = 521, SEC_H = FIG_SEC_ROWS*ROW_PT;
const SEC_PLOT = {x:0.04, y:0.09, w:0.92, h:0.76};
/* 讓 x、y 每公分等長：依繪圖區實際尺寸放大較緊的一向 */
function equalAxes(Wcm, Hcm, chartW, chartH){
  const pw = chartW*SEC_PLOT.w, ph = chartH*SEC_PLOT.h;
  const s = Math.max(Wcm*1.12/pw, Hcm*1.12/ph);          // cm / pt
  return {hx:+(pw*s/2).toFixed(2), hy:+(ph*s/2).toFixed(2), ptPerCm:1/s};
}

/* ---------- 柱：圖表資料（全部公式） ---------- */
const CD = '圖表資料';
function buildColumnChartData(ws, R){
  const r = f => ({f:R(f,CD)});
  ws.columns = Array(16).fill(0).map(()=>({width:11}));
  put(ws,'A1','斷面配筋圖資料（座標 cm，原點為斷面中心；#N/A＝不繪）',{sec:true});
  const heads = ['外框 x','外框 y','圓外框 x','圓外框 y','中空 x','中空 y','箍筋 x','箍筋 y','圓箍 x','圓箍 y','內箍 x','內箍 y','繫筋 x','繫筋 y','主筋 x','主筋 y'];
  heads.forEach((h,j)=>put(ws,colL(j)+'2',h,{head:true}));
  const rect = (c1, c2, hx, hy, cond) => {
    const P=[[1,1],[1,-1],[-1,-1],[-1,1],[1,1]];
    P.forEach(([sx,sy],i)=>{ put(ws,c1+(3+i), r(`IF(${cond},${sx}*(${hx}),${NA})`),{fmt:'0.0'}); put(ws,c2+(3+i), r(`IF(${cond},${sy}*(${hy}),${NA})`),{fmt:'0.0'}); });
  };
  const circ = (c1, c2, rad, cond) => {
    for(let k=0;k<=72;k++){ const th=`(${k}*PI()/36)`;
      put(ws,c1+(3+k), r(`IF(${cond},(${rad})*COS(${th}),${NA})`),{fmt:'0.0'}); put(ws,c2+(3+k), r(`IF(${cond},(${rad})*SIN(${th}),${NA})`),{fmt:'0.0'}); }
  };
  const off = '({db}/2+{dt}/2)';
  rect('A','B','{Be}/2','{He}/2','{isCirc}=0');
  circ('C','D','{Din}/2','{isCirc}=1');
  rect('E','F','({Be}/2-{tw})','({He}/2-{tw})','{isBox}=1');
  rect('G','H',`({hxO}+${off})`,`({hyO}+${off})`,'{isCirc}=0');
  circ('I','J',`({rb}+${off})`,'{isCirc}=1');
  rect('K','L',`({hxI}+${off})`,`({hyI}+${off})`,'{inOK}=1');
  // 繫筋：每支 6 列（兩段＋間隔）；垂直（B 邊 x_i）接著水平（H 邊 y_j）
  const C = "'配筋座標'!";
  let n=3;
  for(const dir of ['v','h']){
    for(let i=1;i<=NMAX;i++){
      const cr=i+2, flag = dir==='v' ? `${C}D${cr}` : `${C}E${cr}`, pos = dir==='v' ? `${C}B${cr}` : `${C}C${cr}`;
      const ho = dir==='v' ? '{hyO}' : '{hxO}', hi = dir==='v' ? '{hyI}' : '{hxI}';
      const on = `AND({isCirc}=0,${flag}=1)`;
      // 段 1：外緣 → 對側外緣（實心／無內層）或 外緣 → 內層（雙層箱型）
      const ends = [[ho, `IF({inOK}=1,${hi},-${ho})`], [`-${ho}`, `-${hi}`]];
      ends.forEach((e,si)=>{
        const segOn = si===0 ? on : `AND(${on},{inOK}=1)`;
        [e[0], e[1]].forEach((v,k)=>{
          const xv = dir==='v' ? pos : v, yv = dir==='v' ? v : pos;
          put(ws,'M'+n, r(`IF(${segOn},${xv},${NA})`),{fmt:'0.0'}); put(ws,'N'+n, r(`IF(${segOn},${yv},${NA})`),{fmt:'0.0'}); n++;
        });
        put(ws,'M'+n,{f:NA}); put(ws,'N'+n,{f:NA}); n++;
      });
    }
  }
  const tieEnd = n-1;
  // 主筋：外層上下列、外層左右（去角隅）、內層上下列、內層左右、圓形
  n=3;
  const bar = (x, y, cond) => { put(ws,'O'+n, r(`IF(${cond},${x},${NA})`),{fmt:'0.0'}); put(ws,'P'+n, r(`IF(${cond},${y},${NA})`),{fmt:'0.0'}); n++; };
  for(let i=1;i<=NMAX;i++){ const cr=i+2; bar(`${C}B${cr}`,'{hyO}',`AND({isCirc}=0,${i}<={nB})`); bar(`${C}B${cr}`,'-{hyO}',`AND({isCirc}=0,${i}<={nB})`); }
  for(let i=1;i<=NMAX;i++){ const cr=i+2; bar('{hxO}',`${C}C${cr}`,`AND({isCirc}=0,${i}>=2,${i}<={nH}-1)`); bar('-{hxO}',`${C}C${cr}`,`AND({isCirc}=0,${i}>=2,${i}<={nH}-1)`); }
  for(const sx of [1,-1]) for(const sy of [1,-1]) bar(`${sx}*{hxI}`,`${sy}*{hyI}`,'{inOK}=1');
  for(let i=1;i<=NMAX;i++){ const cr=i+2; bar(`${C}B${cr}`,'{hyI}',`${C}F${cr}=1`); bar(`${C}B${cr}`,'-{hyI}',`${C}F${cr}=1`); }
  for(let i=1;i<=NMAX;i++){ const cr=i+2; bar('{hxI}',`${C}C${cr}`,`${C}G${cr}=1`); bar('-{hxI}',`${C}C${cr}`,`${C}G${cr}=1`); }
  for(let i=1;i<=NMAX;i++){ const cr=i+2; bar(`${C}I${cr}`,`${C}J${cr}`,`AND({isCirc}=1,${i}<={nC})`); }
  const barEnd = n-1;
  // 圖名（連結輸入，隨尺寸更新）
  put(ws,'R1', r(`"斷面配筋圖　"&IF({isCirc}=1,"D = "&TEXT({Din},"0")&" cm",TEXT({Be},"0")&" × "&TEXT({He},"0")&" cm"&IF({isBox}=1,"，tw = "&TEXT({tw},"0")&" cm",""))&"　主筋 "&{bar}&" × "&{nBars}&"、橫向筋 "&{tie}&" @ "&TEXT({sUse},"0.0")&" cm"`));
  // 載重點：X 軸（圓形取合彎矩）、Y 軸
  ['M_x','P','M_y','P'].forEach((h,j)=>put(ws,colL(18+j)+'2',['載重 Mx','載重 P','載重 My','載重 P'][j],{head:true}));
  for(let i=0;i<NCB;i++){
    const lr=4+i, on=`'載重組合'!H${lr}=1`;
    put(ws,'S'+(3+i), r(`IF(${on},IF({isCirc}=1,SQRT('載重組合'!C${lr}^2+'載重組合'!D${lr}^2),ABS('載重組合'!C${lr})),${NA})`),{fmt:'#,##0.0'});
    put(ws,'T'+(3+i), {f:`IF(${on},'載重組合'!B${lr},${NA})`},{fmt:'#,##0.0'});
    put(ws,'U'+(3+i), {f:`IF(${on},ABS('載重組合'!D${lr}),${NA})`},{fmt:'#,##0.0'});
    put(ws,'V'+(3+i), {f:`IF(${on},'載重組合'!B${lr},${NA})`},{fmt:'#,##0.0'});
  }
  return {tieEnd, barEnd};
}

/* 柱圖表規格（依匯出當下尺寸定座標軸範圍與主筋標記大小） */
function columnCharts(inp, info, a4Start){
  const B = inp.type==='circle'?inp.D:inp.B, H = inp.type==='circle'?inp.D:inp.H;
  const db = (BARS.find(b=>b[0]===inp.barSize)||[0,2.5])[1];
  const eq = equalAxes(B, H, SEC_W, SEC_H);
  const mk = db*eq.ptPerCm*1.1;
  const s = (c1,c2,r2) => [`${qs(CD)}!$${c1}$3:$${c1}$${r2}`, `${qs(CD)}!$${c2}$3:$${c2}$${r2}`];
  const ser = (name, xy, line, marker) => ({name, x:xy[0], y:xy[1], line, marker});
  const circle = inp.type==='circle', box = inp.type==='box';
  const sec = scatterChartXml({
    title:'斷面配筋圖', titleRef:`${qs(CD)}!$R$1`, legend:true, plot:SEC_PLOT,
    xAxis:{min:-eq.hx, max:eq.hx, hidden:true}, yAxis:{min:-eq.hy, max:eq.hy, hidden:true},
    legendDel: [circle?0:1, ...(box?[]:[2]), circle?3:4, ...(box&&inp.dbl?[]:[5]), ...(circle?[6]:[])],
    series:[
      ser('混凝土', s('A','B',7), {color:'8A97A5', w:1.75}),
      ser('混凝土（圓）', s('C','D',75), {color:'8A97A5', w:1.75}),
      ser('中空區', s('E','F',7), {color:'8A97A5', w:1.25, dash:'dash'}),
      ser('外閉合箍筋', s('G','H',7), {color:'0F5F6B', w:2}),
      ser(inp.spiral?'螺旋箍筋':'圓形箍筋', s('I','J',75), {color:'0F5F6B', w:2}),
      ser('內層閉合箍筋', s('K','L',7), {color:'0F5F6B', w:1.5, dash:'dash'}),
      ser('繫筋', s('M','N',info.tieEnd), {color:'B0407A', w:1.25}),
      ser(`主筋 ${inp.barSize}`, s('O','P',info.barEnd), null, {symbol:'circle', size:mk, color:'151C24', fill:'151C24'})
    ]});
  const pm = ax => {
    const P = ax==='x' ? 'P-M_X' : 'P-M_Y', rng = (c,a,b)=>`${qs(P)}!$${c}$${a}:$${c}$${b}`;

    return scatterChartXml({
      title: circle ? 'P-M 互制曲線（圓形：合彎矩）' : `P-M 互制曲線（繞 ${ax.toUpperCase()} 軸）`,
      xTitle: circle ? 'Mr = √(Mx² + My²) (tf·m)' : `M${ax} (tf·m)`, yTitle:'P (tf)，壓為正',
      xAxis:{fmt:'#,##0'}, yAxis:{fmt:'#,##0'},
      series:[
        ser('標稱 Pn–Mn', [rng('G',PM_H+1,PM_H+NPM), rng('F',PM_H+1,PM_H+NPM)], {color:'8A97A5', w:1.25}),
        ser('設計 φPn–φMn', [rng('Q',PM_H+1,PM_H+POLY_N), rng('R',PM_H+1,PM_H+POLY_N)], {color:'0F5F6B', w:2.25}),
        ser('載重組合', ax==='x' ? [`${qs(CD)}!$S$3:$S$${2+NCB}`, `${qs(CD)}!$T$3:$T$${2+NCB}`] : [`${qs(CD)}!$U$3:$U$${2+NCB}`, `${qs(CD)}!$V$3:$V$${2+NCB}`], null, {symbol:'diamond', size:7, color:'B03A2E', fill:'B03A2E'})
      ]});
  };
  const A4 = '結構計算書(A4)';
  return [
    {sheet:A4, name:'斷面配筋圖', from:[1,a4Start], to:[6,a4Start+FIG_SEC_ROWS], xml:sec},
    {sheet:A4, name:'P-M 互制曲線 X', from:[1,a4Start+FIG_SEC_ROWS+1], to:[6,a4Start+FIG_SEC_ROWS+1+FIG_PM_ROWS], xml:pm('x')},
    ...(circle ? [] : [{sheet:A4, name:'P-M 互制曲線 Y', from:[1,a4Start+FIG_SEC_ROWS+2+FIG_PM_ROWS], to:[6,a4Start+FIG_SEC_ROWS+2+2*FIG_PM_ROWS], xml:pm('y')}])
  ];
}

/* ---------- 梁／版：圖表資料 ---------- */
function buildBeamChartData(ws, R){
  const r = f => ({f:R(f,CD)});
  ws.columns = Array(8).fill(0).map(()=>({width:11}));
  put(ws,'A1','梁／版斷面圖資料（x 以寬度中心為 0，y 自底面起算，cm）',{sec:true});
  ['外框 x','外框 y','箍筋 x','箍筋 y','主筋 x','主筋 y','受壓區 x','受壓區 y'].forEach((h,j)=>put(ws,colL(j)+'2',h,{head:true}));
  // 外框：T 型 9 點；矩形 5 點後補 #N/A
  const T = [['-{be}/2','{h}'],['{be}/2','{h}'],['{be}/2','{h}-{hf}'],['{bw}/2','{h}-{hf}'],['{bw}/2','0'],['-{bw}/2','0'],['-{bw}/2','{h}-{hf}'],['-{be}/2','{h}-{hf}'],['-{be}/2','{h}']];
  const Rc = [['-{bw}/2','{h}'],['{bw}/2','{h}'],['{bw}/2','0'],['-{bw}/2','0'],['-{bw}/2','{h}']];
  for(let i=0;i<9;i++){
    const t=T[i], q=Rc[i];
    put(ws,'A'+(3+i), r(q ? `IF({isT}=1,${t[0]},${q[0]})` : `IF({isT}=1,${t[0]},${NA})`),{fmt:'0.0'});
    put(ws,'B'+(3+i), r(q ? `IF({isT}=1,${t[1]},${q[1]})` : `IF({isT}=1,${t[1]},${NA})`),{fmt:'0.0'});
  }
  // 箍筋（版不繪）
  [[1,1],[1,-1],[-1,-1],[-1,1],[1,1]].forEach(([sx,sy],i)=>{
    put(ws,'C'+(3+i), r(`IF({isSlab}=1,${NA},${sx}*({bw}/2-{cover}-{dt}/2))`),{fmt:'0.0'});
    put(ws,'D'+(3+i), r(`IF({isSlab}=1,${NA},{h}/2+${sy}*({h}/2-{cover}-{dt}/2))`),{fmt:'0.0'});
  });
  // 主筋：每層 41 列（梁：等分排列；版：依間距自中心向兩側）
  let n=3;
  const layer = (nKey, yKey, dbKey, spKey) => {
    for(let k=-20;k<=20;k++){
      const i = k+21;
      const xBeam = `(-{bw}/2+{cover}+{dt}+${dbKey}/2)+(({bw}/2-{cover}-{dt}-${dbKey}/2)-(-{bw}/2+{cover}+{dt}+${dbKey}/2))*(${i}-1)/MAX(1,ROUND(${nKey},0)-1)`;
      const xSlab = `${k}*${spKey}`;
      const onBeam = `AND({isSlab}=0,${i}<=ROUND(${nKey},0),ROUND(${nKey},0)>0)`;
      const onSlab = `AND({isSlab}=1,${spKey}>0,ABS(${k}*${spKey})<={bw}/2-${dbKey}/2)`;
      put(ws,'E'+n, r(`IF(${onBeam},IF(ROUND(${nKey},0)=1,0,${xBeam}),IF(${onSlab},${xSlab},${NA}))`),{fmt:'0.0'});
      put(ws,'F'+n, r(`IF(OR(${onBeam},${onSlab}),${yKey},${NA})`),{fmt:'0.0'});
      n++;
    }
  };
  layer('{r1}','{y1}','{dbB}','{spB}');
  layer('IF({isSlab}=1,0,{r2})','{y2}','{dbB}','0');
  layer('{nTop}','{yT}','{dbT}','{spT}');
  const barEnd=n-1;
  // 受壓區（應力塊 a）
  const yA = 'IF({isPos}=1,{h},{a})', yB = 'IF({isPos}=1,{h}-{a},0)', wC = 'IF(AND({isT}=1,{isPos}=1),{be},{bw})/2';
  [['-1',yA],['1',yA],['1',yB],['-1',yB],['-1',yA]].forEach(([sx,y],i)=>{
    put(ws,'G'+(3+i), r(`${sx}*${wC}`),{fmt:'0.0'}); put(ws,'H'+(3+i), r(y),{fmt:'0.0'});
  });
  put(ws,'J1', r(`IF({isSlab}=1,"版斷面圖（每公尺寬）　h = "&TEXT({h},"0")&" cm　底 "&{botS}&" @ "&TEXT({spB},"0.0")&IF({spT}>0,"、頂 "&{topS}&" @ "&TEXT({spT},"0.0"),"")&" cm","梁斷面圖　"&IF({isT}=1,"T 型 be = "&TEXT({be},"0")&"、","")&"bw × h = "&TEXT({bw},"0")&" × "&TEXT({h},"0")&" cm　底 "&ROUND({nBot},0)&"-"&{botS}&"、頂 "&ROUND({nTop},0)&"-"&{topS})`));
  return {barEnd};
}
function beamCharts(inp, info, a4Start){
  const W = inp.slab ? 100 : (inp.type==='T' ? Math.min(Math.max(inp.bw, inp.L/4, inp.bw+16*inp.hf), Math.max(inp.bw, inp.sSpacing||inp.bw)) : inp.bw);
  const db = (BARS.find(b=>b[0]===inp.botSize)||[0,2.5])[1];
  const eq = equalAxes(W, inp.h, SEC_W, SEC_H);
  const mk = db*eq.ptPerCm*1.1;
  const s = (c1,c2,r2) => [`${qs(CD)}!$${c1}$3:$${c1}$${r2}`, `${qs(CD)}!$${c2}$3:$${c2}$${r2}`];
  const xml = scatterChartXml({
    title: inp.slab?'版斷面圖（每公尺寬）':'梁斷面圖', titleRef:`${qs(CD)}!$J$1`, plot:SEC_PLOT,
    xAxis:{min:-eq.hx, max:eq.hx, hidden:true}, yAxis:{min:+(inp.h/2-eq.hy).toFixed(2), max:+(inp.h/2+eq.hy).toFixed(2), hidden:true},
    legendDel: inp.slab ? [2] : [],
    series:[
      {name:'混凝土', x:s('A','B',11)[0], y:s('A','B',11)[1], line:{color:'8A97A5', w:1.75}},
      {name:'受壓區 a', x:s('G','H',7)[0], y:s('G','H',7)[1], line:{color:'8A5A12', w:1.25, dash:'dash'}},
      {name:'箍筋', x:s('C','D',7)[0], y:s('C','D',7)[1], line:{color:'0F5F6B', w:2}},
      {name:'主筋', x:s('E','F',info.barEnd)[0], y:s('E','F',info.barEnd)[1], marker:{symbol:'circle', size:mk, color:'151C24', fill:'151C24'}}
    ]});
  return [{sheet:'結構計算書(A4)', name:'斷面圖', from:[1,a4Start], to:[6,a4Start+FIG_SEC_ROWS], xml}];
}

/* 附圖頁：預留圖表列，回傳第一列（0 起算） */
function a4Figures(ws, a, title, nPm){
  ws.getRow(a.n).addPageBreak();
  a.chap(title);
  const start = a.n;                 // 0 起算之列號＝已用列數
  const total = FIG_SEC_ROWS + 1 + (nPm ? 1 + nPm*FIG_PM_ROWS + (nPm-1) : 0);
  for(let i=0;i<total;i++){ const n=a.row(); ws.getRow(n).height=ROW_PT; if(nPm && i===FIG_SEC_ROWS) ws.getRow(n).addPageBreak(); }
  return start;
}
/* 寫出：ExcelJS 產生後注入原生圖表 */
async function toBuffer(out, JSZip){
  const buf = await out.wb.xlsx.writeBuffer();
  return injectCharts(buf, out.charts||[], JSZip);
}
/* ---------- 輸出 ---------- */
const api = { buildColumn, buildBeam, toBuffer };
if(typeof module!=='undefined' && module.exports) module.exports = api;
else root.RCXLSX = api;
})(typeof window!=='undefined' ? window : globalThis);

# 專案規則

## 操作教學與常見問題 PDF 於合併到 main 時重建

- `docs/RC斷面設計工具_操作教學.pdf` 與 `docs/RC斷面設計工具_常見問題.pdf` 由 `docs/tutorial/build.js` 分別依 `docs/tutorial/tutorial.html`、`docs/tutorial/faq.html` 建置（共用截圖）；說明見 `docs/tutorial/README.md`。
- **分支提交不要重建或提交 PDF 與 `stamp.txt`**（每份約 6～7 MB，會讓儲存庫快速變大）。推送（合併）到 main 時，GitHub Actions（`tutorial-pdf.yml`）會重建並以機器人提交推回 main。
- 修改 `index.html`、`calc-xlsx.js` 或教學原稿時，若介面、操作步驟、檢核項目或規範條文有變，仍須同步修改 `docs/tutorial/tutorial.html` 與 `docs/tutorial/faq.html` 的說明文字（必要時在 `build.js` 增加截圖）；`README.md` 亦同。
- 要確認原稿能正確建置時可在本機 `cd docs/tutorial && npm run build`，但建置產物不要提交（`git checkout -- docs/*.pdf docs/tutorial/stamp.txt`）。
- 教學與常見問題中的公式與符號一律用 KaTeX（`\( … \)`），與網頁、計算書、Excel 的 LaTeX 風格一致。

## 規範依據

- 規範原文 PDF 在 `specs/`（索引與核對狀態見 `specs/README.md`）。新增或修改任何規範檢核時，以原文核對條號、公式與係數，並更新 `specs/README.md` 的核對狀態。
- 建築依《建築物混凝土結構設計規範》112 年版（土木401-112）。橋梁：長細效應與保護層依《公路橋梁設計規範》；容量設計剪力、剪力（φ = 0.85、V_c）、圍束筋、主筋比、材料依《公路橋梁耐震設計規範》；斷面撓曲強度與彎矩 φ 維持 112 年版。

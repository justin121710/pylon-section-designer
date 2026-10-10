# 專案規則

## 操作教學 PDF 必須在每個提交保持最新

- `docs/RC斷面設計工具_操作教學.pdf` 由 `docs/tutorial/build.js` 依 `docs/tutorial/tutorial.html` 建置；說明見 `docs/tutorial/README.md`。
- 先確認已啟用 hook：`git config core.hooksPath .githooks`（pre-commit 會在 PDF 過期時自動重建並加入提交）。
- 修改 `index.html`、`calc-xlsx.js` 或教學原稿時：
  1. 若介面、操作步驟、檢核項目或規範條文有變，同步修改 `docs/tutorial/tutorial.html` 的說明文字（必要時在 `build.js` 增加截圖）；`FAQ.md`、`README.md` 亦同。
  2. 提交前執行 `sh docs/tutorial/check.sh`；過期就 `cd docs/tutorial && npm run build`，PDF 與 `stamp.txt` 與程式放在同一個提交。
- 教學中的公式與符號一律用 KaTeX（`\( … \)`），與網頁、計算書、Excel 的 LaTeX 風格一致。

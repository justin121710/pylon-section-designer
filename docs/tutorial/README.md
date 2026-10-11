# 操作教學與常見問題 PDF 建置

輸出（A4，含標註截圖；公式與符號以 KaTeX 排版，與網頁、計算書、Excel 的 LaTeX 風格一致）：
- `docs/RC斷面設計工具_操作教學.pdf`：原稿 `tutorial.html`
- `docs/RC斷面設計工具_常見問題.pdf`：原稿 `faq.html`（原 FAQ.md，改為附圖問答；截圖 `img/f-*.png` 由 build.js 另行拍攝，其餘沿用教學截圖）

| 檔案 | 用途 |
|---|---|
| `faq.html` | 常見問題原稿，規則同 `tutorial.html` |
| `tutorial.html` | 教學原稿。公式、符號一律寫成 `\( … \)`（KaTeX），例如 `\(V_c\)`、`\(\dfrac{kl_u}{r}\le 22\)`；不要用 `<sub>`／`<sup>` 或純文字 φ、≤ |
| `build.js` | 開啟 `index.html` 依教學步驟操作截圖（紅框編號）、匯出 Excel 轉圖、列印 PDF，並寫入 `stamp.txt` |
| `inputs.txt` | 影響 PDF 內容的檔案清單；任一檔變更即須重建 |
| `stamp.txt` | 建置當時各輸入檔的 git blob 雜湊（由 build.js 產生，勿手改） |
| `check.sh` | 比對 `stamp.txt` 與目前提交內容；過期時回傳失敗 |

## 合併到 main 時重建 PDF

- **GitHub Actions**（`.github/workflows/tutorial-pdf.yml`）：推送（合併）到 `main` 時執行 `check.sh`，PDF 過期即重建，
  以 `[skip ci]` 機器人提交推回 `main`，並上傳為 artifact；亦可在 Actions 頁面手動執行。
- 分支提交**不附 PDF**（每份 PDF 約 6～7 MB，每個提交都附會讓儲存庫快速變大）；分支上的 PDF 暫時過期屬正常。
- **pre-commit hook**（`.githooks/pre-commit`）：只提示 PDF 與程式不同步，不重建、不阻擋。啟用：`git config core.hooksPath .githooks`
- 介面、操作流程或規範檢核有改動時，除了重建，也要**同步修改 `tutorial.html` 與 `faq.html` 的文字**；截圖會自動更新，
  但步驟說明不會。新增的按鈕或畫面請在 `build.js` 增加截圖與標註。

## 手動建置

```sh
cd docs/tutorial
npm install                 # playwright、katex、exceljs、jszip
npx playwright install chromium   # 若本機尚無 Chromium（或設 CHROMIUM_PATH 指向既有執行檔）
npm run build               # 約 1 分鐘
npm run check
```

另需 LibreOffice（`soffice`）、poppler-utils（`pdftoppm`、`pdftotext`、`pdfinfo`）與 python3 + Pillow（Excel 頁面轉圖與裁邊）。

#!/bin/sh
# 檢查操作教學與常見問題 PDF 是否對應目前（已暫存／已提交）的程式與原始檔。
# stamp.txt 記錄建置當時各輸入檔的 git blob 雜湊；不一致代表 PDF 過期，須執行 npm run build。
# 用於 .githooks/pre-commit 與 GitHub Actions；只需要 git。
cd "$(git rev-parse --show-toplevel)" || exit 2
T=docs/tutorial
PDF="docs/RC斷面設計工具_操作教學.pdf"
FAQ="docs/RC斷面設計工具_常見問題.pdf"
exp=$(grep -v '^#' "$T/inputs.txt" | sed '/^[[:space:]]*$/d' | while read -r f; do
  b=$(git ls-files -s -- "$f" | awk '{print $2}')
  [ -n "$b" ] || b=$(git hash-object -- "$f" 2>/dev/null)
  printf '%s  %s\n' "$b" "$f"
done)
cur=$(git show ":$T/stamp.txt" 2>/dev/null || cat "$T/stamp.txt" 2>/dev/null)
for p in "$PDF" "$FAQ"; do
  if [ -z "$(git ls-files -- "$p")" ]; then
    echo "✗ PDF 尚未加入 git：$p"; exit 1
  fi
done
if [ "$exp" != "$cur" ]; then
  echo "✗ 操作教學／常見問題 PDF 已過期：以下輸入檔與上次建置時不同"
  printf '%s\n' "$exp" > /tmp/.rcsd-exp.$$; printf '%s\n' "$cur" > /tmp/.rcsd-cur.$$
  diff /tmp/.rcsd-cur.$$ /tmp/.rcsd-exp.$$ | sed -n 's/^> [0-9a-f]*  /  · /p'
  rm -f /tmp/.rcsd-exp.$$ /tmp/.rcsd-cur.$$
  echo "  請執行：cd docs/tutorial && npm install && npm run build，並一併提交兩份 PDF 與 stamp.txt"
  exit 1
fi
echo "✓ 操作教學與常見問題 PDF 為最新"

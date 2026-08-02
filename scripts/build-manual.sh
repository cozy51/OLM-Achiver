#!/usr/bin/env bash
# docs/manual.html から利用マニュアルの PDF を生成します。
#   ./scripts/build-manual.sh [出力先.pdf]
# Chromium / Edge / Chrome のいずれかが必要です。日本語フォント(IPAゴシック等)が
# 入っていない環境では文字化けするため、事前に導入してください。
set -euo pipefail

root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
source_html="$root/docs/manual.html"
output="${1:-$root/docs/Outlook-Mail-Archiver-manual.pdf}"

browser=""
for candidate in \
  "${CHROME:-}" \
  /opt/pw-browsers/chromium-*/chrome-linux/chrome \
  "$(command -v chromium || true)" \
  "$(command -v chromium-browser || true)" \
  "$(command -v google-chrome || true)" \
  "$(command -v microsoft-edge || true)"
do
  if [ -n "$candidate" ] && [ -x "$candidate" ]; then browser="$candidate"; break; fi
done

if [ -z "$browser" ]; then
  echo "Chromium系ブラウザーが見つかりません。CHROME=<実行ファイルのパス> を指定してください。" >&2
  exit 1
fi

"$browser" --headless --no-sandbox --disable-gpu \
  --no-pdf-header-footer \
  --print-to-pdf="$output" \
  "file://$source_html" 2>/dev/null

echo "作成しました: $output"

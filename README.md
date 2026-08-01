# Outlook Mail PDF Archiver for Edge

Outlook on the web で現在開いているメールを PDF 化する Microsoft Edge 拡張機能です。PDF の先頭には元メールを開くクリック可能な URL を挿入します。

## インストール

1. Edge で `edge://extensions` を開きます。
2. **開発者モード**をオンにします。
3. **展開して読み込み**を選び、このフォルダーを指定します。

すでに Outlook を開いている場合も、そのまま利用できます。拡張機能は必要に応じてメール読み取り処理を現在のタブへ自動的に読み込みます。

## 使い方

1. Outlook on the web で保存したいメールを1通開きます。
2. ツールバーの拡張機能アイコンを押し、**📩 このメールを PDF 保存**を選びます。
3. 保存先を選択します。

ファイル名は `📩_差出人_年月日_日時_件名.pdf` です。件名先頭の `RE:`、`FW:`、`FWD:`、`返信:`、`転送:`（重複を含む）は自動的に取り除きます。

> PDF 作成中に Edge の「デバッガーがこのブラウザーに接続しました」という通知が一時的に表示されます。これは印刷 API を利用するための正常な動作です。

## 対応 URL

- `outlook.office.com`
- `outlook.office365.com`
- `outlook.cloud.microsoft`

Outlook の画面構造が変更された場合は、メール本文を検出できなくなる可能性があります。

## トラブルシューティング

- `Could not establish connection. Receiving end does not exist.` が表示された場合は、最新版へ更新してください。初回実行時に必要な処理を自動注入してから再試行します。
- `Cannot access a chrome-extension:// URL of different extension` が表示されていた問題は修正済みです。拡張機能のポップアップではなく、同じウィンドウ内の Outlook タブだけを選択して処理します。
- Outlook のサインイン画面やメール一覧だけを表示している場合は保存できません。対象メールを開いてから実行してください。

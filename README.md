# Outlook Mail PDF Archiver for Edge

Outlook on the web で現在開いているメールを PDF 化する Microsoft Edge 拡張機能です。PDF の先頭には元メールを開くクリック可能な URL を挿入します。

## インストール

1. Edge で `edge://extensions` を開きます。
2. **開発者モード**をオンにします。
3. **展開して読み込み**を選び、このフォルダーを指定します。

すでに Outlook を開いている場合も、そのまま利用できます。

## 使い方

1. Outlook on the web で保存したいメールを1通開きます。
2. ツールバーの拡張機能アイコンを押し、**📩 このメールを PDF 保存**を選びます。
3. 保存先を選択します。

組織のセキュリティ設定により Edge のPDF機能へ直接接続できない場合は、自動的に標準の印刷画面へ切り替わります。印刷画面の「プリンター」で **PDFとして保存** を選択してください。

ファイル名は `📩_差出人_年月日_日時_件名.pdf` です。件名先頭の `RE:`、`FW:`、`FWD:`、`返信:`、`転送:`（重複を含む）は自動的に取り除きます。重要度記号付きの`[!]RE:`や、すでに区切りが置換された`[!]RE_`も削除します。

`年月日_日時`にはPDFを作成した時刻ではなく、Outlookのメールヘッダーに表示される受信日時（例: `2026-07-13 (月) 20:41`）を使用します。

> PDF 作成中に Edge の「デバッガーがこのブラウザーに接続しました」という通知が一時的に表示されます。これは印刷 API を利用するための正常な動作です。

## 対応 URL

- `outlook.office.com`
- `outlook.office365.com`
- `outlook.cloud.microsoft`

Outlook の画面構造が変更された場合は、メール本文を検出できなくなる可能性があります。

Outlookの画面構造による差出人の取り違えを防ぐため、差出人は自動取得しません。Outlookに表示されている差出人名をコピーしてから拡張機能を開き、`差出人名`欄へ貼り付けてください。

入力例には実在の利用者を特定しない`山田 太郎 (Yamada,T.)`を使用しています。

アドレスバーが`https://outlook.office.com/mail/`のままでも、選択中のメール行からItem IDを取得します。`outlook.office.com/mail/inbox/id/...`形式は直接埋め込まず、`https://outlook.cloud.microsoft/mail/deeplink/read/...?...&exvsurl=1`形式へ変換してPDFへ埋め込みます。個別URLを特定できない場合、メールトップへのリンクで代用せずエラーを表示します。

## トラブルシューティング

- `Could not establish connection` や `Cannot access a chrome-extension:// URL of different extension` が表示されていた問題は修正済みです。最新版ではコンテンツスクリプトを注入せず、PDF 出力にも使用する Edge のデバッグ API 経由で印刷用ページを準備します。
- Outlook のサインイン画面やメール一覧だけを表示している場合は保存できません。対象メールを開いてから実行してください。
- Outlookの画面に閲覧ペインを識別する属性がない場合も、ページ全体を探索してメール情報を取得します。
- エラーの先頭に `【失敗箇所: ...】` を表示します。`Edge のPDF機能への接続` ならURL取得・PDF作成より前のセキュリティ制限、`メール内容とURLの取得` ならOutlook画面の読み取り、`PDFの作成` ならPDF変換段階の問題です。
- `Cannot access a chrome-extension:// URL of different extension` をEdgeが返した場合は、標準の印刷画面へ自動的に切り替えます。そこで **PDFとして保存** を選択すれば、元メールリンク付きの内容と推奨ファイル名を維持して保存できます。

# Spreadsheet -> WordPress 投稿 Apps Script

このリポジトリには、Google スプレッドシートに紐づける Apps Script (`Code.gs`) を置いています。

## できること

- 1行目に指定ヘッダーを設定
- A列が `入稿待ち` の行だけ処理
- B列の Google ドキュメントURLから本文とタイトルを取得
- WordPress REST API に投稿（新規/更新）
- 初回投稿時に `WP_POST_ID` を G列に保存し、次回は更新として再入稿
- Google ドキュメント本文の記法を、投稿時に装飾HTMLへ自動変換（マーカー / ボックス / FAQ）

## スクリプト プロパティ

Apps Script エディタの **プロジェクトの設定 → スクリプト プロパティ** に次を設定してください。

- `WP_BASE_URL`（WordPress がルートにある場合はルートドメイン、サブディレクトリにある場合はその格納ディレクトリまで含める。例: `https://example.com` / `https://example.com/wordpress`）
- `WP_USERNAME`
- `WP_APP_PASSWORD`

## シート列定義

- A: 入稿ステータス
- B: GoogleドキュメントURL
- C: WP_STATUS (`draft` / `publish` / `pending` など)
- D: WP_SLUG
- E: WP_CATEGORIES (カンマ区切り)
- F: WP_TAGS (カンマ区切り)
- G: WP_POST_ID
- H: WP_POST_URL
- I: 最終実行日時
- J: 最終結果
- K: エラーメッセージ

## 使い方

1. スプレッドシートにこのスクリプトを紐づける。
2. メニュー `WordPress連携` → `1行目にヘッダーを設定` を実行。
3. 必要な行を入力し、A列を `入稿待ち` にする。
4. メニュー `WordPress連携` → `入稿待ちを処理` を実行。

## 本文の自動装飾ルール

Google ドキュメント本文に以下の記法を書いておくと、投稿時にHTML装飾へ変換されます。

### 1) マーカー（蛍光ペン風）

- `==強調したい文章==` → `<mark>強調したい文章</mark>`
- H2ごとに0〜1箇所を目安に使ってください。

### 2) SWELLボックス装飾（アイコン付き）

以下の開始タグ〜終了タグで囲むと、対応するボックスHTMLに変換されます。

- `[POINT] ... [/POINT]` : ポイント・メリット（✅）
- `[CAUTION] ... [/CAUTION]` : 注意点・デメリット（⚠️）
- `[NOTE] ... [/NOTE]` : 補足説明・備考（💡）
- `[SUMMARY] ... [/SUMMARY]` : 要約・まとめ（📝）

### 3) キャプションボックス

- `[CAPTION:タイトル] ... [/CAPTION]`
- タイトル省略時は「補足情報」になります。
- 具体例 / 引用 / 参考情報などの補足に使えます。

### 4) FAQブロック

- `## よくある質問` セクション内の `Q:` / `A:` を自動でFAQブロックに変換します。
- 例:
  - `Q: 〇〇ですか？`
  - `A: はい、〇〇です。`

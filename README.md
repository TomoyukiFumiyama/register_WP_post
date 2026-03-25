# Spreadsheet -> WordPress 投稿 Apps Script

このリポジトリには、Google スプレッドシートに紐づける Apps Script (`Code.gs`) を置いています。

## できること

- 1行目に指定ヘッダーを設定
- A列が `入稿待ち` の行だけ処理
- B列の Google ドキュメントURLから本文とタイトルを取得
- WordPress REST API に投稿（新規/更新）
- 初回投稿時に `WP_POST_ID` を G列に保存し、次回は更新として再入稿

## スクリプト プロパティ

Apps Script エディタの **プロジェクトの設定 → スクリプト プロパティ** に次を設定してください。

- `WP_BASE_URL` (例: `https://example.com`)
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

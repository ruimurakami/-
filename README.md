# Sheets AI Assistant (Streamlit版)

自然言語の指示を Claude が pandas コードに変換し、確認の上で Google スプレッドシートへ反映するデモアプリです。
少人数でのテスト・ポートフォリオ公開を目的としたプロトタイプであり、商用運用向けの堅牢性(レート制限・CSRF対策等)は備えていません。

## 構成

- `sheets_core.py` — コアロジック(Claudeへの操作内容生成依頼、安全チェック、スプレッドシート読み書き、undo/redo用スナップショット)
- `streamlit_app.py` — Streamlit製のWebアプリ本体(ログイン、シート読み込み、指示入力、確認・実行、操作履歴)

## ローカルでの起動

1. 依存パッケージをインストール
   ```
   pip install -r requirements.txt
   ```
2. `.env` に以下を設定
   ```
   ANTHROPIC_API_KEY=...
   GOOGLE_APPLICATION_CREDENTIALS=credentials.json
   WEB_APP_PASSWORD=好きなパスワード
   ```
3. 起動
   ```
   streamlit run streamlit_app.py
   ```
4. ブラウザで自動的に開く画面(`http://localhost:8501`)で `WEB_APP_PASSWORD` を入力してログイン

## 使い方

1. 対象のGoogleスプレッドシートを、サービスアカウントのメールアドレス(`credentials.json` 内の `client_email`)に「編集者」として共有する
2. アプリにスプレッドシートのURLと、シート名(任意)を入力して「読み込む」
   - **一度読み込んだシートはセッション内で保持されるため、指示を変えるたびにURLを入力し直す必要はありません。**
   - シートの中身が変わった場合は「データを再読み込み」、別のシートを使う場合は「別のシートを読み込む」を押してください。
3. やってほしいことを入力して「内容を生成」
4. 生成された内容(pandasコードまたは数式)を確認し、問題なければ「実行する」で反映
   - 書き込み位置の指定がない指示は、現在開いているシート全体を更新します(デフォルト)。「新しいシートに」と明示した場合のみ別シートに書き込みます。
   - 「H1セルに書いて」のように位置を指定した場合は、そのセルだけを書き換え、他のデータは変更しません。
   - 「自動で反映されるように」「動的に」と指示すると、元データの変化に自動追従するスプレッドシート数式(SORT/FILTER/QUERYなど)として書き込みます。
   - 「Sheet2を削除して」「シート1〜シート4を削除して」のようにシートの作成・削除も指示できます。
5. 操作は履歴として記録され、「直前の操作を取り消す」「取り消しをやり直す」「すべての操作を取り消す」「指定した時点まで戻す」で管理できます。「元に戻して」のような自然文の指示でも取り消せます。

## Render へのデプロイ

このリポジトリには `render.yaml` を含めています。Render のダッシュボードで「New Blueprint」からこのリポジトリを指定するとサービスが作成されます。

デプロイ後、以下の環境変数をRenderのダッシュボードで設定してください(`sync: false` のため自動では入りません):

| 変数名 | 内容 |
| --- | --- |
| `ANTHROPIC_API_KEY` | Anthropic APIキー |
| `GOOGLE_APPLICATION_CREDENTIALS_JSON` | サービスアカウントJSON鍵の中身をそのまま1行の文字列として貼り付け |
| `WEB_APP_PASSWORD` | アプリへのログインパスワード(知人に共有する用) |

Render環境では `credentials.json` ファイルを配置できないため、`GOOGLE_APPLICATION_CREDENTIALS` ではなく `GOOGLE_APPLICATION_CREDENTIALS_JSON`(JSON文字列)を使う点に注意してください(`sheets_core.get_client()` が両方に対応しています)。

Streamlit CommunityCloud (streamlit.io) でも同様の環境変数を `st.secrets` 経由で設定してデプロイできます。

## セキュリティ上の注意

- 公開URLは共有パスワードのみで保護しています。パスワードは知人にのみ伝え、SNS等では公開しないでください。
- 生成されたコードは実行前に画面で確認できます。内容を確認してから「実行」してください。
- `FORBIDDEN_PATTERNS` による簡易チェックはありますが、完全な安全性を保証するものではありません。信頼できる相手とのテスト利用に留めてください。

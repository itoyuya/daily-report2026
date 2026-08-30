# CCBT日報システム（daily-report2026）

詳細はまず `README.md` を読むこと（構成表・セットアップ手順・PDF生成・氏名マッピングの全仕様がある）。

## 要点

- フォーム: https://itoyuya.github.io/daily-report2026/ （index.html を GitHub Pages で公開）
- `gas_code.js` / `gas_code_admin.js` は **このリポジトリが正本**。実行環境は各Googleスプレッドシートの Apps Script なので、編集後は該当スプレッドシートの Apps Script に手動で貼り付けてデプロイが必要（push だけでは反映されない）
- index.html の変更は push すれば GitHub Pages に反映される

## ワークフロー

- ブランチ運用: `feature/*` `docs/*` を切って main にマージ（既存の運用を踏襲）
- コミットは日本語1行要約
- スプレッドシートのデータ・数式は氏名キー（カタカナ）に依存している。**データ側の氏名表記を変えないこと**（PDF表示のみ `PDF_DISPLAY_NAME_MAP` で置換する設計）

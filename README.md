# テクニカルサポート業務日報システム

業務委託メンバーが日報を入力し、Google スプレッドシートにデータを蓄積。指定日のデータを別紙6様式のPDFとして出力するシステム。

## 構成

| ファイル | 内容 |
|---|---|
| `index.html` | 日報入力フォーム（GitHub Pages で公開） |
| `gas_code.js` | Google Apps Script コード（doPost + PDF生成） |
| `日報ヘッダー.csv` | スプレッドシートのヘッダー行 |

## フォーム URL

https://itoyuya.github.io/daily-report2026/

## セットアップ手順

### 1. Google スプレッドシート

1. 新規スプレッドシートを作成
2. `日報ヘッダー.csv` をインポート
3. シート名を「日報」に変更

### 2. Google Apps Script

1. スプレッドシートの「拡張機能 → Apps Script」を開く
2. `gas_code.js` の内容を貼り付け
3. `CONFIG` の `DRIVE_FOLDER_ID` を保存先フォルダのIDに設定
4. デプロイ → ウェブアプリ（実行: 自分、アクセス: 全員）
5. 発行されたURLを `index.html` の `GAS_URL` に設定して push

### 3. PDF 生成

Apps Script エディタで以下を実行：

```javascript
generateDailyReport("2026-04-06")
```

指定日の全メンバー分を1ファイル（`業務日報_YYYY-MM-DD.pdf`）にまとめて Google Drive に保存。

## スプレッドシートのカラム構成

| A | B | C | D | E | F | G | H | I |
|---|---|---|---|---|---|---|---|---|
| タイムスタンプ | 日付 | 氏名 | 開始時間 | 終了時間 | 業務タイトル | 業務内容 | 気づき・ふりかえり | 勤務時間 |

勤務時間は開始・終了時間から自動算出（単位: 時間）。

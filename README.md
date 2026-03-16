# BNI Chiyoda VisitorHost - Activeチャプター 名簿・割り振りシステム

BNI Activeチャプターの定例会運営を支援する、Google スプレッドシート上で動作する Google Apps Script (GAS) アプリケーションです。

## 概要

本システムは、毎週の定例会に向けた以下の準備作業を半自動化し、作業時間を大幅に削減します。

| 機能 | 説明 |
|------|------|
| **CSV取込・名簿作成** | SpreadingからエクスポートしたCSV（Shift_JIS）を解析し、名前補正・招待者マッチングを行い、ビジター様リストのシートとPDFを自動生成 |
| **ルーム割り振り** | ブレイクアウトルーム・オリエンテーションの割り振りをドラッグ＆ドロップUIで直感的に作成。Gemini AIによる自動割り振りにも対応 |
| **メール一括送信** | テンプレートに基づき、参加者全員への案内メールをプレビュー・編集・一括送信 |
| **PDF管理** | 作成済みのビジターリスト・割り振り表・メンバーブックPDFへのクイックアクセス |

## システム要件

- Google Workspace アカウント（Gmail, Google Drive, Google Sheets）
- [clasp](https://github.com/google/clasp)（ローカル開発・デプロイ時）
- Gemini API キー（AI自動割り振り機能を使用する場合）

## プロジェクト構成

```
BNI_Chiyoda_VisitorHost/
├── コード.js              # サーバーサイド全ロジック（GAS）
├── appsscript.json        # GAS マニフェスト（タイムゾーン・依存サービス定義）
├── dialog.html            # CSV取込・名簿作成ダイアログ
├── allocation.html        # ルーム割り振りダイアログ（D&D UI + AI連携）
├── email.html             # メール確認・一括送信ダイアログ
├── pdf_links.html         # 作成済みPDF確認ダイアログ
├── pdf.html               # メンバーリストOCRアップロード
├── memberbook.html        # メンバーブックPDFアップロード
├── holiday.html           # 休会日管理
├── template.html          # メールテンプレート設定
├── allocation_note.html   # 割り振り表の特記事項設定
├── visitor_host.html      # ビジターホスト設定
├── api_settings.html      # Gemini API・モデル設定
├── webapp_settings.html   # Webアプリ（送信元）設定
├── .claspignore           # clasp push 除外設定
├── .gitignore             # Git 除外設定
├── README.md              # 本ファイル（システム概要）
└── MANUAL.md              # 詳細な利用マニュアル
```

## セットアップ

### 1. clasp のインストールとログイン

```bash
npm install -g @google/clasp
clasp login
```

### 2. プロジェクトのクローンまたは作成

既存のGASプロジェクトに紐づける場合:

```bash
clasp clone <スクリプトID>
```

新規作成の場合:

```bash
clasp create --type sheets --title "BNI VisitorHost"
```

### 3. コードのプッシュ

```bash
clasp push
```

> `.claspignore` により、`README.md`、`MANUAL.md`、`.git/` などのドキュメント・開発用ファイルは自動的に除外されます。

### 4. Webアプリとしてデプロイ（メール送信元の固定）

特定のGoogleアカウントからメールを送信したい場合、Webアプリとしてのデプロイが必要です。

```bash
# 初回デプロイ
clasp deploy --description "メール送信用"

# デプロイIDの確認
clasp deployments
```

または GASエディタから: `デプロイ` > `新しいデプロイ` > `ウェブアプリ`（「次のユーザーとして実行: 自分」「アクセス: 全員」）

デプロイ後、取得したWebアプリURLをスプレッドシートの `名簿システム` > `Webアプリ(送信元)設定` に登録してください。

> **セキュリティ**: WebアプリURLはスクリプトプロパティに保存されるため、ソースコードやGitリポジトリには含まれません。

詳細な手順は [MANUAL.md の「Webアプリのデプロイ」](MANUAL.md#7-webアプリのデプロイ送信元アカウントの設定) を参照してください。

### 5. 初期設定

スプレッドシートを開き、メニューバーの **「名簿システム」** から各種設定を行ってください。詳細は [MANUAL.md](MANUAL.md) を参照してください。

## 使い方（毎週のワークフロー）

1. **事前準備**: SpreadingからCSV、メンバーリストPDFをダウンロード
2. **名簿作成**: `名簿システム` > `1. CSVから名簿・PDF作成` でビジターリストPDFを生成
3. **割り振り作成**: `名簿システム` > `3. ルーム・オリエン割り振り表` で割り振り表PDFを生成
4. **メール送信**: `名簿システム` > `2. メールの確認・一括送信` で案内メールを配信

## 使用技術

- **Google Apps Script (V8)** - サーバーサイドロジック
- **Google Drive API v3** - ファイル管理・OCR
- **Gemini API** - AI自動割り振り
- **HTML5 + CSS3 + Vanilla JS** - クライアントサイドUI（ドラッグ＆ドロップ等）

## ドキュメント

- [利用マニュアル (MANUAL.md)](MANUAL.md) - 画面ごとの操作手順、設定方法、トラブルシューティング

## ライセンス

Copyright Mitsunori KIMURA. All rights reserved.

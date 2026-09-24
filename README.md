# BNI Chiyoda VisitorHost - Activeチャプター 名簿・割り振りシステム

BNI Activeチャプターの定例会運営を支援する、Google スプレッドシート上で動作する Google Apps Script (GAS) アプリケーションです。

## 概要

本システムは、毎週の定例会に向けた以下の準備作業を半自動化し、作業時間を大幅に削減します。

| 機能 | 説明 |
|------|------|
| **CSV取込・名簿作成** | SpreadingからエクスポートしたCSVを解析し、名前補正・招待者マッチングを行い、ビジター様リストのシートとPDFを自動生成 |
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

### 3. デプロイ

```bash
clasp push
```

> `.claspignore` により、`README.md`、`MANUAL.md`、`.git/` などのドキュメント・開発用ファイルは自動的に除外されます。

### 4. 初期設定

スプレッドシートを開き、メニューバーの **「名簿システム」** から各種設定を行ってください。詳細は [MANUAL.md](MANUAL.md) を参照してください。

## ウェブアプリの更新（URLを変えずに）

ダイアログが表示されない環境向けに、ウェブアプリとしても公開できます。
`clasp push` はスクリプトを更新するだけで、**ウェブアプリには反映されません。**
公開中のデプロイを更新する必要があります。

### 1. デプロイIDを調べる

```bash
clasp deployments
```

```
2 Deployments.
- AKfycbwAAA...  @HEAD
- AKfycbxeZur-sqvWW_vr3i1A2VPJ8Bd4...  @5 - ウェブアプリ
```

`@HEAD` の行は**テスト用**なので使いません。`@数字` が付いている方が公開中のデプロイです。

> **URLから読み取ることもできます。**
> デプロイIDは、ウェブアプリURLの `/s/` と `/exec` の間の文字列そのものです。
> ```
> https://script.google.com/macros/s/AKfycbxeZur-.../exec
>                                    ~~~~~~~~~~~~~~~ これがデプロイID
> ```

### 前提：appsscript.json にウェブアプリの宣言が要る

```json
"webapp": {
  "executeAs": "USER_DEPLOYING",
  "access": "MYSELF"
}
```

**この宣言が無いと、`clasp deploy` はウェブアプリではなくライブラリとしてデプロイし、
元のURLでアクセスできなくなります。**
画面から作ったデプロイはその場の設定を使いますが、claspからのデプロイは
`appsscript.json` を見るためです。

| 項目 | 値 | 意味 |
|---|---|---|
| `executeAs` | `USER_ACCESSING` | **アクセスした人として実行**（現在の設定） |
| | `USER_DEPLOYING` | 自分として実行 |
| `access` | `ANYONE` | **Googleアカウントがある人全員**（現在の設定） |
| | `DOMAIN` | 同じ組織の人 |
| | `MYSELF` | 自分だけ |

**現在の組み合わせ（`USER_ACCESSING` ＋ `ANYONE`）は、共有する場合の安全な設定です。**
URLは誰でも開けますが、**中身はその人自身の権限で動く**ため、
スプレッドシートへのアクセス権が無い人には何も見えません。

> **`USER_DEPLOYING` ＋ `ANYONE` にはしないでください。**
> その組み合わせだと、URLを知っている人が**所有者の権限で**スプレッドシートを
> 操作できてしまいます。

### ⚠ 画面で設定を変えたら、この値も合わせること

`appsscript.json` の値と、画面でデプロイしたときの設定が食い違っていると、
**`clasp push` のたびに画面での設定が上書きされます。**
画面側を変えたときは、`clasp pull` して差分を確認し、
このファイルにも同じ値を書いてコミットしてください。

### 2. 同じデプロイを更新する

```bash
clasp push
clasp deploy -i <デプロイID> -d "更新内容のメモ"
```

`-i`（`--deploymentId`）を付けると**既存のデプロイを上書き**するので、**URLは変わりません。**
新しいバージョンが自動で作られ、それがそのURLに割り当てられます。

`-i` を付けずに `clasp deploy` を実行すると**新しいデプロイが作られ、URLも別物になります。**
ブックマークが効かなくなるので注意してください。

### まとめ（毎回この2行）

```bash
clasp push
clasp deploy -i AKfycbxeZur-sqvWW_vr3i1A2VPJ8Bd4iJ9QQ_UOlQSrmIa4LSdDsNDogomVsMhxkNQ-TEzXqA -d "更新"
```

### `clasp deploy -i` で更新されないとき

環境によっては `-i` を付けても反映されないことがあります。順に試してください。

**その1：バージョンを明示する**

```bash
clasp push
clasp version "更新内容のメモ"      # → 「Created version 7」のように番号が出る
clasp deploy -i <デプロイID> -V 7
```

`clasp deploy` 単体だと、バージョンの作成と割り当てをまとめて行おうとして
失敗することがあります。分けると通ることがあります。

**その2：反映されたか必ず確かめる**

```bash
clasp deployments
```

`@5` のような番号が上がっていれば成功です。変わっていなければ失敗しています。
画面側でも、ウェブアプリの下部に出ている **版の日付**で確認できます。

**その3：それでも駄目なら画面から**

`デプロイ` ＞ `デプロイを管理` ＞ 鉛筆アイコン ＞ バージョン `新バージョン` ＞ `デプロイ`。
確実です。**この場合も、画面で選んだ「実行するユーザー」「アクセスできるユーザー」が
`appsscript.json` と一致しているか確認してください。**

- `clasp --version` でバージョンを確認してください。clasp 3系では
  コマンド名や挙動が変わっています。

### うまくいかないとき

- **ライブラリになってしまう / URLで開けない**
  `appsscript.json` の `webapp` の宣言が抜けています。上記を追加し、
  `clasp push` してから `clasp deploy -i <ID>` をやり直してください。
  同じデプロイIDのままウェブアプリに戻ります。
- **メニューは出るがリンクの先が真っ白**
  ウェブアプリの画面は `googleusercontent.com` のiframeの中で動くため、
  `href="?p=..."` のような相対リンクはiframeのURLを基準にしてしまい、
  別の場所へ飛びます。リンクは必ず絶対URL（`ScriptApp.getService().getUrl()`）で
  組み立て、`target="_top"` を付けてください。

### 補足

- `-d`（説明）は省略できます。あとで `clasp deployments` を見たときに分かりやすいので付けるのがおすすめです。
- 特定のバージョンを割り当て直したいときは `-V <バージョン番号>` を併用します。
- コマンド名はclaspのバージョンで変わることがあります。`clasp --version` で確認してください。
- デプロイをやめるときは `clasp undeploy <デプロイID>` です。

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

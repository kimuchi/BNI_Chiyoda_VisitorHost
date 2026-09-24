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

サーバー側は `*.js`、画面は同名の `*.html`。
**GASはファイル名を拡張子抜きで見る**ため、`foo.js` と `foo.html` は同居できない。
そのためサーバー側のファイルには `_srv` を付けている（`tools/check_gas_names.py` が検査する）。

```
BNI_Chiyoda_VisitorHost/
├── コード.js                 # メニュー・名簿作成・PDF・メール・割り振り（本体）
├── appsscript.json           # GASマニフェスト（タイムゾーン・ウェブアプリの宣言）
│
├── webapp_srv.js             # ウェブアプリの入口（doGet）と画面一覧
│   └ webapp_home.html        #   トップページ
├── home_srv.js               # メニュー代わりのホーム画面
│   └ menu_home.html
│
│  ── 毎週の作業 ──
├── dialog.html               # CSV取込・名簿作成
├── allocation.html           # ルーム・オリエン割り振り表
├── email.html                # 案内メールの確認・一括送信
├── pdf_links.html            # 作成済みPDFの確認
├── archive.html              # シートの整理（アーカイブ）
│
│  ── スライド・冊子 ──
├── slides_visitor_srv.js     # ビジター・ゲスト・代理スライド
│   └ slides_visitor.html
├── member_presen_srv.js      # メンバープレゼンスライド
│   └ member_presen.html
├── meeting_slides_srv.js     # 定例会スライドの自動更新
│   └ slides_meeting.html
├── routine_srv.js            # ルーティンチェックシートから、その日の決めごとを読む
├── memberbook_srv.js         # メンバーブック（配布PDFの登録）
│   ├ memberbook.html
│   ├ memberbook_editor.html  #   冊子の編集画面
│   └ memberbook_render.html  #   冊子の組版
├── zoom_guide.html           # Zoom入室案内
├── ooxml.js                  # pptx(OOXML)を文字列で書き換える共通処理
│
│  ── 名簿・素材 ──
├── member_master_srv.js      # メンバー名簿マスタ（全機能の正本）
│   ├ member_master.html
│   └ member_photos.html
├── spreading_srv.js          # Spreadingから名簿を更新
│   └ spreading.html
├── assets.js                 # 素材フォルダ・写真・テンプレートの保管
│   ├ asset_settings.html
│   └ template_files.html
├── big_templates_srv.js      # 大きなpptxをDriveリンクで登録・加工
│   └ big_templates.html
├── pdf.html                  # メンバーリスト(OCR)の取り込み
│
│  ── その他 ──
├── roulette_srv.js           # 抽選ルーレットのビジター招待数
├── holiday.html              # 休会日
├── template.html             # メールテンプレート
├── allocation_note.html      # 割り振り表の特記事項
├── visitor_host.html         # ビジターホスト・優先順位
├── ai_documents.html         # AI参考資料
├── api_settings.html         # Gemini API・モデル
│
├── MANUAL.md                 # 利用マニュアル（正本）
├── manual.html               # ↑から生成。画面に出すもの
├── templates/                # ビジター用pptxテンプレート
└── tools/                    # 検査・生成スクリプト（GASには送らない）
    ├── build_manual.py       #   MANUAL.md → manual.html
    ├── check_gas_names.py    #   名前の衝突・参照先HTMLの検査
    ├── check_html.py         #   divの対応・スクリプトの構文
    ├── check_manual_links.py #   マニュアルの導線 ↔ 実際のメニュー
    ├── check_member_presen.py#   メンバープレゼン生成の通し検査
    ├── check_meeting_slides.js  # 第○回・日付の書き換えとコアバリューのページ
    ├── check_routine.js      #   ルーティンチェックシートの読み取り
    ├── check_meeting_output.js  # 前半スライドを実際に生成してみる
    ├── build_member_presen_template.py  # メンバープレゼンのテンプレートを作り直す
    ├── build_meeting_first_template.py  # 前半スライドの出力に差し込み口を入れる
    ├── build_meeting_second_template.py # 後半スライドの出力に差し込み口を入れる
    └── extract_slide_media.py #   スライドの音楽・動画を曲名で取り出す
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

**いちばん確実なのは、実際に使っているウェブアプリURLから読み取る方法です。**
デプロイIDは、URLの `/s/` と `/exec` の間の文字列そのものです。

```
https://script.google.com/macros/s/AKfycbxeZur-sqvWW_vr3i1A2VPJ8Bd4.../exec
                                   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ これがデプロイID
```

Apps Scriptの `デプロイ` ＞ `デプロイを管理` にも「デプロイ ID」として表示されています。

> **⚠ `clasp deployments` の一覧から選ぶときは要注意**
>
> ```
> 3 Deployments.
> - AKfycbwWgjw...  @HEAD
> - AKfycbxeZur...  @5 - ウェブアプリ
> - AKfycbzQQQQ...  @2 - 古いもの
> ```
>
> - **`@HEAD` の行は使えません。**テスト用で、バージョンを割り当てられません。
> - 過去のデプロイも並ぶので、**別の行を選ぶと**
>   `Requested entity was not found.` になります。
> - デプロイIDはどれも `AKfycb` で始まるため、**先頭数文字だけ見ると見分けがつきません。**
>   必ず全体を見比べてください。

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

### `Requested entity was not found.` と出るとき

**デプロイIDが違います。**指定したIDのデプロイが存在しません。

実際に使っているウェブアプリURLの `/s/` と `/exec` の間をコピーし直してください。
`clasp deployments` の一覧から選んだ場合、`@HEAD` の行や古いデプロイの行を
拾ってしまっていることがよくあります。

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

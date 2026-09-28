# BNI Chiyoda VisitorHost - BNIチャプター 名簿・割り振りシステム

BNIチャプターの定例会運営を支援する、Google スプレッドシート上で動作する Google Apps Script (GAS) アプリケーションです。
もとは Activeチャプター用に作ったもので、何も設定しなければ Activeチャプターの値で動きます。
チャプター名・リージョン・期の番号・定例会の曜日と回数・プレゼンの秒数は `⚙️ 設定` > `チャプター（名前・期・定例会の回数・プレゼンの秒数）` で変えられます
（`chapter_srv.js`）。

## 概要

本システムは、毎週の定例会に向けた以下の準備作業を半自動化し、作業時間を大幅に削減します。

| 機能 | 説明 |
|------|------|
| **CSV取込・名簿作成** | SpreadingからエクスポートしたCSVを解析し、名前補正・招待者マッチングを行い、ビジター様リストのシートとPDFを自動生成 |
| **ルーム割り振り** | ブレイクアウトルーム・オリエンテーションの割り振りをドラッグ＆ドロップUIで直感的に作成。Gemini AIによる自動割り振りにも対応 |
| **メール一括送信** | テンプレートに基づき、参加者全員への案内メールをプレビュー・編集・一括送信 |
| **PDF管理** | 作成済みのビジターリスト・割り振り表・メンバーブックPDFへのクイックアクセス |
| **ビジター情報の投稿文** | 参加者シートから「第○回(M/D)定例会ビジター情報」の投稿文（入金済み／未入金つき）を作ってコピー |
| **役職ごとの入力（定例会の準備）** | 13の役職ごとの画面から、次回の定例会について入力してルーティンチェックシートに保存。前回の内容・名簿・参加者シートから推定した初期値入り。どの役職の必須項目（期日のある項目）が足りないかを一覧で表示。役職の担当者（チームのリーダー）と各チームのサポートメンバー（ビジターホスト・メンバーシップ委員会など）は半期（4〜9月・10〜3月）ごとに登録（次の期も前もって登録できる）。期が替わるとメンバー名簿の「役職」にも反映。割り振りのビジターホストは「役職から読み取る」で読み取って手で直せる |
| **スピーカーローテーション** | 書記兼会計のメインプレゼンの順番の管理（並び順・対象外・休会日を飛ばす・確定した回・Facebookの投稿文と添付する表の画像・MP選出ツールからの取り込み）。前半スライドの表とメインプレゼンの初期値に使う |
| **事前MTG（朝イチMTG）のパワポ** | 役職ごとの入力（今週の共有事項・お願い事項・人数・代理・欠席など）から、まとめ1枚＋役職ごとのページ（担当者の写真つき・短いもの同士は2人で1枚）を作成。「BNI週次役職情報共有 パワポ自動生成ツール」と同じ中身とデザイン。ひな形は同梱の既定のものか、登録したpptx |
| **ビジター・代理スライド** | 参加者シートから紹介スライド・ビジタープレゼンのPowerPointを作成 |
| **定例会スライド（前半）** | ルーティンチェックシートと名簿から、開催回・コアバリュー・メインプレゼン・委員会報告・一般規定を差し込み。スピーカーローテーションの5回分の表（以前は画像）も作る。ウィークリープレゼン（業種区分ごとの扉ページと個人ページ、ウィークリー／スタートアッププレゼンの秒数のカウントダウン・自動送り）とアンバサダー・ディレクターのページも1回で入る。役職のメンバー紹介（リーダーシップチーム・コーディネーター・メンバーシップ委員会・ビジターホストチーム・各サポートチーム）は、その期の役職・チームのお名前・写真・会社名・カテゴリーに入れ替える |
| **定例会スライド（後半）** | 更新状況（書記兼会計による報告の表）・推薦のことば（何組でも。アフター分は抽選コーナーのあと）・抽選コーナー・リファーラル発表（名簿のNo.順）・音楽の差し替えと音量 |
| **メンバーブック・Zoom案内** | 名簿から配布用の冊子（PDF/HTML）とZoom入室案内を作成 |
| **公式ファイルから雛形**（初期設定のときだけ） | BNI 素材フォルダの画面の小さなリンクから開く。BNIメンバーがダウンロードできる定例会の公式スライド（.pptx）を共有ドライブに置いてリンクを設定すると、ビジター紹介・ゲスト紹介・代理紹介・ビジタープレゼン・メンバープレゼン・定例会の前半と後半の雛形を作って「01_テンプレート」に保存・登録。音・動画はそのまま、カウントダウンはチャプターの設定の秒数（ウィークリー・スタートアッププレゼン・ビジタープレゼン・リファーラル発表）。次からは作った雛形をPowerPointで直して使う |
| **ウェブアプリ** | ダイアログが開けない環境でも、全機能を普通のタブで使える |

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
├── chapter_srv.js            # チャプターの設定（名前・リージョン・期の番号・定例会の曜日と回数）
│   └ chapter_settings.html
├── setup_srv.js              # 初回の準備（空のスプレッドシートで開いたときにシートを作る）・足りないシートを作る
│
│  ── 毎週の作業 ──
├── dialog.html               # CSV取込・名簿作成
├── allocation.html           # ルーム・オリエン割り振り表
├── email.html                # 案内メールの確認・一括送信
├── pdf_links.html            # 作成済みPDFの確認
├── visitor_post_srv.js       # ビジター情報の投稿文（入金済み／未入金つきの一覧）
│   └ visitor_post.html
├── role_input_srv.js         # 役職ごとの入力（ルーティンチェックシートの担当の列から項目を作り、推定・保存・入力状況）
│   └ role_input.html         #   入力状況の一覧・役職ごとの入力・スピーカーローテーション（?p=role_input&role=vice など）
├── speaker_rotation_srv.js   # スピーカーローテーション（メインプレゼンの順番。前半スライドの表もここで作る）
├── premtg_srv.js             # 事前MTG（朝イチMTG）のパワポ（画面は role_input.html の一覧。入口は役職ごとの入力と同じ）
│   └ premtg_template.html    #   既定のひな形（docs/templates/BNI_テンプレート_事前MTG.pptx をbase64にしたもの）
├── archive.html              # シートの整理（アーカイブ）
│
│  ── スライド・冊子 ──
├── slides_visitor_srv.js     # ビジター・ゲスト・代理スライド
│   └ slides_visitor.html
├── member_presen_srv.js      # メンバープレゼン（前半のウィークリープレゼンのページ。画面は前半に統合）
├── meeting_slides_srv.js     # 定例会スライド（前半・後半）
│   ├ slides_meeting_first.html   #   前半の画面（メンバーのページ・アンバサダー・ディレクターも）
│   ├ slides_meeting_second.html  #   後半の画面（開くと後半テンプレートを読み込む）
│   └ slides_meeting_common.html  #   前半・後半で共通の部品（読み込み中の表示など。include で読む）
├── referral_srv.js           # リファーラル発表のページを人数ぶんに増やす
├── role_intro_srv.js         # 前半の役職のメンバー紹介（その期の役職・チームの方を入れる。差し込み口・ページの作りで見分ける）
├── meeting_pages_srv.js      # 前半の新メンバー・更新メンバー（1人1枚）・バイスプレジデントによる報告・ネットワーキングリーダー
├── splice_srv.js             # 別のpptxのページを差し込む（メンバープレゼン → 前半スライド）
├── routine_srv.js            # ルーティンチェックシートから、その日の決めごとを読む
├── slides_layout.html        # 会社名・カテゴリーの組版とメンバーのページの並び（前半・後半で共通。include で読む）
├── memberbook_srv.js         # メンバーブック（配布PDFの登録・編集画面から1人ずつ名簿に保存）
│   ├ memberbook.html
│   ├ memberbook_editor.html  #   冊子の編集画面
│   └ memberbook_render.html  #   冊子の組版
├── talk_script_srv.js        # トークスクリプト（台本）：ひな形の {…} にその回の担当者・参加者・チェックシートの内容を入れてシートに作る
│   ├ talk_script.html        #   台本を作る／ひな形の編集
│   └ talk_script_default.js  #   既定のひな形（docs/samples/トークスクリプトのひな形.tsv と同じ内容）
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
├── official_srv.js           # 公式ファイル（約100MB）から雛形を作って登録（使う部品だけをDriveから範囲を指定して読む）
│   ├ official_build_srv.js   #   雛形ごとの組み立て（使うページの選び方・差し込み口の入れ方）
│   └ official_templates.html
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
├── manual.html               # ↑から生成。画面に出すもの（画像も埋め込む）
├── docs/images/              # マニュアルのスクリーンショット（WebP。架空のデータで撮ったもの。
│                             #   Google認証の auth*.webp は実際の画面で、メールアドレスと顔写真を見本に置き換えたもの）
├── docs/samples/             # 見本（チャプター名を {チャプター} にしたトークスクリプトのひな形など）
├── docs/templates/           # スライドの雛形の決まり（docs/templates/README.md）と事前MTGの見本。
│                             #   ビジター用・メンバープレゼン・定例会の雛形は、設置者が置いた公式ファイルから作る（リポジトリには入れない）
├── docs/SETUP.md             # 初期導入の手順（新しいチャプターが空のスプレッドシートから始めるとき）
└── tools/                    # 検査・生成スクリプト（GASには送らない）
    │  ── 毎回走らせる検査 ──
    ├── check_gas_names.py    #   名前の衝突・参照先HTMLの検査
    ├── check_html.py         #   divの対応・スクリプトの構文・共通部品の読み込み漏れ
    ├── check_manual_links.py #   マニュアルの導線 ↔ 実際のメニュー
    ├── build_manual.py       #   MANUAL.md → manual.html（docs/images の画像を埋め込む）
    ├── make_manual_shots.js  #   マニュアルのスクリーンショットを架空のデータで撮り直す（Chromium・Pillow を使う）
    ├── clean_pptx_meta.py    #   pptxのファイルの情報（作成者・コメントの作成者・変更の記録・縮小画像）を消す／--check で確かめる
    ├── check_memberbook.js   #   メンバーブック（Chromium で組版を測る：長い文字を枠に収める・会社での役職とBNIの役職／
    │                         #     編集画面の「反映」ですぐ名簿に保存／期ごとのプレジデント設定）。Googleフォントを写したディレクトリを渡すとその書体で測る
    ├── check_allocation_ai.js #  割り振り表の「AIに提案させる」（待機メンバーの情報はメンバー名簿から文字で渡す・送らない情報・返事の反映）
    │  ── 画面をNodeで動かす（ブラウザ無し）──
    ├── lib_minidom.js        #   画面をNodeで動かすための簡易DOM
    ├── check_meeting_dialog.js  # 定例会スライドの画面（前半・後半・ウィークリープレゼン）
    ├── check_visitor_post.js #   ビジター情報の投稿文（入金の読み方・文面・画面）
    ├── check_role_input_dialog.js # 役職ごとの入力の画面（一覧・入力・保存までサーバーとつないで通す）
    ├── check_webapp_pages.js #   ウェブアプリの画面一覧（?p= での開き方・トップページのリンク）
    ├── check_save_pages.js   #   設定の画面（休会日など）は保存しても画面が残り、ウェブアプリでは「← メニューに戻る」が出る
    ├── check_chapter.js      #   チャプターの設定（名前・期の付け直し・定例会の曜日と回数・初回の準備）
    ├── check_setup.js        #   初回の準備（空のスプレッドシート・チェックシートを作る・前の期を写す）
    ├── check_official_dialog.js # 公式ファイルから雛形を作る画面（リンクの確かめ・登録状況・順に作る・作り直す前の確認）
    │  ── ルーティンチェックシート（実物をExcelに書き出したもので確かめる）──
    │     書き出した routine.json・members.json はリポジトリに入れない（名簿にあたるため）。
    │     検査の中の氏名は架空のもの（実物を使うときは、書き出したものの氏名も同じ架空の氏名に置き換えてから使う）
    ├── routine_dump.py       #   xlsx → routine.json
    ├── check_routine.js      #   開催回・コアバリュー・メインプレゼンなどの読み取り
    ├── check_weekly_start.js #   ウィークリープレゼンの始まりの業種区分（記載・前回からの繰り上げ）
    ├── check_role_input.js   #   役職ごとの入力（項目・推定・入力状況・保存・行の追加・衝突）
    ├── lib_role_fixture.js   #   役職ごとの入力・ローテーション・事前MTGの検査で使う、書き込める偽のシートと参加者シート
    ├── check_speaker_rotation.js # スピーカーローテーション（ツールと同じ割り当て・休会日・確定した回・取り込み）
    ├── check_premtg.js       #   事前MTGのパワポ（人数の数え方・まとめ・役職のページ・写真・帯の色・登録したひな形）
    ├── check_talk_script.js  #   トークスクリプト（差し込みの中身・敬称・シート・ひな形の保存・画面）
    ├── lib_zip.js            #   ↑で使う、pptxの展開・再梱包
    ├── check_countdown.js    #   カウントダウンの秒数を変えても、ベル・動画（音）の指示が残る（ビジタープレゼンの雛形も）
    ├── check_role_intro.js   #   役職のメンバー紹介（作り物のページで：ローマ字・枠の見分け方・替わった方だけ入れる・非表示・写真）
    ├── build_premtg_template.py # 事前MTGの見本pptxから既定のひな形（と premtg_template.html）を作る
    │  ── メンバープレゼン ──
    ├── check_member_presen.py#   テンプレートに対する通し検査（巡回の検査と↓の4つを順に呼ぶ）
    ├── mp_harness_prepare.py #     テンプレートを展開し、写真の見本を用意
    ├── mp_plan.js            #     画面の組版ロジックで items.json を作る
    ├── mp_harness.js         #     本番のコードでpptxのパーツを作る
    ├── mp_harness_check.py   #     出来上がりを組み立てて中身を確かめる
    ├── mp_check_rotation.js  #   業種区分の巡回（従来の計算）を元ツールと突き合わせる
    ├── mp_preview.py         #   出来上がりの配置を確認用のHTMLにする
    ├── build_member_presen_template.py  # テンプレートを作り直す（61枚 → 2枚）
    ├── mp_make_template.js   #     ↑のうち、スライド本文の書き換え
    │  ── 定例会スライド（前半・後半）──
    ├── check_meeting_slides.js  # 第○回・日付の書き換えとコアバリューのページ
    ├── check_meeting_output.js  # 前半・後半スライドを実際に生成してみる
    ├── mtg_zip_check.py      #   ↑の出力をpptxに固めて中身を確かめる
    ├── build_meeting_first_template.py  # 前半スライドの出力に差し込み口を入れる
    ├── make_meeting_first_template.js   #   ↑のうち、スライド本文の書き換え
    ├── build_meeting_second_template.py # 後半スライドの出力に差し込み口を入れる
    ├── make_meeting_second_template.js  #   ↑のうち、スライド本文の書き換え
    ├── lib_meeting_tokens.js #   前半・後半で共通の差し込み口の入れ方
    ├── extract_slide_media.py #  スライドの音楽・動画を曲名で取り出す
    │  ── 公式ファイルから雛形（公式ファイルはリポジトリに入れない）──
    ├── check_official.js     #   公式ファイルから7種類の雛形を実際に作り、中身・秒数・音の指示・登録を確かめる
    ├── lib_gas_fake.js       #   ↑で使う、GASの Utilities・UrlFetchApp（Driveの範囲指定の読み出し）・DriveApp の代わり
    ├── pptx_integrity.py     #   pptxの部品のつながり（関係・種類の登録・図形の番号・アニメーションの指す先）
    │  ── ビジター用テンプレート ──
    └── extract_visitor_templates.py  # 運用中のpptxから1枚ずつ取り出す（自分のチャプターの雛形を使うとき）
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

**新しいチャプターで、空のスプレッドシートから始めるとき**は、[docs/SETUP.md](docs/SETUP.md)（初期導入の手順）を見てください。
空のスプレッドシートで初めて開くと、使うシート（メンバー名簿・業種区分マスタ・休会日）を自動で作り、
チャプターの設定を保存したときにルーティンチェックシートを作ります（`setup_srv.js`）。

## ウェブアプリの更新（URLを変えずに）

ダイアログが表示されない環境向けに、ウェブアプリとしても公開できます。
`clasp push` はスクリプトを更新するだけで、**ウェブアプリには反映されません。**
公開中のデプロイを更新する必要があります。

### 1. デプロイIDを調べる

**いちばん確実なのは、実際に使っているウェブアプリURLから読み取る方法です。**
デプロイIDは、URLの `/s/` と `/exec` の間の文字列そのものです。

```
https://script.google.com/macros/s/AKfycbwwG3oVenqKtbf0uQtdO1VfORDE.../exec
                                   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ これがデプロイID
```

Apps Scriptの `デプロイ` ＞ `デプロイを管理` にも「デプロイ ID」として表示されています。

**今のURLは、定例会情報共有シートに載っているもの（`AKfycbwwG3oV…`）です。**
以前は `AKfycbxeZur…` に出していました。古いIDに出しても、メンバーが開く画面は変わりません。

> **⚠ `clasp deployments` の一覧から選ぶときは要注意**
>
> ```
> 4 Deployments.
> - AKfycbwWgjw...   @HEAD
> - AKfycbwwG3oV...  @9 - ウェブアプリ（今のURL）
> - AKfycbxeZur...   @5 - 以前のウェブアプリ
> - AKfycbzQQQQ...   @2 - 古いもの
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
  "executeAs": "USER_ACCESSING",
  "access": "ANYONE"
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
git pull                              # GitHub の最新を取ってくる
clasp push                            # Apps Script に送る（この時点ではウェブアプリはまだ古いまま）
clasp version "更新内容のメモ"        # 送った中身を番号つきで保存 → 「Created version 34」のように番号が出る
clasp deploy -i <デプロイID> -V 34    # その番号を、今のURLに割り当てる
```

**この順番が大事です。** `clasp version` は「最後に `clasp push` した中身」を保存します。
`git pull`・`clasp push` より先に `clasp version` を作ると、新しいコードが入りません。
`-V` には、直前の `clasp version` で出た番号を書きます（毎回1つずつ増えます）。

`-i`（`--deploymentId`）を付けると**既存のデプロイを上書き**するので、**URLは変わりません。**
`-i` を付けずに `clasp deploy` を実行すると**新しいデプロイが作られ、URLも別物になります。**
ブックマークが効かなくなるので注意してください。

`-V` を付けずに `clasp deploy -i <デプロイID> -d "メモ"` とすると、バージョンの作成と割り当てを
1回で行います。ただし環境によっては反映されないことがあるため、上のように分けるほうが確実です。

### まとめ（毎回この4行）

```bash
git pull
clasp push
clasp version "V2.5"     # → Created version 34（番号は毎回1つずつ増える）
clasp deploy -i AKfycbwwG3oVenqKtbf0uQtdO1VfORDEXvzMPsZoA3_YiZkqGXt_iiu-RAX9hzBuJjvw7vGu -V 34
```

終わったら、ウェブアプリのトップページの下の **「版 …」** が `コード.js` の `SYSTEM_VERSION_` と
同じになっているか確かめます。`clasp version` の説明に版（例: `"V2.5 2026-09-27o"`）を入れておくと、
`デプロイを管理` の一覧でどのコードのバージョンか分かります。

### `Requested entity was not found.` と出るとき

**デプロイIDが違います。**指定したIDのデプロイが存在しません。

実際に使っているウェブアプリURLの `/s/` と `/exec` の間をコピーし直してください。
`clasp deployments` の一覧から選んだ場合、`@HEAD` の行や古いデプロイの行を
拾ってしまっていることがよくあります。

### デプロイしたのに新しくならないとき

順に確かめてください。

**その1：順番を確かめる**

`clasp version` の前に `git pull` と `clasp push` をしたか確かめてください。
`clasp push` より前に作ったバージョンには、新しいコードが入っていません。
その場合は `clasp push` → `clasp version` → `clasp deploy -i <デプロイID> -V <新しい番号>` をやり直します。

**その2：反映されたか必ず確かめる**

```bash
clasp deployments
```

今のURLのIDの行が、`@34` のように直前に作った番号になっていれば成功です。変わっていなければ失敗しています。
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
  `clasp push` → `clasp version` → `clasp deploy -i <ID> -V <番号>` をやり直してください。
  同じデプロイIDのままウェブアプリに戻ります。
- **メニューは出るがリンクの先が真っ白**
  ウェブアプリの画面は `googleusercontent.com` のiframeの中で動くため、
  `href="?p=..."` のような相対リンクはiframeのURLを基準にしてしまい、
  別の場所へ飛びます。リンクは必ず絶対URL（`ScriptApp.getService().getUrl()`）で
  組み立て、`target="_top"` を付けてください。

### 補足

- `-d`（説明）は省略できます。あとで `clasp deployments` を見たときに分かりやすいので付けるのがおすすめです。
- 前のバージョンに戻したいときは、前の番号を `-V` に書きます（例: `clasp deploy -i <ID> -V 33`）。
  URLはそのままで、中身だけ前の状態に戻ります。
- コマンド名はclaspのバージョンで変わることがあります。`clasp --version` で確認してください。
- デプロイをやめるときは `clasp undeploy <デプロイID>` です。

## 使い方（毎週のワークフロー）

1. **事前準備**: SpreadingからCSVをダウンロード
2. **名簿作成**: `名簿システム` > `1. CSVから名簿・PDF作成` でビジターリストPDFを生成
3. **割り振り作成**: `名簿システム` > `3. ルーム・オリエン割り振り表` で割り振り表PDFを生成
4. **メール送信**: `名簿システム` > `2. メールの確認・一括送信` で案内メールを配信
5. **投稿文**: `名簿システム` > `5. ビジター情報の投稿文` でグループへの投稿文をコピー
6. **役職ごとの入力**: 各役職の方が `名簿システム` > `📝 役職ごとの入力（定例会の準備）`（またはウェブアプリ）で
   次回の定例会について入力。`入力状況の一覧・事前MTGのパワポ` でどの役職が入力済みかを確かめ、
   同じ画面で事前MTGのパワポを作る
7. **スライド**: `名簿システム` > `📊 スライド・冊子をつくる` から、ビジター・代理スライドと
   定例会スライド（前半）・（後半）をそれぞれ作成。前半にはメンバーのページも1回で入る

画面ごとの操作は [MANUAL.md](MANUAL.md) を参照してください。

### 更新が反映されたかの確かめ方

`コード.js` の `SYSTEM_VERSION_`（例: `2026-09-28a`）が、ウェブアプリのトップページの下と
「定例会スライド（前半）」「定例会スライド（後半）」の画面の右上に **「版 …」** として出ます。
機能を変えたらこの値を変え、`clasp push`・`clasp version`・`clasp deploy` のあとに画面で確かめてください。

## 使用技術

- **Google Apps Script (V8)** - サーバーサイドロジック
- **Google Drive API v3** - ファイル管理・OCR
- **Gemini API** - AI自動割り振り
- **HTML5 + CSS3 + Vanilla JS** - クライアントサイドUI（ドラッグ＆ドロップ等）

## ドキュメント

- [利用マニュアル (MANUAL.md)](MANUAL.md) - 画面ごとの操作手順、設定方法、トラブルシューティング

## ライセンス

Copyright Mitsunori KIMURA. All rights reserved.

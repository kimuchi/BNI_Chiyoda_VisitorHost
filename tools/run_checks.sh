#!/bin/bash
# 検査一式。変更したら、コミットの前に必ず走らせる（1つでも NG なら終了コード 1）。
#
#   bash tools/run_checks.sh
#
# ルーティンチェックシートを書き出したもの（tools/routine_dump.py）と名簿（members.json）があれば、
# それを使う検査（役職ごとの入力・ローテーション・事前MTG・台本・初回の準備など10本）も走らせる。
# どちらもリポジトリには入れない（名簿にあたるため）。手元に置いて場所を渡す:
#
#   ROUTINE_JSON=/どこか/routine.json MEMBERS_JSON=/どこか/members.json bash tools/run_checks.sh
#
# メンバーブックの組版を Google フォントで測るときは、フォントを写したディレクトリも渡せる（FONTS_DIR=…）。
# 必要なもの: Node.js・Python 3（openpyxl は書き出しのときだけ）・Playwright の Chromium（メンバーブックの検査）。

cd "$(dirname "$0")/.." || exit 1
ng=0
run() {
  local label="$1"; shift
  local out rc
  out=$("$@" 2>&1); rc=$?
  if [ $rc -eq 0 ]; then
    echo "OK  $label"
  else
    echo "NG  $label"
    echo "$out" | grep -v '^\s*at ' | tail -20 | sed 's/^/      /'
    ng=1
  fi
}

echo "── 毎週の流れ（名簿・PDF・メール・割り振り表）"
run "毎週の流れ（候補→CSV→名簿とPDF→再編集→次の週→抽選ルーレット→メール→割り振り表→PDFの確認）" node tools/check_weekly_flow.js
run "ビジターリスト・割り振り表のシートとPDF（作った日は隠さない・真っ白のPDFを作らない）" node tools/check_visitor_pdf.js
run "CSVから名簿・PDF作成の画面（ブラウザ：過ぎた回の再編集はその回に書き戻す・Shift_JIS のCSV・リンクで共有できなかったとき）" node tools/check_csv_dialog.js
run "メールの確認・一括送信の画面（ブラウザ：リンク・キャンセル・記号・二重に送らない・代理送信が動かないとき・返事を受け取れないとき・Web App のテスト）" node tools/check_email_dialog.js
run "ビジター情報の投稿文" node tools/check_visitor_post.js
run "割り振り表のAI（待機リストに居ないビジターホスト・知らない番号・確度とメモ）" node tools/check_allocation_ai.js
run "割り振り表の画面（ブラウザ：読み込んだ回に書き込む・Spreadingの文字を動かさない）" node tools/check_allocation_dialog.js
run "期の替わり目の担当者と名簿の役職（未登録の期で役職を消さない・取り込みを前の期に戻さない・設定の保存で期をずらさない）" node tools/check_role_terms.js
run "スピーカーローテーション（保存していないとき：見た日で発表者が変わらない）" node tools/check_rotation_unsaved.js
run "メインプレゼンターの画像（ブラウザ：写真・お名前・ご紹介して欲しいカテゴリー・回を替える・保存・書体は Meiryo UI）" node tools/check_mp_image.js
run "スピーカーローテーションの画面で定例会を変える（ブラウザ：画面はそのまま・ご案内する回と画像がその日の回に・戻ると選んだ日の書記兼会計の入力・続けて変えたとき）" node tools/check_rotation_meeting.js
run "事前MTGのパワポの役職のページ（「書記兼会計より」などの行からも作る・入力済みの数え方・XMLに書けない文字）" node tools/check_premtg_roles.js
run "事前MTGのパワポの割り振り表のページ（ビジターホストコーディネーターのすぐあと・ビジターの表だけ・字を小さく／ページを分ける・割り振り表が無い日）" node tools/check_premtg_allocation.js
run "本日の招待者を参加者シートから（ビジター・ゲストの招待者・代理とキャンセルの方は入れない・トークスクリプト）" node tools/check_inviters.js
run "事前MTGの熱烈歓迎のページ（新入会の方ごと・まとめのあと・写真・土台の違うひな形はマスターごと写す・pptxとして壊れていない）" node tools/check_welcome.js
run "ルーティンチェックシートの名前の照合（同じ名字・似た氏名・名簿に無い新メンバー・かなのお名前・カテゴリーの書き方）" node tools/check_routine_names.js
run "メンバー名簿の画面（ブラウザ：読み込めていないときは保存させない・ほかの方の変更を消さない・列を足したシート・貼り付けの改行）" node tools/check_member_master_dialog.js
run "PowerPointテンプレートの登録（開けない登録を未登録と出さない・素材フォルダを開けない方は止める）" node tools/check_template_files.js
run "メンバー写真（アップロードで写真索引を作り直す＝新しい方の写真がすぐ見つかる・同じ方の新しい写真・HEIC・ブラウザの画面の知らせ）" node tools/check_member_photos.js

echo "── スライド"
run "定例会スライド（指示どおりか）" node tools/check_meeting_slides.js
run "ビジター・代理スライド（「$」の入った名前・テンプレートの2ページ目・消した枠・キャンセルの方・会社名44pt／カテゴリー32pt・どのページも自動で次へ進めない・ファイル名）" node tools/check_visitor_slides.js
run "会社名・カテゴリーの組版（メンバー・リファーラル：44pt／32ptにそろえる・入らなければ同じ大きさで2行・語の途中で折らない・名簿の改行を外す）" node tools/check_layout_sizes.js
run "リファーラル発表のカテゴリー（1行は32ptのまま・2行は前面に出し数字にかからないところまで・カウントダウンを下げる／数字の箱の上を縮める・会社名も2行の方も32pt・帯の外や下の文字に隠れない・名簿の改行）" node tools/check_referral_category.js
run "欠席の方のページを作らない（ウィークリープレゼン・リファーラル発表：欄の読み方・名字を空白で並べた書き方（事前MTGの人数も）・名簿の方に合わせる・NEXT／次の発表者・外す・足す）" node tools/check_absent_skip.js
run "推薦のことばの受け取ったスライド（ブラウザ：pptx・pdf・画像を画像にして置く・その組のページのすぐあとに入れる・開催日ごとに覚える・準備中は作成させない・置けない理由は組の下に）" node tools/check_reco_slides.js
run "後半の音楽を推薦のことばで止める（前のサウンドを停止・鳴らす枚数を直す・非表示のページは数えない）" node tools/check_music_stop.js
run "後半の1枚目のメインプレゼンのページ（前半のテンプレートから見た目ごと写す・お名前・会社名・カテゴリー）" node tools/check_second_main_page.js
run "新メンバー・更新メンバーのページ" node tools/check_member_pages.js
run "ウィークリープレゼンの送り方（業種区分の扉ページは1秒で次へ・前半の「保存済みのタイミングを使用」）" node tools/check_weekly_timing.js
run "ネットワーキングリーダーのページ（見出しをお名前の枠と取り違えない・お2人の部門・まとめのページで該当者のいない部門は見出しも出さない）" node tools/check_networking_leaders.js
run "役職のメンバー紹介（学習コーナーのお名前の大きさも）" node tools/check_role_intro.js
run "カウントダウンと音" node tools/check_countdown.js
run "公式ファイルから雛形の画面" node tools/check_official_dialog.js
run "定例会スライドの画面（架空の名簿：読み込み中・推薦のことば・「更新した」で手直しが戻らない・後半の1枚目のメインプレゼン）" node tools/check_meeting_dialog.js

echo "── 名簿・メンバーブック"
run "メンバーリスト(OCR)の取り込み（PDFに無い方はチェックした方だけ消す）" node tools/check_member_ocr.js
# shellcheck disable=SC2086
run "メンバーブック（組版・編集画面）" node tools/check_memberbook.js $FONTS_DIR
# shellcheck disable=SC2086
run "メンバーブックのPDF・ドライブの差し替え" node tools/check_memberbook_pdf.js $FONTS_DIR

echo "── 画面・ウェブアプリ・文書"
run "設定の画面の保存" node tools/check_save_pages.js
run "ウェブアプリの画面一覧" node tools/check_webapp_pages.js
run "HTML（divの対応・スクリプトの構文・共通部品）" python3 tools/check_html.py
run "名前の衝突・参照先のHTML" python3 tools/check_gas_names.py
run "マニュアルの導線" python3 tools/check_manual_links.py
run "pptxのファイルの情報" bash -c 'python3 tools/clean_pptx_meta.py --check docs/templates/*.pptx'

if [ -n "$ROUTINE_JSON" ] && [ -n "$MEMBERS_JSON" ]; then
  echo "── ルーティンチェックシートの写しを使う検査"
  run "ルーティンチェックシートの読み取り" node tools/check_routine.js "$ROUTINE_JSON"
  run "チャプターの設定" node tools/check_chapter.js "$ROUTINE_JSON" "$MEMBERS_JSON"
  run "初回の準備" node tools/check_setup.js "$ROUTINE_JSON" "$MEMBERS_JSON"
  run "役職ごとの入力" node tools/check_role_input.js "$ROUTINE_JSON" "$MEMBERS_JSON"
  run "役職ごとの入力の画面" node tools/check_role_input_dialog.js "$ROUTINE_JSON" "$MEMBERS_JSON"
  run "スピーカーローテーション" node tools/check_speaker_rotation.js "$ROUTINE_JSON" "$MEMBERS_JSON"
  run "ウィークリープレゼンの始まり" node tools/check_weekly_start.js "$ROUTINE_JSON" "$MEMBERS_JSON"
  run "事前MTGのパワポ" node tools/check_premtg.js "$ROUTINE_JSON" "$MEMBERS_JSON"
  run "トークスクリプト（台本）" node tools/check_talk_script.js "$ROUTINE_JSON" "$MEMBERS_JSON"
  run "定例会スライドの画面（手元の名簿）" node tools/check_meeting_dialog.js "$MEMBERS_JSON"
else
  echo "（ROUTINE_JSON・MEMBERS_JSON が無いため、ルーティンチェックシートの写しを使う検査10本は飛ばしました）"
fi

if [ $ng -eq 0 ]; then echo "すべて OK"; else echo "NG があります。直してからコミットしてください。"; fi
exit $ng

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
run "毎週の流れ（候補→CSV→名簿とPDF→再編集→次の週→メール→割り振り表→PDFの確認）" node tools/check_weekly_flow.js
run "ビジターリスト・割り振り表のシートとPDF（作った日は隠さない・真っ白のPDFを作らない）" node tools/check_visitor_pdf.js
run "CSVから名簿・PDF作成の画面（ブラウザ：過ぎた回の再編集はその回に書き戻す・Shift_JIS のCSV）" node tools/check_csv_dialog.js
run "メールの確認・一括送信の画面（ブラウザ：リンク・キャンセル・記号・二重に送らない）" node tools/check_email_dialog.js
run "ビジター情報の投稿文" node tools/check_visitor_post.js
run "割り振り表のAI（待機リストに居ないビジターホスト・知らない番号・確度とメモ）" node tools/check_allocation_ai.js
run "割り振り表の画面（ブラウザ：読み込んだ回に書き込む・Spreadingの文字を動かさない）" node tools/check_allocation_dialog.js
run "期の替わり目の担当者と名簿の役職（未登録の期で役職を消さない）" node tools/check_role_terms.js

echo "── スライド"
run "定例会スライド（指示どおりか）" node tools/check_meeting_slides.js
run "新メンバー・更新メンバーのページ" node tools/check_member_pages.js
run "役職のメンバー紹介" node tools/check_role_intro.js
run "カウントダウンと音" node tools/check_countdown.js
run "公式ファイルから雛形の画面" node tools/check_official_dialog.js

echo "── 名簿・メンバーブック"
run "メンバーリスト(OCR)の取り込み" node tools/check_member_ocr.js
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
  run "定例会スライドの画面" node tools/check_meeting_dialog.js "$MEMBERS_JSON"
else
  echo "（ROUTINE_JSON・MEMBERS_JSON が無いため、ルーティンチェックシートの写しを使う検査10本は飛ばしました）"
fi

if [ $ng -eq 0 ]; then echo "すべて OK"; else echo "NG があります。直してからコミットしてください。"; fi
exit $ng

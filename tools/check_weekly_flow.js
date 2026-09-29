// 毎週の流れを、はじめから終わりまで通して確かめる（見せかけのスプレッドシート・ドライブ・Gmail で。lib_sheet_fake.js）。
//
//   node tools/check_weekly_flow.js
//
//   開催日の候補 → CSVの読み込み → 名簿とPDFの作成 → 再編集して作り直す（PDFはURLそのまま）→
//   次の週の名簿（前の週はアーカイブ）→ メールの下書き・送信 → 割り振り表（作る・読み戻す）→ 作成済みPDFの確認
//
// 1つ1つの機能の検査（check_visitor_pdf.js など）では見つからない、機能のつなぎ目の誤り
// （シート名・開催日・PDFのURLの受け渡し）を見つけるためのもの。名前・メールアドレスはすべて架空。

process.env.TZ = 'Asia/Tokyo';   // スクリプトのタイムゾーン（appsscript.json）に合わせる
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeEnv } = require('./lib_sheet_fake');
const { loadPage } = require('./lib_minidom');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }
function step(name, fn) {
  try { fn(); } catch (e) { fails.push(name + ' で止まった: ' + (e && e.stack ? e.stack.split('\n').slice(0, 3).join(' / ') : e)); }
}

// 2026/9/29(火) 10:00 に使う（次の定例会は 9/30(水)）
const env = makeEnv({ now: new Date(2026, 8, 29, 10, 0, 0) });
const srv = Object.assign({}, env.globals);
vm.createContext(srv);
for (const f of fs.readdirSync(ROOT).filter((x) => /\.js$/.test(x)).sort()) {   // 本番と同じく全部のファイルを1つの場所に
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), srv, { filename: f });
}

// ---- はじめの状態：メンバー名簿（架空）・休会日 ----
const MEMBER_HEAD = ['No', '業種区分', '氏名', 'ふりがな', 'カテゴリー', '会社名', '役職', 'メモ', '写真ファイル名', '一言コメント',
  '紹介してほしい人', '協業したい人', '入会日', '更新日', '更新期限日', '会社での役職'];
const MEMBERS = [['1', '見本 一郎'], ['2', '試験 花子'], ['3', '架空 三郎'], ['4', '仮名 四郎'], ['5', '例示 五月'], ['6', '模擬 六助'], ['7', '空想 七海'], ['8', '見本 みさき']];
const rosterRows = [MEMBER_HEAD].concat(MEMBERS.map(([no, name]) => MEMBER_HEAD.map((h, i) => (i === 0 ? no : i === 2 ? name : ''))));
env.reset([
  ['メンバー名簿', false, rosterRows],
  ['休会日', true, [['2026/05/06'], ['2026/08/12'], ['2026/12/30']]],
  ['20260923参加者', false, [['No.', '参加者氏名', 'ふりがな', 'カテゴリー', '会社名', '招待者', '備考', '種別', 'メール'], ['V01', '前回 来人', 'ぜんかい くると', '', '', '', '', 'Visitor', 'prev@example.com']]],
  ['20260923参加者_印刷用', false, [['前回の名簿']]],
], { MEMBER_BOOK_URL: 'https://drive.example/MEMBERBOOK' });

const CSV_0930 = [
  'Name,Furigana,Company Name,Business Category,Inviter,Email,Type,Memo For Printing,Payment Status',
  '見本太郎,みほんたろう,見本商事,税理士,試験花子さん,taro@example.com,Visitor,,Paid',
  '例示 次郎,れいじ じろう,例示工務店,工務店,架空 三郎,jiro@example.com,Visitor,"よろしく, お願いします",',
  '仮設 月子,かせつ つきこ,,,見本 一郎,tsuki@example.com,Guest,,',
  '代理 太一,だいり たいち,代理商会,保険,仮名 四郎,taichi@example.com,Substitute,,',
  ',,,,,,,,',
  '無メール 氏,むめーる うじ,,,,,Visitor,,',
].join('\r\n') + '\r\n';
const CSV_1007 = [
  '参加者氏名,ふりがな,会社名,カテゴリー,招待者,メール,種別',
  '翌週 来子,よくしゅう らいこ,翌週社,デザイン,例示 五月,raiko@example.com,Visitor',
].join('\n');

let cand = null, analyzed = null, url0930 = '', id0930 = '', url1007 = '';

// ---- 1) 開催日の候補 ----
step('開催日の候補', () => {
  const list = srv.getMeetingCandidates();
  cand = list[0];
  ck(list.length === 4, '1) 候補が4件でない: ' + list.length);
  ck(cand && cand.dateValue === '2026/09/30' && cand.display === '2026/9/30(水) 第535回', '1) 次の定例会が 2026/9/30(水) 第535回 にならない: ' + JSON.stringify(cand));
  ck(list[1] && list[1].dateValue === '2026/10/07', '1) 2番目の候補が 10/07 でない: ' + JSON.stringify(list[1]));
});

// ---- 2) CSVの読み込み ----
step('CSVの読み込み', () => {
  analyzed = srv.analyzeCsvData(CSV_0930);
  const r = analyzed.rows;
  const byName = (n) => r.find((x) => x['参加者氏名'] === n);
  ck(r.length === 5, '2) 空の行を除いて5人にならない: ' + r.length);
  ck(r.map((x) => x._No).join(',') === 'V01,V02,V03,G01,代理4', '2) 番号が違う: ' + r.map((x) => x._No + ':' + x['参加者氏名']).join(','));
  ck(byName('見本 太郎') && byName('見本 太郎')._needsNameReview === true, '2) 名前にスペースを補って確認に回していない');
  ck(byName('見本 太郎') && byName('見本 太郎')['招待者'] === '試験 花子', '2) 招待者を名簿の氏名に合わせていない: ' + (byName('見本 太郎') || {})['招待者']);
  ck(byName('例示 次郎') && byName('例示 次郎')['メモ（ビジターリストに表示）'] === 'よろしく, お願いします', '2) 引用符の中のカンマで列がずれた');
  ck(analyzed.header.indexOf('メール') >= 0 && analyzed.header.indexOf('種別') >= 0, '2) 英語の見出しを日本語にしていない: ' + analyzed.header.join(','));
  ck(byName('仮設 月子') && byName('仮設 月子')['招待者'] === '見本 一郎' && !byName('仮設 月子')._needsInviterReview, '2) 同じ字の招待者なのに確認に回した・別の方にした');
  // 招待者の照合：姓だけで同じ姓の方が2人 → 黄色で確かめる。「さん」付き・名前に「さ」「ん」がある方も当たる。代理の番号はその方の番号
  const amb = srv.analyzeCsvData(['Name,Furigana,Inviter,Email,Type',
    '姓だけ 招待,せいだけ しょうたい,見本さん,a@example.com,Visitor',
    'さ行 招待,さぎょう しょうたい,見本 みさきさん,b@example.com,Visitor',
    '代理 さき,だいり さき,見本みさき,c@example.com,Substitute'].join('\n')).rows;
  const an = (n) => amb.find((x) => x['参加者氏名'] === n) || {};
  ck(an('姓だけ 招待')._needsInviterReview === true, '2) 姓だけで同じ姓の方が2人いるのに、確かめずに決めた: ' + an('姓だけ 招待')['招待者']);
  ck(an('さ行 招待')['招待者'] === '見本 みさき' && !an('さ行 招待')._needsInviterReview, '2) 名前に「さ」がある方（「さん」付き）が当たらない: ' + an('さ行 招待')['招待者']);
  ck(an('代理 さき')._No === '代理8', '2) 代理の番号が違う: ' + an('代理 さき')._No);
});

// ---- 3) 名簿とPDFを作る ----
step('名簿とPDFを作る', () => {
  const html = srv.createFinalSheet(cand.dateValue, cand.display, analyzed.rows, analyzed.header);
  ck(/処理が完了しました/.test(html), '3) 完了の知らせが出ない');
  const data = env.values('20260930参加者'), print = env.values('20260930参加者_印刷用');
  ck(data && data[0].slice(0, 7).join(',') === 'No.,参加者氏名,ふりがな,カテゴリー,会社名,招待者,備考', '3) 参加者シートの見出しが違う: ' + (data && data[0].join(',')));
  ck(data && data.length === 6, '3) 参加者シートの行数が違う: ' + (data && data.length));
  ck(print && /定例会へようこそ/.test(print[0][0]) && print[2][0] === cand.display, '3) 印刷用シートの題・開催日が違う');
  ck(env.visibleNames().includes('20260930参加者') && env.visibleNames().includes('20260930参加者_印刷用'), '3) 作ったシートが隠れた: ' + env.visibleNames().join(','));
  ck(env.hiddenNames().includes('20260923参加者'), '3) 前の週のシートがアーカイブされない');
  url0930 = env.props.LATEST_VISITOR_LIST_URL || '';
  id0930 = env.props['VISITOR_PDF_ID_20260930参加者'] || '';
  ck(id0930 && url0930 === 'https://drive.example/' + id0930, '3) PDFのURL・IDの控えが違う: ' + url0930 + ' / ' + id0930);
  const text = env.fileText(id0930);
  ck(['見本 太郎', '例示 次郎', '仮設 月子', '代理 太一', cand.display].every((s) => text.includes(s)), '3) PDFに名簿の中身が入っていない（真っ白）: ' + JSON.stringify(text.slice(0, 80)));
  ck(!/taro@example\.com/.test(text), '3) PDF（配る名簿）にメールアドレスが出ている');
  ck(env.drive.files[id0930] && env.drive.files[id0930].sharing === 'ANYONE_WITH_LINK/VIEW', '3) PDFを「リンクを知っている全員が閲覧」にしていない');
  ck(env.props.LATEST_MEETING_DATE === '2026/09/30', '3) 最新の開催日の控えが違う: ' + env.props.LATEST_MEETING_DATE);
});

// ---- 4) 再編集して作り直す：PDFは同じファイル（URLそのまま） ----
step('再編集して作り直す', () => {
  const sheets = srv.getExistingVisitorSheets();
  ck(sheets[0] === '20260930参加者', '4) 再編集の一覧の先頭が今回の開催日でない: ' + sheets.join(','));
  const loaded = srv.loadSheetData('20260930参加者');
  ck(loaded.rows.map((x) => x._No + ':' + x['参加者氏名']).join(',') === 'V01:見本 太郎,V02:無メール 氏,V03:例示 次郎,G01:仮設 月子,代理4:代理 太一',
     '4) 読み戻した番号・氏名が違う: ' + loaded.rows.map((x) => x._No + ':' + x['参加者氏名']).join(','));
  ck(loaded.rows[2]['メモ（ビジターリストに表示）'] === 'よろしく, お願いします', '4) 備考がメモの欄に戻らない');
  ck(loaded.meeting && loaded.meeting.dateValue === '2026/09/30' && loaded.meeting.display === cand.display,
     '4) 読み込んだシートの開催日（書き戻す先）が違う: ' + JSON.stringify(loaded.meeting));
  ck(loaded.rows[0]['メール'] === 'taro@example.com' && loaded.rows[0]['種別'] === 'Visitor', '4) メール・種別が読み戻せない');
  loaded.rows[2]['会社名'] = '例示工務店（新）';
  loaded.rows[4]['招待者'] = '例示 五月';   // 代理の方の招待者を直す（画面の番号は「代理4」のまま）
  const before = env.drive.created.length;
  const html4 = srv.createFinalSheet(cand.dateValue, cand.display, loaded.rows, loaded.header);
  ck(!/新しいPDFを作りました/.test(html4), '4) 差し替えられたのに「新しいPDFを作りました」と出た');
  ck(env.values('20260930参加者').some((r) => r[0] === '代理5' && r[1] === '代理 太一'), '4) 招待者を直した代理の番号が付け直されない: ' + env.values('20260930参加者').map((r) => r[0]).join(','));
  loaded.rows[4]['招待者'] = '仮名 四郎';
  srv.createFinalSheet(cand.dateValue, cand.display, loaded.rows, loaded.header);
  ck(env.drive.created.length === before, '4) 作り直したのに新しいPDFができた（URLが変わる）');
  ck(env.props.LATEST_VISITOR_LIST_URL === url0930, '4) 作り直したらPDFのURLが変わった');
  ck(env.fileText(id0930).includes('例示工務店（新）'), '4) 作り直したPDFに直した内容が入っていない');
  ck(env.values('20260930参加者').length === 6, '4) 作り直したら参加者シートの行数が変わった: ' + env.values('20260930参加者').length);
});

// ---- 4a) 招待者が空の代理を読み戻しても、名簿の最初の方の代理にならない ----
step('招待者が空の代理', () => {
  env.ss.insertSheet('20261014参加者').getRange(1, 1, 2, 9).setValues([
    ['No.', '参加者氏名', 'ふりがな', 'カテゴリー', '会社名', '招待者', '備考', '種別', 'メール'],
    ['代理??', '空欄 代理', 'くうらん だいり', '', '', '', '', 'Substitute', '']]);
  const r = srv.loadSheetData('20261014参加者').rows[0];
  ck(r._No === '代理??', '4a) 招待者が空の代理が、名簿の最初の方の代理になった: ' + r._No);
  env.ss.deleteSheet(env.sheet('20261014参加者'));
});

// ---- 4b) 前のPDFをゴミ箱に入れてしまっていても、作り直せば戻して差し替える（ゴミ箱のままだとリンクが開けない） ----
step('ゴミ箱のPDFを作り直す', () => {
  env.drive.files[id0930].trashed = true;
  ck(srv.generateEmailDrafts('20260930参加者').visitorList === '', '4b) ゴミ箱のPDFのリンクを、メールに入れようとした');
  const loaded = srv.loadSheetData('20260930参加者');
  srv.createFinalSheet(loaded.meeting.dateValue, loaded.meeting.display, loaded.rows, loaded.header);
  ck(env.drive.files[id0930].trashed === false && env.props['VISITOR_PDF_ID_20260930参加者'] === id0930, '4b) ゴミ箱のPDFを戻して差し替えていない（URLが変わる・開けないまま）');
  ck(srv.generateEmailDrafts('20260930参加者').visitorList === url0930, '4b) 作り直したあと、メールにリンクが入らない');
});

// ---- 5) 次の週の名簿：前の週（今週）はアーカイブされる ----
step('次の週の名簿', () => {
  const a2 = srv.analyzeCsvData(CSV_1007);
  srv.createFinalSheet('2026/10/07', '2026/10/7(水) 第536回', a2.rows, a2.header);
  url1007 = env.props.LATEST_VISITOR_LIST_URL;
  ck(url1007 && url1007 !== url0930, '5) 次の週のPDFが別のファイルにならない');
  ck(env.visibleNames().includes('20261007参加者'), '5) 次の週のシートが隠れた');
  ck(env.visibleNames().includes('20260930参加者') && env.visibleNames().includes('20260930参加者_印刷用'), '5) 翌週分を先に作ったら、次回（9/30）のシートまで隠れた');
  ck(env.hiddenNames().includes('20260923参加者'), '5) 過ぎた回（9/23）のシートがアーカイブされない');
  ck(env.fileText(id0930).includes('見本 太郎'), '5) 次の週を作ったら、今週のPDFが書き換わった');
  const vctx = srv.getVisitorSlideContext();
  ck(vctx.ok && vctx.defaultSheet === '20260930参加者', '5) ビジター・代理スライドの既定が次回（9/30）でない（翌週を先に作ったとき）: ' + vctx.defaultSheet);
  const vp = loadPage('slides_visitor.html', { fails, server: {
    getVisitorSlideContext: () => JSON.parse(JSON.stringify(srv.getVisitorSlideContext())),
    previewVisitorSlideData: (n) => JSON.parse(JSON.stringify(srv.previewVisitorSlideData(n))) } });
  vp.step('ビジター・代理スライドの画面を開く', () => vp.window.onload());
  const pvCall = vp.log.calls.find((c) => c.name === 'previewVisitorSlideData');
  ck(vp.els.sheet.value === '20260930参加者' && pvCall && pvCall.args[0] === '20260930参加者',
     '5) ビジター・代理スライドの画面が、次回（9/30）ではなく ' + vp.els.sheet.value + ' を開いた（読み込み: ' + (pvCall && pvCall.args[0]) + '）');
});

// ---- 5.5) 抽選ルーレットのビジター招待数：翌週（10/07）の名簿を先に作ってあっても、次回（9/30）の分を数える ----
step('抽選ルーレットのビジター招待数', () => {
  env.ss.insertSheet('抽選ルーレット作成用①').getRange(1, 1, 5, 8).setValues([
    ['姓', '名', '', '', '', '', '', 'ビ'], ['試験', '花子', '', '', '', '', '', ''], ['架空', '三郎', '', '', '', '', '', ''],
    ['例示', '五月', '', '', '', '', '', ''], ['見本', '一郎', '', '', '', '', '', '']]);
  const r = srv.fillRouletteVisitorCounts();
  const h = env.values('抽選ルーレット作成用①').slice(1).map((row) => row[0] + row[1] + '=' + row[7]).join(',');
  ck(r.ok && /20260930参加者/.test(r.message) && h === '試験花子=1,架空三郎=1,例示五月=0,見本一郎=0',
     '5.5) 抽選ルーレットが次回（9/30）の名簿で数えない（翌週の名簿を数えた？）: ' + h + ' / ' + r.message);
  // Spreadingでキャンセルのビジターは数えない
  env.ss.insertSheet('20261028参加者').getRange(1, 1, 4, 10).setValues([
    ['No.', '参加者氏名', 'ふりがな', 'カテゴリー', '会社名', '招待者', '備考', '種別', 'メール', 'ステータス'],
    ['V01', '参加 する', '', '', '', '試験 花子', '', 'Visitor', '', '参加予定'],
    ['V02', '取消 した', '', '', '', '試験 花子', '', 'Visitor', '', 'キャンセル'],
    ['V03', '参加 二人目', '', '', '', '架空 三郎', '', 'Visitor', '', '']]);
  const t = srv.countVisitorsByInviter_(env.sheet('20261028参加者'));
  ck(t.visitors === 2 && t.counts['試験花子'] === 1 && t.counts['架空三郎'] === 1 && JSON.stringify(t.cancelled) === '["取消 した"]',
     '5.5) キャンセルのビジターを抽選ルーレットの数に入れた: ' + JSON.stringify(t));
  env.ss.deleteSheet(env.sheet('20261028参加者'));
  env.ss.deleteSheet(env.sheet('抽選ルーレット作成用①'));
});

// ---- 6) メール：今日（9/29）は次回の 9/30 が既定。下書きの日付・リンクは選んだ開催日のもの ----
let drafts = null;
step('メールの下書き', () => {
  const ctx = srv.getEmailContext();
  ck(ctx.ok && ctx.defaultSheet === '20260930参加者', '6) メールの既定の開催日が次回（9/30）でない: ' + ctx.defaultSheet);
  const res = srv.generateEmailDrafts('20260930参加者');
  drafts = res.drafts;
  ck(res.date === '2026年9月30日', '6) 下書きの日付が違う: ' + res.date);
  ck(drafts.map((d) => d.email).join(',') === 'taro@example.com,jiro@example.com,tsuki@example.com,taichi@example.com',
     '6) 宛先が違う（メールの無い方は外す）: ' + drafts.map((d) => d.email).join(','));
  ck(drafts.every((d) => d.subject.includes('2026年9月30日')), '6) 件名の日付が違う');
  const v = drafts.find((d) => d.type === 'Visitor');
  ck(v && v.body.includes('見本 太郎 様'), '6) 本文の宛名が違う');
  ck(v && v.body.includes(url0930), '6) 9/30 の案内に 9/30 のビジターリストのURLが入っていない（最後に作った別の開催日のURLになる）: ' + (v && (v.body.match(/https:\/\/drive\.example\/\w+/g) || []).join(',')));
  ck(v && !v.body.includes(url1007), '6) 9/30 の案内に次の週（10/07）のビジターリストのURLが入っている');
  ck(v && v.body.includes('https://drive.example/MEMBERBOOK'), '6) メンバーブックのURLが入っていない');
  ck(JSON.stringify(res.noEmail) === '["無メール 氏"]', '6) メールアドレスの無い方を画面に返していない: ' + JSON.stringify(res.noEmail));
  // PDFを作った方以外（そのファイルの権限が無い方）がメールの画面を開いても、その開催日のリンクが入る
  env.denied.add(env.idOfUrl(url0930));
  const other = srv.generateEmailDrafts('20260930参加者');
  env.denied.clear();
  const ov = other.drafts.find((d) => d.type === 'Visitor');
  ck(other.visitorList === url0930 && ov && ov.body.includes(url0930), '6) PDFの権限が無い方が開くと、リンクが入らない: ' + other.visitorList);
  env.errors.length = 0;   // 開けなかったことの記録（console.warn ではなく error なら残る）は、ここでは想定どおり
  const res2 = srv.generateEmailDrafts('20261007参加者');
  ck(res2.date === '2026年10月7日' && res2.drafts.length === 1 && res2.drafts[0].body.includes(url1007), '6) 10/07 の下書きの日付・URLが違う');
});

// PDFの無い開催日（9/23 は名簿だけある）は、別の回のリンクで代わりにせず「【未生成】」にして知らせる
step('PDFの無い開催日のメール', () => {
  const r = srv.generateEmailDrafts('20260923参加者');
  ck(r.visitorList === '' && r.drafts.length === 1 && r.drafts[0].body.includes('【未生成】') && !/F\d/.test(r.drafts[0].body.replace('MEMBERBOOK', '')),
     '6) PDFの無い開催日のメールに、別の回のリンクが入った: ' + (r.drafts[0] && r.drafts[0].body));
  ck(r.memberBook === 'https://drive.example/MEMBERBOOK', '6) メンバーブックのURLを画面に返していない');
});

// メールの画面：開く → 下書き → 送る（宛先・件名・本文がそのまま届く）。PDFの無い開催日では警告と送る前の確認
step('メールの画面', () => {
  const clone = (x) => JSON.parse(JSON.stringify(x));
  const server = {
    getEmailContext: () => clone(srv.getEmailContext()),
    generateEmailDrafts: (n) => clone(srv.generateEmailDrafts(n)),
    sendSingleEmail: (e, cc, bcc) => clone(srv.sendSingleEmail(e, cc, bcc)),
  };
  const page = loadPage('email.html', { fails, server });
  page.step('メールの画面を開く', () => page.window.onload());
  ck(page.els.meetSel.value === '20260930参加者', '6) メールの画面の既定の開催日が 9/30 でない: ' + page.els.meetSel.value);
  ck(!/⚠/.test(page.els.meetNote.innerText), '6) リンクがそろっているのに警告が出た: ' + page.els.meetNote.innerText);
  ck(page.els.subj_0 && page.els.subj_0.value === drafts[0].subject && page.els.body_0.value === drafts[0].body, '6) 画面の件名・本文が下書きと違う');
  page.els.body_1.value = page.els.body_1.value + '\n（手で足した一文）';
  page.els.chk_2.checked = false;
  const before = env.mail.length;
  page.step('送る', () => page.run('startSending()'));
  const sent = env.mail.slice(before);
  ck(sent.map((m) => m.to).join(',') === 'taro@example.com,jiro@example.com,taichi@example.com', '6) 送った宛先が違う（チェックを外した方には送らない）: ' + sent.map((m) => m.to).join(','));
  ck(sent[1] && /（手で足した一文）$/.test(sent[1].body), '6) 画面で直した本文で送っていない');
  const withList = sent.filter((m) => /ビジターリスト/.test(m.body));   // 既定のひな形では、ビジターあての本文にだけリンクがある
  ck(withList.length === 2 && withList.every((m) => m.body.includes(url0930) && !m.body.includes(url1007)), '6) 送ったメールのビジターリストのURLが違う');
  ck(page.log.confirms.length === 1 && !/⚠/.test(page.log.confirms[0]), '6) 送る前の確認の出方が違う');

  // 開き直すと、送った方はチェックが外れ「送信済み」と出る。送っていない方（チェックを外した方）はチェックが入る
  const again = loadPage('email.html', { fails, server });
  again.step('開き直す', () => again.window.onload());
  ck(again.els.chk_0 && again.els.chk_0.checked === false && again.els.chk_1.checked === false && again.els.chk_3.checked === false && again.els.chk_2.checked === true,
     '6) 開き直したら、送った方にもチェックが入っている（二重に送る）: ' + [0, 1, 2, 3].map((i) => again.els['chk_' + i] && again.els['chk_' + i].checked).join(','));
  ck(/すでに送った方・Spreadingでキャンセルの方（3名）は、チェックを外してあります/.test(again.els.meetNote.innerText), '6) 送信済みであることが画面に出ない: ' + again.els.meetNote.innerText);
  again.els.chk_0.checked = true;   // わざと送った方にもう一度チェック → 送る前の確認で知らせる
  const b0 = env.mail.length;
  again.step('もう一度送る', () => again.run('startSending()'));
  ck(again.log.confirms.length === 1 && /すでに送っています/.test(again.log.confirms[0]), '6) 送った方にもう一度送るとき、確認で知らせない');
  ck(env.mail.length === b0 + 2, '6) 開き直して送った数が違う: ' + (env.mail.length - b0));

  // 件名に「"」や「<」があっても、切れずにそのまま送る
  const tricky = loadPage('email.html', { fails, server: Object.assign({}, server, {
    generateEmailDrafts: (n) => { const r = clone(srv.generateEmailDrafts(n)); r.drafts = r.drafts.slice(0, 1); r.drafts[0].send = true; r.drafts[0].sentAt = ''; r.drafts[0].subject = '【"見本" <定例会>】ご案内 & 資料'; return r; },
  }) });
  tricky.step('開く', () => tricky.window.onload());
  const b2 = env.mail.length;
  tricky.step('送る', () => tricky.run('startSending()'));
  ck(env.mail[b2] && env.mail[b2].subject === '【"見本" <定例会>】ご案内 & 資料', '6) 件名の記号で件名が切れた: ' + (env.mail[b2] && env.mail[b2].subject));

  // PDFの無い開催日：警告し、送る前の確認にも出す
  const warn = loadPage('email.html', { fails, server });
  warn.step('開く', () => warn.window.onload());
  warn.els.meetSel.value = '20260923参加者';
  warn.step('9/23 を選ぶ', () => warn.run('loadDrafts()'));
  ck(/ビジターリストのPDFがありません/.test(warn.els.meetNote.innerText), '6) PDFの無い開催日で警告が出ない: ' + warn.els.meetNote.innerText);
  ck(/過ぎた回（2026年9月23日）/.test(warn.els.meetNote.innerText), '6) 過ぎた回を選んだことを知らせない: ' + warn.els.meetNote.innerText);
  const b3 = env.mail.length;
  warn.step('送る', () => warn.run('startSending()'));
  ck(warn.log.confirms.length === 1 && /ビジターリストのPDFがありません/.test(warn.log.confirms[0]), '6) 送る前の確認に、リンクが無いことが出ない');
  ck(env.mail.length === b3 + 1, '6) 確認でOKしたのに送られない');
});

// Spreadingでキャンセルになった方は、はじめからチェックを外し、印を出す
step('キャンセルの方', () => {
  env.ss.insertSheet('20261021参加者').getRange(1, 1, 3, 10).setValues([
    ['No.', '参加者氏名', 'ふりがな', 'カテゴリー', '会社名', '招待者', '備考', '種別', 'メール', 'ステータス'],
    ['V01', '参加 する', 'さんか する', '', '', '', '', 'Visitor', 'yes@example.com', '参加予定'],
    ['V02', '取消 した', 'とりけし した', '', '', '', '', 'Visitor', 'no@example.com', 'キャンセル']]);
  const r = srv.generateEmailDrafts('20261021参加者');
  const yes = r.drafts.find((d) => d.email === 'yes@example.com'), no = r.drafts.find((d) => d.email === 'no@example.com');
  ck(yes && yes.send === true && no && no.send === false && no.cancelled === true, '6) キャンセルの方にもチェックが入る: ' + JSON.stringify(r.drafts.map((d) => [d.email, d.send, d.cancelled])));
  const pv = srv.previewVisitorSlideData('20261021参加者');
  ck(pv.ok && pv.visitors.map((v) => v.name).join(',') === '参加 する' && JSON.stringify(pv.cancelled) === '["取消 した"]' && /キャンセル/.test(pv.message),
     '6) キャンセルの方をビジター紹介のスライドに入れる: ' + JSON.stringify(pv && { v: pv.visitors, c: pv.cancelled }));
  const sp = loadPage('slides_visitor.html', { fails, server: {
    getVisitorSlideContext: () => ({ ok: true, sheets: ['20261021参加者'], defaultSheet: '20261021参加者', templates: { intro: true, presen: true, dairi: true } }),
    previewVisitorSlideData: (n) => JSON.parse(JSON.stringify(srv.previewVisitorSlideData(n))) } });
  sp.step('ビジター・代理スライドの画面', () => sp.window.onload());
  ck(/取消 した/.test(sp.els.msg.innerText) && !/取消 した/.test(sp.els.prev.innerHTML) && /参加 する/.test(sp.els.prev.innerHTML),
     '6) ビジター・代理スライドの画面で、キャンセルの方を知らせない・一覧に入れた: ' + sp.els.msg.innerText);
  const page = loadPage('email.html', { fails, server: { getEmailContext: () => ({ ok: true, sheets: [{ sheet: '20261021参加者', label: '2026年10月21日' }], defaultSheet: '20261021参加者' }),
    generateEmailDrafts: (n) => JSON.parse(JSON.stringify(srv.generateEmailDrafts(n))) } });
  page.step('開く', () => page.window.onload());
  ck(page.els.chk_0 && page.els.chk_0.checked === true && page.els.chk_1 && page.els.chk_1.checked === false && /キャンセルの方（1名）は、チェックを外してあります/.test(page.els.meetNote.innerText),
     '6) 画面でキャンセルの方のチェックが外れていない・知らせが出ない: ' + page.els.meetNote.innerText);
  env.ss.deleteSheet(env.sheet('20261021参加者'));
});

step('メールの送信', () => {
  const b = env.mail.length;
  const r = srv.sendSingleEmail(drafts[0], ' cc@example.com ', '');
  ck(r.success === true, '6) 送れない: ' + JSON.stringify(r));
  const m = env.mail[b] || {};
  ck(m.to === 'taro@example.com' && m.subject === drafts[0].subject && m.options.cc === 'cc@example.com' && !('bcc' in m.options),
     '6) 送ったメールの宛先・CC・BCCが違う: ' + JSON.stringify(m));
  const zen = srv.sendSingleEmail({ email: 'ｚｅｎ＠ｅｘａｍｐｌｅ．ｃｏｍ', subject: 's', body: 'b' }, '', '');
  ck(zen.success === true && env.mail[env.mail.length - 1].to === 'zen@example.com', '6) 全角で入力されたメールアドレスに送れない: ' + JSON.stringify(zen));
  env.mail.pop();
  const bad = srv.sendSingleEmail({ email: 'no-at-mark', subject: 's', body: 'b' }, '', '');
  ck(bad.success === false && env.mail.length === b + 1, '6) おかしなアドレスに送ろうとした');
});

// ---- 7) 割り振り表：作る → PDF → 読み戻す ----
step('割り振り表', () => {
  const d = srv.getAllocationData('2026/09/30');
  ck(d.visitors.map((v) => v.no + ':' + v.name + ':' + v.inviter).join(',') === 'V01:見本 太郎:試験 花子,V02:無メール 氏:,V03:例示 次郎:架空 三郎,G01:仮設 月子:見本 一郎,代理4:代理 太一:仮名 四郎',
     '7) 割り振りのビジター一覧が違う: ' + d.visitors.map((v) => v.no + ':' + v.name + ':' + v.inviter).join(','));
  ck(d.pool.map((m) => m.no).join(',') === '5,6,7,8', '7) 招待者を除いた待機メンバーが違う: ' + d.pool.map((m) => m.no).join(','));
  const facil = { V01: '5', V03: '6' }, orien = { V01: ['7'] }, room = { V01: ['6'], V03: ['7'] }, conn = { V01: 'IT関係の方' };
  const html = srv.saveAllocationSheet('2026/09/30', cand.display, d.visitors, d.pool, facil, orien, room, conn, {});
  ck(/作成完了/.test(html), '7) 割り振り表の完了の知らせが出ない');
  ck(env.visibleNames().includes('20260930割り振り表') && env.visibleNames().includes('20260930オリエン') && env.visibleNames().includes('20260930オープンネット'),
     '7) 割り振り表・ルームの表が見えない: ' + env.visibleNames().join(','));
  ck(env.visibleNames().includes('20260930参加者'), '7) 割り振り表を作ったのに、その日の参加者シートが隠れたまま');
  const aid = env.props['ALLOC_PDF_ID_20260930割り振り'];
  ck(aid && env.props.LATEST_ALLOCATION_URL === 'https://drive.example/' + aid, '7) 割り振り表のPDFの控えが違う');
  const t = env.fileText(aid);
  ck(t.includes('見本 太郎') && t.includes('例示 五月') && t.includes('IT関係の方'), '7) 割り振り表のPDFに中身が入っていない: ' + JSON.stringify(t.slice(0, 80)));
  const back = srv.getAllocationData('2026/09/30');
  ck(JSON.stringify(back.facilAlloc) === JSON.stringify(facil), '7) ファシリテーターが読み戻せない: ' + JSON.stringify(back.facilAlloc));
  ck(JSON.stringify(back.roomAlloc.V01) === '["6"]' && JSON.stringify(back.roomAlloc.V03) === '["7"]', '7) ルームメンバーが読み戻せない: ' + JSON.stringify(back.roomAlloc));
  ck(JSON.stringify(back.orienAlloc.V01) === '["7"]', '7) オリエンが読み戻せない: ' + JSON.stringify(back.orienAlloc));
  ck(back.connectReq.V01 === 'IT関係の方', '7) つなげたいメンバーが読み戻せない: ' + JSON.stringify(back.connectReq));
  // 名前が1字だけ違うメンバー（見本 一郎・見本 二郎）がいても、招待者を取り違えない。待機リストからは本当の招待者を除く
  env.sheet('メンバー名簿').appendRow(MEMBER_HEAD.map((h, i) => (i === 0 ? '9' : i === 2 ? '見本 二郎' : '')));
  env.ss.insertSheet('20261104参加者').getRange(1, 1, 2, 9).setValues([
    ['No.', '参加者氏名', 'ふりがな', 'カテゴリー', '会社名', '招待者', '備考', '種別', 'メール'],
    ['V01', '似名 来人', 'にな くると', '', '', '見本 二郎', '', 'Visitor', '']]);
  const sim = srv.getAllocationData('2026/11/04');
  ck(sim.visitors[0].inviter === '見本 二郎', '7) 名前が1字違いのメンバーに、招待者を取り違えた: ' + sim.visitors[0].inviter);
  ck(!sim.pool.some((m) => m.no === '9') && sim.pool.some((m) => m.no === '1'), '7) 待機リストから、本当の招待者でない方を除いた: ' + sim.pool.map((m) => m.no).join(','));
  env.ss.deleteSheet(env.sheet('20261104参加者'));
  const rosterSheet = env.sheet('メンバー名簿');
  rosterSheet._values = rosterSheet._values.filter((r) => r[0] !== '9');
  const orienSheet = env.values('20260930オリエン'), openNet = env.values('20260930オープンネット');
  ck(orienSheet && orienSheet.some((r) => String(r[1]).includes('見本 太郎') && String(r[2]).includes('空想 七海')), '7) オリエンの表が違う');
  ck(openNet && openNet.some((r) => String(r[1]).includes('例示 次郎') && r[3] === '模擬 六助'), '7) オープンネットの表が違う');
});

// ---- 7b) ビジターホスト・優先順位は、名簿の番号がずれても同じ方のまま（OCR・Spreadingの取り込みで番号がずれる） ----
step('ビジターホストと番号のずれ', () => {
  srv.saveVisitorHosts(['3', '4']);                    // 架空 三郎・仮名 四郎
  srv.saveMemberPriorities({ 3: 1, 4: 2 });
  const sh = env.sheet('メンバー名簿'), vals = sh._values;
  // 新メンバーが No.3 に入り、3番から後ろが1つずつずれた
  const shifted = [vals[0], vals[1], vals[2], MEMBER_HEAD.map((h, i) => (i === 0 ? '3' : i === 2 ? '新入 太郎' : ''))]
    .concat(vals.slice(3).map((r) => r.map((v, i) => (i === 0 ? String(Number(v) + 1) : v))));
  const saved = vals.map((r) => r.slice());
  sh._values = shifted;
  ck(JSON.stringify(srv.getVisitorHosts()) === JSON.stringify(['4', '5']), '7b) 番号がずれたら、ビジターホストが別の方になった: ' + JSON.stringify(srv.getVisitorHosts()));
  ck(JSON.stringify(srv.getMemberPriorities()) === JSON.stringify({ 4: 1, 5: 2 }), '7b) 番号がずれたら、優先順位が別の方に移った: ' + JSON.stringify(srv.getMemberPriorities()));
  sh._values = saved;
  // 前の版で番号だけ保存したもの：最初に読んだときに、いまの名簿の氏名を付けて控える
  delete env.props.VISITOR_HOSTS_N; env.props.VISITOR_HOSTS = JSON.stringify(['5']);
  ck(JSON.stringify(srv.getVisitorHosts()) === '["5"]' && /例示 ?五月/.test(env.props.VISITOR_HOSTS_N || ''), '7b) 前の版の保存に、氏名を付けて控えない');
  srv.saveVisitorHosts([]); srv.saveMemberPriorities({});
});

// ---- 8) 作成済みPDFの確認・PDFのみ再作成 ----
step('作成済みPDFの確認', () => {
  // 前の回（9/23）のPDFを作り直しても、「作成済みPDFの確認」は次回（9/30）のものを出す
  srv.regeneratePdfOnly('20260923参加者');
  const links = srv.getPdfLinks();
  ck(links.date === '2026年9月30日' && links.visitorList === url0930, '8) 作成済みPDFのビジターリストが次回（9/30）のものでない: ' + JSON.stringify(links));
  ck(links.allocation === env.props.LATEST_ALLOCATION_URL && links.memberBook === 'https://drive.example/MEMBERBOOK', '8) 作成済みPDFのリンクが違う: ' + JSON.stringify(links));
  // 印刷用シートを手で直して「PDFのみ再作成」→ 同じファイルに直した内容
  const ps = env.sheet('20260930参加者_印刷用');
  ps.getRange(6, 5).setValue('手で直した会社');
  srv.regeneratePdfOnly('20260930参加者');
  ck(env.fileText(id0930).includes('手で直した会社'), '8) PDFのみ再作成で、手で直した内容が入らない');
  ck(srv.getPdfLinks().visitorList === url0930, '8) PDFのみ再作成のあと、作成済みPDFのリンクが変わった');
});

// ---- 8b) PDFを作る方が、スプレッドシートのフォルダに書き込めない・リンクで共有できないとき ----
step('PDFの作る場所と共有', () => {
  const rows = srv.analyzeCsvData(CSV_1007);
  env.readOnly.add('FOLDER');                            // スプレッドシートのフォルダは閲覧だけの共有
  const created = env.drive.created.length;
  srv.createFinalSheet('2026/10/14', '2026/10/14(水) 第537回', rows.rows, rows.header);
  const id = env.props['VISITOR_PDF_ID_20261014参加者'], f = env.drive.files[id];
  ck(env.drive.created.length === created + 1 && f && f.parent !== 'FOLDER' && f.sharing === 'ANYONE_WITH_LINK/VIEW' && env.fileText(id).includes('翌週 来子'),
     '8b) 書き込めるほかの場所にPDFを作れない: ' + JSON.stringify(f && { parent: f.parent, sharing: f.sharing }));
  env.readOnly.clear();
  env.sharingBlocked = true;                              // 組織の設定でリンクの共有が禁止されている
  let err = null;
  try { srv.createFinalSheet('2026/10/21', '2026/10/21(水) 第538回', rows.rows, rows.header); } catch (e) { err = e; }
  const last = env.drive.created[env.drive.created.length - 1];
  ck(err && /リンクを知っている全員が閲覧可/.test(err.message), '8b) リンクで共有できないことを知らせない: ' + (err && err.message));
  ck(!env.props['VISITOR_PDF_ID_20261021参加者'] && last && env.drive.files[last.id].trashed === true, '8b) 共有できないPDFを登録した・残した');
  env.sharingBlocked = false;
  env.errors.length = 0;
});

// ---- 9) スプレッドシートを開けない方（ウェブアプリのURLだけ知っている方）は、設定を読めない・書き換えられない ----
step('設定の読み書きは編集者だけ', () => {
  const realGet = srv.SpreadsheetApp.getActiveSpreadsheet, realOpen = srv.SpreadsheetApp.openById;
  env.props.BNI_SPREADSHEET_ID = 'SSID'; env.props.GEMINI_API_KEY = 'secret-key'; env.props.MAIL_TPL_BCC = 'host@example.com';
  srv.SpreadsheetApp.getActiveSpreadsheet = () => null;                       // ウェブアプリ（開いているスプレッドシートが無い）
  srv.SpreadsheetApp.openById = () => { throw new Error('You do not have permission to access the requested document.'); };
  const tries = [['getApiSettings', []], ['saveApiSettings', [{ apiKey: 'x', modelName: 'y' }]], ['getTemplates', []],
                 ['saveTemplates', [{ bcc: 'evil@example.com' }]], ['getMailWebAppSettings', []], ['saveMailWebAppSettings', [{ webAppUrl: 'https://evil.example/' }]],
                 ['saveAllocationNote', ['書き換え']], ['getChapterSettings', []], ['saveChapterSettings', [{ name: '乗っ取り' }]]];
  tries.forEach(([fn, args]) => {
    let out = null, err = null;
    try { out = srv[fn](...args); } catch (e) { err = e; }
    const leaked = JSON.stringify(out || '').includes('secret-key') || JSON.stringify(out || '').includes('host@example.com');
    ck((err || (out && out.ok === false)) && !leaked, '9) スプレッドシートを開けない方が ' + fn + ' を使えた: ' + JSON.stringify(out).slice(0, 120));
  });
  ck(env.props.GEMINI_API_KEY === 'secret-key' && env.props.MAIL_TPL_BCC === 'host@example.com' && !env.props.MAIL_WEB_APP_URL
     && env.props.ALLOCATION_NOTE === undefined, '9) スプレッドシートを開けない方が、設定を書き換えた');
  srv.SpreadsheetApp.getActiveSpreadsheet = realGet; srv.SpreadsheetApp.openById = realOpen;
  ck(srv.getApiSettings().apiKey === 'secret-key', '9) 編集者（開ける方）が設定を読めない');
  env.errors.length = 0;
});

// ---- 10) トークスクリプトの参加者：移行前の4桁の名前（0916参加者）しか無い回も読む ----
step('トークスクリプトの参加者（移行前の4桁の名前）', () => {
  env.ss.insertSheet('0916参加者').getRange(1, 1, 2, 9).setValues([
    ['No.', '参加者氏名', 'ふりがな', 'カテゴリー', '会社名', '招待者', '備考', '種別', 'メール'],
    ['V01', '旧名 来人', 'きゅうめい くると', '', '', '見本 一郎', '', 'Visitor', '']]);
  const people = srv.talkEnv_(new Date(2026, 8, 16)).people();
  ck(people && people.map((p) => p.name).join(',') === '旧名 来人', '10) 4桁の名前の参加者シートを、トークスクリプトが読まない: ' + JSON.stringify(people));
  env.ss.deleteSheet(env.sheet('0916参加者'));
});

// ---- 11) 移行前の4桁の名前のシート：半年前の回を来年の回にしない。同じ開催日に両方あれば、年の付いた方を既定に ----
step('移行前の4桁の名前のシート', () => {
  const HEADROW = ['No.', '参加者氏名', 'ふりがな', 'カテゴリー', '会社名', '招待者', '備考', '種別', 'メール'];
  env.ss.insertSheet('0325参加者').getRange(1, 1, 2, 9).setValues([HEADROW, ['V01', '三月 来人', '', '', '', '', '', 'Visitor', 'march@example.com']]);
  env.ss.insertSheet('0930参加者').getRange(1, 1, 2, 9).setValues([HEADROW, ['V01', '旧九月 来人', '', '', '', '', '', 'Visitor', 'old0930@example.com']]);
  const d = srv.meetingDateFromKey_('0325');
  ck(d && d.getFullYear() === 2026 && d.getMonth() === 2 && d.getDate() === 25, '11) 4桁の「0325」を来年の回にした: ' + d);
  const ctx = srv.getEmailContext();
  const m = ctx.sheets.find((x) => x.sheet === '0325参加者');
  ck(m && m.past === true && /2026年3月25日/.test(m.label), '11) 半年前の4桁のシートが、これからの回として出る: ' + JSON.stringify(m));
  ck(ctx.defaultSheet === '20260930参加者', '11) 同じ開催日に4桁のシートもあると、そちらが既定になる: ' + ctx.defaultSheet);
  const ld = srv.loadSheetData('0325参加者');
  ck(ld.meeting && ld.meeting.dateValue === '2026/03/25', '11) 4桁のシートの再編集で、来年の日付に書き戻す: ' + JSON.stringify(ld.meeting));
  env.ss.deleteSheet(env.sheet('0325参加者'));
  env.ss.deleteSheet(env.sheet('0930参加者'));
});

// ---- 12) 番号（No）がまだ無いメンバー（Spreadingから足したばかりの新メンバー）：招待者が名簿に合い、その方の代理は「代理??」にしない ----
step('番号の無いメンバー', () => {
  const roster = env.sheet('メンバー名簿'), last = roster.getLastRow();
  roster.getRange(last + 1, 1, 1, MEMBER_HEAD.length).setValues([MEMBER_HEAD.map((h, i) => (i === 2 ? '新規 十子' : ''))]);
  const a = srv.analyzeCsvData([
    '参加者氏名,ふりがな,会社名,カテゴリー,招待者,メール,種別',
    '十子 招待,とおこ しょうたい,招待社,保険,新規 十子さん,t1@example.com,Visitor',
    '十子 代理,とおこ だいり,代理社,保険,新規 十子,t2@example.com,Substitute'].join('\n'));
  const v = a.rows.find((r) => r['種別'] === 'Visitor'), sub = a.rows.find((r) => r['種別'] === 'Substitute');
  ck(v && v['招待者'] === '新規 十子' && !v._needsInviterReview, '12) 番号の無いメンバーを、招待者として名簿に合わせない: ' + JSON.stringify(v && [v['招待者'], v._needsInviterReview]));
  ck(sub && sub._No === '代理', '12) 番号の無いメンバーの代理が「代理??」になる: ' + (sub && sub._No));
  env.ss.insertSheet('20261104参加者').getRange(1, 1, 2, 9).setValues([
    ['No.', '参加者氏名', 'ふりがな', 'カテゴリー', '会社名', '招待者', '備考', '種別', 'メール'],
    ['代理??', '十子 代理', '', '', '', '新規 十子', '', 'Substitute', '']]);
  const ld = srv.loadSheetData('20261104参加者');
  ck(ld.rows[0] && ld.rows[0]._No === '代理', '12) 再編集で、番号の無いメンバーの代理が「代理??」のまま: ' + (ld.rows[0] && ld.rows[0]._No));
  ck(!srv.getMembersList().some((m) => m.name === '新規 十子'), '12) 番号で引く一覧（割り振り表・ビジターホスト）に、番号の無い方が入った');
  env.ss.deleteSheet(env.sheet('20261104参加者'));
  roster.getRange(last + 1, 1, 1, MEMBER_HEAD.length).clearContent();
});

// ---- 13) スプレッドシートを「コピーを作成」した：コピー元のチャプターが共有したPDF・メンバーブックを、コピーの中身で上書きしない ----
step('スプレッドシートのコピー', () => {
  const orig = env.fileText(id0930);
  env.props.BNI_SPREADSHEET_ID = 'ORIGINAL_SS';          // 控えのIDはコピー元のもの（プロパティごとコピーされた）
  env.props.MEMBER_BOOK_ID = 'ORIGINAL_BOOK';
  srv.getSS_();
  ck(!env.props['VISITOR_PDF_ID_20260930参加者'] && !env.props.MEMBER_BOOK_ID && !env.props.LATEST_VISITOR_LIST_URL && env.props.BNI_SPREADSHEET_ID === 'SSID',
     '13) コピーで、コピー元のファイルの控えが残った: ' + JSON.stringify(Object.keys(env.props).filter((k) => /PDF|BOOK|LATEST/.test(k))));
  const a = srv.analyzeCsvData(CSV_0930);
  srv.createFinalSheet('2026/09/30', '2026/9/30(水) 第535回', a.rows, a.header);
  ck(env.fileText(id0930) === orig && env.props['VISITOR_PDF_ID_20260930参加者'] && env.props['VISITOR_PDF_ID_20260930参加者'] !== id0930,
     '13) コピーで作った名簿のPDFで、コピー元の共有したPDFを上書きした');
  ck(env.props.GEMINI_API_KEY === 'secret-key' && env.props.MAIL_TPL_BCC === 'host@example.com', '13) コピーで、ファイル以外の設定（APIキー・メールのBCC）まで消した');
});

ck(env.errors.length === 0, '途中でエラーの記録が出た: ' + env.errors.slice(0, 3).join(' / '));

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('毎週の流れ（候補・CSV・名簿とPDF・再編集・次の週・メール・割り振り表・PDFの確認）: 検査 ' + checks + ' 件 OK');

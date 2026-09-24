// ビジター情報の投稿文（visitor_post_srv.js / visitor_post.html）を、Node上で動かして確かめる。
//
//   node tools/check_visitor_post.js [参加者.json]
//
// 参加者.json を渡すと、その中身（参加者シートの行と同じ形の配列）で投稿文を作って表示する。
// 渡さなければ、架空の参加者で決まった文面になるかを確かめる。

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { loadPage } = require('./lib_minidom');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }

// ---- サーバー側 ----
const HOLIDAYS = ['2026/05/06', '2026/08/12', '2026/12/30'];
const srv = {
  console,
  Utilities: { formatDate: (d, tz, f) => {
    const p = (n) => ('0' + n).slice(-2);
    return f.replace('yyyy', d.getFullYear()).replace('MM', p(d.getMonth() + 1)).replace('dd', p(d.getDate()));
  } },
  PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
};
vm.createContext(srv);
for (const f of ['コード.js', 'visitor_post_srv.js']) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), srv, { filename: f });
}
srv.getHolidays = () => HOLIDAYS;

// 開催回の数え方（2026/3/18 が第509回。休会日は数えない）
ck(srv.meetingCountOf_(new Date(2026, 2, 18)) === 509, '2026/3/18 が第509回にならない');
ck(srv.meetingCountOf_(new Date(2026, 8, 30)) === 535, '2026/9/30 が第535回にならない: ' + srv.meetingCountOf_(new Date(2026, 8, 30)));
ck(srv.meetingCountOf_(new Date(2026, 7, 12)) === 0, '休会日（8/12）なのに回数が付いた');
ck(srv.meetingCountOf_(new Date(2026, 0, 7)) === 0, '基準より前なのに回数が付いた');

// 入金の状態の読み方
const paid = (row) => srv.visitorPaidOf_(row).paid;
ck(paid({ 支払いステータス: '支払済み' }) === true, '「支払済み」が入金済みにならない');
ck(paid({ 支払いステータス: '支払い済み' }) === true, '「支払い済み」が入金済みにならない');
ck(paid({ 支払いステータス: '未払い' }) === false, '「未払い」が未入金にならない');
ck(paid({ 支払いステータス: 'Paid' }) === true, '「Paid」が入金済みにならない');
ck(paid({ 支払いステータス: 'Unpaid' }) === false, '「Unpaid」が未入金にならない（paid を含むので順番が大事）');
ck(paid({ 支払状況: '入金済' }) === true, '日本語の列名「支払状況」を読めない');
ck(paid({ 費用: '3,000', 支払い済み: '3000' }) === true, '金額で入金済みと判断できない');
ck(paid({ 費用: '3000', 支払い済み: '0' }) === false, '金額で未入金と判断できない');
ck(paid({ 参加者氏名: 'だれか' }) === null, '手がかりが無いのに入金の状態を決めている');

// 行の読み方（種別・キャンセル・空白）
const pp = srv.visitorPostPerson_({ _No: 'G01', 種別: 'Guest', 参加者氏名: '見本　花子 ', ふりがな: 'みほん はなこ',
                                     カテゴリー: 'テスト', 招待者: '名簿 太郎', ステータス: 'キャンセル' });
ck(pp.type === 'guest' && pp.name === '見本 花子' && pp.cancelled === true, '行の読み方がおかしい: ' + JSON.stringify(pp));
ck(srv.visitorPostPerson_({ _No: '代理12', 種別: 'Substitute', 参加者氏名: 'A B' }).type === 'sub', '代理を見分けられない');
ck(srv.visitorPostDateOf_('20260930参加者').getDate() === 30, 'シート名から開催日を読めない');

// ---- 架空の参加者シート ----
const ROWS = process.argv[2] ? JSON.parse(fs.readFileSync(process.argv[2], 'utf8')) : [
  { _No: 'V01', 種別: 'Visitor', 参加者氏名: '青山 一郎', ふりがな: 'あおやま いちろう', カテゴリー: '税理士', 招待者: '名簿 太郎', 支払いステータス: '支払済み', ステータス: '有効' },
  { _No: 'V02', 種別: 'Visitor', 参加者氏名: '井上 二郎', ふりがな: 'いのうえ じろう', カテゴリー: '工務店', 招待者: '名簿 花子', 支払いステータス: '未払い', ステータス: '有効' },
  { _No: 'V03', 種別: 'Visitor', 参加者氏名: '上野 三郎', ふりがな: 'うえの さぶろう', カテゴリー: '保険', 招待者: '名簿 太郎', 支払いステータス: '支払済み', ステータス: 'キャンセル' },
  { _No: 'V04', 種別: 'Visitor', 参加者氏名: '江藤 四郎', ふりがな: '', カテゴリー: '', 招待者: '', 支払いステータス: '', ステータス: '有効' },
  { _No: 'G01', 種別: 'Guest', 参加者氏名: '大野 五子', ふりがな: 'おおの いつこ', カテゴリー: '司法書士', 招待者: '名簿 次郎', 支払いステータス: '支払済み', ステータス: '有効' },
  { _No: '代理12', 種別: 'Substitute', 参加者氏名: '加藤 六美', ふりがな: 'かとう むつみ', カテゴリー: '', 招待者: '名簿 三郎', 支払いステータス: '未払い', ステータス: '有効' },
];
const HEADER = Object.keys(ROWS[0]).filter((k) => k !== '_No');
srv.loadSheetData = () => ({ rows: ROWS.map((r) => Object.assign({}, r)), header: HEADER });
srv.getRoutineInfo = () => ({ ok: true, found: true, meetingNo: '535' });
srv.getExistingVisitorSheets = () => ['20260930参加者', '20260923参加者'];

const data = srv.getVisitorPostData('20260930参加者');
ck(data.ok && data.meetingNo === '535' && data.month === 9 && data.day === 30, '開催回・開催日: ' + JSON.stringify([data.meetingNo, data.month, data.day]));
ck(data.hasPayColumn === true, '入金の状態の列があるのに無いことになっている');

// ---- 画面 ----
let savedFooter = null;
const page = loadPage('visitor_post.html', { fails, server: {
  getVisitorPostContext: () => srv.getVisitorPostContext(),
  getVisitorPostData: (n) => JSON.parse(JSON.stringify(srv.getVisitorPostData(n))),
  saveVisitorPostFooter: (t) => { savedFooter = t; return { ok: true, footer: t, message: '保存しました' }; },
} });
const { els, window, run, step } = page;
step('画面を開く', () => window.onload());
ck(els.sheet.value === '20260930参加者', '開催日の既定が ' + els.sheet.value);
ck(els.no.value === '535', '開催回が入っていない: ' + els.no.value);

const out = () => els.out.value;
if (!process.argv[2]) {
  const EXPECT = [
    '第535回(9/30)定例会ビジター情報 　確定',
    '',
    '● ビジター：2名',
    '',
    '◾️入金済み：1名',
    '•青山 一郎（あおやま いちろう）［税理士］招待者 名簿 太郎',
    '',
    '◾️未入金：1名',
    '•井上 二郎（いのうえ じろう）［工務店］招待者 名簿 花子',
    '',
    '● ゲスト：1名',
    '',
    '◽️入金済み：1名',
    '•大野 五子（おおの いつこ）［司法書士］招待者 名簿 次郎',
    '',
    '◽️未入金：0名',
    '',
    '● 代理：1名',
    '',
    '•加藤 六美（かとう むつみ）名簿 三郎さんの代理',
    '',
    '',
    '※招待者の方は[spreading]にご入力ください。',
    '',
    '■メンバーの皆様はSpreadingにてビジター様の詳細情報を事前に確認してください。',
  ].join('\n');
  // 江藤さんは入金の状態が読めないので「未入金」になる。ここでは載せない扱いにしてから比べる
  ck(/江藤 四郎/.test(els.note.innerHTML), '入金の状態が読めない方のお知らせが出ていない');
  ck(/上野 三郎/.test(els.note.innerHTML), 'キャンセルの方のお知らせが出ていない');
  ck(!/上野 三郎/.test(out()), 'キャンセルの方が投稿文に載っている');
  ck(/◾️未入金：2名/.test(out()), '読めない方が未入金に入っていない');
  step('江藤さんを載せない', () => run("setState(3,'skip')"));
  ck(out() === EXPECT, '投稿文が見本と違う:\n' + out() + '\n---- 見本 ----\n' + EXPECT);

  // 入金済みに切り替える → 数が変わる
  step('井上さんを入金済みに', () => run("setState(1,'paid')"));
  ck(/◾️入金済み：2名/.test(out()) && /◾️未入金：0名/.test(out()), '切り替えが投稿文に反映されない');
  ck(/入金済み 2・未入金 0/.test(els.sum.innerHTML), '件数の表示が変わらない');

  // 見出しの最後・回数
  step('速報にする', () => { els.word.value = '速報'; run('render()'); });
  ck(out().split('\n')[0] === '第535回(9/30)定例会ビジター情報 　速報', '見出しが「速報」にならない: ' + out().split('\n')[0]);
  step('見出しの最後なし', () => { els.word.value = ''; run('render()'); });
  ck(out().split('\n')[0] === '第535回(9/30)定例会ビジター情報', '見出しの最後を消せない: ' + out().split('\n')[0]);

  // 手で直したあとは、勝手に作り直さない
  step('手で直す', () => { els.out.value = '手で直した'; run('edited=true'); run("setState(1,'unpaid')"); });
  ck(out() === '手で直した', '手で直した内容が上書きされた');
  ck(/作り直す/.test(els.editNote.innerText), '作り直していないことのお知らせが無い');
  step('作り直す', () => run('render(true)'));
  ck(page.log.confirms.length === 1 && /◾️未入金：1名/.test(out()), '「作り直す」で作り直されない');

  // 末尾の文面を保存・コピー
  step('末尾を変えて保存', () => { els.footer.value = '※よろしくお願いします。'; run('saveFooter()'); });
  ck(savedFooter === '※よろしくお願いします。' && /よろしくお願いします。$/.test(out()), '末尾の文面が保存・反映されない');
  step('コピー', () => run('copyText()'));
  ck(page.log.copies === 1 && page.log.selected === 'out', 'コピーの操作が行われていない');
} else {
  console.log(out());
}

console.log(`\nビジター情報の投稿文: 検査 ${checks} 件`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.slice(0, 20).forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: 開催回・入金の読み方・投稿文・切り替え・保存・コピー');

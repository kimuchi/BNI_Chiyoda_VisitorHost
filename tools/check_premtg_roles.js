// 事前MTG（朝イチMTG）のパワポの役職のページ（premtg_srv.js）を、実データなしで確かめる。
// ルーティンチェックシートは架空の項目の並び（見せかけのスプレッドシート lib_sheet_fake.js）。
// 役職ごとの入力と同じ保存（saveRoleInput）で書き込み、同梱の既定のひな形から作る。名前はすべて架空。
//
//   node tools/check_premtg_roles.js
//
// 確かめること
//   ・ルーティンチェックシートの「書記兼会計より」「プレジデントより」の行（朝一MTGでその役職が話すこと）に書けば、
//     「今週の共有事項（…）」が空でも、その役職のページを作る
//     （以前は「今週の共有事項」だけを見ていて、「書記兼会計より」に書いた書記兼会計のページが出なかった）
//   ・両方に書けば両方入る。同じ中身なら1回だけ。「なし」はページを作らない。「その他のお知らせ」は役職のページに入れない
//   ・役職ごとの入力：「○○より」に書いてあれば、「今週の共有事項」は入力済みに数える（画面にもそう出す）。逆は数えない
//   ・PowerPointの改行（垂直タブ）など、XMLに書けない文字が入っていても、pptxが壊れない（そのページが消えない）

process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeEnv } = require('./lib_sheet_fake');
const { readZip } = require('./lib_zip');
const { loadPage } = require('./lib_minidom');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }
const J = (x) => JSON.stringify(x);

// ---- 架空の名簿・ルーティンチェックシート ----
const HEAD = ['No', '業種区分', '氏名', 'ふりがな', 'カテゴリー', '会社名', '役職', 'メモ', '写真ファイル名', '一言コメント',
  '紹介してほしい人', '協業したい人', '入会日', '更新日', '更新期限日', '会社での役職'];
const NAMES = ['見本 一郎', '試験 花子', '架空 三郎', '仮名 四郎', '例示 五月', '模擬 六助'];
const rosterRows = [HEAD].concat(NAMES.map((n, i) => HEAD.map((h) => (h === 'No' ? String(i + 1) : h === '氏名' ? n : ''))));
const PREV = '2026/09/30', DATE = '2026/10/07', NEXT = '2026/10/14';
// A:空 B:No C:内容 D・E:細目 F:担当 G:期日 H:曜日目安 I:備考 J〜:開催日（1列おき）
const row = (no, c, d, role, due, vals) => ['', no, c, d, '', role || '', due || '', '', '',
  (vals || [])[0] || '', '', (vals || [])[1] || '', '', (vals || [])[2] || ''];
const ROUTINE = [
  ['', '', '開催日', '', '', '', '', '', '', PREV, '', DATE, '', NEXT],
  ['', '', '定例会回数', '', '', '', '', '', '', '535', '', '536', '', '537'],
  ['', 'No', '内容', '', '', '担当', '期日', '曜日目安', '備考', '', '', '', '', ''],
  row('1', '朝一MTG', ''),
  row('', '', 'プレジデントより', 'プレジ', '2日前', ['前回のプレジデントより']),
  row('', '', 'バイスプレジデントより', 'バイス', '2日前'),
  row('', '', '書記兼会計より', '書記兼会計', '2日前'),
  row('', '', 'その他のお知らせ', 'プレジ'),
  row('2', '代理・欠席', ''),
  row('', '', '代理', 'バイス', '当日'),
  row('', '', '欠席', 'バイス', '当日'),
  row('', '', '医療欠席', 'バイス', '当日'),
  row('3', 'ウィークリープレゼン', '', 'プレジ'),
  row('4', 'メインプレゼン', '', '書記兼会計', '5日前'),
];
const SHEET = '【24期】ルーティンチェックシート';

// 2026/10/5（月）10:00 に使う（次の定例会は 10/7）
const env = makeEnv({ now: new Date(2026, 9, 5, 10, 0, 0) });
const F = Object.assign({}, env.globals);
vm.createContext(F);
for (const f of fs.readdirSync(ROOT).filter((x) => /\.js$/.test(x)).sort()) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), F, { filename: f });
}
env.reset([
  ['メンバー名簿', false, rosterRows],
  ['休会日', true, [['2026/12/30']]],
  [SHEET, false, ROUTINE],
], { BNI_CHAPTER: J({ name: '見本', termBase: 23, meetingBaseDate: '2026/03/18', meetingBaseCount: 509 }) });
// 既定のひな形（premtg_template.html）を読む。保存先（Drive）と写真は使わない
F.HtmlService = { createHtmlOutputFromFile: (n) => ({ getContent: () => fs.readFileSync(path.join(ROOT, n + '.html'), 'utf8') }) };
let SAVED = null;
F.saveOutputFile_ = (blob, name) => { SAVED = { blob, name }; return { id: 'out', url: 'https://example/' + name, downloadUrl: 'https://example/dl/' + name }; };
F.findPhotoIdForName_ = () => '';

const holders = F.saveRoleHolders({ president: NAMES[0], vice: NAMES[1], secretary: NAMES[2], vhc: NAMES[3], mentor: NAMES[4], ec: NAMES[5] }, 24, DATE);
ck(holders && holders.ok, '担当者を保存できない: ' + J(holders && holders.message));

// 役職ごとの入力と同じ保存
function saveAs(role, values) {
  const ctx = F.getRoleInputContext(DATE, role);
  if (!ctx.ok) throw new Error(role + ' の画面を開けない: ' + ctx.message);
  const entries = Object.entries(values).map(([title, value]) => {
    const id = ctx.order.find((k) => ctx.items[k].title === title && ctx.items[k].roles.includes(role));
    if (!id) throw new Error(`${role} に「${title}」の項目がありません: ` + ctx.order.filter((k) => ctx.items[k].roles.includes(role)).map((k) => ctx.items[k].title).join(' / '));
    return { id, value, orig: ctx.items[id].value };
  });
  const r = F.saveRoleInput(DATE, role, entries);
  if (!r.ok || (r.skipped && r.skipped.length)) throw new Error(`${role} の保存: ${r.message} ${J(r.skipped || [])}`);
  return r;
}
const statusOf = (role) => {
  const ctx = F.getRoleInputContext(DATE, '');
  return (ctx.roles || []).find((r) => r.key === role).status;
};
const missingTitles = (role) => statusOf(role).missing.map((m) => m.title);

// ===== 1) 役職ごとの入力 =====
// バイスプレジデント：今週の共有事項だけ → 「バイスプレジデントより」は入力済みに数えない（逆は数えない）
saveAs('vice', { '今週の共有事項（バイスプレジデント）': '・チャプターの人数：目標53名、いま49名です。' });
ck(missingTitles('vice').includes('バイスプレジデントより') && !missingTitles('vice').includes('今週の共有事項（バイスプレジデント）'),
   '1) 今週の共有事項を書いたら「バイスプレジデントより」も入力済みになった（または今週の共有事項が未入力のまま）: ' + J(missingTitles('vice')));
// 書記兼会計：「書記兼会計より」だけに書く（PowerPointから写した改行＝垂直タブ入り）
const SEC_TEXT = '・10月分のチャプター運営費のお振込みをお願いします。\u000B・領収書は来週お渡しします。';
ck(missingTitles('secretary').includes('今週の共有事項（書記兼会計）'), '1) 書く前から書記兼会計の今週の共有事項が入力済み: ' + J(missingTitles('secretary')));
saveAs('secretary', { '書記兼会計より': SEC_TEXT, 'メインプレゼン': '見本さん、試験さん' });
ck(!missingTitles('secretary').includes('今週の共有事項（書記兼会計）'),
   '1) 「書記兼会計より」に書いたのに、今週の共有事項が未入力のまま: ' + J(missingTitles('secretary')));
{
  const ctx = F.getRoleInputContext(DATE, 'secretary');
  const share = ctx.items[ctx.order.find((k) => ctx.items[k].title === '今週の共有事項（書記兼会計）')];
  ck(share && share.alt && share.alt.title === '書記兼会計より', '1) 今週の共有事項に「書記兼会計より」との結びつきが無い: ' + J(share && share.alt));
}
// プレジデント：両方に同じ中身。バイス：「バイスプレジデントより」にも別の中身
const PRES = '・24期がはじまりました。役職の引継ぎをお願いします。';
saveAs('president', { 'プレジデントより': PRES, '今週の共有事項（プレジデント）': PRES, 'その他のお知らせ': '見本のその他のお知らせ' });
saveAs('vice', { 'バイスプレジデントより': '・ビジターは3名の予定です。' });
saveAs('mentor', { '今週の共有事項（メンターコーディネーター）': 'なし' });
saveAs('ec', { '今週の共有事項（エデュケーションコーディネーター）': '・今週のエデュケーション：見本さん' });

// ===== 2) 作る前の確かめ（割り付け）=====
const pv = F.getPreMeetingPreview(DATE);
ck(pv.ok, '2) 作る前の確かめ: ' + pv.message);
const labels = (pv.pages || []).map((pg) => pg.map((r) => r.label));
const flat = [].concat(...labels);
ck(flat.includes('書記兼会計') && !(pv.skipped || []).includes('書記兼会計'),
   '2) 「書記兼会計より」に書いたのに、書記兼会計のページを作らない: ' + J({ pages: labels, skipped: pv.skipped }));
ck(J(flat) === J(['プレジデント', 'バイスプレジデント', '書記兼会計', 'エデュケーションコーディネーター']),
   '2) 役職のページの並び: ' + J(labels));
ck((pv.skipped || []).includes('メンターコーディネーター'), '2) 「なし」の役職のページを作る: ' + J(pv.skipped));
const sec = [].concat(...(pv.pages || [])).find((r) => r.label === '書記兼会計') || {};
ck(J(sec.from) === J(['書記兼会計より']), '2) 書記兼会計のページの出どころ: ' + J(sec.from));

// ===== 3) 中身（premtgData_）=====
const data = F.premtgData_(new Date(2026, 9, 7));
const textOf = (label) => (data.roles.find((r) => r.label === label) || {}).text;
ck(textOf('書記兼会計') === '・10月分のチャプター運営費のお振込みをお願いします。\n・領収書は来週お渡しします。',
   '3) 書記兼会計の共有事項（「書記兼会計より」・垂直タブは改行に）: ' + J(textOf('書記兼会計')));
ck(textOf('プレジデント') === PRES, '3) 同じ中身が2回入る: ' + J(textOf('プレジデント')));
ck(textOf('バイスプレジデント') === '・チャプターの人数：目標53名、いま49名です。\n・ビジターは3名の予定です。',
   '3) 両方に書いた中身（今週の共有事項 → バイスプレジデントより）: ' + J(textOf('バイスプレジデント')));
ck(!data.roles.some((r) => /その他のお知らせ/.test(r.text)), '3) 「その他のお知らせ」が役職のページに入った');

// ===== 4) パワポ =====
const res = F.generatePreMeetingSlides(DATE);
ck(res.ok, '4) パワポを作れない: ' + res.message);
if (res.ok && SAVED) {
  const zip = readZip(SAVED.blob._buf);
  const slides = Object.keys(zip).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n));
  const xmlOf = (n) => zip[n].toString('utf8');
  const texts = slides.map((n) => (xmlOf(n).match(/<a:t>([^<]*)<\/a:t>/g) || []).map((t) => t.replace(/<\/?a:t>/g, '')).join('|'));
  ck(texts.some((t) => /書記兼会計/.test(t) && /運営費のお振込み/.test(t) && /領収書は来週/.test(t)),
     '4) パワポに書記兼会計のページ（「書記兼会計より」の中身）が無い: ' + J(texts.map((t) => t.slice(0, 80))));
  ck(texts.join('|').split('役職の引継ぎ').length === 2, '4) プレジデントの共有事項が2回入った');
  ck(!texts.some((t) => /見本のその他のお知らせ/.test(t)), '4) 「その他のお知らせ」が入った');
  // XMLに書けない文字（垂直タブなど）が残っていない（残るとPowerPointが修復を求め、ページが消えることがある）
  const bad = slides.filter((n) => /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(xmlOf(n)));
  ck(!bad.length, '4) XMLに書けない文字がスライドに残った: ' + J(bad));
  ck(/^定例会20261007_（事前MTG）\d{8}\.pptx$/.test(SAVED.name), '4) ファイル名: ' + SAVED.name);
}

// ===== 5) XMLに書けない文字（ほかのスライドの組み立てでも使う escapeXml_）=====
const esc = F.escapeXml_('A\u000BB\u0001C😀D\uD83DE&<');
ck(esc === 'A BC😀DE&amp;&lt;', '5) XMLに書けない文字の扱い（垂直タブ→空白・制御文字と片割れのサロゲートは消す・絵文字は残す）: ' + J(esc));

// ===== 6) 役職ごとの入力の画面：「書記兼会計より」に書いてあれば、今週の共有事項は入力済み・その旨を出す =====
{
  const server = {
    getSystemVersion: () => 'test',
    getRoleInputContext: (d, r) => F.getRoleInputContext(d, r),
  };
  const page = loadPage('role_input.html', {
    server, fails, now: '2026-10-05T10:00:00',
    preprocess: (p) => p.replace(/<\?\s*var roleParam[\s\S]*?\?>/, '').replace('<?= roleParam ?>', 'secretary').replace('<?= viewParam ?>', ''),
  });
  page.flush();
  page.step('書記兼会計の画面を開く', () => page.window.onload());
  const i = page.run(`ids.findIndex(function(id){ return ctx.items[id].title==='今週の共有事項（書記兼会計）'; })`);
  const box = page.els['it_' + i];
  ck(i >= 0 && box && !/need/.test(box.className || '') && /書記兼会計より/.test(box.innerHTML || ''),
     '6) 画面で、今週の共有事項が未入力のまま・「書記兼会計より」のことが出ていない: ' + J({ i, cls: box && box.className, html: box && String(box.innerHTML).replace(/<[^>]+>/g, ' ').slice(0, 300) }));
  ck(!/まだ空欄/.test(page.els.progress.innerHTML), '6) 必須の項目がまだ空欄と出る: ' + page.els.progress.innerHTML.replace(/<[^>]+>/g, ''));
}

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('事前MTGの役職のページ: 検査 ' + checks + ' 件 OK: 「書記兼会計より」などの行からもページを作る・同じ中身は1回・'
  + '「なし」は作らない・入力済みの数え方・XMLに書けない文字');

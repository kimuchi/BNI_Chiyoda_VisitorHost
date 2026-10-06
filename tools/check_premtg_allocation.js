// 事前MTG（朝イチMTG）のパワポの割り振り表のページ（premtg_srv.js）を、実データなしで確かめる。
// ルーティンチェックシート・割り振り表は架空のもの（見せかけのスプレッドシート lib_sheet_fake.js）。名前はすべて架空。
//
//   node tools/check_premtg_allocation.js
//
// 確かめること
//   ・ビジターホストコーディネーターのページのすぐあとに、その開催日の「割り振り表」シートのビジターごとの表を入れる
//     （No.〜オリエンテーションの列そのまま。上の集計の行・下のメンバー別のルーム・オリエン担当表は入れない）
//   ・ビジターホストコーディネーターのページは、うしろの役職と2人で1枚にしない（割り振り表がすぐあとに来るように）
//   ・ビジターホストコーディネーターの共有事項が無い日は、その順番のところ（書記兼会計のページのあと）に入れる
//   ・行が多いときは字を小さくし（10ptまで）、入らなければページを分ける（見出しの行はどのページにも。どのビジターも1回だけ）。
//     表はページの中に収まる
//   ・割り振り表がまだ無い日は入れず、お知らせする。作る前の確かめ（画面）にも出す
//   ・できたファイルの部品のつながり（tools/pptx_integrity.py）

process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const os = require('os');
const path = require('path');
const vm = require('vm');
const { spawnSync } = require('child_process');
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
const NAMES = ['見本 一郎', '試験 花子', '架空 三郎', '仮名 四郎', '例示 五月', '模擬 六助', '空想 七海', '見本 みさき'];
const rosterRows = [HEAD].concat(NAMES.map((n, i) => HEAD.map((h) => (h === 'No' ? String(i + 1) : h === '氏名' ? n : ''))));
const DATE = '2026/10/07', NEXT = '2026/10/14';
const row = (no, c, d, role, vals) => ['', no, c, d, '', role || '', '', '', '', (vals || [])[0] || '', '', (vals || [])[1] || ''];
const ROUTINE = [
  ['', '', '開催日', '', '', '', '', '', '', DATE, '', NEXT],
  ['', '', '定例会回数', '', '', '', '', '', '', '536', '', '537'],
  ['', 'No', '内容', '', '', '担当', '期日', '曜日目安', '備考', '', '', ''],
  row('1', '朝一MTG', ''),
  row('', '', 'プレジデントより', 'プレジ'),
  row('', '', '書記兼会計より', '書記兼会計'),
  row('2', '代理・欠席', ''),
  row('', '', '欠席', 'バイス'),
];
const SHEET = '【24期】ルーティンチェックシート';

// ---- 割り振り表（3. ルーム・オリエン割り振り表 で作るシートと同じ並び：saveAllocationSheet）----
// 12名。ルームメンバーなどは改行で何人も（いただいた画面と同じくらいの量）
const VIS = Array.from({ length: 12 }, (_, i) => {
  const no = i < 10 ? 'V' + String(i + 1).padStart(2, '0') : '代理' + (20 + i);
  const pick = (k, n) => Array.from({ length: n }, (_, j) => NAMES[(i + k + j) % NAMES.length]).join('\n');
  return [no, '訪問' + '甲乙丙丁戊己庚辛壬癸子丑'[i] + ' 太郎', i % 3 ? '見本の業種' : '自動車用品の製造・販売（見本の長い業種名）',
          NAMES[i % NAMES.length], pick(1, 2 + (i % 3)), i === 4 ? '【合同】訪問乙 太郎 と同室' : NAMES[(i + 2) % NAMES.length],
          i === 4 ? '（同上）' : pick(3, 3), pick(5, 1 + (i % 2))];
});
const ALLOC_HEADS = ['No.', 'お名前', 'カテゴリー', '招待者', 'つなげたいメンバー', 'ファシリテーター', 'ルームメンバー', 'オリエンテーション'];
const blank8 = () => ['', '', '', '', '', '', '', ''];
const ALLOC = [
  ['見本チャプター 10/7 定例会 ビジター・見学者・代理様 割り振り表', '', '', '', '', '', '', ''],
  ['【ダッシュボード】 ビジター: 10名 / ゲスト: 0名 / 代理: 2名 / ルーム数: 11 / 1ルーム平均総人数: 5.1名', '', '', '', '', '', '', ''],
  blank8(),
  ALLOC_HEADS,
].concat(VIS, [blank8(), ['※ 見本の注意書き（割り振り表の下の文）', '', '', '', '', '', '', ''], blank8(),
  ['【メンバー別 ルーム・オリエン担当表】', '', '', '', '', '', '', ''], ['No.', 'メンバー名', '担当ビジター・役割', '', '', '', '', ''],
  ['1', '見本 一郎', '訪問甲 太郎 様 (ファシリ)', '', '', '', '', '']]);

// 2026/10/5（月）10:00 に使う（次の定例会は 10/7）
const env = makeEnv({ now: new Date(2026, 9, 5, 10, 0, 0) });
const F = Object.assign({}, env.globals);
vm.createContext(F);
for (const f of fs.readdirSync(ROOT).filter((x) => /\.js$/.test(x)).sort()) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), F, { filename: f });
}
const reset = (withAlloc) => env.reset([
  ['メンバー名簿', false, rosterRows],
  ['休会日', true, [['2026/12/30']]],
  [SHEET, false, ROUTINE],
].concat(withAlloc ? [['20261007割り振り表', false, ALLOC]] : []),
{ BNI_CHAPTER: J({ name: '見本', termBase: 23, meetingBaseDate: '2026/03/18', meetingBaseCount: 509 }) });
F.HtmlService = { createHtmlOutputFromFile: (n) => ({ getContent: () => fs.readFileSync(path.join(ROOT, n + '.html'), 'utf8') }) };
let SAVED = null;
F.saveOutputFile_ = (blob, name) => { SAVED = { blob, name }; return { id: 'out', url: 'https://example/' + name, downloadUrl: 'https://example/dl/' + name }; };
F.findPhotoIdForName_ = () => '';

function setup(withAlloc, vhcText) {
  reset(withAlloc);
  const h = F.saveRoleHolders({ president: NAMES[0], vice: NAMES[1], secretary: NAMES[2], vhc: NAMES[3], mentor: NAMES[4], ec: NAMES[5] }, 24, DATE);
  if (!h || !h.ok) throw new Error('担当者を保存できない: ' + J(h));
  const saveAs = (role, values) => {
    const ctx = F.getRoleInputContext(DATE, role);
    const entries = Object.entries(values).map(([title, value]) => {
      const id = ctx.order.find((k) => ctx.items[k].title === title && ctx.items[k].roles.includes(role));
      if (!id) throw new Error(`${role} に「${title}」の項目がありません`);
      return { id, value, orig: ctx.items[id].value };
    });
    const r = F.saveRoleInput(DATE, role, entries);
    if (!r.ok) throw new Error(role + ' の保存: ' + r.message);
  };
  saveAs('president', { '今週の共有事項（プレジデント）': '・見本のお知らせです。' });
  saveAs('secretary', { '今週の共有事項（書記兼会計）': '・見本の運営費のお知らせ。' });
  if (vhcText) saveAs('vhc', { '今週の共有事項（ビジターホストコーディネーター）': vhcText });
  saveAs('mentor', { '今週の共有事項（メンターコーディネーター）': '・見本のメンターのお知らせ。' });   // VHC と同じく短い
}
// できたパワポ：ページの並び（各ページの文字）・割り振り表のページの表
function deck() {
  const zip = readZip(SAVED.blob._buf);
  const xmlOf = (n) => zip[n].toString('utf8');
  const prs = xmlOf('ppt/presentation.xml'), rels = xmlOf('ppt/_rels/presentation.xml.rels');
  const order = [...prs.matchAll(/<p:sldId id="\d+" r:id="(rId\d+)"\/>/g)].map((m) => {
    const t = rels.match(new RegExp('<Relationship\\b[^>]*\\bId="' + m[1] + '"[^>]*>'))[0].match(/Target="([^"]+)"/)[1];
    return 'ppt/' + t;
  });
  const text = (x) => (x.match(/<a:t>([^<]*)<\/a:t>/g) || []).map((t) => t.replace(/<\/?a:t>/g, '')).join('|');
  const sz = prs.match(/<p:sldSz\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
  return { zip, order, xml: order.map(xmlOf), text: order.map((n) => text(xmlOf(n))), W: +sz[1], H: +sz[2] };
}
const kindOf = (t) => (/ビジター割り振り表/.test(t) ? '割り振り表' : /ビジターホストコーディネーター/.test(t) ? 'VHC'
  : /メンターコーディネーター/.test(t) ? 'メンター' : /書記兼会計/.test(t) ? '書記兼会計' : /プレジデント/.test(t) ? 'プレジデント' : '他');
// 表の行（セルの文字。段落は改行でつなぐ）
const tableRows = (x) => [...x.matchAll(/<a:tr\b[^>]*>([\s\S]*?)<\/a:tr>/g)].map((tr) =>
  [...tr[1].matchAll(/<a:tc>([\s\S]*?)<\/a:tc>/g)].map((tc) => [...tc[1].matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)]
    .map((p) => (p[1].match(/<a:t>([^<]*)<\/a:t>/g) || []).map((t) => t.replace(/<\/?a:t>/g, '')).join('')).join('\n')));

// ===== 1) 割り振り表があり、ビジターホストコーディネーターの共有事項もある日 =====
setup(true, '・ビジターは12名の予定です。');
{
  const pv = F.getPreMeetingPreview(DATE);
  ck(pv.ok && pv.allocation && pv.allocation.found && pv.allocation.rows === 12 && J(pv.allocation.heads) === J(ALLOC_HEADS),
     '1) 作る前の確かめに、割り振り表（12行・列の見出し）が出ない: ' + J(pv.allocation));
  const pages = (pv.pages || []).map((pg) => pg.map((r) => r.label).join('+'));
  ck(pages.includes('ビジターホストコーディネーター') && !pages.some((p) => /^ビジターホストコーディネーター\+/.test(p)),
     '1) ビジターホストコーディネーターのページを、うしろの役職と2人で1枚にした（割り振り表がすぐあとに来ない）: ' + J(pages));
  const res = F.generatePreMeetingSlides(DATE);
  ck(res.ok && /割り振り表 \d枚/.test(res.message) && /20261007割り振り表/.test(res.message), '1) パワポを作れない・お知らせに割り振り表が無い: ' + res.message);
  const d = deck(), kinds = d.text.map(kindOf);
  const vi = kinds.indexOf('VHC'), ai = kinds.indexOf('割り振り表');
  ck(vi >= 0 && ai === vi + 1, '1) 割り振り表のページがビジターホストコーディネーターのページのすぐあとに無い: ' + J(kinds));
  ck(kinds.indexOf('メンター') > kinds.lastIndexOf('割り振り表'), '1) 割り振り表のページのあとに、うしろの役職のページが無い: ' + J(kinds));
  const allocIdx = kinds.map((k, i) => (k === '割り振り表' ? i : -1)).filter((i) => i >= 0);
  // 表：見出しの行（どのページにも）＋ビジターの行。どのビジターも1回だけ・シートのとおり（改行も）
  const got = [];
  let headOk = true, fit = true;
  allocIdx.forEach((i, n) => {
    const x = d.xml[i], rows = tableRows(x);
    if (J(rows[0]) !== J(ALLOC_HEADS)) headOk = false;
    got.push(...rows.slice(1));
    const fr = x.match(/<p:graphicFrame>[\s\S]*?<p:xfrm><a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/);
    if (!fr || +fr[1] + +fr[3] > d.W || +fr[2] + +fr[4] > d.H) fit = false;
    ck(allocIdx.length === 1 || new RegExp('（' + (n + 1) + '/' + allocIdx.length + '）').test(d.text[i]), '1) ページを分けたのに（1/2）などが無い: ' + d.text[i].slice(0, 60));
  });
  ck(headOk, '1) 見出しの行が、割り振り表のどのページにも無い');
  ck(J(got) === J(VIS), '1) 割り振り表のビジターの行がシートのとおりでない（抜け・重なり・改行）: ' + J(got.map((r) => r[0])));
  ck(fit, '1) 表がページからはみ出す');
  const tablePt = Math.min(...[...d.xml[allocIdx[0]].matchAll(/<a:tc>[\s\S]*?sz="(\d+)"/g)].map((m) => +m[1] / 100));
  ck(tablePt >= 10 && tablePt <= 16, '1) 表の字の大きさ（10〜16pt）: ' + tablePt);
  ck(allocIdx.length >= 2, '1) 12名・何人ものルームメンバーで、1ページに詰め込んだ（字が小さくなりすぎる）: ' + allocIdx.length);
  // 上の集計の行・下のメンバー別の表・注意書きは入れない
  const all = allocIdx.map((i) => d.text[i]).join('|');
  ck(!/ダッシュボード|メンバー別|担当ビジター・役割|見本の注意書き/.test(all), '1) 集計の行・メンバー別の表・注意書きが入った: ' + all.slice(0, 200));
  // 部品のつながり
  const f = path.join(os.tmpdir(), 'premtg_alloc_check.pptx');
  fs.writeFileSync(f, SAVED.blob._buf);
  const r = spawnSync('python3', [path.join(ROOT, 'tools', 'pptx_integrity.py'), f], { encoding: 'utf8' });
  ck(r.status === 0, '1) pptx の部品のつながり: ' + (r.stdout || '') + (r.stderr || ''));
  fs.unlinkSync(f);
}

// ===== 2) ビジターホストコーディネーターの共有事項が無い日：その順番のところ（書記兼会計のページのあと）=====
setup(true, '');
{
  const res = F.generatePreMeetingSlides(DATE);
  const kinds = deck().text.map(kindOf);
  const si = kinds.lastIndexOf('書記兼会計'), ai = kinds.indexOf('割り振り表');
  ck(res.ok && !kinds.includes('VHC') && si >= 0 && ai === si + 1 && kinds.indexOf('メンター') > ai,
     '2) ビジターホストコーディネーターのページが無い日の割り振り表の場所: ' + J(kinds));
}

// ===== 3) 割り振り表がまだ無い日：入れずに知らせる =====
setup(false, '・ビジターは12名の予定です。');
{
  const pv = F.getPreMeetingPreview(DATE);
  ck(pv.ok && pv.allocation && !pv.allocation.found && pv.allocation.sheetName === '20261007割り振り表', '3) 作る前の確かめ: ' + J(pv.allocation));
  const res = F.generatePreMeetingSlides(DATE);
  const kinds = deck().text.map(kindOf);
  ck(res.ok && !kinds.includes('割り振り表') && /割り振り表（20261007割り振り表）がまだ無い/.test(res.message),
     '3) 割り振り表が無い日に入れた・知らせない: ' + J({ kinds, msg: res.message }));
  // 画面：作る前の確かめに「まだありません」と出す
  const page = loadPage('role_input.html', {
    server: { getSystemVersion: () => 'test', getRoleInputContext: (dd, rr) => F.getRoleInputContext(dd, rr) }, fails, now: '2026-10-05T10:00:00',
    preprocess: (p) => p.replace(/<\?\s*var roleParam[\s\S]*?\?>/, '').replace('<?= roleParam ?>', '').replace('<?= viewParam ?>', ''),
  });
  page.flush();
  page.step('一覧を開く', () => page.window.onload());
  page.step('作る前の確かめを出す', () => page.run('renderPremtg(' + J(pv) + ')'));
  ck(/割り振り表（20261007割り振り表）がまだありません/.test(page.els.premtgOut.innerHTML.replace(/<[^>]+>/g, '')),
     '3) 画面に「割り振り表がまだありません」が出ない: ' + page.els.premtgOut.innerHTML.replace(/<[^>]+>/g, '').slice(0, 300));
}

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('事前MTGの割り振り表のページ: 検査 ' + checks + ' 件 OK: ビジターホストコーディネーターのページのすぐあと・ビジターの表だけ（集計・メンバー別の表は入れない）・'
  + '字を小さく／ページを分ける・共有事項が無い日の場所・割り振り表が無い日・部品のつながり');

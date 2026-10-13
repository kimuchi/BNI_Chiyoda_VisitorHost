// メインプレゼンが1人の回・ディレクター（リージョンのディレクター・アンバサダー）が担当する回の、
// スピーカーローテーション（speaker_rotation_srv.js・書記兼会計の画面 role_input.html）と、
// それを使う所（前半スライドの表・Facebookの文・メインプレゼンターの画像・書記兼会計の初期値・トークスクリプト）を、実データなしで確かめる。
//
//   node tools/check_main_presen_rotation.js
//
// 確かめること
//   1) 回ごとの指定が無い保存（以前の保存）は、これまでと同じ割り当て（2名ずつ・前にも後ろにも数えられる）
//   2) 回ごとの指定（1名・ディレクター＋1名・ディレクターのみ）：その回はそのメンバーの人数だけ並びを進める
//      （ほかの方の番を飛ばさない）。起点より前の回へさかのぼっても同じ割り当て。過ぎた回の指定は保存で捨てる
//   3) ルーティンチェックシートで確定した回（「①見本さん」「①空想ED　②試験さん」「②なし」）も、書いてあるメンバーの人数で数える
//   4) 回の形（kind・people の kind/role）。ディレクターは前半テンプレートのページの氏名・カテゴリー（リージョン名）・会社名。
//      読めない・当たらないときは書いたまま（ローテーションは止まらない）
//   5) 前半スライドの表：1人の回はまん中に1つのセル（gridSpan）・ディレクターはカテゴリーの行にリージョン名・「なし」は出さない
//   6) Facebookの文：1人の回は（１）だけ・ディレクターは「氏名さん/リージョン名」でつながりたい人の行なし・さんさんにしない
//   7) トークスクリプト：1人の回の {メインプレゼン2} は「なし」・ディレクターは氏名のまま・{スピーカーローテーション} に「なし」が出ない
//   8) 書記兼会計の初期値（メインプレゼン・挨拶）：1名の回・ディレクターの回。書いた初期値をルーティンとして読み直せる
//   9) 画面（簡易DOM）：人数を選んで保存する・ディレクターを選ぶ（一覧が無ければ手で入れる）・画面とサーバーの割り当てと投稿文が同じ・
//      メインプレゼンターの画像（1人の回はまん中に1つ・ディレクターはそのページの写真とリージョン名）・表の画像（1人の回は1つのセル）
// 本番の *.js を全部、見せかけのスプレッドシート（lib_sheet_fake.js）の上で動かす。名前はすべて架空。

process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeEnv } = require('./lib_sheet_fake');
const { loadPage } = require('./lib_minidom');

const ROOT = path.join(__dirname, '..');
const SRC = fs.readdirSync(ROOT).filter((f) => /\.js$/.test(f)).sort().map((f) => [f, fs.readFileSync(path.join(ROOT, f), 'utf8')]);
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }
const J = (x) => JSON.stringify(x);

// --- 架空の名簿・ルーティンチェックシート・アンバサダー・ディレクターのページ ---
const HEAD = ['No', '業種区分', '氏名', 'ふりがな', 'カテゴリー', '会社名', '役職', 'メモ', '写真ファイル名', '一言コメント',
  '紹介してほしい人', '協業したい人', '入会日', '更新日', '更新期限日', '会社での役職'];
const NAMES = ['見本 一郎', '試験 二郎', '架空 三郎', '仮名 四郎', '例示 五郎', '模擬 六郎', '標本 七郎', '新入 八郎'];
const ROSTER = [HEAD].concat(NAMES.map((n, i) => HEAD.map((h) => (h === 'No' ? String(i + 1) : h === '氏名' ? n
  : h === 'カテゴリー' ? '見本の業種' + (i + 1) : h === '会社名' ? '見本会社' + (i + 1) : h === '協業したい人' ? '協業A' + (i + 1) + '・協業B' + (i + 1) : ''))));
const DATES = ['2026/10/07', '2026/10/14', '2026/10/21', '2026/10/28', '2026/11/04', '2026/11/11', '2026/11/18', '2026/11/25'];
const routineSheet = (main) => {
  const row = (label, vals) => ['', '', label, '', '', '書記', '', '', ''].concat(...DATES.map((d, i) => [vals[i] || '', '']));
  return [
    ['', '', '開催日', '', '', '', '', '', ''].concat(...DATES.map((d) => [d, ''])),
    ['', '', '定例会回数', '', '', '', '', '', ''].concat(...DATES.map((d, i) => [String(536 + i), ''])),
    ['', 'No', '内容', '', '', '担当', '期日', '曜日目安', '備考'],
    row('各担当者とポジティブな挨拶', []),
    row('メインプレゼン', main || []),
  ];
};
// 前半テンプレートのアンバサダー・ディレクターのページ（getWeeklyGuests の結果）。カテゴリーの枠にはリージョン名が書いてある
const GUESTS = [
  { name: '空想 太郎', role: 'BNI見本リージョン', company: '見本リージョン\n株式会社', hasPhoto: true, hidden: true, intro: false },
  { name: '仮想 花子', role: 'エグゼクティブディレクター', company: '', hasPhoto: false, hidden: true, intro: false },
];
const STATE = (extra) => Object.assign({ order: NAMES.slice(), excluded: [], anchor: { date: '2026/10/07', pointer: 0 },
  header: 'メインプレゼンテーション（各４分45秒）', notes: ['見本の注意書き'], updated: '2026-10-01T00:00:00.000Z' }, extra || {});

// 本番の *.js を読み込んだサーバー。o.main … ルーティンチェックシートの「メインプレゼン」（DATES の順）、o.state … 保存してある状態
function server(o) {
  const env = makeEnv({ now: o.now || new Date(2026, 9, 5, 10, 0, 0) });
  const F = Object.assign({}, env.globals);
  vm.createContext(F);
  for (const [f, s] of SRC) vm.runInContext(s, F, { filename: f });
  const props = { BNI_CHAPTER: J({ name: '見本', termBase: 23, meetingBaseDate: '2026/03/18', meetingBaseCount: 509 }) };
  if (o.state) props.BNI_SPEAKER_ROTATION = J(o.state);
  env.reset([['メンバー名簿', false, ROSTER], ['休会日', true, [['2026/12/30']]],
             ['【24期】ルーティンチェックシート', false, routineSheet(o.main)]], props);
  F.guestCalls = 0;
  F.getWeeklyGuests = o.guests || (() => { F.guestCalls++; return { ok: true, guests: JSON.parse(J(GUESTS)) }; });
  return { env, F };
}
const names = (w) => w.people.map((p) => p.name).join('・');
const plan = (weeks) => weeks.map((w) => w.date.slice(5) + ':' + w.kind + ':' + names(w)).join(' ');

// ===================== 1) 以前の保存（回ごとの指定なし）=====================
{
  const { F } = server({ state: STATE() });
  const r = F.getSpeakerRotationWeeks('2026/10/07');
  ck(r.ok && plan(r.weeks) === '10/07:two:見本 一郎・試験 二郎 10/14:two:架空 三郎・仮名 四郎 10/21:two:例示 五郎・模擬 六郎 '
     + '10/28:two:標本 七郎・新入 八郎 11/04:two:見本 一郎・試験 二郎', '1) 以前の保存の割り当て: ' + plan(r.weeks || []));
  ck(r.weeks.every((w) => w.people.length === 2 && w.people.every((p) => p.kind === 'member' && p.inMaster && !('role' in p))),
     '1) 以前の保存の回の形（kind: member）: ' + J(r.weeks[0]));
  // 起点より前の回（さかのぼって数える）
  const back = F.getSpeakerRotationWeeks('2026/09/30');
  ck(back.ok && names(back.weeks[0]) === '標本 七郎・新入 八郎' && names(back.weeks[1]) === '見本 一郎・試験 二郎', '1) さかのぼった回: ' + plan(back.weeks));
  ck(J(F.rotPairFor_(new Date(2026, 9, 21))) === J(['例示 五郎', '模擬 六郎']), '1) 書記兼会計の初期値（10/21）: ' + J(F.rotPairFor_(new Date(2026, 9, 21))));
  ck(F.guestCalls === 0, '1) ディレクターの回が無いのに、前半テンプレート（27MB）を読んだ');
  const all = F.getSpeakerRotation();
  ck(all.ok && J(all.weekSettings) === '{}' && all.guests === null && all.weeks.every((w) => w.kind === 'two'), '1) 管理画面の材料: ' + J({ ws: all.weekSettings, g: all.guests }));
  // 以前の並びの数え方（2名ずつ）と同じになる：いろいろな並び・対象外・起点で比べる
  const oldPairs = (order, skip, start, count) => {
    const n = order.length, pairs = []; if (!n) return { pairs: [], next: 0 };
    let idx = ((start % n) + n) % n, guard = 0; const limit = n * 4 + 10;
    while (pairs.length < count && guard < limit) {
      const picked = [];
      while (picked.length < 2 && guard < limit) { if (!skip[order[idx]]) picked.push(order[idx]); idx = (idx + 1) % n; guard++; }
      if (picked.length === 2) pairs.push(picked); else break;
    }
    return { pairs, next: idx };
  };
  const oldPrev = (order, skip, p) => {
    const n = order.length; if (!n) return p;
    let idx = ((p % n) + n) % n, guard = 0; const found = [];
    while (found.length < 2 && guard < n * 4 + 10) { idx = (idx - 1 + n) % n; if (!skip[order[idx]]) found.push(idx); guard++; }
    return found.length === 2 ? found[1] : p;
  };
  let seed = 7, diff = 0, compared = 0;
  const rnd = (k) => {                                     // 決まった並びの乱数（mulberry32）
    seed = (seed + 0x6D2B79F5) | 0; let x = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    x = (x + Math.imul(x ^ (x >>> 7), 61 | x)) ^ x; return ((x ^ (x >>> 14)) >>> 0) % k;
  };
  for (let t = 0; t < 400; t++) {
    const n = 3 + rnd(10), order = Array.from({ length: n }, (_, i) => 'M' + i), skip = {};
    order.forEach((m) => { if (rnd(4) === 0) skip[m] = 'excluded'; });
    const start = rnd(n * 2), count = 1 + rnd(6), st = { order };
    const a = oldPairs(order, skip, start, count);
    if (a.pairs.length < count) continue;                 // 以前の数え方の打ち切り（全員が対象外に近いとき）は比べない
    compared++;
    const b = F.rotPairsFrom_(st, skip, start, count);
    if (J(a) !== J(b) || oldPrev(order, skip, start) !== F.rotPrevPointer_(st, skip, start)) diff++;
  }
  ck(compared > 200 && diff === 0, '1) 以前の並びの数え方と違う: ' + diff + '/' + compared);
}

// ===================== 2) 回ごとの指定 =====================
const WEEKS2 = { '2026/10/14': { members: 1 }, '2026/10/21': { members: 0, guest: '空想 太郎' }, '2026/10/28': { members: 1, guest: '空想' } };
const PLAN2 = '10/07:two:見本 一郎・試験 二郎 10/14:one:架空 三郎 10/21:director:空想 太郎 10/28:director1:空想 太郎・仮名 四郎 '
  + '11/04:two:例示 五郎・模擬 六郎';
{
  const { F, env } = server({ state: STATE({ weeks: WEEKS2 }) });
  const r = F.getSpeakerRotationWeeks('2026/10/07');
  ck(r.ok && plan(r.weeks) === PLAN2, '2) 回ごとの指定の割り当て（だれも飛ばさない）: ' + plan(r.weeks || []));
  const w21 = r.weeks[2], d = w21.people[0];
  ck(w21.people.length === 1 && d.kind === 'director' && d.role === 'BNI見本リージョン' && d.title === 'BNI見本リージョン'
     && d.company === '見本リージョン株式会社' && d.collab === '' && d.inMaster === false, '2) ディレクターの方（ページの氏名・リージョン名・会社名）: ' + J(d));
  ck(r.weeks[3].people[0].kind === 'director' && r.weeks[3].people[1].kind === 'member', '2) ディレクター＋1名はディレクターが先: ' + J(r.weeks[3].people));
  ck(r.weeks[1].people.length === 1, '2) 1名の回に空の2人目がある: ' + J(r.weeks[1].people));
  ck(F.guestCalls === 1, '2) 前半テンプレートのページを1回の呼び出しで何度も読んだ: ' + F.guestCalls);
  ck(names(F.getSpeakerRotationWeeks('2026/11/11').weeks[0]) === '標本 七郎・新入 八郎', '2) 先の回の割り当て');
  // 起点を後ろ（10/28 の仮名さんの位置）にしても、さかのぼって同じ割り当て
  const { F: G } = server({ state: STATE({ weeks: WEEKS2, anchor: { date: '2026/10/28', pointer: 3 } }) });
  ck(plan(G.getSpeakerRotationWeeks('2026/10/07').weeks) === PLAN2, '2) 起点より前へさかのぼった割り当て: ' + plan(G.getSpeakerRotationWeeks('2026/10/07').weeks));
  // 書記兼会計の初期値の材料（その回のローテーション）
  ck(J(F.rotPairFor_(new Date(2026, 9, 14))) === J(['架空 三郎']) && J(F.rotPairFor_(new Date(2026, 9, 21))) === J([])
     && J(F.rotPairFor_(new Date(2026, 9, 28))) === J(['仮名 四郎']), '2) rotPairFor_ が回ごとの指定どおりでない');
  ck(F.rotWeekFor_(new Date(2026, 9, 21)).kind === 'director' && F.rotWeekFor_(new Date(2026, 9, 21)).people[0].name === '空想 太郎', '2) rotWeekFor_ のディレクター');
  // 保存：過ぎた回の指定は捨てる・送られてこなければ（以前の画面）そのまま・ツールの取り込みでも残す
  const cur = F.getSpeakerRotation();
  ck(cur.ok && J(cur.weekSettings) === J(WEEKS2) && Array.isArray(cur.guests) && cur.guests[0].company === '見本リージョン株式会社',
     '2) 管理画面に回ごとの指定・ディレクターの一覧: ' + J({ ws: cur.weekSettings, g: cur.guests }));
  let sv = F.saveSpeakerRotation({ order: cur.order, excluded: [], anchor: cur.anchor, base: cur.updated, header: cur.header, notes: cur.notes,
    weekSettings: Object.assign({ '2026/09/30': { members: 1 }, '2026/11/04': { members: 2 } }, WEEKS2) });
  const saved = JSON.parse(env.props.BNI_SPEAKER_ROTATION);
  ck(sv.ok && J(saved.weeks) === J(WEEKS2), '2) 過ぎた回・2名の回の指定を保存した: ' + J(saved.weeks));
  sv = F.saveSpeakerRotation({ order: sv.order, excluded: [], anchor: sv.anchor, base: sv.updated, header: sv.header, notes: sv.notes });
  ck(sv.ok && J(JSON.parse(env.props.BNI_SPEAKER_ROTATION).weeks) === J(WEEKS2), '2) 回ごとの指定を送らない保存で、指定が消えた');
  const im = F.importSpeakerRotation(J({ order: NAMES, excluded: [], pointer: 0, nextMeetingDate: '2026-10-07' }));
  ck(im.ok && J(JSON.parse(env.props.BNI_SPEAKER_ROTATION).weeks) === J(WEEKS2) && plan(im.weeks.slice(0, 5)) === PLAN2,
     '2) ツールから取り込むと回ごとの指定が消えた: ' + plan((im.weeks || []).slice(0, 5)));
  sv = F.saveSpeakerRotation({ order: im.order, excluded: [], anchor: im.anchor, base: im.updated, weekSettings: {} });
  ck(sv.ok && !('weeks' in JSON.parse(env.props.BNI_SPEAKER_ROTATION)) && sv.weeks.slice(0, 5).every((w) => w.kind === 'two'),
     '2) 回ごとの指定を全部外して保存できない');
}

// ===================== 3)〜6) ルーティンチェックシートで確定した回・表・Facebookの文 =====================
const MAIN3 = ['①見本さん', '①空想ED　②試験さん', '①架空さん　②なし', '', '', '①未知EDさん', '①仮想エグゼクティブディレクター'];
{
  const { F } = server({ state: STATE(), main: MAIN3 });
  const r = F.getSpeakerRotationWeeks('2026/10/07');
  ck(r.ok && plan(r.weeks) === '10/07:one:見本 一郎 10/14:director1:空想 太郎・試験 二郎 10/21:one:架空 三郎 '
     + '10/28:two:仮名 四郎・例示 五郎 11/04:two:模擬 六郎・標本 七郎', '3) 確定した回の人数で数える: ' + plan(r.weeks || []));
  ck(r.weeks.slice(0, 3).every((w) => w.source === 'routine') && r.weeks[1].people[0].role === 'BNI見本リージョン', '3) 確定した回の形: ' + J(r.weeks[1]));
  const all = F.getSpeakerRotation();
  ck(all.ok && all.rebased && J(all.anchor) === J({ date: '2026/10/28', pointer: 3 }), '3) 起点を確定した回のあとへ（仮名さんの位置）: ' + J(all.anchor));
  ck(all.routineSlots['2026/10/07'] === 1 && all.routineSlots['2026/10/14'] === 1 && all.routineSlots['2026/10/21'] === 1
     && all.routineSlots['2026/11/18'] === 0, '3) 確定した回のメンバーの人数: ' + J(all.routineSlots));
  const w = (md) => all.weeks.find((x) => x.date.slice(5) === md);
  ck(names(w('11/11')) === '未知ED' && w('11/11').people[0].kind === 'director' && w('11/11').people[0].role === '', '4) ページに無いディレクターは書いたまま: ' + J(w('11/11')));
  ck(names(w('11/18')) === '仮想 花子' && w('11/18').people[0].role === 'エグゼクティブディレクター', '4) 「仮想エグゼクティブディレクター」→ 名字でページの方に: ' + J(w('11/18')));
  // 前半テンプレートを読めなくても、ローテーションは出る（お名前は書いたまま）
  F.getWeeklyGuests = () => { throw new Error('テンプレートが未登録'); };
  const nr = F.getSpeakerRotationWeeks('2026/10/07');
  ck(nr.ok && names(nr.weeks[1]) === '空想ED・試験 二郎' && nr.weeks[1].kind === 'director1', '4) 前半テンプレートを読めないとき: ' + J(nr));
  F.getWeeklyGuests = () => ({ ok: true, guests: JSON.parse(J(GUESTS)) });

  // 5) 前半スライドの表
  const xml = F.rotTableXml_(10, 0, 0, 9000000, 6000000, { weeks: r.weeks, header: 'メインプレゼンテーション', notes: [] }).xml;
  const rows = [...xml.matchAll(/<a:tr\b[^>]*>([\s\S]*?)<\/a:tr>/g)].map((m) => [...m[1].matchAll(/<a:tc\b([^>]*)>([\s\S]*?)<\/a:tc>/g)]
    .map((c) => ({ span: (c[1].match(/gridSpan="(\d)"/) || [])[1] || '', merged: /hMerge="1"/.test(c[1]), text: (c[2].match(/<a:t>([^<]*)<\/a:t>/) || [])[1] || '' })));
  const week = (i) => rows.slice(1 + i * 3, 4 + i * 3);
  const one = week(0);
  ck(one.every((cells) => cells.length === 3 && cells[1].span === '2' && cells[2].merged) && one[1][1].text === '見本 一郎' && one[0][1].text === '見本の業種1'
     && one[2][1].text === '協業A1・協業B1', '5) 1人の回が1つのセル（gridSpan）になっていない: ' + J(one));
  const dir = week(1);
  ck(dir[0][1].text === 'BNI見本リージョン' && dir[1][1].text === '空想 太郎' && dir[2][1].text === '' && dir[1][2].text === '試験 二郎'
     && dir[0][2].text === '見本の業種2' && !dir[0][1].span, '5) ディレクターの回のカテゴリー・氏名: ' + J(dir));
  ck(!/なし/.test(xml) && week(2)[1][1].span === '2' && week(2)[1][1].text === '架空 三郎', '5) 「②なし」の回: ' + J(week(2)));
  ck(week(3).every((cells) => cells.length === 3 && !cells[1].span && !cells[2].merged), '5) 2名の回がこれまでどおりでない: ' + J(week(3)));
  // スライドの画面から1回目を差し替えたとき（2人目が空・「なし」）も、空の列にしない
  const xml2 = F.rotTableXml_(10, 0, 0, 9000000, 6000000, { weeks: [{ date: '2026/10/07', label: '10月7日', no: '536',
    people: [{ name: '試験 二郎', title: '見本の業種2', collab: '' }, { name: '', title: '', collab: '' }] },
    { date: '2026/10/14', label: '10月14日', no: '537', people: [{ name: '見本 一郎', title: 'X', collab: '' }, { name: 'なし' }] }], header: 'H', notes: [] }).xml;
  ck((xml2.match(/gridSpan="2"/g) || []).length === 7 && !/なし/.test(xml2), '5) 画面で差し替えた1回目（2人目が空・なし）: ' + (xml2.match(/gridSpan="2"/g) || []).length);

  // 6) Facebookの文
  const fb1 = F.rotFbText_(r.weeks[0], '例示 五郎');
  ck(/（１）見本 一郎さん\/見本の業種1\n　　　→つながりたい人：協業A1・協業B1\nスピーカーの繋がりたい方/.test(fb1) && !/（２）/.test(fb1) && !/決まっていない/.test(fb1),
     '6) 1人の回の投稿文: ' + fb1);
  const fb2 = F.rotFbText_(r.weeks[1], '例示 五郎');
  ck(/（１）空想 太郎さん\/BNI見本リージョン\n（２）試験 二郎さん\/見本の業種2\n　　　→つながりたい人：協業A2・協業B2\n/.test(fb2)
     && (fb2.match(/つながりたい人/g) || []).length === 1, '6) ディレクター＋1名の投稿文: ' + fb2);
  const fb3 = F.rotFbText_(all.weeks.find((x) => x.date === '2026/11/11'), '例示 五郎');
  ck(/（１）未知EDさん\/\n/.test(fb3) && !/さんさん/.test(fb3) && !/つながりたい人/.test(fb3) && !/スピーカーの繋がりたい方/.test(fb3),
     '6) ディレクターだけの投稿文（さんさんにしない・つながりたい人なし）: ' + fb3);
  const two = F.rotFbText_(r.weeks[3], '例示 五郎');
  ck(/（１）仮名 四郎さん\/見本の業種4\n　　　→つながりたい人：協業A4・協業B4\n（２）例示 五郎さん\/見本の業種5\n　　　→つながりたい人：協業A5・協業B5\nスピーカーの繋がりたい方を/.test(two),
     '6) 2名の投稿文がこれまでどおりでない: ' + two);
  ck(/決まっていない/.test(F.rotFbText_({ label: 'x', people: [] }, '')), '6) 決まっていない回');

  // 7) トークスクリプト
  const talk = (d) => { const e = F.talkEnv_(new Date(d)); return (k) => F.talkValue_(k, e); };
  let tv = talk('2026-10-07T00:00:00');
  ck(J(tv('メインプレゼン2')) === J({ value: 'なし', state: 'none' }) && tv('メインプレゼン').value === '見本さん' && tv('メインプレゼン1').value === '見本',
     '7) 1人の回の差し込み: ' + J(['メインプレゼン', 'メインプレゼン1', 'メインプレゼン2'].map(tv)));
  const fill = (text, t) => F.talkFill_(text, t, () => {});
  ck(fill('続いてのメインプレゼンターは　{メインプレゼン2}さんです！', tv) === '続いてのメインプレゼンターは　なしです！', '7) 1人の回の台本: ' + fill('{メインプレゼン2}さん', tv));
  const rot7 = tv('スピーカーローテーション').value;
  ck(/^10\/14\(水\) 第537回　空想 太郎さん・試験さん\n10\/21\(水\) 第538回　架空さん\n10\/28\(水\) 第539回　仮名さん・例示さん\n/.test(rot7)
     && !/なし|さんさん/.test(rot7), '7) {スピーカーローテーション}: ' + rot7);
  tv = talk('2026-10-14T00:00:00');
  ck(tv('メインプレゼン1').value === '空想 太郎' && tv('メインプレゼン2').value === '試験' && tv('メインプレゼン').value === '空想 太郎さん・試験さん'
     && fill('トップバッターは、そう！{メインプレゼン1}さん', tv) === 'トップバッターは、そう！空想 太郎さん', '7) ディレクターの差し込み: ' + J(['メインプレゼン', 'メインプレゼン1'].map(tv)));
  tv = talk('2026-10-21T00:00:00');
  ck(tv('メインプレゼン').value === '架空さん' && tv('メインプレゼン2').state === 'none', '7) 「②なし」の回: ' + J(tv('メインプレゼン')));
  tv = talk('2026-11-11T00:00:00');
  ck(tv('メインプレゼン1').value === '未知ED' && tv('メインプレゼン').value === '未知EDさん', '7) ページに無いディレクター（さんさんにしない）: ' + J(tv('メインプレゼン')));

  // 8) 書記兼会計の初期値：挨拶（ルーティンに書いてある回）
  const est = (title, d) => F.roleEstimate_({ title, parent: '' }, F.roleEstimateEnv_(new Date(d), [], { order: [], items: {} }));
  ck((est('各担当者とポジティブな挨拶', '2026-10-14T00:00:00') || {}).value === '空想 太郎さん　試験さん', '8) ディレクターの回の挨拶: ' + J(est('各担当者とポジティブな挨拶', '2026-10-14T00:00:00')));
  ck((est('各担当者とポジティブな挨拶', '2026-10-21T00:00:00') || {}).value === '架空さん', '8) 「②なし」の回の挨拶: ' + J(est('各担当者とポジティブな挨拶', '2026-10-21T00:00:00')));
  ck((est('各担当者とポジティブな挨拶', '2026-11-11T00:00:00') || {}).value === '未知EDさん', '8) ページに無いディレクターの挨拶: ' + J(est('各担当者とポジティブな挨拶', '2026-11-11T00:00:00')));
}
{
  // 8) 書記兼会計の初期値：メインプレゼン（ローテーションの回ごとの指定から）・挨拶（ルーティンが空の回）
  const { F } = server({ state: STATE({ weeks: WEEKS2 }) });
  const est = (title, d) => F.roleEstimate_({ title, parent: '' }, F.roleEstimateEnv_(new Date(d), [], { order: [], items: {} }));
  const mr = (d) => (est('メインプレゼン', d) || {}).value;
  ck(mr('2026-10-07T00:00:00') === '①見本さん　②試験さん', '8) 2名の回のメインプレゼン: ' + mr('2026-10-07T00:00:00'));
  ck(mr('2026-10-14T00:00:00') === '①架空さん', '8) 1名の回のメインプレゼン: ' + mr('2026-10-14T00:00:00'));
  ck(mr('2026-10-21T00:00:00') === '①空想 太郎ディレクター', '8) ディレクターのみの回のメインプレゼン: ' + mr('2026-10-21T00:00:00'));
  ck(mr('2026-10-28T00:00:00') === '①空想 太郎ディレクター　②仮名さん', '8) ディレクター＋1名の回のメインプレゼン: ' + mr('2026-10-28T00:00:00'));
  // 書いた初期値をルーティンチェックシートとして読み直すと、同じ回になる（ディレクターと分かり、ページの方に当たる）
  const back = F.routineMainPresenters_(mr('2026-10-28T00:00:00'));
  ck(back.length === 2 && back[0].kind === 'director' && F.guestByName_(GUESTS, back[0].guest).name === '空想 太郎' && back[1].name === '仮名 四郎',
     '8) 初期値を読み直したとき: ' + J(back));
  ck((est('各担当者とポジティブな挨拶', '2026-10-28T00:00:00') || {}).value === '空想 太郎さん　仮名さん'
     && (est('各担当者とポジティブな挨拶', '2026-10-14T00:00:00') || {}).value === '架空さん', '8) ルーティンが空の回の挨拶（ローテーションから）: '
     + J(est('各担当者とポジティブな挨拶', '2026-10-28T00:00:00')));
}

// ===================== 9) 書記兼会計の画面（簡易DOM）=====================
// 画像は、描いた文字・写真・線を残す canvas の代わりで確かめる
function richCanvas() {
  const ops = [];
  let font = '';
  const cv = { width: 0, height: 0, _ops: ops };
  const ctx = {
    fillStyle: '', strokeStyle: '', lineWidth: 1, textAlign: '', textBaseline: '', lineCap: '', lineJoin: '',
    set font(v) { font = String(v); }, get font() { return font; },
    measureText(t) { const m = /(\d+(?:\.\d+)?)px/.exec(font); const s = m ? parseFloat(m[1]) : 10; let w = 0; for (const ch of String(t)) w += ch.charCodeAt(0) < 128 ? 0.55 : 1; return { width: w * s }; },
    fillText(t, x, y) { ops.push({ op: 'text', text: String(t), x, y, color: this.fillStyle, align: this.textAlign }); },
    drawImage(im, ...a) { ops.push({ op: 'image', src: im && im.src, x: a.length >= 8 ? a[4] : a[0] }); },
    moveTo(x, y) { this._p = [x, y]; },
    lineTo(x, y) { if (this._p) ops.push({ op: 'line', x1: this._p[0], y1: this._p[1], x2: x, y2: y }); this._p = [x, y]; },
    fillRect(x, y, w, h) { ops.push({ op: 'rect', x, y, w, h, color: this.fillStyle }); },
  };
  ['scale', 'strokeRect', 'save', 'restore', 'beginPath', 'rect', 'clip', 'closePath', 'fill', 'stroke', 'arc', 'arcTo'].forEach((k) => { ctx[k] = () => {}; });
  cv.getContext = () => ctx;
  cv.toDataURL = () => 'data:image/png;base64,' + Buffer.from(J({ w: cv.width, h: cv.height, ops })).toString('base64');
  return cv;
}
const decode = (url) => JSON.parse(Buffer.from(String(url || '').split(',')[1] || '', 'base64').toString() || '{"ops":[]}');

function openScreen(F, extra) {
  const calls = [];
  const srv = Object.assign({
    getSystemVersion: () => 'test',
    getRoleInputContext: () => ({ ok: false, message: '検査では読まない' }),
    getSpeakerRotation: () => { calls.push(['rot']); return F.getSpeakerRotation(); },
    saveSpeakerRotation: (d) => { calls.push(['rotSave', JSON.parse(J(d))]); return F.saveSpeakerRotation(JSON.parse(J(d))); },
    getRotationGuests: () => { calls.push(['guests']); return F.getRotationGuests(); },
    getMemberPhotosBase64: (n) => { calls.push(['photos', n]); return { ok: true, map: { '見本 一郎': 'data:image/png;base64,TUVNQkVS' } }; },
    getWeeklyGuestPhotos: () => { calls.push(['guestPhotos']); return { ok: true, photos: { '空想 太郎': 'data:image/png;base64,R1VFU1Q=' } }; },
    saveSpeakerRotationImage: () => ({ ok: true }),
  }, extra || {});
  const page = loadPage('role_input.html', {
    server: srv, fails, now: '2026-10-05T10:00:00',
    preprocess: (p) => p.replace(/<\?\s*var roleParam[\s\S]*?\?>/, '').replace('<?= roleParam ?>', 'secretary').replace('<?= viewParam ?>', 'rotation'),
  });
  const realCreate = page.document.createElement;
  page.document.createElement = (tag) => (tag === 'canvas' ? richCanvas() : realCreate(tag));
  page.sandbox.Image = class { set src(v) { this._src = v; this.width = 300; this.height = 400; if (this.onload) this.onload(); } get src() { return this._src; } };
  page.calls = calls;
  page.flush();
  page.window.onload();
  page.step('ローテーションを開く', () => {});
  return page;
}
{
  const { F, env } = server({ state: STATE({ weeks: WEEKS2 }) });
  const page = openScreen(F);
  const { els, run, step } = page;
  const rot = () => run('rot');
  const clientPlan = () => run('JSON.stringify(rot.weeks.map(function(w){ return w.date.slice(5)+":"+weekPeople(w).map(function(p){ return p.name; }).join("・"); }))');
  const serverPlan = (r) => J(r.weeks.map((w) => w.date.slice(5) + ':' + names(w)));
  ck(rot() && clientPlan() === serverPlan(F.getSpeakerRotation()), '9) 画面とサーバーの割り当てが違う: ' + clientPlan() + ' / ' + serverPlan(F.getSpeakerRotation()));
  ck((els.rotWeeks.innerHTML.match(/<tr/g) || []).length === 13, '9) 予定の表の行の数');
  const rowOf = (md) => { const m = els.rotWeeks.innerHTML.match(new RegExp(md.replace('/', '\\/') + '\\(水\\)[^]*?<\\/tr>')); return m ? m[0] : ''; };
  ck(/架空 三郎/.test(rowOf('10/14')) && els.rwk_1.value === 'one', '9) 1名の回の行: ' + rowOf('10/14').replace(/<[^>]+>/g, ' '));
  ck(/空想 太郎 <span class="badge">ディレクター<\/span>/.test(rowOf('10/21')) && /BNI見本リージョン/.test(rowOf('10/21')) && !/名簿にいない/.test(rowOf('10/21'))
     && els.rwk_2.value === 'director' && els.rwg_2.value === '空想 太郎', '9) ディレクターの回の行: ' + rowOf('10/21').replace(/<[^>]+>/g, ' '));
  ck(els.rwk_3.value === 'director1' && els.rwg_3.tag === 'select', '9) ディレクター＋1名の回の行');
  // Facebookの文：どの回も、サーバーと同じ文面
  const srvRot = F.getSpeakerRotation();
  let same = 0;
  for (let i = 0; i < 5; i++) {
    step('ご案内する回を選ぶ', () => { els.rotFbWeek.value = String(i); run('rotFbWeekChanged()'); });
    if (els.rotFb.value === F.rotFbText_(srvRot.weeks[i], srvRot.weeks[i].secretary || '')) same++;
    else fails.push('9) 投稿文がサーバーと違う（' + srvRot.weeks[i].date + '）: ' + els.rotFb.value.slice(0, 200));
  }
  checks++;
  // メインプレゼンターの画像：1人の回（10/14）
  const mp = () => decode(run('mpData')).ops;
  const texts = () => mp().filter((o) => o.op === 'text').map((o) => o.text);
  step('10/14 のご案内にする', () => { els.rotFbWeek.value = '1'; run('rotFbWeekChanged()'); });
  let t = texts();
  ck(run('mpCount') === 1 && els.mpName0.value === '架空 三郎' && t.includes('PRESENTER') && !t.some((x) => /PRESENTER[①②]/.test(x)),
     '9) 1人の回の画像（PRESENTER は1つ）: ' + J(t.slice(0, 8)));
  ck(!mp().some((o) => o.op === 'line' && o.x1 === 960 && o.x2 === 960), '9) 1人の回に真ん中の仕切りを描いた');
  ck(mp().some((o) => o.op === 'text' && o.text === '架空 三郎' && o.x === 681 + 475), '9) 1人の回のパネルがまん中にない: ' + J(mp().filter((o) => o.text === '架空 三郎')));
  ck(t.includes('ご紹介して欲しいカテゴリー') && t.includes('協業A3'), '9) 1人の回のカテゴリー');
  // 表の画像：1人の回は2つの列をまとめたセルのまん中に
  const rimg = decode(run('rotImgData')).ops.filter((o) => o.op === 'text');
  const w1 = 1000 - 24, c0 = Math.round(w1 * 0.17), c1 = Math.round((w1 - c0) / 2), c2 = w1 - c0 - c1;
  ck(rimg.some((o) => o.text === '架空 三郎' && Math.abs(o.x - (12 + c0 + (c1 + c2) / 2)) < 1), '9) 表の画像で1人の回が1つのセルのまん中にない: ' + J(rimg.filter((o) => o.text === '架空 三郎')));
  // ディレクターのみの回（10/21）：写真はそのページの写真、カテゴリーの所にリージョン名（見出しは描かない）
  step('10/21 のご案内にする', () => { els.rotFbWeek.value = '2'; run('rotFbWeekChanged()'); });
  t = texts();
  ck(run('mpCount') === 1 && els.mpName0.value === '空想 太郎' && els.mpRole0.value === 'BNI見本リージョン' && t.includes('BNI見本リージョン')
     && !t.includes('ご紹介して欲しいカテゴリー'), '9) ディレクターの画像: ' + J(t));
  ck(page.calls.some((c) => c[0] === 'guestPhotos') && mp().some((o) => o.op === 'image' && o.src === 'data:image/png;base64,R1VFU1Q='),
     '9) ディレクターの写真（前半テンプレートのページ）を描いていない: ' + J(mp().filter((o) => o.op === 'image')));
  ck(/ディレクターのページから/.test(els.mpPhotoNote0.innerText), '9) ディレクターの写真の知らせ: ' + els.mpPhotoNote0.innerText);
  ck(!page.calls.some((c) => c[0] === 'photos' && c[1].includes('空想 太郎')), '9) ディレクターをメンバー写真で探した');
  // ディレクター＋1名（10/28）：パネル2つ・仕切りあり
  step('10/28 のご案内にする', () => { els.rotFbWeek.value = '3'; run('rotFbWeekChanged()'); });
  t = texts();
  ck(run('mpCount') === 2 && els.mpName0.value === '空想 太郎' && els.mpName1.value === '仮名 四郎' && t.includes('PRESENTER①') && t.includes('PRESENTER②')
     && mp().some((o) => o.op === 'line' && o.x1 === 960), '9) ディレクター＋1名の画像: ' + J(t.slice(0, 10)));
  // 2名の回（10/7）はこれまでどおり
  step('10/7 のご案内にする', () => { els.rotFbWeek.value = '0'; run('rotFbWeekChanged()'); });
  ck(run('mpCount') === 2 && texts().includes('PRESENTER②') && els.mpName1.value === '試験 二郎' && els.mpCat0_0.value === '協業A1', '9) 2名の回の画像');

  // 人数を選ぶ：11/4 を1名に → 11/11 は模擬さんから（だれも飛ばさない）
  step('11/4 を1名にする', () => { els.rwk_4.value = 'one'; run('rotWeekKind(4)'); });
  ck(/例示 五郎/.test(rowOf('11/4')) && !/模擬 六郎/.test(rowOf('11/4')) && /模擬 六郎[^]*標本 七郎/.test(rowOf('11/11'))
     && /保存していない変更/.test(els.rotDirty.innerText), '9) 1名にした回のあと: ' + rowOf('11/11').replace(/<[^>]+>/g, ' '));
  // 11/18 をディレクターのみに：お名前が無いと保存しない → 一覧から選んで保存
  step('11/18 をディレクターのみにする', () => { els.rwk_6.value = 'director'; run('rotWeekKind(6)'); });
  ck(els.rwg_6 && els.rwg_6.tag === 'select' && els.rwg_6.options.some((o) => o.value === '仮想 花子'), '9) ディレクターの一覧から選べない');
  const nSave = page.calls.filter((c) => c[0] === 'rotSave').length;
  step('お名前なしで保存', () => run('rotSave()'));
  ck(page.calls.filter((c) => c[0] === 'rotSave').length === nSave && /11\/18/.test(els.rotMsg.innerHTML) && /お名前/.test(els.rotMsg.innerHTML),
     '9) ディレクターのお名前が無いのに保存した: ' + els.rotMsg.innerHTML);
  step('ディレクターを選ぶ', () => { els.rwg_6.value = '仮想 花子'; run('rotWeekGuest(6)'); });
  ck(/仮想 花子/.test(rowOf('11/18')) && /エグゼクティブディレクター/.test(rowOf('11/18')), '9) 選んだディレクターが予定に出ない: ' + rowOf('11/18').replace(/<[^>]+>/g, ' '));
  step('保存する', () => run('rotSave()'));
  const sent = page.calls.filter((c) => c[0] === 'rotSave').pop();
  ck(sent && J(sent[1].weekSettings) === J(Object.assign({}, WEEKS2, { '2026/11/04': { members: 1 }, '2026/11/18': { members: 0, guest: '仮想 花子' } })),
     '9) 送った回ごとの指定: ' + J(sent && sent[1].weekSettings));
  const stored = JSON.parse(env.props.BNI_SPEAKER_ROTATION).weeks;
  ck(stored['2026/11/04'].members === 1 && stored['2026/11/18'].guest === '仮想 花子' && /保存しました/.test(els.rotMsg.innerHTML), '9) 保存した状態: ' + J(stored));
  ck(clientPlan() === serverPlan(F.getSpeakerRotation()), '9) 保存したあと、画面とサーバーの割り当てが違う: ' + clientPlan());
  ck(F.getSpeakerRotation().weeks.find((w) => w.date === '2026/11/18').people[0].role === 'エグゼクティブディレクター', '9) 保存したディレクターのカテゴリー');
  // 一覧に無い方：「ほかの方」→ お名前を手で入れる
  step('ほかの方を選ぶ', () => { els.rwg_6.value = '__other__'; run('rotWeekGuest(6)'); });
  ck(els.rwg_6.tag === 'input', '9) ほかの方を選んでも、お名前を入れる欄にならない');
  step('お名前を入れる', () => { els.rwg_6.value = '未知 次郎'; run('rotWeekGuest(6)'); });
  ck(/未知 次郎 <span class="badge">ディレクター/.test(rowOf('11/18')), '9) 手で入れたディレクター: ' + rowOf('11/18').replace(/<[^>]+>/g, ' '));
  // 2名に戻す
  step('2名に戻す', () => { els.rwk_6.value = 'two'; run('rotWeekKind(6)'); });
  ck(run('JSON.stringify(rWeeks["2026/11/18"]||null)') === 'null' && /11\/18[^]*<td>[^<]+<\/td><td>[^<]+<\/td>/.test(rowOf('11/18')), '9) 2名に戻せない: ' + rowOf('11/18').replace(/<[^>]+>/g, ' '));
  // 画面の並びの数え方がサーバーと同じ（いろいろな人数で）
  let bad = 0;
  for (let i = 0; i < 40; i++) {
    const sizes = Array.from({ length: 6 }, (_, k) => (i * 7 + k * 3) % 3), start = i % 8;
    const c = run('JSON.stringify(rPairs(' + start + ', 6, ' + J(sizes) + '))');
    const s = J(F.rotPairsFrom_({ order: run('rOrder') }, {}, start, 6, sizes));
    if (c !== s) bad++;
    if (run('rPrev(' + start + ', ' + (i % 3) + ')') !== F.rotPrevPointer_({ order: run('rOrder') }, {}, start, i % 3)) bad++;
  }
  ck(bad === 0, '9) 画面とサーバーの並びの数え方が違う: ' + bad);
}
{
  // ディレクターの一覧が無い（前半テンプレートにページが無い）ときは、お名前を手で入れる
  const { F } = server({ state: STATE(), guests: () => ({ ok: true, guests: [] }) });
  const page = openScreen(F);
  const { els, run, step } = page;
  step('10/14 をディレクター＋1名にする', () => { els.rwk_1.value = 'director1'; run('rotWeekKind(1)'); });
  ck(page.calls.some((c) => c[0] === 'guests') && els.rwg_1 && els.rwg_1.tag === 'input', '9) 一覧が無いとき、お名前を入れる欄にならない');
  ck(/見つかりません/.test(els.rotGuestNote.innerText), '9) 一覧が無いことを知らせない: ' + els.rotGuestNote.innerText);
  step('お名前を入れる', () => { els.rwg_1.value = '空想 太郎'; run('rotWeekGuest(1)'); });
  step('保存する', () => run('rotSave()'));
  const sent = page.calls.filter((c) => c[0] === 'rotSave').pop();
  ck(sent && J(sent[1].weekSettings) === J({ '2026/10/14': { members: 1, guest: '空想 太郎' } }), '9) 手で入れたディレクターの保存: ' + J(sent && sent[1].weekSettings));
  // 10/14 は ディレクター＋架空さん、10/21 は仮名さん・例示さん（架空さんの相方の番は飛ばさない）
  const w = F.getSpeakerRotation().weeks;
  ck(names(w[1]) === '空想 太郎・架空 三郎' && names(w[2]) === '仮名 四郎・例示 五郎', '9) 保存したあとの割り当て: ' + plan(w.slice(0, 3)));
}

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('メインプレゼン（1人の回・ディレクターの回）のローテーション: 検査 ' + checks + ' 件 OK: 以前の保存は同じ割り当て・回ごとの指定とルーティンで確定した回はその人数だけ進む・'
  + 'ディレクターはページの氏名とリージョン名・表（1人はまん中のセル）・投稿文・トークスクリプト・書記兼会計の初期値・画面（人数を選んで保存・画像）');

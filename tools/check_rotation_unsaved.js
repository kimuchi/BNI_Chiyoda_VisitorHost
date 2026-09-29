// スピーカーローテーションを一度も保存していないとき（speaker_rotation_srv.js）を、実データなしで確かめる。
//
//   node tools/check_rotation_unsaved.js
//
//   ・同じ回（10/14）の発表者が、見た日（9/29・10/2・10/9）によって変わらない
//     （以前は見るたびに起点が「次の開催日の先頭から」になり、毎週ずれていた：前半スライドの表・書記兼会計の初期値・告知の文）
//   ・並び順は、保存するまでメンバー名簿の順のまま（名簿に足した方も入る）。「まだ保存していない」ことを画面に知らせる
//   ・並び順を保存したあとは、知らせない
// 本番の *.js を全部、見せかけのスプレッドシート（lib_sheet_fake.js）の上で動かす。名前はすべて架空。

process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeEnv } = require('./lib_sheet_fake');

const ROOT = path.join(__dirname, '..');
const FILES = fs.readdirSync(ROOT).filter((f) => /\.js$/.test(f)).sort();
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }
const J = (x) => JSON.stringify(x);

const HEAD = ['No', '業種区分', '氏名'];
const NAMES = ['見本 一郎', '試験 二郎', '架空 三郎', '仮名 四郎', '例示 五郎', '模擬 六郎', '空想 七郎', '新入 八郎'];
const roster = (names) => [HEAD].concat(names.map((n, i) => [String(i + 1), '', n]));
let props = {};                                   // 日をまたいで残るもの（スクリプトのプロパティ）
function onDay(date, names) {
  const env = makeEnv({ now: date });
  const srv = Object.assign({}, env.globals);
  vm.createContext(srv);
  for (const f of FILES) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), srv, { filename: f });
  env.reset([['メンバー名簿', false, roster(names || NAMES)], ['休会日', true, [['2026/12/30']]]], props);
  return { env, srv, keep: () => { props = Object.assign({}, env.props); } };
}
const pairOn = (srv, d) => {
  const r = srv.getSpeakerRotationWeeks(d);
  return r.ok ? { names: r.weeks[0].people.map((p) => p.name).join('・'), provisional: r.provisional } : { names: 'NG ' + r.message };
};

// 9/29・10/2・10/9 に、10/14 の回を見る
const seen = [];
[new Date(2026, 8, 29, 10, 0), new Date(2026, 9, 2, 10, 0), new Date(2026, 9, 9, 10, 0)].forEach((d) => {
  const { srv, keep } = onDay(d);
  seen.push(pairOn(srv, '2026/10/14'));
  keep();
});
ck(seen[0].names && !/^NG/.test(seen[0].names) && seen.every((x) => x.names === seen[0].names),
   '保存していないと、同じ回（10/14）の発表者が見た日で変わる: ' + J(seen.map((x) => x.names)));
ck(seen.every((x) => x.provisional === true), '保存していないことを、前半スライドの画面に知らせない: ' + J(seen));

// 名簿に足した方も、並び（名簿の順）に入る
{
  const { srv } = onDay(new Date(2026, 9, 9, 10, 0), NAMES.concat(['追加 九郎']));
  const all = srv.getSpeakerRotation();
  ck(all.ok && all.order[all.order.length - 1] === '追加 九郎' && all.provisional === true,
     '保存していないとき、名簿に足した方が並びに入らない・知らせない: ' + J({ order: all.order, provisional: all.provisional }));
}

// 並び順を保存したら、知らせない
{
  const { srv, keep } = onDay(new Date(2026, 9, 9, 10, 0));
  const cur = srv.getSpeakerRotation();
  const r = srv.saveSpeakerRotation({ order: cur.order, excluded: [], anchor: cur.anchor, base: cur.updated });
  ck(r.ok && r.provisional === false, '並び順を保存しても「まだ保存していない」と出る: ' + J({ ok: r.ok, provisional: r.provisional, msg: r.message }));
  keep();
  const { srv: srv2 } = onDay(new Date(2026, 9, 12, 10, 0));
  ck(pairOn(srv2, '2026/10/14').provisional === false, '保存したあと、前半スライドの画面に「まだ保存していない」と出る');
}

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('スピーカーローテーション（保存していないとき）: 検査 ' + checks + ' 件 OK: 同じ回の発表者が見た日で変わらない・名簿の順・まだ保存していないことを知らせる');

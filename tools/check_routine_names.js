// ルーティンチェックシートに書かれた名前を、メンバー名簿の方に合わせる処理（routine_srv.js の
// routineMemberName_・routineRosterName_）を、実データなしで確かめる。名前はすべて架空。
//
//   node tools/check_routine_names.js
//
// メインプレゼン・スタートアップ・推薦のことば・抽選・トークスクリプトの招待者・新メンバーなどが、これを通る。
//   ・同じ名字の方・似た氏名の方がいても、書いてある方に合わせる（名簿で先に並んでいる別の方にしない）
//   ・名字だけで2人以上に当たるときは決めない（画面で選んでもらう）
//   ・異体字（川辺／川邉）・「さん」「へ」「番」・カテゴリーの前置き・（代理）は、これまでどおり読む
//   ・名簿に無い新メンバーを、お名前の重なる方・1字違いの方にしない

process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeEnv } = require('./lib_sheet_fake');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }

const env = makeEnv({ now: new Date(2026, 8, 29, 10, 0, 0) });
const srv = Object.assign({}, env.globals);
vm.createContext(srv);
for (const f of fs.readdirSync(ROOT).filter((x) => /\.js$/.test(x)).sort()) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), srv, { filename: f });
const HEAD = vm.runInContext('MEMBER_HEADERS_', srv);
// 名簿の並び：似た方が先に来るようにしてある（以前は先に並んでいる方を選んでいた）
const ROSTER = [
  ['1', '見本 学', '税理士'], ['2', '見本 誠', '司法書士'], ['3', '川邉 真由子', '社労士'], ['4', '架空 次郎', '内装業'],
  ['5', '試験 花子', '保険'], ['6', '仮名 誠一', '印刷'], ['7', '新人 花代', 'デザイン'], ['8', '単独 太郎', '税理士'],
];
env.reset([['メンバー名簿', false, [HEAD].concat(ROSTER.map(([no, n, t]) => HEAD.map((h) => (h === 'No' ? no : h === '氏名' ? n : h === 'カテゴリー' ? t : ''))))]], {});
const who = (raw) => srv.routineMemberName_(raw);

// ---- 書いてある方に合わせる（先に並んでいる似た方にしない）----
[['見本誠さん', '見本 誠'], ['見本 誠', '見本 誠'], ['見本　誠さんへ', '見本 誠'], ['見本 学さん', '見本 学'],
 ['川辺さん', '川邉 真由子'], ['川辺真由子さん', '川邉 真由子'], ['内装業　架空さん', '架空 次郎'], ['３番架空さん', '架空 次郎'],
 ['試験 花子さん（代理）', '試験 花子'], ['単独さん', '単独 太郎'], ['単独くん', '単独 太郎'], ['見本 誠くん', '見本 誠'],
 ['試験 花了', '試験 花子']].forEach(([raw, want]) => {
  const r = who(raw);
  ck(r.matched && r.name === want, `「${raw}」→「${r.name}」（${want} のはず）`);
});

// ---- 名字だけで2人以上に当たるときは決めない ----
['見本さん', '見本くん'].forEach((raw) => {
  const r = who(raw);
  ck(!r.matched && r.name === '', `「${raw}」は名字が同じ方が2人いるのに「${r.name}」に決めた`);
});

// ---- 名簿に無い新メンバー（新入会の欄）：お名前が重なる方・1字違いの方にしない ----
const roster = ROSTER.map(([, n, t]) => ({ name: n, title: t }));
[['（行政書士）新人 花子さん', ''], ['新人 花子さん', ''], ['別人 花子さん', ''], ['見本 誠さん（司法書士）', '見本 誠'],
 ['司法書士 見本さん', '見本 誠'], ['川辺さん', '川邉 真由子']].forEach(([raw, want]) => {
  const got = srv.routineRosterName_(raw, roster, raw.indexOf('司法書士') >= 0 ? '司法書士' : '');
  ck(got === want, `新入会「${raw}」→「${got}」（${want || '名簿に無い'} のはず）`);
});

// ---- トークスクリプトの招待者（参加者シートの招待者の氏名）も、書いてある方のまま ----
ck(who('見本 誠').name === '見本 誠' && who('仮名 誠一').name === '仮名 誠一', 'トークスクリプトの招待者を、似た別の方にした');

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('ルーティンチェックシートの名前の照合: 検査 ' + checks + ' 件 OK: 同じ名字・似た氏名の方を取り違えない・名字だけで2人なら決めない・異体字や前置き・名簿に無い新メンバー');

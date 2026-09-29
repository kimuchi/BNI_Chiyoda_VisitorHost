// 期が替わるときの「役職・チーム（半期ごと）」の担当者と、メンバー名簿の「役職」を、実データなしで確かめる。
//
//   node tools/check_role_terms.js
//
// 本番と同じ *.js を全部、見せかけのスプレッドシート（lib_sheet_fake.js）の上で動かす。名前はすべて架空。
// 確かめること
//   ・担当者を何も保存していないとき、どの期も「未登録」。期が替わった日（10/1）に名簿の「役職」を消さない
//   ・前の版が保存した「空の24期」があっても、登録した期として扱わず、名簿の「役職」を消さない。24期は23期の担当者を使う
//   ・9月に23期の担当者を保存しても、空の24期を一緒に保存しない
//   ・24期の担当者を登録してあれば、10/1 に名簿の「役職」をその期の内容にする（これまでどおり）

process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeEnv } = require('./lib_sheet_fake');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }
const J = (x) => JSON.stringify(x);

const FILES = fs.readdirSync(ROOT).filter((f) => /\.js$/.test(f)).sort();
const HEAD = ['No', '業種区分', '氏名', 'ふりがな', 'カテゴリー', '会社名', '役職', 'メモ', '写真ファイル名', '一言コメント',
  '紹介してほしい人', '協業したい人', '入会日', '更新日', '更新期限日', '会社での役職'];
const ROSTER = [['1', '見本 一郎', 'プレジデント'], ['2', '試験 花子', '書記兼会計'], ['3', '架空 三郎', 'ビジターホスト'],
  ['4', '仮名 四郎', ''], ['5', '例示 五月', '']];

function server(now, props) {
  const env = makeEnv({ now });
  const srv = Object.assign({}, env.globals);
  vm.createContext(srv);
  for (const f of FILES) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), srv, { filename: f });
  env.reset([['メンバー名簿', false, [HEAD].concat(ROSTER.map(([no, name, role]) => HEAD.map((h) => (h === 'No' ? no : h === '氏名' ? name : h === '役職' ? role : ''))))]],
    Object.assign({ BNI_CHAPTER: J({ name: '見本', termBase: 23, meetingBaseDate: '2026/03/18', meetingBaseCount: 509 }) }, props || {}));
  return { env, srv };
}
const roles = (env) => {
  const v = env.values('メンバー名簿'), col = HEAD.indexOf('役職');
  return v.slice(1).map((r) => r[2] + ':' + (r[col] || '')).join(' | ');
};
const BEFORE = '見本 一郎:プレジデント | 試験 花子:書記兼会計 | 架空 三郎:ビジターホスト | 仮名 四郎: | 例示 五月:';
const blank = () => ({ president: '', vice: '', secretary: '', vhc: '', mentor: '', ec: '', web: '', support: '', training: '', event: '', bcp: '', spreading: '', gbc: '' });

// ---- A) 何も保存していない：10/1 に名簿を開いても、役職は消えない ----
{
  const { env, srv } = server(new Date(2026, 9, 1, 9, 0));
  const all = srv.roleHolderTerms_();
  ck(Object.keys(all).length === 0, 'A) 何も保存していないのに、登録した期がある: ' + J(all));
  ck(srv.roleHoldersOfTerm_(all, 24).registered === false, 'A) 24期が登録済みになっている');
  const mm = srv.getMemberMaster();
  ck(mm.ok, 'A) 名簿を開けない: ' + mm.message);
  ck(roles(env) === BEFORE, 'A) 10/1 に名簿を開いたら、役職が消えた: ' + roles(env));
}

// ---- B) 前の版が保存した「空の24期」がある：登録した期として扱わず、24期は23期の担当者を使う ----
{
  const h23 = Object.assign(blank(), { president: '見本 一郎', secretary: '試験 花子' });
  const { env, srv } = server(new Date(2026, 9, 1, 9, 0), { BNI_ROLE_HOLDERS_TERMS: J({ 23: h23, 24: blank() }) });
  srv.getMemberMaster();
  ck(roles(env) === BEFORE, 'B) 空の24期で、10/1 に名簿の役職が消えた: ' + roles(env));
  const h = srv.roleHoldersOfTerm_(srv.roleHolderTerms_(), 24);
  ck(h.registered === false && h.from === 23 && h.holders.president === '見本 一郎' && h.holders.secretary === '試験 花子',
     'B) 24期が、23期の担当者を使っていない: ' + J({ registered: h.registered, from: h.from, president: h.holders.president }));
  ck(srv.roleHolders_(new Date(2026, 9, 7)).secretary === '試験 花子', 'B) 10/7 の担当者（書記兼会計）が空になる');
}

// ---- C) 9月に23期の担当者を保存しても、空の24期は保存しない ----
{
  const { env, srv } = server(new Date(2026, 8, 29, 10, 0));
  const res = srv.saveRoleHolders({ president: '見本 一郎', secretary: '試験 花子' }, 23, '2026/09/30');
  const saved = JSON.parse(env.props.BNI_ROLE_HOLDERS_TERMS || '{}');
  ck(res.ok && Object.keys(saved).join(',') === '23', 'C) 23期を保存したら、ほかの期（空の24期）も保存された: ' + Object.keys(saved).join(','));
}

// ---- D) 24期の担当者を登録してあれば、10/1 に名簿の役職をその期の内容にする ----
{
  const h23 = Object.assign(blank(), { president: '見本 一郎', secretary: '試験 花子' });
  const h24 = Object.assign(blank(), { president: '例示 五月', secretary: '試験 花子' });
  const { env, srv } = server(new Date(2026, 9, 1, 9, 0), {
    BNI_ROLE_HOLDERS_TERMS: J({ 23: h23, 24: h24 }), BNI_ROLE_ROSTER_STATE: J({ applied: 23, seen: 23 }) });
  srv.getMemberMaster();
  ck(roles(env) === '見本 一郎: | 試験 花子:書記兼会計 | 架空 三郎:ビジターホスト | 仮名 四郎: | 例示 五月:プレジデント',
     'D) 24期を登録してあるのに、10/1 に名簿の役職がその期の内容にならない: ' + roles(env));
}

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('期の替わり目の担当者と名簿の役職: 検査 ' + checks + ' 件 OK: 未登録の期で役職を消さない・空の期を登録扱いしない・前の期の担当者を使う・登録した期は反映する');

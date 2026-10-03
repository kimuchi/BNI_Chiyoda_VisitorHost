// 「本日の招待者」（ルーティンチェックシートの項目。担当はwebマスター）を、参加者シートから自動で入れることを、実データなしで確かめる。
// 本番の *.js を見せかけのスプレッドシート（lib_sheet_fake.js）の上で動かす。名前はすべて架空。
//
//   node tools/check_inviters.js
//
// 確かめること
//   ・役職ごとの入力の初期値：参加者シート（20261007参加者）の、ビジター・ゲストを招待したメンバー（同じ方は1回）。
//     代理の方の「招待者」（代わりを頼んだメンバー）・キャンセルの方・招待者が空の方は入れない。同じ名字の方がいればフルネーム
//   ・参加者シートがまだ無い回は、初期値を出さない。ビジター・ゲストがいなければ「なし」
//   ・トークスクリプトの {ルーティン:本日の招待者}：チェックシートが空なら参加者シートから。書いてあれば書いてあるとおり

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

const HEAD = ['No', '業種区分', '氏名', 'ふりがな', 'カテゴリー', '会社名', '役職', 'メモ', '写真ファイル名', '一言コメント',
  '紹介してほしい人', '協業したい人', '入会日', '更新日', '更新期限日', '会社での役職'];
const NAMES = ['見本 一郎', '試験 花子', '架空 三郎', '仮名 四郎', '例示 五月', '模擬 六助', '見本 みさき'];
const rosterRows = [HEAD].concat(NAMES.map((n, i) => HEAD.map((h) => (h === 'No' ? String(i + 1) : h === '氏名' ? n : ''))));
const DATE = '2026/10/07', NEXT = '2026/10/14';
const ROUTINE = [
  ['', '', '開催日', '', '', '', '', '', '', DATE, '', NEXT],
  ['', '', '定例会回数', '', '', '', '', '', '', '536', '', '537'],
  ['', 'No', '内容', '', '', '担当', '期日', '曜日目安', '備考', '', '', ''],
  ['', '1', '本日の招待者', '', '', 'WEB', '', '', '【ビジター】PeatixのCSV貼り付けシートより招待者欄をコピーして貼り付け。重複を削除。', '', '', ''],
  ['', '2', 'アフターMTG', '', '', 'プレジ', '', '', '', '見本のアフターMTG', '', ''],
];
const PHEAD = ['No.', '参加者氏名', 'ふりがな', 'カテゴリー', '会社名', '招待者', '備考', '種別', 'ステータス'];
const PART = [PHEAD,
  ['V01', '来訪 一子', '', '税理士', '', '試験 花子', '', 'Visitor', ''],
  ['V02', '来訪 二子', '', '工務店', '', '架空 三郎', '', 'Visitor', ''],
  ['V03', '来訪 三子', '', '保険', '', '試験 花子', '', 'Visitor', ''],          // 同じ方が2人招待
  ['V04', '来訪 四子', '', '印刷', '', '仮名 四郎', '', 'Visitor', 'キャンセル'],   // キャンセル
  ['V05', '来訪 五子', '', '', '', '', '', 'Visitor', ''],                      // 招待者が空
  ['V06', '来訪 六子', '', 'デザイン', '', '見本 一郎', '', 'Visitor', ''],      // 同じ名字の方がいる
  ['G01', '来訪 七子', '', '司法書士', '', '例示 五月', '', 'Guest', ''],
  ['代理12', '代理 八子', '', '', '', '模擬 六助', '', 'Substitute', ''],          // 代理の方（代わりを頼んだメンバー）
];
const WANT = '試験さん、架空さん、見本一郎さん、例示さん';

const env = makeEnv({ now: new Date(2026, 9, 5, 10, 0, 0) });
const F = Object.assign({}, env.globals);
vm.createContext(F);
for (const f of fs.readdirSync(ROOT).filter((x) => /\.js$/.test(x)).sort()) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), F, { filename: f });
}
env.reset([
  ['メンバー名簿', false, rosterRows],
  ['休会日', true, [['2026/12/30']]],
  ['【24期】ルーティンチェックシート', false, ROUTINE],
  ['20261007参加者', false, PART],
], { BNI_CHAPTER: J({ name: '見本', termBase: 23, meetingBaseDate: '2026/03/18', meetingBaseCount: 509 }) });

const itemOf = (ctx) => ctx.items[ctx.order.find((k) => ctx.items[k].title === '本日の招待者')];

// ===== 1) 役職ごとの入力（webマスター）の初期値 =====
{
  const ctx = F.getRoleInputContext(DATE, 'web');
  const it = itemOf(ctx);
  ck(ctx.ok && it && it.roles.includes('web'), '1) webマスターの画面に「本日の招待者」が無い: ' + J({ ok: ctx.ok, msg: ctx.message }));
  const e = it && it.estimate;
  ck(e && e.value === WANT && e.mode === 'prefill' && /20261007参加者/.test(e.source) && /ビジター・ゲスト/.test(e.source),
     '1) 初期値（ビジター・ゲストの招待者・同じ方は1回・代理とキャンセルの方は入れない）: ' + J(e));
}
// 参加者シートがまだ無い回は、初期値を出さない
{
  const it = itemOf(F.getRoleInputContext(NEXT, 'web'));
  ck(it && !it.estimate, '1) 参加者シートが無い回に初期値が出る: ' + J(it && it.estimate));
}
// ビジター・ゲストがいない（代理の方とキャンセルの方だけ）なら「なし」
{
  const only = [PHEAD, PART[4], PART[8]];
  env.ss.getSheetByName('20261007参加者')._values = only.map((r) => r.slice());
  const e = itemOf(F.getRoleInputContext(DATE, 'web')).estimate;
  ck(e && e.value === 'なし', '1) ビジター・ゲストがいない回の初期値: ' + J(e));
  env.ss.getSheetByName('20261007参加者')._values = PART.map((r) => r.slice());
}

// ===== 2) トークスクリプトの {ルーティン:本日の招待者} =====
const talk = () => F.talkValue_('ルーティン:本日の招待者', F.talkEnv_(new Date(2026, 9, 7)));
{
  const r = talk();
  ck(r.state === 'ok' && r.value === WANT, '2) チェックシートが空のとき、参加者シートから入らない: ' + J(r));
  const built = F.talkBuild_(new Date(2026, 9, 7));
  const line = built.rows.map((x) => x.talk).find((t) => /本日ご招待いただきました/.test(t)) || '';
  ck(line.includes('本日ご招待いただきました' + WANT + '、ありがとうございました'), '2) 台本の文: ' + J(line));
  ck(!built.missing.includes('ルーティン:本日の招待者'), '2) 入らなかった差し込みに数えた: ' + J(built.missing));
}
// 役職ごとの入力で保存したら、書いてあるとおり
{
  const ctx = F.getRoleInputContext(DATE, 'web'), it = itemOf(ctx);
  const res = F.saveRoleInput(DATE, 'web', [{ id: it.id, value: '試験さん、例示さん', orig: it.value }]);
  ck(res.ok, '2) 保存できない: ' + res.message);
  const r = talk();
  ck(r.state === 'ok' && r.value === '試験さん、例示さん', '2) チェックシートに書いてあるのに、書いてあるとおりにならない: ' + J(r));
}

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('本日の招待者: 検査 ' + checks + ' 件 OK: 参加者シートのビジター・ゲストの招待者（同じ方は1回・代理とキャンセルの方は入れない）・'
  + '参加者シートが無い回・トークスクリプト（空なら参加者シートから・書いてあれば書いてあるとおり）');

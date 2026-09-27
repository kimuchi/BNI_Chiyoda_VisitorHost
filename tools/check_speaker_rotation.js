// スピーカーローテーション（speaker_rotation_srv.js）を、ルーティンチェックシートの実物で確かめる。
// 初期値は MP選出・ローテーション管理ツールの状態（10/14 第537回は並びの6番目から）。
//
//   node tools/check_speaker_rotation.js <routine.json> <members.json>
//
// 確かめること
//   ・ツールと同じ割り当てになるか（対象外は飛ばす・休会日の週は飛ばす）
//   ・ルーティンチェックシートにメインプレゼンがある回は、そちらを出す（確定）
//   ・並びを直すと、確定した回のあとへ起点を進めてから変わる（確定した回は動かない）
//   ・保存のぶつかり・ツールからの取り込み・Facebookの文

const { makeRoleServer } = require('./lib_role_fixture');

const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }

const S = makeRoleServer(process.argv[2], process.argv[3]);
const F = S.F;
const pairs = (r) => r.weeks.map((w) => `${w.md.replace(/\(.\)/, '')}:${w.people.map((p) => p.name.split(' ')[0]).join('・')}${w.source === 'routine' ? '*' : ''}`);
const byDate = (r, md) => r.weeks.find((w) => w.md.indexOf(md + '(') === 0);

// ===================== 初期値（ツールの状態）=====================
let r = F.getSpeakerRotation();
ck(r.ok, '読み込みに失敗: ' + r.message);
console.log('これからの予定: ' + pairs(r).join(' ／ '));
ck(r.anchor.date === '2026/10/14' && r.anchor.pointer === 5 && !r.rebased, '起点: ' + JSON.stringify(r.anchor) + ' rebased=' + r.rebased);
ck(r.openDate === '2026/10/07', 'メインプレゼンがまだ空のいちばん近い回: ' + r.openDate);
const w930 = byDate(r, '9/30');
ck(w930 && w930.source === 'routine' && w930.no === '535' && /^山本/.test(w930.people[0].name) && /^金子/.test(w930.people[1].name),
   '9/30（ルーティンで確定）: ' + JSON.stringify(w930));
ck(pairs(r).slice(1, 6).join(' ') === '10/7:渡邉・岡本 10/14:仲宗根・葉山 10/21:中込・竹田 10/28:星本・長見 11/4:徳山・平井',
   'ツールと同じ割り当て（熊谷さんなど対象外は飛ばす）: ' + pairs(r).slice(1, 6).join(' '));
ck(r.weeks.map((w) => w.no).join(',') === '535,536,537,538,539,540,541,542,543,544,545,546', '開催回: ' + r.weeks.map((w) => w.no).join(','));
ck(JSON.stringify(r.holidayHints) === JSON.stringify(['11/11', '2/17']), '休会日のお知らせ（ルーティンで開催回が空）: ' + JSON.stringify(r.holidayHints));
ck(/＜【9月30日定例会】メインプレゼンターのご案内＞/.test(r.fbText) && /（１）山本 登一郎さん\//.test(r.fbText)
   && /書記兼会計の原田 雅人まで/.test(r.fbText), 'Facebookの文: ' + r.fbText.slice(0, 120));
ck(r.missing.length === 0, '名簿にいない方: ' + JSON.stringify(r.missing));
ck(JSON.stringify(F.rotPairFor_(new Date('2026-10-07T00:00:00'))) === JSON.stringify(['渡邉 真理子', '岡本 翔太']), '10/7 の2名（役職ごとの入力で使う）');

// 前半スライドの表：その日から5回ぶん
const rw = F.getSpeakerRotationWeeks('2026/09/30');
ck(rw.ok && rw.weeks.length === 5 && rw.weeks[4].md.indexOf('10/28') === 0 && /各４分45秒/.test(rw.header) && rw.notes.length === 4,
   '前半の表の材料: ' + JSON.stringify(rw).slice(0, 160));

// ===================== 休会日の週は飛ばす =====================
const hol = S.sheets.find((s) => s.getName() === '休会日');
hol._grid.push([new Date('2026-10-21T00:00:00')]);
r = F.getSpeakerRotation();
ck(!byDate(r, '10/21') && pairs(r).slice(1, 5).join(' ') === '10/7:渡邉・岡本 10/14:仲宗根・葉山 10/28:中込・竹田 11/4:星本・長見',
   '10/21 を休会日にしたとき: ' + pairs(r).slice(1, 5).join(' '));
ck(byDate(r, '10/28').no === '538', '休会日のあとの開催回: ' + byDate(r, '10/28').no);
hol._grid.pop();

// ===================== 確定した回のあとへ起点を進める =====================
// ルーティンチェックシートに 10/7・10/14 のメインプレゼンを書く（書記兼会計が準備した回）
const s24 = S.sheets.find((s) => s.getName() === '【24期】ルーティンチェックシート');
const rowMp = s24._grid.findIndex((row) => String(row[3] || '').trim() === 'メインプレゼン');
const col = (d) => s24._grid[0].indexOf(d);
s24._grid[rowMp][col('2026/10/07')] = '①渡邉さん　②岡本さん';
s24._grid[rowMp][col('2026/10/14')] = '①仲宗根さん　②葉山さん';
r = F.getSpeakerRotation();
ck(r.rebased && r.anchor.date === '2026/10/21' && r.anchor.pointer === 8, '起点を 10/21（中込さんの位置）へ: ' + JSON.stringify(r.anchor));
ck(pairs(r).slice(1, 5).join(' ') === '10/7:渡邉・岡本* 10/14:仲宗根・葉山* 10/21:中込・竹田 10/28:星本・長見', '起点を進めても割り当ては同じ: ' + pairs(r).slice(1, 5).join(' '));

// 並びを直して保存（画面と同じく、位置で入れ替える。起点の位置はそのまま）
const order = r.order.slice(), iA = order.indexOf('長見 響児'), iB = order.indexOf('徳山 京介');
[order[iA], order[iB]] = [order[iB], order[iA]];
let sv = F.saveSpeakerRotation({ order, excluded: r.excluded, anchor: r.anchor, base: r.updated, header: r.header, notes: r.notes });
ck(sv.ok && /保存しました/.test(sv.message), '保存: ' + sv.message);
ck(pairs(sv).slice(1, 6).join(' ') === '10/7:渡邉・岡本* 10/14:仲宗根・葉山* 10/21:中込・竹田 10/28:星本・徳山 11/4:長見・平井',
   '入れ替えたあと（確定した回は動かない）: ' + pairs(sv).slice(1, 6).join(' '));
// 画面を開いたあとに、ほかの方が保存していたら止める
const stale = F.saveSpeakerRotation({ order, excluded: r.excluded, anchor: r.anchor, base: r.updated, header: r.header, notes: r.notes });
ck(!stale.ok && stale.conflict, 'ほかの方の保存とぶつかったのに保存した');
// 対象外にする（10/21 の中込さん）→ 竹田さん・星本さんに繰り上がる
r = F.getSpeakerRotation();
sv = F.saveSpeakerRotation({ order: r.order, excluded: r.excluded.concat(['中込 渉']), anchor: r.anchor, base: r.updated, header: r.header, notes: r.notes });
ck(sv.ok && byDate(sv, '10/21').people.map((p) => p.name.split(' ')[0]).join('・') === '竹田・星本', '中込さんを対象外にしたとき: ' + pairs(sv).slice(3, 5).join(' '));
// 名簿にいない方を並びに入れると知らせる。見出し・注意書きも保存される
r = sv;
sv = F.saveSpeakerRotation({ order: r.order.concat(['見本 退会']), excluded: r.excluded, anchor: r.anchor, base: r.updated,
                            header: 'メインプレゼンテーション（各５分）', notes: ['注意1', '', '注意2'] });
ck(sv.ok && JSON.stringify(sv.missing) === JSON.stringify(['見本 退会']) && sv.header === 'メインプレゼンテーション（各５分）'
   && JSON.stringify(sv.notes) === JSON.stringify(['注意1', '注意2']), '名簿にいない方・見出し・注意書き: ' + JSON.stringify([sv.missing, sv.header, sv.notes]));
ck(!F.saveSpeakerRotation({ order: [], excluded: [], anchor: sv.anchor, base: sv.updated }).ok, '空の並びを保存した');

// ===================== ツールから取り込む =====================
s24._grid[rowMp][col('2026/10/07')] = '';
s24._grid[rowMp][col('2026/10/14')] = '';
const toolJson = JSON.stringify({ order: vmDefault('order'), excluded: vmDefault('excluded'), pointer: 5,
  holidays: ['2026-08-12', '2026-12-30'], nextMeetingDate: '2026-10-14', nextMeetingNumber: 537, secretaryName: '原田 雅人' });
function vmDefault(k) { return JSON.parse(JSON.stringify(S.sandbox.ROT_DEFAULT_[k])); }
const im = F.importSpeakerRotation(toolJson);
ck(im.ok && im.anchor.date === '2026/10/14' && im.anchor.pointer === 5 && im.order.length === 46 && im.excluded.length === 6,
   '取り込み: ' + (im.message || '') + JSON.stringify(im.anchor));
ck(/12\/30/.test(im.message) && !/8\/12/.test(im.message.replace('2026/08/12', '')), '休会日に無いお休みのお知らせ: ' + im.message);
ck(pairs(im).slice(1, 4).join(' ') === '10/7:渡邉・岡本 10/14:仲宗根・葉山 10/21:中込・竹田', '取り込んだあとの割り当て: ' + pairs(im).slice(1, 4).join(' '));
ck(!F.importSpeakerRotation('{"order":[]}').ok, '空のデータを取り込んだ');

console.log(`\nスピーカーローテーション: 検査 ${checks} 件`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: ツールと同じ割り当て・休会日・確定した回・並びの直し方・ぶつかり・取り込み・Facebookの文');

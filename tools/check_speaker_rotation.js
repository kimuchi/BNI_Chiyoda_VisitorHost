// スピーカーローテーション（speaker_rotation_srv.js）を、ルーティンチェックシートの実物で確かめる。
// 始めの状態は、MP選出・ローテーション管理ツールの状態（10/14 第537回は並びの6番目から）。
// 本番ではコードに持たないので、検査の土台（lib_role_fixture.js）でスクリプトのプロパティに入れてある。
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
const TOOL = JSON.parse(S.props.BNI_SPEAKER_ROTATION);          // ツールの状態（取り込みの検査で使う）
const pairs = (r) => r.weeks.map((w) => `${w.md.replace(/\(.\)/, '')}:${w.people.map((p) => p.name.split(' ')[0]).join('・')}${w.source === 'routine' ? '*' : ''}`);
const byDate = (r, md) => r.weeks.find((w) => w.md.indexOf(md + '(') === 0);

// ===================== 始めの状態（ツールの状態）=====================
let r = F.getSpeakerRotation();
ck(r.ok, '読み込みに失敗: ' + r.message);
console.log('これからの予定: ' + pairs(r).join(' ／ '));
ck(r.anchor.date === '2026/10/14' && r.anchor.pointer === 5 && !r.rebased, '起点: ' + JSON.stringify(r.anchor) + ' rebased=' + r.rebased);
ck(r.openDate === '2026/10/07', 'メインプレゼンがまだ空のいちばん近い回: ' + r.openDate);
const w930 = byDate(r, '9/30');
ck(w930 && w930.source === 'routine' && w930.no === '535' && /^山内/.test(w930.people[0].name) && /^金井/.test(w930.people[1].name),
   '9/30（ルーティンで確定）: ' + JSON.stringify(w930));
ck(pairs(r).slice(1, 6).join(' ') === '10/7:川邉・岡林 10/14:仲里根・羽生 10/21:中西・梅田 10/28:星野・長尾 11/4:徳永・平松',
   'ツールと同じ割り当て（熊田さんなど対象外は飛ばす）: ' + pairs(r).slice(1, 6).join(' '));
ck(r.weeks.map((w) => w.no).join(',') === '535,536,537,538,539,540,541,542,543,544,545,546', '開催回: ' + r.weeks.map((w) => w.no).join(','));
ck(JSON.stringify(r.holidayHints) === JSON.stringify(['11/11', '2/17']), '休会日のお知らせ（ルーティンで開催回が空）: ' + JSON.stringify(r.holidayHints));
ck(/＜【9月30日定例会】メインプレゼンターのご案内＞/.test(r.fbText) && /（１）山内 登志夫さん\//.test(r.fbText)
   && /書記兼会計の原口 雅樹まで/.test(r.fbText), 'Facebookの文: ' + r.fbText.slice(0, 120));
ck(r.missing.length === 0, '名簿にいない方: ' + JSON.stringify(r.missing));
ck(JSON.stringify(F.rotPairFor_(new Date('2026-10-07T00:00:00'))) === JSON.stringify(['川邉 真由子', '岡林 翔吾']), '10/7 の2名（役職ごとの入力で使う）');

// 前半スライドの表：その日から5回ぶん
const rw = F.getSpeakerRotationWeeks('2026/09/30');
ck(rw.ok && rw.weeks.length === 5 && rw.weeks[4].md.indexOf('10/28') === 0 && /各４分45秒/.test(rw.header) && rw.notes.length === 4,
   '前半の表の材料: ' + JSON.stringify(rw).slice(0, 160));

// ===================== 休会日の週は飛ばす =====================
const hol = S.sheets.find((s) => s.getName() === '休会日');
hol._grid.push([new Date('2026-10-21T00:00:00')]);
r = F.getSpeakerRotation();
ck(!byDate(r, '10/21') && pairs(r).slice(1, 5).join(' ') === '10/7:川邉・岡林 10/14:仲里根・羽生 10/28:中西・梅田 11/4:星野・長尾',
   '10/21 を休会日にしたとき: ' + pairs(r).slice(1, 5).join(' '));
ck(byDate(r, '10/28').no === '538', '休会日のあとの開催回: ' + byDate(r, '10/28').no);
hol._grid.pop();

// ===================== 確定した回のあとへ起点を進める =====================
// ルーティンチェックシートに 10/7・10/14 のメインプレゼンを書く（書記兼会計が準備した回）
const s24 = S.sheets.find((s) => s.getName() === '【24期】ルーティンチェックシート');
const rowMp = s24._grid.findIndex((row) => String(row[3] || '').trim() === 'メインプレゼン');
const col = (d) => s24._grid[0].indexOf(d);
s24._grid[rowMp][col('2026/10/07')] = '①川邉さん　②岡林さん';
s24._grid[rowMp][col('2026/10/14')] = '①仲里根さん　②羽生さん';
r = F.getSpeakerRotation();
ck(r.rebased && r.anchor.date === '2026/10/21' && r.anchor.pointer === 8, '起点を 10/21（中西さんの位置）へ: ' + JSON.stringify(r.anchor));
ck(pairs(r).slice(1, 5).join(' ') === '10/7:川邉・岡林* 10/14:仲里根・羽生* 10/21:中西・梅田 10/28:星野・長尾', '起点を進めても割り当ては同じ: ' + pairs(r).slice(1, 5).join(' '));

// 並びを直して保存（画面と同じく、位置で入れ替える。起点の位置はそのまま）
const order = r.order.slice(), iA = order.indexOf('長尾 響一'), iB = order.indexOf('徳永 京平');
[order[iA], order[iB]] = [order[iB], order[iA]];
let sv = F.saveSpeakerRotation({ order, excluded: r.excluded, anchor: r.anchor, base: r.updated, header: r.header, notes: r.notes });
ck(sv.ok && /保存しました/.test(sv.message), '保存: ' + sv.message);
ck(pairs(sv).slice(1, 6).join(' ') === '10/7:川邉・岡林* 10/14:仲里根・羽生* 10/21:中西・梅田 10/28:星野・徳永 11/4:長尾・平松',
   '入れ替えたあと（確定した回は動かない）: ' + pairs(sv).slice(1, 6).join(' '));
// 画面を開いたあとに、ほかの方が保存していたら止める
const stale = F.saveSpeakerRotation({ order, excluded: r.excluded, anchor: r.anchor, base: r.updated, header: r.header, notes: r.notes });
ck(!stale.ok && stale.conflict, 'ほかの方の保存とぶつかったのに保存した');
// 対象外にする（10/21 の中西さん）→ 梅田さん・星野さんに繰り上がる
r = F.getSpeakerRotation();
sv = F.saveSpeakerRotation({ order: r.order, excluded: r.excluded.concat(['中西 渉']), anchor: r.anchor, base: r.updated, header: r.header, notes: r.notes });
ck(sv.ok && byDate(sv, '10/21').people.map((p) => p.name.split(' ')[0]).join('・') === '梅田・星野', '中西さんを対象外にしたとき: ' + pairs(sv).slice(3, 5).join(' '));
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
  holidays: ['2026-08-12', '2026-12-30'], nextMeetingDate: '2026-10-14', nextMeetingNumber: 537, secretaryName: '原口 雅樹' });
function vmDefault(k) { return JSON.parse(JSON.stringify(TOOL[k])); }
const im = F.importSpeakerRotation(toolJson);
ck(im.ok && im.anchor.date === '2026/10/14' && im.anchor.pointer === 5 && im.order.length === 46 && im.excluded.length === 6,
   '取り込み: ' + (im.message || '') + JSON.stringify(im.anchor));
ck(/12\/30/.test(im.message) && !/8\/12/.test(im.message.replace('2026/08/12', '')), '休会日に無いお休みのお知らせ: ' + im.message);
ck(pairs(im).slice(1, 4).join(' ') === '10/7:川邉・岡林 10/14:仲里根・羽生 10/21:中西・梅田', '取り込んだあとの割り当て: ' + pairs(im).slice(1, 4).join(' '));
ck(!F.importSpeakerRotation('{"order":[]}').ok, '空のデータを取り込んだ');

// ===================== Facebookに添付する画像を Drive に保存 =====================
{
  const zlib = require('zlib');
  const chunk = (type, data) => {
    const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
    const td = Buffer.concat([Buffer.from(type, 'ascii'), data]);
    const crc = Buffer.alloc(4); crc.writeUInt32BE(zlib.crc32(td));
    return Buffer.concat([len, td, crc]);
  };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(40, 0); ihdr.writeUInt32BE(20, 4); ihdr[8] = 8; ihdr[9] = 2;
  const png = Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr),
    chunk('IDAT', zlib.deflateSync(Buffer.alloc((40 * 3 + 1) * 20, 200))), chunk('IEND', Buffer.alloc(0))]);
  let saved = null;
  F.saveOutputFile_ = (blob, name) => { saved = { blob, name }; return { id: 'x', url: 'https://example/' + name, downloadUrl: 'https://example/dl/' + name }; };
  let r = F.saveSpeakerRotationImage('data:image/png;base64,' + png.toString('base64'), '20261007_スピーカーローテーション.png');
  ck(r.ok && saved && saved.name === '20261007_スピーカーローテーション.png' && saved.blob.getContentType() === 'image/png'
     && Buffer.compare(saved.blob._buf, png) === 0 && /03_生成物/.test(r.message), '画像の保存: ' + JSON.stringify(r));
  r = F.saveSpeakerRotationImage(png.toString('base64'), 'a/b:c');
  ck(r.ok && saved.name === 'a_b_c.png', 'ファイル名の整え方: ' + saved.name);
  saved = null;
  r = F.saveSpeakerRotationImage('', 'x.png');
  ck(!r.ok && !saved, '空の画像を保存した');
}

// ===================== まだ一度も保存していないとき =====================
// 並びはメンバー名簿の順・対象外なし・次回の定例会が並びの先頭から（メンバーの氏名はコードに持たない）
{
  delete S.props.BNI_SPEAKER_ROTATION;
  const fresh = F.getSpeakerRotation();
  const roster = F.getMemberMaster({ membersOnly: true }).members.map((m) => m.name);
  const loaded = F.rotLoad_();
  ck(loaded.anchor.pointer === 0 && loaded.anchor.date === '2026/09/30', '保存していないときの起点（次回の定例会・並びの先頭）: ' + JSON.stringify(loaded.anchor));
  // 9/30 はルーティンチェックシートで確定しているので、画面を開くと起点は 10/7（並びの3人目）へ進む
  ck(fresh.ok && JSON.stringify(fresh.order) === JSON.stringify(roster) && fresh.excluded.length === 0
     && fresh.rebased && fresh.anchor.date === '2026/10/07' && fresh.anchor.pointer === 2, '保存していないときの並び: ' + JSON.stringify([fresh.order.slice(0, 3), fresh.anchor]));
  const w = fresh.weeks.find((x) => x.source === 'rotation');
  ck(w && w.people.length === 2 && roster.indexOf(w.people[0].name) >= 0, '保存していないときの割り当て: ' + JSON.stringify(w && w.people.map((p) => p.name)));
  ck(!/\b(order|excluded)\b/.test(Object.keys(S.sandbox.ROT_DEFAULT_).join(' ')), 'コードに並び順（氏名）を持っている');
}

console.log(`\nスピーカーローテーション: 検査 ${checks} 件`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: ツールと同じ割り当て・休会日・確定した回・並びの直し方・ぶつかり・取り込み・Facebookの文・画像の保存・保存していないときの並び');

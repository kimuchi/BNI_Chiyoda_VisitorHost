// ルーティンチェックシートの読み取りを、実物のデータで確かめる。
// 本番と同じ routine_srv.js を、SpreadsheetApp まわりだけ差し替えてNodeで動かす。
//
//   node tools/check_routine.js <routine.json>

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const DATA = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));

// --- シートを真似る ---
function fakeSheet(name, grid) {
  return {
    getName: () => name,
    getLastRow: () => grid.length,
    getLastColumn: () => (grid[0] ? grid[0].length : 0),
    getRange(r, c, nr, nc) {
      return {
        getValues() {
          const out = [];
          for (let i = 0; i < nr; i++) {
            const row = grid[r - 1 + i] || [];
            out.push(row.slice(c - 1, c - 1 + nc));
          }
          return out;
        },
      };
    },
  };
}
const sheets = Object.keys(DATA).map((n) => fakeSheet(n, DATA[n]));

// --- 名簿（検証用。実際の氏名は使わない）---
const MEMBERS = ['谷口 大輔', '梅田 明日斗', '本間 圭吾', '梅中 公', '桒田 美香',
                 '山岸 由紀子', '田村 浩美', '藤本 礼子', '金井 美里', '細川 哲也',
                 '宮崎 佳子', '豊島 恵', '高村 舟', '平松 成子', '舩川 ちなつ', '杉山 太郎',
                 '山内 登志夫', '金井 高志', '泉 ゆか', '田村 秀二', '小西 美和',
                 '北川 光男', '若松 勇人', '岩渕 裕太', '文野 雅彦', '西澤 浩平', '上原 誠',
                 '川邉 真由子', '齋藤 一郎', '齊藤 次郎']
  .map((n) => ({ no: '1', name: n }));

const sandbox = {
  console,
  SpreadsheetApp: {},
  getSS_: () => ({ getSheets: () => sheets }),
  getMembersList: () => MEMBERS,
  normName_: (s) => String(s == null ? '' : s).replace(/[\s　]/g, ''),
  parseDate_(v) {
    if (!v) return null;
    let d;
    if (Object.prototype.toString.call(v) === '[object Date]') d = new Date(v.getTime());
    else {
      const s = String(v).trim().replace(/[年月]/g, '/').replace(/日/g, '');
      if (!s) return null;
      d = new Date(s);
    }
    if (isNaN(d.getTime())) return null;
    d.setHours(0, 0, 0, 0);
    return d;
  },
  fmtDate_(d) {
    if (!d) return '';
    const p = (n) => ('0' + n).slice(-2);
    return d.getFullYear() + '/' + p(d.getMonth() + 1) + '/' + p(d.getDate());
  },
  // コード.js の照合をそのまま持ってくる（姓だけでも当たるように）
  matchInviterToMember(inviterName, list) {
    if (!inviterName) return '';
    const s = String(inviterName).replace(/[\s　さん]/g, '');
    if (!s) return '';
    for (const m of list) {
      const t = m.name.replace(/[\s　]/g, '');
      if (t === s || t.startsWith(s) || s.startsWith(t)) return m.name;
    }
    return String(inviterName);
  },
};
sandbox.global = sandbox;
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'routine_srv.js'), 'utf8'), sandbox, { filename: 'routine_srv.js' });

// --- 全開催日を通して読んでみる ---
let dates = [];
for (const name of Object.keys(DATA)) {
  const row1 = DATA[name][0] || [];
  row1.forEach((v) => { if (/^\d{4}\/\d{2}\/\d{2}$/.test(String(v))) dates.push(String(v)); });
}
dates = [...new Set(dates)].sort();

let ok = 0, noSheet = 0, withCore = 0, withPres = 0, unmatched = [], badCore = [];
for (const d of dates) {
  const r = sandbox.getRoutineInfo(d);
  if (!r.ok) { console.log('  読み取り失敗', d, r.message); continue; }
  if (!r.found) { noSheet++; continue; }
  ok++;
  if (r.coreValue) withCore++;
  else if (r.coreValueRaw && !/^(済|休会|なし|無し)/.test(r.coreValueRaw)) badCore.push([d, r.coreValueRaw]);
  if (r.longPresenter) withPres++;
  if (r.longPresenterUnmatched) unmatched.push([d, r.longPresenterRaw]);
}
console.log(`開催日 ${dates.length} 件 → 列が見つかった ${ok} 件・見つからない ${noSheet} 件`);
console.log(`  コアバリューを判別できた: ${withCore} 件`);
console.log(`  スタートアッププレゼンの方あり : ${withPres} 件`);
if (badCore.length) {
  console.log(`  判別できなかったコアバリュー ${badCore.length} 件:`);
  badCore.slice(0, 8).forEach(([d, v]) => console.log(`    ${d}  ${JSON.stringify(v).slice(0, 60)}`));
}
if (unmatched.length) {
  console.log(`  名簿と一致しなかった方 ${unmatched.length} 件（検証用の名簿には少ししか入れていないため）:`);
  unmatched.slice(0, 8).forEach(([d, v]) => console.log(`    ${d}  ${v}`));
}

// --- スライドに載せる項目 ---
console.log('\nスライドに載せる項目:');
for (const d of ['2026/04/01', '2026/07/29', '2026/09/23', '2026/09/30']) {
  const r = sandbox.getRoutineInfo(d);
  if (!r.found) { console.log(`  ${d} (該当なし)`); continue; }
  const mp = (r.mainPresenters || []).map((m) => `${m.raw}→${m.name || '(未一致)'}`).join(' / ');
  console.log(`  ${d}`);
  console.log(`     メインプレゼン : ${mp || '(なし)'}`);
  console.log(`     一般規定       : ${r.generalPolicy || '(なし)'}番  ［${r.generalPolicyRaw}］`);
  console.log(`     求める専門分野 : ${JSON.stringify(r.wantedCategories)}`);
  console.log(`     開放カテゴリー : ${r.openCategory || '(空欄)'}`);
  console.log(`     審査中         : ${r.reviewCategory || '(空欄)'}`);
  const WHEN = { during: '', after: '[アフター]', later: '[翌週以降]' };
  const rc = (r.recommendations || []).map((x) =>
    `${WHEN[x.when] || ''}${x.giver.name || x.giver.raw || '?'}→${x.receiver.name || x.receiver.raw || '?'}`).join(' / ');
  console.log(`     推薦のことば   : ${rc || '(なし)'}`);
  console.log(`     リージョン参加者: ${r.regionGuestsRaw || '(空欄)'}`);
}

// --- リージョン参加者（アンバサダー・ディレクターのページの初期値に使う）---
const region = dates.map((d) => [d, (sandbox.getRoutineInfo(d) || {}).regionGuestsRaw || ''])
  .filter(([, v]) => v && !/^(なし|無し)/.test(v));
console.log(`\nリージョン参加者の記載がある日: ${region.length} 件`);
region.slice(0, 10).forEach(([d, v]) => console.log(`  ${d}  ${v}`));

// --- 決まった日で中身を確かめる ---
const samples = ['2026/04/01', '2026/04/08', '2026/09/30', '2026/10/07'];
console.log('\n抜き取り:');
for (const d of samples) {
  const r = sandbox.getRoutineInfo(d);
  console.log(`  ${d}  ${r.found ? r.sheetName : '(該当なし)'}`
    + (r.found ? `  第${r.meetingNo}回  コアバリュー=${r.coreValue || '(不明)'} `
               + `${r.coreValueRaw ? '［' + String(r.coreValueRaw).replace(/\s+/g, ' ').slice(0, 30) + '］' : ''}`
               + `  スタートアップ=${r.longPresenter || (r.longPresenterRaw ? r.longPresenterRaw + '(未一致)' : 'なし')}` : ''));
}

// --- 推薦の言葉の読み取り（決まった書き方で確かめる）---
const fails = [];
let checks = 0;
const ck = (c, m) => { checks++; if (!c) fails.push(m); };
const recoOf = (t) => sandbox.routineRecommendations_(t)
  .map((x) => `${x.when}:${x.giver.raw}>${x.receiver.raw}`).join(' | ');
[
  // 見出しで時間を分ける。丸数字の種類（⓵と①）が混ざっていてもよい
  ['定例会中\n⓵藤本さん⇒金井さん、\n②山岸さん⇒山内さん\nアフター\n①船越さん→金井さん',
   'during:藤本さん>金井さん | during:山岸さん>山内さん | after:船越さん>金井さん'],
  // 1行に何組も。翌週以降は later
  ['定例会中 ①Aさん→Bさん ②Cさん→Dさん\nアフター 船越さん→金井さん 藤本さん→長尾さん\n翌週以降\n岩淵→辻',
   'during:Aさん>Bさん | during:Cさん>Dさん | after:船越さん>金井さん | after:藤本さん>長尾さん | later:岩淵>辻'],
  // 見出しなし（定例会中として扱う）。括弧の但し書きは落とす
  ['大森さん →川邉さん\n溝川さん→大森さん（先週繰り越し分）', 'during:大森さん>川邉さん | during:溝川さん>大森さん'],
  // 見出しと組が同じ行。※以降のメモは落とす
  ['定例会後：北川さん→舩川さん ※時間があれば', 'after:北川さん>舩川さん'],
  ['アフター分　成田→中西さん', 'after:成田>中西さん'],
  ['定例会中なし\nアフター\n成田→中西さん', 'after:成田>中西さん'],
  // 【】で囲んだ見出しと組が同じ行
  ['【定例会中】泉さん→梅中さん\n【アフター】山内さん→梅中さん、船越さん→梅中さん',
   'during:泉さん>梅中さん | after:山内さん>梅中さん | after:船越さん>梅中さん'],
  // 組のうしろの括弧に時間（その組だけ）
  ['加納さん→近江さん（定例会中）\n高梨さん→泉さん（アフター）\n泉さん→高梨さん（※5/7よりスライド、アフター）\n原口さん→泉さん',
   'during:加納さん>近江さん | after:高梨さん>泉さん | after:泉さん>高梨さん | during:原口さん>泉さん'],
  // 行頭の「・」
  ['【定例会中】※3枚読み上げます！\n・船越さん→小早川さん\n・山岸さん→文野さん、・山内さん→仲里根さん',
   'during:船越さん>小早川さん | during:山岸さん>文野さん | during:山内さん>仲里根さん'],
  // 書き方の見本は組にしない
  ['○○さん　⇒○○さん\n○○さん　⇒○○さん', ''],
  // 翌週以降の欄の「岩淵（辻、船越）」のような予定の書き方は組にしない
  ['定例会中\n①北川さん→徳永さん\n\nアフターは、割振り表次第で決定。\n①長尾くん→徳永さん\n\n翌週以降\n岩淵（辻、船越、梅中、徳永）',
   'during:北川さん>徳永さん | after:長尾くん>徳永さん'],
  ['なし', ''],
  ['', ''],
].forEach(([t, want]) => ck(recoOf(t) === want, `推薦の言葉「${t.replace(/\n/g, '／')}」→ ${recoOf(t)}（${want} のはず）`));

// 名簿の氏名に合わせる：字の違い（川辺／川邉・岩淵／岩渕）でも、名字が1人に決まれば合わせる
const nameOf = (raw) => sandbox.routineMemberName_(raw).name;
ck(nameOf('川辺さん') === '川邉 真由子', '「川辺さん」→ ' + nameOf('川辺さん'));
ck(nameOf('岩淵') === '岩渕 裕太', '「岩淵」→ ' + nameOf('岩淵'));
ck(nameOf('桑田さん') === '桒田 美香', '「桑田さん」→ ' + nameOf('桑田さん'));
ck(nameOf('西沢さん') === '西澤 浩平', '「西沢さん」→ ' + nameOf('西沢さん'));
// 「山内さんへ」「若松くん」のような書き方
ck(nameOf('山内さんへ') === '山内 登志夫', '「山内さんへ」→ ' + nameOf('山内さんへ'));
ck(nameOf('若松くん') === '若松 勇人', '「若松くん」→ ' + nameOf('若松くん'));
// 名字が2人に当たるとき（斉藤 → 齋藤・齊藤）は決めない
ck(nameOf('斉藤さん') === '', '「斉藤さん」が1人に決まってしまった: ' + nameOf('斉藤さん'));
ck(nameOf('大森さん') === '', '名簿に無い「大森さん」が誰かに決まってしまった: ' + nameOf('大森さん'));

// --- 新入会・更新式（1人ずつに分け、更新は1年・2年も読む。書き方の見本は作り物）---
const listOf = (t) => sandbox.routineMemberList_(t)
  .map((x) => (x.name || '×' + x.raw) + (x.years ? '/' + x.years : '') + (x.category ? '[' + x.category + ']' : '')).join(' | ');
[
  ['本間さん、平松さん、谷口さん', '本間 圭吾 | 平松 成子 | 谷口 大輔'],
  ['梅田さん（1年更新）、文野さん（2年）', '梅田 明日斗/1 | 文野 雅彦/2'],
  ['文野さん（2年更新）北川さん（1年更新）', '文野 雅彦/2 | 北川 光男/1'],
  // 前に業種・カテゴリーが付く書き方（うしろが名字に一致する方）
  ['内装業　店舗オフィス岩渕さん（1年）', '岩渕 裕太/1'],
  ['法人コスト削減：上原さん', '上原 誠'],
  ['カテゴリー測量、細川さん（1年）　税理士（事業承継）文野さん（2年）', '細川 哲也/1 | 文野 雅彦/2'],
  // 括弧の中のお名前・「○○さんのあと」のような文の中のお名前は数えない
  ['小西さん（泉さんは10/7）', '小西 美和'],
  ['上原誠さん（エステサロン）\n若松さんのあと、31番に入ります', '上原 誠'],
  // よみがなの括弧のあとに「さん」／どこにも「さん」が無い書き方
  ['北川光男（きたがわみつお）さん', '北川 光男'],
  ['北川 光男（きたがわ みつお）（新規事業ITサポート）', '北川 光男'],
  // 名簿に無い方（まだ名簿に入っていない新メンバー）は、書いてあったお名前とカテゴリー
  ['（業務用冷凍冷蔵設備) 新井 花子さん', '×新井 花子さん[業務用冷凍冷蔵設備]'],
  ['新井 花子さん（エステサロン）メンバーリスト20番', '×新井 花子さん[エステサロン]'],
  // 同じ名字が2人（金井）は決めない
  ['金井さん（2年更新）', '×金井さん/2'],
  ['なし（9/28月曜8時時点）', ''], ['対面BOD', ''], ['休会', ''], ['', ''],
].forEach(([t, want]) => ck(listOf(t) === want, `新入会・更新式「${t.replace(/\n/g, '／')}」→ ${listOf(t)}（${want} のはず）`));

// --- バイスプレジデントによる報告（読み上げる文から、スライドの数字を拾う）---
const J = (x) => JSON.stringify(x);
{
  const vp = sandbox.routineVpReport_('チャプター設立以来、月間リファーラル数の平均は308件、2026年8月の月間リファーラル数は281件(70件/週） '
    + '2026年3月から2026年8月の半年間のリファーラル数の合計はちょうど1,927件となります。 （クリック）チャプターが発足されてから交わされた'
    + 'ビジネスのサンキュー額つまり売上は、 54億0,272万円となります。（このあとスライド通り速報を読み上げる）');
  ck(J(vp) === J({ avg: '308', month: '2026-08', count: '281', perWeek: '70', from: '2026-03', to: '2026-08', total: '1,927', thanks: '54億272万円' }),
     'バイスプレジデントによる報告: ' + J(vp));
  ck(sandbox.routineVpReport_('ー') === null && sandbox.routineVpReport_('') === null, 'バイスプレジデントによる報告の空欄');
}
// --- ネットワーキングリーダー（原稿から、部門ごとの数と受賞者。同数の2人・「わたくし」・該当者なし）---
{
  const nl = sandbox.routineNetworkingLeaders_('2026年8月のネットワーキングリーダーの発表をさせて頂きます。\nまずCEUからです。\n'
    + 'CEUとは、BNIのトレーニングやワークショップなどを通じて学んだ、エデュケーションをポイント化したものです。\n'
    + '先月のCEU部門は23ポイントで、エンターテイメントショー：泉さんです。\n\n次にサンキュー部門、つまり売上に貢献されたメンバーを発表します。\n'
    + 'なんと！1,200,6000円の売上に貢献頂きました、税理士（事業承継）：文野さんです！\n\n次に外部リファーラル部門。\n16件で、わたくし、ベリーダンス：山岸です。\n\n'
    + '次に1to1の回数ですが、今回はおふたりいらっしゃいます。\n21回で、鍼灸師：梅中さんと、デジタル広告制作（ホームページ・動画）：若松さんです！\n\n'
    + '最後にビジター招待数部門です。\n3名以上招待された方がいらっしゃらなかったので、該当者なしです。\n\n'
    + 'それでは、今回の受賞者を代表し、泉さんより、一言頂きたいと思います。\n素晴らしい貢献を頂きました、泉さん、文野さんへ、改めて大きな拍手をお願い致します。');
  const got = nl ? nl.month + ' ' + nl.items.map((it) => `${it.key}=${it.value}${it.unit}:${it.winners.map((w) => (w.self ? '*' : '') + (w.name || '×' + w.raw)).join('+')}`).join(' / ') : 'null';
  ck(got === '2026-08 ceu=23ポイント:泉 ゆか / thanks=12,006,000円:文野 雅彦 / ext=16件:*山岸 由紀子 / oto=21回:梅中 公+若松 勇人 / visitor=3名:',
     'ネットワーキングリーダー: ' + got);
  ck(sandbox.routineNetworkingLeaders_('ー') === null, 'ネットワーキングリーダーの「ー」');
}
// --- その月の最初の定例会か（休会日の週は飛ばす）---
sandbox.getHolidays = () => ['2026/05/06'];
[['2026-09-02', true], ['2026-09-09', false], ['2026-09-30', false], ['2026-05-13', true], ['2026-05-20', false]].forEach(([d, want]) => {
  const got = sandbox.routineFirstOfMonth_(new Date(d + 'T00:00:00'));
  ck(got === want, `月の最初の定例会 ${d} → ${got}（${want} のはず）`);
});

console.log(`\n推薦の言葉・氏名の照らし合わせ・新入会／更新式・バイス報告・ネットワーキングリーダー: 検査 ${checks} 件`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: 見出し（定例会中・アフター・翌週以降）での分け方・1行に何組も・字の違い・新入会／更新式・バイス報告・ネットワーキングリーダー');

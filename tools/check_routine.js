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
console.log(`  2分30秒プレゼンの方あり : ${withPres} 件`);
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
               + `  2分30秒=${r.longPresenter || (r.longPresenterRaw ? r.longPresenterRaw + '(未一致)' : 'なし')}` : ''));
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

console.log(`\n推薦の言葉・氏名の照らし合わせ: 検査 ${checks} 件`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: 見出し（定例会中・アフター・翌週以降）での分け方・1行に何組も・字の違い');

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
const MEMBERS = ['谷村 大輔', '竹田 明日翔', '本田 圭吾', '竹中 公', '桒原 美穂',
                 '山口 由美子', '田中 浩子', '藤田 礼恵', '金子 美緒', '細田 哲雄',
                 '野崎 佳子', '豊田 恵', '高瀬 舟', '平井 成美', '舩山 ちひろ', '杉浦 太郎',
                 '山本 登一郎', '金子 高志', '瞳 ゆり', '田中 秀一', '小池 美咲',
                 '木村 光範', '若林 勇貴', '石渕 裕介', '分銅 雅一', '三澤 浩三', '上野 誠',
                 '渡邉 真理子', '齋藤 一郎', '齊藤 次郎']
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
  ['定例会中\n⓵藤田さん⇒金子さん、\n②山口さん⇒山本さん\nアフター\n①船木さん→金子さん',
   'during:藤田さん>金子さん | during:山口さん>山本さん | after:船木さん>金子さん'],
  // 1行に何組も。翌週以降は later
  ['定例会中 ①Aさん→Bさん ②Cさん→Dさん\nアフター 船木さん→金子さん 藤田さん→長見さん\n翌週以降\n石淵→尹',
   'during:Aさん>Bさん | during:Cさん>Dさん | after:船木さん>金子さん | after:藤田さん>長見さん | later:石淵>尹'],
  // 見出しなし（定例会中として扱う）。括弧の但し書きは落とす
  ['大野さん →渡邉さん\n溝口さん→大野さん（先週繰り越し分）', 'during:大野さん>渡邉さん | during:溝口さん>大野さん'],
  // 見出しと組が同じ行。※以降のメモは落とす
  ['定例会後：木村さん→舩山さん ※時間があれば', 'after:木村さん>舩山さん'],
  ['アフター分　成毛→中込さん', 'after:成毛>中込さん'],
  ['定例会中なし\nアフター\n成毛→中込さん', 'after:成毛>中込さん'],
  // 【】で囲んだ見出しと組が同じ行
  ['【定例会中】瞳さん→竹中さん\n【アフター】山本さん→竹中さん、船木さん→竹中さん',
   'during:瞳さん>竹中さん | after:山本さん>竹中さん | after:船木さん>竹中さん'],
  // 組のうしろの括弧に時間（その組だけ）
  ['佐藤さん→近藤さん（定例会中）\n高木さん→瞳さん（アフター）\n瞳さん→高木さん（※5/7よりスライド、アフター）\n原田さん→瞳さん',
   'during:佐藤さん>近藤さん | after:高木さん>瞳さん | after:瞳さん>高木さん | during:原田さん>瞳さん'],
  // 行頭の「・」
  ['【定例会中】※3枚読み上げます！\n・船木さん→小石川さん\n・山口さん→分銅さん、・山本さん→仲宗根さん',
   'during:船木さん>小石川さん | during:山口さん>分銅さん | during:山本さん>仲宗根さん'],
  // 書き方の見本は組にしない
  ['○○さん　⇒○○さん\n○○さん　⇒○○さん', ''],
  // 翌週以降の欄の「石淵（尹、船木）」のような予定の書き方は組にしない
  ['定例会中\n①木村さん→徳山さん\n\nアフターは、割振り表次第で決定。\n①長見くん→徳山さん\n\n翌週以降\n石淵（尹、船木、竹中、徳山）',
   'during:木村さん>徳山さん | after:長見くん>徳山さん'],
  ['なし', ''],
  ['', ''],
].forEach(([t, want]) => ck(recoOf(t) === want, `推薦の言葉「${t.replace(/\n/g, '／')}」→ ${recoOf(t)}（${want} のはず）`));

// 名簿の氏名に合わせる：字の違い（渡辺／渡邉・石淵／石渕）でも、名字が1人に決まれば合わせる
const nameOf = (raw) => sandbox.routineMemberName_(raw).name;
ck(nameOf('渡辺さん') === '渡邉 真理子', '「渡辺さん」→ ' + nameOf('渡辺さん'));
ck(nameOf('石淵') === '石渕 裕介', '「石淵」→ ' + nameOf('石淵'));
ck(nameOf('桑原さん') === '桒原 美穂', '「桑原さん」→ ' + nameOf('桑原さん'));
ck(nameOf('三沢さん') === '三澤 浩三', '「三沢さん」→ ' + nameOf('三沢さん'));
// 「山本さんへ」「若林くん」のような書き方
ck(nameOf('山本さんへ') === '山本 登一郎', '「山本さんへ」→ ' + nameOf('山本さんへ'));
ck(nameOf('若林くん') === '若林 勇貴', '「若林くん」→ ' + nameOf('若林くん'));
// 名字が2人に当たるとき（斉藤 → 齋藤・齊藤）は決めない
ck(nameOf('斉藤さん') === '', '「斉藤さん」が1人に決まってしまった: ' + nameOf('斉藤さん'));
ck(nameOf('大野さん') === '', '名簿に無い「大野さん」が誰かに決まってしまった: ' + nameOf('大野さん'));

console.log(`\n推薦の言葉・氏名の照らし合わせ: 検査 ${checks} 件`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: 見出し（定例会中・アフター・翌週以降）での分け方・1行に何組も・字の違い');

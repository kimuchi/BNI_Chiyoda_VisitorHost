// ウィークリープレゼンの「始まりの業種区分」を、ルーティンチェックシートから決める処理を確かめる。
// 本番と同じ routine_srv.js / member_presen_srv.js を、シートまわりだけ差し替えてNodeで動かす。
//
//   node tools/check_weekly_start.js <routine.json> <members.json>
//
//   routine.json … tools/routine_dump.py で書き出したルーティンチェックシート
//   members.json … メンバー名簿（{ name, cat } の配列。区分の人数に使う）
//
// 確かめること
//   ・「ウィークリープレゼン」の記載が、どれも業種区分として読めるか
//   ・記載が無い日の繰り上げ（開催した回ごとに1つ進み、誰もいない区分は飛ばす）が、
//     実際の次の記載と合うか（名簿が今と同じ時期＝最近の分だけを判定に使う）

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
const ROUTINE = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const MEMBERS = JSON.parse(fs.readFileSync(process.argv[3], 'utf8'));
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }

function fakeSheet(name, grid) {
  return {
    getName: () => name,
    getLastRow: () => grid.length,
    getLastColumn: () => (grid[0] ? grid[0].length : 0),
    getRange: (r, c, nr, nc) => ({
      getValues: () => Array.from({ length: nr }, (_, i) => (grid[r - 1 + i] || []).slice(c - 1, c - 1 + nc)),
    }),
  };
}
const sheets = Object.keys(ROUTINE).map((n) => fakeSheet(n, ROUTINE[n]));

// 開催日の一覧（1行目）。本番は「休会日」シートを使うが、ここではルーティンチェックシートから
// 見当をつける：水曜なのに列が無い週と、列はあっても2行目（定例会回数）が空の週を休会日とみなす
// （2026/5/6・8/12 は列だけあって回数が空）。
let dates = [];
const held = {};
for (const n of Object.keys(ROUTINE)) {
  const g = ROUTINE[n];
  (g[0] || []).forEach((v, c) => {
    if (!/^\d{4}\/\d{2}\/\d{2}$/.test(String(v))) return;
    dates.push(String(v));
    if (String((g[1] || [])[c] || '').trim()) held[String(v)] = true;
  });
}
dates = [...new Set(dates)].sort();
const fmt = (d) => d.getFullYear() + '/' + ('0' + (d.getMonth() + 1)).slice(-2) + '/' + ('0' + d.getDate()).slice(-2);
const HOLIDAYS = [];
for (let d = new Date(dates[0] + ' 00:00:00'); fmt(d) <= dates[dates.length - 1]; d.setDate(d.getDate() + 7)) {
  if (!held[fmt(d)]) HOLIDAYS.push(fmt(d));
}

const sandbox = {
  console,
  getSS_: () => ({ getSheets: () => sheets }),
  getMembersList: () => MEMBERS.map((m) => ({ no: m.no, name: m.name })),
  getHolidays: () => HOLIDAYS,
  normName_: (s) => String(s == null ? '' : s).replace(/[\s　]/g, ''),
  matchInviterToMember(inviterName, list) {
    const s = String(inviterName || '').replace(/[\s　さん]/g, '');
    if (!s) return '';
    for (const m of list) {
      const t = m.name.replace(/[\s　]/g, '');
      if (t === s || t.startsWith(s) || s.startsWith(t)) return m.name;
    }
    return String(inviterName);
  },
  parseDate_(v) {
    if (!v) return null;
    const d = new Date(String(v).trim().replace(/[年月]/g, '/').replace(/日/g, ''));
    if (isNaN(d.getTime())) return null;
    d.setHours(0, 0, 0, 0);
    return d;
  },
  fmtDate_: (d) => (d ? fmt(d) : ''),
};
vm.createContext(sandbox);
for (const f of ['member_master_srv.js', 'routine_srv.js', 'member_presen_srv.js']) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), sandbox, { filename: f });
}
// 業種区分マスタは初期値（member_master_srv.js の DEFAULT_CATEGORIES_）を使う
sandbox.getCategoryMaster = () => vm.runInContext('DEFAULT_CATEGORIES_', sandbox)
  .map((r) => ({ key: r[0], label: r[1], block: r[4], order: r[5] }));

const members = MEMBERS.map((m) => ({ name: m.name, cat: m.cat, blockKey: '' }));
const cycle = sandbox.mpBlocks_(members);
const blockName = (k) => (cycle.find((c) => c.gkey === k) || {}).block || '(なし)';
console.log('区分と人数: ' + cycle.map((c) => `${c.block} ${c.count}名`).join(' / '));

// 1) 記載がすべて区分として読めるか
const rows = sandbox.routineRowValues_(vm.runInContext('ROUTINE_WEEKLY_LABELS_', sandbox));
const keys = Object.keys(rows).sort();
const parsed = {};
const unread = [];
for (const d of keys) {
  const k = sandbox.mpBlockFromText_(cycle, members, rows[d]);
  if (k) parsed[d] = k; else unread.push(d + '  ' + rows[d]);
}
// 昔の記載には、今は無い区分名（研修＆育成・建設＆住まい）や、番号とお名前だけのものがある。
// 判定に使うのは最近1年ぶんだけにして、それより前は件数だけ出す。
const lastKey = keys[keys.length - 1] || '';
const yearAgo = fmt(new Date(new Date(lastKey + ' 00:00:00').getTime() - 365 * 86400000));
const recentUnread = unread.filter((x) => x.slice(0, 10) >= yearAgo);
console.log(`\n「ウィークリープレゼン」の記載 ${keys.length} 件 → 区分が読めた ${Object.keys(parsed).length} 件`
  + `（読めない ${unread.length} 件のうち、最近1年ぶん ${recentUnread.length} 件）`);
recentUnread.forEach((x) => console.log('  読めない: ' + x));
ck(keys.length > 0, '「ウィークリープレゼン」の記載が1件も見つからない');
ck(recentUnread.length === 0, `最近1年の記載に、区分として読めないものがある（${recentUnread.length}件）`);

// 2) 前回の記載から繰り上げた区分が、実際の次の記載と合うか。
//    名簿は今のものなので、判定は最近（今の名簿と同じ区分構成の時期）の分だけ。
const RECENT = keys.slice(-9);          // お盆（8/12 休会）をまたぐ分も入れる
let agree = 0, total = 0;
console.log('\n最近の記載と、前回の記載から繰り上げた区分:');
for (let i = 1; i < RECENT.length; i++) {
  const prev = RECENT[i - 1], cur = RECENT[i];
  if (!parsed[prev] || !parsed[cur]) continue;
  const steps = sandbox.mpMeetingsBetween_(sandbox.parseDate_(prev), sandbox.parseDate_(cur), HOLIDAYS);
  const guess = sandbox.mpAdvance_(cycle, parsed[prev], steps);
  const actual = sandbox.mpAdvance_(cycle, parsed[cur], 0);
  total++;
  if (guess === actual) agree++;
  console.log(`  ${cur}  記載=${blockName(actual)}  繰り上げ=${blockName(guess)}（${prev} から ${steps}回）${guess === actual ? '' : '  ← 違う'}`);
}
ck(total > 0 && agree === total, `繰り上げが実際の記載と合わない（${agree}/${total}）`);

// 3) 画面に渡す値（その日の記載・前回からの繰り上げ）
const last = keys[keys.length - 1];
const next = fmt(new Date(new Date(last + ' 00:00:00').getTime() + 7 * 86400000));
const onDay = sandbox.mpStartFromRoutine_(cycle, members, rows, last, HOLIDAYS);
ck(onDay && onDay.from === 'routine' && onDay.key === sandbox.mpAdvance_(cycle, parsed[last], 0),
   `記載のある日（${last}）がその記載どおりにならない: ${JSON.stringify(onDay)}`);
const after = sandbox.mpStartFromRoutine_(cycle, members, rows, next, HOLIDAYS);
ck(after && after.from === 'previous' && after.date === last, `記載の無い日（${next}）が前回から数えられていない: ${JSON.stringify(after)}`);
console.log(`\n${last}（記載あり）→ ${blockName(onDay && onDay.key)}`);
console.log(`${next}（記載なし）→ ${blockName(after && after.key)}（${last} の「${rows[last]}」から${after && after.steps}回ぶん）`);

// 4) 誰もいない区分の飛ばし方
const A = [{ gkey: 'a', count: 2 }, { gkey: 'b', count: 0 }, { gkey: 'c', count: 3 }, { gkey: 'd', count: 1 }];
ck(sandbox.mpAdvance_(A, 'a', 1) === 'c', '誰もいない区分を飛ばさない');
ck(sandbox.mpAdvance_(A, 'b', 0) === 'c', '誰もいない区分から始めると、次の区分にならない');
ck(sandbox.mpAdvance_(A, 'b', 3) === 'a', '誰もいない区分からの繰り上げがずれる');
ck(sandbox.mpAdvance_(A, 'd', 1) === 'a', '最後の区分から先頭に戻らない');
ck(sandbox.mpAdvance_(A, 'c', 6) === 'c', '2周ぶんの繰り上げで元に戻らない');

// 5) 画面（メンバープレゼン）まで通す：サーバーの getMemberPresenContext() の結果で画面を開く
sandbox.getMemberMaster = () => ({ members: MEMBERS });
sandbox.findPhotoIdForName_ = () => 'photo';
sandbox.getBigTemplateStatus = () => ({ templates: [{ kind: 'memberPresen', registered: true }] });
sandbox.getMeetingCandidates = () => [last, next].map((d) => ({ dateValue: d, display: d }));
const ctx = sandbox.getMemberPresenContext();
ck(ctx.ok, 'getMemberPresenContext が失敗した: ' + ctx.message);
const cands = ctx.candidates || [];
ck(cands[0] && cands[0].startFrom === 'routine' && cands[1] && cands[1].startFrom === 'previous',
   '候補の出どころ: ' + JSON.stringify(cands.map((c) => c.startFrom)));
const { loadPage } = require('./lib_minidom');
const page = loadPage('member_presen.html', { fails, server: { getMemberPresenContext: () => JSON.parse(JSON.stringify(ctx)) } });
page.step('メンバープレゼンの画面を開く', () => page.window.onload());
ck(page.els.start.value === cands[0].start, `画面の始まりの区分が ${page.els.start.value}（${cands[0].start} のはず）`);
ck(/ルーティンチェックシートの記載/.test(page.els.startNote.textContent), '画面に出どころが出ていない: ' + page.els.startNote.textContent);
page.step('次の開催日に切り替える', () => { page.els.date.value = '1'; page.run('onDate()'); });
ck(page.els.start.value === cands[1].start, `次の開催日の始まりの区分が ${page.els.start.value}（${cands[1].start} のはず）`);
ck(/前回の記載/.test(page.els.startNote.textContent), '前回からの繰り上げが画面に出ていない: ' + page.els.startNote.textContent);
console.log(`画面: ${last} → ${page.els.startNote.textContent ? '「' + blockName(cands[0].start) + '」' : ''}／${next} → 「${blockName(cands[1].start)}」`);

console.log(`\n始まりの業種区分: 検査 ${checks} 件`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: 記載の読み取り・前回からの繰り上げ・誰もいない区分の飛ばし方');

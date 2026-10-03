// 欠席の方（ルーティンチェックシートの「代理・欠席」の「欠席」「医療欠席」の欄）の、
// ウィークリープレゼン（前半のメンバーのページ）・リファーラル発表（後半）のページを作らないことを、実データなしで確かめる。
// サーバーは本番の *.js を見せかけのスプレッドシート（lib_sheet_fake.js）の上で動かし、画面は簡易DOM（lib_minidom.js）で動かす。
// 名前はすべて架空。
//
//   node tools/check_absent_skip.js
//
// 確かめること
//   ・欄の読み方：かっこの中（当欠など）・「遅刻：」・「○○さん遅刻」・「代理：」「→」・「復帰」・「なし」・「○時時点でなし」・
//     敬称ごとに1人・うしろの書き足し。名簿の方に合わせる（同じ名字の方が2人なら決めず、画面で選んでもらう）
//   ・代理を立てた方（「代理」の欄）は外さない（代理の方が発表する）。医療欠席の方は外す
//   ・前半：欠席の方の個人ページを作らない。扉ページの表・先頭の写真・NEXT にも出さない。
//     チェックを外せば作る・「欠席の方を足す」で選べば外す。その日の列が無い日は、前の日の欠席の方を残さない
//   ・後半：欠席の方のリファーラル発表のページを作らない。「次の発表者」も出席の方

process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeEnv } = require('./lib_sheet_fake');
const { loadPage } = require('./lib_minidom');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }
const J = (x) => JSON.stringify(x);

// ---- 架空の名簿（見本さんは2人いる）・ルーティンチェックシート ----
const HEAD = ['No', '業種区分', '氏名', 'ふりがな', 'カテゴリー', '会社名', '役職', 'メモ', '写真ファイル名', '一言コメント',
  '紹介してほしい人', '協業したい人', '入会日', '更新日', '更新期限日', '会社での役職'];
const NAMES = ['見本 一郎', '試験 花子', '架空 三郎', '仮名 四郎', '例示 五月', '模擬 六助', '空想 七海', '見本 みさき'];
const CATS = ['企業サポート', '不動産関連'];
const rosterRows = [HEAD].concat(NAMES.map((n, i) => HEAD.map((h) => (h === 'No' ? String(i + 1) : h === '氏名' ? n
  : h === '業種区分' ? CATS[i % 2] : h === 'カテゴリー' ? 'カテゴリー' + (i + 1) : h === '会社名' ? '見本会社' + (i + 1) : ''))));
const DATE = '2026/10/07', NEXT = '2026/10/14';
const row = (c, d, vals) => ['', '', c, d, '', d || c ? 'バイス' : '', '', '', '', (vals || [])[0] || '', '', (vals || [])[1] || ''];
const ROUTINE = [
  ['', '', '開催日', '', '', '', '', '', '', DATE, '', NEXT],
  ['', '', '定例会回数', '', '', '', '', '', '', '536', '', '537'],
  ['', 'No', '内容', '', '', '担当', '期日', '曜日目安', '備考', '', '', ''],
  ['', '1', '遅刻・欠席担当(7:00開始)', '', '', 'バイス', '', '', '', '見本の担当', '', ''],
  ['', '2', '代理・欠席', '', '', '', '', '', '', '', '', ''],
  row('', '代理', ['例示さん代理：外部 太郎さん']),
  row('', '欠席', ['試験さん（当欠）、架空 三郎さん／遅刻：模擬さん', 'なし（10/13 8時時点）']),
  row('', '医療欠席', ['仮名さん']),
];
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
], { BNI_CHAPTER: J({ name: '見本', termBase: 23, meetingBaseDate: '2026/03/18', meetingBaseCount: 509 }) });

// ===== 1) 欄の読み方（routineAbsentTokens_）=====
[
  ['試験さん、架空さん', ['試験', '架空']],
  ['試験さん（当欠）／遅刻：仮名さん', ['試験']],
  ['なし', []], ['ー', []], ['', []], ['24日21時時点でなし', []], ['なし（10/1火曜8時時点）', []], ['休会', []],
  ['例示さん代理：外部 太郎さん', ['例示']],
  ['模擬さん→外部 次郎さん(8/20時点)', ['模擬']],
  ['空想さん復帰につきプレゼン・リファーラルスライド変更', []],
  ['試験 花子さん　架空さん', ['試験 花子', '架空']],
  ['仮名さん休会中', ['仮名']],
  ['※模擬さん遅刻', []],
  ['当欠：空想さん', ['空想']],
  ['見本さん、試験さん、架空さん(当欠)、※仮名さん遅刻', ['見本', '試験', '架空']],
  ['例示ED', ['例示ED']],
].forEach(([raw, want]) => {
  const got = F.routineAbsentTokens_(raw);
  ck(J(got) === J(want), `1) 「${raw}」→ ${J(got)}（${J(want)} のはず）`);
});
// 名簿の方に合わせる（同じ名字の方が2人＝見本さんは決めない）。敬称のある書き方が混ざっていれば、敬称の無いもの（「例示ED」）は読まない
{
  const r = F.routineAbsentees_('試験さん（当欠）、見本さん、例示ED', '仮名さん');
  ck(J(r.names) === J(['試験 花子', '仮名 四郎']) && J(r.unknown) === J(['見本']),
     '1) 名簿の方に合わせる: ' + J(r));
  ck(r.raw === '欠席：試験さん（当欠）、見本さん、例示ED　／　医療欠席：仮名さん', '1) 欄の文字: ' + J(r.raw));
  const r2 = F.routineAbsentees_('空想 七海、例示ED', '');
  ck(J(r2.names) === J(['空想 七海']) && J(r2.unknown) === J(['例示ED']), '1) 敬称の無い書き方: ' + J(r2));
}

// ===== 2) ルーティンチェックシートから（getRoutineInfo）=====
const R = F.getRoutineInfo(DATE, { firstHalf: true });
ck(R.ok && R.found && R.absentees, '2) getRoutineInfo に欠席の方が無い: ' + J({ ok: R.ok, found: R.found, message: R.message }));
const ABSENT = ['試験 花子', '架空 三郎', '仮名 四郎'];
ck(R.absentees && J(R.absentees.names) === J(ABSENT) && R.absentees.unknown.length === 0,
   '2) 欠席・医療欠席の方（遅刻の方・代理を立てた方は入れない）: ' + J(R.absentees));
const R2 = F.getRoutineInfo(NEXT);
ck(R2.absentees && R2.absentees.names.length === 0 && /^欠席：なし/.test(R2.absentees.raw), '2) 「なし」の日: ' + J(R2.absentees));

// ===== 3) 前半の画面：欠席の方のメンバーのページを作らない =====
const gk = (x) => String(x || '').normalize('NFKC').replace(/[\s・･＆&と]/g, '');
const MEMBERS = NAMES.map((n, i) => ({ no: String(i + 1), name: n, company: '見本会社' + (i + 1), title: 'カテゴリー' + (i + 1), hasPhoto: true }));
const mm = NAMES.map((n, i) => ({ name: n, company: '見本会社' + (i + 1), title: 'カテゴリー' + (i + 1), cat: CATS[i % 2], blockKey: gk(CATS[i % 2]), hasPhoto: true }));
const blocks = CATS.map((b, i) => ({ gkey: gk(b), block: b, order: i + 1, known: true, count: mm.filter((m) => m.blockKey === gk(b)).length }));
let routine = R;
const sent = [];
const SERVER = {
  getSystemVersion: () => 'test',
  getMeetingSlideContext: () => ({
    ok: true, meetings: [{ dateValue: DATE, display: '第536回 2026/10/07' }], defaultMeeting: { dateValue: DATE, display: '第536回 2026/10/07' },
    lists: { ok: true, text: { newMembers: '該当者なし', renewMembers: '該当者なし', d90: '該当者なし', d60: '該当者なし', d30: '該当者なし', overdue: '該当者なし' },
             newMembers: [], renewMembers: [], d90: [], d60: [], d30: [], overdue: [], noDate: [], done: [], leaving: [] },
    templates: { meetingFirst: true, meetingSecond: true, memberPresen: true }, routine,
    seconds: { weekly: 30, startup: 150, visitor: 20, referral: 7 }, coreValues: [], memberCount: MEMBERS.length, members: MEMBERS,
  }),
  getRoutineInfo: () => routine,
  getWeeklyGuests: () => ({ ok: true, guests: [] }),
  getMemberPresenContext: () => ({ ok: true, members: mm, blocks, rowsPerPage: 7, unusedCategories: [], seconds: { weekly: 30, startup: 150 },
    candidates: [{ dateValue: DATE, start: gk(CATS[0]), longPresenter: '', longPresenterRaw: '', startFrom: 'routine', startRaw: '企業サポート', startRawDate: DATE, startSteps: 0 }] }),
  getSpeakerRotationWeeks: () => ({ ok: true, header: '', notes: [], weeks: [] }),
  getRoleIntroPreview: () => ({ ok: true, term: 24, label: '', holdersOk: true, holders: [], teams: [] }),
  getMeetingTemplateInfo: () => ({ ok: true, message: '', list: [],
    referralBoxes: { companyTall: { x: 5087424, y: 2849608, cx: 6893161, cy: 1446550 }, categoryLow: 4071101, categoryWidth: 6893161, hasNext: true, slides: 1 } }),
  computeRenewalLists: () => SERVER.getMeetingSlideContext().lists,
  getMeetingMusicFiles: () => ({ ok: true, files: [] }),
  generateMeetingSlides: (...args) => { sent.push(args); return { ok: true, message: '作成しました', url: 'u' }; },
};
const lastOpts = () => (sent.length ? sent[sent.length - 1][3] || {} : {});
{
  const page = loadPage('slides_meeting_first.html', { server: SERVER, fails });
  const { els, window, run, step } = page;
  const html = (id) => String(els[id] ? els[id].innerHTML : '');
  page.flush();
  step('前半を開く', () => window.onload());
  const boxes = () => NAMES.map((n, i) => [els['ab_' + i] && els['ab_' + i].checked, n]);
  ck(ABSENT.every((n, i) => els['ab_' + i] && els['ab_' + i].checked && html('abBox').includes(n)),
     '3) 前半：欠席の方のチェックが出ていない: ' + html('abBox').replace(/<[^>]+>/g, ' ').slice(0, 300));
  ck(/欠席：試験さん（当欠）/.test(html('abBox')) && /代理を立てた方/.test(html('abBox')), '3) 前半：欄の記載・代理の方の案内が無い');
  ck(/欠席 2名/.test(html('mpOrder')) && /欠席の方 3名のページは作りません/.test(html('mpNote2')),
     '3) 前半：順番の表・枚数の案内に欠席の方が出ていない: ' + html('mpOrder').replace(/<[^>]+>/g, ' ') + ' / ' + html('mpNote2'));
  step('前半を作る', () => run('gen()'));
  const pages = () => lastOpts().memberPresen || [];
  const ind = () => pages().filter((x) => x.kind === 'individual');
  const ov = () => pages().filter((x) => x.kind === 'overview');
  ck(J(ind().map((x) => x.name)) === J(['見本 一郎', '例示 五月', '空想 七海', '模擬 六助', '見本 みさき']),
     '3) 前半：個人ページ（欠席の方を除く）: ' + J(ind().map((x) => x.name)));
  ck(J(ov().map((x) => [x.block, x.photoName, x.rows.map((r) => r.name)])) === J([['企業サポート', '見本 一郎', ['見本 一郎', '例示 五月', '空想 七海']],
     ['不動産関連', '模擬 六助', ['模擬 六助', '見本 みさき']]]), '3) 前半：扉ページの表・先頭の写真: ' + J(ov().map((x) => [x.block, x.photoName, x.rows.map((r) => r.name)])));
  ck(J(ind().map((x) => x.nextName)) === J(['例示 五月', '空想 七海', '', '見本 みさき', '']), '3) 前半：NEXT に欠席の方が出る: ' + J(ind().map((x) => x.nextName)));
  // チェックを外せば作る・足せば外す
  step('試験さんのチェックを外す', () => { els.ab_0.checked = false; run('absentToggle(0)'); run('gen()'); });
  ck(ind().some((x) => x.name === '試験 花子') && ind().find((x) => x.name === '試験 花子').nextName === '模擬 六助',
     '3) 前半：チェックを外しても作らない: ' + J(ind().map((x) => [x.name, x.nextName])));
  step('模擬さんを足す', () => { els.abAdd.value = '模擬 六助'; run('absentAdd()'); run('gen()'); });
  ck(!ind().some((x) => x.name === '模擬 六助') && html('abBox').includes('模擬 六助'),
     '3) 前半：「欠席の方を足す」で選んだ方のページを作る: ' + J(ind().map((x) => x.name)));
  // その日の列が無い日に選び直す：前の日の欠席の方を残さない
  routine = { ok: true, found: false, message: 'ルーティンチェックシートに、この開催日の列が見つかりませんでした。' };
  step('列の無い日に選び直す', () => run('reload()'));
  step('列の無い日で作る', () => run('gen()'));
  ck(ind().length === NAMES.length && /いません/.test(html('abBox')), '3) 前半：列の無い日に、前の日の欠席の方が残った: ' + J(ind().map((x) => x.name)));
  routine = R;
}

// ===== 4) 後半の画面：欠席の方のリファーラル発表のページを作らない =====
{
  const page = loadPage('slides_meeting_second.html', { server: SERVER, fails });
  const { els, window, run, step } = page;
  const html = (id) => String(els[id] ? els[id].innerHTML : '');
  page.flush();
  step('後半を開く', () => window.onload());
  ck(ABSENT.every((n, i) => els['ab_' + i] && els['ab_' + i].checked), '4) 後半：欠席の方のチェックが出ていない: ' + html('abBox').slice(0, 200));
  ck(/<b>5名ぶん<\/b>/.test(html('rfNote')) && /欠席の方 3名のページは作りません/.test(html('rfNote')), '4) 後半：人数の案内: ' + html('rfNote'));
  step('後半を作る', () => run('gen()'));
  const rf = () => lastOpts().referral || [];
  ck(J(rf().map((x) => x.name)) === J(['見本 一郎', '例示 五月', '模擬 六助', '空想 七海', '見本 みさき']),
     '4) 後半：リファーラル発表のページ（名簿のNo.順・欠席の方を除く）: ' + J(rf().map((x) => x.name)));
  ck(J(rf().map((x) => x.nextName)) === J(['例示 五月', '模擬 六助', '空想 七海', '見本 みさき', '']), '4) 後半：次の発表者に欠席の方が出る: ' + J(rf().map((x) => x.nextName)));
  step('架空さんのチェックを外す', () => { els.ab_1.checked = false; run('absentToggle(1)'); run('gen()'); });
  ck(J(rf().map((x) => x.name)) === J(['見本 一郎', '架空 三郎', '例示 五月', '模擬 六助', '空想 七海', '見本 みさき']),
     '4) 後半：チェックを外しても作らない: ' + J(rf().map((x) => x.name)));
}

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('欠席の方のページ: 検査 ' + checks + ' 件 OK: 欄の読み方（当欠・遅刻・代理・復帰・なし）・名簿の方に合わせる・'
  + '前半のメンバーのページ（扉ページ・NEXT）・後半のリファーラル発表（次の発表者）・チェックを外す・足す・列の無い日');

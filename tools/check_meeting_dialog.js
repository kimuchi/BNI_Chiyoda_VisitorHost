// 定例会スライドの画面（slides_meeting.html）を、Node上の簡易DOMで動かしてみる。
// 構文は正しくても、ボタンを押したときに初めて出るエラー（関数が無い・値の渡し忘れ）を
// 見つけるため。サーバーの返事は、それらしい固定値で代用する。
//
//   node tools/check_meeting_dialog.js <members.json>

const fs = require('fs');
const { loadPage } = require('./lib_minidom');

const MEMBERS = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }

// ---- サーバーの代わり ----
const gk = (x) => String(x || '').normalize('NFKC').replace(/[\s・･＆&と]/g, '');
const ORDER = ['企業サポート', '研修・教育', '不動産関連', '建築・住まい', 'プロモーション',
               '暮らし・生活', '美容と健康', '飲食・エンタメ'];
const mm = MEMBERS.map((m) => ({ name: m.name, company: m.company, title: m.title, cat: m.cat,
                                  blockKey: gk(m.cat), hasPhoto: true }));
const blocks = ORDER.map((b, i) => ({ gkey: gk(b), block: b, order: i + 1,
                                      count: mm.filter((m) => m.blockKey === gk(b)).length }));
const DATE = '2026/09/23';
let routine = { ok: true, found: true, sheetName: '【23期】ルーティンチェックシート', meetingNo: '534',
                coreValue: 'Accountability', mainPresenters: [], wantedCategories: [],
                recommendations: [], regionGuestsRaw: '吉田ED' };
const GUESTS = [{ name: '坂爪　達也', role: 'Activeチャプター担当アンバサダー', hidden: true },
                { name: '吉田　まり子', role: 'エクゼティブディレクター', hidden: true }];
const sent = [];
const SERVER = {
  getMeetingSlideContext: () => ({
    ok: true, meetings: [{ dateValue: DATE, display: '第534回 2026/09/23' }],
    defaultMeeting: { dateValue: DATE, display: '第534回 2026/09/23' }, lists: null, stats: {},
    templates: { meetingFirst: true, meetingSecond: true }, routine,
    coreValues: ['Givers Gain', 'Accountability'], memberCount: MEMBERS.length,
    members: MEMBERS.map((m) => ({ no: m.no, name: m.name, company: m.company, title: m.title, hasPhoto: true })),
  }),
  getRoutineInfo: () => routine,
  getWeeklyGuests: () => ({ ok: true, guests: GUESTS }),
  getMemberPresenContext: () => ({
    ok: true, members: mm, blocks, rowsPerPage: 7,
    candidates: [{ dateValue: DATE, start: gk('建築・住まい'), longPresenter: '', longPresenterRaw: '',
                   startFrom: 'routine', startRaw: '建築　住まい　２２番　熊谷さん', startRawDate: DATE, startSteps: 0 }],
  }),
  getMeetingTemplateInfo: () => ({
    ok: true, list: [], message: '',
    referralBoxes: { companyTall: { x: 5087424, y: 2849608, cx: 6893161, cy: 1446550 },
                     categoryLow: 4071101, categoryWidth: 6893161, hasNext: true, slides: 1 },
  }),
  computeRenewalLists: () => ({ ok: false, message: '（省略）' }),
  getMeetingMusicFiles: () => ({ ok: true, files: [] }),
  getSystemVersion: () => 'test',
  generateMeetingSlides: (...args) => { sent.push(args); return { ok: true, message: '作成しました', url: 'u' }; },
};
const page = loadPage('slides_meeting.html', { server: SERVER, fails });
const { els, window, run, step } = page;
const lastOpts = () => (sent.length ? sent[sent.length - 1][3] || {} : {});
const pagesOf = (o) => (Array.isArray(o.memberPresen) ? o.memberPresen : []);

// 1) 開く：ルーティンチェックシートに「吉田ED」→ 吉田さんだけチェック
step('画面を開く', () => window.onload());
ck(els.gs_0 && !els.gs_0.checked, '坂爪さんにチェックが入っている（リージョン参加者は「吉田ED」）');
ck(els.gs_1 && els.gs_1.checked, '吉田さんにチェックが入っていない');
ck(/吉田ED/.test(els.gsNote.innerHTML), 'リージョン参加者の記載が表示されていない');

// 2) メンバープレゼンを差し込む → 前半を作る
step('メンバープレゼンを入れる', () => { els.mpOn.checked = true; run('toggleMP()'); });
ck(els.mpStart.value === gk('建築・住まい'), '始まりの業種区分が選ばれていない: ' + els.mpStart.value);
ck(/ルーティンチェックシートの記載（建築　住まい　２２番　熊谷さん）/.test(els.mpStartNote.textContent),
   '始まりの業種区分の出どころが表示されていない: ' + els.mpStartNote.textContent);
step('前半を作る', () => run("gen('meetingFirst')"));
let o = lastOpts();
ck(JSON.stringify(o.weeklyGuests) === JSON.stringify(['吉田　まり子']), 'weeklyGuests が ' + JSON.stringify(o.weeklyGuests));
ck(o.weeklyAuto === true, '自動送りが渡っていない');
ck(pagesOf(o).length > 40, 'メンバープレゼンのページが ' + pagesOf(o).length);
const last = pagesOf(o)[pagesOf(o).length - 1] || {};
ck(last.kind === 'individual' && last.nextName === '吉田　まり子', '最後の方の NEXT が ' + last.nextName);
ck((pagesOf(o)[0] || {}).kind === 'overview' && (pagesOf(o)[0] || {}).block === '建築・住まい',
   '最初のページが「建築・住まい」の扉ではない');
ck(JSON.stringify(o.referral) === '[]', '前半なのにリファーラル発表が渡っている');

// 3) 坂爪さんも入れる／自動送りを外す
step('坂爪さんも入れる', () => { els.gs_0.checked = true; els.mpAuto.checked = false; run('renderMP()'); run("gen('meetingFirst')"); });
o = lastOpts();
ck(JSON.stringify(o.weeklyGuests) === JSON.stringify(['坂爪　達也', '吉田　まり子']), 'weeklyGuests が ' + JSON.stringify(o.weeklyGuests));
ck(o.weeklyAuto === false, '自動送りを外したのに true のまま');
ck(pagesOf(o).length > 0 && pagesOf(o).every((x) => !x.autoAdvanceMs), '自動送りを外したのに自動で進むページがある');
ck((pagesOf(o)[pagesOf(o).length - 1] || {}).nextName === '坂爪　達也', 'NEXT が先頭の方（坂爪さん）になっていない');

// 4) 開催日を変えて「なし」→ チェックが外れる
routine = Object.assign({}, routine, { regionGuestsRaw: 'なし' });
step('開催日を選び直す', () => run('reload()'));
ck(!els.gs_0.checked && !els.gs_1.checked, '「なし」なのにチェックが残っている');
step('前半を作る（来訪なし）', () => run("gen('meetingFirst')"));
o = lastOpts();
ck(JSON.stringify(o.weeklyGuests) === '[]', '来訪なしの weeklyGuests が ' + JSON.stringify(o.weeklyGuests));

// 5) 後半：リファーラル発表（組版の関数を使うので、読み込み漏れがあるとここで止まる）
step('後半テンプレートを読む', () => run('loadTemplate()'));
step('リファーラル発表を作る', () => { els.rfOn.checked = true; run('renderReferral()'); run("gen('meetingSecond')"); });
o = lastOpts();
ck((o.referral || []).length === MEMBERS.length, 'リファーラル発表が ' + (o.referral || []).length + '名ぶん');
ck((o.referral || []).every((x) => x.auto === false && x.seconds === 7), 'リファーラル発表が自動送り、または7秒でない');
ck(o.weeklyGuests === null, '後半なのにアンバサダー・ディレクターが渡っている');
ck(JSON.stringify(o.memberPresen) === '[]', '後半なのにメンバープレゼンが渡っている');

console.log(`画面の動作確認: 検査 ${checks} 件`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.slice(0, 20).forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: 開く・前半（メンバープレゼン＋アンバサダー・ディレクター）・後半（リファーラル発表）');

// 定例会スライドの画面（前半 slides_meeting_first.html／後半 slides_meeting_second.html）を、
// Node上の簡易DOMで動かしてみる。構文は正しくても、ボタンを押したときに初めて出るエラー
// （関数が無い・値の渡し忘れ）や、「読み込み中」の出し忘れを見つけるため。
// サーバーの返事は、それらしい固定値で代用する。
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
const blocks = ORDER.map((b, i) => ({ gkey: gk(b), block: b, order: i + 1, known: true,
                                      count: mm.filter((m) => m.blockKey === gk(b)).length }));
const DATE = '2026/09/23';
let routine = { ok: true, found: true, sheetName: '【23期】ルーティンチェックシート', meetingNo: '534',
                coreValue: 'Accountability', mainPresenters: [], wantedCategories: [],
                recommendations: [], regionGuestsRaw: '吉田ED' };
const GUESTS = [{ name: '坂爪　達也', role: 'Activeチャプター担当アンバサダー', hidden: true },
                { name: '吉田　まり子', role: 'エクゼティブディレクター', hidden: true }];
const LISTS = { ok: true, text: { newMembers: '該当者なし', renewMembers: '該当者なし',
                                  d90: 'Aさん', d60: '該当者なし', d30: 'Bさん', overdue: '該当者なし' },
                newMembers: [], renewMembers: [], d90: [{ name: 'A', date: '2026/12/01', left: 70 }], d60: [],
                d30: [{ name: 'B', date: '2026/10/10', left: 17 }], overdue: [], noDate: [], done: [], leaving: [] };
const sent = [];
const SERVER = {
  getMeetingSlideContext: () => ({
    ok: true, meetings: [{ dateValue: DATE, display: '第534回 2026/09/23' }],
    defaultMeeting: { dateValue: DATE, display: '第534回 2026/09/23' }, lists: LISTS, stats: {},
    templates: { meetingFirst: true, meetingSecond: true, memberPresen: true }, routine,
    coreValues: ['Givers Gain', 'Accountability'], memberCount: MEMBERS.length,
    members: MEMBERS.map((m) => ({ no: m.no, name: m.name, company: m.company, title: m.title, hasPhoto: true })),
  }),
  getRoutineInfo: () => routine,
  getWeeklyGuests: () => ({ ok: true, guests: GUESTS }),
  getMemberPresenContext: () => ({
    ok: true, members: mm, blocks, rowsPerPage: 7, unusedCategories: ['研修・教育'],
    candidates: [{ dateValue: DATE, start: gk('建築・住まい'), longPresenter: '', longPresenterRaw: '',
                   startFrom: 'routine', startRaw: '建築　住まい　２２番　熊谷さん', startRawDate: DATE, startSteps: 0 }],
  }),
  getMeetingTemplateInfo: () => ({
    ok: true, message: '',
    list: [{ slide: 'ppt/slides/slide1.xml', slideNo: 1, spid: '3', name: '貢献発表054', volume: 6, video: false, key: 'slide1.xml#3' },
           { slide: 'ppt/slides/slide21.xml', slideNo: 21, spid: '3', name: '170989_1280x720', volume: 80, video: true, key: 'slide21.xml#3' }],
    referralBoxes: { companyTall: { x: 5087424, y: 2849608, cx: 6893161, cy: 1446550 },
                     categoryLow: 4071101, categoryWidth: 6893161, hasNext: true, slides: 1 },
  }),
  computeRenewalLists: () => LISTS,
  getMeetingMusicFiles: () => ({ ok: true, files: [{ id: 'f1', name: '曲A.mp3', sizeMB: 3.2 }] }),
  saveMeetingStats: () => ({ ok: true, message: '保存しました' }),
  setRenewalMark: () => ({ ok: true }),
  getSystemVersion: () => 'test',
  generateMeetingSlides: (...args) => { sent.push(args); return { ok: true, message: '作成しました', url: 'u' }; },
};
const lastCall = () => (sent.length ? sent[sent.length - 1] : []);
const lastOpts = () => lastCall()[3] || {};
const pagesOf = (o) => (Array.isArray(o.memberPresen) ? o.memberPresen : []);
const shown = (el) => !!el && el.style.display !== 'none' && el.style.display !== undefined;

// ===================== 前半 =====================
{
  const page = loadPage('slides_meeting_first.html', { server: SERVER, fails });
  const { els, window, run, step } = page;
  page.flush();                                      // 読み込み時に呼ぶ「版」の返事を先に届けておく

  // 開いた直後は「読み込み中」で、作成ボタンは押せない
  window.onload();
  ck(shown(els.loading) && /読み込み中/.test(els.loading.innerHTML), '前半：開いた直後に「読み込み中」が出ていない');
  ck(els.genBtn.disabled === true, '前半：読み込み中なのに作成ボタンが押せる');
  page.flushOne();                                   // 開催日などが届く → テンプレートとメンバーを読みに行く
  ck(/前半テンプレート/.test(els.loading.innerHTML) && /メンバーと写真/.test(els.loading.innerHTML),
     '前半：テンプレートとメンバーの読み込み中の表示が無い: ' + els.loading.innerHTML);
  ck(els.genBtn.disabled === true, '前半：テンプレートの読み込み中なのに作成ボタンが押せる');
  step('前半：読み込みを終える', () => {});
  ck(!shown(els.loading), '前半：読み込みが終わっても「読み込み中」が消えない');
  ck(els.genBtn.disabled === false, '前半：読み込みが終わっても作成ボタンが押せない');

  // アンバサダー・ディレクター（リージョン参加者は「吉田ED」）
  ck(els.gs_0 && !els.gs_0.checked, '坂爪さんにチェックが入っている');
  ck(els.gs_1 && els.gs_1.checked, '吉田さんにチェックが入っていない');
  ck(/吉田ED/.test(els.gsNote.innerHTML), 'リージョン参加者の記載が表示されていない');
  // メンバーのページ（既定で入れる）
  ck(els.mpOn.checked === true, 'メンバーのページを入れるチェックが既定で入っていない');
  ck(els.mpStart.value === gk('建築・住まい'), '始まりの業種区分: ' + els.mpStart.value);
  ck(/ルーティンチェックシートの記載（建築　住まい　２２番　熊谷さん）/.test(els.mpStartNote.textContent),
     '始まりの業種区分の出どころが表示されていない: ' + els.mpStartNote.textContent);
  ck(/建築・住まい/.test(els.mpOrder.innerHTML) && /first/.test(els.mpOrder.innerHTML), '順番の一覧が出ていない');
  ck(/研修・教育/.test(els.mpWarn.innerHTML) && /tidy\(\)/.test(els.mpWarn.innerHTML), '使われていない業種区分の案内が出ていない');
  ck(els.t_新メンバー.value === '該当者なし', '新メンバーが入っていない');
  ck(!els.t_更新90, '前半に更新状況一覧（後半のページ）が出ている');

  step('前半を作る', () => run('gen()'));
  let o = lastOpts();
  ck(lastCall()[0] === 'meetingFirst', '前半の作成で種類が ' + lastCall()[0]);
  ck(JSON.stringify(o.weeklyGuests) === JSON.stringify(['吉田　まり子']), 'weeklyGuests が ' + JSON.stringify(o.weeklyGuests));
  ck(o.weeklyAuto === true, '自動送りが渡っていない');
  ck(pagesOf(o).length > 40, 'メンバーのページが ' + pagesOf(o).length);
  const last = pagesOf(o)[pagesOf(o).length - 1] || {};
  ck(last.kind === 'individual' && last.nextName === '吉田　まり子', '最後の方の NEXT が ' + last.nextName);
  ck((pagesOf(o)[0] || {}).block === '建築・住まい', '最初のページが「建築・住まい」の扉ではない');
  ck(o.referral === undefined && o.music === undefined, '前半なのにリファーラル発表・音楽が渡っている');
  ck(lastCall()[1]['新メンバー'] === '該当者なし' && lastCall()[1]['更新90'] === undefined, '前半の差し込む値がおかしい');

  step('坂爪さんも入れる', () => { els.gs_0.checked = true; els.mpAuto.checked = false; run('renderMP()'); run('gen()'); });
  o = lastOpts();
  ck(JSON.stringify(o.weeklyGuests) === JSON.stringify(['坂爪　達也', '吉田　まり子']), 'weeklyGuests が ' + JSON.stringify(o.weeklyGuests));
  ck(o.weeklyAuto === false && pagesOf(o).every((x) => !x.autoAdvanceMs), '自動送りを外したのに自動で進む');
  ck((pagesOf(o)[pagesOf(o).length - 1] || {}).nextName === '坂爪　達也', 'NEXT が先頭の方（坂爪さん）になっていない');

  // 開催日を選び直すと、読み込み中の間は作成できない
  routine = Object.assign({}, routine, { regionGuestsRaw: 'なし' });
  const before = sent.length;
  run('reload()');
  ck(shown(els.loading) && els.genBtn.disabled === true, '前半：選び直した直後に「読み込み中」になっていない');
  run('gen()');
  ck(sent.length === before, '前半：読み込み中なのに作成が始まった');
  step('前半：選び直しを終える', () => {});
  ck(!els.gs_0.checked && !els.gs_1.checked, '「なし」なのにチェックが残っている');
  step('メンバーのページを入れずに作る', () => { els.mpOn.checked = false; run('toggleMP()'); run('gen()'); });
  o = lastOpts();
  ck(sent.length === before + 1 && pagesOf(o).length === 0 && JSON.stringify(o.weeklyGuests) === '[]',
     'メンバーのページなしの作成がおかしい: ' + JSON.stringify({ n: pagesOf(o).length, g: o.weeklyGuests }));
}

// ===================== 後半 =====================
{
  // 推薦のことば：定例会中2組・アフター1組（名簿に無い方が1人）・翌週以降1組
  const N = (i) => MEMBERS[i].name;
  const who = (i) => ({ raw: N(i).split(' ')[0] + 'さん', name: N(i), matched: true });
  routine = Object.assign({}, routine, {
    recommendationsRaw: '定例会中\n①A→B\n②C→D\nアフター\nE→大野さん\n翌週以降\nF→G',
    recommendations: [
      { giver: who(0), receiver: who(1), raw: 'A→B', when: 'during' },
      { giver: who(2), receiver: who(3), raw: 'C→D', when: 'during' },
      { giver: who(4), receiver: { raw: '大野さん', name: '', matched: false }, raw: 'E→大野さん', when: 'after' },
      { giver: who(5), receiver: who(6), raw: 'F→G', when: 'later' },
    ] });
  const page = loadPage('slides_meeting_second.html', { server: SERVER, fails });
  const { els, window, run, step } = page;
  page.flush();

  window.onload();
  ck(shown(els.loading) && /読み込み中/.test(els.loading.innerHTML), '後半：開いた直後に「読み込み中」が出ていない');
  ck(els.genBtn.disabled === true, '後半：読み込み中なのに作成ボタンが押せる');
  page.flushOne();                                   // 開催日などが届く → 後半テンプレートを読みに行く
  ck(/後半テンプレート/.test(els.loading.innerHTML), '後半：テンプレートの読み込み中の表示が無い: ' + els.loading.innerHTML);
  ck(/読み込み中/.test(els.rfNote.innerHTML), '後半：リファーラル発表の欄に読み込み中が出ていない');
  ck(els.genBtn.disabled === true, '後半：テンプレートの読み込み中なのに作成ボタンが押せる');
  step('後半：読み込みを終える', () => {});
  ck(!shown(els.loading) && els.genBtn.disabled === false, '後半：読み込みが終わっても作成できない');

  ck(els.rfOn.checked === true && /名ぶん/.test(els.rfNote.innerHTML), 'リファーラル発表が既定で入っていない: ' + els.rfNote.innerHTML);
  ck(!!els.au_0 && !els.au_1 && !!els.av_1, '音楽の行に差し替え欄が無い、または動画の行に差し替え欄が出ている');
  ck(!!els.au_0 && els.au_0.options.some((x) => x.text === '曲A.mp3'), '登録した音楽が差し替えの候補に出ていない');
  ck(els.t_更新90.value === 'Aさん' && /B/.test(els.mkList.innerHTML), '更新状況一覧が入っていない');
  ck(!els.mpOn && !els.core, '後半に前半の欄が出ている');

  step('後半を作る', () => { els.au_0.value = 'f1'; els.av_1.value = '50'; run('gen()'); });
  const o = lastOpts();
  ck(lastCall()[0] === 'meetingSecond', '後半の作成で種類が ' + lastCall()[0]);
  ck((o.referral || []).length === MEMBERS.length && o.referral.every((x) => x.auto === false && x.seconds === 7),
     'リファーラル発表: ' + (o.referral || []).length + '名ぶん');
  ck(JSON.stringify(o.music) === JSON.stringify({ 'slide1.xml#3': { fileId: 'f1' }, 'slide21.xml#3': { volume: 50 } }),
     '音楽の指定: ' + JSON.stringify(o.music));
  ck(o.memberPresen === undefined && o.weeklyGuests === undefined, '後半なのにメンバーのページが渡っている');
  ck(lastCall()[1]['更新90'] === 'Aさん' && lastCall()[1]['新メンバー'] === undefined, '後半の差し込む値がおかしい');

  // 推薦のことば：ルーティンチェックシートの組が、定例会中・アフターに分かれて入っている
  const rows = (w) => (els[w === 'during' ? 'recoDuring' : 'recoAfter'].innerHTML.match(/id="rp_\w+_g_\d+"/g) || []).length;
  ck(rows('during') === 2 && rows('after') === 1, '推薦のことばの組の数: 定例会中' + rows('during') + '・アフター' + rows('after'));
  ck(els.rp_during_g_0.value === N(0) && els.rp_during_r_0.value === N(1) && els.rp_during_g_1.value === N(2)
     && els.rp_during_r_1.value === N(3), '定例会中の組が入っていない');
  ck(els.rp_after_g_0.value === N(4) && els.rp_after_r_0.value === '', 'アフターの組が入っていない');
  ck(/名簿の氏名と一致しなかった方：大野さん/.test(els.rcNote.innerHTML), '名簿に無い方の案内が出ていない: ' + els.rcNote.innerHTML);
  ck(/翌週以降の分（F→G）は入れていません/.test(els.rcNote.innerHTML), '翌週以降の分の案内が出ていない');
  const pairsOf = (x) => (x.recommendPairs || []).map((q) => (q.after ? 'A:' : 'D:') + q.giver.name + '>' + q.receiver.name).join(',');
  ck(pairsOf(o) === `D:${N(0)}>${N(1)},D:${N(2)}>${N(3)},A:${N(4)}>`, '渡した組: ' + pairsOf(o));
  const p0 = (o.recommendPairs || [])[0] || {};
  ck(p0.giver && p0.giver.company === MEMBERS[0].company && p0.giver.category === '【' + MEMBERS[0].title + '】',
     '推薦する人の会社名・カテゴリーが渡っていない: ' + JSON.stringify(p0.giver));
  ck(o.recommenders === undefined && lastCall()[1]['推薦のことば1氏名'] === undefined, '推薦のことばを1組だけの形でも渡している');

  // 組を足す・外す（外しても、ほかの組の選んだ値は残る）
  step('定例会中に組を足す', () => {
    els.rp_during_g_0.value = N(7);                  // 画面で選び直した値は、足したあとも残る
    run("addPair('during')");
    els.rp_during_g_2.value = N(8); els.rp_during_r_2.value = N(9);
    run("addPair('after')");                         // 2人とも空の組は渡さない
    run('gen()');
  });
  ck(rows('during') === 3 && rows('after') === 2, '組を足したあとの数: 定例会中' + rows('during') + '・アフター' + rows('after'));
  ck(pairsOf(lastOpts()) === `D:${N(7)}>${N(1)},D:${N(2)}>${N(3)},D:${N(8)}>${N(9)},A:${N(4)}>`,
     '組を足したあとに渡した組: ' + pairsOf(lastOpts()));
  step('定例会中の1組目を外す', () => { run("delPair('during',0)"); run('gen()'); });
  ck(rows('during') === 2 && els.rp_during_g_0.value === N(2) && els.rp_during_g_1.value === N(8),
     '外したあとの並び: ' + els.rp_during_g_0.value + ' / ' + els.rp_during_g_1.value);
  ck(pairsOf(lastOpts()) === `D:${N(2)}>${N(3)},D:${N(8)}>${N(9)},A:${N(4)}>`, '外したあとに渡した組: ' + pairsOf(lastOpts()));
  step('定例会中の組を全部外す', () => { run("delPair('during',0)"); run("delPair('during',0)"); run('gen()'); });
  ck(rows('during') === 0 && /（なし）/.test(els.recoDuring.innerHTML), '定例会中の組が無いときの表示: ' + els.recoDuring.innerHTML);
  ck(pairsOf(lastOpts()) === `A:${N(4)}>`, '定例会中なしで渡した組: ' + pairsOf(lastOpts()));

  // 開催日を選び直すと、その日のルーティンチェックシートの組に入れ替わる（記載が無ければ空の1組）
  routine = Object.assign({}, routine, { recommendations: [], recommendationsRaw: '' });
  step('後半：開催日を選び直す', () => run('reload()'));
  ck(rows('during') === 1 && rows('after') === 0 && els.rp_during_g_0.value === '',
     '選び直したあとの組: 定例会中' + rows('during') + '・アフター' + rows('after'));
  step('推薦のことばなしで作る', () => run('gen()'));
  ck(Array.isArray(lastOpts().recommendPairs) && lastOpts().recommendPairs.length === 0,
     '組が無いときに渡した値: ' + JSON.stringify(lastOpts().recommendPairs));
}

console.log(`画面の動作確認: 検査 ${checks} 件`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.slice(0, 30).forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: 前半（読み込み中・メンバーのページ・アンバサダー・ディレクター）／後半（読み込み中・推薦のことば・リファーラル発表・音楽）');

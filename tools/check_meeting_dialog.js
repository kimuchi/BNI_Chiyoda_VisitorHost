// 定例会スライドの画面（前半 slides_meeting_first.html／後半 slides_meeting_second.html）を、
// Node上の簡易DOMで動かしてみる。構文は正しくても、ボタンを押したときに初めて出るエラー
// （関数が無い・値の渡し忘れ）や、「読み込み中」の出し忘れを見つけるため。
// サーバーの返事は、それらしい固定値で代用する。
//
//   node tools/check_meeting_dialog.js               … 架空の60名の名簿で動かす
//   node tools/check_meeting_dialog.js <members.json> … 手元の名簿（リポジトリには入れない）で動かす

const fs = require('fs');
const { loadPage } = require('./lib_minidom');

// 架空の名簿（60名。ブロックは6つに分かれ、研修・教育と飲食・エンタメは0名）
function fictionalMembers() {
  const fams = ['見本', '試験', '架空', '仮名', '例示', '模擬', '空想', '新入', '単独', '別人',
                '甲野', '乙野', '丙野', '丁野', '戊野', '己野', '庚野', '辛野', '壬野', '癸野'];
  const gives = ['一郎', '二郎', '三郎'];
  const cats = ['企業サポート', '不動産関連', '建築・住まい', 'プロモーション', '暮らし・生活', '美容と健康'];
  return Array.from({ length: 60 }, (_, i) => ({
    no: String(i + 1), name: fams[i % 20] + ' ' + gives[Math.floor(i / 20)], kana: '', cat: cats[i % 6],
    title: 'カテゴリー' + (i + 1), company: '見本会社' + (i + 1), role: '', position: '', comment: '', refer: '', collab: '' }));
}
const MEMBERS = process.argv[2] ? JSON.parse(fs.readFileSync(process.argv[2], 'utf8')) : fictionalMembers();
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
// 前半の新メンバー・更新メンバー／バイスプレジデントによる報告／ネットワーキングリーダー（ルーティンチェックシートから読んだ形）
//   新メンバー … 名簿の方と、名簿に無い方。更新 … 2年と、年数の記載なし。ネットワーキングリーダー … 1to1 はお2人（1人は名簿に無い）
const NM = (i) => MEMBERS[i].name;
const FIRST_HALF = {
  newMembers: [{ raw: 'Aさん', name: NM(20), matched: true, years: 0, category: '' },
               { raw: '新井 花子さん', name: '', matched: false, years: 0, category: 'エステサロン' }],
  newMembersRaw: 'Aさん、新井 花子さん（エステサロン）',
  renewMembers: [{ raw: 'Bさん', name: NM(34), matched: true, years: 2, category: '' },
                 { raw: 'Cさん', name: NM(0), matched: true, years: 0, category: '' }],
  renewMembersRaw: 'Bさん（2年更新）、Cさん',
  vpReport: { avg: '308', month: '2026-08', count: '281', perWeek: '70', from: '2026-03', to: '2026-08', total: '1,927', thanks: '54億8,074万円' },
  vpReportFrom: '2026/09/16',
  networkingLeaders: { month: '2026-08', items: [
    { key: 'ceu', label: 'CEU', value: '23', unit: 'ポイント', winners: [{ raw: 'Dさん', name: NM(44), matched: true, category: '' }] },
    { key: 'oto', label: '1to1', value: '19', unit: '回', winners: [{ raw: 'Eさん', name: NM(29), matched: true, category: '' },
                                                                      { raw: '大森さん', name: '', matched: false, category: '' }] }] },
  networkingLeadersRaw: '2026年8月のネットワーキングリーダーの発表…',
  firstOfMonth: false,
};
let routine = { ok: true, found: true, sheetName: '【23期】ルーティンチェックシート', meetingNo: '534',
                coreValue: 'Accountability', mainPresenters: [], wantedCategories: [],
                recommendations: [], regionGuestsRaw: '大庭ED', firstHalf: FIRST_HALF };
const GUESTS = [{ name: '坂上　達彦', role: 'Activeチャプター担当アンバサダー', hidden: true },
                { name: '大庭　まり子', role: 'エクゼティブディレクター', hidden: true }];
const LISTS = { ok: true, text: { newMembers: '該当者なし', renewMembers: '該当者なし',
                                  d90: 'Aさん', d60: '該当者なし', d30: 'Bさん', overdue: '該当者なし' },
                newMembers: [], renewMembers: [], d90: [{ name: 'A', date: '2026/12/01', left: 70 }], d60: [],
                d30: [{ name: 'B', date: '2026/10/10', left: 17 }], overdue: [], noDate: [], done: [], leaving: [] };
const sent = [];
const J = (x) => JSON.stringify(x);
const SERVER = {
  getMeetingSlideContext: () => ({
    ok: true, meetings: [{ dateValue: DATE, display: '第534回 2026/09/23' }],
    defaultMeeting: { dateValue: DATE, display: '第534回 2026/09/23' }, lists: LISTS, stats: {},
    templates: { meetingFirst: true, meetingSecond: true, memberPresen: true }, routine,
    seconds: { weekly: 45, startup: 180, visitor: 20, referral: 9 },          // チャプターの設定（既定と違う秒数）
    coreValues: ['Givers Gain', 'Accountability'], memberCount: MEMBERS.length,
    members: MEMBERS.map((m) => ({ no: m.no, name: m.name, company: m.company, title: m.title, hasPhoto: true })),
  }),
  getRoutineInfo: () => routine,
  getWeeklyGuests: () => ({ ok: true, guests: GUESTS }),
  getMemberPresenContext: () => ({
    ok: true, members: mm, blocks, rowsPerPage: 7, unusedCategories: ['研修・教育'], seconds: { weekly: 45, startup: 180 },
    candidates: [{ dateValue: DATE, start: gk('建築・住まい'), longPresenter: '', longPresenterRaw: '',
                   startFrom: 'routine', startRaw: '建築　住まい　２２番　熊田さん', startRawDate: DATE, startSteps: 0 }],
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
  setRenewalMark: () => ({ ok: true }),
  getSystemVersion: () => 'test',
  // 役職のメンバー紹介（その開催日の期の役職・チーム。メンターコーディネーターは未登録）
  getRoleIntroPreview: (d) => ({ ok: true, term: 24, label: '2026年10月〜2027年3月', holdersOk: true, holdersFrom: 24, teamsOk: true,
    holders: [{ label: 'プレジデント', name: MEMBERS[0].name }, { label: 'バイスプレジデント', name: MEMBERS[1].name },
              { label: 'メンターコーディネーター', name: '' }],
    teams: [{ name: 'メンバーシップ委員会', count: 4 }, { name: 'ビジターホスト', count: 12 }] }),
  generateMeetingSlides: (...args) => { sent.push(args); return { ok: true, message: '作成しました', url: 'u' }; },
  // スピーカーローテーション（その日から5回ぶん。1回目はローテーションの予定＝ルーティンにメインプレゼンが無い日）
  getSpeakerRotationWeeks: (d) => ({ ok: true, header: 'メインプレゼンテーション（各４分45秒）', notes: ['注意書き'],
    weeks: [0, 1, 2, 3, 4].map((i) => ({ date: i ? '2026/10/0' + (i + 1) : d, no: String(534 + i), md: 'M/D', label: (i + 1) + '月',
      source: i ? 'rotation' : 'rotation',
      people: [MEMBERS[2 * i], MEMBERS[2 * i + 1]].map((m) => ({ name: m.name, title: m.title, collab: '協業' + i })) })) }),
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

  // アンバサダー・ディレクター（リージョン参加者は「大庭ED」）
  ck(els.gs_0 && !els.gs_0.checked, '坂上さんにチェックが入っている');
  ck(els.gs_1 && els.gs_1.checked, '大庭さんにチェックが入っていない');
  ck(/大庭ED/.test(els.gsNote.innerHTML), 'リージョン参加者の記載が表示されていない');
  // メンバーのページ（既定で入れる）
  ck(els.mpOn.checked === true, 'メンバーのページを入れるチェックが既定で入っていない');
  ck(els.mpStart.value === gk('建築・住まい'), '始まりの業種区分: ' + els.mpStart.value);
  ck(/ルーティンチェックシートの記載（建築　住まい　２２番　熊田さん）/.test(els.mpStartNote.textContent),
     '始まりの業種区分の出どころが表示されていない: ' + els.mpStartNote.textContent);
  ck(/建築・住まい/.test(els.mpOrder.innerHTML) && /first/.test(els.mpOrder.innerHTML), '順番の一覧が出ていない');
  ck(/研修・教育/.test(els.mpWarn.innerHTML) && /tidy\(\)/.test(els.mpWarn.innerHTML), '使われていない業種区分の案内が出ていない');
  ck(!els.t_更新90, '前半に更新状況一覧（後半のページ）が出ている');
  // 新メンバー・更新メンバー：ルーティンチェックシートの方が1行ずつ。名簿に無い方・年数の記載なしの案内
  ck(els.nm_new.innerHTML.includes('value="' + NM(20) + '" selected') && /「新井 花子さん」（名簿にありません）/.test(els.nm_new.innerHTML),
     '新メンバーの行: ' + els.nm_new.innerHTML.slice(0, 200));
  ck(els.nm_renew.innerHTML.includes('value="' + NM(34) + '" selected') && /<option value="2" selected>2年更新/.test(els.nm_renew.innerHTML)
     && /年数の記載なし/.test(els.nm_renew.innerHTML), '更新メンバーの行: ' + els.nm_renew.innerHTML.slice(0, 200));
  ck(/新入会「Aさん、新井 花子さん（エステサロン）」/.test(els.nmDetail.innerText), '新入会の記載が出ていない: ' + els.nmDetail.innerText);
  // バイスプレジデントによる報告：読み上げる文から拾った数字。この日の欄が空なら前の回の記載
  ck(els.vp_avg.value === '308' && els.vp_month.value === '2026-08' && els.vp_count.value === '281' && els.vp_perWeek.value === '70'
     && els.vp_from.value === '2026-03' && els.vp_to.value === '2026-08' && els.vp_total.value === '1,927' && els.vp_thanks.value === '54億8,074万円',
     'バイスプレジデントによる報告の欄: ' + [els.vp_avg.value, els.vp_month.value, els.vp_total.value, els.vp_thanks.value].join(' / '));
  ck(/2026\/09\/16<\/b> の記載から読みました/.test(els.vpNote.innerHTML) && els.vp_date.textContent === '9/23',
     'バイスプレジデントによる報告の案内: ' + els.vpNote.innerHTML + ' / ' + els.vp_date.textContent);
  // ネットワーキングリーダー：原稿があるので表示。月の最初ではないことも出す
  ck(els.nlOn.checked === true && /最初の定例会ではありません/.test(els.nlWhy.innerHTML) && /原稿から/.test(els.nlWhy.innerHTML),
     'ネットワーキングリーダーの表示: ' + els.nlOn.checked + ' ' + els.nlWhy.innerHTML);
  ck(els.nl_month.value === '2026-08' && els.nlTable.innerHTML.includes('value="' + NM(44) + '" selected')
     && /「大森さん」（名簿にありません）/.test(els.nlTable.innerHTML) && /名簿の氏名と一致しなかった方：大森さん/.test(els.nlNote.innerHTML),
     'ネットワーキングリーダーの表: ' + els.nlTable.innerHTML.slice(0, 200));

  // 役職のメンバー紹介：その期の方と、未登録の役職を出す。既定で入れる
  ck(els.roleOn.checked === true && /24期（2026年10月〜2027年3月）/.test(els.rolePreview.innerHTML)
     && els.rolePreview.innerHTML.includes('プレジデント：' + MEMBERS[0].name) && /未登録の役職：メンターコーディネーター/.test(els.rolePreview.innerHTML)
     && /メンバーシップ委員会 4名、ビジターホスト 12名/.test(els.rolePreview.innerHTML),
     '前半：役職のメンバー紹介の表示: ' + els.rolePreview.innerHTML.slice(0, 200));
  step('前半を作る', () => run('gen()'));
  let o = lastOpts();
  ck(lastCall()[0] === 'meetingFirst', '前半の作成で種類が ' + lastCall()[0]);
  ck(o.roleIntro === true, '前半：役職のメンバー紹介を入れる指定が渡らない: ' + o.roleIntro);
  step('前半：役職のメンバー紹介を外して作る', () => { els.roleOn.checked = false; run('gen()'); });
  ck(lastOpts().roleIntro === false, '前半：役職のメンバー紹介を外しても渡る: ' + lastOpts().roleIntro);
  els.roleOn.checked = true;
  step('前半を作り直す', () => run('gen()'));
  o = lastOpts();
  ck(JSON.stringify(o.weeklyGuests) === JSON.stringify(['大庭　まり子']), 'weeklyGuests が ' + JSON.stringify(o.weeklyGuests));
  ck(o.weeklyAuto === true, '自動送りが渡っていない');
  ck(pagesOf(o).length > 40, 'メンバーのページが ' + pagesOf(o).length);
  // カウントダウンの秒数はチャプターの設定（ウィークリー45秒）。スタートアッププレゼンの方はいない日
  ck(pagesOf(o).filter((x) => x.kind === 'individual').every((x) => x.countdownSec === 45 && x.auto === true),
     'メンバーのページの秒数・自動送り: ' + J(pagesOf(o).filter((x) => x.kind === 'individual').slice(0, 2).map((x) => [x.countdownSec, x.auto])));
  ck(/カウントダウンは 45秒。/.test(els.mpNote2.innerHTML), '秒数の案内: ' + els.mpNote2.innerHTML);
  const last = pagesOf(o)[pagesOf(o).length - 1] || {};
  ck(last.kind === 'individual' && last.nextName === '大庭　まり子', '最後の方の NEXT が ' + last.nextName);
  ck((pagesOf(o)[0] || {}).block === '建築・住まい', '最初のページが「建築・住まい」の扉ではない');
  ck(o.referral === undefined && o.music === undefined, '前半なのにリファーラル発表・音楽が渡っている');
  // スピーカーローテーション：ルーティンにメインプレゼンが無い日は、ローテーションの2名が入り、表も渡る
  ck(els.mp1.value === MEMBERS[0].name && els.mp2.value === MEMBERS[1].name && /スピーカーローテーションのお2人/.test(els.mpNote.innerHTML),
     'ローテーションの2名が入っていない: ' + els.mp1.value + ' / ' + els.mp2.value);
  ck(/第534回/.test(els.rotPreview.innerHTML) && (els.rotPreview.innerHTML.match(/<br>/g) || []).length === 4, 'ローテーションの予定が出ていない');
  ck(o.speakerRotation && o.speakerRotation.weeks.length === 5 && o.speakerRotation.header && o.speakerRotation.notes.length === 1,
     'ローテーションの表のデータが渡っていない: ' + JSON.stringify(o.speakerRotation).slice(0, 120));
  ck(lastCall()[1]['新メンバー'] === NM(20) + 'さん、新井 花子さん' && lastCall()[1]['更新メンバー'] === NM(34) + 'さん（2年更新）、' + NM(0) + 'さん（1年更新）'
     && lastCall()[1]['更新90'] === undefined, '前半の差し込む値がおかしい: ' + lastCall()[1]['新メンバー'] + ' / ' + lastCall()[1]['更新メンバー']);
  // 新メンバー・更新メンバー／バイスプレジデントによる報告／ネットワーキングリーダーが渡る
  ck(J(o.memberPages) === J({ newMembers: [{ name: NM(20), raw: 'Aさん', category: '', years: 0 }, { name: '', raw: '新井 花子さん', category: 'エステサロン', years: 0 }],
                              renewMembers: [{ name: NM(34), raw: 'Bさん', category: '', years: 2 }, { name: NM(0), raw: 'Cさん', category: '', years: 1 }] }),
     '新メンバー・更新メンバーの渡し方: ' + J(o.memberPages));
  ck(o.vpReport && o.vpReport.avg === '308' && o.vpReport.thanks === '54億8,074万円' && o.vpReport.weekCount === '' && o.vpReport.from === '2026-03',
     'バイスプレジデントによる報告の渡し方: ' + J(o.vpReport));
  ck(o.networkingLeaders && o.networkingLeaders.show === true && o.networkingLeaders.month === '2026-08'
     && J(o.networkingLeaders.items.map((it) => [it.key, it.value, it.winners.map((w) => w.name || w.raw)]))
        === J([['ceu', '23', [NM(44)]], ['oto', '19', [NM(29), '大森さん']]]),
     'ネットワーキングリーダーの渡し方: ' + J(o.networkingLeaders));
  step('新メンバー・更新メンバーを直す', () => {
    run("nmDel('new',1)"); run("nmYears(1,'2')"); run("nmAdd('renew')"); run("nmSet('renew',2,'" + NM(5) + "')");
    els.vp_week.value = '12'; els.vp_weekExt.value = '3';
    run("nlAdd('ceu')"); run("nlSet('ceu',1,'" + NM(3) + "')"); run("nlSet('oto',1,'')");
    run('gen()');
  });
  ck(J(lastOpts().memberPages.newMembers.map((x) => x.name)) === J([NM(20)])
     && J(lastOpts().memberPages.renewMembers.map((x) => [x.name, x.years])) === J([[NM(34), 2], [NM(0), 2], [NM(5), 1]]),
     '直したあとの新メンバー・更新メンバー: ' + J(lastOpts().memberPages));
  ck(lastOpts().vpReport.weekCount === '12' && lastOpts().vpReport.weekExt === '3', '速報の数が渡らない: ' + J(lastOpts().vpReport));
  ck(J(lastOpts().networkingLeaders.items.map((it) => [it.key, it.winners.map((w) => w.name || w.raw)]))
     === J([['ceu', [NM(44), NM(3)]], ['oto', [NM(29)]]]), '直したあとのネットワーキングリーダー: ' + J(lastOpts().networkingLeaders.items));
  step('ネットワーキングリーダーのページを出さない', () => { els.nlOn.checked = false; run('renderNL()'); run('gen()'); });
  ck(lastOpts().networkingLeaders.show === false && els.nlBox.style.display === 'none', 'ネットワーキングリーダーを出さない指定が渡らない');

  step('スタートアッププレゼンの方を選ぶ', () => { els.mpLong.value = MEMBERS[3].name; run('renderMP()'); run('gen()'); });
  {
    const ind = pagesOf(lastOpts()).filter((x) => x.kind === 'individual');
    const lp = ind.filter((x) => x.name === MEMBERS[3].name);
    ck(lp.length === 1 && lp[0].countdownSec === 180 && ind.filter((x) => x.countdownSec === 180).length === 1,
       'スタートアッププレゼンの方だけ3分: ' + J(lp.map((x) => x.countdownSec)));
    ck(new RegExp('カウントダウンは 45秒（' + MEMBERS[3].name + 'さんは 3分）').test(els.mpNote2.innerHTML), '秒数の案内（スタートアップ）: ' + els.mpNote2.innerHTML);
  }
  step('スタートアッププレゼンの方を外す', () => { els.mpLong.value = ''; run('renderMP()'); });
  step('坂上さんも入れる', () => { els.gs_0.checked = true; els.mpAuto.checked = false; run('renderMP()'); run('gen()'); });
  o = lastOpts();
  // メインプレゼンを選び直すと、表の1回目もその2名になる。表を作らないチェックなら渡さない
  step('メインプレゼンを選び直す', () => { els.mp1.value = MEMBERS[5].name; run('gen()'); });
  ck(lastOpts().speakerRotation.weeks[0].people[0].name === MEMBERS[5].name, '表の1回目が画面のメインプレゼンになっていない');
  step('表を作らない', () => { els.rotOn.checked = false; run('gen()'); });
  ck(lastOpts().speakerRotation === null, '表を作らないのにデータが渡っている');
  step('表を作る・元に戻す', () => { els.rotOn.checked = true; els.mp1.value = MEMBERS[0].name; });
  ck(JSON.stringify(o.weeklyGuests) === JSON.stringify(['坂上　達彦', '大庭　まり子']), 'weeklyGuests が ' + JSON.stringify(o.weeklyGuests));
  ck(o.weeklyAuto === false && pagesOf(o).every((x) => !x.auto), '自動送りを外したのに自動で進む');
  ck((pagesOf(o)[pagesOf(o).length - 1] || {}).nextName === '坂上　達彦', 'NEXT が先頭の方（坂上さん）になっていない');

  // 開催日を選び直すと、読み込み中の間は作成できない。
  // 選び直した日は月の最初の定例会で、ネットワーキングリーダーの原稿はまだ空欄・新メンバーと更新メンバーはいない
  routine = Object.assign({}, routine, { regionGuestsRaw: 'なし', firstHalf: Object.assign({}, FIRST_HALF, {
    newMembers: [], newMembersRaw: '', renewMembers: [], renewMembersRaw: '', networkingLeaders: null, networkingLeadersRaw: 'ー',
    vpReport: null, vpReportFrom: '', firstOfMonth: true }) });
  const before = sent.length;
  run('reload()');
  ck(shown(els.loading) && els.genBtn.disabled === true, '前半：選び直した直後に「読み込み中」になっていない');
  run('gen()');
  ck(sent.length === before, '前半：読み込み中なのに作成が始まった');
  step('前半：選び直しを終える', () => {});
  ck(!els.gs_0.checked && !els.gs_1.checked, '「なし」なのにチェックが残っている');
  ck(/いません/.test(els.nm_new.innerHTML) && /いません/.test(els.nm_renew.innerHTML), '新メンバー・更新メンバーがいない日の表示');
  ck(els.nlOn.checked === true && /最初の定例会/.test(els.nlWhy.innerHTML) && /まだ空欄/.test(els.nlWhy.innerHTML) && els.nl_month.value === '2026-08',
     '月の最初の定例会（原稿はまだ）: ' + els.nlWhy.innerHTML);
  ck(els.vp_avg.value === '' && /記載がありません/.test(els.vpNote.innerHTML), 'バイスプレジデントによる報告の記載が無い日: ' + els.vpNote.innerHTML);
  step('メンバーのページを入れずに作る', () => { els.mpOn.checked = false; run('toggleMP()'); run('gen()'); });
  o = lastOpts();
  ck(sent.length === before + 1 && pagesOf(o).length === 0 && JSON.stringify(o.weeklyGuests) === '[]',
     'メンバーのページなしの作成がおかしい: ' + JSON.stringify({ n: pagesOf(o).length, g: o.weeklyGuests }));
  ck(J(o.memberPages) === J({ newMembers: [], renewMembers: [] }) && o.networkingLeaders.show === true && o.networkingLeaders.items.length === 0
     && lastCall()[1]['新メンバー'] === '該当者なし', '新メンバーなし・ネットワーキングリーダーの原稿なしの渡し方: ' + J([o.memberPages, o.networkingLeaders]));
}

// ===================== 後半 =====================
{
  // 推薦のことば：定例会中2組・アフター1組（名簿に無い方が1人）・翌週以降1組
  const N = (i) => MEMBERS[i].name;
  const who = (i) => ({ raw: N(i).split(' ')[0] + 'さん', name: N(i), matched: true });
  routine = Object.assign({}, routine, {
    recommendationsRaw: '定例会中\n①A→B\n②C→D\nアフター\nE→大森さん\n翌週以降\nF→G',
    recommendations: [
      { giver: who(0), receiver: who(1), raw: 'A→B', when: 'during' },
      { giver: who(2), receiver: who(3), raw: 'C→D', when: 'during' },
      { giver: who(4), receiver: { raw: '大森さん', name: '', matched: false }, raw: 'E→大森さん', when: 'after' },
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
  ck((o.referral || []).length === MEMBERS.length && o.referral.every((x) => x.auto === false && x.seconds === 9),
     'リファーラル発表（秒数はチャプターの設定の9秒）: ' + (o.referral || []).length + '名ぶん・' + J((o.referral || [])[0] && o.referral[0].seconds));
  ck(els.rfSec.value === '9', 'リファーラル発表の秒数の欄: ' + els.rfSec.value);
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
  ck(/名簿の氏名と一致しなかった方：大森さん/.test(els.rcNote.innerHTML), '名簿に無い方の案内が出ていない: ' + els.rcNote.innerHTML);
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

  // 更新状況の「更新した」を押しても、手で直した推薦のことば・抽選はそのまま（以前は全部を読み直して、
  // ルーティンチェックシートの内容に戻っていた）。読み直すのは更新状況の一覧だけ
  step('推薦のことば・抽選を手で直して、「更新した」を押す', () => {
    run("addPair('during')");
    els.rp_during_g_0.value = N(10); els.rp_during_r_0.value = N(11);
    els.lt1.value = N(12); els.lt2.value = N(13);
    page.log.calls.length = 0;
    run("mark(0,'done')");
  });
  const calledNames = page.log.calls.map((c) => c.name);
  ck(calledNames.includes('setRenewalMark') && calledNames.includes('computeRenewalLists') && !calledNames.includes('getRoutineInfo'),
     '「更新した」で読み直したもの: ' + calledNames.join(','));
  ck(els.rp_during_g_0.value === N(10) && els.rp_during_r_0.value === N(11) && els.lt1.value === N(12) && els.lt2.value === N(13),
     '「更新した」を押したら、手で直した推薦のことば・抽選が戻った: ' + [els.rp_during_g_0.value, els.rp_during_r_0.value, els.lt1.value, els.lt2.value].join(' / '));
  step('手で直したまま作る', () => run('gen()'));
  ck(pairsOf(lastOpts()) === `D:${N(10)}>${N(11)},A:${N(4)}>`, '「更新した」のあとに渡した組: ' + pairsOf(lastOpts()));
  ck(lastCall()[1]['抽選1氏名'] === N(12) && lastCall()[1]['抽選2氏名'] === N(13),
     '「更新した」のあとに渡した抽選: ' + [lastCall()[1]['抽選1氏名'], lastCall()[1]['抽選2氏名']].join(' / '));
  step('定例会中の組を外す（次の検査のため）', () => run("delPair('during',0)"));

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
console.log('OK: 前半（読み込み中・メンバーのページ・アンバサダー・ディレクター）／後半（読み込み中・推薦のことば・リファーラル発表・音楽・「更新した」で手直しが戻らない）');

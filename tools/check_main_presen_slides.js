// メインプレゼンが1人の週・ディレクターが担当する週の、定例会スライドを確かめる（名前はすべて架空）。
//
//   node tools/check_main_presen_slides.js
//
// 画面（slides_meeting_first.html・slides_meeting_second.html を lib_minidom で動かす）
//   ・ルーティンチェックシートが「①見本さん」（1人の週）：①に見本さん、②は（なし）。色の付いた注意は出さず「1人の週です」
//   ・「①空想ED②試験さん」：①はアンバサダー・ディレクターのページの空想 太郎さん（一覧が届いてから入る）。
//     会社名・カテゴリー（リージョンの名前を【】で）はその方のページの文字。表の1回目はディレクターが先・kind 'director'
//   ・ディレクターだけの週、テンプレートにページの無いディレクター（手で選ぶ知らせ）、スピーカーローテーションの1人・ディレクターの回
//   ・後半：①②は同じ。抽選はメンバーの方だけ（1人なら①だけ）。抽選の欄でもディレクターを選べる
// サーバー（本番の *.js を lib_sheet_fake の上で。メインプレゼンのページは公式ファイルからの雛形と同じ offTwoPersonPage_ で作る）
//   ・1人の週：もう1人の文字箱と写真の枠を消し、残る方（①でも②でも）の写真の枠と文字箱をまん中に。上下はそのまま。2人の週は前のまま。
//     1人ずつの後ろの台紙（白いカード）は、残る方のものを一緒にまん中へ・もう1人のものを消す（2人にまたがる帯・端の飾りはそのまま）
//   ・後半の1枚目（前半のテンプレートから写したページ）・抽選も同じ。抽選の飾りの画像に重なるときは動かさない
//   ・ディレクターの写真：前半はその方のページの画像を指す。後半は前半のテンプレートから画像を写す（種類の登録も足す）。
//     画面から会社名・カテゴリーが来なければ、その方のページから入れる。写真の無い方の知らせには、どのページかを付ける
//   ・グループの中の図形もページの座標でまん中に。推薦のことばは1人でも動かさない。pptxとして壊れていない
process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const os = require('os');
const path = require('path');
const vm = require('vm');
const zlib = require('zlib');
const { spawnSync } = require('child_process');
const { loadPage } = require('./lib_minidom');
const { makeEnv } = require('./lib_sheet_fake');
const { readZip, writeZip } = require('./lib_zip');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }
const J = (x) => JSON.stringify(x);

// ===================== 画面 =====================
const MEMBERS = [['見本 一郎', '見本会計事務所', '税理士'], ['試験 二郎', '試験工務店', '工務店'], ['架空 三郎', '架空商事', '保険'],
                 ['仮名 四郎', '仮名デザイン', 'デザイン']]
  .map(([name, company, title], i) => ({ no: String(i + 1), name, company, title, collab: '協業' + (i + 1), hasPhoto: true }));
const GUESTS = [{ name: '空想 太郎', role: 'BNI見本リージョン', company: '空想リージョン株式会社', hasPhoto: true, hidden: true, intro: false },
                { name: '仮想 花子', role: 'アンバサダー', company: '', hasPhoto: false, hidden: true, intro: false }];
const DATE = '2026/10/07';
const gk = (x) => String(x || '').normalize('NFKC').replace(/[\s・･＆&と]/g, '');
const mm = MEMBERS.map((m) => ({ name: m.name, company: m.company, title: m.title, cat: '企業サポート', blockKey: gk('企業サポート'), hasPhoto: true }));
const member = (raw, name) => ({ raw, name, matched: true, kind: 'member' });
const director = (raw, guest) => ({ raw, name: '', matched: false, kind: 'director', guest });
let routine = null;
let rotation = null;
const setRoutine = (raw, mps) => { routine = { ok: true, found: true, sheetName: '【24期】ルーティンチェックシート', meetingNo: '536',
  coreValue: '', mainPresenters: mps, mainPresentersRaw: raw, wantedCategories: [], recommendations: [], regionGuestsRaw: '', firstHalf: {} }; };
setRoutine('①見本さん', [member('見本さん', '見本 一郎')]);
const rotWeek = (people) => ({ ok: true, header: 'メインプレゼンテーション', notes: [],
  weeks: [{ date: DATE, no: '536', md: '10/7', label: '', source: 'rotation', kind: people.length > 1 ? 'two' : 'one', people },
          { date: '2026/10/14', no: '537', md: '10/14', label: '', source: 'rotation', kind: 'two',
            people: [MEMBERS[2], MEMBERS[3]].map((m) => ({ name: m.name, title: m.title, collab: m.collab, company: m.company, inMaster: true, kind: 'member' })) }] });
rotation = rotWeek([MEMBERS[0], MEMBERS[1]].map((m) => ({ name: m.name, title: m.title, collab: m.collab, company: m.company, inMaster: true, kind: 'member' })));
let guestList = GUESTS;
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
  getWeeklyGuests: () => ({ ok: true, guests: guestList }),
  getMemberPresenContext: () => ({ ok: true, members: mm, blocks: [{ gkey: gk('企業サポート'), block: '企業サポート', order: 1, known: true, count: mm.length }],
    rowsPerPage: 7, unusedCategories: [], seconds: { weekly: 30, startup: 150 },
    candidates: [{ dateValue: DATE, start: gk('企業サポート'), longPresenter: '', longPresenterRaw: '', startFrom: 'routine', startRaw: '', startRawDate: DATE, startSteps: 0 }] }),
  getSpeakerRotationWeeks: () => rotation,
  getRoleIntroPreview: () => ({ ok: true, term: 24, label: '', holdersOk: true, holders: [], teams: [] }),
  getMeetingTemplateInfo: () => ({ ok: true, message: '', list: [], referralBoxes: null }),
  computeRenewalLists: () => SERVER.getMeetingSlideContext().lists,
  getMeetingMusicFiles: () => ({ ok: true, files: [] }),
  getRecommendationSlides: () => ({ ok: true, slides: [] }),
  generateMeetingSlides: (...args) => { sent.push(args); return { ok: true, message: '作成しました', url: 'u' }; },
};
const last = () => sent[sent.length - 1] || [];
const orange = (h) => /#c60/.test(h);

// ---- 前半の画面 ----
{
  const page = loadPage('slides_meeting_first.html', { server: SERVER, fails });
  const { els, window, run, step } = page;
  page.flush();
  // 1) 1人の週（①見本さん）
  step('前半：開く（①見本さん）', () => window.onload());
  ck(els.mp1.value === '見本 一郎' && els.mp2.value === '', '1) 1人の週の①②: ' + J([els.mp1.value, els.mp2.value]));
  ck(/1人の週です/.test(els.mpNote.innerHTML) && !orange(els.mpNote.innerHTML), '1) 1人の週の知らせ（注意の色は付けない）: ' + els.mpNote.innerHTML);
  ck(els.mp1.options.some((o) => o.value === '空想 太郎') && els.mp1.options[0].value === '',
     '1) ①②の選択肢にアンバサダー・ディレクター・（なし）が無い: ' + J(els.mp1.options.map((o) => o.value)));
  step('前半：作る（1人）', () => run('gen()'));
  let [, vals, , o] = last();
  ck(J(o.mainPresenters) === J(['見本 一郎', '']) && vals['メインプレゼン2氏名'] === '' && vals['メインプレゼン1カテゴリー'] === '【税理士】',
     '1) 1人の週に渡す値: ' + J([o.mainPresenters, vals['メインプレゼン1カテゴリー'], vals['メインプレゼン2氏名']]));
  ck(o.speakerRotation && o.speakerRotation.weeks[0].people.length === 1 && o.speakerRotation.weeks[0].kind === 'one'
     && o.speakerRotation.weeks[0].people[0].kind === 'member' && o.speakerRotation.weeks[0].people[0].inMaster === true,
     '1) 表の1回目（1人だけ・空の2人目を足さない）: ' + J(o.speakerRotation && o.speakerRotation.weeks[0]));

  // 2) ディレクター＋メンバー（①空想ED②試験さん）。一覧はあとから届く
  setRoutine('①空想ED②試験さん', [director('空想ED', '空想'), member('試験さん', '試験 二郎')]);
  run('reload()');
  page.flushOne();                                   // ルーティンチェックシートだけ先に届く（一覧はもう読んである）
  step('前半：選び直す（ディレクター）', () => {});
  ck(els.mp1.value === '空想 太郎' && els.mp2.value === '試験 二郎', '2) ディレクターの週の①②: ' + J([els.mp1.value, els.mp2.value]));
  ck(/ディレクター：空想 太郎さん/.test(els.mpNote.innerHTML) && !orange(els.mpNote.innerHTML) && !/一致しませんでした/.test(els.mpNote.innerHTML),
     '2) ディレクターの知らせ（名簿に無いという注意は出さない）: ' + els.mpNote.innerHTML);
  step('前半：作る（ディレクター）', () => run('gen()'));
  [, vals, , o] = last();
  ck(J(o.mainPresenters) === J(['空想 太郎', '試験 二郎']), '2) mainPresenters: ' + J(o.mainPresenters));
  ck(vals['メインプレゼン1氏名'] === '空想 太郎' && vals['メインプレゼン1会社名'] === '空想リージョン株式会社'
     && vals['メインプレゼン1カテゴリー'] === '【BNI見本リージョン】' && vals['メインプレゼン2カテゴリー'] === '【工務店】',
     '2) ディレクターの会社名・カテゴリー（その方のページの文字）: ' + J([vals['メインプレゼン1会社名'], vals['メインプレゼン1カテゴリー']]));
  const w0 = o.speakerRotation.weeks[0];
  ck(w0.kind === 'director1' && w0.people.length === 2 && w0.people[0].kind === 'director' && w0.people[0].name === '空想 太郎'
     && w0.people[0].role === 'BNI見本リージョン' && w0.people[0].title === 'BNI見本リージョン' && w0.people[0].inMaster === false
     && w0.people[0].collab === '' && w0.people[0].company === '空想リージョン株式会社' && w0.people[1].name === '試験 二郎' && w0.people[1].kind === 'member',
     '2) 表の1回目（ディレクターの形）: ' + J(w0));
  // ①メンバー・②ディレクターに選び直しても、表はディレクターが先
  step('前半：①②を入れ替える', () => { els.mp1.value = '試験 二郎'; els.mp2.value = '空想 太郎'; run('gen()'); });
  ck(J(last()[3].speakerRotation.weeks[0].people.map((p) => p.name)) === J(['空想 太郎', '試験 二郎']), '2) 表の1回目はディレクターが先: '
     + J(last()[3].speakerRotation.weeks[0].people.map((p) => p.name)));

  // 3) ディレクターだけの週（一覧がまだ届いていないときは、届いてから入る）
  setRoutine('空想ディレクター', [director('空想ディレクター', '空想')]);
  run('guests=null');
  run('reload()');
  page.flushOne();
  ck(els.mp1.value === '' && /読み込んでから入れます/.test(els.mpNote.innerHTML), '3) 一覧が届く前: ' + J([els.mp1.value, els.mpNote.innerHTML]));
  step('前半：一覧を読み直す', () => run('loadGuests()'));
  ck(els.mp1.value === '空想 太郎' && els.mp2.value === '' && /1人の週です/.test(els.mpNote.innerHTML) && /ディレクター：空想 太郎さん/.test(els.mpNote.innerHTML)
     && !orange(els.mpNote.innerHTML), '3) ディレクターだけの週: ' + J([els.mp1.value, els.mp2.value, els.mpNote.innerHTML]));
  step('前半：作る（ディレクターだけ）', () => run('gen()'));
  ck(J(last()[3].mainPresenters) === J(['空想 太郎', '']) && last()[3].speakerRotation.weeks[0].kind === 'director'
     && last()[3].speakerRotation.weeks[0].people.length === 1, '3) ディレクターだけの週に渡す値: ' + J([last()[3].mainPresenters, last()[3].speakerRotation.weeks[0]]));

  // 4) テンプレートにページの無いディレクター → 手で選ぶよう知らせる
  setRoutine('①架空ED', [director('架空ED', '架空')]);
  guestList = GUESTS.slice(1);
  run('guests=null');
  step('前半：ページの無いディレクター', () => { run('reload()'); run('loadGuests()'); });
  ck(els.mp1.value === '' && orange(els.mpNote.innerHTML) && /架空ED は、前半テンプレートのアンバサダー・ディレクターのページに見つかりませんでした/.test(els.mpNote.innerHTML),
     '4) ページの無いディレクターの知らせ: ' + els.mpNote.innerHTML);
  guestList = GUESTS;
  step('前半：一覧を戻す', () => run('loadGuests()'));

  // 5) 名簿の氏名と合わない書き方は、これまでどおり注意
  setRoutine('①見本さん②謎さん', [member('見本さん', '見本 一郎'), { raw: '謎さん', name: '', matched: false, kind: 'member' }]);
  step('前半：合わない書き方', () => run('reload()'));
  ck(orange(els.mpNote.innerHTML) && /謎さん は名簿の氏名と一致しませんでした/.test(els.mpNote.innerHTML), '5) 合わない書き方の注意: ' + els.mpNote.innerHTML);

  // 6) ルーティンチェックシートに無い日：スピーカーローテーションのその回（1人／ディレクター）
  setRoutine('', []);
  rotation = rotWeek([{ name: '空想 太郎', title: 'BNI見本リージョン', collab: '', company: '空想リージョン株式会社', inMaster: false, kind: 'director', role: 'BNI見本リージョン' }]);
  step('前半：ローテーションのディレクターの回', () => run('reload()'));
  ck(els.mp1.value === '空想 太郎' && els.mp2.value === '' && /1人の週です/.test(els.mpNote.innerHTML) && /ディレクターの空想 太郎さん/.test(els.mpNote.innerHTML),
     '6) ローテーションのディレクターの回: ' + J([els.mp1.value, els.mp2.value, els.mpNote.innerHTML]));
  ck(/空想 太郎（BNI見本リージョン）/.test(els.rotPreview.innerHTML), '6) 予定の表示にリージョンの名前が無い: ' + els.rotPreview.innerHTML);
  rotation = rotWeek([{ name: '仮名 四郎', title: 'デザイン', collab: '協業4', company: '仮名デザイン', inMaster: true, kind: 'member' }]);
  step('前半：ローテーションの1人の回', () => run('reload()'));
  ck(els.mp1.value === '仮名 四郎' && els.mp2.value === '' && /1人の週です/.test(els.mpNote.innerHTML), '6) ローテーションの1人の回: ' + J([els.mp1.value, els.mpNote.innerHTML]));
  rotation = rotWeek([MEMBERS[0], MEMBERS[1]].map((m) => ({ name: m.name, title: m.title, collab: m.collab, company: m.company, inMaster: true, kind: 'member' })));
  step('前半：ローテーションの2人の回', () => run('reload()'));
  ck(els.mp1.value === '見本 一郎' && els.mp2.value === '試験 二郎' && /スピーカーローテーションのお2人を入れました/.test(els.mpNote.innerHTML),
     '6) ふだんの2人の回: ' + J([els.mp1.value, els.mp2.value, els.mpNote.innerHTML]));
}

// ---- 後半の画面 ----
{
  setRoutine('①空想ED②試験さん', [director('空想ED', '空想'), member('試験さん', '試験 二郎')]);
  const page = loadPage('slides_meeting_second.html', { server: SERVER, fails });
  const { els, window, run, step } = page;
  page.flush();
  step('後半：開く（ディレクター＋メンバー）', () => window.onload());
  ck(page.log.calls.some((c) => c.name === 'getWeeklyGuests'), '7) 後半でアンバサダー・ディレクターを読んでいない');
  ck(els.mp1.value === '空想 太郎' && els.mp2.value === '試験 二郎', '7) 後半の①②: ' + J([els.mp1.value, els.mp2.value]));
  ck(els.lt1.value === '試験 二郎' && els.lt2.value === '' && /ディレクターは入れていません/.test(els.ltNote.innerHTML),
     '7) 抽選はメンバーの方だけ: ' + J([els.lt1.value, els.lt2.value, els.ltNote.innerHTML]));
  ck(els.lt1.options.some((o) => o.value === '空想 太郎') && els.mp2.options.some((o) => o.value === '仮想 花子'), '7) 後半の欄でディレクターを選べない');
  step('後半：作る', () => run('gen()'));
  let [, vals, , o] = last();
  ck(J(o.mainPresenters) === J(['空想 太郎', '試験 二郎']) && J(o.lottery) === J(['試験 二郎', '']), '7) 後半に渡す①②・抽選: ' + J([o.mainPresenters, o.lottery]));
  ck(vals['メインプレゼン1カテゴリー'] === '【BNI見本リージョン】' && vals['メインプレゼン1会社名'] === '空想リージョン株式会社'
     && vals['抽選1氏名'] === '試験 二郎' && vals['抽選1カテゴリー'] === '【工務店】' && vals['抽選2氏名'] === '',
     '7) 後半の値: ' + J([vals['メインプレゼン1カテゴリー'], vals['抽選1氏名'], vals['抽選2氏名']]));
  setRoutine('①見本さん', [member('見本さん', '見本 一郎')]);
  step('後半：1人の週', () => run('reload()'));
  ck(els.mp1.value === '見本 一郎' && els.mp2.value === '' && els.lt1.value === '見本 一郎' && els.lt2.value === ''
     && /1人の週です/.test(els.mpNote.innerHTML) && !orange(els.mpNote.innerHTML), '8) 後半の1人の週: ' + J([els.mp1.value, els.mp2.value, els.lt1.value, els.mpNote.innerHTML]));
  // 抽選でディレクターを選べば、その方の値が渡る
  step('後半：抽選にディレクターを選ぶ', () => { els.lt2.value = '空想 太郎'; run('gen()'); });
  ck(last()[1]['抽選2カテゴリー'] === '【BNI見本リージョン】' && J(last()[3].lottery) === J(['見本 一郎', '空想 太郎']), '8) 抽選のディレクター: ' + J(last()[3].lottery));
}

// ===================== サーバー =====================
const HEAD = ['No', '業種区分', '氏名', 'ふりがな', 'カテゴリー', '会社名'];
const ROSTER = [['1', '', '見本 一郎', 'みほん いちろう', '税理士', '見本会計事務所'], ['2', '', '試験 二郎', 'しけん じろう', '工務店', '試験工務店']];
const env = makeEnv({ now: new Date(2026, 9, 5, 10, 0, 0) });
const F = Object.assign({}, env.globals);
vm.createContext(F);
for (const f of fs.readdirSync(ROOT).filter((x) => /\.js$/.test(x)).sort()) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), F, { filename: f });
env.reset([['メンバー名簿', false, [HEAD].concat(ROSTER)], ['休会日', true, [['2026/12/30']]]],
          { BNI_CHAPTER: J({ name: '見本', termBase: 23, meetingBaseDate: '2026/03/18', meetingBaseCount: 509 }) });
let SAVED = null;
F.saveOutputFile_ = (blob, name) => { SAVED = { blob, name }; return { id: 'out', url: 'https://example/' + name, downloadUrl: 'https://example/dl/' + name }; };

function png(w, h, v) {
  const row = Buffer.alloc(1 + w * 3, v); row[0] = 0;
  const raw = Buffer.concat(Array.from({ length: h }, () => row));
  const chunk = (type, data) => { const len = Buffer.alloc(4); len.writeUInt32BE(data.length); const td = Buffer.concat([Buffer.from(type), data]), crc = Buffer.alloc(4); crc.writeUInt32BE(zlib.crc32(td) >>> 0); return Buffer.concat([len, td, crc]); };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2;
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}
// 写真：名簿の方は写真フォルダに（ディレクターは無い）
const PHOTOS = {};
['見本 一郎', '試験 二郎'].forEach((n, i) => { const id = 'PH' + i; env.addFile(id, new env.FakeBlob(png(40, 60, 100 + i), 'image/png', n + '.png'), n + '.png'); PHOTOS[F.normName_(n)] = id; });
F.findPhotoIdForName_ = (n) => PHOTOS[F.normName_(n)] || '';

const EMU = 12700;
const tpl = (f) => fs.readFileSync(path.join(ROOT, 'docs', 'templates', f));
const partsOf = (buf) => { const z = readZip(buf), m = {}; Object.keys(z).forEach((k) => { m[k] = new env.FakeBlob(z[k], '', k); }); return m; };
const zipOf = (parts) => { const o = {}; Object.keys(parts).forEach((k) => { o[k] = parts[k]._buf || Buffer.from(parts[k].getDataAsString(), 'utf8'); }); return o; };
const integrity = (z, label) => {
  const f = path.join(os.tmpdir(), 'check_main_presen_' + process.pid + '_' + label + '.pptx');
  fs.writeFileSync(f, writeZip(z));
  const r = spawnSync('python3', [path.join(ROOT, 'tools', 'pptx_integrity.py'), f], { encoding: 'utf8' });
  fs.unlinkSync(f);
  return { ok: r.status === 0, out: (r.stdout || '') + (r.stderr || '') };
};
// 公式ファイルのメインプレゼンテーションのページに似せたページ（題名・写真・「メンバー名／会社名／カテゴリー」）→ お2人のページ
function twoPersonPage(z, prefix, title) {
  let s = z['ppt/slides/slide1.xml'].toString('utf8');
  s = s.replace('name="見出し"', 'name="Title 1"').replace('🎉 熱烈歓迎 🎉', title);
  const para = (t) => '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="ja-JP" altLang="en-US" sz="2800" dirty="0"/><a:t>' + t + '</a:t></a:r></a:p>';
  s = s.replace(/(<p:cNvPr id="5" name="カテゴリー"\/>[\s\S]*?<a:lstStyle\/>)[\s\S]*?(<\/p:txBody>)/, '$1' + para('メンバー名') + para('会社名') + para('カテゴリー') + '$2');
  s = F.removeShape_(F.removeShape_(s, '6'), '7');
  const pg = { xml: s, rels: z['ppt/slides/_rels/slide1.xml.rels'].toString('utf8') };
  return F.offTwoPersonPage_(pg, pg, prefix, title, false);
}
// 文字箱（WEEKLY PRESENTATION のページ）
const box = (id, y, cx, t) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="t${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="4000000" y="${y}"/><a:ext cx="${cx}" cy="600000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ja-JP" sz="2800"/><a:t>${t}</a:t></a:r></a:p></p:txBody></p:sp>`;
const picXml = (id, x, y, cx, cy, rid) => `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="写真${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
const page = (body, hidden) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"${hidden ? ' show="0"' : ''}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${body}</p:spTree></p:cSld></p:sld>`;
const GUEST_MEDIA = 'ppt/media/guest_kuso.jpg';                 // 中身は PNG（大きさを読めるように）。拡張子は後半のひな形に無い jpg
// 前半：見出し → 空想 太郎さんの WEEKLY PRESENTATION（非表示・写真あり）→ 全員終わりましたか？ → メインプレゼンのページ
function firstHalf(prefix) {
  const z = readZip(tpl('BNI_テンプレート_熱烈歓迎.pptx'));
  const two = twoPersonPage(z, prefix || 'メインプレゼン', 'メインプレゼンテーション');
  z['ppt/slides/slide1.xml'] = Buffer.from(two.xml, 'utf8');
  z['ppt/slides/_rels/slide1.xml.rels'] = Buffer.from(two.rels, 'utf8');
  z[GUEST_MEDIA] = png(30, 50, 60);
  const parts = partsOf(writeZip(z));
  const rels = (img) => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
    + (img ? '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/guest_kuso.jpg"/>' : '') + '</Relationships>';
  const a = F.addSlidePart_(parts, page(box(2, 100000, 9000000, 'ウィークリープレゼンテーション')), rels(false));
  const g = F.addSlidePart_(parts, page(box(2, 100000, 9000000, 'WEEKLY PRESENTATION') + box(3, 1500000, 7000000, '空想 太郎')
    + box(4, 2500000, 7000000, '空想リージョン株式会社') + box(5, 3500000, 7000000, 'BNI見本リージョン')
    + picXml(9, 500000, 1500000, 3200000, 3600000, 'rId2'), true), rels(true));
  const d = F.addSlidePart_(parts, page(box(2, 100000, 9000000, '全員終わりましたか？ウィークリープレゼンテーション')), rels(false));
  const main = F.slideEntries_(parts).filter((e) => e.path === 'ppt/slides/slide1.xml')[0];
  F.setSlideEntries_(parts, [a, g, d, main]);
  return writeZip(zipOf(parts));
}
// 後半：事前MTGのひな形（3枚）の3枚目を抽選のページに。deco … 2人のあいだに飾りの画像（まん中に動かすと重なる）。jpg の種類の登録は外す
function secondHalf({ deco = false } = {}) {
  const z = readZip(tpl('BNI_テンプレート_事前MTG.pptx'));
  const w = readZip(tpl('BNI_テンプレート_熱烈歓迎.pptx'));
  const two = twoPersonPage(w, '抽選', '賞品の抽選');
  let xml = two.xml, rels = two.rels;
  if (deco) {
    xml = xml.replace('</p:spTree>', picXml(90, 462 * EMU, 300 * EMU, 36 * EMU, 100 * EMU, 'rIdDeco') + '</p:spTree>');
    rels = rels.replace('</Relationships>', '<Relationship Id="rIdDeco" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/welcome_photo.png"/></Relationships>');
  }
  z['ppt/slides/slide3.xml'] = Buffer.from(xml, 'utf8');
  z['ppt/media/welcome_photo.png'] = w['ppt/media/welcome_photo.png'];
  // レイアウトの関係は、行き先だけ後半のひな形のものに（関係IDは熱烈歓迎のページのまま。重ならないように）
  const lay = (z['ppt/slides/_rels/slide3.xml.rels'].toString('utf8').match(/<Relationship\b[^>]*slideLayout[^>]*Target="([^"]+)"/) || [])[1];
  z['ppt/slides/_rels/slide3.xml.rels'] = Buffer.from(rels.replace(/(<Relationship\b[^>]*slideLayout[^>]*Target=")[^"]+/, '$1' + lay), 'utf8');
  z['[Content_Types].xml'] = Buffer.from(z['[Content_Types].xml'].toString('utf8').replace(/<Default Extension="jpg"[^>]*\/>/, ''), 'utf8');
  return writeZip(z);
}
const roster = { '見本 一郎': ['見本会計事務所', '税理士'], '試験 二郎': ['試験工務店', '工務店'], '空想 太郎': ['空想リージョン株式会社', 'BNI見本リージョン'] };
// 画面（twoPersonValues）と同じ値。noGuest … ディレクターの会社名・カテゴリーを渡さない（サーバーが埋める）
function values(prefix, names, noGuest) {
  const m = { 開催回: '536', 開催日: '2026/10/7(水) 第536回', 開催日付: '2026/10/07' };
  [0, 1].forEach((i) => {
    const n = names[i] || '', r = roster[n] || ['', ''], skip = noGuest && n === '空想 太郎';
    m[prefix + (i + 1) + '氏名'] = n; m[prefix + (i + 1) + '会社名'] = skip ? '' : r[0]; m[prefix + (i + 1) + 'カテゴリー'] = skip || !r[1] ? '' : '【' + r[1] + '】';
  });
  return m;
}
const slideOf = (parts, re) => Object.keys(parts).filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p)).find((p) => re.test(F.xmlOf_(parts, p) || ''));
const shapes = (xml, prefix, n) => F.offShapes_(xml).filter((s) => s.name.indexOf(prefix + n + ' ') === 0);
const centerOf = (list) => { const lo = Math.min(...list.map((s) => s.x)), hi = Math.max(...list.map((s) => s.x + s.cx)); return (lo + hi) / 2; };
const FIRST = firstHalf();
const SLIDE_W = parseInt(F.xmlOf_(partsOf(FIRST), 'ppt/presentation.xml').match(/<p:sldSz\s+cx="(\d+)"/)[1], 10);
const BEFORE = (() => { const p = partsOf(FIRST), x = F.xmlOf_(p, 'ppt/slides/slide1.xml'); return { 1: shapes(x, 'メインプレゼン', 1), 2: shapes(x, 'メインプレゼン', 2) }; })();
const sameY = (a, b) => a.length === b.length && a.every((s, i) => s.y === b[i].y && s.cx === b[i].cx);
const centered = (list) => list.length && Math.abs(centerOf(list) - SLIDE_W / 2) <= 2;
const relTarget = (parts, slide, picShape) => {
  const rid = (picShape.xml.match(/r:embed="([^"]+)"/) || [])[1];
  const rels = F.xmlOf_(parts, F.relsPathOf_(slide)) || '';
  return ((rels.match(new RegExp('<Relationship\\b[^>]*Id="' + rid + '"[^>]*>')) || [''])[0].match(/Target="([^"]+)"/) || [])[1] || '';
};
function runFirst(names, opt) {
  const parts = partsOf(opt && opt.buf || FIRST), map = values('メインプレゼン', names, opt && opt.noGuest);
  const r = F.editMeetingSlides_(parts, map, [], { mainPresenters: names });
  const slide = slideOf(parts, /メインプレゼン[12] 氏名/), xml = F.xmlOf_(parts, slide);
  return { parts, r, slide, xml, s1: shapes(xml, 'メインプレゼン', 1), s2: shapes(xml, 'メインプレゼン', 2), text: F.slideText_(xml), msg: r.photos.message };
}

// 9) 2人の週は前のまま
{
  const t = runFirst(['見本 一郎', '試験 二郎']);
  ck(t.s1.length === 4 && t.s2.length === 4 && Math.abs(centerOf(t.s1) - centerOf(BEFORE[1])) < 1 && Math.abs(centerOf(t.s2) - centerOf(BEFORE[2])) < 1
     && !/1人なので/.test(t.msg) && /見本 一郎/.test(t.text) && /試験 二郎/.test(t.text), '9) 2人の週が変わった: ' + J({ msg: t.msg, n: [t.s1.length, t.s2.length] }));
}
// 10) 1人の週（①だけ・②だけ・1人ぶんだけ渡された）：もう1人の欄を消し、まん中に。上下はそのまま
[[['見本 一郎', ''], 1, 2], [['', '見本 一郎'], 2, 1], [['見本 一郎'], 1, 2]].forEach(([names, keep, gone]) => {
  const t = runFirst(names), kept = keep === 1 ? t.s1 : t.s2, lost = keep === 1 ? t.s2 : t.s1;
  ck(lost.length === 0, '10) ' + J(names) + ' もう1人の欄が残った: ' + J(lost.map((s) => s.name)));
  ck(kept.length === 4 && centered(kept) && sameY(kept, BEFORE[keep]), '10) ' + J(names) + ' 残る方がまん中にない・上下が変わった: '
     + J({ c: centerOf(kept), want: SLIDE_W / 2, n: kept.length }));
  ck(/メインプレゼンは1人なので、まん中に置きました/.test(t.msg) && /見本 一郎/.test(t.text) && /見本会計事務所/.test(t.text) && /【税理士】/.test(t.text) && !/\{\{/.test(t.text),
     '10) ' + J(names) + ' 知らせ・文字: ' + J({ msg: t.msg, text: t.text }));
  const pics = (x) => (x.match(/<p:pic\b/g) || []).length;
  ck(pics(t.xml) === pics(F.xmlOf_(partsOf(FIRST), 'ppt/slides/slide1.xml')) - 1, '10) ' + J(names) + ' もう1人の写真の枠（テンプレートの写真）が残った: '
     + pics(t.xml));
});
// 10b) 1人ずつの後ろの台紙（白いカード）：残る方の台紙は一緒にまん中へ、もう1人の台紙は消す
//      （以前は人だけまん中に動き、2枚の台紙のすき間に乗っていた）。2人にまたがる帯・端の飾りは動かさない。2人の週は前のまま
{
  const z = readZip(FIRST);
  const bx = (list) => ({ x0: Math.min(...list.map((s) => s.x)), x1: Math.max(...list.map((s) => s.x + s.cx)),
                          y0: Math.min(...list.map((s) => s.y)), y1: Math.max(...list.map((s) => s.y + s.cy)) });
  const sp = (id, x, y, cx, cy) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="台紙${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>`
    + `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>`
    + `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:spPr></p:sp>`;
  const card = (id, b) => sp(id, b.x0 - 100000, b.y0 - 100000, b.x1 - b.x0 + 200000, b.y1 - b.y0 + 200000);
  const b1 = bx(BEFORE[1]), b2 = bx(BEFORE[2]);
  // 902 2人にまたがる帯・903 左下の飾り・904 題名の後ろの左の帯（欄のまん中に届かない）・905/906 左右半分ずつの背景（高さいっぱい）
  const extra = sp(905, 0, 0, SLIDE_W / 2, 6858000) + sp(906, SLIDE_W / 2, 0, SLIDE_W / 2, 6858000)
    + card(900, b1) + card(901, b2) + sp(902, 0, b1.y0 + 200000, SLIDE_W, 300000) + sp(903, 100000, 6100000, 900000, 500000)
    + sp(904, 0, 200000, 5600000, 1100000);
  z['ppt/slides/slide1.xml'] = Buffer.from(z['ppt/slides/slide1.xml'].toString('utf8')
    .replace(/<p:grpSpPr\/>|<p:grpSpPr>[\s\S]*?<\/p:grpSpPr>/, (m) => m + extra), 'utf8');
  const buf = writeZip(z), before = F.offShapes_(F.xmlOf_(partsOf(buf), 'ppt/slides/slide1.xml'));
  const geo = (xml, id) => F.offShapes_(xml).find((s) => s.id === id);
  const same = (xml, id) => { const a = geo(xml, id), b = before.find((s) => s.id === id); return a && b && a.x === b.x && a.y === b.y && a.cx === b.cx; };
  [[['見本 一郎', ''], '900', '901'], [['', '見本 一郎'], '901', '900']].forEach(([names, keep, gone]) => {
    const t = runFirst(names, { buf }), kept = geo(t.xml, keep), was = before.find((s) => s.id === keep);
    const person = keep === '900' ? t.s1 : t.s2;
    ck(kept && Math.abs(kept.x + kept.cx / 2 - SLIDE_W / 2) <= 2 && kept.y === was.y && kept.cx === was.cx
       && Math.abs((kept.x + kept.cx / 2) - centerOf(person)) <= 2,
       '10b) ' + J(names) + ' 残る方の台紙が一緒にまん中へ動かない: ' + J({ kept: kept && [kept.x, kept.cx], person: centerOf(person), mid: SLIDE_W / 2 }));
    ck(!geo(t.xml, gone), '10b) ' + J(names) + ' もう1人の台紙が残った');
    ck(['902', '903', '904', '905', '906'].every((id) => same(t.xml, id)),
       '10b) ' + J(names) + ' 2人にまたがる帯・端の飾り・題名の後ろの帯・半分ずつの背景を動かした・消した: '
       + J(['902', '903', '904', '905', '906'].filter((id) => !same(t.xml, id))));
    ck(/メインプレゼンは1人なので、まん中に置きました/.test(t.msg), '10b) ' + J(names) + ' 知らせ: ' + t.msg);
  });
  const t2 = runFirst(['見本 一郎', '試験 二郎'], { buf });
  ck(['900', '901', '902', '903', '904', '905', '906'].every((id) => same(t2.xml, id)), '10b) 2人の週に台紙を動かした');
}
// 11) ディレクター＋メンバー：ディレクターの写真は、前半テンプレートのその方のページの画像
{
  const t = runFirst(['空想 太郎', '試験 二郎']);
  const pic = t.s1.find((s) => s.tag === 'p:pic');
  ck(pic && relTarget(t.parts, t.slide, pic) === '../media/guest_kuso.jpg' && /<a:srcRect\b[^>]*b="\d+"/.test(pic.xml),
     '11) ディレクターの写真がその方のページの画像でない: ' + (pic ? relTarget(t.parts, t.slide, pic) : '枠なし'));
  ck(/ディレクター（空想 太郎さん）の写真は、前半テンプレートのその方のページから入れました/.test(t.msg) && !/写真が見つからない/.test(t.msg),
     '11) ディレクターの写真の知らせ: ' + t.msg);
  ck(/空想リージョン株式会社/.test(t.text) && /【BNI見本リージョン】/.test(t.text) && !/1人なので/.test(t.msg) && Math.abs(centerOf(t.s1) - centerOf(BEFORE[1])) < 1,
     '11) ディレクターの文字・位置: ' + t.text);
  const g = slideOf(t.parts, /WEEKLY PRESENTATION/);
  ck(/guest_kuso\.jpg/.test(F.xmlOf_(t.parts, F.relsPathOf_(g))) && t.parts[GUEST_MEDIA], '11) その方のページの写真がなくなった');
  const it = integrity(zipOf(t.parts), 'first_dir');
  ck(it.ok, '11) 前半のpptxが壊れている: ' + it.out);
}
// 12) ディレクターだけの週（会社名・カテゴリーを画面から渡さない → その方のページの文字）・写真の無い名簿に無い方
{
  const t = runFirst(['空想 太郎', ''], { noGuest: true });
  const pic = t.s1.find((s) => s.tag === 'p:pic');
  ck(t.s2.length === 0 && centered(t.s1) && pic && relTarget(t.parts, t.slide, pic) === '../media/guest_kuso.jpg',
     '12) ディレクターだけの週の置き方・写真: ' + J({ n2: t.s2.length, c: centerOf(t.s1) }));
  ck(/空想リージョン株式会社/.test(t.text) && /【BNI見本リージョン】/.test(t.text) && /まん中に置きました/.test(t.msg) && /ディレクター（空想 太郎さん）/.test(t.msg),
     '12) ディレクターの会社名・カテゴリーを、その方のページから入れていない: ' + J({ text: t.text, msg: t.msg }));
  const it = integrity(zipOf(t.parts), 'first_only');
  ck(it.ok, '12) 前半のpptxが壊れている: ' + it.out);
  const u = runFirst(['未知 太郎', '見本 一郎']);
  ck(/メインプレゼンで写真が見つからない方（写真なし）: 未知 太郎/.test(u.msg) && !/ディレクター（/.test(u.msg), '12) 写真の無い方の知らせにページの名前が無い: ' + u.msg);
}
// 13) 後半：1枚目（前半から写したページ）と抽選。ディレクターの写真は前半のテンプレートから写す（jpg の種類の登録も足す）
env.addFile('TPL1', new env.FakeBlob(FIRST, 'application/vnd.openxmlformats-officedocument.presentationml.presentation', 'first.pptx'), 'first.pptx');
env.props.BNI_TPL_MEETING_FIRST_ID = 'TPL1';
const setSecond = (buf) => {
  env.addFile('TPL2', new env.FakeBlob(buf, 'application/vnd.openxmlformats-officedocument.presentationml.presentation', 'second.pptx'), 'second.pptx');
  env.props.BNI_TPL_MEETING_SECOND_ID = 'TPL2';
};
function runSecond(opts, map) {
  const res = F.generateMeetingSlides('meetingSecond', map, '2026/10/07', opts);
  if (!res.ok) return { res };
  const parts = partsOf(SAVED.blob._buf);
  const main = slideOf(parts, /メインプレゼン[12] 氏名/), lot = slideOf(parts, /抽選[12] 氏名/);
  const mx = main ? F.xmlOf_(parts, main) : '', lx = lot ? F.xmlOf_(parts, lot) : '';
  return { res, parts, main, lot, mx, lx, m1: shapes(mx, 'メインプレゼン', 1), m2: shapes(mx, 'メインプレゼン', 2), l1: shapes(lx, '抽選', 1), l2: shapes(lx, '抽選', 2) };
}
const LOT_BEFORE = (() => { const p = partsOf(secondHalf()), x = F.xmlOf_(p, 'ppt/slides/slide3.xml'); return { 1: shapes(x, '抽選', 1), 2: shapes(x, '抽選', 2) }; })();
{
  setSecond(secondHalf());
  const t = runSecond({ mainPage: true, mainPresenters: ['空想 太郎', ''], lottery: ['', '試験 二郎'] },
                      Object.assign(values('メインプレゼン', ['空想 太郎', '']), values('抽選', ['', '試験 二郎'])));
  ck(t.res.ok, '13) 後半を作れない: ' + t.res.message);
  if (t.res.ok) {
    const pic = t.m1.find((s) => s.tag === 'p:pic'), target = pic ? relTarget(t.parts, t.main, pic) : '';
    const media = 'ppt/' + target.replace(/^\.\.\//, '');
    ck(t.m2.length === 0 && t.m1.length === 4 && centered(t.m1), '13) 後半の1枚目：1人をまん中に: ' + J({ n: [t.m1.length, t.m2.length], c: t.m1.length && centerOf(t.m1) }));
    ck(/^\.\.\/media\/guestphoto\d+\.jpg$/.test(target) && t.parts[media] && t.parts[media]._buf.equals(png(30, 50, 60)),
       '13) 後半の1枚目のディレクターの写真（前半のテンプレートから写す）: ' + target);
    ck(/<Default Extension="jpg"/.test(F.xmlOf_(t.parts, '[Content_Types].xml')), '13) 写した画像の種類（jpg）を登録していない');
    ck(t.l1.length === 0 && t.l2.length === 4 && centered(t.l2) && sameY(t.l2, LOT_BEFORE[2]), '13) 抽選：②だけの1人をまん中に: '
       + J({ n: [t.l1.length, t.l2.length], c: t.l2.length && centerOf(t.l2) }));
    const msg = t.res.message;
    ck(/メインプレゼンは1人なので、まん中に置きました/.test(msg) && /抽選は1人なので、まん中に置きました/.test(msg)
       && /ディレクター（空想 太郎さん）の写真は、前半テンプレートのその方のページから入れました/.test(msg), '13) 後半の知らせ: ' + msg);
    const text = F.slideText_(t.mx) + '|' + F.slideText_(t.lx);
    ck(/空想 太郎/.test(text) && /【BNI見本リージョン】/.test(text) && /試験 二郎/.test(text) && /【工務店】/.test(text) && !/\{\{/.test(text), '13) 後半の文字: ' + text);
    const it = integrity(zipOf(t.parts), 'second');
    ck(it.ok, '13) 後半のpptxが壊れている: ' + it.out);
  }
}
// 14) 抽選の飾りの画像に重なるときは動かさない（もう1人の欄だけ消す）
{
  setSecond(secondHalf({ deco: true }));
  const t = runSecond({ mainPage: false, lottery: ['見本 一郎', ''] }, values('抽選', ['見本 一郎', '']));
  ck(t.res.ok && t.l2.length === 0 && t.l1.length === 4 && Math.abs(centerOf(t.l1) - centerOf(LOT_BEFORE[1])) < 1
     && /抽選は1人なので、もう1人の欄を消しました（まん中に動かすと飾りの画像に重なるので、位置はテンプレートのままです）/.test(t.res.message),
     '14) 飾りの画像に重なるのに動かした: ' + J({ msg: t.res && t.res.message, c: t.l1 && t.l1.length && centerOf(t.l1) }));
  ck(t.res.ok && /写真90/.test(t.lx), '14) 飾りの画像を消した');
}
// 15) 1枚目を入れずに、抽選だけにディレクターを選んだ：前半のテンプレートを開いて、その方の写真を写す
{
  setSecond(secondHalf());
  const t = runSecond({ mainPage: false, lottery: ['空想 太郎', '試験 二郎'] }, values('抽選', ['空想 太郎', '試験 二郎'], true));
  const pic = t.res.ok ? t.l1.find((s) => s.tag === 'p:pic') : null;
  ck(t.res.ok && !t.main && pic && /^\.\.\/media\/guestphoto\d+\.jpg$/.test(relTarget(t.parts, t.lot, pic))
     && /抽選のディレクター（空想 太郎さん）の写真は、前半テンプレートのその方のページから入れました/.test(t.res.message)
     && /【BNI見本リージョン】/.test(F.slideText_(t.lx)), '15) 抽選だけのディレクター: ' + J(t.res.message));
  const it = t.res.ok ? integrity(zipOf(t.parts), 'lottery') : { ok: false, out: '' };
  ck(it.ok, '15) 後半のpptxが壊れている: ' + it.out);
}
// 16) 推薦のことばは1人でも動かさない（2人のページのまま）
{
  const parts = partsOf(firstHalf('推薦のことば'));
  const r = F.editMeetingSlides_(parts, values('推薦のことば', ['見本 一郎', '']), [], { recommenders: ['見本 一郎', ''] });
  const x = F.xmlOf_(parts, 'ppt/slides/slide1.xml'), a = shapes(x, '推薦のことば', 1), b = shapes(x, '推薦のことば', 2);
  ck(a.length === 4 && b.length === 3 && Math.abs(centerOf(a) - centerOf(BEFORE[1])) < 1 && !/1人なので/.test(r.photos.message),
     '16) 推薦のことばを1人の形にした: ' + J({ n: [a.length, b.length], msg: r.photos.message }));
}
// 17) グループの中の図形も、ページの座標でまん中に（グループが半分に縮めているとき）
{
  const sp = (id, x, t) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="8000000"/><a:ext cx="4000000" cy="600000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>${t}</a:t></a:r></a:p></p:txBody></p:sp>`;
  const grp = (body) => `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="50" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="1000000" y="0"/><a:ext cx="4000000" cy="6000000"/><a:chOff x="0" y="0"/><a:chExt cx="8000000" cy="12000000"/></a:xfrm></p:grpSpPr>${body}</p:grpSp>`;
  const xml = page(grp(sp(11, 0, '{{X1氏名}}') + sp(12, 0, '{{X1会社名}}') + sp(13, 0, '{{X1カテゴリー}}') + picXml(14, 1000000, 0, 2000000, 7000000, 'rId2'))
    + sp(21, 7000000, '{{X2氏名}}') + sp(22, 7000000, '{{X2会社名}}') + sp(23, 7000000, '{{X2カテゴリー}}'));
  const plan = F.twoPersonSinglePlan_(xml, 'X', 0, [{ id: '14' }, null], 12000000);
  const out = F.twoPersonSingleApply_(xml, plan), geo = F.shapeBoxes_(out);
  const ids = ['11', '12', '13', '14'], lo = Math.min(...ids.map((i) => geo[i].x)), hi = Math.max(...ids.map((i) => geo[i].x + geo[i].cx));
  ck(plan.moved && Math.abs((lo + hi) / 2 - 6000000) <= 2 && !/\{\{X2/.test(out) && /\{\{X1氏名/.test(out),
     '17) グループの中の図形をまん中にしていない: ' + J({ plan: { dx: plan.dx, moved: plan.moved }, c: (lo + hi) / 2 }));
}

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('メインプレゼンの1人の週・ディレクターの週のスライド: 検査 ' + checks + ' 件 OK: 画面（1人の週は②を（なし）・ディレクターを選べる／'
  + 'ルーティンチェックシートの「空想ED」から入る・会社名とリージョンのカテゴリー・表の1回目の形・ローテーションの回・後半の抽選はメンバーだけ）／'
  + 'サーバー（もう1人の欄を消してまん中に・①でも②でも・後半の1枚目と抽選・飾りに重なるときは動かさない・ディレクターの写真をその方のページから・'
  + '後半は写して種類も登録・会社名とカテゴリーを埋める・写真の無い方の知らせにページの名前・推薦のことばはそのまま・グループの中も・pptxが壊れていない）');

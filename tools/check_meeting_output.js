// 前半スライドのテンプレートから、実際に出力を作ってみる。
// 本番と同じ ooxml.js / routine_srv.js / meeting_slides_srv.js を、
// Driveとスプレッドシートのところだけ差し替えてNodeで動かす。
//
//   node tools/check_meeting_output.js <パーツのディレクトリ> <routine.json> <members.json> <開催日>

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const DIR = process.argv[2];
const ROUTINE = JSON.parse(fs.readFileSync(process.argv[3], 'utf8'));
const MEMBERS = JSON.parse(fs.readFileSync(process.argv[4], 'utf8'));
const DATE = process.argv[5] || '2026/09/30';

// --- シート・Blob を真似る ---
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

function fileBlob(p, type) {
  let name = p;
  return {
    _src: p,
    getDataAsString: () => fs.readFileSync(p, 'utf8'),
    getBytes: () => Array.from(fs.readFileSync(p)).map((b) => (b > 127 ? b - 256 : b)),
    getContentType: () => type || 'image/jpeg',
    setName(n) { name = n; return this; },
    getName: () => name,
  };
}
function memBlob(c, type, n) {
  let name = n;
  return { _mem: c, getDataAsString: () => c, getContentType: () => type,
           setName(x) { name = x; return this; }, getName: () => name };
}

// 写真は、テンプレートに入っている画像を借りて代用する
const media = fs.readdirSync(path.join(DIR, 'ppt/media'))
  .filter((n) => /\.(jpe?g|png)$/i.test(n)).sort();
const PHOTO_OF = {};
MEMBERS.forEach((m, i) => { PHOTO_OF[m.name.replace(/[\s　]/g, '')] = media[i % media.length]; });

const sandbox = {
  console,
  SpreadsheetApp: {}, PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
  CacheService: undefined,
  Utilities: { newBlob: (c, t, n) => memBlob(c, t, n),
               formatDate: (d, tz, f) => {
                 const p = (n) => ('0' + n).slice(-2);
                 return f.replace('yyyy', d.getFullYear()).replace('MM', p(d.getMonth() + 1)).replace('dd', p(d.getDate()));
               } },
  DriveApp: { getFileById: (id) => ({ getBlob: () => fileBlob(path.join(DIR, 'ppt/media', id), /\.png$/i.test(id) ? 'image/png' : 'image/jpeg') }) },
  getSS_: () => ({ getSheets: () => sheets }),
  getMembersList: () => MEMBERS.map((m) => ({ no: m.no, name: m.name })),
  getMemberMaster: () => ({ ok: true, members: MEMBERS }),
  normName_: (s) => String(s == null ? '' : s).replace(/[\s　]/g, ''),
  findPhotoIdForName_: (n) => PHOTO_OF[String(n).replace(/[\s　]/g, '')] || '',
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
  BIG_TEMPLATE_KINDS_: { meetingFirst: { label: '定例会スライド（前半）' } },
};
sandbox.global = sandbox;
vm.createContext(sandbox);
// member_presen_srv.js は写真の取り込み（mpAddPhoto_）を使うため読み込む
for (const f of ['ooxml.js', 'routine_srv.js', 'member_presen_srv.js', 'referral_srv.js',
                 'splice_srv.js', 'meeting_slides_srv.js']) {
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', f), 'utf8'), sandbox, { filename: f });
}
const F = sandbox;

// メンバープレゼンのテンプレート（前半に差し込むときに使う）。展開済みのフォルダを渡す。
const MP_DIR = process.env.MTG_MP_DIR || '';
function dirToMap(dir) {
  const map = {};
  (function walk(d) {
    for (const n of fs.readdirSync(d)) {
      const p = path.join(d, n);
      if (fs.statSync(p).isDirectory()) walk(p);
      else map[path.relative(dir, p).replace(/\\/g, '/')] = fileBlob(p);
    }
  })(dir);
  return map;
}
sandbox.getBigTemplateFile_ = (kind) => ({ getBlob: () => ({ __dir: MP_DIR }) });
sandbox.unzipToMap_ = (blob) => dirToMap(blob.__dir);

// 画面側の組版（slides_layout.html）も読み込む。canvasはNodeに無いので、
// 全角1文字ぶん・半角0.5文字ぶんで測る簡易版に差し替える。
function fakeCtx() {
  let size = 44, bold = false;
  return {
    set font(v) { const m = /(\d+(?:\.\d+)?)px/.exec(v); size = m ? parseFloat(m[1]) : 44; bold = /bold/.test(v); },
    get font() { return `${bold ? 'bold ' : ''}${size}px`; },
    measureText(t) {
      let w = 0;
      for (const ch of String(t)) w += ch.charCodeAt(0) < 128 ? 0.5 : 1.0;
      return { width: w * size };
    },
  };
}
const layoutBox = { console, document: { createElement: () => ({ getContext: () => fakeCtx() }) } };
layoutBox.global = layoutBox;
vm.createContext(layoutBox);
{
  const html = fs.readFileSync(path.join(__dirname, '..', 'slides_layout.html'), 'utf8');
  const js = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map((m) => m[1]).join('\n');
  vm.runInContext(js, layoutBox, { filename: 'slides_layout.html' });
}

// --- ルーティンチェックシートから、その日の決めごとを読む ---
const R = F.getRoutineInfo(DATE);
console.log(`ルーティンチェックシート: ${R.found ? R.sheetName : '(該当なし)'}`);
console.log(`  第${R.meetingNo}回 / コアバリュー=${R.coreValue || '(なし)'} / 一般規定=${R.generalPolicy || '(なし)'}番`);
console.log(`  メインプレゼン: ${(R.mainPresenters || []).map((m) => `${m.raw}→${m.name || '(未一致)'}`).join(' / ') || '(なし)'}`);
console.log(`  求める専門分野: ${JSON.stringify(R.wantedCategories)}`);
console.log(`  開放=${R.openCategory || '(空欄)'} / 審査中=${R.reviewCategory || '(空欄)'}`);

// --- 画面の collect() と同じ形で差し込む値を組み立てる ---
const byName = {};
MEMBERS.forEach((m) => { byName[m.name] = m; });
const map = { 開催回: R.meetingNo, 開催日付: DATE };
const mainNames = (R.mainPresenters || []).map((m) => m.name || '');
const reco = (R.recommendations || [])[0];
const recoNames = reco ? [reco.giver.name || '', reco.receiver.name || ''] : [];
const lotteryNames = mainNames.slice();          // 抽選はメインプレゼンと同じお2人
[['メインプレゼン', mainNames], ['推薦のことば', recoNames], ['抽選', lotteryNames]].forEach(([prefix, names]) => {
  names.forEach((nm, i) => {
    const who = byName[nm] || {};
    map[`${prefix}${i + 1}氏名`] = nm || '';
    map[`${prefix}${i + 1}会社名`] = who.company || '';
    map[`${prefix}${i + 1}カテゴリー`] = who.title ? `【${who.title}】` : '';
  });
});
for (let i = 1; i <= 12; i++) map[`求める専門分野${i}`] = (R.wantedCategories || [])[i - 1] || '';
map['開放カテゴリー'] = R.openCategory || '';
map['審査中カテゴリー'] = R.reviewCategory || '';

// --- パーツを読み込んで書き換える ---
const parts = {};
(function walk(d) {
  for (const n of fs.readdirSync(d)) {
    const p = path.join(d, n);
    if (fs.statSync(p).isDirectory()) walk(p);
    else parts[path.relative(DIR, p).replace(/\\/g, '/')] = fileBlob(p);
  }
})(DIR);

const d = new Date(DATE.replace(/\//g, '-') + 'T00:00:00');
const rules = F.meetingPatternRules_(R.meetingNo, d);
// 音楽：入っているものを一覧し、1つ目の音量だけ変えてみる（差し替えは動作確認用に別途）
const audioList = F.listMeetingAudio_(parts);
const music = {};
if (audioList.length) {
  music[audioList[0].key] = { volume: 25 };
  if (process.env.MTG_MUSIC_ID) music[audioList[0].key].fileId = process.env.MTG_MUSIC_ID;
}
if (audioList.length) {
  console.log('\n入っている音楽・動画:');
  audioList.forEach((a) => console.log(`  ${a.slide.replace(/^.*\//, '')} spid=${a.spid} `
    + `${a.video ? '[動画]' : '[音声]'} ${a.name || a.file}  いまの音量=${a.volume}%`));
}

// --- リファーラル発表（名簿のNo.順・全員ぶん）---
let referral = [];
const rfBoxes = F.rfLayoutBoxes_(parts, F.RF_TITLE_);
if (rfBoxes && process.env.MTG_REFERRAL !== '0') {
  layoutBox.setLayoutBoxes(rfBoxes);
  const list = MEMBERS.slice().sort((a, b) => (parseFloat(a.no) || 9999) - (parseFloat(b.no) || 9999));
  referral = list.map((m, i) => {
    const co = layoutBox.layoutCompany(m.company || '');
    const ca = layoutBox.layoutCategory(m.title || '');
    return { name: m.name, photoName: m.name, seconds: 7, auto: false,   // リファーラルはクリックで進む
             companyLines: co.lines, companyPt: co.fontPt, companyGeom: co.geom, categoryTop: co.categoryTop,
             categoryLines: ca.lines, categoryPt: ca.fontPt, categoryTight: ca.tight,
             nextName: (list[i + 1] ? list[i + 1].name : '') };
  });
  console.log(`\nリファーラル発表: ${referral.length}名ぶん（${referral[0].name} → … → ${referral[referral.length - 1].name}）`);
  console.log(`  ひな形の枠: 会社名 ${JSON.stringify(rfBoxes.companyTall)} カテゴリー上端=${rfBoxes.categoryLow}`);
}

// --- アンバサダー・ディレクターのページ（画面と同じく「リージョン参加者」の名字で選ぶ）---
// MTG_REGION で「リージョン参加者」の記載を差し替えられる（例: MTG_REGION='坂爪アンバサダー'）
let weeklyGuests = null;
const guestPages = F.weeklyGuestPages_(parts);
if (guestPages.length) {
  const raw = process.env.MTG_REGION !== undefined ? process.env.MTG_REGION : (R.regionGuestsRaw || '');
  const flat = raw.replace(/[\s　]/g, '');
  weeklyGuests = guestPages.filter((g) => {
    const ps = g.name.trim().split(/[\s　]+/);
    const sur = ps.length > 1 ? ps[0] : g.name.trim().slice(0, 2);
    return raw ? (!!sur && flat.includes(sur)) : !g.hidden;
  }).map((g) => g.name);
  console.log(`\nアンバサダー・ディレクター: ${guestPages.map((g) => `${g.name}（${g.role}）`).join(' / ')}`);
  console.log(`  リージョン参加者「${raw || '空欄'}」→ 表示: ${weeklyGuests.join('、') || 'なし'}`);
}

// --- メンバープレゼンを前半に差し込む（画面の memberPresenItems() と同じ並び）---
let memberPresen = [];
if (MP_DIR && F.weeklyAnchor_(parts)) {
  const gk = (x) => String(x || '').normalize('NFKC').replace(/[\s・･＆&と]/g, '');
  const ORDER = ['企業サポート', '研修・教育', '不動産関連', '建築・住まい', 'プロモーション',
                 '暮らし・生活', '美容と健康', '飲食・エンタメ'];
  const mm = MEMBERS.map((m) => Object.assign({}, m, { blockKey: gk(m.cat) }));
  const blocks = ORDER.map((b, i) => ({ gkey: gk(b), block: b, order: i + 1,
                                        count: mm.filter((m) => m.blockKey === gk(b)).length }));
  const mctx = { blocks, members: mm, rowsPerPage: 7 };
  const longName = (R.longPresenter || '');
  memberPresen = layoutBox.memberPresenItems(mctx, gk('プロモーション'), longName, true);
  // 後ろにアンバサダー・ディレクターが続くときは、最後の方の「NEXT」にその方を出す（画面の buildMP() と同じ）
  const last = memberPresen[memberPresen.length - 1];
  if (weeklyGuests && weeklyGuests.length && last && last.kind === 'individual' && !last.nextName) {
    last.nextName = weeklyGuests[0];
  }
  const ind = memberPresen.filter((x) => x.kind === 'individual').length;
  console.log(`\nメンバープレゼン: ${memberPresen.length}ページ（扉${memberPresen.length - ind}＋個人${ind}）`
    + `／2分30秒=${longName || 'なし'}／差し込み先=${F.weeklyAnchor_(parts)}`);
}

const info = F.editMeetingSlides_(parts, map, rules, {
  memberPresen: memberPresen,
  weeklyGuests: weeklyGuests,
  weeklyAuto: true,
  referral: referral,
  coreValue: R.coreValue,
  generalPolicy: R.generalPolicy,
  mainPresenters: mainNames.filter((x) => x),
  recommenders: recoNames.filter((x) => x),
  lottery: lotteryNames.filter((x) => x),
  music: music,
});

console.log('\n書き換え結果:');
console.log(`  差し替えたページ: ${info.touched}枚（うち「第○回・日付」${info.byPattern}か所）`);
if (info.core) console.log('  ' + info.core.message);
if (info.policy) console.log('  ' + info.policy.message);
if (info.photos && info.photos.message) console.log('  ' + info.photos.message.split('\n').join('\n  '));
if (info.audio && info.audio.message) console.log('  ' + info.audio.message);
if (info.referral && info.referral.message) console.log('  ' + info.referral.message.split('\n').join('\n  '));
if (info.weekly && info.weekly.message) console.log('  ' + info.weekly.message.split('\n').join('\n  '));
if (info.guests && info.guests.message) console.log('  ' + info.guests.message);

// --- 書き出し ---
const OUT = process.env.MTG_OUT || path.join(DIR, '..', 'out');
fs.rmSync(OUT, { recursive: true, force: true });
fs.mkdirSync(OUT, { recursive: true });
const plan = {};
for (const p of Object.keys(parts)) {
  const b = parts[p];
  if (b._mem !== undefined) {
    const dst = path.join(OUT, 'gen', p);
    fs.mkdirSync(path.dirname(dst), { recursive: true });
    fs.writeFileSync(dst, b._mem, 'utf8');
    plan[p] = { from: 'gen' };
  } else {
    plan[p] = { from: 'src', src: b._src };
  }
}
fs.writeFileSync(path.join(OUT, 'plan.json'), JSON.stringify({ plan, map, info: {
  touched: info.touched, byPattern: info.byPattern,
  core: info.core, policy: info.policy,
  photos: info.photos ? info.photos.message : '',
  audio: info.audio ? info.audio.message : '', audioList: audioList,
  referral: info.referral || null, referralPlan: referral,
  weekly: info.weekly || null, memberPresenPlan: memberPresen,
  guests: info.guests || null, guestPages: guestPages,
} }, null, 1));
console.log(`\n書き出し: ${Object.keys(plan).length} パーツ → ${OUT}`);

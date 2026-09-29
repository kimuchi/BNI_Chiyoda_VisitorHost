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

// 写真は、1人ずつ中身の違う小さなPNGを作って代用する。
// テンプレートの画像を借りると、飾りの画像と同じものが混ざって「誰の写真か」を見分けられないため。
const zlib = require('zlib');
const os = require('os');
function crc32(buf) {
  let c, crc = 0xffffffff;
  for (let n = 0; n < buf.length; n++) {
    c = (crc ^ buf[n]) & 0xff;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crc = (crc >>> 8) ^ c;
  }
  return (crc ^ 0xffffffff) >>> 0;
}
function pngOf(w, h, rgb) {
  const chunk = (type, data) => {
    const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
    const td = Buffer.concat([Buffer.from(type, 'ascii'), data]);
    const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(td));
    return Buffer.concat([len, td, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2; ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;
  const raw = Buffer.alloc((w * 3 + 1) * h);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) raw.set(rgb, y * (w * 3 + 1) + 1 + x * 3);
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr),
                        chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}
const PHOTO_DIR = fs.mkdtempSync(path.join(os.tmpdir(), 'mtg_photos_'));
const PHOTO_OF = {}, PHOTO_FILE = {};
// MTG_NOPHOTO で「写真が無い方」を作れる（名簿の並びの番号をカンマ区切りで。例: MTG_NOPHOTO='4,9'）
const NO_PHOTO = new Set(String(process.env.MTG_NOPHOTO || '').split(',').filter((x) => x !== '').map(Number));
MEMBERS.forEach((m, i) => {
  if (NO_PHOTO.has(i)) return;
  const id = 'photo' + i + '.png';
  fs.writeFileSync(path.join(PHOTO_DIR, id), pngOf(30 + (i % 7), 40 + (i % 5), [(i * 37) % 256, (i * 91) % 256, (i * 53) % 256]));
  PHOTO_OF[m.name.replace(/[\s　]/g, '')] = id;
  PHOTO_FILE[m.name] = path.join(PHOTO_DIR, id);
});

const sandbox = {
  console,
  SpreadsheetApp: {}, PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
  CacheService: undefined,
  Utilities: { newBlob: (c, t, n) => memBlob(c, t, n),
               formatDate: (d, tz, f) => {
                 const p = (n) => ('0' + n).slice(-2);
                 return f.replace('yyyy', d.getFullYear()).replace('MM', p(d.getMonth() + 1)).replace('dd', p(d.getDate()));
               } },
  DriveApp: { getFileById: (id) => ({ getBlob: () => fileBlob(path.join(PHOTO_DIR, id), 'image/png') }) },
  getSS_: () => ({ getSheets: () => sheets }),
  getMembersList: () => MEMBERS.map((m) => ({ no: m.no, name: m.name })),
  // 名簿の「つながりたい人（協業）」は検査用の値（長いものは表で文字が小さくなるか確かめる）
  getMemberMaster: () => ({ ok: true, members: MEMBERS.map((m, i) => Object.assign({
    collab: i % 3 === 0 ? '税理士・社会保険労務士・司法書士・創業支援者・金融機関の融資担当' : '不動産賃貸管理' }, m)) }),
  normName_: (s) => String(s == null ? '' : s).replace(/[\s　]/g, ''),
  findPhotoIdForName_: (n) => PHOTO_OF[String(n).replace(/[\s　]/g, '')] || '',
  // コード.js の fuzzyNameMatch と同じ（routineMemberName_ が、書き間違い・1字違いのときに使う）
  fuzzyNameMatch(searchStr, memberStr) {
    if (searchStr === memberStr) return true;
    if (searchStr.indexOf(memberStr) !== -1 || memberStr.indexOf(searchStr) !== -1) return true;
    if (searchStr.length === memberStr.length && searchStr.length >= 2) {
      let diff = 0;
      for (let i = 0; i < searchStr.length; i++) if (searchStr[i] !== memberStr[i]) diff++;
      if (diff <= 1) return true;
    }
    return false;
  },
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
for (const f of ['ooxml.js', 'chapter_srv.js', 'routine_srv.js', 'member_presen_srv.js', 'referral_srv.js',
                 'splice_srv.js', 'meeting_slides_srv.js', 'speaker_rotation_srv.js', 'role_input_srv.js', 'role_intro_srv.js',
                 'meeting_pages_srv.js']) {
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', f), 'utf8'), sandbox, { filename: f });
}
const F = sandbox;
// 休会日：ルーティンチェックシートで、列はあっても「定例会回数」が空の週（これまでの分）
const HOLIDAYS = [];
for (const n of Object.keys(ROUTINE)) {
  const g = ROUTINE[n];
  (g[0] || []).forEach((v, c) => {
    if (/^\d{4}\/\d{2}\/\d{2}$/.test(String(v)) && String(v) < '2026/10/01' && !String((g[1] || [])[c] || '').trim()) HOLIDAYS.push(String(v));
  });
}
sandbox.getHolidays = () => HOLIDAYS;
// 開催回（2026/3/18 が第509回。休会日の週は数えない）
sandbox.meetingCountOf_ = (d) => {
  const cur = new Date('2026-03-18T00:00:00'); let n = 509;
  const f = (x) => x.getFullYear() + '/' + ('0' + (x.getMonth() + 1)).slice(-2) + '/' + ('0' + x.getDate()).slice(-2);
  for (let i = 0; i < 2000 && cur.getTime() <= d.getTime(); i++) {
    if (cur.getTime() === d.getTime()) return HOLIDAYS.indexOf(f(cur)) < 0 ? n : 0;
    if (HOLIDAYS.indexOf(f(cur)) < 0) n++;
    cur.setDate(cur.getDate() + 7);
  }
  return 0;
};
sandbox.getMeetingCandidates = () => [];

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
const R = F.getRoutineInfo(DATE, { firstHalf: true });
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
const lotteryNames = mainNames.slice();          // 抽選はメインプレゼンと同じお2人
[['メインプレゼン', mainNames], ['抽選', lotteryNames]].forEach(([prefix, names]) => {
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

// 推薦のことば：画面と同じく、定例会中とアフター・定例会後の組に分ける（翌週以降の分は入れない）。
// MTG_RECO で組を差し替えられる（名簿の並びの番号。d=定例会中・a=アフター・定例会後、x=空）
//   例: MTG_RECO='d:0-1,d:2-3,a:4-x'   MTG_RECO='' … 組なし
const personOf = (nm) => {
  const who = byName[nm] || {};
  return { name: nm || '', company: who.company || '', category: who.title ? `【${who.title}】` : '' };
};
let recoPairs;
if (process.env.MTG_RECO !== undefined) {
  const nameAt = (k) => (k === 'x' ? '' : (MEMBERS[parseInt(k, 10)] || {}).name || '');
  recoPairs = process.env.MTG_RECO.split(',').filter((x) => x.trim()).map((x) => {
    const mm = x.trim().match(/^([da]):(\w+)-(\w+)$/);
    if (!mm) throw new Error('MTG_RECO の書き方が違います: ' + x);
    return { giver: personOf(nameAt(mm[2])), receiver: personOf(nameAt(mm[3])), after: mm[1] === 'a' };
  });
} else {
  recoPairs = (R.recommendations || []).filter((x) => x.when !== 'later').map((x) => ({
    giver: personOf(x.giver.name || ''), receiver: personOf(x.receiver.name || ''), after: x.when === 'after' }));
}
recoPairs = recoPairs.filter((p) => p.giver.name || p.receiver.name);   // 2人とも空の組は渡さない（画面と同じ）

// --- パーツを読み込んで書き換える ---
const parts = {};
(function walk(d) {
  for (const n of fs.readdirSync(d)) {
    const p = path.join(d, n);
    if (fs.statSync(p).isDirectory()) walk(p);
    else parts[path.relative(DIR, p).replace(/\\/g, '/')] = fileBlob(p);
  }
})(DIR);

// 後半のテンプレートか（画面では、後半だけが推薦のことば・更新状況一覧を渡す）。
// 前半にも「書記兼会計」の文字はある（役員紹介）ので、後半にしか無いページで見分ける。MTG_KIND=first/second で指定もできる
const SECOND = process.env.MTG_KIND ? process.env.MTG_KIND === 'second'
  : !!(F.findSlideWithText_(parts, 'REFERRAL PRESENTATION') || F.findSlideWithText_(parts, '更新を迎えるメンバー'));
if (SECOND) {
  console.log(`  推薦のことば: ${recoPairs.map((p) => `${p.after ? '[アフター]' : ''}${p.giver.name || '(空)'}→${p.receiver.name || '(空)'}`).join(' / ') || '(なし)'}`);
  // 更新状況一覧：画面では名簿の更新期限日から出す一覧（「○○ ○○さん、…」）。ここでは名簿の並びから作る。
  // 期限切れはわざと多くして、2行に収まるよう文字が小さくなるかを確かめる。MTG_RENEWAL=0 なら渡さない
  const namesOf = (a, b) => MEMBERS.slice(a, b).map((m) => m.name + 'さん').join('、') || '該当者なし';
  if (process.env.MTG_RENEWAL !== '0') {
    map['更新90'] = namesOf(3, 7);
    map['更新60'] = '';                       // 空 → 「該当者なし」
    map['更新30'] = namesOf(10, 11);
    map['更新超過'] = namesOf(12, 26);
  }
}

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
// MTG_REGION で「リージョン参加者」の記載を差し替えられる（例: MTG_REGION='坂上アンバサダー'）
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
    + `／スタートアッププレゼン=${longName || 'なし'}／差し込み先=${F.weeklyAnchor_(parts)}`);
}

// --- スピーカーローテーションの表（前半。画面と同じく、1回目はメインプレゼンのお2人）---
let speakerRotation = null;
if (!SECOND && process.env.MTG_ROTATION !== '0') {
  const rw = F.getSpeakerRotationWeeks(DATE);
  if (rw.ok) {
    speakerRotation = { weeks: rw.weeks, header: rw.header, notes: rw.notes };
    if (mainNames.some((n) => n)) {
      speakerRotation.weeks[0].people = mainNames.map((n) => {
        const who = F.getMemberMaster().members.find((m) => m.name === n) || {};
        return { name: n || '', title: who.title || '', collab: who.collab || '' };
      });
    }
    console.log('\nスピーカーローテーション: ' + speakerRotation.weeks.map((w) => `第${w.no}回 ${w.md} ${w.people.map((p) => p.name).join('・')}${w.source === 'routine' ? '（ルーティン）' : ''}`).join(' ／ '));
  } else console.log('\nスピーカーローテーション: ' + rw.message);
}

// --- 役職のメンバー紹介（前半）：名簿の並びで、架空の担当者・チームを決める（MTG_ROLES=0 なら入れない）---
//   13の役職 … 名簿の1〜13番目の方。ただしメンターコーディネーターは空（1人のページが非表示になる）
//   メンバーシップ委員会 … 14〜16番目の3名（枠が4つあれば、4つ目は空になる）
//   ビジターホスト … 17〜31番目の15名
//   1番目の方にだけ、ふりがな（架空）を付ける → ローマ字
const ROLE_KEYS = ['president', 'vice', 'secretary', 'vhc', 'mentor', 'ec', 'web', 'support', 'training', 'event', 'bcp', 'spreading', 'gbc'];
let roleIntroData = null;
if (!SECOND && process.env.MTG_ROLES !== '0') {
  const holders = {};
  ROLE_KEYS.forEach((k, i) => { holders[k] = k === 'mentor' ? '' : ((MEMBERS[i] || {}).name || ''); });
  const teams = [
    { key: 'membership', name: 'メンバーシップ委員会', members: MEMBERS.slice(13, 16).map((m) => ({ name: m.name })) },
    { key: 'role:vhc', name: 'ビジターホスト', members: MEMBERS.slice(16, 31).map((m) => ({ name: m.name })) },
  ];
  const withKana = MEMBERS.map((m, i) => (i === 0 ? Object.assign({ kana: 'みほん いちろう' }, m) : m));
  roleIntroData = F.riDataFrom_(24, '2026年10月〜2027年3月', holders, 24, teams, true, withKana);
}

// --- 前半：新メンバー・更新メンバー／バイスプレジデントによる報告／ネットワーキングリーダー ---
// 画面と同じく、ルーティンチェックシートの記載（firstHalf）から組み立てる。環境変数で差し替えられる:
//   MTG_NEW='0,1'（名簿の並びの番号）・MTG_RENEW='2:1,3:2'（番号:年数）… '' なら「いない」
//   MTG_NL=on/off（表示するか。既定は、月の最初の定例会か、その日に記載があるとき）
//   MTG_NL_FROM='2026/09/02' … その日の記載を使う（表示にする）
//   MTG_NL_TIE='oto:2,visitor:3' … その部門の受賞者を、名簿の並びの番号の方を足して増やす
//   MTG_WEEK='12' … 速報の「今週のリファーラル」の数
let memberPages = null, vpReport = null, networkingLeaders = null;
if (!SECOND) {
  const FH = R.firstHalf || {};
  const idxList = (env) => process.env[env].split(',').map((x) => x.trim()).filter((x) => x);
  memberPages = {
    newMembers: process.env.MTG_NEW !== undefined
      ? idxList('MTG_NEW').map((i) => ({ name: MEMBERS[+i].name }))
      : (FH.newMembers || []).map((x) => ({ name: x.name, raw: x.raw, category: x.category })),
    renewMembers: process.env.MTG_RENEW !== undefined
      ? idxList('MTG_RENEW').map((x) => { const [i, y] = x.split(':'); return { name: MEMBERS[+i].name, years: +(y || 1) }; })
      : (FH.renewMembers || []).map((x) => ({ name: x.name, raw: x.raw, category: x.category, years: x.years || 1 })),
  };
  vpReport = FH.vpReport ? Object.assign({ weekCount: process.env.MTG_WEEK || '', weekExt: process.env.MTG_WEEK ? '3' : '' }, FH.vpReport) : null;
  let nlSrc = FH.networkingLeaders, nlShow = !!(FH.firstOfMonth || nlSrc);
  if (process.env.MTG_NL_FROM) {
    nlSrc = (F.getRoutineInfo(process.env.MTG_NL_FROM, { firstHalf: true }).firstHalf || {}).networkingLeaders;
    nlShow = true;
  }
  if (process.env.MTG_NL) nlShow = process.env.MTG_NL === 'on';
  const nlItems = nlSrc ? nlSrc.items.map((it) => ({ key: it.key, value: it.value, unit: it.unit,
    winners: it.winners.map((w) => ({ name: w.name, raw: w.raw, category: w.category })) })) : [];
  String(process.env.MTG_NL_TIE || '').split(',').filter((x) => x.trim()).forEach((x) => {
    const [k, i] = x.split(':'), it = nlItems.find((y) => y.key === k);
    if (it) it.winners.push({ name: MEMBERS[+i].name });
  });
  networkingLeaders = { show: nlShow, month: nlSrc ? nlSrc.month : '', items: nlItems };
  console.log(`\n新メンバー: ${memberPages.newMembers.map((x) => x.name || x.raw).join('、') || 'なし'}`
    + ` ／ 更新メンバー: ${memberPages.renewMembers.map((x) => `${x.name || x.raw}（${x.years}年）`).join('、') || 'なし'}`);
  console.log(`バイスプレジデントによる報告: ${vpReport ? JSON.stringify(vpReport) : '(記載なし)'}${FH.vpReportFrom ? '（' + FH.vpReportFrom + ' の記載）' : ''}`);
  console.log(`ネットワーキングリーダー: ${nlShow ? '表示' : '非表示'}（月の最初=${FH.firstOfMonth}）`
    + (nlItems.length ? ' ' + nlItems.map((it) => `${it.key}=${it.value}${it.unit}:${it.winners.map((w) => w.name || w.raw).join('+')}`).join(' / ') : ''));
}

const info = F.editMeetingSlides_(parts, map, rules, {
  memberPresen: memberPresen,
  weeklyGuests: weeklyGuests,
  weeklyAuto: true,
  referral: referral,
  coreValue: R.coreValue,
  generalPolicy: R.generalPolicy,
  // 画面と同じく、並び（左・右）はそのまま渡す（名簿に無い方は空）
  mainPresenters: mainNames,
  recommendPairs: SECOND ? recoPairs : undefined,
  speakerRotation: speakerRotation,
  lottery: lotteryNames,
  music: music,
  roleIntro: !!roleIntroData,
  roleIntroData: roleIntroData,
  memberPages: memberPages,
  vpReport: vpReport,
  networkingLeaders: networkingLeaders,
  meetingDate: d,
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
if (info.reco && info.reco.message) console.log('  ' + info.reco.message.split('\n').join('\n  '));
if (info.renewal && info.renewal.message) console.log('  ' + info.renewal.message);
if (info.rotation && info.rotation.message) console.log('  ' + info.rotation.message);
if (info.roles && info.roles.message) console.log('  ' + info.roles.message.split('\n').join('\n  '));
if (info.members && info.members.message) console.log('  ' + info.members.message.split('\n').join('\n  '));
if (info.vp && info.vp.message) console.log('  ' + info.vp.message);
if (info.leaders && info.leaders.message) console.log('  ' + info.leaders.message.split('\n').join('\n  '));

// --- 役職のメンバー紹介の中身を確かめる ---
// 入れたお名前がそのページにあること・写真の枠の写真がその方のものであること・写真の無い方の枠は外したこと・
// 差し込み口が残っていないこと・空の役職の1人のページが非表示になったこと・ローマ字
let roleChecks = 0;
const roleFails = [];
if (roleIntroData && info.roles) {
  const rck = (ok, msg) => { roleChecks++; if (!ok) roleFails.push(msg); };
  const txt = (p) => F.slideText_(F.xmlOf_(parts, p) || '');
  const filled = info.roles.filled || [];
  rck(filled.length > 0, '役職のメンバー紹介：入れたところが無い');
  filled.forEach((f) => {
    if (f.name) rck(txt(f.path).includes(f.name), `${f.path}：${f.base}に「${f.name}」が無い`);
  });
  const vals = F.riValues_(roleIntroData);
  const noPhoto = new Set(MEMBERS.filter((m, i) => NO_PHOTO.has(i)).map((m) => m.name));
  for (const p of F.slideOrder_(parts)) {
    const x = F.xmlOf_(parts, p) || '', t = F.slideText_(x);
    const left = [...t.matchAll(/\{\{([^{}]{1,60})\}\}/g)].map((m) => m[1]).filter((k) => k in vals.map);
    rck(!left.length, p + '：役職の差し込み口が残っている: ' + left.join('、'));
    const rels = F.xmlOf_(parts, p.replace(/^(ppt\/slides\/)(slide\d+\.xml)$/, '$1_rels/$2.rels')) || '';
    for (const m of x.matchAll(/<p:pic>[\s\S]*?<\/p:pic>/g)) {
      const nm = (m[0].match(/<p:cNvPr\b[^>]*\sname="\{\{([^"{}]+)写真\}\}"/) || [])[1];
      if (!nm) continue;
      const who = vals.persons[nm];
      rck(!!who && !noPhoto.has(who.name), `${p}：${nm}の写真の枠が残っている（入る方がいない・写真が無い）`);
      if (!who) continue;
      const rid = (m[0].match(/r:embed="(rId\d+)"/) || [])[1];
      const rel = (rels.match(new RegExp('<Relationship\\b[^>]*\\bId="' + rid + '"[^>]*>')) || [''])[0];
      const tgt = ((rel.match(/Target="([^"]+)"/) || [])[1] || '').replace('../', 'ppt/');
      const blob = parts[tgt];
      rck(blob && blob._src === PHOTO_FILE[who.name], `${p}：${nm}（${who.name}）の写真が違う: ${tgt}`);
    }
  }
  // 写真の無い方（MTG_NOPHOTO）は、写真の枠を外して知らせる
  filled.filter((f) => noPhoto.has(f.name) && vals.persons[f.base]).forEach((f) => {
    rck(info.roles.noPhoto.includes(f.name) || !(F.xmlOf_(parts, f.path) || '').includes('{{' + f.base + '写真}}'),
        `${f.name}さん（写真なし）が知らせに無い`);
  });
  // 担当者の居ない役職（メンターコーディネーター）の1人のページは非表示
  if (info.roles.hidden.includes('メンターコーディネーター')) {
    rck(F.slideOrder_(parts).some((p) => /メンターコーディネーター/.test(txt(p)) && /<p:sld\b[^>]*\sshow="0"/.test(F.xmlOf_(parts, p))),
        'メンターコーディネーターのページが非表示になっていない');
  }
  // ローマ字（1人目の方：みほん いちろう → ICHIRO MIHON）。英語の役職名のある1人のページがある雛形だけ
  rck(roleIntroData.roles.president.romaji === 'ICHIRO MIHON', 'ローマ字: ' + roleIntroData.roles.president.romaji);
  const presSingle = filled.find((f) => f.base === 'プレジデント' && /President/.test(txt(f.path)));
  if (presSingle) rck(/ICHIRO MIHON/.test(txt(presSingle.path)), 'プレジデントの1人のページにローマ字が入っていない');
  console.log(`  役職のメンバー紹介の中身: 検査 ${roleChecks} 件` + (roleFails.length ? `　NG ${roleFails.length} 件` : '　OK'));
  roleFails.slice(0, 20).forEach((f) => console.log('    NG ' + f));
}

// --- 新メンバー・更新メンバー／バイスプレジデントによる報告／ネットワーキングリーダーの中身を確かめる ---
let pageChecks = 0;
const pageFails = [];
if (!SECOND) {
  const pck = (ok, msg) => { pageChecks++; if (!ok) pageFails.push(msg); };
  const order = F.slideOrder_(parts), sz = F.fpSize_(parts);
  const txt = (p) => F.slideText_(F.xmlOf_(parts, p) || '');
  const flat = (p) => txt(p).normalize('NFKC').replace(/\s/g, '');
  const shown = (p) => !/<p:sld\b[^>]*\sshow="0"/.test(F.xmlOf_(parts, p) || '');
  const noPhoto = new Set(MEMBERS.filter((m, i) => NO_PHOTO.has(i)).map((m) => m.name));
  // 写真の枠 picId の画像が、その方の写真か（写真の無い方は枠が隠れているか）
  const photoOk = (p, picId, name) => {
    const x = F.xmlOf_(parts, p), r = F.findShapeRange_(x, picId);
    if (!r) return false;
    const seg = x.substring(r.start, r.end);
    if (noPhoto.has(name) || !PHOTO_FILE[name]) return /<p:cNvPr\b[^>]*\shidden="1"/.test(seg);
    const rels = F.xmlOf_(parts, p.replace(/^(ppt\/slides\/)(slide\d+\.xml)$/, '$1_rels/$2.rels')) || '';
    const rid = (seg.match(/r:embed="(rId\d+)"/) || [])[1];
    const rel = (rels.match(new RegExp('<Relationship\\b[^>]*\\bId="' + rid + '"[^>]*>')) || [''])[0];
    const tgt = ((rel.match(/Target="([^"]+)"/) || [])[1] || '').replace('../', 'ppt/');
    return !!parts[tgt] && parts[tgt]._src === PHOTO_FILE[name] && !/<p:cNvPr\b[^>]*\shidden="1"/.test(seg);
  };
  // 新メンバー・更新メンバー：1人1枚。表示のページの数と、お名前・写真・年数
  if (memberPages) {
    [['new', '新メンバー', memberPages.newMembers], ['renew', '更新メンバー', memberPages.renewMembers]].forEach(([k, label, list]) => {
      const pages = order.filter((p) => F.fpMemberKind_(F.xmlOf_(parts, p)) === k && F.fpMemberUnit_(F.xmlOf_(parts, p), sz.W, sz.H));
      const vis = pages.filter(shown);
      pck(vis.length === list.length, `${label}：表示のページが ${vis.length}枚（${list.length}名のはず）`);
      list.forEach((x, i) => {
        const p = vis[i];
        if (!p) return;
        const nm = x.name || x.raw;
        pck(txt(p).includes(nm), `${label}：${p} に ${nm} が無い`);
        const who = byName[nm];
        if (who && who.title) pck(flat(p).includes(`【${who.title}】`.normalize('NFKC').replace(/\s/g, '')), `${label}：${p} に ${nm} のカテゴリーが無い`);
        if (who && who.company) pck(flat(p).includes(who.company.normalize('NFKC').replace(/\s/g, '')), `${label}：${p} に ${nm} の会社名が無い`);
        if (k === 'renew') pck(flat(p).includes(`${x.years || 1}年更新`), `${label}：${p} に「${x.years || 1}年更新」が無い`);
        const u = F.fpMemberUnit_(F.xmlOf_(parts, p), sz.W, sz.H);
        if (u && u.photo && who) pck(photoOk(p, u.photo.id, nm), `${label}：${p} の写真が ${nm} さんのものでない`);
        if (noPhoto.has(nm)) pck(!/トリミング/.test(txt(p)), `${label}：${p}（写真の無い方）に見本の文字が残っている`);
      });
    });
  }
  // 倫理規定：新メンバー・更新メンバーのまとまりの、それぞれすぐうしろに1枚。人がいるまとまりのうしろだけ表示
  if (memberPages) {
    const isEth = (p) => /倫理規定/.test(flat(p));
    [['new', '新メンバー', memberPages.newMembers], ['renew', '更新メンバー', memberPages.renewMembers]].forEach(([k, label, list]) => {
      const pages = order.filter((p) => F.fpMemberKind_(F.xmlOf_(parts, p)) === k && F.fpMemberUnit_(F.xmlOf_(parts, p), sz.W, sz.H));
      if (!pages.length) return;
      const next = order[Math.max(...pages.map((p) => order.indexOf(p))) + 1];
      pck(next && isEth(next), `倫理規定：${label}のすぐうしろに倫理規定のページが無い`);
      if (next && isEth(next)) pck(shown(next) === (list.length > 0), `倫理規定：${label}のあとの倫理規定が${shown(next) ? '表示' : '非表示'}（${list.length}名）`);
    });
    const vis = order.filter(shown);
    vis.forEach((p, i) => { if (i && isEth(p)) pck(!isEth(vis[i - 1]), `倫理規定：表示のページが2枚続いている（${vis[i - 1]}・${p}）`); });
    pck((info.members.ethics || []).length >= 1, '倫理規定：知らせが無い');
    // 並び：新メンバー → 倫理規定 → 更新メンバー → 倫理規定（雛形が「更新メンバー → 新メンバー」の並びでも）
    const seqOk = (q, list) => {
      const unit = (p) => { const x = F.xmlOf_(q, p) || '', k = F.fpMemberKind_(x); return k && k !== 'both' && F.fpMemberUnit_(x, sz.W, sz.H) ? k : ''; };
      const ks = list.map(unit), iN = ks.indexOf('new'), lN = ks.lastIndexOf('new'), iR = ks.indexOf('renew'), lR = ks.lastIndexOf('renew');
      const eth = (i) => i >= 0 && i < list.length && /倫理規定/.test(F.slideText_(F.xmlOf_(q, list[i]) || ''));
      return iN >= 0 && iR >= 0 && eth(lN + 1) && lN + 2 === iR && eth(lR + 1);
    };
    if (memberPages.newMembers.length && memberPages.renewMembers.length) {
      pck(seqOk(parts, vis), '新メンバー → 倫理規定 → 更新メンバー → 倫理規定 の並びでない');
    }
    // 雛形の並びが違っても（倫理規定がまとまりの前・あいだに別のページ・2枚目の写しがある）、
    // 表示の倫理規定は「人がいるまとまり」の数だけで、続けて出ない。4つの並び × 4つの人数で、雛形を読み直して作る
    const fresh = () => {
      const q = {};
      (function walk(d) {
        for (const n of fs.readdirSync(d)) {
          const x = path.join(d, n);
          if (fs.statSync(x).isDirectory()) walk(x); else q[path.relative(DIR, x).replace(/\\/g, '/')] = fileBlob(x);
        }
      })(DIR);
      return q;
    };
    const t0 = fresh(), o0 = F.slideOrder_(t0), ethOf = (q, p) => /倫理規定/.test(F.slideText_(F.xmlOf_(q, p) || ''));
    const blk = (k) => o0.filter((p) => F.fpMemberKind_(F.xmlOf_(t0, p) || '') === k && F.fpMemberUnit_(F.xmlOf_(t0, p), sz.W, sz.H));
    const R0 = blk('renew'), N0 = blk('new'), E0 = o0.find((p) => ethOf(t0, p));
    if (R0.length && N0.length && E0) {
      const reorder = (q, fn) => {
        const prs = F.xmlOf_(q, 'ppt/presentation.xml'), ids = prs.match(/<p:sldId\b[^>]*\/>/g), o = F.slideOrder_(q);
        const out = fn(ids.map((x, i) => ({ x, p: o[i] })));
        F.putXml_(q, 'ppt/presentation.xml', prs.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, '<p:sldIdLst>' + out.map((y) => y.x).join('') + '</p:sldIdLst>'));
      };
      const move = (arr, p, before) => {
        const [it] = arr.splice(arr.findIndex((y) => y.p === p), 1);
        arr.splice(before ? arr.findIndex((y) => y.p === before) : arr.length, 0, it);
        return arr;
      };
      const other = o0.find((p) => !ethOf(t0, p) && !R0.includes(p) && !N0.includes(p) && o0.indexOf(p) > o0.indexOf(E0) + 1) || o0[0];
      const LAYOUTS = {
        'いつもの並び': null,
        '倫理規定がまとまりの前': (q) => reorder(q, (a) => move(a, E0, R0[0])),
        '新メンバーと倫理規定のあいだに別のページ': (q) => reorder(q, (a) => move(a, other, E0)),
        '倫理規定の写しが2枚': (q) => F.fpClonePages_(q, E0, 1, E0),
      };
      const ppl = (n) => Array.from({ length: n }, (_, i) => ({ name: MEMBERS[i].name, raw: MEMBERS[i].name, years: 1 }));
      const LISTS = { 両方: [1, 1], 更新だけ: [1, 0], 新だけ: [0, 1], だれもいない: [0, 0] };
      Object.entries(LAYOUTS).forEach(([ln, fn]) => Object.entries(LISTS).forEach(([cn, [r, n]]) => {
        const q = fresh();
        if (fn) fn(q);
        F.applyMemberPages_(q, { renewMembers: ppl(r), newMembers: ppl(n) }, { by: {}, seq: 0 });
        const vis2 = F.slideOrder_(q).filter((p) => !/<p:sld\b[^>]*\sshow="0"/.test(F.xmlOf_(q, p) || ''));
        const ve = vis2.filter((p) => ethOf(q, p));
        pck(ve.length === r + n, `倫理規定（${ln}・${cn}）：表示の倫理規定が ${ve.length}枚（${r + n}枚のはず）`);
        vis2.forEach((p, i) => { if (i && ethOf(q, p)) pck(!ethOf(q, vis2[i - 1]), `倫理規定（${ln}・${cn}）：表示のページが2枚続いている`); });
        if (r && n) pck(seqOk(q, vis2), `倫理規定（${ln}・${cn}）：新メンバー → 倫理規定 → 更新メンバー → 倫理規定 の並びでない`);
      }));
    }
  }
  // バイスプレジデントによる報告：数字と速報の日付
  if (vpReport) {
    const vpPages = order.filter((p) => /バイスプレジデント/.test(txt(p)) && /報告/.test(txt(p)));
    const all = vpPages.map(flat).join('|');
    const want = [vpReport.avg && vpReport.avg + '件', vpReport.count && vpReport.count + '件', vpReport.total && vpReport.total + '件', vpReport.thanks,
                  vpReport.perWeek && '(' + vpReport.perWeek + '件/週)'].filter(Boolean);
    want.forEach((w) => pck(all.includes(w.normalize('NFKC')), `バイスプレジデントによる報告に「${w}」が無い`));
    if (vpReport.month) {
      const [y, m] = vpReport.month.split('-').map(Number);
      pck(all.includes(`${y}年${m}月の月間リファーラル数`), `バイスプレジデントによる報告の年月（${y}年${m}月）が無い`);
    }
    if (vpReport.from && vpReport.to) {
      const [y1, m1] = vpReport.from.split('-').map(Number), [y2, m2] = vpReport.to.split('-').map(Number);
      pck(all.includes(`${y1}年${m1}月から${y2}年${m2}月`), `累計の期間（${y1}年${m1}月から${y2}年${m2}月）が無い`);
    }
    if (vpPages.some((p) => /速報/.test(txt(p)))) pck(all.includes(`${d.getMonth() + 1}/${d.getDate()}速報`), '速報の日付が開催日でない');
    if (vpReport.weekCount) pck(all.includes(`今週のリファーラル${vpReport.weekCount}件`), '速報の今週のリファーラルの数が無い');
  }
  // ネットワーキングリーダー：月の最初の定例会だけ表示。部門のページ・まとめのページに受賞者
  if (networkingLeaders) {
    const nlPages = order.filter((p) => /ネットワーキングリーダー/.test(flat(p)));
    if (!networkingLeaders.show) nlPages.forEach((p) => pck(!shown(p), `ネットワーキングリーダー：${p} が表示のまま（非表示のはず）`));
    else {
      const vis = nlPages.filter(shown);
      pck(vis.length > 0, 'ネットワーキングリーダー：表示のページが無い');
      const filled = (info.leaders && info.leaders.filled) || [];
      networkingLeaders.items.forEach((it) => {
        const names = it.winners.map((w) => (byName[w.name] ? w.name : '')).filter(Boolean);
        if (!names.length) return;
        const pages = vis.filter((p) => { const inf = F.fpNlPage_(F.xmlOf_(parts, p), sz.W, sz.H); return inf.type === 'kind' && inf.kind === it.key; });
        pck(pages.length >= 1, `ネットワーキングリーダー：${it.key} のページが表示されていない`);
        names.forEach((nm) => {
          const onKind = pages.filter((p) => txt(p).includes(nm));
          pck(onKind.length === 1, `ネットワーキングリーダー：${it.key} の ${nm} のページが ${onKind.length}枚`);
          onKind.forEach((p) => {
            const inf = F.fpNlPage_(F.xmlOf_(parts, p), sz.W, sz.H), u = inf.units.find((x) => F.slideText_(F.xmlOf_(parts, p).substring(F.findShapeRange_(F.xmlOf_(parts, p), x.name.id).start, F.findShapeRange_(F.xmlOf_(parts, p), x.name.id).end)).includes(nm));
            pck(!!u, `ネットワーキングリーダー：${p} の ${nm} の枠が見つからない`);
            if (u && u.photo) pck(photoOk(p, u.photo.id, nm), `ネットワーキングリーダー：${p} の写真が ${nm} さんのものでない`);
            if (inf.value) pck(flat(p).includes(String(it.value).normalize('NFKC').replace(/\s/g, '').replace(/(万円|円)$/, '')), `ネットワーキングリーダー：${p} に数 ${it.value} が無い`);
          });
          const sum = vis.filter((p) => F.fpNlPage_(F.xmlOf_(parts, p), sz.W, sz.H).type === 'summary');
          sum.forEach((p) => pck(txt(p).includes(nm), `ネットワーキングリーダー：まとめのページに ${nm} が無い`));
        });
        if (names.length === 2) pck(pages.some((p) => names.every((nm) => txt(p).includes(nm))), `ネットワーキングリーダー：${it.key} のお2人のページが無い`);
      });
      pck(filled.length > 0, 'ネットワーキングリーダー：入れたところが無い');
      if (networkingLeaders.month) {
        const m = +networkingLeaders.month.split('-')[1], fw = String(m).replace(/[0-9]/g, (c) => String.fromCharCode(c.charCodeAt(0) + 0xFEE0));
        pck(vis.some((p) => new RegExp(`(${m}|${fw})月度`).test(txt(p).replace(/\s/g, ''))) || !nlPages.some((p) => /月度/.test(txt(p))), 'ネットワーキングリーダー：見出しの「○月度」が対象の月でない');
      }
    }
  }
  console.log(`  新メンバー・更新メンバー／バイス報告／ネットワーキングリーダー: 検査 ${pageChecks} 件` + (pageFails.length ? `　NG ${pageFails.length} 件` : '　OK'));
  pageFails.slice(0, 30).forEach((f) => console.log('    NG ' + f));
}

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
  // 写真の取り違いを確かめるため：お名前 → 代わりの写真のファイル、ページごとに入るはずのお2人
  photoOf: PHOTO_FILE,
  twoPerson: { メインプレゼン: mainNames, 抽選: lotteryNames },
  // 推薦のことば（組ごとのページ）と、書記兼会計による報告（更新状況一覧）
  recoPairs: SECOND ? recoPairs : null, reco: info.reco || null,
  renewal: info.renewal || null,
  // スピーカーローテーションの表（前半）
  rotation: info.rotation || null, rotationPlan: speakerRotation,
  // 役職のメンバー紹介（前半）
  roles: info.roles ? { message: info.roles.message, hidden: info.roles.hidden, filled: info.roles.filled } : null,
  // 新メンバー・更新メンバー／バイスプレジデントによる報告／ネットワーキングリーダー（前半）
  members: info.members || null, vp: info.vp || null, leaders: info.leaders || null,
  memberPages: memberPages, vpReport: vpReport, networkingLeaders: networkingLeaders,
} }, null, 1));
console.log(`\n書き出し: ${Object.keys(plan).length} パーツ → ${OUT}`);
if (roleFails.length || pageFails.length) process.exitCode = 1;

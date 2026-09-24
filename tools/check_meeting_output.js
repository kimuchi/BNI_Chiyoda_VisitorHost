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
for (const f of ['ooxml.js', 'routine_srv.js', 'member_presen_srv.js', 'meeting_slides_srv.js']) {
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', f), 'utf8'), sandbox, { filename: f });
}
const F = sandbox;

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
(R.mainPresenters || []).forEach((mp, i) => {
  const who = byName[mp.name] || {};
  map[`メインプレゼン${i + 1}氏名`] = mp.name || '';
  map[`メインプレゼン${i + 1}会社名`] = who.company || '';
  map[`メインプレゼン${i + 1}カテゴリー`] = who.title ? `【${who.title}】` : '';
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
const info = F.editMeetingSlides_(parts, map, rules, {
  coreValue: R.coreValue,
  generalPolicy: R.generalPolicy,
  mainPresenters: (R.mainPresenters || []).map((m) => m.name).filter((x) => x),
});

console.log('\n書き換え結果:');
console.log(`  差し替えたページ: ${info.touched}枚（うち「第○回・日付」${info.byPattern}か所）`);
if (info.core) console.log('  ' + info.core.message);
if (info.policy) console.log('  ' + info.policy.message);
if (info.photos && info.photos.message) console.log('  ' + info.photos.message);

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
} }, null, 1));
console.log(`\n書き出し: ${Object.keys(plan).length} パーツ → ${OUT}`);

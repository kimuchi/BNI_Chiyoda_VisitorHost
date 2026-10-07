// 公式ファイルから雛形を作り、その雛形でこのシステムの作成処理を実際に動かして確かめる。
//
//   node tools/check_official.js <公式ファイル.pptx> [出力先のフォルダ]
//
// 公式ファイル（BNIメンバーがダウンロードできる定例会のスライド）はリポジトリに入れていない。
// 手元にダウンロードしたものを渡す。Apps Script の API は tools/lib_gas_fake.js で真似る
// （Drive API の範囲指定のダウンロードも、手元のファイルから返す。50MBを超える Blob は止める）。
//
// 確かめること
//   ・7つの雛形が作れて、部品のつながりが正しい（tools/pptx_integrity.py）。1つあたり50MBより十分小さい
//   ・ビジター紹介・ゲスト・代理：3人1枚で作れる。「専門分野：」「招待者：」の見出しが残り、空き枠は消える
//   ・ビジタープレゼン：カウントダウンがビジタープレゼンの秒数。卵時計の動画（音）が残る
//   ・メンバープレゼン：扉と個人のページ。枠の寸法をテンプレートから採る。スタートアッププレゼンの方は長い秒数。
//     自動で次へ進む時間は、卵時計のぶん遅れて始まるカウントダウンの終わり＋1秒
//   ・定例会（前半）：表紙の第○回・日付、メンバーシップ委員会の表、メインプレゼンのお2人、メンバーのページの差し込み、
//     スピーカーローテーションの表
//   ・定例会（後半）：リファーラル発表のページ（人数ぶん・リファーラル発表の秒数・クリックで始まる・次の発表者）、
//     推薦のことば（2組）、書記兼会計の更新状況の表

const fs = require('fs');
const os = require('os');
const path = require('path');
const vm = require('vm');
const { execFileSync } = require('child_process');
const { makeGas, FakeBlob } = require('./lib_gas_fake');
const { readZip } = require('./lib_zip');

const ROOT = path.join(__dirname, '..');
const SRC = process.argv[2];
if (!SRC || !fs.existsSync(SRC)) {
  console.log('使い方: node tools/check_official.js <公式ファイル.pptx> [出力先のフォルダ]');
  process.exit(2);
}
const OUT = process.argv[3] || fs.mkdtempSync(path.join(os.tmpdir(), 'official_'));
fs.mkdirSync(OUT, { recursive: true });
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }
const J = (x) => JSON.stringify(x);

// --- 写真（名前 → 手元の画像）。縦長の小さなPNGを作って使う ---
function png(w, h) {
  const zlib = require('zlib');
  const row = Buffer.alloc(1 + w * 3, 200); row[0] = 0;
  const raw = Buffer.concat(Array.from({ length: h }, () => row));
  const chunk = (type, data) => {
    const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
    const td = Buffer.concat([Buffer.from(type), data]), crc = Buffer.alloc(4);
    crc.writeUInt32BE(zlib.crc32(td) >>> 0);
    return Buffer.concat([len, td, crc]);
  };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2;
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr),
                        chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}
const photoFile = path.join(OUT, '_photo.png');
fs.writeFileSync(photoFile, png(40, 60));

const FILES = { OFFICIAL: SRC, PHOTO1: photoFile, PHOTO2: photoFile };
const gas = makeGas({ files: FILES });
const PHOTOS = { '見本一郎': 'PHOTO1', '見本花子': 'PHOTO2' };
let mpBlob = null;
const sb = Object.assign({}, gas.globals, {
  findPhotoIdForName_: (n) => PHOTOS[String(n || '').replace(/[\s　]/g, '')] || '',
  getBigTemplateFile_: (kind) => ({ getBlob: () => mpBlob, getName: () => kind, getId: () => kind }),
});
vm.createContext(sb);
for (const f of ['ooxml.js', 'chapter_srv.js', 'assets.js', 'big_templates_srv.js', 'member_presen_srv.js', 'referral_srv.js',
                 'splice_srv.js', 'meeting_slides_srv.js', 'speaker_rotation_srv.js', 'slides_visitor_srv.js',
                 'role_input_srv.js', 'role_intro_srv.js', 'official_srv.js', 'official_build_srv.js', 'routine_srv.js', 'meeting_pages_srv.js']) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), sb, { filename: f });
}
sb.getBigTemplateFile_ = (kind) => ({ getBlob: () => mpBlob, getName: () => kind, getId: () => kind });
sb.findPhotoIdForName_ = (n) => PHOTOS[String(n || '').replace(/[\s　]/g, '')] || '';

// 画面の組版（slides_layout.html）。canvas の代わりに、全角1文字・半角0.5文字で測る
const layout = (() => {
  const html = fs.readFileSync(path.join(ROOT, 'slides_layout.html'), 'utf8');
  const js = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map((m) => m[1]).join('\n');
  const ctx = () => { let size = 44; return { set font(v) { const m = /(\d+(?:\.\d+)?)px/.exec(v); size = m ? +m[1] : 44; },
    measureText(t) { let w = 0; for (const ch of String(t)) w += ch.charCodeAt(0) < 128 ? 0.5 : 1; return { width: w * size }; } }; };
  const L = { console, document: { createElement: () => ({ getContext: ctx }) } };
  vm.createContext(L);
  vm.runInContext(js, L, { filename: 'slides_layout.html' });
  return L;
})();

const SEC = { weekly: 30, startup: 150, visitor: 20, referral: 7 };
const opt = { chapter: 'サンプル', seconds: SEC, meetingNo: '41', meetingDate: new Date(2026, 9, 6) };
const zipOf = (blob) => readZip(blob._buf);
const text = (x) => [...String(x).matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join('');
const slidesOf = (files) => {
  const prs = files['ppt/presentation.xml'].toString('utf8'), rels = files['ppt/_rels/presentation.xml.rels'].toString('utf8');
  const map = {};
  for (const m of rels.matchAll(/Id="(rId\d+)"[^>]*Target="slides\/(slide\d+\.xml)"/g)) map[m[1]] = 'ppt/slides/' + m[2];
  return [...prs.matchAll(/<p:sldId [^>]*r:id="(rId\d+)"/g)].map((m) => map[m[1]]);
};
const hides = (x) => (String(x).match(/presetClass="exit"/g) || []).length;

// ===== 1. 雛形を作る =====
const t0 = Date.now();
const zip = sb.offZipOpen_('OFFICIAL');
const made = {};
for (const kind of ['intro', 'guest', 'dairi', 'presen', 'memberPresen', 'meetingFirst', 'meetingSecond']) {
  let built = null;
  try { built = sb.offBuildTemplate_(zip, kind, opt); } catch (e) { ck(false, kind + ' を作れない: ' + e.message); continue; }
  const blob = sb.zipFromMap_(built.map, kind + '.pptx');
  made[kind] = blob;
  fs.writeFileSync(path.join(OUT, kind + '.pptx'), blob._buf);
  ck(blob._buf.length < 30 * 1024 * 1024, kind + ' が大きすぎる: ' + (blob._buf.length / 1e6).toFixed(1) + 'MB');
}
console.log('雛形を作りました（' + ((Date.now() - t0) / 1000).toFixed(1) + '秒・読んだ量 ' + (gas.log.fetchedBytes / 1e6).toFixed(0)
  + 'MB・' + gas.log.fetches + '回）: ' + Object.keys(made).map((k) => k + ' ' + (made[k]._buf.length / 1e6).toFixed(1) + 'MB').join('、'));
function integrity(files, label) {
  try {
    execFileSync('python3', [path.join(__dirname, 'pptx_integrity.py')].concat(files), { stdio: 'pipe' });
    ck(true, '');
  } catch (e) {
    ck(false, label + ' の部品のつながりが壊れている:\n' + String(e.stdout || e.message).split('\n').slice(0, 12).join('\n'));
  }
}
integrity(Object.keys(made).map((k) => path.join(OUT, k + '.pptx')), '雛形');

// 公式ファイルの音（動画）がどれも残っているか
const official = readZip(fs.readFileSync(SRC));
const videosOf = (files) => Object.keys(files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
  .reduce((a, n) => a + (files[n].toString('utf8').match(/<a:videoFile\b/g) || []).length, 0);

// ===== 2. ビジター紹介・ゲスト・代理 =====
const V = (n, cat, inv) => ({ name: n, category: cat, inviter: inv, company: '株式会社' + n.replace(/\s|　/g, '') });
const visitors = [V('見本　一郎', 'デザイン印刷', '紹介　太郎'), V('見本　花子', '税理士', '紹介　次郎'),
                  V('見本　三郎', 'カメラマン', '紹介　太郎'), V('見本　四郎', '社会保険労務士', '紹介　花子')];
for (const kind of ['intro', 'guest', 'dairi']) {
  if (!made[kind]) continue;
  const v = sb.validateTemplate_(made[kind], sb.TEMPLATE_KINDS_[kind].ids);
  ck(v.ok, kind + ' の雛形が図形の番号の確かめに通らない: ' + v.message);
  const out = sb.buildPptxFromTemplate_(made[kind], sb.makeGroups_(visitors), sb.buildGroupXml_, kind + '_out.pptx');
  const f = zipOf(out), ss = slidesOf(f), t1 = text(f[ss[0]]), t2 = text(f[ss[1]]);
  ck(ss.length === 2, kind + ': 4人で2枚にならない: ' + ss.length);
  ck(t1.includes('見本　一郎 様') && t1.includes('専門分野：デザイン印刷') && t1.includes('招待者：紹介　太郎')
     && t1.includes('見本　三郎 様') && !t1.includes('氏名'), kind + ': 1枚目の文字: ' + t1.slice(0, 160));
  ck(t2.includes('見本　四郎 様') && (t2.match(/専門分野：/g) || []).length === 1 && (t2.match(/招待者：/g) || []).length === 1
     && !t2.includes('氏名'), kind + ': 2枚目（空き枠は見出しごと消える）: ' + t2);
  const title = kind === 'intro' ? '本日のビジター' : kind === 'guest' ? '本日のゲスト' : '代理出席の方々';
  ck(t1.includes(title) && (kind !== 'dairi' || !t1.includes('歓迎')), kind + ': 題名: ' + t1.slice(0, 30));
  ck(videosOf(f) === 2, kind + ': 右上の動画（音）が残っていない: ' + videosOf(f));
}
// まとめて作る（紹介スライドの一括作成と同じ：題名だけ差し替える）
if (made.intro) {
  const items = [{ trio: [visitors[0], null, null], title: '歓迎 本日のゲスト' }];
  const out = sb.buildPptxFromTemplate_(made.intro, items, (xml, it) => sb.buildGroupXml_(sb.setTextInShape_(xml, sb.INTRO_TITLE_SHAPE_ID_, it.title), it.trio), 'all.pptx');
  const t = text(zipOf(out)[slidesOf(zipOf(out))[0]]);
  ck(t.startsWith('歓迎 本日のゲスト') && t.includes('見本　一郎 様'), 'まとめて作る（題名の差し替え）: ' + t.slice(0, 40));
}

// ===== 3. ビジタープレゼン =====
if (made.presen) {
  const v = sb.validateTemplate_(made.presen, sb.TEMPLATE_KINDS_.presen.ids);
  ck(v.ok, 'ビジタープレゼンの雛形が図形の番号の確かめに通らない: ' + v.message);
  for (const sec of [20, 30]) {
    const out = sb.buildPptxFromTemplate_(made.presen, visitors.slice(0, 2), (xml, x) => sb.buildPresenXml_(xml, x, sec), 'p.pptx');
    const f = zipOf(out), ss = slidesOf(f), x = f[ss[0]].toString('utf8');
    ck(ss.length === 2 && text(x).includes('見本　一郎 様') && text(x).includes('株式会社見本一郎') && text(x).includes('【デザイン印刷】'),
       'ビジタープレゼン ' + sec + '秒: 文字: ' + text(x).slice(0, 80));
    ck(sb.mpCountdownSeconds_(x) === sec && hides(x) === sec, 'ビジタープレゼン ' + sec + '秒: カウントダウンが ' + sb.mpCountdownSeconds_(x));
    ck(/<a:videoFile\b/.test(x) && /presetClass="mediacall"/.test(x) && /<p:video>/.test(x),
       'ビジタープレゼン ' + sec + '秒: 卵時計の動画（音）が残っていない');
    ck(/<p:cond delay="indefinite"\/>/.test(x.slice(x.indexOf('nodeType="mainSeq"'))), 'ビジタープレゼン: クリックで始まらない');
  }
}

// ===== 4. メンバープレゼン =====
const gk = (x) => String(x).replace(/[\s・･＆&と]/g, '');
const members = [
  { name: '見本　一郎', company: 'サンプル株式会社', title: '税理士', cat: '企業サポート' },
  { name: '見本　花子', company: '株式会社とても長い名前のサンプル商事ホールディングス', title: '社会保険労務士（労務の相談）', cat: '企業サポート' },
  { name: '見本　三郎', company: '三郎写真室', title: 'カメラマン', cat: '企業サポート' },
  { name: '見本　四郎', company: 'シロウ不動産', title: '売買仲介', cat: '不動産関連' },
].map((m) => Object.assign(m, { blockKey: gk(m.cat), hasPhoto: true }));
const blocks = [{ gkey: gk('企業サポート'), block: '企業サポート', count: 3 }, { gkey: gk('不動産関連'), block: '不動産関連', count: 1 }];
let items = [], boxes = null;
if (made.memberPresen) {
  const map = sb.unzipToMap_(made.memberPresen);
  boxes = sb.mpLayoutBoxes_(map);
  ck(boxes && boxes.companyPt > 40 && boxes.categoryPt > 10 && boxes.companyDefault && boxes.categoryWidth > 0,
     'メンバープレゼン：枠の寸法をテンプレートから採れない: ' + J(boxes));
  const mctx = { blocks, members, rowsPerPage: 7, seconds: { weekly: SEC.weekly, startup: SEC.startup }, layoutBoxes: boxes };
  items = JSON.parse(JSON.stringify(layout.memberPresenItems(mctx, gk('企業サポート'), '見本　花子', true)));
  ck(items.length === 6 && items.filter((x) => x.kind === 'individual').every((x) => x.companyGeom && x.companyGeom.x === boxes.companyDefault.x),
     'メンバープレゼン：組版が雛形の枠を使っていない: ' + J(items.map((x) => x.companyGeom && x.companyGeom.x)));
  mpBlob = made.memberPresen;
  const info = sb.buildMemberPresenSlides_(map, items);
  const out = sb.zipFromMap_(map, 'mp.pptx');
  fs.writeFileSync(path.join(OUT, 'memberPresen_out.pptx'), out._buf);
  integrity([path.join(OUT, 'memberPresen_out.pptx')], 'メンバープレゼンの出来上がり');
  const f = zipOf(out), ss = slidesOf(f);
  ck(ss.length === 6 && info.photos === 2, 'メンバープレゼン：ページ数・写真: ' + ss.length + ' / ' + info.photos);
  items.forEach((it, i) => {
    const x = f[ss[i]].toString('utf8'), t = text(x);
    if (it.kind === 'overview') {
      ck(t.includes(it.block) && it.rows.every((r) => t.includes(r.name)) && t.includes(it.nextName),
         'メンバープレゼン：扉 ' + it.block + ': ' + t.slice(0, 120));
      ck(/advTm="1000"/.test(x), 'メンバープレゼン：扉 ' + it.block + ' が1秒で次へ進まない');
      return;
    }
    const sec = it.name === '見本　花子' ? SEC.startup : SEC.weekly;
    const end = sb.mpCountdownEndMs_(x), adv = (x.match(/advTm="(\d+)"/) || [])[1];
    ck(t.includes(it.name) && t.includes(it.companyLines.join('')) && t.includes(it.categoryLines.join('')),
       'メンバープレゼン：' + it.name + ' の文字: ' + t.slice(0, 120));
    ck(sb.mpCountdownSeconds_(x) === sec && hides(x) === sec, 'メンバープレゼン：' + it.name + ' のカウントダウン ' + sb.mpCountdownSeconds_(x) + '（' + sec + ' のはず）');
    ck(end > sec * 1000 && String(end + 1000) === adv, 'メンバープレゼン：' + it.name + ' の自動送り ' + adv + '（卵時計のあとで終わる ' + end + ' ＋1秒のはず）');
    ck(/<a:videoFile\b/.test(x) && /<p:video>/.test(x), 'メンバープレゼン：' + it.name + ' のページに卵時計の動画（音）が無い');
    ck(it.nextName ? t.includes(it.nextName) : !t.includes('次のプレゼンター'), 'メンバープレゼン：' + it.name + ' の次の方: ' + it.nextName);
  });
  ck(/useTimings="1"/.test(f['ppt/presProps.xml'].toString('utf8')), 'メンバープレゼン：保存済みのタイミングを使う設定が無い');
}

// ===== 5. 定例会（前半）=====
if (made.meetingFirst && made.memberPresen) {
  const parts = sb.unzipToMap_(made.meetingFirst);
  const map = { '開催回': '42', '審査中カテゴリー': 'ITコンサルタント', 'メインプレゼン1氏名': '見本　一郎', 'メインプレゼン1会社名': 'サンプル株式会社',
                'メインプレゼン1カテゴリー': '【税理士】', 'メインプレゼン2氏名': '見本　花子', 'メインプレゼン2会社名': '花子商事',
                'メインプレゼン2カテゴリー': '【社会保険労務士】' };
  ['結婚相談所', '電気工事業', 'SNS運用', 'イベント企画', '工務店', '美容室経営', '相続コンサルタント'].forEach((c, i) => { map['求める専門分野' + (i + 1)] = c; });
  for (let i = 8; i <= 12; i++) map['求める専門分野' + i] = '';
  const rules = sb.meetingPatternRules_('42', new Date(2026, 9, 13));
  const weeks = [0, 1, 2, 3, 4].map((i) => ({ date: '2026/10/' + (13 + 7 * i), no: String(42 + i), md: '10/' + (13 + 7 * i), label: (i + 1) + '回目',
    source: 'rotation', people: [{ name: '見本　一郎', title: '税理士', collab: '' }, { name: '見本　花子', title: '社労士', collab: '' }] }));
  // 役職のメンバー紹介：雛形のリーダーシップチーム・サポートチームのページに、差し込み口と写真の枠の名前が入っている
  const tpl = sb.slideOrder_(parts).map((p) => sb.xmlOf_(parts, p));
  ck(tpl.some((x) => sb.slideText_(x).includes('{{プレジデント氏名}}') && x.includes('name="{{プレジデント写真}}"')
     && x.includes('name="{{書記兼会計写真}}"'))
     && tpl.some((x) => ['{{メンバーシップ委員会1氏名}}', '{{エデュケーションコーディネーター氏名}}', '{{ビジターホストコーディネーター氏名}}',
       '{{ビジターホストの一覧}}'].every((k) => sb.slideText_(x).includes(k))),
     '前半の雛形：役職のメンバー紹介の差し込み口が無い');
  // その期の役職・チーム（架空）。見本　三郎さんは写真が無い、メンターコーディネーターは空
  const roleData = sb.riDataFrom_(24, '2026年10月〜2027年3月',
    { president: '見本　一郎', vice: '見本　花子', secretary: '見本　三郎', ec: '見本　四郎', mentor: '', vhc: '見本　一郎' }, 24,
    [{ key: 'membership', name: 'メンバーシップ委員会', members: [{ name: '見本　花子' }, { name: '見本　三郎' }] },
     { key: 'role:vhc', name: 'ビジターホスト', members: [{ name: '見本　三郎' }, { name: '見本　四郎' }, { name: '見本　五郎' }] }],
    true, members.concat([{ name: '見本　五郎', company: '', title: '' }]));
  sb.getMemberMaster = () => ({ ok: true, members });
  // 新規および更新メンバー（1枚に両方の欄）・バイスプレジデントによる報告・ネットワーキングリーダー
  const memberPages = { newMembers: [{ name: '見本　一郎' }, { raw: '新井さん', category: 'エステサロン' }],
                        renewMembers: [{ name: '見本　花子', years: 2 }] };
  const vpReport = { avg: '308', month: '2026-09', count: '281', perWeek: '70', from: '2026-04', to: '2026-09', total: '1,927', thanks: '54億8,074万円' };
  const leaders = { show: true, month: '2026-09', items: [{ key: 'ext', value: '17', unit: '件', winners: [{ name: '見本　四郎' }] }] };
  const res = sb.editMeetingSlides_(parts, map, rules, { memberPresen: items, weeklyGuests: [], weeklyAuto: true,
    mainPresenters: ['見本　一郎', '見本　花子'], speakerRotation: { weeks, header: 'メインプレゼンテーション', notes: ['注意書き'] },
    roleIntro: true, roleIntroData: roleData, memberPages, vpReport, networkingLeaders: leaders, meetingDate: new Date(2026, 9, 13) });
  const out = sb.zipFromMap_(parts, 'first.pptx');
  fs.writeFileSync(path.join(OUT, 'meetingFirst_out.pptx'), out._buf);
  integrity([path.join(OUT, 'meetingFirst_out.pptx')], '前半の出来上がり');
  const f = zipOf(out), ss = slidesOf(f), all = ss.map((p) => text(f[p]));
  ck(all[0].includes('BNI サンプルチャプター') && all[0].includes('第42回') && all[0].includes('2026年10月13日'), '前半：表紙: ' + all[0]);
  const mem = all.find((t) => t.includes('メンバーシップ委員会による報告')) || '';
  ck(mem.includes('結婚相談所') && mem.includes('相続コンサルタント') && mem.includes('ITコンサルタント') && !mem.includes('{{')
     && !/First\s*&(amp;)?\s*last name/i.test(mem),
     '前半：メンバーシップ委員会の表: ' + mem.slice(0, 160));
  const mainI = all.findIndex((t) => t.startsWith('メインプレゼンテーション') && t.includes('見本　一郎'));
  ck(mainI >= 0 && all[mainI].includes('見本　花子') && all[mainI].includes('【社会保険労務士】'), '前半：メインプレゼン: ' + (all[mainI] || '').slice(0, 120));
  const mainX = mainI >= 0 ? f[ss[mainI]].toString('utf8') : '';
  ck((res.photos && /見本　一郎/.test(res.photos.message)) && (mainX.match(/mpphoto/g) || []).length === 0
     && Object.keys(f).some((n) => /ppt\/media\/mpphoto/.test(n)), '前半：メインプレゼンの写真: ' + (res.photos && res.photos.message));
  const head = all.findIndex((t) => t === 'ウィークリープレゼンテーション');
  const done = all.findIndex((t) => t.includes('終わりましたか'));
  ck(head >= 0 && done === head + 1 + items.length && all[head + 1].includes('企業サポート') && all[head + 2].includes('見本　一郎'),
     '前半：メンバーのページの差し込み位置: 見出し ' + head + '・終わりましたか ' + done + '（' + items.length + '枚）');
  const lead = ss.map((p) => f[p].toString('utf8')).find((x) => text(x).includes('リーダーシップチーム')) || '';
  ck(['見本　一郎', '見本　花子', '見本　三郎'].every((n) => text(lead).includes(n))
     && /name="\{\{プレジデント写真\}\}"/.test(lead) && /name="\{\{バイスプレジデント写真\}\}"/.test(lead)
     && !/name="\{\{書記兼会計写真\}\}"/.test(lead) && /見本　三郎/.test(res.roles.message),
     '前半：リーダーシップチーム（お名前・写真・写真の無い方の枠を外す）: ' + text(lead).slice(0, 80) + ' / ' + res.roles.message);
  const sup = text(ss.map((p) => f[p].toString('utf8')).find((x) => text(x).includes('サポートチーム')) || '');
  ck(sup.includes('見本　花子') && sup.includes('見本　四郎') && sup.includes('見本　一郎')
     && sup.includes('見本　三郎、見本　四郎、見本　五郎') && !sup.includes('氏名'),
     '前半：サポートチーム（メンバーシップ委員会・コーディネーター・ビジターホスト）: ' + sup.slice(0, 160));
  ck(/24期/.test(res.roles.message) && res.roles.pages.length === 2, '前半：役職のメンバー紹介の知らせ: ' + res.roles.message);
  const rot = ss.map((p) => f[p].toString('utf8')).find((x) => text(x).includes('スピーカーローテーション')) || '';
  ck(/第42回/.test(text(rot)) && /見本　花子/.test(text(rot)) && (rot.match(/<a:tbl>/g) || []).length === 1, '前半：スピーカーローテーションの表: ' + text(rot).slice(0, 120));
  ck(!all.some((t) => t.includes('{{')), '前半：差し込み口が残っている: ' + all.filter((t) => t.includes('{{')).map((t) => t.slice(0, 40)));
  ck(videosOf(f) >= 3 + items.filter((x) => x.kind === 'individual').length, '前半：動画（音）の数: ' + videosOf(f));
  ck(out._buf.length < 40 * 1024 * 1024, '前半の出来上がりが大きい: ' + (out._buf.length / 1e6).toFixed(1) + 'MB');
  // 新規および更新メンバー：新メンバーの欄に上から、更新メンバーの欄に年数つきで。余った「氏名」は空
  const nmX = ss.map((p) => f[p].toString('utf8')).find((x) => text(x).includes('新規および更新メンバー')) || '';
  ck(text(nmX).includes('見本　一郎') && text(nmX).includes('新井') && text(nmX).includes('見本　花子（2年）') && !text(nmX).includes('氏名')
     && !/<p:sld\b[^>]*\sshow="0"/.test(nmX), '前半：新規および更新メンバー: ' + text(nmX).slice(0, 100));
  ck(res.members && /新メンバーのページ：見本　一郎さん、新井さん/.test(res.members.message) && /名簿に無い方.*新井/.test(res.members.message),
     '前半：新メンバー・更新メンバーの知らせ: ' + (res.members && res.members.message));
  // バイスプレジデントによる報告：「月間リファーラル数の平均：」のあとに数を入れる（ほかの項目は雛形のまま）
  const vpX = ss.map((p) => f[p].toString('utf8')).find((x) => text(x).includes('バイスプレジデントによる報告')) || '';
  ck(text(vpX).includes('月間リファーラル数の平均：308件') && text(vpX).includes('月間ビジター数の平均：'),
     '前半：バイスプレジデントによる報告: ' + text(vpX).slice(0, 100));
  // ネットワーキングリーダーの紹介のページ（部門のページではない作り）は、表示にするだけ
  const nlX = ss.map((p) => f[p].toString('utf8')).find((x) => text(x).includes('ネットワーキングリーダー')) || '';
  ck(nlX && !/<p:sld\b[^>]*\sshow="0"/.test(nlX) && res.leaders, '前半：ネットワーキングリーダーのページが表示でない');
  // いない日・月の最初でない日：新規および更新メンバー・ネットワーキングリーダーのページは非表示
  const parts2 = sb.unzipToMap_(made.meetingFirst);
  sb.editMeetingSlides_(parts2, map, rules, { memberPages: { newMembers: [], renewMembers: [] }, networkingLeaders: { show: false, items: [] } });
  const hid = (t) => sb.slideOrder_(parts2).map((p) => sb.xmlOf_(parts2, p)).filter((x) => sb.slideText_(x).includes(t))
    .every((x) => /<p:sld\b[^>]*\sshow="0"/.test(x));
  ck(hid('新規および更新メンバー') && hid('ネットワーキングリーダー') && hid('倫理規定'), '前半：いない日のページ（倫理規定も）が非表示になっていない');
  // 新規および更新メンバーのいる日：そのすぐうしろの倫理規定は表示（複製はしない）
  const ordF = ss.map((p) => text(f[p])), iNm = ordF.findIndex((t) => t.includes('新規および更新メンバー'));
  ck(iNm >= 0 && ordF[iNm + 1].includes('倫理規定') && !/<p:sld\b[^>]*\sshow="0"/.test(f[ss[iNm + 1]].toString('utf8'))
     && ordF.filter((t) => t.includes('倫理規定')).length === 1, '前半：新規および更新メンバーのあとの倫理規定');
}

// ===== 6. 定例会（後半）=====
if (made.meetingSecond) {
  const parts = sb.unzipToMap_(made.meetingSecond);
  const rb = sb.rfLayoutBoxes_(parts, sb.RF_TITLE_);
  ck(rb && rb.companyPt && rb.hasNext, '後半：リファーラル発表のひな形が見つからない・寸法を採れない: ' + J(rb));
  const list = members.slice(0, 3);
  const referral = JSON.parse(JSON.stringify(layout.withLayoutBoxes(rb, () => list.map((m, i) => {
    const co = layout.layoutCompany(m.company), ca = layout.layoutCategory(m.title);
    return { name: m.name, photoName: m.name, seconds: SEC.referral, auto: false, companyLines: co.lines, companyPt: co.fontPt,
             companyGeom: co.geom, categoryTop: co.categoryTop, categoryLines: ca.lines, categoryPt: ca.fontPt,
             categoryTight: ca.tight, nextName: list[i + 1] ? list[i + 1].name : '' };
  }))));
  const who = (n, c, t) => ({ name: n, company: c, category: '【' + t + '】' });
  const pairs = [{ giver: who('見本　一郎', 'サンプル株式会社', '税理士'), receiver: who('見本　花子', '花子商事', '社労士'), after: false },
                 { giver: who('見本　三郎', '三郎写真室', 'カメラマン'), receiver: who('見本　一郎', 'サンプル株式会社', '税理士'), after: true }];
  const map = { '更新90': '見本　一郎さん', '更新60': '', '更新30': '見本　花子さん、見本　三郎さん', '更新超過': '' };
  const res = sb.editMeetingSlides_(parts, map, [], { referral, recommendPairs: pairs });
  const out = sb.zipFromMap_(parts, 'second.pptx');
  fs.writeFileSync(path.join(OUT, 'meetingSecond_out.pptx'), out._buf);
  integrity([path.join(OUT, 'meetingSecond_out.pptx')], '後半の出来上がり');
  const f = zipOf(out), ss = slidesOf(f), xs = ss.map((p) => f[p].toString('utf8')), all = xs.map(text);
  const rf = xs.map((x, i) => i).filter((i) => xs[i].includes('name="REFERRAL PRESENTATION"'));
  ck(rf.length === 3 && res.referral && /3枚/.test(res.referral.message), '後半：リファーラル発表のページ ' + rf.length + '枚: ' + (res.referral && res.referral.message));
  rf.forEach((i, k) => {
    const x = xs[i], t = all[i];
    ck(t.includes(list[k].name) && t.includes(referral[k].categoryLines.join('')), '後半：リファーラル ' + list[k].name + ' の文字: ' + t.slice(0, 100));
    ck(sb.mpCountdownSeconds_(x) === SEC.referral && hides(x) === SEC.referral, '後半：リファーラルのカウントダウン ' + sb.mpCountdownSeconds_(x));
    ck(/<p:cond delay="indefinite"\/>/.test(x.slice(x.indexOf('nodeType="mainSeq"'))) && !/advTm=/.test(x), '後半：リファーラルがクリックで始まらない・自動で進む');
    ck(k < 2 ? (t.includes('次の発表者') && t.includes(list[k + 1].name)) : !t.includes('次の発表者'), '後半：リファーラルの次の発表者 ' + k + ': ' + t.slice(-40));
    ck(/<a:videoFile\b/.test(x), '後半：リファーラルのページの右上の動画（音）が無い');
  });
  const reco = all.filter((t) => t.startsWith('推薦のことば') && t.includes('見本'));
  ck(reco.length === 2 && reco[0].includes('見本　一郎') && reco[0].includes('見本　花子') && reco[0].includes('【社労士】'),
     '後半：推薦のことば: ' + reco.map((t) => t.slice(0, 60)).join(' | '));
  const ren = all.find((t) => t.includes('更新を迎えるメンバー')) || '';
  ck(ren.includes('見本　一郎さん') && ren.includes('見本　花子さん') && (ren.match(/該当者なし/g) || []).length === 2,
     '後半：書記兼会計の更新状況: ' + ren.slice(0, 200));
  ck(!all.some((t) => t.includes('{{')), '後半：差し込み口が残っている');
}

// ===== 7. 画面から作る（01_テンプレート に保存して登録する）=====
{
  const folder = {
    files: {},
    getFilesByName(n) { const f = this.files[n]; return { hasNext: () => !!f, next: () => f }; },
    createFile(blob) {
      const id = 'MADE' + (Object.keys(this.files).length + 1), dst = path.join(OUT, id + '.pptx');
      fs.writeFileSync(dst, blob._buf);
      FILES[id] = dst;
      const f = { getId: () => id, getUrl: () => 'https://drive.google.com/file/d/' + id, setTrashed() {} };
      this.files[blob.getName()] = f;
      return f;
    },
  };
  let updated = 0;
  sb.getAssetFolder_ = (kind) => { ck(kind === 'template', '保存先が 01_テンプレート でない: ' + kind); return folder; };
  sb.Drive = { Files: { update: (m, id, blob) => { updated++; fs.writeFileSync(FILES[id], blob._buf); } } };
  sb.getMeetingCandidates = () => [{ dateValue: '2026/10/13', display: '2026/10/13(火) 第42回' }];
  sb.PropertiesService.getScriptProperties().setProperty('BNI_OFFICIAL_FILE_ID', 'OFFICIAL');
  let r = sb.buildOfficialTemplate('intro');
  const props = gas.props;
  ck(r.ok && props.BNI_TPL_INTRO_ID === 'MADE1' && JSON.parse(props.BNI_OFFICIAL_MADE).intro.id === 'MADE1' && r.sizeMB > 1,
     '画面から作る（ビジター紹介）: ' + J(r).slice(0, 200));
  r = sb.buildOfficialTemplate('meetingFirst');
  ck(r.ok && props.BNI_TPL_MEETING_FIRST_ID === 'MADE2', '画面から作る（前半）: ' + J(r).slice(0, 200));
  const cover = text(readZip(fs.readFileSync(FILES.MADE2))['ppt/slides/slide1.xml']);
  ck(cover.includes('BNI Activeチャプター') && cover.includes('第42回') && cover.includes('2026年10月13日'), '画面から作る（前半の表紙）: ' + cover);
  r = sb.buildOfficialTemplate('intro');                        // 作り直すと、同じファイルの中身を入れ替える（リンクは変わらない）
  ck(r.ok && updated === 1 && props.BNI_TPL_INTRO_ID === 'MADE1', '作り直し（同じファイルを更新）: ' + J(r).slice(0, 120));
  const st = sb.getOfficialTemplateStatus(), k = st.kinds.find((x) => x.kind === 'intro');
  ck(st.ok && k.registered && k.fromOfficial && /^\d{4}\/\d{1,2}\/\d{1,2}$/.test(k.madeAt)
     && Object.keys(folder.files).includes('BNI_テンプレート_ビジター紹介（公式から作成）.pptx'), '登録状況（公式ファイルから作成）: ' + J(k));
}

console.log(`公式ファイルから雛形: 検査 ${checks} 件（出力: ${OUT}）`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: 7つの雛形・ビジター紹介（見出しを残す）・ビジタープレゼン（卵時計・秒数）・メンバープレゼン（雛形の寸法・秒数・自動送り）'
  + '・前半（表紙・メンバーシップ・メインプレゼン・差し込み・ローテーション・新規および更新メンバー・バイス報告・ネットワーキングリーダー）'
  + '・後半（リファーラル・推薦のことば・更新状況）');

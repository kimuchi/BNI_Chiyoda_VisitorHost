// 前半スライドのメンバーのページ（ウィークリープレゼン）と、後半スライドのリファーラル発表の会社名の行を、
// 作り物の会社名で確かめる。会社名・お名前はすべて架空。
//
//   node tools/check_company_lines_client.js
//
// 確かめること
//   1) layoutCompany（slides_layout.html）は CB_ で行を分ける：法人の種類の前後・空白・語の切れ目で折り、
//      語の途中（「一般社団法人日本見／本協会」）・英単語の途中・行の頭に来てはいけない字（ー・）など）の前で折らない。
//      どの行も枠の内側に入る。名簿のセルの中の改行・㈱ は、そろえてから決める
//   2) 大きさと枠：1行は44pt・1行の枠、2行は44pt・2行の枠。2行で28ptまで小さくしても入らないときだけ3行（28pt以下）
//   3) 同じ会社は、メンバーのページとリファーラル発表（元ツールの寸法のひな形）で同じ行・同じ大きさ。幅の少し違うひな形
//      （公式ファイルから作ったもの）でも、44ptで入る会社はほとんど同じ切れ目。
//      画面の memberPresenItems（前半）と buildReferral（後半）の出力でも同じ
//   4) PCの書体によらない：canvas の測り方を変えても・canvas が無くても、会社名の行は同じ（カテゴリーは canvas で測るので変わる）
//   5) 画面（slides_meeting_first.html・slides_meeting_second.html）が、slides_layout より前に include('company_break') を読む。
//      推薦のことば（後半の personValue）が渡す会社名は CB_.normalize したもの
//   6) サーバー（member_presen_srv.js の mpIndividualSlide_・referral_srv.js の rfBuildSlide_）が、画面で決めた行を
//      1行1段落のまま書く（字・大きさ・枠も）。会社名の自動調整（雛形に残った縮小の割合）を切る
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { loadPage } = require('./lib_minidom');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const J = (x) => JSON.stringify(x);
const read = (f) => fs.readFileSync(path.join(ROOT, f), 'utf8');

// ===== 画面の組版（slides_layout.html）を、canvas の測り方を変えて読む =====
const layoutJs = [...read('slides_layout.html').matchAll(/<script>([\s\S]*?)<\/script>/g)].map((m) => m[1]).join('\n');
// measure … 1字の幅（全角1字に対する割合）。null なら canvas を作れない（作ろうとしたら止まる）
function layoutWith(measure) {
  const ctx = () => { let size = 44; return { set font(v) { const m = /(\d+(?:\.\d+)?)px/.exec(v); size = m ? +m[1] : 44; },
    measureText(t) { let w = 0; for (const ch of String(t)) w += measure(ch); return { width: w * size }; } }; };
  const doc = { createElement: () => { if (!measure) throw new Error('canvas を使った'); return { getContext: ctx }; } };
  const L = { console, document: doc };
  vm.createContext(L);
  vm.runInContext(read('company_break.js'), L, { filename: 'company_break.js' });
  vm.runInContext(layoutJs, L, { filename: 'slides_layout.html' });
  return L;
}
const A = layoutWith((ch) => (ch.charCodeAt(0) < 128 ? 0.5 : 1));                 // tools/lib_minidom.js と同じ測り方
const B = layoutWith((ch) => (ch.charCodeAt(0) < 128 ? 0.75 : /[ァ-ヶー]/.test(ch) ? 1.25 : 1.1));   // 幅の広い書体のPC
const C = layoutWith(null);                                                        // canvas を使えない
const CB = A.CB_;

// 枠（EMU）。メンバーのページは slides_layout.html の既定（元ツールの寸法）
const REF_TOOL = { companyTall: { x: 5087424, y: 2776058, cx: 6893161, cy: 1446550 }, categoryLow: 4071101, categoryWidth: 7311519 };
const REF_OFFICIAL = { companyDefault: { x: 4830858, y: 2849608, cx: 6950000, cy: 600000 }, categoryTop: 3500000, categoryWidth: 7000000 };
const co = (L, box, s) => JSON.parse(J(L.withLayoutBoxes(box, () => {
  const r = L.layoutCompany(s);
  return { lines: r.lines, pt: r.fontPt, kind: r.geom === L.COMPANY_DEFAULT ? 1 : r.geom === L.COMPANY_TALL ? 2 : 3,
           geom: r.geom, top: r.categoryTop, tall: r.geom !== L.COMPANY_DEFAULT,
           inner: Math.min(L.COMPANY_DEFAULT.cx, L.COMPANY_TALL.cx) / 12700 - 14.4, catTop: L.CATEGORY_TOP, catLow: L.CATEGORY_TOP_LOW };
})));
const show = (r) => r.lines.join('／') + ' ' + r.pt + 'pt';

// ===== 1) 自然な切れ目（メンバーのページ）=====
const WANT = [
  ['見本商事', '見本商事'],
  ['試験デザイン事務所', '試験デザイン事務所'],
  ['株式会社見本コンサルティング', '株式会社／見本コンサルティング'],
  ['見本テクノロジーズ株式会社', '見本テクノロジーズ／株式会社'],
  ['一般社団法人日本見本協会', '一般社団法人／日本見本協会'],                     // 以前は「一般社団法人日本見／本協会」
  ['税理士法人見本会計事務所', '税理士法人／見本会計事務所'],
  ['特定非営利活動法人見本子ども支援センター', '特定非営利活動法人／見本子ども支援センター'],
  ['見本建築設計事務所一級建築士事務所', '見本建築設計事務所／一級建築士事務所'],
  ['株式会社見本不動産コンサルティングオフィス', '株式会社見本不動産／コンサルティングオフィス', 40],   // 語の途中で折らず、少し小さく
  ['見本生命保険株式会社 新宿支社 第十一営業所', '見本生命保険株式会社／新宿支社 第十一営業所'],
  ['Mihon Design Studio合同会社', 'Mihon Design Studio／合同会社'],
  ['株式会社SAMPLE Creative Works', '株式会社SAMPLE／Creative Works'],
  ['㈱見本エンジニアリング', '株式会社／見本エンジニアリング'],
  ['見本生命保険㈱', '見本生命保険株式会社'],                                     // ㈱を開いて1行に入る（以前は略記だけ2行にしていた）
  ['株式会社見本\nコンサルティング', '株式会社／見本コンサルティング'],             // 名簿のセルの中の改行
];
WANT.forEach(([t, want, pt]) => {
  const r = co(A, null, t);
  ck(r.lines.join('／') === want && r.pt === (pt || 44), `1) 「${t.replace(/\n/g, '⏎')}」→「${show(r)}」（${want} ${pt || 44}pt のはず）`);
});

// 作り物の名前の組み合わせ（決まりごと）
const NO_START = /^[ー‐・･、。，．,.：:；;！？!?）)」』】〕〉》］\]々ゝゞヽヾぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶ％%]/;
const CORPS = ['株式会社', '有限会社', '合同会社', '一般社団法人', '税理士法人', '特定非営利活動法人'];
const parts = {
  pre: ['', '株式会社', '一般社団法人', '税理士法人', '特定非営利活動法人'],
  body: ['見本', '試験', 'ミホン', 'SAMPLE', 'Mihon', '日本見本まちづくり'],
  mid: ['', 'コンサルティング', 'テクノロジーズ', '不動産', 'ビューティー＆ウェルネス', ' Creative Works', '総合法律事務所',
        'インターナショナルトレーディング', '建設工業', '（ミホン）', '・アンド・パートナーズ'],
  post: ['', '株式会社', '合同会社', ' 東京支社 第十一営業所', '研究所', 'センター'],
};
const ALL = [];
for (const a of parts.pre) for (const b of parts.body) for (const c of parts.mid) for (const d of parts.post) ALL.push(a + b + c + d);
ALL.push('一般社団法人日本見本まちづくり協会連合会 東京支部 関東地区本部 事務局 見本出張所');
let n = 0, three = 0, shrunk = 0;
const same = { ref: 0, at44: 0, official: 0, officialDiff: [] };
for (const t of ALL) {
  const r = co(A, null, t);
  n++;
  const lines = r.lines, flat = CB.normalize(t).replace(/[ 　]/g, '');
  const tag = `「${t}」→「${show(r)}」`;
  if (lines.join('').replace(/[ 　]/g, '') !== flat) { ck(false, '1) 字が欠けた・増えた: ' + tag); continue; }
  if (lines.some((x, i) => i > 0 && NO_START.test(x))) ck(false, '1) 行の頭に来てはいけない字: ' + tag);
  if (lines.some((x) => !CB.fits(x, r.pt, r.inner))) ck(false, '1) 枠の内側に入らない行: ' + tag);
  for (let i = 1; i < lines.length; i++) {
    const joint = lines[i - 1] + '|' + lines[i];
    if (CORPS.some((w) => { for (let k = 1; k < w.length; k++) if (joint.includes(w.slice(0, k) + '|' + w.slice(k))) return true; return false; }))
      ck(false, '1) 法人の種類の途中で折った: ' + tag);
    if (/[A-Za-z0-9]$/.test(lines[i - 1]) && /^[A-Za-z0-9]/.test(lines[i]) && !CB.normalize(t).includes(lines[i - 1] + ' ' + lines[i].split(' ')[0]))
      ck(false, '1) 英単語の途中で折った: ' + tag);
  }
  // 2) 大きさと枠
  if (lines.length === 1) ck(r.kind === 1 && r.pt === 44 && r.top === r.catTop, '2) 1行なのに1行の枠・44pt でない: ' + tag + ' ' + J(r.geom));
  else if (lines.length === 2 && r.pt === 44) ck(r.kind === 2 && r.top === r.catLow, '2) 44ptの2行なのに2行の枠でない: ' + tag);
  else {
    shrunk++;
    ck(r.kind === 3 && r.pt < 44 && r.pt >= 20 && (lines.length === 2 ? r.pt >= 28 : r.pt <= 28) && lines.length <= 3,
       '2) 小さくするとき（2行は28ptまで・3行は28pt以下）: ' + tag);
    ck(r.geom.cy === Math.round(91440 + lines.length * r.pt * 12700 * 1.213) && r.top === r.geom.y + r.geom.cy - 151507,
       '2) 小さくしたときの枠の高さ・【カテゴリー】の位置: ' + tag + ' ' + J([r.geom, r.top]));
    if (lines.length === 3) three++;
  }
  // 3) リファーラル発表のひな形でも同じ
  const rf = co(A, REF_TOOL, t);
  if (J([rf.lines, rf.pt]) === J([r.lines, r.pt])) same.ref++;
  else ck(false, '3) リファーラル発表（元ツールの寸法）で行が違う: ' + tag + ' / ' + show(rf));
  const of = co(A, REF_OFFICIAL, t);
  if (r.pt === 44) {
    same.at44++;
    if (J(of.lines) === J(r.lines)) same.official++;
    else same.officialDiff.push(tag + ' / ' + show(of));
  }
  // 4) canvas によらない
  const rb = co(B, null, t);
  let rc;
  try { rc = co(C, null, t); } catch (e) { rc = { lines: ['（' + e.message + '）'], pt: 0 }; }
  ck(J(rb) === J(r) && J(rc) === J(r), '4) canvas の測り方で会社名の行が変わる: ' + tag + ' / ' + show(rb) + ' / ' + show(rc));
}
ck(n > 1500 && three >= 1 && shrunk >= 3, '1) 試した数・3行・小さくした数が少ない: ' + J({ n, three, shrunk }));
// 公式ファイルから作ったひな形は、枠の幅が少し違うことがある（ここでは 528pt と 533pt）。幅が近ければ、44pt で入る会社は
// ほとんど同じ切れ目（違うのは、少し広いと「少し小さくして良い切れ目」に入るときだけ。小さくする会社は大きさも変わってよい）
ck(same.official >= same.at44 * 0.98, '3) 幅の近いリファーラル発表のひな形で、切れ目の違う会社が多い: ' + same.officialDiff.length + '件 / ' + same.at44
   + ' ' + J(same.officialDiff.slice(0, 5)));
WANT.forEach(([t]) => { const a = co(A, null, t), b = co(A, REF_OFFICIAL, t); ck(J(a.lines) === J(b.lines), '3) 公式ファイルのひな形で行が違う: ' + show(a) + ' / ' + show(b)); });
// カテゴリーは、いまも canvas で測る（この検査の canvas の測り方が効いていること）
ck(J(A.layoutCategory('見本コンサルティング事業').lines) !== J(B.layoutCategory('見本コンサルティング事業').lines),
   '4) canvas の測り方を変えてもカテゴリーが変わらない（この検査の測り方が効いていない）');

// ===== 3)・5) 画面（前半・後半）から =====
{
  const pf = [];
  const server = { getSystemVersion: () => 'test' };
  const first = loadPage('slides_meeting_first.html', { server, fails: pf });
  const second = loadPage('slides_meeting_second.html', { server, fails: pf });
  ck(!pf.length, '画面を読めない: ' + J(pf));
  ['slides_meeting_first.html', 'slides_meeting_second.html'].forEach((f) => {
    const s = read(f), a = s.indexOf("include('company_break')"), b = s.indexOf("include('slides_layout')");
    ck(a >= 0 && a < b, `5) ${f} が slides_layout より前に include('company_break') を読んでいない`);
  });
  ck(typeof first.sandbox.CB_ === 'object' && typeof second.sandbox.CB_ === 'object', '5) 画面に CB_ が無い');
  const NAMES = ['見本 一郎', '試験 花子', '架空 三郎', '仮名 四郎', '見本 五郎', '試験 六子'];
  const COMPANIES = ['一般社団法人日本見本協会', '株式会社見本不動産コンサルティングオフィス', '見本商事', 'Mihon Design Studio合同会社',
                     '㈱見本エンジニアリング', '見本生命保険株式会社\n新宿支社 第十一営業所'];
  const members = NAMES.map((nm, i) => ({ no: String(i + 1), name: nm, company: COMPANIES[i], title: '見本カテゴリー', cat: '見本の区分',
                                          blockKey: 'A', hasPhoto: true }));
  const mctx = { rowsPerPage: 7, seconds: { weekly: 30, startup: 150 }, blocks: [{ gkey: 'A', block: '見本の区分', count: members.length }], members };
  const mp = first.sandbox.memberPresenItems(mctx, 'A', '', false).filter((x) => x.kind === 'individual');
  second.sandbox.ctx = { members, seconds: { referral: 7 } };
  second.sandbox.rfBoxes = REF_TOOL;
  const rf = second.sandbox.buildReferral();
  ck(mp.length === 6 && rf.length === 6, '3) 画面のページの数: ' + mp.length + ' / ' + rf.length);
  members.forEach((m, i) => {
    const want = co(A, null, m.company), a = mp[i] || {}, b = rf.filter((x) => x.name === m.name)[0] || {};
    ck(J(a.companyLines) === J(want.lines) && a.companyPt === want.pt && a.companyTall === want.tall && a.categoryTop === want.top,
       '3) 前半のメンバーのページの会社名: ' + J([a.companyLines, a.companyPt, a.companyTall]) + '（' + show(want) + ' のはず）');
    ck(J(b.companyLines) === J(want.lines) && b.companyPt === want.pt, '3) 後半のリファーラル発表の会社名: ' + J([b.companyLines, b.companyPt]) + '（' + show(want) + ' のはず）');
  });
  // 推薦のことば：名簿の会社名をそろえて渡す（セルの中の改行・㈱・空白）
  second.sandbox.ctx = { members: [{ name: '見本 一郎', company: '㈱見本\nコンサルティング', title: '税理士' },
                                   { name: '試験 花子', company: '  SAMPLE\nWORKS   合同会社 ', title: '' }] };
  const pv = [second.sandbox.personValue('見本 一郎'), second.sandbox.personValue('試験 花子'), second.sandbox.personValue('')];
  ck(J(pv.map((p) => p.company)) === J(['株式会社見本コンサルティング', 'SAMPLE WORKS 合同会社', '']) && pv[0].category === '【税理士】',
     '5) 推薦のことばの会社名をそろえていない: ' + J(pv));
}

// ===== 6) サーバーが、画面で決めた行をそのまま書く =====
{
  const sb = { console: { log() {}, warn() {}, error() {} } };
  vm.createContext(sb);
  for (const f of ['ooxml.js', 'chapter_srv.js', 'member_presen_srv.js', 'referral_srv.js', 'role_intro_srv.js'])
    vm.runInContext(read(f), sb, { filename: f });
  // カウントダウンは、この検査の外（tools/check_countdown.js）で確かめる
  sb.mpSetCountdown_ = (x) => x;
  sb.mpNoAutoAdvance_ = (x) => x;
  const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    + 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
  const xf = (x, y, cx, cy) => `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`;
  // 会社名の枠は、雛形に「はみ出すときは縮小」の割合が残っているもの
  const AUTO = '<a:bodyPr wrap="square" anchor="b"><a:normAutofit fontScale="85000" lnSpcReduction="10000"/></a:bodyPr>';
  const sp = (id, g, text, sz, body) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Text ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`
    + `<p:spPr>${xf(...g)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>`
    + `<p:txBody>${body || '<a:bodyPr wrap="square"/>'}<a:lstStyle/><a:p><a:pPr algn="l"/><a:r><a:rPr lang="ja-JP" sz="${sz}" b="1"/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
  const pic = (id, g) => `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="写真"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>`
    + `<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xf(...g)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  const slide = (body) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld><p:spTree>`
    + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' + body
    + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
  const RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>';
  // メンバーのページの個人ページ（図形IDは member_presen_srv.js の MP_INDIVIDUAL_）
  const MEMBER_PAGE = slide(pic(3, [500000, 1300000, 3800000, 4400000]) + sp(56, [4830858, 1500000, 7000000, 900000], '氏名', 5400)
    + sp(2, [4830858, 2849608, 7461517, 769441], '会社名', 4400, AUTO) + sp(12, [4830481, 3456634, 7311519, 584404], '【カテゴリー】', 3200)
    + sp(6, [4830858, 6200000, 1500000, 500000], 'NEXT➡', 2400) + sp(14, [6500000, 6200000, 3000000, 500000], '次の 方', 2400));
  // リファーラル発表のページ（図形は役割で見分ける：referral_srv.js の presenterShapes_）
  const REF_PAGE = slide(sp(20, [600000, 200000, 11000000, 700000], 'REFERRAL PRESENTATION', 4000) + pic(21, [500000, 1300000, 3800000, 4400000])
    + sp(22, [4830858, 1500000, 7000000, 900000], '見本 太郎', 5400) + sp(23, [5087424, 2776058, 6893161, 1446550], '見本株式会社', 4400, AUTO)
    + sp(24, [4830481, 4071101, 7311519, 584404], '【見本カテゴリー】', 3200)
    + sp(25, [4830858, 6200000, 1500000, 500000], 'NEXT➡', 2400) + sp(26, [6500000, 6200000, 3000000, 500000], '次の 方', 2400));
  const shapeOf = (xml, id) => {
    const r = sb.findShapeRange_(xml, id), seg = xml.substring(r.start, r.end);
    const paras = sb.findTagRanges_(seg, 'a:p').map((p) => sb.slideText_(seg.substring(p.start, p.end)));
    const sz = [...new Set([...seg.matchAll(/<a:rPr\b[^>]*\ssz="(\d+)"/g)].map((m) => +m[1] / 100))];
    const bp = (seg.match(/<a:bodyPr\b[^>]*(?:\/>|>[\s\S]*?<\/a:bodyPr>)/) || [''])[0];
    return { paras, sz, geom: sb.readShapeGeomEmu_(xml, id), bp, brs: (seg.match(/<a:br\b/g) || []).length };
  };
  const companies = ['一般社団法人日本見本協会', '見本商事', 'SAMPLE & Mihon株式会社', '株式会社見本不動産コンサルティングオフィス',
                     '一般社団法人日本見本まちづくり協会連合会 東京支部 関東地区本部 事務局 見本出張所', '見本生命保険㈱'];
  const items = A.memberPresenItems({ rowsPerPage: 7, blocks: [{ gkey: 'A', block: '見本の区分', count: companies.length }],
    members: companies.map((c, i) => ({ name: '見本 ' + (i + 1) + '郎', company: c, title: '見本カテゴリー', blockKey: 'A' })) }, 'A', '', false)
    .filter((x) => x.kind === 'individual');
  const rfItems = A.withLayoutBoxes(REF_TOOL, () => companies.map((c) => {
    const r = A.layoutCompany(c), ca = A.layoutCategory('見本カテゴリー');
    return { name: '見本 太郎', companyLines: r.lines, companyPt: r.fontPt, companyGeom: r.geom, categoryTop: r.categoryTop,
             categoryLines: ca.lines, categoryPt: ca.fontPt, categoryTight: ca.tight, nextName: '次の 方', seconds: 7, auto: false };
  }));
  const plain = (x) => JSON.parse(J(x));
  companies.forEach((c, i) => {
    const it = plain(items[i]);
    const x = sb.mpIndividualSlide_(MEMBER_PAGE, RELS, it, null, { W: 12192000, H: 6858000 }).xml, s = shapeOf(x, 2);
    ck(J(s.paras) === J(it.companyLines) && s.brs === 0, '6) メンバーのページの会社名がそのまま入らない: ' + J(s.paras) + '（' + J(it.companyLines) + ' のはず）');
    ck(J(s.sz) === J([it.companyPt]) && J(s.geom) === J(it.companyGeom), '6) メンバーのページの会社名の大きさ・枠: ' + J([s.sz, s.geom, it.companyPt, it.companyGeom]));
    ck(/<a:noAutofit\/>/.test(s.bp) && !/normAutofit|spAutoFit/.test(s.bp) && /wrap="square"/.test(s.bp),
       '6) メンバーのページの会社名の自動調整を切っていない（折り返しは雛形のまま）: ' + s.bp);
    const rt = plain(rfItems[i]);
    const SH = sb.presenterShapes_(REF_PAGE, 12192000);
    SH.W = 12192000; SH.H = 6858000;
    const y = sb.rfBuildSlide_(REF_PAGE, RELS, rt, null, SH).xml, t = shapeOf(y, SH.company);
    ck(SH.company === '23', '6) リファーラル発表のひな形の会社名の枠を見分けられない: ' + J(SH));
    ck(J(t.paras) === J(rt.companyLines) && J(t.paras) === J(it.companyLines), '6) リファーラル発表の会社名がそのまま入らない・メンバーのページと違う: '
       + J(t.paras) + ' / ' + J(it.companyLines));
    ck(J(t.sz) === J([rt.companyPt]) && J(t.geom) === J(rt.companyGeom) && /<a:noAutofit\/>/.test(t.bp) && !/normAutofit/.test(t.bp),
       '6) リファーラル発表の会社名の大きさ・枠・自動調整: ' + J([t.sz, t.geom, t.bp]));
  });
  // 枠を渡さない古い画面のときは、リファーラル発表の雛形の自動調整に任せる（枠の高さが雛形のままのため）
  const SH = sb.presenterShapes_(REF_PAGE, 12192000);
  const old = shapeOf(sb.rfBuildSlide_(REF_PAGE, RELS, { name: '見本 太郎', companyLines: ['見本商事'], categoryLines: ['【見本】'], seconds: 7 }, null, SH).xml, SH.company);
  ck(J(old.paras) === J(['見本商事']) && /normAutofit/.test(old.bp), '6) 枠を渡さない画面のとき、雛形の自動調整を変えた: ' + old.bp);
}

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.slice(0, 40).forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('メンバーのページ・リファーラル発表の会社名: 検査 ' + checks + ' 件 OK（会社名 ' + ALL.length + ' 通り）: 語の切れ目で折る（CB_）・行頭禁則・'
  + '英単語と法人の種類の途中で折らない・枠に入る・1行/2行は44pt・3行は28pt以下・メンバーのページとリファーラル発表で同じ行・'
  + 'canvas によらない・推薦のことばの会社名をそろえる・サーバーがそのまま書く（自動調整を切る）');

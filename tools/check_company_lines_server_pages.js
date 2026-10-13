// 前半スライドの1人ぶんのページと、事前MTGの熱烈歓迎のページの会社名の行を、作り物のページで確かめる。
// 会社名・お名前はすべて架空。雛形の実物は使わない。
//
//   node tools/check_company_lines_server_pages.js
//
// どのページも、会社名の行は会社名の改行の決め方（company_break.js の CB_）で決めて書く（PowerPoint に折らせない）。
// 確かめること
//   ・新メンバー・更新メンバー（1人1枚）／ネットワーキングリーダーの部門のページ／熱烈歓迎 … 1行の枠：
//     入ればそのまま・入らなければ1行のまま小さく（見積もりは CB_.lineWidthPt。枠の左右の余白 lIns・rIns も引く）・
//     下限（元の半分・熱烈歓迎は12pt）でも入らなければ、自然な切れ目の2行（同じ段落の中の <a:br>。下の【カテゴリー】の段落はそのまま）
//   ・ネットワーキングリーダーのまとめのページ … 狭い列：元の大きさで自然な切れ目の2行まで（並べ直して狭くなった列も）
//   ・役職のメンバー紹介 … 1人のページ（枠を見分けて差し込み口を書くページ）・差し込み口の雛形（グループの中の枠も）：
//     1行に入れば1行、入らなければ自然な切れ目の2行（半分まで小さく）。ほかの字と同じ段落・大きさのわからない差し込み口の枠は1行のまま
//   ・名簿の文字をそろえる（セルの中の改行・㈱）。<a:t> の中に改行が残らない
//   ・どの行も、見積もりの幅が枠の内側に入る。切れ目は、その幅での CB_.layout と同じ。幅が近ければ、ページが違っても同じ切れ目
//   ・XMLとして壊れていない（熱烈歓迎はできたpptxを tools/pptx_integrity.py で）
process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const os = require('os');
const path = require('path');
const vm = require('vm');
const { spawnSync } = require('child_process');
const { makeEnv } = require('./lib_sheet_fake');
const { readZip, writeZip } = require('./lib_zip');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const J = (x) => JSON.stringify(x);

function blob(content, type, name) {
  let nm = name || '';
  const buf = Buffer.isBuffer(content) ? content : Buffer.from(String(content), 'utf8');
  return { _buf: buf, getDataAsString: () => buf.toString('utf8'), getBytes: () => Array.from(buf).map((b) => (b > 127 ? b - 256 : b)),
           getContentType: () => type, setName(n) { nm = n; return this; }, getName: () => nm };
}

// ===== 作り物の会社名 =====
const CO = {
  short: '見本商事',
  newline: '㈱見本\nエンジニアリング',                                                   // 名簿のセルの中の改行・㈱
  shrink: '株式会社見本マーケティングソリューションズ東京支社',                            // 1行のまま小さくする
  insets: '社会保険労務士法人見本人事サポートセンター',                                      // 余白が広い枠でだけ小さくする
  long: '一般社団法人日本見本インターナショナルコンサルティングマーケティングソリューションズ協会東京支部事務局',
  nlMid: '株式会社見本マーケティングソリューションズ',
  nlLong: '一般社団法人日本見本インターナショナルコンサルティング協会',
  col1: '株式会社見本コンサルティング',
  col2: '見本テクノロジーズ株式会社',
  col3: '一般社団法人日本見本協会',
  col4: '税理士法人見本会計事務所',
  single: '株式会社見本不動産コンサルティングオフィス',
};
const ROSTER = [
  { name: '見本 新一', company: CO.short, title: '税理士' },
  { name: '見本 新二', company: CO.newline, title: '司法書士' },
  { name: '見本 新三', company: CO.shrink, title: 'Web制作' },
  { name: '見本 新四', company: CO.long, title: '行政書士' },
  { name: '見本 新五', company: CO.insets, title: '社会保険労務士' },
  { name: '試験 一郎', company: CO.short, title: '工務店' },
  { name: '試験 二郎', company: CO.nlMid, title: '生命保険' },
  { name: '試験 三郎', company: CO.nlLong, title: '贈答用生鮮食品' },
  { name: '試験 四郎', company: CO.newline, title: '保険代理店' },
  { name: '架空 一子', company: CO.col1, title: 'コンサルタント' },
  { name: '架空 二子', company: CO.col2, title: 'システム開発' },
  { name: '架空 三子', company: CO.col3, title: '協会運営' },
  { name: '架空 四子', company: CO.col4, title: '税理士' },
  { name: '架空 五子', company: CO.short, title: '印刷' },
  { name: '架空 六子', company: CO.single, title: '不動産' },
];

// ===== 本番の *.js を読む（Drive・名簿のところだけ差し替え）=====
const sb = {
  console: { log() {}, warn() {}, error: console.error },
  Utilities: { newBlob: (c, t, n) => blob(c, t, n) },
  DriveApp: { getFileById: () => { throw new Error('写真は無い'); } },
  normName_: (s) => String(s == null ? '' : s).normalize('NFKC').replace(/[\s　]/g, ''),
  findPhotoIdForName_: () => '',
};
vm.createContext(sb);
for (const f of ['ooxml.js', 'company_break.js', 'chapter_srv.js', 'member_presen_srv.js', 'referral_srv.js', 'splice_srv.js',
                 'meeting_slides_srv.js', 'role_input_srv.js', 'role_intro_srv.js', 'routine_srv.js', 'meeting_pages_srv.js']) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), sb, { filename: f });
}
sb.getMemberMaster = () => ({ ok: true, members: ROSTER.map((m) => Object.assign({}, m)) });
const F = sb, CB = sb.CB_;

// ===== 作り物のページ =====
const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
  + 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const E = 12700;
const xf = (x, y, w, h) => `<a:xfrm><a:off x="${Math.round(x * E)}" y="${Math.round(y * E)}"/><a:ext cx="${Math.round(w * E)}" cy="${Math.round(h * E)}"/></a:xfrm>`;
const para = (t, sz) => `<a:p><a:r><a:rPr lang="ja-JP" sz="${sz}"/><a:t>${t}</a:t></a:r><a:endParaRPr lang="ja-JP" sz="${sz}"/></a:p>`;
// paras … [[文字, 大きさ], …]（大きさが null なら sz を書かない）。o.ins … 左右の余白（EMU）、o.ph … 差し込み口の枠、o.title
const box = (id, x, y, w, h, paras, o) => {
  o = o || {};
  const bp = o.ins ? `<a:bodyPr wrap="square" lIns="${o.ins}" rIns="${o.ins}"/>` : '<a:bodyPr wrap="square"/>';
  const nv = o.title ? '<p:nvPr><p:ph type="title"/></p:nvPr>' : o.ph ? '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr>' : '<p:nvPr/>';
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Text ${id}"/><p:cNvSpPr txBox="1"/>${nv}</p:nvSpPr>`
    + `<p:spPr>${xf(x, y, w, h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody>${bp}<a:lstStyle/>`
    + paras.map(([t, sz]) => (sz ? para(t, sz) : `<a:p><a:r><a:rPr lang="ja-JP"/><a:t>${t}</a:t></a:r></a:p>`)).join('') + '</p:txBody></p:sp>';
};
const sp = (id, x, y, w, h, lines, sz, o) => box(id, x, y, w, h, lines.map((t) => [t, sz]), o);
const pic = (id, x, y, w, h, rid, name) => `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${name || '図 ' + id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>`
  + `<p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>`
  + `<p:spPr>${xf(x, y, w, h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
// グループ（子の座標は 0.8 倍に縮めて置く。グループの中の枠の幅を、スライド上の幅に直せているかを確かめるため）
const grp = (id, x, y, w, h, kids) => `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="グループ ${id}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`
  + `<p:grpSpPr><a:xfrm><a:off x="${x * E}" y="${y * E}"/><a:ext cx="${w * E}" cy="${h * E}"/>`
  + `<a:chOff x="0" y="0"/><a:chExt cx="${Math.round(w / 0.8 * E)}" cy="${Math.round(h / 0.8 * E)}"/></a:xfrm></p:grpSpPr>${kids}</p:grpSp>`;
const slideXml = (inner) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld><p:spTree>`
  + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'
  + inner + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
function makeParts(slides) {
  const parts = {};
  const put = (p, s) => { parts[p] = blob(s, 'application/xml', p); };
  put('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + slides.map((k, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('') + '</Types>');
  put('ppt/presentation.xml', `<?xml version="1.0"?><p:presentation ${NS}><p:sldIdLst>`
    + slides.map((k, i) => `<p:sldId id="${256 + i}" r:id="rId${100 + i}"/>`).join('') + '</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>');
  put('ppt/_rels/presentation.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + slides.map((k, i) => `<Relationship Id="rId${100 + i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`).join('')
    + '</Relationships>');
  slides.forEach((inner, i) => {
    put(`ppt/slides/slide${i + 1}.xml`, slideXml(inner));
    put(`ppt/slides/_rels/slide${i + 1}.xml.rels`, '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + [2, 3, 4, 5].map((n) => `<Relationship Id="rId${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${i + 1}_${n}.png"/>`).join('')
      + '</Relationships>');
  });
  return parts;
}
const shown = (parts) => F.slideOrder_(parts).map((p) => F.xmlOf_(parts, p)).filter((x) => !/<p:sld\b[^>]*\sshow="0"/.test(x));

// ===== 見る道具 =====
const shapeSeg = (xml, id) => { const r = F.findShapeRange_(xml, id); return r ? xml.substring(r.start, r.end) : ''; };
const parasOf = (seg) => F.findTagRanges_(seg, 'a:p').map((r) => seg.substring(r.start, r.end));
const linesOf = (p) => p.split(/<a:br\b[^>]*\/>|<a:br\b[^>]*>[\s\S]*?<\/a:br>/).map((x) => F.slideText_(x));
const sizesOf = (p) => [...p.matchAll(/<a:(?:rPr|endParaRPr)\b[^>]*\ssz="(\d+)"/g)].map((m) => +m[1]);
const innerOf = (seg) => {
  const cx = +seg.match(/<a:ext cx="(\d+)"/)[1], bp = (seg.match(/<a:bodyPr\b[^>]*>/) || [''])[0];
  const ins = (k) => +((bp.match(new RegExp('\\s' + k + '="(\\d+)"')) || [])[1] || 91440);
  return (cx - ins('lIns') - ins('rIns')) / E;
};
// XMLの形（開く・閉じるのつり合い・名前の付いた文字の書き方・<a:t> の中の改行）
function wellFormed(xml) {
  const s = xml.replace(/^<\?xml[^>]*\?>/, ''), stack = [];
  const re = /<(\/?)([A-Za-z_][\w:.-]*)((?:\s+[\w:.-]+="[^"<]*")*)\s*(\/?)>/y;
  for (let i = 0; i < s.length;) {
    const lt = s.indexOf('<', i);
    const text = s.slice(i, lt < 0 ? s.length : lt);
    if (/&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9A-Fa-f]+);)/.test(text)) return '文字の中の & : ' + text.slice(0, 40);
    if (lt < 0) break;
    re.lastIndex = lt;
    const m = re.exec(s);
    if (!m) return '読めないタグ: ' + s.slice(lt, lt + 60);
    if (m[1]) { if (stack.pop() !== m[2]) return '閉じるタグが合わない: ' + m[2]; } else if (!m[4]) stack.push(m[2]);
    i = re.lastIndex;
  }
  if (stack.length) return '閉じていないタグ: ' + stack.join(',');
  if (/<a:t(?:\s[^>]*)?>[^<]*[\n\r][^<]*<\/a:t>/.test(xml)) return '<a:t> の中に改行が残っている';
  return '';
}
// 1行の枠の決まり（入る→そのまま／1行のまま小さく（0.5pt刻み・下限まで）／下限でも入らない→ twoPt〜下限の自然な2行）
function expectOne(text, inner, pt, minPt, twoPt) {
  if (CB.fits(text, pt, inner)) return { kind: 'fit', lines: [text], pt };
  const fit = Math.floor(inner / CB.lineWidthPt(text, 1) * 2) / 2;
  if (fit >= minPt) return { kind: 'shrink', lines: [text], pt: fit };
  const r = CB.layout(text, { width: inner, pt: Math.max(minPt, twoPt), minPt, maxLines: 2 });
  return { kind: 'two', lines: r.lines, pt: r.pt, natural: r.natural };
}
// 2行まで書ける枠の決まり（元の大きさで1行か自然な2行・半分まで小さく。2行では語の途中でしか折れなければ、2行ぶんの高さの自然な3行）
function expectTwo(text, inner, pt, minPt) {
  const r = CB.layout(text, { width: inner, pt, minPt, maxLines: 2 });
  if (r.natural) return r;
  const t = CB.layout(text, { width: inner, pt, minPt, maxLines: 3, accept: (ls, q) => ls.length * q <= 2 * pt + 1e-6 });
  return t.natural ? t : r;
}
// 会社名の段落を確かめる。want … { lines, pt }。base … 元の大きさ（pt）。budget … 元の大きさで何行ぶんの高さまでよいか
function checkCompany(label, seg, idx, want, inner, base, budget) {
  const p = parasOf(seg)[idx] || '', got = linesOf(p), szs = sizesOf(p);
  ck(J(got) === J(want.lines), `${label}: 行 ${J(got)}（${J(want.lines)} のはず）`);
  ck(szs.length && szs.every((z) => z === Math.round(want.pt * 100)), `${label}: 大きさ ${J(szs)}（${want.pt}pt のはず）`);
  ck(got.every((l) => CB.lineWidthPt(l, want.pt) <= inner + 1e-6), `${label}: 枠の内側（${inner}pt）に入らない行 ${J(got.map((l) => [l, CB.lineWidthPt(l, want.pt)]))}`);
  if (got.length > 1) {
    ck(want.natural !== false, `${label}: 語の途中で折った ${J(got)}`);
    ck(got.length * want.pt <= (budget || 1) * base + 1e-6, `${label}: 元の${budget || 1}行ぶんより高い ${J([got.length, want.pt, base])}`);
  }
  return got;
}

// ===== 1) 新メンバー・更新メンバー（1人1枚）=====
// お名前（44pt）の下に、会社名と【カテゴリー】の枠（幅480pt・20pt）。余白の広い枠（左右 0.4インチ）のページも
{
  const person = (head, ins) => sp(2, 0, 20, 960, 70, [head], 4000, { title: true })
    + pic(5, 330, 95, 300, 200, 'rId3') + sp(3, 240, 310, 480, 70, ['見本 太郎'], 4400)
    + sp(4, 240, 390, 480, 70, ['見本株式会社', '【見本カテゴリー】'], 2000, { ins });
  const parts = makeParts([sp(2, 100, 200, 700, 80, ['ようこそ　BNI 見本チャプター'], 4000), person('新メンバー'), person('更新メンバー', 365760),
                           sp(2, 0, 200, 960, 80, ['ウィークリープレゼンテーション'], 5400)]);
  const people = (names) => names.map((n) => ({ name: n, raw: n, years: 1 }));
  const res = F.applyMemberPages_(parts, { newMembers: people(['見本 新一', '見本 新二', '見本 新三', '見本 新四']),
                                           renewMembers: people(['見本 新五', '見本 新三']) }, { by: {}, seq: 0 });
  const pages = shown(parts).filter((x) => F.fpMemberKind_(x));
  const want = [['見本 新一', CO.short, 'fit', 0], ['見本 新二', '株式会社見本エンジニアリング', 'fit', 0], ['見本 新三', CO.shrink, 'shrink', 0],
                ['見本 新四', CO.long, 'two', 0], ['見本 新五', CO.insets, 'shrink', 365760], ['見本 新三', CO.shrink, 'shrink', 365760]];
  ck(pages.length === 6 && res.new.length === 4 && res.renew.length === 2, '1) ページの数: ' + pages.length);
  pages.forEach((x, i) => {
    const [nm, co, kind] = want[i], seg = shapeSeg(x, 4), inner = innerOf(seg), e = expectOne(CB.normalize(co), inner, 20, 10, 10);
    const label = `1) ${nm}（${co.replace(/\n/g, '\\n')}・内側 ${inner}pt）`;
    ck(F.slideText_(shapeSeg(x, 3)) === nm, label + ': お名前');
    ck(e.kind === kind, `${label}: 作り物の会社名の長さが、確かめたい場合（${kind}）になっていない: ${e.kind}`);
    checkCompany(label, seg, 0, e, inner, 20);
    const cat = parasOf(seg)[1] || '';
    ck(parasOf(seg).length === 2 && /^【.+】$/.test(F.slideText_(cat)) && J(sizesOf(cat)) === J([2000, 2000]), label + ': 【カテゴリー】の段落が変わった: ' + F.slideText_(cat));
    const bad = wellFormed(x);
    ck(!bad, label + ': XML ' + bad);
  });
  // 余白の広い枠：同じ会社名でも、内側が狭いぶん小さくなる（既定の余白の枠なら20ptのまま入る）
  ck(CB.fits(CO.insets, 20, 480 - 14.4) && !CB.fits(CO.insets, 20, 480 - 57.6), '1) 余白を確かめる会社名の長さ');
  // 2行のとき：半分の大きさ（10pt）の自然な2行。法人の種類・カタカナ語の途中で折らない
  const two = linesOf(parasOf(shapeSeg(pages[3], 4))[0]);
  ck(J(two) === J(['一般社団法人日本見本インターナショナルコンサルティング', 'マーケティングソリューションズ協会東京支部事務局']), '1) 2行の切れ目: ' + J(two));
  ck(/<a:br><a:rPr lang="ja-JP" sz="1000"\/><\/a:br>/.test(shapeSeg(pages[3], 4)), '1) 2行目を同じ段落の中の改行（同じ書式）にしていない');
}

// ===== 2) ネットワーキングリーダーの部門のページ（1行の枠）とまとめのページ（2行まで）=====
const TITLE = sp(2, 180, 5, 760, 60, ['ネットワーキングリーダーの発表'], 4000, { title: true });
const kindPage = (label, value) => TITLE
  + pic(3, 90, 110, 300, 300, 'rId2', '月桂樹') + sp(4, 140, 220, 200, 70, label, 3600) + sp(5, 80, 420, 320, 60, [value], 4000)
  + pic(6, 560, 110, 250, 300, 'rId3') + sp(7, 500, 420, 370, 50, ['見本 太郎'], 3600) + sp(8, 500, 470, 370, 70, ['見本株式会社', '【見本カテゴリー】'], 2400);
const LABELS = [['CEU'], ['サンキュー'], ['外部', 'リファーラル'], ['1to1'], ['ビジター', '招待数']];
const KEYS = ['ceu', 'thanks', 'ext', 'oto', 'visitor'];
const summary = sp(2, 100, 10, 760, 60, ['６月のネットワーキングリーダーの皆さま'], 3600, { title: true })
  + LABELS.map((label, i) => {
    const x = 30 + i * 180, id = 10 + i * 10;
    return sp(id, x + 10, 90, 160, 50, label, 2000) + pic(id + 1, x + 30, 150, 120, 150, 'rId' + (2 + (i % 3)))
      + sp(id + 2, x, 310, 180, 40, ['見本 太郎'], 2000) + sp(id + 3, x, 350, 180, 60, ['見本株式会社', '【見本カテゴリー】'], 1400);
  }).join('');
const nlItems = (wins) => KEYS.map((k) => ({ key: k, value: '1', unit: '', winners: (wins[k] || []).map((n) => ({ name: n, raw: n, category: '' })) }));
{
  const parts = makeParts([TITLE].concat(['25PT', '4,000万円', '17件', '17件', '8名'].map((v, i) => kindPage(LABELS[i], v))));
  F.applyNetworkingLeaders_(parts, { show: true, month: '2026-06', items: nlItems({ ceu: ['試験 一郎'], thanks: ['試験 二郎'], ext: ['試験 三郎'], visitor: ['試験 四郎'] }) },
                            { by: {}, seq: 0 });
  const want = { '試験 一郎': [CO.short, 'fit'], '試験 二郎': [CO.nlMid, 'shrink'], '試験 三郎': [CO.nlLong, 'two'], '試験 四郎': ['株式会社見本エンジニアリング', 'fit'] };
  const pages = shown(parts).filter((x) => F.fpNlPage_(x, 12192000, 6858000).type === 'kind');
  ck(pages.length === 4, '2) 部門のページの数: ' + pages.length);
  pages.forEach((x) => {
    const nm = F.slideText_(shapeSeg(x, 7)), [co, kind] = want[nm] || ['', ''], seg = shapeSeg(x, 8), inner = innerOf(seg);
    const e = expectOne(co, inner, 24, 12, 12), label = `2) 部門のページ ${nm}（${co}・内側 ${inner}pt）`;
    ck(e.kind === kind, `${label}: 作り物の会社名の長さが、確かめたい場合（${kind}）になっていない: ${e.kind}`);
    checkCompany(label, seg, 0, e, inner, 24);
    ck(/^【.+】$/.test(F.slideText_(parasOf(seg)[1] || '')), label + ': 【カテゴリー】の段落');
    const bad = wellFormed(x);
    ck(!bad, label + ': XML ' + bad);
  });
}
// まとめのページ：列の幅180pt・14pt。全部門に1人ずつ（並べ直さない）と、お2人の部門がある月（列が狭くなる）
const summaryLines = {};
[[{ ceu: ['架空 一子'], thanks: ['架空 二子'], ext: ['架空 三子'], oto: ['架空 四子'], visitor: ['架空 五子'] }, false],
 [{ ceu: ['架空 一子'], thanks: ['架空 二子', '架空 三子'], ext: ['架空 四子'], oto: ['試験 三郎'], visitor: ['架空 五子'] }, true]].forEach(([wins, relaid], t) => {
  const parts = makeParts([TITLE, summary]);
  F.applyNetworkingLeaders_(parts, { show: true, month: '2026-06', items: nlItems(wins) }, { by: {}, seq: 0 });
  const x = F.xmlOf_(parts, 'ppt/slides/slide2.xml'), byName = {};
  ROSTER.forEach((m) => { byName[m.name] = CB.normalize(m.company); });
  // 会社名の枠：お名前の枠のすぐ下の、会社名と【カテゴリー】の2段落の枠
  const units = F.riShapes_(x).filter((s) => s.tag === 'p:sp' && s.paras.length === 2 && /^【/.test(s.paras[1]));
  const names = F.riShapes_(x).filter((s) => s.tag === 'p:sp' && s.paras.length === 1 && byName[s.text]);
  ck(units.length === (relaid ? 6 : 5), `2) まとめ${t}: 会社名の枠の数 ${units.length}`);
  units.forEach((u) => {
    const nm = names.filter((n) => Math.abs(n.ax - u.ax) < E && n.ay < u.ay).sort((a, b) => b.ay - a.ay)[0];
    const co = nm ? byName[nm.text] : '', seg = shapeSeg(x, u.id), inner = innerOf(seg);
    const r = expectTwo(co, inner, 14, 7), label = `2) まとめ${t} ${nm && nm.text}（${co}・内側 ${Math.round(inner * 10) / 10}pt）`;
    const got = checkCompany(label, seg, 0, { lines: r.lines, pt: r.pt, natural: r.natural }, inner, 14, 2);
    if (!relaid) summaryLines[co] = r.lines;
    ck(parasOf(seg).length === 2 && /^【.+】$/.test(F.slideText_(parasOf(seg)[1])), label + ': 【カテゴリー】の段落が変わった');
    // 狭くなった列のとても長い会社名：2行では自然な切れ目で入らないので、2行ぶんの高さの3行（語の途中で詰めて折らない）
    if (co === CO.nlLong) {
      ck(J(got) === J(['一般社団法人', '日本見本インターナショナル', 'コンサルティング協会']) && r.pt * 3 <= 28,
         label + ': とても長い会社名の3行: ' + J([got, r.pt]));
    }
    ck(Math.abs(inner - (relaid ? innerOf(seg) : 165.6)) < 0.01, label + ': 列の幅');
  });
  ck(J(summaryLines[CO.col1]) === J(['株式会社', '見本コンサルティング']) && J(summaryLines[CO.col2]) === J(['見本テクノロジーズ', '株式会社'])
     && J(summaryLines[CO.col3]) === J(['一般社団法人', '日本見本協会']) && J(summaryLines[CO.col4]) === J(['税理士法人', '見本会計事務所'])
     && J(summaryLines[CO.short]) === J([CO.short]), '2) まとめの自然な切れ目: ' + J(summaryLines));
  const bad = wellFormed(x);
  ck(!bad, `2) まとめ${t}: XML ` + bad);
});

// ===== 3) 役職のメンバー紹介 =====
{
  const MEMBERS = [
    { name: '架空 六子', kana: 'かくう むつこ', company: CO.single, title: '不動産' },
    { name: '架空 一子', kana: '', company: CO.col1, title: 'コンサルタント' },
    { name: '見本 新二', kana: '', company: CO.newline, title: '司法書士' },
    { name: '見本 新一', kana: '', company: CO.short, title: '税理士' },
    { name: '見本 新四', kana: '', company: CO.long, title: '行政書士' },
  ];
  // 1人のページ（枠を見分けて差し込み口を書く）：お名前66pt・ローマ字24pt・会社名32pt・（カテゴリー）28pt。枠の幅475pt
  const single = (title, nm) => sp(3, 20, 16, 459, 93, [title], 4000, { title: true })
    + box(6, 0, 191, 475, 271, [[nm, 6600], ['ZENNIN ICHIRO', 2400], ['前任商事', 3200], ['（旧カテゴリー）', 2800]])
    + pic(7, 506, 88, 353, 433, 'rId3');
  // 差し込み口の雛形：ふつうの枠（幅300pt・24pt）・グループの中の枠（幅180pt・14pt）・ほかの字と同じ段落・
  // 大きさの無い差し込み口の枠（p:ph）・短い会社名
  const tokens = sp(2, 0, 10, 960, 60, ['役職とチームの紹介'], 4000, { title: true })
    + sp(20, 40, 100, 300, 60, ['{{バイスプレジデント会社名}}'], 2400)
    + grp(30, 400, 100, 180, 60, sp(31, 0, 0, 180 / 0.8, 60 / 0.8, ['{{メンバーシップ委員会1会社名}}'], 1400))
    + sp(40, 40, 200, 300, 60, ['会社名：{{書記兼会計会社名}}'], 2400)
    + box(50, 40, 300, 300, 60, [['{{エデュケーションコーディネーター会社名}}', null]], { ph: true })
    + sp(60, 400, 300, 300, 60, ['{{メンターコーディネーター会社名}}'], 2400)
    + sp(70, 40, 400, 300, 60, ['{{ビジターホストコーディネーター会社名}}'], 2400);
  const parts = makeParts([single('プレジデント', '前任　一郎'), tokens, sp(2, 0, 200, 960, 80, ['ウィークリープレゼンテーション'], 5400)]);
  const holders = { president: '架空 六子', vice: '架空 六子', secretary: '見本 新四', ec: '見本 新四', mentor: '見本 新一', vhc: '見本 新二' };
  const data = F.riDataFrom_(24, '', holders, 24, [{ key: 'membership', name: 'membership', members: [{ name: '架空 一子' }] }], true, MEMBERS);
  F.applyRoleIntro_(parts, data, { by: {}, seq: 0 });
  const s1 = F.xmlOf_(parts, 'ppt/slides/slide1.xml'), s2 = F.xmlOf_(parts, 'ppt/slides/slide2.xml');
  // 1人のページ：会社名の段落（3つ目）だけ自然な2行。お名前・ローマ字・カテゴリーの段落はそのまま
  {
    const seg = shapeSeg(s1, 6), inner = innerOf(seg), r = expectTwo(CO.single, inner, 32, 16);
    const got = checkCompany('3) 1人のページ', seg, 2, { lines: r.lines, pt: r.pt, natural: r.natural }, inner, 32, 2);
    ck(J(got) === J(['株式会社見本不動産', 'コンサルティングオフィス']), '3) 1人のページの切れ目: ' + J(got));
    const ps = parasOf(seg);
    ck(ps.length === 4 && F.slideText_(ps[0]) === '架空 六子' && J(sizesOf(ps[0])) === J([6600, 6600]) && J(sizesOf(ps[3])) === J([2800, 2800])
       && F.slideText_(ps[3]) === '（不動産）', '3) 1人のページのほかの段落: ' + J(ps.map((p) => [F.slideText_(p), sizesOf(p)])));
  }
  // 差し込み口の雛形
  const tok = (id, base, co, label) => {
    const seg = shapeSeg(s2, id), inner = id === 31 ? 180 - 14.4 : innerOf(seg), r = expectTwo(co, inner, base, base / 2);
    return checkCompany('3) ' + label, seg, 0, { lines: r.lines, pt: r.pt, natural: r.natural }, inner, base, 2);
  };
  ck(J(tok(20, 24, CO.single, '差し込み口（幅300pt・24pt）')) === J(['株式会社見本不動産', 'コンサルティングオフィス']), '3) 差し込み口の切れ目');
  const inGroup = tok(31, 14, CO.col1, 'グループの中の差し込み口（幅180pt・14pt）');
  ck(J(inGroup) === J(summaryLines[CO.col1]), '3) 幅の近い枠（まとめのページの列）と切れ目が違う: ' + J([inGroup, summaryLines[CO.col1]]));
  ck(J(tok(60, 24, CO.short, '短い会社名')) === J([CO.short]) && !/<a:br\b/.test(shapeSeg(s2, 60)), '3) 短い会社名を2行にした');
  ck(tok(70, 24, '株式会社見本エンジニアリング', '名簿の文字をそろえた会社名').join('') === '株式会社見本エンジニアリング', '3) 名簿の文字をそろえない');
  ck(F.slideText_(shapeSeg(s2, 40)) === '会社名：' + CB.normalize(CO.long) && !/<a:br\b/.test(shapeSeg(s2, 40)),
     '3) ほかの字と同じ段落の差し込み口は1行のまま: ' + F.slideText_(shapeSeg(s2, 40)));
  ck(F.slideText_(shapeSeg(s2, 50)) === CB.normalize(CO.long) && !/<a:br\b/.test(shapeSeg(s2, 50)) && !/\ssz="/.test(shapeSeg(s2, 50)),
     '3) 大きさのわからない差し込み口の枠（p:ph）は1行のまま・大きさを書かない');
  ck(!/\{\{/.test(F.slideText_(s1) + F.slideText_(s2)), '3) 差し込み口が残った');
  [s1, s2].forEach((x, i) => { const bad = wellFormed(x); ck(!bad, `3) ページ${i + 1}: XML ` + bad); });
}

// ===== 4) 熱烈歓迎（事前MTG。既定のひな形の会社名の枠：幅540pt・28pt・下限12pt）=====
{
  const HEAD = ['No', '業種区分', '氏名', 'ふりがな', 'カテゴリー', '会社名', '役職', 'メモ', '写真ファイル名', '一言コメント',
    '紹介してほしい人', '協業したい人', '入会日', '更新日', '更新期限日', '会社での役職'];
  const W = [['見本 一子', 'みほん いちこ', '税理士', '見本会計事務所'], ['見本 二子', 'みほん にこ', '行政書士', '㈱試験\nリーガルサービス'],
             ['見本 三子', 'みほん みこ', 'Web制作', CO.shrink], ['見本 四子', 'みほん よんこ', '司法書士', CO.long]];
  const rows = [HEAD].concat(W.map(([n, k, t, c], i) => HEAD.map((h) => ({ No: String(i + 1), 氏名: n, ふりがな: k, カテゴリー: t, 会社名: c }[h] || ''))));
  const DATE = '2026/10/07';
  const row = (c, role, v) => ['', '', c, '', '', role, '', '', '', v || '', '', '', '', '', '', ''];
  const ROUTINE = [['', '', '開催日', '', '', '', '', '', '', DATE, '', '', '', '', '', ''], ['', '', '定例会回数', '', '', '', '', '', '', '536', '', '', '', '', '', ''],
    ['', 'No', '内容', '', '', '担当', '期日', '曜日目安', '備考', '', '', '', '', '', '', ''],
    row('新入会', 'バイス', '見本 一子さん、見本 二子さん、見本 三子さん、見本 四子さん'), row('ウィークリープレゼン', 'プレジ', '')];
  const env = makeEnv({ now: new Date(2026, 9, 5, 10, 0, 0) });
  const G = Object.assign({}, env.globals);
  vm.createContext(G);
  for (const f of fs.readdirSync(ROOT).filter((x) => /\.js$/.test(x)).sort()) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), G, { filename: f });
  env.reset([['メンバー名簿', false, rows], ['休会日', true, [['2026/12/30']]], ['【24期】ルーティンチェックシート', false, ROUTINE]],
            { BNI_CHAPTER: J({ name: '見本', termBase: 23, meetingBaseDate: '2026/03/18', meetingBaseCount: 509 }) });
  G.HtmlService = { createHtmlOutputFromFile: (n) => ({ getContent: () => fs.readFileSync(path.join(ROOT, n + '.html'), 'utf8') }) };
  let saved = null;
  G.saveOutputFile_ = (b, name) => { saved = b; return { id: 'out', url: 'https://example/' + name, downloadUrl: 'https://example/dl/' + name }; };
  G.findPhotoIdForName_ = () => '';
  G.CacheService = undefined;
  G.saveRoleHolders({ president: '見本 一子' }, 24, DATE);
  const c0 = G.getRoleInputContext(DATE, 'president');
  const sid = c0.order.find((k) => c0.items[k].title === '今週の共有事項（プレジデント）');
  G.saveRoleInput(DATE, 'president', [{ id: sid, value: '・見本の共有事項です。', orig: '' }]);
  const data = G.premtgData_(new Date(2026, 9, 7));
  ck(J(data.newMembers.map((m) => m.company)) === J(['見本会計事務所', '株式会社試験リーガルサービス', CO.shrink, CO.long]),
     '4) 新入会の方の会社名（名簿の文字をそろえる）: ' + J(data.newMembers.map((m) => m.company)));
  const res = G.generatePreMeetingSlides(DATE);
  ck(res.ok && /熱烈歓迎 4枚/.test(res.message), '4) 熱烈歓迎のページを作れない: ' + J(res.message));
  const z = readZip(saved._buf), xml = (p) => (z[p] ? z[p].toString('utf8') : '');
  // 熱烈歓迎のページ（お名前の順）
  const order = ['一', '二', '三', '四'], who = (x) => order.indexOf((/見本 (.)子さん/.exec(G.slideText_(x)) || [])[1]);
  const welcome = Object.keys(z).filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p) && /熱烈歓迎/.test(xml(p)))
    .map((p) => xml(p)).sort((a, b) => who(a) - who(b));
  const want = [['見本 一子', '見本会計事務所', 'fit'], ['見本 二子', '株式会社試験リーガルサービス', 'fit'], ['見本 三子', CO.shrink, 'shrink'], ['見本 四子', CO.long, 'two']];
  ck(welcome.length === 4, '4) 熱烈歓迎のページの数: ' + welcome.length);
  welcome.forEach((x, i) => {
    const [nm, co, kind] = want[i], r = G.findShapeRange_(x, 7), seg = r ? x.substring(r.start, r.end) : '', inner = innerOf(seg);
    const e = expectOne(co, inner, 28, 12, 14), label = `4) 熱烈歓迎 ${nm}（${co}・内側 ${inner}pt）`;
    ck(G.slideText_(x).includes(nm + 'さん'), label + ': お名前');
    ck(e.kind === kind, `${label}: 作り物の会社名の長さが、確かめたい場合（${kind}）になっていない: ${e.kind}`);
    const got = checkCompany(label, seg, 0, e, inner, 28);
    if (kind === 'two') ck(got.length === 2 && e.pt >= 12 && e.pt <= 14, label + ': 2行の大きさ ' + e.pt);
    const bad = wellFormed(x);
    ck(!bad, label + ': XML ' + bad);
  });
  // できたpptxが壊れていない
  const f = path.join(os.tmpdir(), 'check_company_lines_' + process.pid + '.pptx');
  fs.writeFileSync(f, writeZip(Object.fromEntries(Object.entries(z))));
  const it = spawnSync('python3', [path.join(ROOT, 'tools', 'pptx_integrity.py'), f], { encoding: 'utf8' });
  fs.unlinkSync(f);
  ck(it.status === 0, '4) pptxとして壊れている: ' + (it.stdout || '') + (it.stderr || ''));
}

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.slice(0, 40).forEach((m) => console.log('  - ' + m));
  process.exit(1);
}
console.log('会社名の行（前半スライドの1人ぶんのページ・熱烈歓迎）: 検査 ' + checks + ' 件 OK: 新メンバー・更新メンバー／部門のページ／熱烈歓迎は'
  + '1行のまま小さく（枠の余白も引く）・下限でも入らなければ自然な2行・まとめのページは自然な2行まで・役職のメンバー紹介（1人のページ・差し込み口・'
  + 'グループの中）は1行か自然な2行・名簿の文字をそろえる・どの行も枠に入る・幅が近ければ同じ切れ目・XMLが壊れていない');

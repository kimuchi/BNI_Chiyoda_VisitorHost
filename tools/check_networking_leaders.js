// 定例会スライド（前半）のネットワーキングリーダーのページ（meeting_pages_srv.js の applyNetworkingLeaders_）を、作り物のページで確かめる。
// 雛形の実物は使わない。名前はすべて架空。
//
//   node tools/check_networking_leaders.js
//
// 作り物のページは、Activeチャプターの雛形と同じ並び：左の月桂樹の中に部門の見出し（CEU・サンキュー…）、その下に数、
// 右に写真とお名前・会社名・【カテゴリー】。外部リファーラルにはお2人のページもある。
//
// 確かめること
//   ・カタカナだけの見出し（「サンキュー」）を、お名前の枠と取り違えない
//     （以前は、受賞者がお2人だと見出しに1人目のお名前・数に1人目の会社名が入り、写真の枠には2人目だけが入った。
//       お1人でも、見出しにお名前が入り、写真の側が空になった）
//   ・お2人の部門は、お2人のページ（ほかの部門のもの）を写して、その部門の見出しと数にする。元の1人のページは非表示
//   ・お1人の部門は、その部門のページに入れる。見出しと数はそのまま（数だけその月の数）
const fs = require('fs');
const path = require('path');
const vm = require('vm');

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
const ROSTER = [
  { name: '見本 一郎', company: '一郎商事株式会社', title: '税理士' },
  { name: '試験 二郎', company: '株式会社二郎企画', title: '司法書士' },
  { name: '架空 三郎', company: '三郎工務店', title: '工務店' },
  { name: '仮名 四郎', company: '四郎保険', title: '生命保険（法人）' },
  { name: '例示 五郎', company: '五郎デザイン', title: 'Web制作' },
  { name: '模擬 六郎', company: '六郎食品', title: '贈答用生鮮食品' },
];
const sb = {
  console: { log() {}, warn() {}, error: console.error },
  Utilities: { newBlob: (c, t, n) => blob(c, t, n) },
  DriveApp: { getFileById: () => { throw new Error('写真は無い'); } },
  normName_: (s) => String(s == null ? '' : s).normalize('NFKC').replace(/[\s　]/g, ''),
  findPhotoIdForName_: () => '',
};
vm.createContext(sb);
for (const f of ['ooxml.js', 'chapter_srv.js', 'member_presen_srv.js', 'referral_srv.js', 'splice_srv.js', 'meeting_slides_srv.js',
                 'role_input_srv.js', 'role_intro_srv.js', 'routine_srv.js', 'meeting_pages_srv.js']) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), sb, { filename: f });
}
sb.getMemberMaster = () => ({ ok: true, members: ROSTER.map((m) => Object.assign({}, m)) });
const F = sb;

// ===== 作り物のページ =====
const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
  + 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const E = 12700;
const xf = (x, y, w, h) => `<a:xfrm><a:off x="${Math.round(x * E)}" y="${Math.round(y * E)}"/><a:ext cx="${Math.round(w * E)}" cy="${Math.round(h * E)}"/></a:xfrm>`;
const para = (t, sz) => `<a:p><a:r><a:rPr lang="ja-JP" sz="${sz}" b="1"/><a:t>${t}</a:t></a:r></a:p>`;
const sp = (id, x, y, w, h, lines, sz, title) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Text ${id}"/><p:cNvSpPr txBox="1"/>`
  + `${title ? '<p:nvPr><p:ph type="title"/></p:nvPr>' : '<p:nvPr/>'}</p:nvSpPr><p:spPr>${xf(x, y, w, h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`
  + `<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>${lines.map((t) => para(t, sz)).join('')}</p:txBody></p:sp>`;
const pic = (id, x, y, w, h, rid, name) => `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${name || '図 ' + id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>`
  + `<p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>`
  + `<p:spPr>${xf(x, y, w, h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
const slideXml = (inner) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld><p:spTree>`
  + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'
  + inner + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
const TITLE = sp(2, 180, 5, 760, 60, ['ネットワーキングリーダーの発表'], 4000, true);
// 1人のページ：左に月桂樹（絵）と、その中の部門の見出し・その下の数。右に写真・お名前・会社名・【カテゴリー】
const kindPage = (label, value) => TITLE
  + pic(3, 90, 110, 300, 300, 'rId2', '月桂樹') + sp(4, 140, 220, 200, 70, label, 3600) + sp(5, 80, 420, 320, 60, [value], 4000)
  + pic(6, 560, 110, 250, 300, 'rId3') + sp(7, 500, 420, 370, 50, ['見本 太郎'], 3600) + sp(8, 500, 470, 370, 70, ['見本株式会社', '【見本カテゴリー】'], 2400);
// お2人のページ：見出しと数は上、左右に写真・お名前・会社名・【カテゴリー】
const pairPage = (label, value) => TITLE
  + sp(4, 330, 75, 300, 60, label, 3200) + sp(5, 330, 135, 300, 50, [value], 3600)
  + pic(6, 110, 190, 230, 250, 'rId3') + sp(7, 60, 450, 330, 45, ['見本 太郎'], 3200) + sp(8, 60, 495, 330, 45, ['見本株式会社', '【見本カテゴリー】'], 2000)
  + pic(9, 620, 190, 230, 250, 'rId4') + sp(10, 570, 450, 330, 45, ['見本 花子'], 3200) + sp(11, 570, 495, 330, 45, ['見本株式会社', '【見本カテゴリー】'], 2000);
const SLIDES = {
  cover: sp(2, 100, 200, 700, 80, ['ようこそ　BNI 見本チャプター'], 4000),
  title: sp(2, 100, 120, 760, 80, ['2026年'], 5400) + sp(3, 100, 220, 760, 80, ['６月度'], 5400) + sp(4, 100, 320, 760, 80, ['ネットワーキングリーダーの発表'], 4400),
  ceu: kindPage(['CEU'], '25PT'),
  thanks: kindPage(['サンキュー'], '4,000万円'),
  ext: kindPage(['外部', 'リファーラル'], '17件'),
  oto: kindPage(['1to1'], '17件'),
  visitor: kindPage(['ビジター', '招待数'], '8名'),
  extPair: pairPage(['外部', 'リファーラル'], '17件'),
  weekly: sp(2, 0, 200, 960, 80, ['ウィークリープレゼンテーション'], 5400),
};
const ORDER = ['cover', 'title', 'ceu', 'thanks', 'ext', 'oto', 'visitor', 'extPair', 'weekly'];
function makeParts() {
  const parts = {};
  const put = (p, s) => { parts[p] = blob(s, 'application/xml', p); };
  put('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + ORDER.map((k, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('') + '</Types>');
  put('ppt/presentation.xml', `<?xml version="1.0"?><p:presentation ${NS}><p:sldIdLst>`
    + ORDER.map((k, i) => `<p:sldId id="${256 + i}" r:id="rId${100 + i}"/>`).join('') + '</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>');
  put('ppt/_rels/presentation.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + ORDER.map((k, i) => `<Relationship Id="rId${100 + i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`).join('')
    + '</Relationships>');
  ORDER.forEach((k, i) => {
    put(`ppt/slides/slide${i + 1}.xml`, slideXml(SLIDES[k]));
    put(`ppt/slides/_rels/slide${i + 1}.xml.rels`, '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + [2, 3, 4].map((n) => `<Relationship Id="rId${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${i + 1}_${n}.png"/>`).join('')
      + '</Relationships>');
  });
  return parts;
}
const W = 12192000, H = 6858000;
const shapeText = (xml, id) => { const r = F.findShapeRange_(xml, id); return r ? F.slideText_(xml.substring(r.start, r.end)) : null; };
// 表示のページのうち、その部門のページ：[{ path, xml, info }]
function kindPages(parts, key) {
  return F.slideOrder_(parts).map((p) => ({ path: p, xml: F.xmlOf_(parts, p) }))
    .filter((pg) => !/<p:sld\b[^>]*\sshow="0"/.test(pg.xml))
    .map((pg) => Object.assign(pg, { info: F.fpNlPage_(pg.xml, W, H) }))
    .filter((pg) => pg.info.type === 'kind' && pg.info.kind === key);
}
const unitNames = (pg) => pg.info.units.map((u) => shapeText(pg.xml, u.name.id));
const winner = (name) => ({ name, raw: name + 'さん', category: '' });

// ===== 0. 雛形のページの見分け方：月桂樹の中の見出し・その下の数は、お名前の枠ではない =====
{
  const parts = makeParts();
  const inf = (k) => F.fpNlPage_(F.xmlOf_(parts, `ppt/slides/slide${ORDER.indexOf(k) + 1}.xml`), W, H);
  ['ceu', 'thanks', 'ext', 'oto', 'visitor'].forEach((k) => {
    const i = inf(k);
    ck(i.type === 'kind' && i.units.length === 1 && i.units[0].name.id === '7',
       `0) ${k} のページ：お名前の枠は右の1つだけのはず: ${J({ type: i.type, units: (i.units || []).map((u) => u.name.id) })}`);
  });
  const p = inf('extPair');
  ck(p.type === 'kind' && p.kind === 'ext' && J(p.units.map((u) => u.name.id)) === J(['7', '10']), '0) お2人のページの枠: ' + J(p.units && p.units.map((u) => u.name.id)));
}

// ===== 1. お2人の部門（サンキュー・1to1）とお1人の部門 =====
{
  const parts = makeParts();
  const nl = { show: true, month: '2026-09', items: [
    { key: 'ceu', value: '30', unit: 'PT', winners: [winner('見本 一郎')] },
    { key: 'thanks', value: '5,000万円', unit: '', winners: [winner('試験 二郎'), winner('架空 三郎')] },
    { key: 'ext', value: '12', unit: '件', winners: [winner('仮名 四郎')] },
    { key: 'oto', value: '21', unit: '回', winners: [winner('例示 五郎'), winner('模擬 六郎')] },
    { key: 'visitor', value: '3', unit: '名', winners: [winner('見本 一郎')] },
  ] };
  const res = F.applyNetworkingLeaders_(parts, nl, { by: {}, seq: 0 });
  ck(res && /お2人のページを作った部門: サンキュー、1to1/.test(res.message), '1) 知らせ: ' + (res && res.message));
  // サンキュー：お2人のページに、見出し「サンキュー」・数・お2人。見出しや数の枠にお名前・会社名が入らない
  const th = kindPages(parts, 'thanks');
  ck(th.length === 1, '1) サンキューの表示のページが1枚でない（見出しがお名前に替わると、部門のページと見分けられない）: ' + th.length);
  th.forEach((pg) => {
    ck(J(unitNames(pg)) === J(['試験 二郎', '架空 三郎']), '1) サンキューのページのお2人: ' + J(unitNames(pg)));
    ck(shapeText(pg.xml, pg.info.label.id) === 'サンキュー', '1) サンキューの見出し: ' + shapeText(pg.xml, pg.info.label.id));
    ck(pg.info.value && shapeText(pg.xml, pg.info.value.id) === '5,000万円', '1) サンキューの数: ' + (pg.info.value && shapeText(pg.xml, pg.info.value.id)));
    ck(!/見本 太郎|見本 花子|見本株式会社/.test(F.slideText_(pg.xml)), '1) サンキューのページに見本の字が残った: ' + F.slideText_(pg.xml));
    ck(/株式会社二郎企画/.test(F.slideText_(pg.xml)) && /三郎工務店/.test(F.slideText_(pg.xml)), '1) サンキューのページの会社名: ' + F.slideText_(pg.xml));
  });
  // 1to1 も同じく、お2人のページ
  const oto = kindPages(parts, 'oto');
  ck(oto.length === 1 && J(unitNames(oto[0])) === J(['例示 五郎', '模擬 六郎']) && shapeText(oto[0].xml, oto[0].info.label.id) === '1to1'
     && shapeText(oto[0].xml, oto[0].info.value.id) === '21件',
     '1) 1to1 のお2人のページ: ' + J(oto.map((pg) => [unitNames(pg), shapeText(pg.xml, pg.info.label.id), shapeText(pg.xml, pg.info.value.id)])));
  // お1人の部門：その部門のページに、見出しはそのまま
  [['ceu', 'CEU', '30PT', '見本 一郎'], ['ext', '外部リファーラル', '12件', '仮名 四郎'], ['visitor', 'ビジター招待数', '3名', '見本 一郎']].forEach(([k, label, value, nm]) => {
    const pgs = kindPages(parts, k);
    ck(pgs.length === 1 && J(unitNames(pgs[0])) === J([nm]) && shapeText(pgs[0].xml, pgs[0].info.label.id) === label
       && shapeText(pgs[0].xml, pgs[0].info.value.id) === value,
       `1) ${k} のページ: ` + J(pgs.map((pg) => [unitNames(pg), shapeText(pg.xml, pg.info.label.id), shapeText(pg.xml, pg.info.value.id)])));
  });
  // 使わなかったページ（サンキュー・1to1 の1人のページ、外部リファーラルのお2人のページ）は非表示
  const hidden = F.slideOrder_(parts).filter((p) => /<p:sld\b[^>]*\sshow="0"/.test(F.xmlOf_(parts, p))).map((p) => ORDER[+p.match(/slide(\d+)\.xml/)[1] - 1] || 'copy');
  ck(J(hidden.sort()) === J(['extPair', 'oto', 'thanks']), '1) 非表示のページ: ' + J(hidden));
}

// ===== 2. サンキューがお1人：サンキューのページに入れ、見出しにお名前を入れない =====
{
  const parts = makeParts();
  const nl = { show: true, month: '2026-09', items: [{ key: 'thanks', value: '800万円', unit: '', winners: [winner('試験 二郎')] }] };
  F.applyNetworkingLeaders_(parts, nl, { by: {}, seq: 0 });
  const th = kindPages(parts, 'thanks');
  ck(th.length === 1 && J(unitNames(th[0])) === J(['試験 二郎']) && shapeText(th[0].xml, th[0].info.label.id) === 'サンキュー'
     && shapeText(th[0].xml, th[0].info.value.id) === '800万円',
     '2) サンキュー（お1人）: ' + J(th.map((pg) => [unitNames(pg), shapeText(pg.xml, pg.info.label.id), shapeText(pg.xml, pg.info.value.id)])));
  ck(!/試験 二郎/.test(shapeText(F.xmlOf_(parts, 'ppt/slides/slide4.xml'), 4) || ''), '2) 月桂樹の中の見出しに、お名前が入った');
}

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('ネットワーキングリーダーのページ: 検査 ' + checks + ' 件 OK: 見出し（サンキュー）をお名前の枠と取り違えない・'
  + 'お2人の部門はお2人のページ・お1人の部門は見出しと数をそのまま');

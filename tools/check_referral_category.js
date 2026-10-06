// リファーラル発表（とメンバーのページ）のカテゴリーが、2行になっても・名簿に改行があっても見えるか。
// 作り物のページ（雛形の実物は使わない）で、referral_srv.js の rfBuildSlide_ と member_presen_srv.js の
// fitPresenterCategory_ を動かして確かめる。名前はすべて架空。
//
//   node tools/check_referral_category.js
//
// 確かめること
//   ・1行のカテゴリーは、雛形どおり32pt（下のカウントダウンの数字の箱に枠が少しかかっていても、小さくしない。
//     2026-10-06c では、カウントダウンの箱を「下の図形」と見て、1行のカテゴリーまで小さくしていた）
//   ・2行のときは、カテゴリーを前面に出し（カウントダウンの白い箱に隠れないように）、数字の字にかからないところまで使う
//     （白い箱の上の余白にはかかってよい）。かかるなら、カウントダウンを下げ（「次の発表者」の帯の上まで）、まだかかるなら
//     数字の箱の上側を縮めて数字を下げる（箱は同じ形のまま・数字の字が隠れる高さより低くしない）。足りないぶんだけ少し小さく（24ptまで）。
//     会社名が2行でカテゴリーが下がった方も、32ptの2行のまま（2026-10-06e までは 24pt 前後まで小さくなっていた）
//   ・下に何も無ければ、32ptの2行のまま、枠の高さを2行ぶんにする（以前は雛形の1行ぶんのまま、2行目が枠の外へ）
//   ・カテゴリーが載っている色の帯（や帯の絵）の下端・下にある文字（「今週のポジティブな貢献は？」など）に届くなら、
//     届かない大きさまで小さくする（2行のまま／1行にして、の大きい方）。字は1つも落とさない
//   ・名簿の改行（「デジタル広告制作」⏎「(ホームページ・動画)」）は外して1行の字にする（画面の古い版から来たときも）
//   ・はみ出した字を切る設定（vertOverflow="clip"）は外す
//   ・役職のメンバー紹介の読み方（riShapes_）が無いところ（メンバープレゼンの検査）でも動く
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const J = (x) => JSON.stringify(x);

function sandbox(files) {
  const sb = { console: { log() {}, warn() {}, error() {} } };
  vm.createContext(sb);
  for (const f of files) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), sb, { filename: f });
  // カウントダウンは、この検査の外（tools/check_countdown.js）で確かめる
  sb.mpSetCountdown_ = (x) => x;
  sb.mpNoAutoAdvance_ = (x) => x;
  return sb;
}
const F = sandbox(['ooxml.js', 'chapter_srv.js', 'member_presen_srv.js', 'referral_srv.js', 'role_intro_srv.js']);

// ===== 作り物のページ（EMU）=====
const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
  + 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const W = 12192000, H = 6858000;
const xf = (x, y, cx, cy) => `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`;
const sp = (id, [x, y, cx, cy], text, sz, body) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Text ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`
  + `<p:spPr>${xf(x, y, cx, cy)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>`
  + `<p:txBody>${body || '<a:bodyPr wrap="square"/>'}<a:lstStyle/><a:p><a:r><a:rPr lang="ja-JP" sz="${sz}" b="1"/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
const band = (id, [x, y, cx, cy]) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="帯 ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>`
  + `<p:spPr>${xf(x, y, cx, cy)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="C8102E"/></a:solidFill></p:spPr></p:sp>`;
const pic = (id, [x, y, cx, cy], name) => `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${name || '図 ' + id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>`
  + `<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xf(x, y, cx, cy)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
const CAT = [4830481, 4071101, 7311519, 584404];         // カテゴリーの枠（雛形は1行ぶんの高さ）
// カウントダウン：数字を書いた白い箱を重ねたもの（同じ大きさ。1つめの箱に消える動き）。カテゴリーの1行ぶんの枠に少しかかる高さ
const CD = [6000000, 4400000, 4000000, 1500000];
const countdown = (ids) => ids.map((id, i) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="数字 ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`
  + `<p:spPr>${xf(...CD)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:spPr>`
  + `<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="9600" b="1"/><a:t>${ids.length - i}</a:t></a:r></a:p></p:txBody></p:sp>`).join('');
const TIMING = (id) => `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek">`
  + `<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>`
  + `<p:par><p:cTn id="4" presetID="10" presetClass="exit" presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>`
  + `<p:set><p:cBhvr><p:cTn id="5" dur="1" fill="hold"><p:stCondLst><p:cond delay="999"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${id}"/></p:tgtEl>`
  + `<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="hidden"/></p:to></p:set>`
  + `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
const page = ({ under = '', below = '', catBody, cd = false } = {}) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld><p:spTree>`
  + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'
  + under
  + sp(2, [600000, 200000, 11000000, 700000], 'REFERRAL PRESENTATION', 4000)
  + pic(3, [500000, 1300000, 3800000, 4400000], '写真')
  + sp(4, [4830858, 1500000, 7000000, 900000], '見本 太郎', 5400)
  + sp(5, [4830858, 2776058, 6893161, 1446550], '見本株式会社', 4400)
  + sp(6, CAT, '【見本カテゴリー】', 3200, catBody)
  + below
  + sp(7, [4830858, 6200000, 1500000, 500000], 'NEXT➡', 2400)
  + sp(8, [6500000, 6200000, 3000000, 500000], '次の 方', 2400)
  + (cd ? countdown([40, 41, 42, 43, 44, 45, 46]) : '')
  + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' + (cd ? TIMING(40) : '') + '</p:sld>';
const RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>';

const TWO = ['【デジタル広告制作(ホームページ・', '動画・SNS運用)】'];
const item = (lines, o) => Object.assign({ name: '見本 太郎', companyLines: ['見本株式会社'], companyPt: 44,
  categoryTop: CAT[1], categoryLines: lines, categoryPt: 32, categoryTight: lines.length > 1, nextName: '次の 方', seconds: 7, auto: false }, o || {});
const build = (xml, it) => {
  const SH = F.presenterShapes_(xml, W);
  SH.W = W; SH.H = H;
  return F.rfBuildSlide_(xml, RELS, it, null, SH).xml;
};
const catOf = (xml) => {
  const r = F.findShapeRange_(xml, 6), seg = xml.substring(r.start, r.end), g = F.readShapeGeomEmu_(xml, 6);
  const paras = F.findTagRanges_(seg, 'a:p').map((p) => F.slideText_(seg.substring(p.start, p.end))).filter((t) => t.trim());
  const sz = [...seg.matchAll(/<a:rPr\b[^>]*\ssz="(\d+)"/g)].map((m) => +m[1] / 100);
  return { paras, pt: Math.max(...sz), top: g.y, bottom: g.y + g.cy, cy: g.cy, body: (seg.match(/<a:bodyPr\b[^>]*>/) || [''])[0] };
};
const LINE = 12700 * 1.213, PAD = 91440;
const need = (n, pt, sp) => PAD + n * pt * LINE * (n > 1 ? sp : 1);
const lastShapeId = (xml) => { const ids = [...xml.replace(/<p:timing>[\s\S]*$/, '').matchAll(/<p:cNvPr id="(\d+)"/g)].map((m) => m[1]); return ids[ids.length - 1]; };

// ===== 0. 1行のカテゴリー：下にカウントダウンの数字の箱がかかっていても、雛形どおり32pt（いただいた画面の件）=====
{
  ['【デジタル広告制作(ホームページ・動画)】', '【生命保険(法人)】'].forEach((t) => {
    const x = build(page({ cd: true }), item([t]));
    const c = catOf(x), g = F.readShapeGeomEmu_(x, 40);
    ck(J(c.paras) === J([t]) && c.pt === 32, '0) 1行のカテゴリーを小さくした: ' + J(c));
    ck(g.y === CD[1], '0) 1行なのにカウントダウンを動かした: ' + g.y);
  });
}

// ===== 0b. 2行のカテゴリー：前面に出し、数字の字にかからないところまで使う（白い箱の上の余白にはかかってよい）。
//          かかるなら、カウントダウンを下げ（次の発表者の帯まで）、まだかかるなら数字の箱の上側を縮めて数字を下げる。
//          会社名が2行でカテゴリーが下がった方（いただいた画面の件）も、32ptの2行のまま =====
{
  const EM = 96 * 12700, NEXT_TOP = 6200000;
  const look = (x) => {
    const cd = F.readShapeGeomEmu_(x, 40);
    return { cd, mid: cd.y + cd.cy / 2, same: [41, 42, 43, 44, 45, 46].every((id) => J(F.readShapeGeomEmu_(x, id)) === J(cd)) };
  };
  const x = build(page({ cd: true }), item(TWO)), c = catOf(x), k = look(x);
  ck(lastShapeId(x) === '6', '0b) 2行のカテゴリーが前面にない（カウントダウンの白い箱に隠れる）: ' + lastShapeId(x));
  ck(J(c.paras) === J(TWO) && c.pt === 32, '0b) 2行・32ptのままでない（会社名が2行の方のカテゴリーが小さくなる）: ' + J(c));
  ck(k.same && k.mid > CD[1] + CD[3] / 2, '0b) カウントダウン（全部の箱が同じ形のまま）を下げていない: ' + J(k));
  ck(k.cd.y + k.cd.cy <= NEXT_TOP - 25400, '0b) カウントダウンの箱が「次の発表者」の帯に重なった: ' + J(k.cd));
  // 数字の箱は、上下 0.45em ずつ（0.9em）より低くしない（上の箱の白で、下の箱の数字を隠しきる）。自動で縮める設定は外す
  ck(k.cd.cy >= 0.9 * EM - 1 && !/normAutofit|spAutoFit/.test(x.match(/<p:cNvPr id="40"[\s\S]*?<\/p:sp>/)[0]),
     '0b) 数字の箱を、数字の字が隠れない高さまで縮めた（下の数字が見える）・自動で縮める設定が残った: ' + J(k.cd));
  // 数字の字（真ん中から上へ 0.39em。Impact など背の高い書体でも）にかからない
  ck(c.bottom <= k.mid - 0.39 * EM, '0b) 2行目がカウントダウンの数字にかかる: ' + J({ bottom: c.bottom, digitTop: Math.round(k.mid - 0.39 * EM) }));
  // 下げる余白が十分なら、箱は縮めない（下げるだけ）
  const roomy = page({ cd: true }).replace(/<a:off x="(4830858|6500000)" y="6200000"\/>/g, '<a:off x="$1" y="6600000"/>');
  const x2 = build(roomy, item(TWO)), c2 = catOf(x2), k2 = look(x2);
  ck(c2.pt === 32 && J(c2.paras) === J(TWO) && k2.cd.cy === CD[3] && k2.same, '0b) 下げる余白があるのに小さくした・箱を縮めた: ' + J({ c2, cd: k2.cd }));
  // 下げる余白も、縮める余白も足りない（次の発表者がすぐ下）：数字にかからない大きさまで少し小さく（24ptまで）
  const tight = page({ cd: true }).replace(/<a:off x="(4830858|6500000)" y="6200000"\/>/g, '<a:off x="$1" y="6000000"/>');
  const x3 = build(tight, item(TWO)), c3 = catOf(x3), k3 = look(x3);
  ck(J(c3.paras) === J(TWO) && c3.pt >= 24 && c3.pt < 32 && c3.bottom <= k3.mid - 0.39 * EM && k3.cd.y + k3.cd.cy <= 6000000 - 25400 && k3.cd.cy >= 0.9 * EM - 1,
     '0b) 余白が足りないとき: ' + J({ c3, cd: k3.cd, digitTop: Math.round(k3.mid - 0.39 * EM) }));
}

// ===== 1. 下に何も無い：32ptの2行のまま、枠を2行ぶんの高さに =====
{
  const c = catOf(build(page(), item(TWO)));
  ck(J(c.paras) === J(TWO) && c.pt === 32, '1) 2行・32ptのままでない: ' + J(c));
  ck(c.cy >= need(2, 32, 0.85) - 2 && c.bottom < 6200000, '1) 枠が2行ぶんの高さになっていない（2行目が枠の外）: ' + J(c));
  const one = catOf(build(page(), item(['【税理士】'])));
  ck(J(one.paras) === J(['【税理士】']) && one.pt === 32 && one.cy === CAT[3], '1) 1行のカテゴリーが変わった: ' + J(one));
}

// ===== 2. カテゴリーが色の帯の上にある：帯の下端を越えない（字は落とさない）=====
{
  const BAND = [4700000, 1400000, 7450000, 3400000];      // 下端 4,800,000
  const c = catOf(build(page({ under: band(20, BAND) }), item(TWO)));
  ck(c.bottom <= BAND[1] + BAND[3] - 25400, '2) 帯の下端を越えた（帯の外の字は見えない）: ' + J(c));
  ck(c.paras.join('') === TWO.join('') && c.pt >= 16 && c.pt < 32, '2) 字が落ちた・小さすぎる: ' + J(c));
  // 帯が絵のときも同じ
  const p = catOf(build(page({ under: pic(21, BAND, '帯の絵') }), item(TWO)));
  ck(p.bottom <= BAND[1] + BAND[3] - 25400 && p.paras.join('') === TWO.join(''), '2) 帯の絵の下端を越えた: ' + J(p));
}

// ===== 3. 下に別の文字（「今週のポジティブな貢献は？」）：2行目がその上端を越えない =====
{
  const ASK = [5000000, 4750000, 5900000, 500000];
  const c = catOf(build(page({ below: sp(30, ASK, '今週のポジティブな貢献は？', 2800) }), item(TWO)));
  ck(c.bottom <= ASK[1] - 25400, '3) 下の文字に重なった: ' + J(c));
  ck(c.paras.join('') === TWO.join('') && c.pt >= 16, '3) 字が落ちた・小さすぎる: ' + J(c));
  // 1行にした方が大きく入るときは1行に（狭い高さ・短いカテゴリー）
  const SHORT = ['【見本カテゴリー', 'の説明】'];
  const t = catOf(build(page({ below: sp(30, [5000000, 4560000, 5900000, 400000], '今週のポジティブな貢献は？', 2800) }), item(SHORT)));
  ck(t.paras.length === 1 && t.paras[0] === SHORT.join('') && t.bottom <= 4560000 - 25400, '3) 1行の方が大きく入るのに2行のまま: ' + J(t));
}

// ===== 4. 名簿の改行（画面の古い版から来た行）：外して1行の字に =====
{
  const c = catOf(build(page(), item(['【デジタル広告制作\n(ホームページ・動画)】'], { name: '見本\n太郎', companyLines: ['見本\r\n株式会社'] })));
  ck(J(c.paras) === J(['【デジタル広告制作(ホームページ・動画)】']), '4) カテゴリーの改行が残った: ' + J(c.paras));
  const x = build(page(), item(['【A\nB】'], { name: '見本\n太郎', companyLines: ['Web\nDesign Inc.'] }));
  ck(!/<a:t>[^<]*\n[^<]*<\/a:t>/.test(x) && x.includes('<a:t>Web Design Inc.</a:t>') && x.includes('<a:t>見本太郎</a:t>'),
     '4) 氏名・会社名の改行が残った（英数字のあいだは空白に）: ' + (x.match(/<a:t>[^<]*<\/a:t>/g) || []).slice(0, 6).join(' '));
  ck(J([F.slideOneLine_('Web\r\nDesign'), F.slideOneLine_('税理士\n'), F.slideOneLine_('普通の\t名前')]) === J(['Web Design', '税理士', '普通の 名前']), '4) 改行の外し方');
}

// ===== 5. はみ出した字を切る設定は外す =====
{
  const c = catOf(build(page({ catBody: '<a:bodyPr wrap="square" vertOverflow="clip" horzOverflow="clip"/>' }), item(TWO)));
  ck(!/Overflow=/.test(c.body), '5) はみ出した字を切る設定が残った: ' + c.body);
}

// ===== 6. 役職のメンバー紹介の読み方（riShapes_）が無いところでも動く（メンバープレゼンの検査と同じ読み込み）=====
{
  const G = sandbox(['ooxml.js', 'member_presen_srv.js']);
  const xml = page({ under: band(20, [4700000, 1400000, 7450000, 3400000]) });
  const x = G.fitPresenterCategory_(G.setLineSpacingInShape_(G.setParagraphsInShape_(xml, 6, TWO), 6, 85), 6, W, H);
  const g = G.readShapeGeomEmu_(x, 6);
  ck(g.y + g.cy <= 4800000 - 25400, '6) riShapes_ が無いと、帯の下端を見ない: ' + J(g));
}

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('リファーラル発表のカテゴリー: 検査 ' + checks + ' 件 OK: 2行なら枠を2行ぶんに・帯の外や下の文字に隠れない大きさに（字は落とさない）・'
  + '名簿の改行を外す・はみ出しを切る設定を外す');

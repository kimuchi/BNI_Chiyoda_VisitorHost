// リファーラル発表（とメンバーのページ）のカテゴリーが、2行になっても・名簿に改行があっても見えるか。
// 作り物のページ（雛形の実物は使わない）で、referral_srv.js の rfBuildSlide_ と member_presen_srv.js の
// fitPresenterCategory_ を動かして確かめる。名前はすべて架空。
//
//   node tools/check_referral_category.js
//
// 確かめること
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
const page = ({ under = '', below = '', catBody } = {}) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld><p:spTree>`
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
  + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
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

// ===== 3. 下に別の文字（「今週のポジティブな貢献は？」）：その上端を越えない =====
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

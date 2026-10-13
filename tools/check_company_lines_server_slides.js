// サーバーで作るスライドの会社名（ビジタープレゼンのページ・お2人のページ）が、どの出力でも同じ会社名の改行の決め方
// （company_break.js の CB_）を使っているかを、作り物の会社名で確かめる。会社名・お名前はすべて架空。
//
//   node tools/check_company_lines_server_slides.js
//
// 確かめること
//   1) ビジタープレゼンのページ（slides_visitor_srv.js の buildPresenXml_ → ooxml.js の fitStackFixed_）。
//      ひな形は tools/check_visitor_slides.js の 4b と同じ並び（お名前 → 会社名 → 【カテゴリー】→ カウントダウン）
//      ・枠の幅は、左右の余白を除いた内側の幅で測る
//      ・会社名を名簿の書き方にそろえる（セルの中の改行を外す・㈱を「株式会社」に）
//      ・カウントダウンまでの余地が十分あるとき：行の分け方・大きさが CB_.layout（内側の幅・44pt・16ptまで・2行まで）と同じ。
//        語の途中で折らない（以前は「見本インターナショ／ナル…」「見本／コンサルティングセンター」など）。
//        幅の近いメンバーのページ（528pt）と同じ切れ目
//      ・どの行も、CB_ の見積もりで枠の内側に入る（1行のときも。英字の多い会社名を PowerPoint が自分で折り返さない）
//      ・余地が少ないとき：【カテゴリー】がカウントダウンに重ならない。法人の種類・語の途中で折らない。
//        1行で入る短い会社名を、余地のために小さな2行に分けない
//   2) お2人のページ（meeting_slides_srv.js：メインプレゼン・抽選・後半の1枚目・推薦のことば）
//      ・会社名は1行のまま。名簿の書き方にそろえて入れる（改行が無い・㈱を開く）。画面から改行の残った会社名が来ても
//      ・縮めた大きさで、CB_.lineWidthPt の幅が枠の内側の幅に入る（英字の多い会社名も）。入る会社名は縮めない
//      ・お名前・カテゴリーの縮め方は以前のまま
process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeEnv } = require('./lib_sheet_fake');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const J = (x) => JSON.stringify(x);

const env = makeEnv({ now: new Date(2026, 9, 5, 10, 0, 0) });
const F = Object.assign({}, env.globals);
vm.createContext(F);
for (const f of fs.readdirSync(ROOT).filter((x) => /\.js$/.test(x)).sort()) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), F, { filename: f });
env.reset([['メンバー名簿', false, [['No', '業種区分', '氏名'], ['1', '', '見本 一郎'], ['2', '', '試験 花子']]]], {});
F.findPhotoIdForName_ = () => '';
const CB = F.CB_;

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
  + 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const slideXml = (inner) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld><p:spTree>`
  + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' + inner
  + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
const shapeOf = (x, id) => (x.match(new RegExp('<p:cNvPr id="' + id + '"[\\s\\S]*?</p:sp>')) || [''])[0];
const sizesOf = (seg) => [...seg.matchAll(/\ssz="(\d+)"/g)].map((m) => +m[1] / 100);
const textsOf = (seg) => [...seg.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]);
const parasOf = (seg) => (seg.match(/<a:p>[\s\S]*?<\/a:p>/g) || []).map((p) => textsOf(p).join(''));

// 会社名として不自然な切れ目（法人の種類の途中・行の頭に来てはいけない字・カタカナ語の途中）
const CORPS = ['株式会社', '有限会社', '合同会社', '一般社団法人', '特定非営利活動法人', '税理士法人'];
const KATA = ['インターナショナル', 'トレーディング', 'コンサルティング', 'テクノロジーズ', 'ホールディングス', 'エンジニアリング', 'オフィス', 'センター', 'デザイン'];
const NO_START = /^[ー‐・･、。，．,.：:；;！？!?）)」』】〕〉》］\]々ゝゞヽヾぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶ％%]/;
function unnatural(lines, company) {
  const norm = CB.normalize(company);
  for (let i = 1; i < lines.length; i++) {
    if (NO_START.test(lines[i])) return '行の頭に来てはいけない字: ' + lines[i];
    const joint = lines[i - 1] + '|' + lines[i];
    for (const w of CORPS.concat(KATA)) for (let k = 1; k < w.length; k++) if (joint.includes(w.slice(0, k) + '|' + w.slice(k))) return '語の途中: ' + w;
    // 英単語・数字の途中（空白の位置で折ったときは、空白が消える）
    if (/[A-Za-z0-9]$/.test(lines[i - 1]) && /^[A-Za-z0-9]/.test(lines[i]) && norm.includes(lines[i - 1] + lines[i])) return '英単語の途中';
  }
  return '';
}

// ===== 1) ビジタープレゼンのページ =====
const E = 12700;
const box = (id, y, h, text, sz, w) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="枠 ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`
  + `<p:spPr><a:xfrm><a:off x="${120 * E}" y="${y * E}"/><a:ext cx="${(w || 620) * E}" cy="${h * E}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`
  + `<p:txBody><a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="ja-JP" sz="${sz}" b="1"/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
// countdown … カウントダウンの上端（pt）。300 は余地が少ない（【カテゴリー】の下から43pt）、420 は十分
const presenTpl = (countdown, width) => slideXml(box(25, 100, 60, 'お名前　様', 5400) + box(27, 170, 45, '会社名', 3200, width)
  + box(29, 217, 40, '【カテゴリー】', 2800) + box(40, countdown, 90, '20', 6000, 200));
const geo = (x, id) => F.readShapeGeomEmu_(x, id);
const bottomOf = (x, id) => (geo(x, id).y + geo(x, id).cy) / E;
const presen = (tpl, company) => {
  const x = F.buildPresenXml_(tpl, { name: '見本 来人', company, category: '税理士' }, 0), seg = shapeOf(x, 27);
  return { x, lines: parasOf(seg), pt: sizesOf(seg)[0], cat: sizesOf(shapeOf(x, 29)), bottom: bottomOf(x, 29) };
};
const PRESEN = { pt: 44, minPt: 16, maxLines: 2 };
const MEMBER = { width: 528, pt: 44, minPt: 28, hardMinPt: 20, maxLines: 3, linePenalty: 10 };   // メンバーのページ（tools/check_company_break.js）
const NAMES = [
  '見本商事', '見本コンサルティング株式会社', '見本コンサルティングオフィス株式会社', '株式会社見本不動産コンサルティングオフィス',
  '見本インターナショナルトレーディング研究所', 'ミホンインターナショナルトレーディングセンター', '見本不動産コンサルティングオフィス 東京支社',
  '一般社団法人日本見本協会', '特定非営利活動法人見本子ども支援センター', '見本建築設計事務所一級建築士事務所', 'Mihon Design Studio合同会社',
  'SAMPLE CREATIVE WORKS', 'SAMPLEテクノロジーズ研究所', 'SAMPLECREATIVEWORKSLIMITED', '株式会社SAMPLE Creative Works', '㈱見本\nエンジニアリング',
  '税理士法人見本会計事務所', '見本ホールディングスセンター', '見本コンサルティングオフィスとても長い名前株式会社', 'SAMPLE\nWORKS株式会社',
];
{
  // 枠の幅：左右の余白（既定 0.1インチ＝7.2pt ずつ）を除いた内側
  const roomy = presenTpl(420), inner = F.readShapeTextBox_(roomy, 27).widthPt;
  ck(Math.abs(inner - (620 - 14.4)) < 1e-6, '1) 会社名の枠の内側の幅が 605.6pt でない: ' + inner);
  const noIns = slideXml(box(27, 170, 45, '会社名', 3200).replace('<a:bodyPr wrap="square">', '<a:bodyPr wrap="square" lIns="0" rIns="0">'));
  ck(Math.abs(F.readShapeTextBox_(noIns, 27).widthPt - 620) < 1e-6, '1) 余白 0 の枠の幅が 620pt でない: ' + F.readShapeTextBox_(noIns, 27).widthPt);

  // 余地が十分あるとき：CB_.layout と同じ行・大きさ。どの行も枠に入る。語の途中で折らない
  NAMES.forEach((c) => {
    const r = presen(roomy, c), want = CB.layout(c, Object.assign({ width: inner }, PRESEN));
    const exp = want.fallback ? null : want;   // 2行に分けられない（英字だけの長い語）ときは、1行のまま小さくする
    if (exp) ck(J(r.lines) === J(exp.lines) && r.pt === exp.pt, `1) 「${c}」が CB_.layout と違う: ${J([r.lines, r.pt])}（${J([exp.lines, exp.pt])} のはず）`);
    else ck(r.lines.length === 1 && r.pt < 44 && r.pt >= 16, `1) 「${c}」（分けられない語）を1行のまま小さくしない: ${J([r.lines, r.pt])}`);
    ck(r.lines.every((l) => CB.fits(l, r.pt, inner)), `1) 「${c}」の行が枠の内側に入らない（PowerPoint が折り返す）: ${J([r.lines, r.pt])}`);
    ck(r.lines.join('').replace(/[ 　]/g, '') === CB.normalize(c).replace(/[ 　]/g, ''), `1) 「${c}」の字が欠けた・増えた: ${J(r.lines)}`);
    ck(!/\n/.test(r.lines.join('')) && !/㈱/.test(r.lines.join('')), `1) 「${c}」：改行・㈱が残った: ${J(r.lines)}`);
    const bad = unnatural(r.lines, c);
    ck(!bad, `1) 「${c}」が不自然な所で折れた（${bad}）: ${J(r.lines)}`);
    ck(r.cat.every((v) => v === 32) && r.bottom <= 420, `1) 「${c}」：【カテゴリー】が32ptでない・カウントダウンに重なる: ${J([r.cat, r.bottom])}`);
  });
  // 自然な切れ目（以前の決め方では、語の途中・片寄った所で折れていたもの）
  [
    ['見本インターナショナルトレーディング研究所', ['見本インターナショナル', 'トレーディング研究所'], 44],   // 以前は「見本インターナショ／ナル…」34pt
    ['見本ホールディングスセンター', ['見本ホールディングス', 'センター'], 44],                             // 以前は「見本／ホールディングスセンター」
    ['株式会社見本不動産コンサルティングオフィス', ['株式会社見本不動産', 'コンサルティングオフィス'], 44],
    ['特定非営利活動法人見本子ども支援センター', ['特定非営利活動法人', '見本子ども支援センター'], 44],
    ['Mihon Design Studio合同会社', ['Mihon Design Studio', '合同会社'], 44],
    ['㈱見本\nエンジニアリング', ['株式会社', '見本エンジニアリング'], 44],                                 // 名簿のセルの改行と ㈱
    ['見本商事', ['見本商事'], 44],
    ['一般社団法人日本見本協会', ['一般社団法人日本見本協会'], 44],
  ].forEach(([c, lines, pt]) => {
    const r = presen(roomy, c);
    ck(J(r.lines) === J(lines) && r.pt === pt, `1) 「${c}」→ ${J([r.lines, r.pt])}（${J([lines, pt])} のはず）`);
  });
  // 社名の部分が44ptで1行に入らないときは、語の途中で折らず、少し小さくして「社名／株式会社」（check_visitor_slides.js の 4b と同じ）
  {
    const r = presen(roomy, '見本コンサルティングオフィス株式会社');
    ck(J(r.lines) === J(['見本コンサルティングオフィス', '株式会社']) && r.pt < 44 && r.pt >= 36, '1) 長い社名の2行: ' + J([r.lines, r.pt]));
  }
  // 英字の多い会社名：以前の見積もり（半角0.5字）では1行に入ることになり、PowerPoint が自分で折り返していた
  {
    const c = 'SAMPLEテクノロジーズ研究所', r = presen(roomy, c);
    ck(F.textWidthUnits_(c) * 44 * 1.05 <= inner && !CB.fits(c, 44, inner), '1) 見本の会社名が、以前の見積もりで1行に入らない（見本の選び方の誤り）');
    ck(r.lines.length === 2 && r.lines.every((l) => CB.fits(l, r.pt, inner)), '1) 英字の多い会社名を、CB_ の見積もりで2行にしない: ' + J([r.lines, r.pt]));
  }
  // 同じ会社は、幅の近いメンバーのページ（528pt）と同じ切れ目
  ['株式会社見本不動産コンサルティングオフィス', '特定非営利活動法人見本子ども支援センター', '見本インターナショナルトレーディング研究所', 'Mihon Design Studio合同会社']
    .forEach((c) => {
      const r = presen(roomy, c), m = CB.layout(c, MEMBER).lines;
      ck(J(r.lines) === J(m), `1) 「${c}」がメンバーのページと違う切れ目: ${J(r.lines)} / ${J(m)}`);
    });
  // 枠の幅が違うひな形（520pt）でも、CB_.layout と同じ
  {
    const narrow = presenTpl(420, 520), w = F.readShapeTextBox_(narrow, 27).widthPt;
    ['見本インターナショナルトレーディング研究所', '株式会社SAMPLE Creative Works', '見本建築設計事務所一級建築士事務所'].forEach((c) => {
      const r = presen(narrow, c), want = CB.layout(c, Object.assign({ width: w }, PRESEN));
      ck(J(r.lines) === J(want.lines) && r.pt === want.pt && r.lines.every((l) => CB.fits(l, r.pt, w)),
         `1) 幅520pt の枠で「${c}」: ${J([r.lines, r.pt])}（${J([want.lines, want.pt])} のはず）`);
    });
  }

  // 余地が少ないとき（カウントダウンが【カテゴリー】の下43pt）：重ならない。どの行も枠に入る。不自然な所で折らない
  const tight = presenTpl(300);
  NAMES.forEach((c) => {
    const r = presen(tight, c);
    ck(r.bottom <= 300 + 1e-6, `1) 余地が少ないとき「${c}」：【カテゴリー】がカウントダウンに重なる: ${r.bottom}`);
    ck(r.lines.length <= 2 && r.lines.every((l) => CB.fits(l, r.pt, inner)) && r.pt >= 16, `1) 余地が少ないとき「${c}」の行が枠に入らない: ${J([r.lines, r.pt])}`);
    const bad = unnatural(r.lines, c);
    ck(!bad, `1) 余地が少ないとき「${c}」が不自然な所で折れた（${bad}）: ${J(r.lines)}`);
    ck(r.cat.every((v) => v === 32), `1) 余地が少ないとき「${c}」：【カテゴリー】が32ptでない: ${J(r.cat)}`);
  });
  // 1行で入る会社名は、余地のために小さな2行に分けない（1行のまま小さくする）
  {
    const r = presen(tight, '見本コンサルティングオフィス株式会社');
    ck(r.lines.length === 1 && r.pt >= 30, '1) 余地が少ないとき、1行で入る会社名を小さな2行にした: ' + J([r.lines, r.pt]));
    const s = presen(tight, '見本商事');
    ck(J(s.lines) === J(['見本商事']) && s.pt === 44, '1) 余地が少ないとき、短い会社名: ' + J([s.lines, s.pt]));
    const v = presen(presenTpl(262), '見本コンサルティング株式会社');   // 【カテゴリー】の下5pt
    ck(v.lines.length === 1 && v.pt >= 30 && v.bottom <= 262 + 1e-6, '1) 余地がほとんど無いとき: ' + J([v.lines, v.pt, v.bottom]));
  }
}

// ===== 2) お2人のページ =====
// 会社名の枠：幅 315pt（内側 300.6pt）・20pt。お名前 28pt・カテゴリー 18pt（2行まで）
const tsp = (id, text, sz, cx) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="t${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>`
  + `<a:xfrm><a:off x="${(id % 2) * 6000000}" y="${id * 300000}"/><a:ext cx="${cx || 4000000}" cy="400000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`
  + `<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r><a:rPr lang="ja-JP" sz="${sz}"/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
const twoPage = (prefix, title) => tsp(2, title, 3600, 9000000)
  + tsp(3, `{{${prefix}1氏名}}`, 2800) + tsp(5, `{{${prefix}1会社名}}`, 2000) + tsp(7, `{{${prefix}1カテゴリー}}`, 1800)
  + tsp(4, `{{${prefix}2氏名}}`, 2800) + tsp(6, `{{${prefix}2会社名}}`, 2000) + tsp(8, `{{${prefix}2カテゴリー}}`, 1800);
const DECK = [['メインプレゼン', '今週のメインプレゼンテーション'], ['推薦のことば', '推薦のことば'], ['抽選', '抽選コーナー']];
function makeParts() {
  const parts = {}, put = (p, s) => { parts[p] = new env.FakeBlob(s, 'application/xml', p); };
  put('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
    + DECK.map((d, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('') + '</Types>');
  put('ppt/presentation.xml', `<?xml version="1.0"?><p:presentation ${NS}><p:sldIdLst>`
    + DECK.map((d, i) => `<p:sldId id="${256 + i}" r:id="rId${100 + i}"/>`).join('') + '</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>');
  put('ppt/_rels/presentation.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + DECK.map((d, i) => `<Relationship Id="rId${100 + i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`).join('') + '</Relationships>');
  put('ppt/slideLayouts/slideLayout1.xml', `<?xml version="1.0"?><p:sldLayout ${NS}><p:cSld><p:spTree/></p:cSld></p:sldLayout>`);
  DECK.forEach(([prefix, title], i) => {
    put(`ppt/slides/slide${i + 1}.xml`, slideXml(twoPage(prefix, title)));
    put(`ppt/slides/_rels/slide${i + 1}.xml.rels`, '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>');
  });
  return parts;
}
const INNER = 4000000 / E - 14.4;   // 会社名の枠の内側の幅（300.6pt）
// 以前の縮め方（お名前・カテゴリーはこのまま）
const oldSize = (text, pt, lines, min) => {
  const w = INNER * 0.95, em = F.textBoxEm_(text);
  return em * pt <= w * lines ? pt : Math.max(min, Math.floor(w * lines / em * 2) / 2);
};
// 会社名の枠の中身と大きさ（差し込み口の番号 5・6 が1人目・2人目）
const companyOf = (x, k) => { const seg = shapeOf(x, k === 1 ? 5 : 6); return { text: textsOf(seg).join(''), paras: parasOf(seg).length, pt: sizesOf(seg)[0] }; };
const checkCompany = (label, got, raw) => {
  const want = CB.normalize(raw);
  ck(got.text === want && got.paras === 1 && !/[\n㈱]/.test(got.text), `2) ${label}：会社名がそろっていない・1行でない: ${J(got)}（${J(want)} のはず）`);
  ck(CB.lineWidthPt(got.text, got.pt) <= INNER + 1e-6, `2) ${label}：「${got.text}」${got.pt}pt が枠の内側（${INNER}pt）に入らない（見積もり ${CB.lineWidthPt(got.text, got.pt).toFixed(1)}pt）`);
  const full = CB.lineWidthPt(want, 20) <= INNER + 1e-6;
  ck(full ? got.pt === 20 : got.pt < 20 && CB.lineWidthPt(got.text, got.pt + 0.5) > INNER, `2) ${label}：「${got.text}」の大きさ ${got.pt}pt（入るなら20pt、入らなければ入るいちばん大きい0.5pt刻み）`);
};
{
  // 会社名（作り物）：英字の多いもの・名簿のセルの改行と ㈱・入るもの・長いもの
  const CO = {
    m1: 'SAMPLE CREATIVE WORKS株式会社', m2: '㈱見本\nエンジニアリング',
    l1: '見本不動産コンサルティングオフィス 東京支社', l2: '試験商事',
    g: 'Mihon\nDesign Studio合同会社', r: '㈲試験\n工務店',
  };
  // 英字の多い会社名：以前の見積もり（半角0.55字・幅の95%）では縮め足りず、CB_ の見積もりでは枠からはみ出す（見本の選び方の確かめ）
  const before = oldSize(CO.m1, 20, 1, 8);
  ck(CB.lineWidthPt(CO.m1, before) > INNER, '2) 見本の会社名が、以前の縮め方でも枠に入る（見本の選び方の誤り）: ' + before);

  // fitTwoPersonText_ 単体：改行の残った会社名でも、そろえた文字で測る
  {
    const xml = slideXml(twoPage('メインプレゼン', '今週のメインプレゼンテーション'));
    const raw = F.fitTwoPersonText_(xml, 'メインプレゼン', { メインプレゼン1会社名: 'SAMPLE\nCREATIVE WORKS株式会社', メインプレゼン2会社名: CO.m2 });
    const flat = F.fitTwoPersonText_(xml, 'メインプレゼン', { メインプレゼン1会社名: CB.normalize('SAMPLE\nCREATIVE WORKS株式会社'), メインプレゼン2会社名: CB.normalize(CO.m2) });
    ck(raw === flat, '2) fitTwoPersonText_：改行の残った会社名を、そろえた文字と違う大きさにした');
    const p = sizesOf(shapeOf(raw, 5))[0];
    ck(p < 20 && CB.lineWidthPt(CB.normalize(CO.m1), p) <= INNER + 1e-6, '2) fitTwoPersonText_：英字の多い会社名の大きさ: ' + p);
    ck(sizesOf(shapeOf(raw, 6))[0] === 20, '2) fitTwoPersonText_：入る会社名（㈱を開いたもの）を縮めた: ' + sizesOf(shapeOf(raw, 6))[0]);
    // fitTokenBox_ 単体：会社名の印（6つ目）があるときだけ CB_ で測る。無ければ以前のまま
    const t = F.fitTokenBox_(xml, '{{メインプレゼン1会社名}}', CO.m1, 1, 8);
    ck(sizesOf(shapeOf(t, 5))[0] === before, '2) fitTokenBox_：会社名の印が無いときの縮め方が以前と違う: ' + sizesOf(shapeOf(t, 5))[0]);
  }

  // 作成（editMeetingSlides_）：メインプレゼン・抽選は画面の値（map）、推薦のことばは組（recommendPairs）から
  const parts = makeParts();
  const map = {
    メインプレゼン1氏名: '見本 一郎', メインプレゼン1会社名: CO.m1, メインプレゼン1カテゴリー: '【とても長い見本のカテゴリーの名前・見本の専門分野の説明・試験の分野の説明】',
    メインプレゼン2氏名: 'SAMPLE TARO MIHON-YAMADA', メインプレゼン2会社名: CO.m2, メインプレゼン2カテゴリー: '【税理士】',
    抽選1氏名: '見本 一郎', 抽選1会社名: CO.l1, 抽選1カテゴリー: '【見本】', 抽選2氏名: '試験 花子', 抽選2会社名: CO.l2, 抽選2カテゴリー: '【試験】',
  };
  const pairs = [{ giver: { name: '見本 一郎', company: CO.g, category: '【見本】' }, receiver: { name: '試験 花子', company: CO.r, category: '【試験】' }, after: false }];
  const mapCopy = J(map);
  F.editMeetingSlides_(parts, map, [], { mainPresenters: ['見本 一郎', 'SAMPLE TARO MIHON-YAMADA'], lottery: ['見本 一郎', '試験 花子'], recommendPairs: pairs });
  ck(J(map) === mapCopy, '2) 渡した値（map）を書き換えた');
  const xmlOf = (i) => F.xmlOf_(parts, `ppt/slides/slide${i}.xml`);
  const main = xmlOf(1), reco = xmlOf(2), lot = xmlOf(3);
  ck(![main, reco, lot].some((x) => /\{\{/.test(F.slideText_(x))), '2) 差し込み口が残った');
  checkCompany('メインプレゼン1', companyOf(main, 1), CO.m1);
  checkCompany('メインプレゼン2', companyOf(main, 2), CO.m2);
  checkCompany('抽選1', companyOf(lot, 1), CO.l1);
  checkCompany('抽選2', companyOf(lot, 2), CO.l2);
  checkCompany('推薦のことば（推薦する方）', companyOf(reco, 1), CO.g);
  checkCompany('推薦のことば（推薦される方）', companyOf(reco, 2), CO.r);
  ck(companyOf(reco, 1).text === 'Mihon Design Studio合同会社' && companyOf(reco, 2).text === '有限会社試験工務店' && companyOf(main, 2).text === '株式会社見本エンジニアリング',
     '2) 会社名のそろえ方: ' + J([companyOf(reco, 1).text, companyOf(reco, 2).text, companyOf(main, 2).text]));
  // お名前・カテゴリーの縮め方は以前のまま
  const nameSz = sizesOf(shapeOf(main, 4))[0], catSz = sizesOf(shapeOf(main, 7))[0];
  ck(nameSz === oldSize('SAMPLE TARO MIHON-YAMADA', 28, 1, 20), '2) お名前の縮め方が変わった: ' + nameSz);
  ck(catSz === oldSize(map.メインプレゼン1カテゴリー, 18, 2, 9) && catSz < 18, '2) カテゴリーの縮め方が変わった: ' + catSz);
  ck(sizesOf(shapeOf(main, 3))[0] === 28 && sizesOf(shapeOf(main, 8))[0] === 18, '2) 短いお名前・カテゴリーを縮めた');
}

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.slice(0, 40).forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('サーバーのスライドの会社名: 検査 ' + checks + ' 件 OK: ビジタープレゼンは CB_.layout と同じ行・大きさ（枠の内側の幅・語の途中で折らない・'
  + 'メンバーのページと同じ切れ目・1行も CB_ の見積もりで入る・余地が少なくてもカウントダウンに重ならない・短い会社名を小さな2行にしない）・'
  + 'お2人のページ（メインプレゼン・抽選・推薦のことば）の会社名は名簿の書き方にそろえて1行・CB_ の見積もりで枠に入る大きさ・お名前とカテゴリーは以前のまま');

// 推薦のことば：受け取ったスライド（pptx・pdf・画像の1枚）を後半スライドに入れる（reco_slide_srv.js・slides_meeting_second.html・
// meeting_slides_srv.js の expandRecommendations_）を、実データなしで確かめる。名前はすべて架空。
//
//   node tools/check_reco_slides.js
//   PDFJS_DIR=/どこか/pdfjs3 node tools/check_reco_slides.js   … pdf.js（pdfjs-dist 3.11.174 の build/pdf.min.js・pdf.worker.min.js）を
//                                                                 置いたフォルダを渡すと、本物の pdf.js でPDFを画像にする
//                                                                 （渡さなければ、pdf.js の代わりの作り物で確かめる）
//
// 確かめること
//   1) サーバー：受け取ったスライド（画像）をドライブの「03_生成物／推薦のことばのスライド」に置き、開催日・組ごとに覚える。
//      同じ組に置き直すと前のものはゴミ箱へ。外す。画像でないもの・組が空のときは置かない。pptx はGoogleスライドに変換してPDFで返す。
//      覚えるのは開催日ごとのプロパティ（1つ 9KB まで。以前は全部を1つに入れていて、数か月ぶんたまると置けなくなった）。
//      以前の1つのものは開催日ごとへ移す（古い開催日のものは移さない）。覚えられなかったときは、置いたファイルをゴミ箱へ
//   2) 作成：その組の推薦のことばのページのすぐあとに、受け取ったスライドのページを入れる（定例会中・アフターとも）。
//      画像は縦横比のままページいっぱいに真ん中へ・まわりは黒・レイアウトの飾りは出さない。開けないファイルはお知らせして入れない。
//      お知らせに、推薦のことばのページと受け取ったスライドが何枚目かを出す
//   3) 画面（Chromium）：画像は長い辺2400pxまでに縮めて送る・PDFは1ページ目を画像にして送る・pptxはPDFにしてもらってから同じく・
//      使えないファイル・組が空のときは送らない（理由はその組のすぐ下にも出す）・置いてあるものは開き直しても同じ組に付く・外す・
//      アフターの組は「抽選コーナーのあとの、この組のページのすぐあと」と出す・準備しているあいだは作成させない・作成に組ごとのファイルが渡る

process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const zlib = require('zlib');
const { makeEnv } = require('./lib_sheet_fake');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }
const J = (x) => JSON.stringify(x);

// --- 作り物の画像（単色のPNG）・JPEGの頭（大きさだけ読めるもの）---
function png(w, h, rgb) {
  const chunk = (type, data) => {
    const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
    const td = Buffer.concat([Buffer.from(type), data]), crc = Buffer.alloc(4);
    crc.writeUInt32BE(zlib.crc32(td) >>> 0);
    return Buffer.concat([len, td, crc]);
  };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2;
  const row = Buffer.alloc(1 + w * 3); for (let x = 0; x < w; x++) row.set(rgb, 1 + x * 3);
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr),
                        chunk('IDAT', zlib.deflateSync(Buffer.concat(Array.from({ length: h }, () => row)))), chunk('IEND', Buffer.alloc(0))]);
}
const jpegHead = (w, h) => Buffer.concat([Buffer.from([0xff, 0xd8, 0xff, 0xc0, 0x00, 0x11, 0x08, h >> 8, h & 255, w >> 8, w & 255, 0x03, 1, 0x22, 0, 2, 0x11, 1, 3, 0x11, 1, 0xff, 0xd9]), Buffer.alloc(16)]);
const pngSize = (buf) => ({ w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) });
const dataUrlOf = (buf, type) => 'data:' + type + ';base64,' + buf.toString('base64');

// ===== サーバー（本番の *.js を見せかけのスプレッドシート・ドライブの上で）=====
const env = makeEnv({ now: new Date(2026, 9, 5, 10, 0, 0) });
const F = Object.assign({}, env.globals);
vm.createContext(F);
for (const f of fs.readdirSync(ROOT).filter((x) => /\.js$/.test(x)).sort()) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), F, { filename: f });
}
// 以前の覚え方（全部の開催日を1つのプロパティに）。古い開催日と、前の回の開催日のもの
env.reset([['メンバー名簿', false, [['No', '業種区分', '氏名'], ['1', '', '見本 一郎'], ['2', '', '試験 花子']]]],
          { BNI_RECO_SLIDES: J({ '2026/01/07': [{ g: '昔 一郎', r: '昔 二郎', id: 'OLD', name: '古い.png', w: 10, h: 10 }],
                                 '2026/09/30': [{ g: '前回 一郎', r: '前回 二郎', id: 'PREV', name: '前の回.png', w: 10, h: 10 }] }) });
// スクリプトのプロパティは、1つの値が 9KB まで（本物と同じく、超えたら止まる）
{
  const real = F.PropertiesService.getScriptProperties;
  F.PropertiesService = Object.assign({}, F.PropertiesService, { getScriptProperties: () => {
    const st = real();
    return Object.assign({}, st, { setProperty: (k, v) => {
      if (Buffer.byteLength(String(v), 'utf8') > 9 * 1024) throw new Error('Argument too large: value');
      return st.setProperty(k, v);
    } });
  } });
}
F.findPhotoIdForName_ = () => '';
const DATE = '2026/10/07';
const folderOf = (id) => (env.drive.files[id] || {}).parent || '';
const saved = (d) => F.getRecommendationSlides(d || DATE).slides || [];

// ----- 1) 置く・覚える・置き直す・外す -----
{
  ck(J(saved('2026/09/30').map((x) => x.id)) === J(['PREV']), '1) 以前の覚え方（1つのプロパティ）のものを読めない: ' + J(saved('2026/09/30')));
  const r1 = F.saveRecommendationSlide({ date: DATE, giver: '見本 一郎', receiver: '試験 花子', name: '推薦のスライド.pdf', data: dataUrlOf(png(1600, 900, [10, 20, 30]), 'image/png') });
  ck(r1.ok && r1.slide && r1.slide.w === 1600 && r1.slide.h === 900 && r1.slide.name === '推薦のスライド.pdf',
     '1) 受け取ったスライドを置けない: ' + J(r1));
  ck(r1.ok && /03_生成物\/推薦のことばのスライド$/.test(folderOf(r1.slide.id)) && /^20261007_推薦のことば_見本一郎→試験花子\.png$/.test(env.drive.files[r1.slide.id].name),
     '1) 置き場所・ファイル名: ' + J({ folder: folderOf(r1.slide && r1.slide.id), name: r1.slide && (env.drive.files[r1.slide.id] || {}).name }));
  ck(J(saved().map((s) => [s.giver, s.receiver, s.id])) === J([['見本 一郎', '試験 花子', r1.slide.id]]), '1) 開催日・組ごとに覚えていない: ' + J(saved()));
  // 開催日ごとに覚える。以前の1つのものは開催日ごとへ移して消し、古い開催日（120日より前）のものは移さない
  ck(env.props.BNI_RECO_SLIDES === undefined && J(JSON.parse(env.props.BNI_RECO_SLIDES_20261007 || '[]').map((x) => x.id)) === J([r1.slide.id])
     && J(JSON.parse(env.props.BNI_RECO_SLIDES_20260930 || '[]').map((x) => x.id)) === J(['PREV']) && env.props.BNI_RECO_SLIDES_20260107 === undefined
     && J(saved('2026/09/30').map((x) => x.id)) === J(['PREV']),
     '1) 開催日ごとに覚えていない・以前のものを移さない・古い開催日のものを消さない: ' + J(Object.keys(env.props).filter((k) => /RECO/.test(k))));
  // 同じ組に置き直す → 前のものはゴミ箱へ。別の組は別に覚える
  const r2 = F.saveRecommendationSlide({ date: DATE, giver: '見本 一郎', receiver: '試験 花子', name: '差し替え.jpg', data: dataUrlOf(jpegHead(800, 600), 'image/jpeg') });
  ck(r2.ok && r2.slide.w === 800 && r2.slide.h === 600 && env.drive.files[r1.slide.id].trashed && /\.jpg$/.test(env.drive.files[r2.slide.id].name),
     '1) 置き直したとき、前のものをゴミ箱に入れない・JPEGを読めない: ' + J({ r2, oldTrashed: env.drive.files[r1.slide.id].trashed }));
  const r3 = F.saveRecommendationSlide({ date: DATE, giver: '試験 花子', receiver: '見本 一郎', name: '別の組.png', data: dataUrlOf(png(400, 300, [1, 2, 3]), 'image/png') });
  ck(r3.ok && saved().length === 2, '1) 別の組を覚えない: ' + J(saved()));
  // 画像でないもの・組が空
  const bad = F.saveRecommendationSlide({ date: DATE, giver: '見本 一郎', receiver: '試験 花子', name: 'x.txt', data: dataUrlOf(Buffer.from('hello world, not an image'), 'text/plain') });
  ck(!bad.ok && /画像として読めません/.test(bad.message) && saved().length === 2, '1) 画像でないものを置いた: ' + J(bad));
  const noPair = F.saveRecommendationSlide({ date: DATE, giver: '', receiver: '', name: 'a.png', data: dataUrlOf(png(10, 10, [0, 0, 0]), 'image/png') });
  ck(!noPair.ok && /先に、推薦する方/.test(noPair.message), '1) 組が空なのに置いた: ' + J(noPair));
  // 数か月ぶん（毎週3組・長いファイル名）たまっても置ける（以前は全部を1つのプロパティに入れていたため、9KB を超えて置けなくなった）
  {
    const longName = '【推薦のことば】' + '見本'.repeat(20) + '様への推薦スライド_最終版.pptx';
    let bad = null;
    for (let w = 0; w < 16 && !bad; w++) {
      const day = new Date(2026, 5, 17 + w * 7), ds = day.getFullYear() + '/' + String(day.getMonth() + 1).padStart(2, '0') + '/' + String(day.getDate()).padStart(2, '0');
      for (let k = 0; k < 3 && !bad; k++) {
        const q = F.saveRecommendationSlide({ date: ds, giver: '長い名前の方 ' + k, receiver: '推薦される方 ' + k, name: longName, data: dataUrlOf(png(16, 9, [k, w, 0]), 'image/png') });
        if (!q.ok) bad = ds + ' ' + q.message;
      }
    }
    const kept = Object.keys(env.props).filter((k) => /^BNI_RECO_SLIDES_/.test(k));
    ck(!bad && kept.length >= 16 && saved('2026/06/17').length === 3 && saved('2026/09/30').length === 4,
       '1) 数か月ぶんたまると置けない（プロパティの 9KB）: ' + J({ bad, kept: kept.length, june: saved('2026/06/17').length, sep30: saved('2026/09/30').length }));
    // 置けなかったときは、ドライブに置いたファイルを残さない
    const before = Object.keys(env.drive.files).length, realSet = F.PropertiesService;
    F.PropertiesService = { getScriptProperties: () => Object.assign({}, realSet.getScriptProperties(), { setProperty: () => { throw new Error('Argument too large: value'); } }) };
    const ng = F.saveRecommendationSlide({ date: DATE, giver: '見本 一郎', receiver: '架空 三郎', name: 'x.png', data: dataUrlOf(png(16, 9, [1, 1, 1]), 'image/png') });
    F.PropertiesService = realSet;
    const made = Object.keys(env.drive.files).slice(before).map((id) => env.drive.files[id]);
    ck(!ng.ok && made.length === 1 && made[0].trashed, '1) 覚えられなかったのに、ドライブにファイルを残した: ' + J({ ng, made: made.map((f) => [f.name, f.trashed]) }));
  }
  // 外す
  const rm = F.removeRecommendationSlide(DATE, r3.slide.id);
  ck(rm.ok && saved().length === 1 && env.drive.files[r3.slide.id].trashed, '1) 外せない: ' + J({ rm, saved: saved() }));
  // pptx → PDF（Googleスライドに変換して書き出す。変換に使ったファイルはゴミ箱へ）
  let created = null;
  F.Drive = { Files: { create: (res, blob, opt) => { created = { res, type: blob.getContentType(), opt }; env.addFile('TMP', blob, res.name); return { id: 'TMP' }; } } };
  const realGet = F.DriveApp.getFileById;
  F.DriveApp = Object.assign({}, F.DriveApp, { getFileById: (id) => Object.assign(realGet(id), { getAs: (t) => new env.FakeBlob(Buffer.from('%PDF-1.4 見本'), t, 'x.pdf') }) });
  const cv = F.convertRecommendationPptx({ name: '推薦.pptx', data: Buffer.from('PK\u0003\u0004 見本のpptx').toString('base64') });
  ck(cv.ok && Buffer.from(cv.pdf, 'base64').toString('utf8') === '%PDF-1.4 見本' && created && created.res.mimeType === 'application/vnd.google-apps.presentation'
     && /presentationml/.test(created.type) && env.drive.files.TMP.trashed, '1) pptx をPDFにできない・変換に使ったファイルを残した: ' + J({ cv: cv.ok ? 'ok' : cv, created }));
  const cvBad = F.convertRecommendationPptx({ name: 'x.pptx', data: Buffer.from('not a zip').toString('base64') });
  ck(!cvBad.ok && /pptx として読めません/.test(cvBad.message), '1) pptx でないものを変換した: ' + J(cvBad));
  F.DriveApp = Object.assign({}, F.DriveApp, { getFileById: realGet });
}

// ----- 2) 作成：その組のページのすぐあとに入れる -----
const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
  + 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const sp = (id, text) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="t${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${id * 100000}" y="${id * 300000}"/>`
  + `<a:ext cx="3000000" cy="400000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ja-JP" sz="2000"/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
const SLIDES = [
  ['表紙', sp(2, '見本チャプター 定例会')],
  ['推薦', sp(2, '推薦のことば') + sp(3, '{{推薦のことば1氏名}}') + sp(4, '{{推薦のことば1会社名}}') + sp(5, '{{推薦のことば2氏名}}') + sp(6, '{{推薦のことば2会社名}}')],
  ['ほか', sp(2, 'リファーラルの時間')],
  ['抽選', sp(2, '抽選コーナー') + sp(3, '{{抽選1氏名}}') + sp(4, '{{抽選2氏名}}')],
  ['締め', sp(2, '締めの言葉')],
];
function makeParts() {
  const parts = {};
  const put = (p, s) => { parts[p] = new env.FakeBlob(s, 'application/xml', p); };
  put('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
    + SLIDES.map((s, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('') + '</Types>');
  put('ppt/presentation.xml', `<?xml version="1.0"?><p:presentation ${NS}><p:sldIdLst>`
    + SLIDES.map((s, i) => `<p:sldId id="${256 + i}" r:id="rId${100 + i}"/>`).join('') + '</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>');
  put('ppt/_rels/presentation.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + SLIDES.map((s, i) => `<Relationship Id="rId${100 + i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`).join('') + '</Relationships>');
  put('ppt/slideLayouts/slideLayout1.xml', `<?xml version="1.0"?><p:sldLayout ${NS}><p:cSld><p:spTree/></p:cSld></p:sldLayout>`);
  put('ppt/slideLayouts/slideLayout7.xml', `<?xml version="1.0"?><p:sldLayout ${NS}><p:cSld><p:spTree/></p:cSld></p:sldLayout>`);
  SLIDES.forEach(([k, inner], i) => {
    put(`ppt/slides/slide${i + 1}.xml`, `<?xml version="1.0"?><p:sld ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>`
      + inner + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>');
    put(`ppt/slides/_rels/slide${i + 1}.xml.rels`, '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout${k === '推薦' ? 7 : 1}.xml"/></Relationships>`);
  });
  return parts;
}
{
  // ドライブに置いた受け取ったスライド：16:9 の PNG と 4:3 の JPEG
  env.addFile('IMG169', new env.FakeBlob(png(1920, 1080, [200, 30, 30]), 'image/png', 'a.png'), 'a.png');
  env.addFile('IMG43', new env.FakeBlob(jpegHead(800, 600), 'image/jpeg', 'b.jpg'), 'b.jpg');
  env.addFile('IMGAFTER', new env.FakeBlob(png(1000, 1000, [30, 30, 200]), 'image/png', 'c.png'), 'c.png');
  const person = (n) => ({ name: n, company: n + '商事', category: '【見本】' });
  const pairs = [
    { giver: person('見本 一郎'), receiver: person('試験 花子'), after: false, slide: { id: 'IMG169', name: '一組目.pdf' } },
    { giver: person('試験 花子'), receiver: person('見本 一郎'), after: false, slide: null },
    { giver: person('架空 三郎'), receiver: person('仮名 四郎'), after: false, slide: { id: 'IMG43', name: '三組目.jpg' } },
    { giver: person('例示 五月'), receiver: person('模擬 六助'), after: true, slide: { id: 'IMGAFTER', name: 'アフター.png' } },
    { giver: person('空想 七海'), receiver: person('見本 一郎'), after: true, slide: { id: 'GONE', name: '消えたファイル.png' } },
  ];
  const gone = F.recoLoadSlides_(pairs);
  ck(/消えたファイル\.png/.test(gone) && !pairs[4].slideImage && pairs[0].slideImage && pairs[0].slideImage.width === 1920 && pairs[2].slideImage.ext === 'jpeg',
     '2) 受け取ったスライドを開く・開けないものを知らせる: ' + J({ gone, ok0: !!pairs[0].slideImage }));
  const parts = makeParts();
  const res = F.expandRecommendations_(parts, pairs, { by: {}, seq: 0 });
  const order = F.slideOrder_(parts);
  const kindOf = (p) => {
    const x = F.xmlOf_(parts, p), t = F.slideText_(x);
    if (/<p:pic>/.test(x) && /受け取ったスライド/.test(x)) return 'スライド:' + (F.xmlOf_(parts, F.relsPathOf_(p)).match(/media\/(recoslide\d+\.\w+)/) || [])[1];
    if (/推薦のことば/.test(t)) return '推薦:' + (t.match(/(?:見本|試験|架空|仮名|例示|模擬|空想) \S{2}(?=商事)/g) || []).join('→');
    return t.replace(/\{\{.*$/, '').slice(0, 6);
  };
  const got = order.map(kindOf);
  ck(J(got) === J(['見本チャプタ', '推薦:見本 一郎→試験 花子', 'スライド:recoslide1.png', '推薦:試験 花子→見本 一郎', '推薦:架空 三郎→仮名 四郎', 'スライド:recoslide2.jpeg',
                   'リファーラル', '抽選コーナー', '推薦:例示 五月→模擬 六助', 'スライド:recoslide3.png', '推薦:空想 七海→見本 一郎', '締めの言葉']),
     '2) 受け取ったスライドのページの場所（その組の推薦のことばのページのすぐあと）: ' + J(got));
  ck(res.slides && res.slides.length === 3 && /受け取ったスライド 3枚/.test(res.message), '2) お知らせ: ' + J(res.message));
  // お知らせに、何枚目に入れたか（PowerPoint の左の一覧の番号）
  ck(/定例会中。2〜6枚目/.test(res.message) && /抽選コーナー（8枚目）のあと（9〜11枚目）/.test(res.message)
     && /見本 一郎 → 試験 花子 … 3枚目、架空 三郎 → 仮名 四郎 … 6枚目、例示 五月 → 模擬 六助 … 10枚目/.test(res.message),
     '2) お知らせに、推薦のことばのページ・受け取ったスライドが何枚目かが無い: ' + J(res.message));
  // ページの中身：縦横比のままページいっぱいに真ん中へ・まわりは黒・レイアウトの飾りは出さない・レイアウトは推薦のことばのページと同じ
  const pics = order.filter((p) => /受け取ったスライド/.test(F.xmlOf_(parts, p)));
  const geo = (p) => { const x = F.xmlOf_(parts, p); const o = x.match(/<a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/><\/a:xfrm><a:prstGeom/); return o ? o.slice(1).map(Number) : null; };
  ck(J(geo(pics[0])) === J([0, 0, 12192000, 6858000]), '2) 16:9 の画像がページいっぱいにならない: ' + J(geo(pics[0])));
  ck(J(geo(pics[1])) === J([1524000, 0, 9144000, 6858000]), '2) 4:3 の画像が真ん中に縦いっぱいにならない: ' + J(geo(pics[1])));
  ck(J(geo(pics[2])) === J([2667000, 0, 6858000, 6858000]), '2) 正方形の画像: ' + J(geo(pics[2])));
  const x0 = F.xmlOf_(parts, pics[0]), r0 = F.xmlOf_(parts, F.relsPathOf_(pics[0]));
  ck(/showMasterSp="0"/.test(x0) && /<p:bg><p:bgPr><a:solidFill><a:srgbClr val="000000"\/>/.test(x0) && /descr="一組目\.pdf"/.test(x0),
     '2) まわりの黒・レイアウトの飾りを出さない・元のファイル名: ' + x0.slice(0, 400));
  ck(/Target="\.\.\/slideLayouts\/slideLayout7\.xml"/.test(r0) && /Target="\.\.\/media\/recoslide1\.png"/.test(r0), '2) 関係（レイアウト・画像）: ' + r0);
  ck(parts['ppt/media/recoslide1.png'] && parts['ppt/media/recoslide2.jpeg'] && parts['ppt/media/recoslide3.png'], '2) 画像が入っていない');
  const ct = F.xmlOf_(parts, '[Content_Types].xml');
  ck(/Extension="png"/.test(ct) && /Extension="jpeg"/.test(ct) && pics.every((p) => ct.includes('PartName="/' + p + '"')), '2) 種類の登録: ' + ct.slice(0, 300));
  ck(!/<p:sld\b[^>]*\sshow="0"/.test(x0), '2) 受け取ったスライドのページが非表示');
}

// ----- 3) 画面（Chromium）-----
const pw = (() => { try { return require('playwright'); } catch (e) { return require('/opt/node22/lib/node_modules/playwright'); } })();
function evalTemplate(file, vars) {
  const src = fs.readFileSync(path.join(ROOT, file), 'utf8');
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  let code = 'var __o=[];with(__v){', i = 0, m;
  const re = /<\?(!=|=)?([\s\S]*?)\?>/g;
  while ((m = re.exec(src))) {
    code += '__o.push(' + J(src.slice(i, m.index)) + ');';
    if (m[1] === '=') code += '__o.push(__e(' + m[2] + '));';
    else if (m[1] === '!=') code += '__o.push(String(' + m[2] + '));';
    else code += m[2] + '\n';
    i = re.lastIndex;
  }
  code += '__o.push(' + J(src.slice(i)) + ');}return __o.join("");';
  const include = (n) => fs.readFileSync(path.join(ROOT, n + '.html'), 'utf8');
  return new Function('__v', '__e', code)(Object.assign({ include }, vars), esc);
}
const MEMBERS = ['見本 一郎', '試験 花子', '架空 三郎', '仮名 四郎'].map((n, i) => ({ no: String(i + 1), name: n, company: n + '商事', title: 'カテゴリー', hasPhoto: true }));
const who = (n) => ({ raw: n.split(' ')[0] + 'さん', name: n, matched: true });
const CTX = {
  ok: true, meetings: [{ dateValue: DATE, display: '第536回 2026/10/07' }], defaultMeeting: { dateValue: DATE, display: '第536回 2026/10/07' },
  lists: { ok: true, text: { d90: '該当者なし', d60: '該当者なし', d30: '該当者なし', overdue: '該当者なし' }, d90: [], d60: [], d30: [], overdue: [], noDate: [], done: [], leaving: [] },
  templates: { meetingSecond: true }, seconds: { referral: 7 }, memberCount: MEMBERS.length, members: MEMBERS,
  routine: { ok: true, found: true, sheetName: '【24期】ルーティンチェックシート', meetingNo: '536', mainPresenters: [],
             recommendations: [{ giver: who('見本 一郎'), receiver: who('試験 花子'), raw: 'A→B', when: 'during' },
                               { giver: who('架空 三郎'), receiver: who('仮名 四郎'), raw: 'C→D', when: 'during' }], recommendationsRaw: 'A→B C→D' },
};
// google.script.run の代わり（呼ばれた関数と引数を window.__calls に残す）。受け取ったスライドを置くと、作り物のIDを返す
const STUB = '<script>window.__calls=[];var __ctx=' + J(CTX).replace(/</g, '\\u003c') + ';var __n=0;var __pdf=' + J(tinyPdf().toString('base64')) + ';'
  + 'var __S={getSystemVersion:function(){return "";},getMeetingSlideContext:function(){return __ctx;},getMeetingMusicFiles:function(){return {ok:true,files:[]};},'
  + 'getMeetingTemplateInfo:function(){return {ok:true,list:[],referralBoxes:null};},computeRenewalLists:function(){return __ctx.lists;},'
  + 'getRecommendationSlides:function(d){return {ok:true,slides:[{giver:"架空 三郎",receiver:"仮名 四郎",id:"SAVED1",name:"前に置いた.png",w:800,h:450}]};},'
  + 'saveRecommendationSlide:function(q){__n++;return {ok:true,slide:{giver:q.giver,receiver:q.receiver,id:"UP"+__n,name:q.name,w:0,h:0},message:"「"+q.name+"」を入れます。"};},'
  + 'removeRecommendationSlide:function(){return {ok:true,message:"受け取ったスライドを外しました。"};},'
  + 'convertRecommendationPptx:function(q){return {ok:true,name:q.name,pdf:__pdf};},'
  + 'generateMeetingSlides:function(){return {ok:true,message:"作成しました",url:"#"};}};'
  + 'var google={script:{host:{close:function(){}},get run(){var ok=null,ng=null,p=new Proxy({},{get:function(_,n){'
  + 'if(n==="withSuccessHandler")return function(f){ok=f;return p;};if(n==="withFailureHandler")return function(f){ng=f;return p;};'
  + 'if(n==="withUserObject")return function(){return p;};'
  + 'return function(){var a=[].slice.call(arguments);window.__calls.push([n,a]);var v=__S[n]?__S[n].apply(null,a):null;'
  + 'setTimeout(function(){ok&&ok(v);},10);};}});return p;}}};</script>';
// pdf.js の代わり：960×540 のページを青で塗る（2ページあるPDF）。PDFJS_DIR を渡したときは本物を使う
const PDFJS_DIR = process.env.PDFJS_DIR || '';
const FAKE_PDFJS = '<script>window.pdfjsWorker={};window.pdfjsLib={GlobalWorkerOptions:{},getDocument:function(o){window.__pdfBytes=o.data.length;'
  + 'return {promise:Promise.resolve({numPages:2,getPage:function(n){return Promise.resolve({getViewport:function(v){return {width:960*v.scale,height:540*v.scale};},'
  + 'render:function(r){r.canvasContext.fillStyle="#3366cc";r.canvasContext.fillRect(0,0,r.viewport.width,r.viewport.height);return {promise:Promise.resolve()};}});}})};}};</script>';
// 本物の pdf.js で読む作り物のPDF（1ページ・960×540pt を青で塗る）
function tinyPdf() {
  const objs = ['<< /Type /Catalog /Pages 2 0 R >>', '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 960 540] /Contents 4 0 R >>'];
  const content = '0.2 0.4 0.8 rg 0 0 960 540 re f';
  objs.push('<< /Length ' + content.length + ' >>\nstream\n' + content + '\nendstream');
  let out = '%PDF-1.4\n', offs = [];
  objs.forEach((o, i) => { offs.push(out.length); out += (i + 1) + ' 0 obj\n' + o + '\nendobj\n'; });
  const xref = out.length;
  out += 'xref\n0 ' + (objs.length + 1) + '\n0000000000 65535 f \n' + offs.map((o) => String(o).padStart(10, '0') + ' 00000 n \n').join('');
  out += 'trailer\n<< /Size ' + (objs.length + 1) + ' /Root 1 0 R >>\nstartxref\n' + xref + '\n%%EOF\n';
  return Buffer.from(out, 'latin1');
}

(async () => {
  const html = evalTemplate('slides_meeting_second.html', {}).replace(/<head>/i, '<head><meta charset="utf-8">' + STUB + (PDFJS_DIR ? '' : FAKE_PDFJS));
  const browser = await pw.chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1000, height: 900 } });
  page.on('pageerror', (e) => fails.push('画面のエラー: ' + e.message));
  const URL = 'https://slides-second.test/';
  await page.route(URL, (rt) => rt.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: html }));
  await page.route(/cdn\.jsdelivr\.net\/npm\/pdfjs-dist@3\.11\.174\/build\/(pdf(?:\.worker)?\.min\.js)$/, (rt) => {
    const f = PDFJS_DIR ? path.join(PDFJS_DIR, rt.request().url().replace(/^.*\//, '')) : '';
    if (f && fs.existsSync(f)) rt.fulfill({ status: 200, contentType: 'application/javascript', body: fs.readFileSync(f) });
    else rt.abort();
  });
  await page.goto(URL);
  const calls = (n) => page.evaluate((nm) => window.__calls.filter((c) => c[0] === nm).map((c) => c[1]), n);
  const text = (id) => page.evaluate((i) => (document.getElementById(i) || {}).innerText || '', id);
  const waitCall = async (n, count) => { try { await page.waitForFunction(([nm, k]) => window.__calls.filter((c) => c[0] === nm).length >= k, [n, count], { timeout: 15000 }); } catch (e) {} };
  const sizeOfUpload = (q) => { const b = Buffer.from(String(q.data).replace(/^data:[^,]*,/, ''), 'base64'); return b[0] === 0x89 ? pngSize(b) : { w: -1, h: -1, jpeg: b[0] === 0xff }; };
  await page.waitForTimeout(600);

  // 置いてあるものは、開き直しても同じ組（架空さん→仮名さん）に付く
  ck(/前に置いた\.png/.test(await text('recoDuring')) && (await calls('getRecommendationSlides')).length === 1, '3) 置いてある受け取ったスライドが同じ組に付かない: ' + await text('recoDuring'));
  // 1組目に画像（3000×1500 → 長い辺2400pxに縮める）
  await page.setInputFiles('#rf_during_0', { name: '推薦の画像.png', mimeType: 'image/png', buffer: png(3000, 1500, [200, 100, 50]) });
  await waitCall('saveRecommendationSlide', 1);
  let up = (await calls('saveRecommendationSlide'))[0] || [{}];
  ck(up[0].date === DATE && up[0].giver === '見本 一郎' && up[0].receiver === '試験 花子' && up[0].name === '推薦の画像.png'
     && J(sizeOfUpload(up[0])) === J({ w: 2400, h: 1200 }), '3) 画像を縮めて送らない: ' + J({ date: up[0].date, g: up[0].giver, r: up[0].receiver, size: up[0].data ? sizeOfUpload(up[0]) : null }));
  await page.waitForTimeout(100);
  ck(/推薦の画像\.png/.test(await text('recoDuring')) && /この組のページのすぐあとに入れます/.test(await text('recoDuring')), '3) 置いたスライドが組に出ない: ' + await text('recoDuring'));
  // 1組目をPDFに置き直す（1ページ目を、長い辺2400pxの画像にして送る）
  await page.setInputFiles('#rf_during_0', { name: '推薦のスライド.pdf', mimeType: 'application/pdf', buffer: tinyPdf() });
  await waitCall('saveRecommendationSlide', 2);
  await page.waitForTimeout(150);
  up = (await calls('saveRecommendationSlide'))[1] || [{}];
  ck(up[0].name === '推薦のスライド.pdf' && J(sizeOfUpload(up[0])) === J({ w: 2400, h: 1350 }), '3) PDFの1ページ目を画像にして送らない: ' + J(up[0].data ? sizeOfUpload(up[0]) : up[0]));
  if (!PDFJS_DIR) ck(/2ページあるうちの、1ページ目/.test(await text('msg')), '3) 2ページ以上のPDFのお知らせが無い: ' + await text('msg'));
  // 色（本物の pdf.js でも作り物でも、ページは青）
  const center = await page.evaluate((d) => new Promise((res) => { const im = new Image(); im.onload = () => { const c = document.createElement('canvas'); c.width = im.width; c.height = im.height;
    const g = c.getContext('2d'); g.drawImage(im, 0, 0); res(Array.from(g.getImageData(im.width / 2, im.height / 2, 1, 1).data).slice(0, 3)); }; im.src = d; }), up[0].data);
  ck(Math.abs(center[0] - 51) < 8 && Math.abs(center[1] - 102) < 8 && Math.abs(center[2] - 204) < 8, '3) PDFの絵が画像にならない（真ん中の色）: ' + J(center));
  // 2組目は pptx（サーバーでPDFにしてもらってから、同じく画像にする）
  await page.setInputFiles('#rf_during_1', { name: '推薦.pptx', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', buffer: Buffer.from('PK\u0003\u0004見本') });
  await waitCall('saveRecommendationSlide', 3);
  await page.waitForTimeout(150);
  const conv = (await calls('convertRecommendationPptx'))[0] || [{}];
  up = (await calls('saveRecommendationSlide'))[2] || [{}];
  ck(conv[0].name === '推薦.pptx' && Buffer.from(conv[0].data || '', 'base64').slice(0, 2).toString('latin1') === 'PK', '3) pptx をサーバーに送らない: ' + J(conv[0] && conv[0].name));
  if (!PDFJS_DIR) ck(up[0].giver === '架空 三郎' && up[0].name === '推薦.pptx' && J(sizeOfUpload(up[0])) === J({ w: 2400, h: 1350 }), '3) pptx のPDFを画像にして送らない: ' + J(up[0].name));
  // 使えないファイル
  await page.setInputFiles('#rf_during_0', { name: '原稿.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from('PK') });
  await page.waitForTimeout(150);
  ck(/使えません/.test(await text('msg')) && (await calls('saveRecommendationSlide')).length === 3, '3) 使えないファイルを送った: ' + await text('msg'));
  // 置けなかった理由は、画面のいちばん下だけでなく、その組のすぐ下にも出す（下のお知らせは見落としやすい）
  ck(/⚠ 「原稿\.docx」は使えません/.test(await text('recoDuring')), '3) 使えないファイルのお知らせが、その組の下に出ない: ' + await text('recoDuring'));
  // 組が空のときは選ばせない
  await page.evaluate(() => { addPair('after'); });
  await page.evaluate(() => { recoChoose('after', 0); });
  ck(/先に、推薦する方・推薦される方を選んで/.test(await text('msg')) && /⚠ 先に、推薦する方・推薦される方を選んで/.test(await text('recoAfter')),
     '3) 組が空なのに選ばせた・お知らせが組の下に出ない: ' + await text('msg') + ' / ' + await text('recoAfter'));
  // アフターの組：入る場所は「抽選コーナーのあとの、この組のページのすぐあと」
  await page.evaluate(() => { pairs.after[0].g = '仮名 四郎'; pairs.after[0].r = '見本 一郎'; renderPairs(); });
  await page.setInputFiles('#rf_after_0', { name: 'アフターの推薦.png', mimeType: 'image/png', buffer: png(800, 450, [10, 120, 200]) });
  await waitCall('saveRecommendationSlide', 4);
  await page.waitForTimeout(150);
  ck(/アフターの推薦\.png … 抽選コーナーのあとの、この組のページのすぐあとに入れます/.test(await text('recoAfter')) && !/⚠/.test(await text('recoAfter')),
     '3) アフターの組の入る場所の書き方・前のお知らせが残る: ' + await text('recoAfter'));
  // 受け取ったスライドを準備しているあいだは作成させない（以前は押せて、スライドの入らないものができた）
  await page.evaluate(() => { pairs.during[0].busy = '「x.pptx」をPDFにしています…'; renderPairs(); });
  const dis = await page.evaluate(() => document.getElementById('genBtn').disabled);
  await page.evaluate(() => gen());
  await page.waitForTimeout(50);
  ck(dis && (await calls('generateMeetingSlides')).length === 0 && /受け取ったスライドを準備しています/.test(await text('msg')),
     '3) 受け取ったスライドを準備しているあいだに作成できた: ' + J({ dis, msg: await text('msg') }));
  await page.evaluate(() => { pairs.during[0].busy = ''; renderPairs(); });
  ck(!(await page.evaluate(() => document.getElementById('genBtn').disabled)), '3) 準備が終わっても作成のボタンを押せない');
  // 作成に、組ごとのファイルが渡る（1組目は最後に置いたPDF・2組目は pptx）
  await page.waitForTimeout(100);
  await page.evaluate(() => gen());
  await waitCall('generateMeetingSlides', 1);
  const gp = ((await calls('generateMeetingSlides'))[0] || [])[3] || {};
  ck(J((gp.recommendPairs || []).map((p) => [p.after, p.slide && p.slide.id])) === J([[false, 'UP2'], [false, 'UP3'], [true, 'UP4']]),
     '3) 作成に受け取ったスライドが渡らない: ' + J(gp.recommendPairs));
  // 外す
  await page.evaluate(() => recoUnset('during', 1));
  await waitCall('removeRecommendationSlide', 1);
  const rmc = (await calls('removeRecommendationSlide'))[0] || [];
  ck(rmc[0] === DATE && rmc[1] === 'UP3' && !/推薦\.pptx/.test(await text('recoDuring')), '3) 外せない: ' + J(rmc));
  await browser.close();

  if (fails.length) {
    console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
    fails.forEach((f) => console.log('  - ' + f));
    process.exit(1);
  }
  console.log('推薦のことばの受け取ったスライド: 検査 ' + checks + ' 件 OK' + (PDFJS_DIR ? '（本物の pdf.js）' : '') + ': ドライブに置く・組ごとに覚える・置き直す・外す・'
    + 'pptx→PDF・その組のページのすぐあとに入れる（縦横比のまま・まわりは黒）・画面（画像を縮める・PDFの1ページ目・pptx・使えないファイル・作成に渡る）');
})().catch((e) => { console.log('NG ' + (e && e.stack ? e.stack : e)); process.exit(1); });

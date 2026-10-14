// 「ビジター・代理スライド作成」（slides_visitor_srv.js・ooxml.js）と、テンプレートの登録（assets.js）を、
// この検査の中で組み立てる小さな見本のテンプレート（BNIの資料は使わない）で、実際に作って確かめる。
// サーバー側は本番の *.js を全部、見せかけのスプレッドシート・ドライブ（lib_sheet_fake.js）の上で動かす。
//
//   node tools/check_visitor_slides.js
//
//   ・お名前・会社名・専門分野の「$&」「$'」「$1」を、そのままの文字で入れる
//     （以前は置き換えの記号として読まれ、見本の文字が混ざったり、XMLが壊れて PowerPoint で開けなかったりした）
//   ・テンプレートに2ページ目以降があっても、出力に残さない（前に作ったファイルを登録したときの先週のビジターのページ・
//     同じページが2回出る）。PowerPoint のセクションにも足したページを入れる
//   ・見本の文字を消した枠（ランの無い枠）にも、お名前・専門分野を入れる（以前は空のまま「作成しました」）
//   ・番号が合っているだけの写真・動画の枠では登録しない（ビジター紹介のテンプレートをビジタープレゼンの欄で登録できた）
//   ・Spreadingでキャンセルの方は入れない
//   ・ビジタープレゼン：会社名44pt・【カテゴリー】32ptにそろえる（入らないときは同じ大きさで2行・下のカウントダウンに重ならない範囲で）。
//     カウントダウンが終わっても自動で次の方へ進めない。ファイル名は「定例会20260930_（ビジタープレゼン）MMddHHmm.pptx」
//   ・できたファイルの部品のつながり（tools/pptx_integrity.py）
// 名前はすべて架空。

process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const os = require('os');
const path = require('path');
const vm = require('vm');
const { execFileSync } = require('child_process');
const { makeEnv } = require('./lib_sheet_fake');
const { readZip, writeZip } = require('./lib_zip');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }
function step(name, fn) {
  try { fn(); } catch (e) { fails.push(name + ' で止まった: ' + (e && e.stack ? e.stack.split('\n').slice(0, 3).join(' / ') : e)); }
}
const J = (x) => JSON.stringify(x);

const env = makeEnv({ now: new Date(2026, 8, 29, 10, 0, 0) });
const srv = Object.assign({}, env.globals);
vm.createContext(srv);
for (const f of fs.readdirSync(ROOT).filter((x) => /\.js$/.test(x)).sort()) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), srv, { filename: f });

// ---- 見本のテンプレート ----
const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
  + 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const PML = 'application/vnd.openxmlformats-officedocument.presentationml.';
// text が null の枠は、PowerPointで見本の文字を消したときの形（ランが無く、段落の終わりの書式だけ）
const sp = (id, text, sz) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="枠 ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`
  + `<p:spPr><a:xfrm><a:off x="100000" y="${id * 150000}"/><a:ext cx="5000000" cy="400000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`
  + '<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>'
  + (text == null ? `<a:p><a:endParaRPr lang="ja-JP" sz="${sz || 2000}" dirty="0"><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill></a:endParaRPr></a:p>`
                  : `<a:p><a:r><a:rPr lang="ja-JP" sz="${sz || 2000}"/><a:t>${text}</a:t></a:r></a:p>`)
  + '</p:txBody></p:sp>';
const pic = (id) => `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="動画 ${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>`
  + '<p:blipFill><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr/></p:pic>';
// after … <p:clrMapOvr> のあとに入れるもの（画面切り替え <p:transition> など）
const slideXml = (inner, after) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld><p:spTree>`
  + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'
  + inner + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' + (after || '') + '</p:sld>';
const notesXml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notes ${NS}><p:cSld><p:spTree>`
  + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:notes>';
const rels = (list) => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  + list.map(([id, type, target]) => `<Relationship Id="${id}" Type="${REL}${type}" Target="${target}"/>`).join('') + '</Relationships>';
// slides … ページの中身（1ページ目がひな形）。[中身, clrMapOvr のあと] の形でも。ノートと PowerPoint のセクションも付ける
function makePptx(slides) {
  const f = {};
  f['[Content_Types].xml'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
    + `<Override PartName="/ppt/presentation.xml" ContentType="${PML}presentation.main+xml"/>`
    + slides.map((s, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="${PML}slide+xml"/>`
      + `<Override PartName="/ppt/notesSlides/notesSlide${i + 1}.xml" ContentType="${PML}notesSlide+xml"/>`).join('') + '</Types>';
  f['_rels/.rels'] = rels([['rId1', 'officeDocument', 'ppt/presentation.xml']]);
  f['ppt/presentation.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation ${NS}><p:sldIdLst>`
    + slides.map((s, i) => `<p:sldId id="${256 + i}" r:id="rId${2 + i}"/>`).join('')
    + '</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/>'
    + '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">'
    + '<p14:section name="見本" id="{00000000-0000-4000-8000-000000000001}"><p14:sldIdLst>'
    + slides.map((s, i) => `<p14:sldId id="${256 + i}"/>`).join('') + '</p14:sldIdLst></p14:section></p14:sectionLst></p:ext></p:extLst></p:presentation>';
  f['ppt/_rels/presentation.xml.rels'] = rels(slides.map((s, i) => [`rId${2 + i}`, 'slide', `slides/slide${i + 1}.xml`]));
  slides.forEach((s, i) => {
    f[`ppt/slides/slide${i + 1}.xml`] = Array.isArray(s) ? slideXml(s[0], s[1]) : slideXml(s);
    f[`ppt/slides/_rels/slide${i + 1}.xml.rels`] = rels([['rId1', 'notesSlide', `../notesSlides/notesSlide${i + 1}.xml`]]);
    f[`ppt/notesSlides/notesSlide${i + 1}.xml`] = notesXml();
    f[`ppt/notesSlides/_rels/notesSlide${i + 1}.xml.rels`] = rels([['rId1', 'slide', `../slides/slide${i + 1}.xml`]]);
  });
  return writeZip(f);
}
// ビジター紹介（3人1枚）。1人目のお名前（21）と2人目の専門分野（25）は、見本の文字を消した枠
const INTRO = sp(15, '歓迎 本日のビジター', 3200)
  + sp(17, '専門分野') + sp(18, '招待者') + sp(19, '見本の専門分野') + sp(20, '見本 招待者') + sp(21, null, 2800)
  + sp(23, '専門分野') + sp(24, '招待者') + sp(25, null) + sp(26, '見本 招待者') + sp(28, '見本 名前 様', 2800)
  + sp(29, '専門分野') + sp(30, '招待者') + sp(31, '見本の専門分野') + sp(32, '見本 招待者') + sp(33, '見本 名前 様', 2800);
const LEFTOVER = sp(2, '先週の見本 来人 様', 2800);            // 前に作ったファイルの2ページ目（先週のビジター）
// ビジタープレゼン（1人1枚）。お名前（25）は見本の文字を消した枠
const PRESEN = sp(25, null, 4000) + sp(27, '会社名', 2400) + sp(29, '【カテゴリー】', 2400);
// カウントダウンが終わったら、クリックしなくても次の方へ進む画面切り替え（元の雛形にあったもの）
const AUTO_NEXT = '<p:transition spd="slow" advClick="0" advTm="31000"/>';
const PRESEN_WRONG = sp(25, '2人目の専門分野') + pic(27) + sp(29, '3人目の見出し');   // ビジター紹介のような作り（27 は動画）

// ---- 参加者シート（架空）。「$」の入ったお名前・会社名・専門分野、キャンセルの方、ゲスト・代理 ----
const HEAD = ['No.', '参加者氏名', 'ふりがな', 'カテゴリー', '会社名', '招待者', '備考', '種別', 'メール', 'ステータス'];
const ROWS = [
  ['V01', 'Q$&A 太郎', '', "Rock$'n 業", '見本$1商事', '見本 一郎', '', 'Visitor', '', '参加予定'],
  ['V02', '$1 花子', '', 'US$100 相談', "$'会社", '試験 花子', '', 'Visitor', '', ''],
  ['V03', '取消 した', '', '税理士', '取消商事', '見本 一郎', '', 'Visitor', '', 'キャンセル'],
  ['V04', '四人目 来人', '', '工務店', '四人目社', '架空 三郎', '', 'Visitor', '', ''],
  ['V05', '五人目 来子', '', 'デザイン', '五人目社', '架空 三郎', '', 'Visitor', '', ''],
  ['G01', 'ゲスト 客人', '', '保険', 'ゲスト社', '見本 一郎', '', 'Guest', '', ''],
  ['代理1', '代理 太一', '', '印刷', '代理社', '試験 花子', '', 'Substitute', '', ''],
];
env.reset([
  ['メンバー名簿', false, [['No', '業種区分', '氏名'], ['1', '', '見本 一郎'], ['2', '', '試験 花子'], ['3', '', '架空 三郎']]],
  ['20260930参加者', false, [HEAD].concat(ROWS)],
], { [vm.runInContext('ASSET_ROOT_KEY_', srv)]: 'ASSETROOT' });

const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'vslides_'));
const b64 = (buf) => Buffer.from(buf).toString('base64');
function outputOf(res, label) {
  const id = env.idOfUrl(res.url), f = env.drive.files[id];
  if (!f) { fails.push(label + ': できたファイルが見つからない ' + J(res)); return null; }
  const buf = f.blob._buf, file = path.join(TMP, label + '.pptx');
  fs.writeFileSync(file, buf);
  try { execFileSync('python3', [path.join(__dirname, 'pptx_integrity.py'), file], { stdio: 'pipe' }); ck(true, ''); }
  catch (e) { ck(false, label + ': 部品のつながりが壊れている（PowerPoint が修復を求める）: ' + String(e.stdout || e).slice(0, 400)); }
  const parts = readZip(buf), text = (p) => (parts[p] ? parts[p].toString('utf8') : '');
  const prs = text('ppt/presentation.xml'), prels = text('ppt/_rels/presentation.xml.rels');
  const order = [...prs.matchAll(/<p:sldId id="(\d+)" r:id="([^"]+)"\/>/g)].map((m) => {
    const t = (prels.match(new RegExp('Id="' + m[2] + '"[^>]*Target="([^"]+)"')) || [])[1];
    return { id: m[1], path: 'ppt/' + t };
  });
  const slides = order.map((o) => {
    const x = text(o.path);
    return { path: o.path, xml: x, text: [...x.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1].replace(/&amp;/g, '&').replace(/&apos;/g, "'")).join('|') };
  });
  const sectionIds = [...prs.matchAll(/<p14:sldId id="(\d+)"\/>/g)].map((m) => m[1]);
  return { parts, slides, order, sectionIds, file };
}

// ---- 1) 登録：番号が合っているだけの動画の枠では登録しない。2ページあるテンプレートは、1ページ目だけを使うと知らせる ----
step('テンプレートの登録', () => {
  const wrong = srv.saveTemplateBase64('presen', b64(makePptx([PRESEN_WRONG])), 'wrong.pptx');
  ck(wrong.ok === false && /27/.test(wrong.message) && /文字の枠ではありません/.test(wrong.message),
     '1) 27番が動画のテンプレートを、ビジタープレゼンとして登録した: ' + J(wrong));
  const intro = srv.saveTemplateBase64('intro', b64(makePptx([[INTRO, AUTO_NEXT], LEFTOVER])), 'intro.pptx');
  ck(intro.ok === true && /2ページありますが、使うのは1ページ目だけ/.test(intro.message), '1) ビジター紹介の登録: ' + J(intro.message));
  const presen = srv.saveTemplateBase64('presen', b64(makePptx([[PRESEN, AUTO_NEXT]])), 'presen.pptx');
  ck(presen.ok === true && !/ページありますが/.test(presen.message), '1) ビジタープレゼンの登録: ' + J(presen.message));
});

// ---- 2) ビジター紹介：先週のページを残さない・「$」をそのまま・消した枠にも入れる・キャンセルの方は入れない ----
step('ビジター紹介', () => {
  const res = srv.generateVisitorSlides('20260930参加者', 'intro');
  ck(res.ok === true && /ビジター紹介/.test(res.message) && /取消 した/.test(res.message), '2) ビジター紹介を作れない: ' + J(res));
  const o = outputOf(res, 'intro');
  if (!o) return;
  ck(o.slides.length === 2, '2) ページ数が 2 でない（4名・3人1枚）: ' + o.slides.length + ' ' + J(o.order));
  ck(!o.slides.some((s) => /先週の見本/.test(s.text)), '2) テンプレートの2ページ目（先週のビジター）が残った');
  ck(new Set(o.order.map((x) => x.path)).size === o.order.length, '2) 同じページが2回出る: ' + J(o.order));
  const p1 = o.slides[0] ? o.slides[0].text : '';
  ck(/Q\$&A 太郎 様/.test(p1) && /\$1 花子 様/.test(p1) && /Rock\$'n 業/.test(p1) && /US\$100 相談/.test(p1),
     '2) 「$」の入ったお名前・専門分野が、そのままの文字で入らない: ' + p1);
  ck(!/見本 名前 様/.test(p1) && !/見本の専門分野/.test(p1), '2) 見本の文字が残った・混ざった: ' + p1);
  ck(/<p:cNvPr id="21"[\s\S]*?<a:t>Q\$&amp;A 太郎 様<\/a:t>/.test(o.slides[0].xml), '2) 見本の文字を消した枠（21）に、お名前が入らない');
  ck(/<p:cNvPr id="25"[\s\S]*?<a:t>US\$100 相談<\/a:t>/.test(o.slides[0].xml), '2) 見本の文字を消した枠（25）に、専門分野が入らない');
  ck(/1F3864/.test((o.slides[0].xml.match(/<p:cNvPr id="21"[\s\S]*?<\/p:sp>/) || [''])[0].match(/<a:rPr[\s\S]*?<\/a:rPr>/) || ''),
     '2) 消した枠に入れた文字が、枠の書式（色）を引き継がない');
  ck(!o.slides.some((s) => /取消 した/.test(s.text)), '2) キャンセルの方をスライドに入れた');
  ck(J(o.sectionIds) === J(o.order.map((x) => x.id)), '2) PowerPoint のセクションと、ページの並びが合わない: ' + J([o.sectionIds, o.order.map((x) => x.id)]));
  ck(!o.parts['ppt/notesSlides/notesSlide3.xml'], '2) 取り除いたページのノートが残った');
  // ビジターのスライドは、紹介のページも自動で次へ進めない（テンプレートの「○秒後に次へ」も外す）
  ck(o.slides.every((x) => !/advTm=/.test(x.xml) && !/advClick="0"/.test(x.xml)),
     '2) ビジター紹介のページが、自動で次へ進む: ' + J(o.slides.map((x) => (x.xml.match(/<p:transition\b[^>]*>/) || [''])[0])));
});

// ---- 3) まとめて作成：ビジター2枚・ゲスト1枚・代理1枚。見出しが節ごとに替わる ----
step('まとめて作成', () => {
  const res = srv.generateAllIntroSlides('20260930参加者');
  ck(res.ok === true && res.slideCount === 4, '3) まとめて作成: ' + J(res));
  const o = outputOf(res, 'all');
  if (!o) return;
  ck(J(o.slides.map((s) => s.text.split('|')[0])) === J(['歓迎 本日のビジター', '歓迎 本日のビジター', '歓迎 本日のゲスト', '代理出席の方々']),
     '3) 見出しの並び: ' + J(o.slides.map((s) => s.text.split('|')[0])));
  ck(/ゲスト 客人 様/.test(o.slides[2].text) && /代理 太一 様/.test(o.slides[3].text), '3) ゲスト・代理のページの中身: ' + o.slides.slice(2).map((s) => s.text).join(' / '));
  ck(!o.slides.some((s) => /先週の見本/.test(s.text)), '3) テンプレートの2ページ目が残った');
  ck(o.slides.every((x) => !/advTm=/.test(x.xml)), '3) まとめて作った紹介のページ（ゲスト・代理も）が、自動で次へ進む');
});

// ---- 4) ビジタープレゼン：1人1枚。お名前の枠（見本の文字を消した枠）にも入る ----
step('ビジタープレゼン', () => {
  const res = srv.generateVisitorSlides('20260930参加者', 'presen');
  ck(res.ok === true, '4) ビジタープレゼンを作れない: ' + J(res));
  const o = outputOf(res, 'presen');
  if (!o) return;
  ck(J(o.slides.map((s) => s.text)) === J(['Q$&A 太郎 様|見本$1商事|【Rock$\'n 業】', '$1 花子 様|$\'会社|【US$100 相談】',
                                            '四人目 来人 様|四人目社|【工務店】', '五人目 来子 様|五人目社|【デザイン】']),
     '4) ビジタープレゼンの中身: ' + J(o.slides.map((s) => s.text)));
  // ファイル名は「定例会20260930_（ビジタープレゼン）MMddHHmm.pptx」（開催日と、作った日時）
  ck(/^定例会20260930_（ビジタープレゼン）09291000\.pptx$/.test(res.fileName), '4) ファイル名: ' + res.fileName);
  // 自動で次の方へ進めない（時間で進むのを外し、クリックで進めるように戻す）
  ck(o.slides.every((x) => !/advTm=/.test(x.xml) && !/advClick="0"/.test(x.xml) && /<p:transition\b/.test(x.xml)),
     '4) ビジタープレゼンのページが、自動で次の方へ進む: ' + J(o.slides.map((x) => (x.xml.match(/<p:transition\b[^>]*>/) || [''])[0])));
  // 会社名44pt・【カテゴリー】32ptにそろえる（1人ずつ小さくしない）
  const szOf = (x, id) => ((x.match(new RegExp('<p:cNvPr id="' + id + '"[\\s\\S]*?</p:sp>')) || [''])[0].match(/\ssz="(\d+)"/g) || []).map((v) => v.replace(/\D/g, ''));
  ck(o.slides.every((x) => szOf(x.xml, 27).every((v) => v === '4400') && szOf(x.xml, 29).every((v) => v === '3200')),
     '4) 会社名・カテゴリーの大きさが 44pt・32pt でない: ' + J(o.slides.map((x) => [szOf(x.xml, 27), szOf(x.xml, 29)])));
  // ほかのスライドのファイル名も、同じ「定例会＋開催日_（種類）＋作った日時」
  const intro = srv.generateVisitorSlides('20260930参加者', 'intro'), all = srv.generateAllIntroSlides('20260930参加者');
  ck(/^定例会20260930_（ビジター紹介）\d{8}\.pptx$/.test(intro.fileName) && /^定例会20260930_（紹介スライド・まとめて）\d{8}\.pptx$/.test(all.fileName),
     '4) 紹介スライドのファイル名: ' + J([intro.fileName, all.fileName]));
});

// ---- 4b) ビジタープレゼンの組み方（実際の雛形に近い並び：お名前の下に会社名・【カテゴリー】、その下にカウントダウン）----
//   会社名が短い方も長い方も、【カテゴリー】は32pt。会社名は44pt（入らないほど長いときだけ小さく）。カウントダウンに重ならない
step('ビジタープレゼンの組み方', () => {
  const E = 12700, box = (id, y, h, text, sz, w) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="枠 ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`
    + `<p:spPr><a:xfrm><a:off x="${120 * E}" y="${y * E}"/><a:ext cx="${(w || 620) * E}" cy="${h * E}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`
    + `<p:txBody><a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="ja-JP" sz="${sz}" b="1"/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
  // お名前 → 会社名（32pt・1行ぶんの高さ）→【カテゴリー】（28pt）→ カウントダウン（【カテゴリー】の下43pt）
  const tpl = slideXml(box(25, 100, 60, 'お名前　様', 5400) + box(27, 170, 45, '会社名', 3200) + box(29, 217, 40, '【カテゴリー】', 2800)
    + box(40, 300, 90, '20', 6000, 200), AUTO_NEXT);
  const geo = (x, id) => srv.readShapeGeomEmu_(x, id), sz = (x, id) => [...((x.match(new RegExp('<p:cNvPr id="' + id + '"[\\s\\S]*?</p:sp>')) || [''])[0]).matchAll(/\ssz="(\d+)"/g)].map((m) => +m[1] / 100);
  const bottom = (x, id) => (geo(x, id).y + geo(x, id).cy) / E;
  const make = (company, category) => srv.buildPresenXml_(tpl, { name: '見本 来人', company, category }, 0);
  // 短い会社名：44pt・32pt。会社名の枠が伸びたぶん【カテゴリー】を下げ、カウントダウン（300pt）には重ならない
  let x = make('見本商事', '税理士');
  ck(sz(x, 27).every((v) => v === 44) && sz(x, 29).every((v) => v === 32), '4b) 短い会社名の大きさ: ' + J([sz(x, 27), sz(x, 29)]));
  ck(geo(x, 29).y / E >= bottom(x, 27) - 0.5 && bottom(x, 29) <= 300, '4b) 【カテゴリー】が会社名かカウントダウンに重なる: ' + J([bottom(x, 27), geo(x, 29).y / E, bottom(x, 29)]));
  ck(/<a:noAutofit\/>/.test(x) && !/spAutoFit/.test(x.match(/<p:cNvPr id="27"[\s\S]*?<\/p:sp>/)[0]), '4b) 会社名の枠の「文字に合わせて調整」が残っている');
  // 長い会社名：【カテゴリー】は32ptのまま。会社名は小さくしてもカウントダウンに重ならない
  x = make('見本コンサルティングオフィスとても長い名前株式会社', 'デジタル広告制作');
  ck(sz(x, 29).every((v) => v === 32) && sz(x, 27).every((v) => v < 44 && v >= 16) && bottom(x, 29) <= 300,
     '4b) 長い会社名: ' + J({ company: sz(x, 27), category: sz(x, 29), bottom: bottom(x, 29) }));
  // 2行にするときは「社名／株式会社」で分ける（カウントダウンまでの余地が大きいとき）
  const roomy = tpl.replace(`<a:off x="${120 * E}" y="${300 * E}"/>`, `<a:off x="${120 * E}" y="${420 * E}"/>`);
  const linesOf = (xx) => (xx.match(/<p:cNvPr id="27"[\s\S]*?<\/p:sp>/)[0].match(/<a:t>[^<]*<\/a:t>/g) || []).map((t) => t.replace(/<[^>]+>/g, ''));
  x = srv.buildPresenXml_(roomy, { name: '見本 来人', company: '見本コンサルティング株式会社', category: '税理士' }, 0);
  ck(J(linesOf(x)) === J(['見本コンサルティング', '株式会社']) && sz(x, 27).every((v) => v === 44) && sz(x, 29).every((v) => v === 32) && bottom(x, 29) <= 420,
     '4b) 2行の会社名（44pt）: ' + J({ lines: linesOf(x), company: sz(x, 27), category: sz(x, 29), bottom: bottom(x, 29) }));
  ck(!/advTm=/.test(x), '4b) 自動で次へ進む');
  // 社名の部分が44ptで1行に入らないほど長いときは、カタカナの語の途中で切らず、少し小さくして「社名／株式会社」
  x = srv.buildPresenXml_(roomy, { name: '見本 来人', company: '見本コンサルティングオフィス株式会社', category: '税理士' }, 0);
  ck(J(linesOf(x)) === J(['見本コンサルティングオフィス', '株式会社']) && sz(x, 27).every((v) => v < 44 && v >= 36) && sz(x, 29).every((v) => v === 32),
     '4b) 長い社名の2行: ' + J({ lines: linesOf(x), company: sz(x, 27), category: sz(x, 29) }));
});

// ---- 5) メールのひな形も、お名前の「$」をそのまま差し込む ----
step('メールのひな形', () => {
  const t = srv.fillMailTemplate_('{{name}} 様 / {{inviter}} / {{date}}', { name: 'Q$&A 太郎', inviter: "$'見本", date: '9/30' });
  ck(t === "Q$&A 太郎 様 / $'見本 / 9/30", '5) メールのひな形の「$」: ' + t);
});

step('省略結果の画面にダウンロードリンクを出さない', () => {
  const { loadPage } = require('./lib_minidom');
  const page = loadPage('slides_visitor.html', { server: {
    generateAllIntroSlides: () => ({ ok: true, skipped: true, message: '未登録のため省略しました。' }),
    generateVisitorSlides: () => ({ ok: true, skipped: true, message: '未登録のため省略しました。' })
  } });
  page.els.sheet.innerHTML = '<option value="20260930参加者">20260930参加者</option>';
  page.els.sheet.value = '20260930参加者';
  page.run("data={visitors:[{name:'見本'}],guests:[],dairi:[]};refresh()");
  ck(!page.els.bAll.disabled, '共通テンプレートなしでまとめて作成できない');
  page.run('genAll()');
  page.flush();
  ck(/省略/.test(page.els.msg.innerText) && !/href=/.test(page.els.msg.innerHTML),
    '全体省略時に無効なリンクを表示した');
  page.run("gen('presen')");
  page.flush();
  ck(/省略/.test(page.els.msg.innerText) && !/href=/.test(page.els.msg.innerHTML),
    '単体省略時に無効なリンクを表示した');
  ck(page.fails.length === 0, '画面の例外: ' + J(page.fails));
});

step('未登録のテンプレートは省略して、登録済みの紹介を生成', () => {
  const props = srv.PropertiesService.getScriptProperties();
  const introKey = srv.TEMPLATE_KINDS_.intro.prop;
  const presenKey = srv.TEMPLATE_KINDS_.presen.prop;
  const introId = props.getProperty(introKey), presenId = props.getProperty(presenKey);
  try {
    ck(srv.saveTemplateBase64('guest', b64(makePptx([INTRO])), 'guest.pptx').ok, 'ゲストひな形の登録');
    ck(srv.saveTemplateBase64('dairi', b64(makePptx([INTRO])), 'dairi.pptx').ok, '代理ひな形の登録');
    props.deleteProperty(introKey);
    props.deleteProperty(presenKey);
    const single = srv.generateVisitorSlides('20260930参加者', 'presen');
    ck(single.ok && single.skipped && !single.url, '未登録の単体生成がエラーになった: ' + J(single));
    const res = srv.generateAllIntroSlides('20260930参加者');
    ck(res.ok && !res.skipped && res.slideCount === 2 && /省略.*ビジター/.test(res.message),
      'ビジターを省略してゲスト・代理を生成できない: ' + J(res));
    const o = outputOf(res, 'available-intro');
    if (o) {
      ck(o.slides.length === 2 && o.slides.some(s => /ゲスト 客人/.test(s.text))
        && o.slides.some(s => /代理 太一/.test(s.text)), '登録済みのゲスト・代理が欠けた');
      ck(!o.slides.some(s => /Q\$&A 太郎/.test(s.text)), '未登録のビジターが生成された');
    }
    props.deleteProperty(srv.TEMPLATE_KINDS_.guest.prop);
    props.deleteProperty(srv.TEMPLATE_KINDS_.dairi.prop);
    const empty = srv.generateAllIntroSlides('20260930参加者');
    ck(empty.ok && empty.skipped && empty.slideCount === 0 && !empty.url, '全ひな形未登録時: ' + J(empty));
    const invalid = srv.generateAllIntroSlides('存在しないシート');
    ck(!invalid.ok, '存在しないシートのエラーまで省略した');
  } finally {
    props.setProperty(introKey, introId);
    props.setProperty(presenKey, presenId);
  }
});

step('定例会に同梱されたメンバープレゼンと後半', () => {
  const wide = (id, text, y) => sp(id, text).replace('cx="5000000"', 'cx="7000000"')
    .replace('y="' + (id * 150000) + '"', 'y="' + y + '"');
  const presenter = title => sp(4, title) + wide(10, '苗字 名前', 1500000)
    + wide(11, '会社名', 2500000) + wide(12, 'カテゴリー', 3500000);
  const toParts = slides => srv.unzipToMap_(env.globals.Utilities.newBlob([...makePptx(slides)], 'application/zip', 'combined.pptx'));
  const parts = toParts([sp(4, 'ウィークリープレゼンテーション'), presenter('WEEKLY PRESENTATION'),
    sp(4, '全員終わりましたか？'), presenter('REFERRAL PRESENTATION'), sp(4, '{{抽選1氏名}}')]);
  const item = { kind: 'individual', name: '試験 花子', companyLines: ['見本会社'], categoryLines: ['税理士'], seconds: 0, auto: false };
  const before = srv.slideOrder_(parts).length;
  const weekly = srv.insertMemberPresen_(parts, [item, { ...item, name: '見本 一郎' }]);
  ck(weekly.paths && weekly.paths.length === 2 && /定例会テンプレート内/.test(weekly.message),
    '同梱ひな形を使えない: ' + J(weekly));
  ck(srv.slideOrder_(parts).length === before + 1, '同梱の元ひな形が重複した、または後半を消した');
  ck(srv.rfFindModels_(parts, srv.RF_TITLE_).length === 1, '同梱のリファーラルひな形が失われた');
  const missing = toParts([sp(1, 'ウィークリープレゼンテーション')]);
  const noWeekly = srv.insertMemberPresen_(missing, [item]);
  ck(noWeekly.skipped && srv.slideOrder_(missing).length === 1, 'メンバー用ひな形も無いときに停止・生成した');
  const guests = toParts([sp(4, 'ウィークリープレゼンテーション'), presenter('WEEKLY PRESENTATION'),
    presenter('WEEKLY PRESENTATION').replace('苗字 名前', 'ゲスト 客人'), sp(4, '全員終わりましたか？')]);
  srv.insertMemberPresen_(guests, [item]);
  ck(srv.slideOrder_(guests).some(p => /ゲスト 客人/.test(srv.slideText_(srv.xmlOf_(guests, p)))),
    '同梱ひな形で生成するとアンバサダー・ディレクターのページを消した');
  const skipped = srv.expandPresenterSlides_(missing, [item], { title: srv.RF_TITLE_ });
  ck(skipped.skipped && srv.slideOrder_(missing).length === 1, 'リファーラルひな形なしで個人ページを生成した');
  const oldRoutine = srv.getRoutineInfo, oldMembers = srv.getMemberMaster, oldLists = srv.computeRenewalLists;
  try {
    srv.getRoutineInfo = () => ({ ok: true, found: true, absentees: { names: ['欠席 一郎'] },
      mainPresenters: [{ name: '試験 花子', matched: true, kind: 'member' }],
      recommendations: [{ giver: { name: '試験 花子' }, receiver: { name: '見本 一郎' }, when: 'during' }] });
    srv.getMemberMaster = () => ({ members: [
      { no: '2', name: '試験 花子', company: '見本会社', title: '税理士' },
      { no: '1', name: '見本 一郎', company: '見本商事', title: '印刷' },
      { no: '3', name: '欠席 一郎', company: '欠席会社', title: '保険' }] });
    srv.computeRenewalLists = () => ({ ok: true, text: { d90: '90日', d60: '60日', d30: '30日', overdue: '超過' } });
    const map = {}, opts = { absentNames: ['欠席 一郎'] };
    srv.combinedMeetingOptions_(parts, map, opts, '2026/10/14');
    ck(opts.referral.length === 2 && opts.referral[0].name === '見本 一郎', '同梱の後半で名簿順・欠席除外が違う');
    ck(opts.recommendPairs.length === 1 && opts.lottery[0] === '試験 花子'
      && map['抽選1氏名'] === '試験 花子' && map['更新90'] === '90日', '同梱の後半の材料が欠けた');
    const info = srv.editMeetingSlides_(parts, map, [], opts);
    ck(info.referral && info.referral.made === 2, '前半の出力内で後半の個人ページを生成できない');
    const saved = srv.saveOutputFile_(srv.zipFromMap_(parts, 'combined-output.pptx'), 'combined-output.pptx');
    const result = outputOf(saved, 'combined-output');
    if (result) {
      ck(result.slides.filter(s => /REFERRAL PRESENTATION/.test(s.text)).length === 2,
        '出力したpptxに後半の個人ページがない');
      ck(result.slides.some(s => s.text === '試験 花子'), '同梱された抽選ページの氏名が更新されない');
    }
    const explicit = { referral: [], recommendPairs: [], lottery: [], absentNames: [] }, values = { 更新90: '手直し' };
    srv.combinedMeetingOptions_(parts, values, explicit, '2026/10/14');
    ck(explicit.referral.length === 0 && explicit.recommendPairs.length === 0 && values.更新90 === '手直し', '指定済みの後半設定を上書きした');
    const noModel = {};
    srv.combinedMeetingOptions_(toParts([sp(1, '{{抽選1氏名}}')]), {}, noModel, '2026/10/14');
    ck(noModel.referral === undefined, 'ひな形なしでリファーラル生成を要求した');
  } finally {
    srv.getRoutineInfo = oldRoutine; srv.getMemberMaster = oldMembers; srv.computeRenewalLists = oldLists;
  }
});

fs.rmSync(TMP, { recursive: true, force: true });
if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('ビジター・代理スライド: 検査 ' + checks + ' 件 OK: 「$」をそのまま・テンプレートの2ページ目を残さない・消した枠にも入れる・'
  + '動画の枠では登録しない・キャンセルの方は入れない・部品のつながり・ビジタープレゼンは会社名44pt／カテゴリー32pt（カウントダウンに重ならない）・'
  + '自動で次の方へ進めない・ファイル名（定例会＋開催日_（種類）＋作った日時）');

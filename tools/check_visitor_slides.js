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
const slideXml = (inner) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld><p:spTree>`
  + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'
  + inner + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
const notesXml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notes ${NS}><p:cSld><p:spTree>`
  + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:notes>';
const rels = (list) => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  + list.map(([id, type, target]) => `<Relationship Id="${id}" Type="${REL}${type}" Target="${target}"/>`).join('') + '</Relationships>';
// slides … ページの中身（1ページ目がひな形）。ノートと PowerPoint のセクションも付ける
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
    f[`ppt/slides/slide${i + 1}.xml`] = slideXml(s);
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
  const intro = srv.saveTemplateBase64('intro', b64(makePptx([INTRO, LEFTOVER])), 'intro.pptx');
  ck(intro.ok === true && /2ページありますが、使うのは1ページ目だけ/.test(intro.message), '1) ビジター紹介の登録: ' + J(intro.message));
  const presen = srv.saveTemplateBase64('presen', b64(makePptx([PRESEN])), 'presen.pptx');
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
});

// ---- 5) メールのひな形も、お名前の「$」をそのまま差し込む ----
step('メールのひな形', () => {
  const t = srv.fillMailTemplate_('{{name}} 様 / {{inviter}} / {{date}}', { name: 'Q$&A 太郎', inviter: "$'見本", date: '9/30' });
  ck(t === "Q$&A 太郎 様 / $'見本 / 9/30", '5) メールのひな形の「$」: ' + t);
});

fs.rmSync(TMP, { recursive: true, force: true });
if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('ビジター・代理スライド: 検査 ' + checks + ' 件 OK: 「$」をそのまま・テンプレートの2ページ目を残さない・消した枠にも入れる・'
  + '動画の枠では登録しない・キャンセルの方は入れない・部品のつながり');

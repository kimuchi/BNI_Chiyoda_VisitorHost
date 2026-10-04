// 事前MTG（朝イチMTG）の「熱烈歓迎」のページ（welcome_srv.js・premtg_srv.js・splice_srv.js の importSlideCopies_）を、
// 実データなしで確かめる。本番の *.js を見せかけのスプレッドシート（lib_sheet_fake.js）の上で動かす。名前はすべて架空。
//
//   node tools/check_welcome.js
//
// 確かめること
//   ・ルーティンチェックシートの「新入会」の方ごとに1枚、まとめのページのすぐあとに入る（いなければ作らない。ひな形も開かない）
//   ・お名前・会社名・カテゴリー・チャプター・開催日が入る。名簿の方は名簿から、名簿に無い方は書いてあったお名前とかっこの中のカテゴリー
//   ・写真は名簿の方のメンバー写真（枠に合わせて切り抜く）。写真が無い方は仮の画像のまま（お知らせする）
//   ・既定のひな形は事前MTGと同じ土台なので、マスターを足さない
//   ・土台の違うひな形を登録すると、マスター・レイアウト・テーマごと写す（番号が重ならない・pptxとして壊れていない）。
//     写真の枠の名前が「写真1」でも入る
//   ・作る前の確かめ（役職ごとの入力の画面）に、熱烈歓迎のページの方とひな形が出る
//   ・画像から作ったひな形（「お名前」「お写真」と書いた四角・ページが 20×11.25インチ）：お名前を入れ、「お写真」の四角を
//     同じ位置・大きさの写真にする（飾りの画像はそのまま）。写真の無い方は「お写真」の文字だけ消す。
//     ページの大きさを事前MTGのパワポに合わせて縮める（図形・文字・写したマスターも同じ割合で。はみ出さない）。
//     縦横の比が違うひな形は、収まる大きさで真ん中に置く
//   ・ひな形を確かめる：お名前・お写真の場所が無ければ、登録したときと作る前の確かめで知らせる
//   ・書体はメイリオ：既定のひな形（事前MTG・熱烈歓迎）のテーマの書体・行間（100%・110%）。作ったパワポのテーマの書体は
//     どれもメイリオ（テーマがＭＳ Ｐゴシックのひな形でも）。ひな形で書体を指定した文字はそのまま

process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const os = require('os');
const path = require('path');
const vm = require('vm');
const zlib = require('zlib');
const { spawnSync } = require('child_process');
const { makeEnv } = require('./lib_sheet_fake');
const { readZip, writeZip } = require('./lib_zip');
const { loadPage } = require('./lib_minidom');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }
const J = (x) => JSON.stringify(x);

function png(w, h, rgb) {
  const chunk = (t, d) => { const l = Buffer.alloc(4); l.writeUInt32BE(d.length); const td = Buffer.concat([Buffer.from(t), d]); const c = Buffer.alloc(4); c.writeUInt32BE(zlib.crc32(td) >>> 0); return Buffer.concat([l, td, c]); };
  const ih = Buffer.alloc(13); ih.writeUInt32BE(w, 0); ih.writeUInt32BE(h, 4); ih[8] = 8; ih[9] = 2;
  const row = Buffer.alloc(1 + w * 3); for (let x = 0; x < w; x++) row.set(rgb, 1 + x * 3);
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ih), chunk('IDAT', zlib.deflateSync(Buffer.concat(Array.from({ length: h }, () => row)))), chunk('IEND', Buffer.alloc(0))]);
}

// ---- 架空の名簿・ルーティンチェックシート ----
const HEAD = ['No', '業種区分', '氏名', 'ふりがな', 'カテゴリー', '会社名', '役職', 'メモ', '写真ファイル名', '一言コメント',
  '紹介してほしい人', '協業したい人', '入会日', '更新日', '更新期限日', '会社での役職'];
const ROSTER = [['見本 一郎', 'みほん いちろう', '税理士', '見本会計事務所'], ['新入 花子', 'しんにゅう はなこ', '行政書士（建設業許可・外国人ビザ）', '株式会社新入リーガルサービス'],
                ['試験 二郎', 'しけん じろう', '工務店', '試験工務店']];
const rosterRows = [HEAD].concat(ROSTER.map(([n, k, t, c], i) => HEAD.map((h) => ({ No: String(i + 1), 氏名: n, ふりがな: k, カテゴリー: t, 会社名: c }[h] || ''))));
const DATE = '2026/10/07', NEXT = '2026/10/14';
const row = (c, role, vals) => ['', '', c, '', '', role, '', '', '', (vals || [])[0] || '', '', (vals || [])[1] || ''];
const ROUTINE = [
  ['', '', '開催日', '', '', '', '', '', '', DATE, '', NEXT],
  ['', '', '定例会回数', '', '', '', '', '', '', '536', '', '537'],
  ['', 'No', '内容', '', '', '担当', '期日', '曜日目安', '備考', '', '', ''],
  row('新入会', 'バイス', ['新入 花子さん、新規 太郎さん（エステサロン）', 'なし']),
  row('ウィークリープレゼン', 'プレジ', ['', '']),
];
const env = makeEnv({ now: new Date(2026, 9, 5, 10, 0, 0) });
const F = Object.assign({}, env.globals);
vm.createContext(F);
for (const f of fs.readdirSync(ROOT).filter((x) => /\.js$/.test(x)).sort()) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), F, { filename: f });
env.reset([['メンバー名簿', false, rosterRows], ['休会日', true, [['2026/12/30']]], ['【24期】ルーティンチェックシート', false, ROUTINE]],
          { BNI_CHAPTER: J({ name: '見本', termBase: 23, meetingBaseDate: '2026/03/18', meetingBaseCount: 509 }) });
const htmlCalls = [];
F.HtmlService = { createHtmlOutputFromFile: (n) => { htmlCalls.push(n); return { getContent: () => fs.readFileSync(path.join(ROOT, n + '.html'), 'utf8') }; } };
let SAVED = null;
F.saveOutputFile_ = (blob, name) => { SAVED = { blob, name }; return { id: 'out', url: 'https://example/' + name, downloadUrl: 'https://example/dl/' + name }; };
env.addFile('PHOTO_NEW', new env.FakeBlob(png(300, 400, [10, 120, 200]), 'image/png', 'p.png'), 'p.png');   // 縦長の写真
F.findPhotoIdForName_ = (n) => (String(n).replace(/[\s　]/g, '') === '新入花子' ? 'PHOTO_NEW' : '');
F.CacheService = undefined;
F.saveRoleHolders({ president: '見本 一郎', vice: '試験 二郎' }, 24, DATE);
// 役職のページも1枚（プレジデントの今週の共有事項）
{
  const c0 = F.getRoleInputContext(DATE, 'president');
  const sid = c0.order.find((k) => c0.items[k].title === '今週の共有事項（プレジデント）');
  const sv = F.saveRoleInput(DATE, 'president', [{ id: sid, value: '・見本の共有事項です。', orig: '' }]);
  if (!sv.ok) throw new Error('共有事項を保存できない: ' + sv.message);
}

// ===== 1) 新入会の方 =====
const data = F.premtgData_(new Date(2026, 9, 7));
ck(J(data.newMembers) === J([
  { name: '新入 花子', raw: '新入 花子さん', matched: true, company: '株式会社新入リーガルサービス', category: '行政書士（建設業許可・外国人ビザ）', kana: 'しんにゅう はなこ' },
  { name: '新規 太郎', raw: '新規 太郎さん', matched: false, company: '', category: 'エステサロン', kana: '' }]),
   '1) 新入会の方（名簿の方は名簿から・名簿に無い方は書いてあったとおり）: ' + J(data.newMembers));
ck(F.premtgData_(new Date(2026, 9, 14)).newMembers.length === 0, '1) 「なし」の日に新入会の方がいる');

// ===== 2) 既定のひな形で作る =====
const slidesOf = () => {
  const z = readZip(SAVED.blob._buf);
  const xml = (p) => (z[p] ? z[p].toString('utf8') : '');
  const prs = xml('ppt/presentation.xml'), prels = xml('ppt/_rels/presentation.xml.rels');
  const rid2 = {}; prels.replace(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="slides\/(slide\d+\.xml)"/g, (a, r, s) => { rid2[r] = 'ppt/slides/' + s; });
  const order = [...prs.matchAll(/<p:sldId id="\d+" r:id="([^"]+)"\/>/g)].map((m) => rid2[m[1]]);
  const text = (p) => (xml(p).match(/<a:t>([^<]*)<\/a:t>/g) || []).map((t) => t.replace(/<\/?a:t>/g, '')).join('|');
  return { z, xml, order, text, rels: (p) => xml(p.replace(/^(.*)\/([^/]+)$/, '$1/_rels/$2.rels')) };
};
const integrity = (z, label) => {
  const f = path.join(os.tmpdir(), 'check_welcome_' + process.pid + '_' + label + '.pptx');
  fs.writeFileSync(f, writeZip(Object.fromEntries(Object.entries(z))));
  const r = spawnSync('python3', [path.join(ROOT, 'tools', 'pptx_integrity.py'), f], { encoding: 'utf8' });
  fs.unlinkSync(f);
  return { ok: r.status === 0, out: (r.stdout || '') + (r.stderr || '') };
};
{
  const res = F.generatePreMeetingSlides(DATE);
  ck(res.ok && /熱烈歓迎 2枚/.test(res.message) && /熱烈歓迎のページ: 新入 花子さん、新規 太郎さん（ひな形: 既定のもの）/.test(res.message),
     '2) 作成のお知らせ: ' + J(res.message));
  ck(/写真が見つからない方（仮の画像のまま）: .*新規 太郎/.test(res.message), '2) 写真が無い方のお知らせ: ' + J(res.message));
  const S = slidesOf();
  const kinds = S.order.map((p) => { const t = S.text(p); return /朝イチMTG/.test(t) ? 'まとめ' : /熱烈歓迎/.test(t) ? '歓迎:' + (t.match(/(新入 花子|新規 太郎)さん/) || [])[1] : '役職'; });
  ck(J(kinds.slice(0, 3)) === J(['まとめ', '歓迎:新入 花子', '歓迎:新規 太郎']) && kinds.slice(3).every((k) => k === '役職'),
     '2) 熱烈歓迎のページの場所（まとめのページのすぐあと・お1人1枚）: ' + J(kinds));
  const w1 = S.order[1], w2 = S.order[2], t1 = S.text(w1), t2 = S.text(w2);
  ck(/行政書士（建設業許可・外国人ビザ）/.test(t1) && /株式会社新入リーガルサービス/.test(t1) && /ようこそ、見本チャプターへ！/.test(t1) && /2026年10月7日（水） 定例会/.test(t1),
     '2) 名簿の方の中身（カテゴリー・会社名・チャプター・開催日）: ' + t1);
  ck(/新規 太郎さん/.test(t2) && /エステサロン/.test(t2) && !/\{\{/.test(t1 + t2), '2) 名簿に無い方の中身・差し込み口が残る: ' + t2);
  const r1 = S.rels(w1), r2 = S.rels(w2);
  const photoOf = (x, r) => { const rid = (x.match(/<p:cNvPr id="\d+" name="写真"[\s\S]*?<a:blip r:embed="([^"]+)"/) || [])[1]; return (r.match(new RegExp('Id="' + rid + '"[^>]*Target="([^"]+)"')) || [])[1] || ''; };
  ck(/mpphoto\d+\.png$/.test(photoOf(S.xml(w1), r1)) && /<a:srcRect\b/.test(S.xml(w1)), '2) 名簿の方の写真が入らない・切り抜かない: ' + photoOf(S.xml(w1), r1));
  ck(/welcome_photo/.test(photoOf(S.xml(w2), r2)), '2) 写真が無い方は仮の画像のまま: ' + photoOf(S.xml(w2), r2));
  ck(Object.keys(S.z).filter((p) => /^ppt\/slideMasters\/[^/]+\.xml$/.test(p)).length === 1, '2) 既定のひな形なのにマスターを足した');
  const it = integrity(S.z, 'default');
  ck(it.ok, '2) pptxとして壊れている: ' + it.out);
}

// ===== 3) 土台の違うひな形を登録する（マスター・テーマごと写す）=====
const builtin = (() => { const h = fs.readFileSync(path.join(ROOT, 'welcome_template.html'), 'utf8'); return Buffer.from(h.match(/PPTX_BASE64_BEGIN([\s\S]*?)PPTX_BASE64_END/)[1].replace(/\s/g, ''), 'base64'); })();
function customTemplate() {
  const z = readZip(builtin), put = (p, s) => { z[p] = Buffer.from(s, 'utf8'); }, get = (p) => z[p].toString('utf8');
  const m = get('ppt/slideMasters/slideMaster1.xml');
  const bg = '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFF4D6"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>';
  put('ppt/slideMasters/slideMaster1.xml', /<p:bg>[\s\S]*?<\/p:bg>/.test(m) ? m.replace(/<p:bg>[\s\S]*?<\/p:bg>/, bg) : m.replace(/(<p:cSld(?: [^>]*)?>)/, '$1' + bg));
  put('ppt/theme/theme1.xml', get('ppt/theme/theme1.xml').replace(/(<a:theme\b[^>]*\bname=")[^"]*/, '$1見本の歓迎テーマ'));
  put('ppt/slides/slide1.xml', get('ppt/slides/slide1.xml').replace('name="写真" descr="写真"', 'name="写真1" descr="写真1"'));
  return writeZip(z);
}
{
  env.addFile('WELCOME_TPL', new env.FakeBlob(customTemplate(), 'application/zip', 'w.pptx'), '熱烈歓迎_見本.pptx');
  env.props.BNI_TPL_WELCOME_ID = 'WELCOME_TPL';
  const res = F.generatePreMeetingSlides(DATE);
  ck(res.ok && /ひな形: 登録したもの（熱烈歓迎_見本\.pptx）/.test(res.message), '3) 登録したひな形で作らない: ' + J(res.message));
  const S = slidesOf();
  const masters = Object.keys(S.z).filter((p) => /^ppt\/slideMasters\/[^/]+\.xml$/.test(p)).sort();
  ck(masters.length === 2 && S.z['ppt/theme/theme2.xml'] && /見本の歓迎テーマ/.test(S.xml('ppt/theme/theme2.xml')), '3) マスター・テーマを写さない: ' + J(masters));
  const ids = [...S.xml('ppt/presentation.xml').matchAll(/<p:sldMasterId id="(\d+)"/g)].map((m) => m[1])
    .concat(...masters.map((p) => [...S.xml(p).matchAll(/<p:sldLayoutId id="(\d+)"/g)].map((m) => m[1])));
  ck(ids.length >= 4 && new Set(ids).size === ids.length && ids.every((x) => +x >= 2147483648), '3) マスター・レイアウトの番号が重なる: ' + J(ids));
  const w1 = S.order[1], lay = (S.rels(w1).match(/Target="\.\.\/slideLayouts\/([^"]+)"/) || [])[1];
  const layMaster = (S.xml('ppt/slideLayouts/_rels/' + lay + '.rels').match(/Target="\.\.\/slideMasters\/([^"]+)"/) || [])[1];
  ck(lay && layMaster === 'slideMaster2.xml' && /FFF4D6/.test(S.xml('ppt/slideMasters/slideMaster2.xml')), '3) 熱烈歓迎のページが写したレイアウト・マスターを使っていない: ' + J({ lay, layMaster }));
  ck(/新入 花子さん/.test(S.text(w1)) && /mpphoto\d+\.png/.test(S.rels(w1)), '3) 「写真1」の枠に写真が入らない: ' + S.rels(w1));
  const role = S.order.find((p) => !/熱烈歓迎|朝イチMTG/.test(S.text(p)));
  ck(role && /slideLayout1\.xml/.test(S.rels(role)), '3) 役職のページのレイアウトが変わった: ' + (role ? S.rels(role) : '役職のページが無い'));
  const it = integrity(S.z, 'custom');
  ck(it.ok, '3) pptxとして壊れている（写したマスター）: ' + it.out);
  delete env.props.BNI_TPL_WELCOME_ID;
}

// ===== 4) 新入会の方がいない日：作らない・ひな形も開かない =====
{
  const before = htmlCalls.filter((n) => n === 'welcome_template').length;
  const res = F.generatePreMeetingSlides(NEXT);
  const S = slidesOf();
  ck(res.ok && !/熱烈歓迎/.test(res.message) && !S.order.some((p) => /熱烈歓迎/.test(S.text(p))), '4) 新入会の方がいない日に熱烈歓迎のページを作った: ' + J(res.message));
  ck(htmlCalls.filter((n) => n === 'welcome_template').length === before, '4) 新入会の方がいないのに熱烈歓迎のひな形を開いた');
}

// ===== 5) 作る前の確かめ（役職ごとの入力の画面）=====
{
  const pv = F.getPreMeetingPreview(DATE);
  ck(pv.ok && J(pv.welcome.map((w) => [w.name, w.matched, w.photo])) === J([['新入 花子', true, true], ['新規 太郎', false, false]])
     && pv.welcomeTemplate && pv.welcomeTemplate.registered === false, '5) 作る前の確かめの熱烈歓迎: ' + J({ welcome: pv.welcome, tpl: pv.welcomeTemplate }));
  const page = loadPage('role_input.html', {
    server: { getSystemVersion: () => 'test', getRoleInputContext: (d, r) => F.getRoleInputContext(d, r), getPreMeetingPreview: (d) => F.getPreMeetingPreview(d) },
    fails, now: '2026-10-05T10:00:00',
    preprocess: (p) => p.replace(/<\?\s*var roleParam[\s\S]*?\?>/, '').replace('<?= roleParam ?>', '').replace('<?= viewParam ?>', 'premtg'),
  });
  page.flush();
  page.step('一覧を開く（事前MTGの確かめ）', () => page.window.onload());
  const h = String(page.els.premtgOut && page.els.premtgOut.innerHTML).replace(/<[^>]+>/g, ' ');
  ck(/熱烈歓迎のページ（2枚・まとめのページのあと）/.test(h) && /新入 花子さん/.test(h) && /新規 太郎さん/.test(h) && /名簿にありません/.test(h) && /写真なし/.test(h) && /既定のもの/.test(h),
     '5) 画面に熱烈歓迎のページが出ない: ' + h.replace(/\s+/g, ' ').slice(0, 400));
}

// ===== 6) 画像から作ったひな形（「お名前」「お写真」と書いた四角・ページが 20×11.25インチ）=====
// 画像を背景に敷いて、枠の中に「お名前」（2つの切れ目に分けて書く）と「お写真」の四角を置いたもの。右上に飾りの画像（ロゴ）。
// size を渡すとそのページの大きさで作る（中身もページに収まるように縮めて置く）。marks:false … 目印の四角を置かない
const BIG = { cx: 18288000, cy: 10287000 };
function imageTemplate(opts) {
  const o = opts || {}, size = o.size || BIG;
  const z = readZip(builtin), put = (p, v) => { z[p] = Buffer.isBuffer(v) ? v : Buffer.from(v, 'utf8'); }, get = (p) => z[p].toString('utf8');
  // theme … テーマの書体をその書体にする（いただいたひな形はＭＳ Ｐゴシック）
  if (o.theme) put('ppt/theme/theme1.xml', get('ppt/theme/theme1.xml').replace(/(<a:(?:latin|ea) typeface=")[^"]*(")/g, '$1' + o.theme + '$2')
    .replace(/(<a:font script="Jpan" typeface=")[^"]*(")/g, '$1' + o.theme + '$2'));
  const k = Math.min(size.cx / BIG.cx, size.cy / BIG.cy), S = (v) => Math.round(v * k);
  put('ppt/presentation.xml', get('ppt/presentation.xml').replace(/<p:sldSz cx="\d+" cy="\d+"/, '<p:sldSz cx="' + size.cx + '" cy="' + size.cy + '"'));
  put('ppt/media/bg.png', png(64, 36, [240, 236, 228]));
  put('ppt/media/logo.png', png(20, 20, [200, 30, 40]));
  const box = (id, name, x, y, cx, cy, fill, runs, sz) => '<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="' + name + '"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
    + '<p:spPr><a:xfrm><a:off x="' + S(x) + '" y="' + S(y) + '"/><a:ext cx="' + S(cx) + '" cy="' + S(cy) + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
    + '<a:solidFill><a:srgbClr val="' + fill + '"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>'
    + '<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/>'
    + runs.map((t) => '<a:r><a:rPr lang="ja-JP" altLang="en-US" sz="' + S(sz) + '" dirty="0"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:rPr><a:t>' + t + '</a:t></a:r>').join('')
    + '</a:p></p:txBody></p:sp>';
  const slide = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    + '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    + 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree>'
    + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    + '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    + '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Freeform 2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + size.cx + '" cy="' + size.cy + '"/></a:xfrm>'
    + '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="' + size.cx + '" h="' + size.cy + '"><a:moveTo><a:pt x="0" y="0"/></a:moveTo>'
    + '<a:lnTo><a:pt x="' + size.cx + '" y="0"/></a:lnTo><a:lnTo><a:pt x="' + size.cx + '" y="' + size.cy + '"/></a:lnTo><a:lnTo><a:pt x="0" y="' + size.cy + '"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>'
    + '<a:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr></p:sp>'
    + '<p:pic><p:nvPicPr><p:cNvPr id="5" name="ロゴ"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId3"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
    + '<p:spPr><a:xfrm><a:off x="' + S(16200000) + '" y="' + S(300000) + '"/><a:ext cx="' + S(1500000) + '" cy="' + S(1500000) + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
    + (o.marks === false ? '' : box(3, '正方形/長方形 2', 7162800, 3924300, 3962400, 5257800, 'D9D9D9', ['お写真'], 4800)
                               + box(4, '正方形/長方形 3', 7162800, 3162300, 3962400, 762000, 'D6DCE5', ['お', '名前'], 4000))
    // explicit … 書体を指定した文字（游明朝）を1つ置く
    + (o.explicit ? box(6, '書体を指定した文字', 1000000, 9000000, 5000000, 600000, 'FFFFFF', ['見本の書体指定'], 2400)
      .replace('<a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:rPr>', '<a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="游明朝"/><a:ea typeface="游明朝"/></a:rPr>') : '')
    + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
  put('ppt/slides/slide1.xml', slide);
  put('ppt/slides/_rels/slide1.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
    + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/bg.png"/>'
    + '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/logo.png"/></Relationships>');
  delete z['ppt/media/welcome_photo.png'];
  return writeZip(z);
}
// 図形の位置と大きさ（[x, y, cx, cy]）・ページからはみ出した図形・画像の行き先
const geoOf = (seg) => (seg.match(/<a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/) || []).slice(1).map(Number);
const overOf = (x, sz) => [...x.matchAll(/<a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/g)].map((m) => m.slice(1).map(Number))
  .filter(([X, Y, CX, CY]) => X < -2 || Y < -2 || X + CX > sz[0] + 2 || Y + CY > sz[1] + 2);
const picNamed = (x, name) => (x.match(new RegExp('<p:pic>(?:(?!</p:pic>)[\\s\\S])*?name="' + name + '"(?:(?!</p:pic>)[\\s\\S])*</p:pic>')) || [''])[0];
const targetOf = (rels, seg) => { const rid = (seg.match(/r:embed="([^"]+)"/) || [])[1]; return (rels.match(new RegExp('Id="' + rid + '"[^>]*Target="([^"]+)"')) || [])[1] || ''; };
// ドライブのファイル：登録のときに種類と大きさを見るので、見せかけのファイルにも足す
{
  const orig = F.DriveApp.getFileById;
  F.DriveApp.getFileById = (id) => { const f = orig(id); return Object.assign({ getMimeType: () => F.PPTX_MIME_, getSize: () => f.getBlob()._buf.length }, f); };
}
{
  env.addFile('WELCOME_IMAGE_TPL_0001', new env.FakeBlob(imageTemplate(), 'application/zip', 'w.pptx'), '熱烈歓迎_画像から.pptx');
  env.props.BNI_TPL_WELCOME_ID = 'WELCOME_IMAGE_TPL_0001';
  const res = F.generatePreMeetingSlides(DATE);
  ck(res.ok && /熱烈歓迎 2枚/.test(res.message), '6) 画像から作ったひな形で作らない: ' + J(res.message));
  ck(/熱烈歓迎のひな形のページの大きさを、事前MTGのパワポの大きさに合わせました（67%）/.test(res.message), '6) 大きさを合わせたお知らせ: ' + J(res.message));
  const S = slidesOf();
  const sz = S.xml('ppt/presentation.xml').match(/<p:sldSz cx="(\d+)" cy="(\d+)"/).slice(1).map(Number);
  const f = Math.min(sz[0] / BIG.cx, sz[1] / BIG.cy), dx = Math.round((sz[0] - BIG.cx * f) / 2), dy = Math.round((sz[1] - BIG.cy * f) / 2);
  const w1 = S.order[1], w2 = S.order[2], x1 = S.xml(w1), x2 = S.xml(w2), t1 = S.text(w1), t2 = S.text(w2);
  ck(/新入 花子/.test(t1) && /新規 太郎/.test(t2) && !/お名前|お写真|\{\{/.test(t1 + t2), '6) お名前が入らない・目印の文字が残る: ' + J([t1, t2]));
  // 「お写真」の四角が、同じ位置・大きさ（縮めたもの）の写真の画像になる。飾りの画像（ロゴ）は写真にしない
  const ph = picNamed(x1, '写真'), want = [Math.round(7162800 * f + dx), Math.round(3924300 * f + dy), Math.round(3962400 * f), Math.round(5257800 * f)];
  ck(ph && geoOf(ph).every((v, i) => Math.abs(v - want[i]) <= 2) && /<a:srcRect\b/.test(ph) && /prst="rect"/.test(ph),
     '6) 「お写真」の四角が写真にならない: ' + J({ geo: geoOf(ph), want }));
  ck(/mpphoto\d+\.png$/.test(targetOf(S.rels(w1), ph)), '6) 写真がメンバー写真でない: ' + targetOf(S.rels(w1), ph));
  const logo = picNamed(x1, 'ロゴ');
  ck(logo && !/mpphoto/.test(targetOf(S.rels(w1), logo)) && /\.png$/.test(targetOf(S.rels(w1), logo)), '6) 飾りの画像（ロゴ）を写真にした: ' + targetOf(S.rels(w1), logo));
  // 写真の無い方：「お写真」の文字だけ消して、四角はそのまま
  ck(!picNamed(x2, '写真') && /name="正方形\/長方形 2"/.test(x2) && /D9D9D9/.test(x2), '6) 写真の無い方の「お写真」の四角が残らない');
  ck(/写真が見つからない方[^\n]*新規 太郎/.test(res.message), '6) 写真の無い方を知らせない: ' + J(res.message));
  // ページからはみ出さない。文字の大きさ・写したマスターも同じ割合で縮める
  ck(!overOf(x1, sz).length && !overOf(x2, sz).length, '6) ページからはみ出す: ' + J(overOf(x1, sz)));
  ck(/sz="2667"/.test(x1) && /sz="3200"/.test(x2), '6) 文字の大きさを同じ割合で縮めない: ' + J([...x1.matchAll(/sz="(\d+)"/g)].map((m) => m[1])));
  const lay = (S.rels(w1).match(/Target="\.\.\/slideLayouts\/([^"]+)"/) || [])[1];
  const mst = (S.xml('ppt/slideLayouts/_rels/' + lay + '.rels').match(/Target="\.\.\/slideMasters\/([^"]+)"/) || [])[1];
  const bz = readZip(builtin), bsz = [...bz['ppt/slideMasters/slideMaster1.xml'].toString('utf8').matchAll(/\bsz="(\d+)"/g)].map((m) => +m[1]);
  const msz = [...S.xml('ppt/slideMasters/' + mst).matchAll(/\bsz="(\d+)"/g)].map((m) => +m[1]);
  ck(mst && mst !== 'slideMaster1.xml' && J(msz) === J(bsz.map((v) => Math.max(100, Math.round(v * f)))) && !overOf(S.xml('ppt/slideMasters/' + mst), sz).length,
     '6) 写したマスターを同じ割合で縮めない: ' + J({ mst, msz: msz.slice(0, 4), bsz: bsz.slice(0, 4) }));
  const it = integrity(S.z, 'image');
  ck(it.ok, '6) pptxとして壊れている（画像から作ったひな形）: ' + it.out);

  // 縦横の比が違うひな形（4:3）→ 収まる大きさで真ん中に置く
  env.addFile('WELCOME_IMAGE_TPL_0001', new env.FakeBlob(imageTemplate({ size: { cx: 9144000, cy: 6858000 } }), 'application/zip', 'w.pptx'), '熱烈歓迎_4対3.pptx');
  const res2 = F.generatePreMeetingSlides(DATE);
  ck(res2.ok && /縦横の比が、事前MTGのパワポと違います。ページに収まる大きさ（100%）にして、真ん中に置きました/.test(res2.message), '6) 縦横の比が違うお知らせ: ' + J(res2.message));
  const S2 = slidesOf(), x3 = S2.xml(S2.order[1]);
  const bgGeo = geoOf((x3.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?name="Freeform 2"[\s\S]*?<\/p:sp>/) || [''])[0]);
  ck(J(bgGeo) === J([Math.round((sz[0] - 9144000) / 2), 0, 9144000, 6858000]) && !overOf(x3, sz).length && /新入 花子/.test(S2.text(S2.order[1])),
     '6) 4:3 のひな形を真ん中に置かない・はみ出す: ' + J({ bgGeo, over: overOf(x3, sz) }));
  ck(integrity(S2.z, 'ratio').ok, '6) pptxとして壊れている（4:3 のひな形）');
}

// ===== 7) ひな形を確かめる（登録したとき・作る前の確かめ）=====
if (typeof F.welcomeTemplateCheck_ !== 'function') fails.push('7) ひな形を確かめる道具（welcomeTemplateCheck_）が無い');
else {
  const unz = (buf) => F.unzipToMap_(new env.FakeBlob(buf, 'application/zip', 't.pptx'));
  const good = F.welcomeTemplateCheck_(unz(imageTemplate()));
  ck(good.name && good.photo === 'mark' && good.notes.length === 0, '7) 「お名前」「お写真」のひな形を確かめる: ' + J(good));
  const std = F.welcomeTemplateCheck_(unz(builtin));
  ck(std.name && std.photo === 'pic' && std.notes.length === 0, '7) 既定のひな形を確かめる: ' + J(std));
  const none = F.welcomeTemplateCheck_(unz(imageTemplate({ marks: false })));
  ck(!none.name && none.photo === '' && none.notes.length === 2 && /お名前を入れる場所がありません/.test(none.notes[0]) && /お写真を入れる場所がありません/.test(none.notes[1]),
     '7) 目印の無いひな形を知らせない: ' + J(none));
  env.addFile('WELCOME_IMAGE_TPL_0001', new env.FakeBlob(imageTemplate(), 'application/zip', 'w.pptx'), '熱烈歓迎_画像から.pptx');
  const reg = F.registerBigTemplate('welcome', 'https://drive.google.com/file/d/WELCOME_IMAGE_TPL_0001/view?usp=sharing');
  ck(reg.ok && /お名前と、お写真（「お写真」の図形）を入れる場所が見つかりました/.test(reg.message), '7) 登録のお知らせ: ' + J(reg.message));
  env.addFile('WELCOME_NO_MARKS_0001', new env.FakeBlob(imageTemplate({ marks: false }), 'application/zip', 'w.pptx'), '熱烈歓迎_目印なし.pptx');
  const reg2 = F.registerBigTemplate('welcome', 'WELCOME_NO_MARKS_0001');
  ck(reg2.ok && /⚠ お名前を入れる場所がありません/.test(reg2.message) && /⚠ お写真を入れる場所がありません/.test(reg2.message), '7) 目印の無いひな形を登録したときのお知らせ: ' + J(reg2.message));
  // 作る前の確かめ：画面に知らせが出る
  const pv = F.getPreMeetingPreview(DATE);
  ck(pv.ok && pv.welcomeTemplate && pv.welcomeTemplate.notes.length === 2, '7) 作る前の確かめにひな形の知らせが無い: ' + J(pv.welcomeTemplate));
  const page = loadPage('role_input.html', {
    server: { getSystemVersion: () => 'test', getRoleInputContext: (d, r) => F.getRoleInputContext(d, r), getPreMeetingPreview: (d) => F.getPreMeetingPreview(d) },
    fails, now: '2026-10-05T10:00:00',
    preprocess: (p) => p.replace(/<\?\s*var roleParam[\s\S]*?\?>/, '').replace('<?= roleParam ?>', '').replace('<?= viewParam ?>', 'premtg'),
  });
  page.flush();
  page.step('一覧を開く（事前MTGの確かめ・目印の無いひな形）', () => page.window.onload());
  const h = String(page.els.premtgOut && page.els.premtgOut.innerHTML).replace(/<[^>]+>/g, ' ');
  ck(/⚠ お名前を入れる場所がありません/.test(h) && /⚠ お写真を入れる場所がありません/.test(h), '7) 画面にひな形の知らせが出ない: ' + h.replace(/\s+/g, ' ').slice(0, 300));
  delete env.props.BNI_TPL_WELCOME_ID;
}

// ===== 8) 書体はメイリオ =====
{
  const fontsOf = (themeXml) => ['majorFont', 'minorFont'].map((t) => {
    const seg = (themeXml.match(new RegExp('<a:' + t + '>([\\s\\S]*?)</a:' + t + '>')) || [])[1] || '';
    return [(seg.match(/<a:latin typeface="([^"]*)"/) || [])[1], (seg.match(/<a:ea typeface="([^"]*)"/) || [])[1], (seg.match(/<a:font script="Jpan" typeface="([^"]*)"/) || [])[1]];
  });
  const allMeiryo = (themeXml) => fontsOf(themeXml).every((f) => f.every((v) => v === 'メイリオ'));
  // 既定のひな形（同梱の埋め込み用と docs/templates の pptx が同じ・テーマの書体・行間）
  const emb = (name) => Buffer.from(fs.readFileSync(path.join(ROOT, name), 'utf8').match(/PPTX_BASE64_BEGIN([\s\S]*?)PPTX_BASE64_END/)[1].replace(/\s/g, ''), 'base64');
  const pm = emb('premtg_template.html'), wl = emb('welcome_template.html');
  ck(pm.equals(fs.readFileSync(path.join(ROOT, 'docs/templates/BNI_テンプレート_事前MTG.pptx'))) && wl.equals(fs.readFileSync(path.join(ROOT, 'docs/templates/BNI_テンプレート_熱烈歓迎.pptx'))),
     '8) 埋め込み用のひな形と docs/templates の pptx が違う');
  const pz = readZip(pm), wz = readZip(wl);
  ck(allMeiryo(pz['ppt/theme/theme1.xml'].toString('utf8')) && allMeiryo(wz['ppt/theme/theme1.xml'].toString('utf8')),
     '8) 既定のひな形の書体がメイリオでない: ' + J([fontsOf(pz['ppt/theme/theme1.xml'].toString('utf8')), fontsOf(wz['ppt/theme/theme1.xml'].toString('utf8'))]));
  const ln = Object.keys(pz).filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p)).sort()
    .map((p) => [...pz[p].toString('utf8').matchAll(/<a:spcPct val="(\d+)"/g)].map((m) => m[1]));
  ck(J(ln) === J([['100000', '100000', '100000'], ['110000', '110000'], ['110000']]), '8) 既定のひな形の行間（メイリオに合わせて詰める）: ' + J(ln));
  // 作ったパワポ：テーマがＭＳ Ｐゴシックのひな形でも、テーマの書体はどれもメイリオ。書体を指定した文字はそのまま
  env.addFile('WELCOME_IMAGE_TPL_0001', new env.FakeBlob(imageTemplate({ theme: 'ＭＳ Ｐゴシック', explicit: true }), 'application/zip', 'w.pptx'), '熱烈歓迎_ＭＳＰゴシック.pptx');
  env.props.BNI_TPL_WELCOME_ID = 'WELCOME_IMAGE_TPL_0001';
  const res = F.generatePreMeetingSlides(DATE);
  const S = slidesOf(), themes = Object.keys(S.z).filter((p) => /^ppt\/theme\/[^/]+\.xml$/.test(p)).sort();
  ck(res.ok && themes.length === 2 && themes.every((p) => allMeiryo(S.xml(p))), '8) 作ったパワポのテーマの書体がメイリオでない: ' + J(themes.map((p) => [p, fontsOf(S.xml(p))])));
  const w1 = S.order[1], x1 = S.xml(w1);
  ck(/新入 花子/.test(S.text(w1)) && /<a:latin typeface="游明朝"\/><a:ea typeface="游明朝"\/>/.test(x1), '8) 書体を指定した文字が変わった');
  ck(!/typeface="ＭＳ Ｐゴシック"/.test(Object.keys(S.z).filter((p) => /^ppt\/(slides|slideLayouts|slideMasters|theme)\//.test(p)).map((p) => S.xml(p)).join('')),
     '8) ＭＳ Ｐゴシックが残る');
  ck(integrity(S.z, 'font').ok, '8) pptxとして壊れている（書体）');
  delete env.props.BNI_TPL_WELCOME_ID;
}

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('熱烈歓迎のページ: 検査 ' + checks + ' 件 OK: 新入会の方ごとに1枚・まとめのページのあと・名簿に無い方・写真・'
  + '土台の違うひな形はマスターごと写す（番号が重ならない・壊れていない）・いない日は作らない・作る前の確かめ・'
  + '画像から作ったひな形（「お名前」「お写真」・大きさを合わせる・縦横の比が違う）・ひな形を確かめる・書体はメイリオ');

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
//     写真の枠の名前が「写真1」でも入る。ページの大きさが違えばお知らせする
//   ・作る前の確かめ（役職ごとの入力の画面）に、熱烈歓迎のページの方とひな形が出る

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
function customTemplate(opts) {
  const z = readZip(builtin), put = (p, s) => { z[p] = Buffer.from(s, 'utf8'); }, get = (p) => z[p].toString('utf8');
  const m = get('ppt/slideMasters/slideMaster1.xml');
  const bg = '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFF4D6"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>';
  put('ppt/slideMasters/slideMaster1.xml', /<p:bg>[\s\S]*?<\/p:bg>/.test(m) ? m.replace(/<p:bg>[\s\S]*?<\/p:bg>/, bg) : m.replace(/(<p:cSld(?: [^>]*)?>)/, '$1' + bg));
  put('ppt/theme/theme1.xml', get('ppt/theme/theme1.xml').replace(/(<a:theme\b[^>]*\bname=")[^"]*/, '$1見本の歓迎テーマ'));
  put('ppt/slides/slide1.xml', get('ppt/slides/slide1.xml').replace('name="写真" descr="写真"', 'name="写真1" descr="写真1"'));
  if (opts && opts.size) put('ppt/presentation.xml', get('ppt/presentation.xml').replace(/<p:sldSz cx="\d+" cy="\d+"/, '<p:sldSz cx="9144000" cy="6858000"'));
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
  // ページの大きさが違うひな形
  env.addFile('WELCOME_TPL', new env.FakeBlob(customTemplate({ size: true }), 'application/zip', 'w.pptx'), '熱烈歓迎_4対3.pptx');
  const res2 = F.generatePreMeetingSlides(DATE);
  ck(res2.ok && /ページの大きさが、事前MTGのひな形と違います/.test(res2.message), '3) ページの大きさが違うことを知らせない: ' + J(res2.message));
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

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('熱烈歓迎のページ: 検査 ' + checks + ' 件 OK: 新入会の方ごとに1枚・まとめのページのあと・名簿に無い方・写真・'
  + '土台の違うひな形はマスターごと写す（番号が重ならない・壊れていない）・いない日は作らない・作る前の確かめ');

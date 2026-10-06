// 後半スライドの1枚目に、前半の最後のページ（メインプレゼンテーション）を入れる（meeting_slides_srv.js の insertMainPresenPage_・
// generateMeetingSlides）を、同梱のひな形から作った作り物の前半・後半のテンプレートで確かめる。名前はすべて架空。
//
//   node tools/check_second_main_page.js
//
// 作り物：前半 … 熱烈歓迎のひな形のページを、メインプレゼンのページ（差し込み口 {{メインプレゼン1氏名}} など）にしたもの
//         後半 … 事前MTGのひな形（3枚）に、PowerPoint のセクションを2つ付けたもの
// 確かめること
//   ・前半のテンプレートのメインプレゼンのページを、見た目（レイアウト・マスター）ごと写して、後半の1枚目に入れる。
//     土台の違うテンプレートでも（マスターごと写す）。pptxとして壊れていない。セクションも並びに合わせる
//   ・後半のテンプレートにもうそのページがあれば、写さない（2枚にしない）。前半のテンプレートに無ければ、知らせて入れない
//   ・作成（generateMeetingSlides）：画面で選んだメインプレゼンのお名前・会社名・カテゴリーが入る。チェックを外せば入れない。
//     前半のテンプレートが未登録なら、知らせて入れない（後半はふだんどおり作る）
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
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }
const J = (x) => JSON.stringify(x);

const HEAD = ['No', '業種区分', '氏名', 'ふりがな', 'カテゴリー', '会社名'];
const ROSTER = [['1', '', '見本 一郎', 'みほん いちろう', '税理士', '見本会計事務所'], ['2', '', '試験 二郎', 'しけん じろう', '工務店', '試験工務店']];
const env = makeEnv({ now: new Date(2026, 9, 5, 10, 0, 0) });
const F = Object.assign({}, env.globals);
vm.createContext(F);
for (const f of fs.readdirSync(ROOT).filter((x) => /\.js$/.test(x)).sort()) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), F, { filename: f });
env.reset([['メンバー名簿', false, [HEAD].concat(ROSTER)], ['休会日', true, [['2026/12/30']]]],
          { BNI_CHAPTER: J({ name: '見本', termBase: 23, meetingBaseDate: '2026/03/18', meetingBaseCount: 509 }) });
let SAVED = null;
F.saveOutputFile_ = (blob, name) => { SAVED = { blob, name }; return { id: 'out', url: 'https://example/' + name, downloadUrl: 'https://example/dl/' + name }; };
F.findPhotoIdForName_ = () => '';

const tpl = (f) => fs.readFileSync(path.join(ROOT, 'docs', 'templates', f));
// 前半：熱烈歓迎のひな形のページを、メインプレゼンのページにする
function firstHalf({ token = true, otherMaster = false } = {}) {
  const z = readZip(tpl('BNI_テンプレート_熱烈歓迎.pptx')), get = (p) => z[p].toString('utf8'), put = (p, s) => { z[p] = Buffer.from(s, 'utf8'); };
  let s = get('ppt/slides/slide1.xml').replace('🎉 熱烈歓迎 🎉', '今週のメインプレゼンテーション');
  if (token) s = s.replace('{{氏名}}さん', '{{メインプレゼン1氏名}}').replace('{{会社名}}', '{{メインプレゼン1会社名}}').replace('{{カテゴリー}}', '{{メインプレゼン1カテゴリー}}');
  put('ppt/slides/slide1.xml', s);
  if (otherMaster) {
    const m = get('ppt/slideMasters/slideMaster1.xml'), bg = '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFF4D6"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>';
    put('ppt/slideMasters/slideMaster1.xml', /<p:bg>[\s\S]*?<\/p:bg>/.test(m) ? m.replace(/<p:bg>[\s\S]*?<\/p:bg>/, bg) : m.replace(/(<p:cSld(?: [^>]*)?>)/, '$1' + bg));
    const l = Object.keys(z).filter((p) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(p));
    l.forEach((p) => put(p, get(p).replace(/(<p:cSld\b[^>]*\bname=")[^"]*/, '$1見本の前半のレイアウト')));
    put('ppt/theme/theme1.xml', get('ppt/theme/theme1.xml').replace(/(<a:theme\b[^>]*\bname=")[^"]*/, '$1見本の前半のテーマ'));
  }
  return writeZip(z);
}
// 後半：事前MTGのひな形（3枚）に、セクションを2つ（はじめ：1枚目／あと：2〜3枚目）
function secondHalf({ token = false } = {}) {
  const z = readZip(tpl('BNI_テンプレート_事前MTG.pptx')), get = (p) => z[p].toString('utf8'), put = (p, s) => { z[p] = Buffer.from(s, 'utf8'); };
  let prs = get('ppt/presentation.xml');
  const ids = [...prs.matchAll(/<p:sldId id="(\d+)"/g)].map((m) => m[1]);
  const sec = '<p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">'
    + `<p14:section name="はじめ" id="{00000000-0000-0000-0000-000000000001}"><p14:sldIdLst><p14:sldId id="${ids[0]}"/></p14:sldIdLst></p14:section>`
    + `<p14:section name="あと" id="{00000000-0000-0000-0000-000000000002}"><p14:sldIdLst>${ids.slice(1).map((i) => `<p14:sldId id="${i}"/>`).join('')}</p14:sldIdLst></p14:section>`
    + '</p14:sectionLst></p:ext>';
  prs = /<p:extLst>/.test(prs) ? prs.replace('<p:extLst>', '<p:extLst>' + sec) : prs.replace('</p:presentation>', '<p:extLst>' + sec + '</p:extLst></p:presentation>');
  put('ppt/presentation.xml', prs);
  if (token) put('ppt/slides/slide2.xml', get('ppt/slides/slide2.xml').replace(/<a:t>[^<]*<\/a:t>/, '<a:t>{{メインプレゼン1氏名}}</a:t>'));
  return writeZip(z);
}
const partsOf = (buf) => { const z = readZip(buf), m = {}; Object.keys(z).forEach((k) => { m[k] = new env.FakeBlob(z[k], '', k); }); return m; };
const zipOf = (parts) => { const o = {}; Object.keys(parts).forEach((k) => { o[k] = parts[k]._buf || Buffer.from(parts[k].getDataAsString(), 'utf8'); }); return o; };
const integrity = (z, label) => {
  const f = path.join(os.tmpdir(), 'check_second_main_' + process.pid + '_' + label + '.pptx');
  fs.writeFileSync(f, writeZip(z));
  const r = spawnSync('python3', [path.join(ROOT, 'tools', 'pptx_integrity.py'), f], { encoding: 'utf8' });
  fs.unlinkSync(f);
  return { ok: r.status === 0, out: (r.stdout || '') + (r.stderr || '') };
};
const masters = (parts) => Object.keys(parts).filter((p) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(p)).length;
const sectionIds = (parts) => [...F.xmlOf_(parts, 'ppt/presentation.xml').matchAll(/<p14:section\b[^>]*name="([^"]+)"[^>]*>([\s\S]*?)<\/p14:section>/g)]
  .map((m) => [m[1], [...m[2].matchAll(/<p14:sldId id="(\d+)"/g)].map((x) => x[1])]);

// ===== 1) 同じ土台のテンプレート：1枚目に入る（マスターは足さない）=====
{
  const parts = partsOf(secondHalf()), src = partsOf(firstHalf()), before = F.slideEntries_(parts);
  const r = F.insertMainPresenPage_(parts, src), after = F.slideEntries_(parts);
  ck(r.path && after.length === before.length + 1 && after[0].path === r.path, '1) 1枚目に入っていない: ' + J({ r, after: after.map((e) => e.path) }));
  ck(F.xmlOf_(parts, r.path).includes('{{メインプレゼン1氏名}}') && /今週のメインプレゼンテーション/.test(F.slideText_(F.xmlOf_(parts, r.path))), '1) 写したページの中身が違う');
  ck(J(after.slice(1).map((e) => e.path)) === J(before.map((e) => e.path)), '1) もとのページの並びが変わった');
  ck(masters(parts) === 1, '1) 同じ土台なのにマスターを足した: ' + masters(parts));
  const secs = sectionIds(parts);
  ck(secs.length === 2 && secs[0][1][0] === String(after[0].id) && secs[0][1].length === 2, '1) セクションが並びに合っていない: ' + J(secs));
  ck(/1枚目/.test(r.message), '1) 知らせ: ' + r.message);
  const it = integrity(zipOf(parts), 'same');
  ck(it.ok, '1) pptxとして壊れている: ' + it.out);
}

// ===== 2) 土台の違うテンプレート：マスターごと写す =====
{
  const parts = partsOf(secondHalf()), src = partsOf(firstHalf({ otherMaster: true }));
  const r = F.insertMainPresenPage_(parts, src);
  ck(r.path && F.slideEntries_(parts)[0].path === r.path && masters(parts) === 2, '2) 土台の違うテンプレートのページ: ' + J({ r, masters: masters(parts) }));
  const it = integrity(zipOf(parts), 'other');
  ck(it.ok, '2) pptxとして壊れている（マスターごと写したとき）: ' + it.out);
}

// ===== 3) 後半にもうある・前半に無い =====
{
  const parts = partsOf(secondHalf({ token: true })), n = F.slideEntries_(parts).length;
  const r = F.insertMainPresenPage_(parts, partsOf(firstHalf()));
  ck(!r.path && !r.message && F.slideEntries_(parts).length === n, '3) 後半のテンプレートにもうあるのに写した: ' + J(r));
  const p2 = partsOf(secondHalf()), r2 = F.insertMainPresenPage_(p2, partsOf(firstHalf({ token: false })));
  ck(!r2.path && /見つからない/.test(r2.message) && F.slideEntries_(p2).length === n, '3) 前半に無いときの知らせ: ' + J(r2));
}

// ===== 4) 作成：お名前・会社名・カテゴリーが入る／チェックを外せば入れない／前半のテンプレートが未登録 =====
{
  env.addFile('TPL1', new env.FakeBlob(firstHalf(), 'application/vnd.openxmlformats-officedocument.presentationml.presentation', 'first.pptx'), 'first.pptx');
  env.addFile('TPL2', new env.FakeBlob(secondHalf(), 'application/vnd.openxmlformats-officedocument.presentationml.presentation', 'second.pptx'), 'second.pptx');
  env.props.BNI_TPL_MEETING_FIRST_ID = 'TPL1';
  env.props.BNI_TPL_MEETING_SECOND_ID = 'TPL2';
  const values = { 開催回: '536', 開催日: '2026/10/7(水) 第536回', 開催日付: '2026/10/07',
    メインプレゼン1氏名: '見本 一郎', メインプレゼン1会社名: '見本会計事務所', メインプレゼン1カテゴリー: '【税理士】',
    メインプレゼン2氏名: '試験 二郎', メインプレゼン2会社名: '試験工務店', メインプレゼン2カテゴリー: '【工務店】' };
  const firstText = () => {
    const z = readZip(SAVED.blob._buf), xml = (p) => (z[p] ? z[p].toString('utf8') : '');
    const rid2 = {}; xml('ppt/_rels/presentation.xml.rels').replace(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="slides\/(slide\d+\.xml)"/g, (a, r, s) => { rid2[r] = 'ppt/slides/' + s; });
    const order = [...xml('ppt/presentation.xml').matchAll(/<p:sldId id="\d+" r:id="([^"]+)"\/>/g)].map((m) => rid2[m[1]]);
    return { text: (xml(order[0]).match(/<a:t>([^<]*)<\/a:t>/g) || []).map((t) => t.replace(/<\/?a:t>/g, '')).join('|'), n: order.length, z };
  };
  let res = F.generateMeetingSlides('meetingSecond', values, '2026/10/07', { mainPage: true, mainPresenters: ['見本 一郎', '試験 二郎'] });
  let s = res.ok && firstText();
  ck(res.ok && /1枚目に、メインプレゼンのページ/.test(res.message), '4) 作成の知らせ: ' + J(res.message));
  ck(s && /今週のメインプレゼンテーション/.test(s.text) && /見本 一郎/.test(s.text) && /見本会計事務所/.test(s.text) && /【税理士】/.test(s.text) && !/\{\{/.test(s.text),
     '4) 1枚目のメインプレゼンのページの中身: ' + (s && s.text));
  ck(s && s.n === 4, '4) ページの数（後半3枚＋1枚）: ' + (s && s.n));
  const it = s && integrity(s.z, 'gen');
  ck(it && it.ok, '4) 作ったpptxが壊れている: ' + (it && it.out));
  res = F.generateMeetingSlides('meetingSecond', values, '2026/10/07', { mainPage: false, mainPresenters: ['見本 一郎', '試験 二郎'] });
  s = res.ok && firstText();
  ck(res.ok && s.n === 3 && !/メインプレゼンテーション/.test(s.text) && !/メインプレゼンのページ/.test(res.message), '4) チェックを外したのに入れた: ' + J(res.message));
  delete env.props.BNI_TPL_MEETING_FIRST_ID;
  res = F.generateMeetingSlides('meetingSecond', values, '2026/10/07', { mainPage: true, mainPresenters: ['見本 一郎', '試験 二郎'] });
  s = res.ok && firstText();
  ck(res.ok && s.n === 3 && /前半のテンプレートを開けません/.test(res.message), '4) 前半のテンプレートが未登録のとき: ' + J(res.message));
}

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('後半の1枚目のメインプレゼンのページ: 検査 ' + checks + ' 件 OK: 前半のテンプレートから見た目ごと写す（土台が違えばマスターも）・'
  + 'セクション・2枚にしない・前半に無い／未登録なら知らせる・お名前・会社名・カテゴリーが入る・チェックを外せば入れない');

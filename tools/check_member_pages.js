// 定例会スライド（前半）の新メンバー・更新メンバーのページ（meeting_pages_srv.js の applyMemberPages_）を、作り物のページで確かめる。
// 雛形の実物は使わない。
//
//   node tools/check_member_pages.js
//
// 確かめること
//   ・並びは「新メンバー → 倫理規定 → 更新メンバー → 倫理規定」。雛形が「更新メンバー → 新メンバー → 倫理規定」の並び
//     （Activeチャプターの雛形）でも、新メンバーを先にする。はじめから新メンバーが先の雛形はそのまま
//     まとまりごとに倫理規定がある雛形・新メンバーのページが2か所にある雛形・見出しや「一言」のページがある雛形も
//   ・1人1枚（人数ぶんページを増やす・書いた順）。いないまとまりのページと、そのうしろの倫理規定は非表示
//   ・表示の倫理規定は「人がいるまとまり」の数だけで、続けて2枚出ない（倫理規定のページにページ番号があって、文字が少し違っても）
//   ・PowerPoint のセクションがある雛形は、セクションも入れ替えた並びに合わせる（syncSections_）
//   ・1枚に両方の欄がある雛形（公式ファイルの「新規および更新メンバー」）は、そのページのうしろに倫理規定1枚
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
// 作り物の名簿（会社名・カテゴリー）。写真は無い
const ROSTER = [
  { name: '見本 新一', company: '新一商事', title: '税理士' }, { name: '見本 新二', company: '新二企画', title: '司法書士' },
  { name: '見本 更一', company: '更一工務店', title: '工務店' }, { name: '見本 更二', company: '更二保険', title: '生命保険' },
  { name: '見本 更三', company: '更三デザイン', title: 'Web制作' },
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
                 'role_input_srv.js', 'role_intro_srv.js', 'meeting_pages_srv.js']) {
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
const para = (t, sz) => `<a:p><a:r><a:rPr lang="ja-JP" sz="${sz}"/><a:t>${t}</a:t></a:r></a:p>`;
const sp = (id, x, y, w, h, lines, sz, title) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Text ${id}"/><p:cNvSpPr txBox="1"/>`
  + `${title ? '<p:nvPr><p:ph type="title"/></p:nvPr>' : '<p:nvPr/>'}</p:nvSpPr><p:spPr>${xf(x, y, w, h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`
  + `<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>${lines.map((t) => para(t, sz)).join('')}</p:txBody></p:sp>`;
const pic = (id, x, y, w, h, rid) => `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="図 ${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>`
  + `<p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>`
  + `<p:spPr>${xf(x, y, w, h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
const slideXml = (inner) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld><p:spTree>`
  + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'
  + inner + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
// 1人1枚のページ：見出し・写真の枠（と見本の文字）・その下にお名前（大きな字）・その下に会社名と【カテゴリー】。更新は「1年更新」も
const person = (head, renew) => sp(2, 0, 20, 960, 70, [head], 4000, true)
  + pic(5, 330, 95, 300, 200, 'rId3') + sp(6, 340, 170, 280, 40, ['写真のサイズを枠に合わせてトリミング'], 1200)
  + sp(3, 240, 310, 480, 70, ['見本 太郎'], 4400) + sp(4, 240, 390, 480, 70, ['見本株式会社', '【見本カテゴリー】'], 2000)
  + (renew ? sp(7, 760, 120, 180, 50, ['1年更新'], 2400) : '');
const SLIDES = {
  cover: sp(2, 100, 200, 700, 80, ['ようこそ　BNI 見本チャプター'], 4000),
  renew: person('更新メンバー', true),
  neu: person('新メンバー', false),
  ethics: sp(2, 0, 20, 960, 70, ['BNI 倫理規定'], 4000, true) + sp(3, 40, 120, 880, 360, ['1. 私は、見本の倫理規定を守ります。'], 2000),
  both: sp(2, 0, 20, 960, 70, ['新規および更新メンバー'], 4000, true)
    + sp(10, 60, 110, 400, 50, ['新メンバー'], 2800) + sp(11, 500, 110, 400, 50, ['更新メンバー'], 2800)
    + [0, 1, 2].map((i) => sp(20 + i, 60, 170 + i * 60, 400, 50, ['氏名'], 2400) + sp(30 + i, 500, 170 + i * 60, 400, 50, ['氏名'], 2400)).join(''),
  weekly: sp(2, 0, 200, 960, 80, ['ウィークリープレゼンテーション'], 5400),
  headRenew: sp(2, 0, 200, 960, 80, ['更新メンバー紹介'], 5400, true),
  headNew: sp(2, 0, 200, 960, 80, ['新メンバー紹介'], 5400, true),
  newNote: sp(2, 0, 20, 960, 70, ['新メンバーからの一言'], 4000, true) + sp(3, 40, 120, 880, 200, ['入会のきっかけと、これからのことを話します。'], 2000),
  headRenewShiki: sp(2, 0, 200, 960, 80, ['更新式'], 5400, true),
  headNewShiki: sp(2, 0, 200, 960, 80, ['新入会メンバー紹介'], 5400, true),
  // 「倫理規定」の言葉が入った一般規定のページ（倫理規定のページではない）
  policy: sp(2, 0, 20, 960, 70, ['一般規定'], 4000, true) + sp(3, 40, 120, 880, 200, ['メンバー規定、倫理規定、BNIコアバリューを守ります。'], 2000),
};
// ページ番号の枠（PowerPoint は番号の文字をページに持つので、置いた場所で文字が変わる）
const numberPh = (n) => '<p:sp><p:nvSpPr><p:cNvPr id="90" name="スライド番号"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
  + `<p:nvPr><p:ph type="sldNum" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr><p:spPr>${xf(880, 500, 60, 30)}</p:spPr>`
  + '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{8A5F1E6B-0C7D-4B7A-9E2B-3F4A5B6C7D8E}" type="slidenum">'
  + `<a:rPr lang="ja-JP" sz="1200"/><a:t>${n}</a:t></a:fld><a:endParaRPr lang="ja-JP" sz="1200"/></a:p></p:txBody></p:sp>`;
// opts.numbers … どのページにもページ番号。opts.sections … PowerPoint のセクション [[名前, [ページの位置（0から）…]], …]
function makeParts(order, opts) {
  opts = opts || {};
  const parts = {};
  const put = (p, s) => { parts[p] = blob(s, 'application/xml', p); };
  put('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + order.map((k, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('') + '</Types>');
  const secs = (opts.sections || []).map(([name, idx], k) => `<p14:section name="${name}" id="{0000000${k}-0000-4000-8000-000000000000}">`
    + (idx.length ? `<p14:sldIdLst>${idx.map((i) => `<p14:sldId id="${256 + i}"/>`).join('')}</p14:sldIdLst>` : '<p14:sldIdLst/>') + '</p14:section>').join('');
  put('ppt/presentation.xml', `<?xml version="1.0"?><p:presentation ${NS}><p:sldIdLst>`
    + order.map((k, i) => `<p:sldId id="${256 + i}" r:id="rId${100 + i}"/>`).join('')
    + '</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/>'
    + (secs ? '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">'
              + `<p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">${secs}</p14:sectionLst></p:ext></p:extLst>` : '')
    + '</p:presentation>');
  put('ppt/_rels/presentation.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + order.map((k, i) => `<Relationship Id="rId${100 + i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`).join('')
    + '</Relationships>');
  order.forEach((k, i) => {
    put(`ppt/slides/slide${i + 1}.xml`, slideXml(SLIDES[k] + (opts.numbers ? numberPh(i + 1) : '')));
    put(`ppt/slides/_rels/slide${i + 1}.xml.rels`, '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${i + 1}.png"/></Relationships>`);
  });
  return parts;
}
// 作ったあとの表示のページ：[種類, お名前]（種類 … 新・更新・倫理・両方・ほか）
function shown(parts) {
  return F.slideOrder_(parts).map((p) => F.xmlOf_(parts, p)).filter((x) => !/<p:sld\b[^>]*\sshow="0"/.test(x)).map((x) => {
    const t = F.slideText_(x), kind = F.fpMemberKind_(x);
    const name = (t.match(/見本 [新更][一二三]/) || [''])[0];
    if (/一般規定/.test(t)) return '一般規定';
    if (/紹介$/.test(t)) return kind === 'renew' ? '更新の見出し' : '新の見出し';
    if (/^更新式$/.test(t)) return '更新の見出し';
    if (/一言/.test(t)) return '新の一言';
    return kind === 'new' ? '新:' + name : kind === 'renew' ? '更新:' + name : kind === 'both' ? '両方' : /倫理規定/.test(t) ? '倫理' : /ウィークリー/.test(t) ? '週' : '表紙';
  });
}
// presentation.xml のセクション [{ name, ids }] と、スライドの並びの番号
function sectionsOf(parts) {
  return (F.xmlOf_(parts, 'ppt/presentation.xml').match(/<p14:section\b[^>]*?(?:\/>|>[\s\S]*?<\/p14:section>)/g) || []).map((x) => ({
    name: (x.match(/name="([^"]*)"/) || [])[1], ids: (x.match(/<p14:sldId id="\d+"/g) || []).map((y) => y.match(/id="(\d+)"/)[1]) }));
}
const sldIds = (parts) => (F.xmlOf_(parts, 'ppt/presentation.xml').match(/<p:sldId id="\d+"/g) || []).map((y) => y.match(/id="(\d+)"/)[1]);
// 開きタグ・閉じタグが対応しているか（簡単な確かめ）
function balanced(xml) {
  const st = [], re = /<(\/?)([\w:]+)[^>]*?(\/?)>/g;
  let m;
  while ((m = re.exec(xml.replace(/<\?[^>]*\?>/g, ''))) !== null) {
    if (m[3]) continue;
    if (!m[1]) st.push(m[2]);
    else if (st.pop() !== m[2]) return false;
  }
  return st.length === 0;
}
const people = (names, years) => names.map((n) => ({ name: n, raw: n, years: years || 1 }));
const LISTS = {
  両方: { newMembers: people(['見本 新一', '見本 新二']), renewMembers: people(['見本 更一', '見本 更二', '見本 更三'], 2) },
  新だけ: { newMembers: people(['見本 新一']), renewMembers: [] },
  更新だけ: { newMembers: [], renewMembers: people(['見本 更一']) },
  だれもいない: { newMembers: [], renewMembers: [] },
};
const WANT = {
  両方: ['表紙', '新:見本 新一', '新:見本 新二', '倫理', '更新:見本 更一', '更新:見本 更二', '更新:見本 更三', '倫理', '週'],
  新だけ: ['表紙', '新:見本 新一', '倫理', '週'],
  更新だけ: ['表紙', '更新:見本 更一', '倫理', '週'],
  だれもいない: ['表紙', '週'],
};
const LAYOUTS = {
  'Activeチャプターの雛形（更新 → 新 → 倫理規定）': ['cover', 'renew', 'neu', 'ethics', 'weekly'],
  '新メンバーが先の雛形（新 → 更新 → 倫理規定）': ['cover', 'neu', 'renew', 'ethics', 'weekly'],
  'まとまりごとに倫理規定がある雛形（更新 → 倫理規定 → 新 → 倫理規定）': ['cover', 'renew', 'ethics', 'neu', 'ethics', 'weekly'],
  '新メンバーのページが2か所にある雛形（新 → 更新 → 新 → 倫理規定）': ['cover', 'neu', 'renew', 'neu', 'ethics', 'weekly'],
  '倫理規定がまとまりの前とあいだにある雛形（倫理規定 → 更新 → 倫理規定 → 新）': ['cover', 'ethics', 'renew', 'ethics', 'neu', 'weekly'],
  '倫理規定の写しが2枚ある雛形（更新 → 新 → 倫理規定 → 倫理規定）': ['cover', 'renew', 'neu', 'ethics', 'ethics', 'weekly'],
};
// ページ番号が無い雛形と、どのページにもページ番号がある雛形（倫理規定のページどうしの文字が番号だけ違う）
[false, true].forEach((numbers) => Object.entries(LAYOUTS).forEach(([ln0, order]) => Object.entries(LISTS).forEach(([cn, lists]) => {
  const ln = ln0 + (numbers ? '・ページ番号あり' : '');
  const parts = makeParts(order, { numbers });
  const res = F.applyMemberPages_(parts, lists, { by: {}, seq: 0 });
  const got = shown(parts);
  ck(J(got) === J(WANT[cn]), `${ln}・${cn}：表示のページ ${J(got)}（${J(WANT[cn])} のはず）`);
  // 知らせの倫理規定の並びも、新メンバーが先
  if (cn === '両方') ck(/倫理規定のページを、新メンバー・更新メンバーのあとに表示しました/.test(res.message), `${ln}：知らせ ${res.message}`);
  if (cn === 'だれもいない') ck(/倫理規定のページも非表示にしました/.test(res.message), `${ln}：だれもいない日の知らせ ${res.message}`);
  // スライドの番号（p:sldId の id）が重ならない
  const ids = sldIds(parts);
  ck(new Set(ids).size === ids.length, `${ln}・${cn}：スライドの番号が重なる`);
})));
// 見出し（「新メンバー紹介」「更新メンバー紹介」）・「新メンバーからの一言」のページがある雛形：そのまとまりと一緒に動く
{
  const cases = {
    '見出しがある雛形（更新の見出し → 更新 → 新の見出し → 新 → 倫理規定）': [['cover', 'headRenew', 'renew', 'headNew', 'neu', 'ethics', 'weekly'],
      ['表紙', '新の見出し', '新:見本 新一', '新:見本 新二', '倫理', '更新の見出し', '更新:見本 更一', '更新:見本 更二', '更新:見本 更三', '倫理', '週']],
    '「新メンバーからの一言」がある雛形（更新 → 新 → 一言 → 倫理規定）': [['cover', 'renew', 'neu', 'newNote', 'ethics', 'weekly'],
      ['表紙', '新:見本 新一', '新:見本 新二', '新の一言', '倫理', '更新:見本 更一', '更新:見本 更二', '更新:見本 更三', '倫理', '週']],
    '「新メンバーからの一言」がある、新メンバーが先の雛形（新 → 一言 → 更新 → 倫理規定）': [['cover', 'neu', 'newNote', 'renew', 'ethics', 'weekly'],
      ['表紙', '新:見本 新一', '新:見本 新二', '新の一言', '倫理', '更新:見本 更一', '更新:見本 更二', '更新:見本 更三', '倫理', '週']],
    '「更新式」「新入会メンバー紹介」の見出しがある雛形': [['cover', 'headRenewShiki', 'renew', 'headNewShiki', 'neu', 'ethics', 'weekly'],
      ['表紙', '新の見出し', '新:見本 新一', '新:見本 新二', '倫理', '更新の見出し', '更新:見本 更一', '更新:見本 更二', '更新:見本 更三', '倫理', '週']],
    // 一般規定のページに「倫理規定」の言葉があっても、写すのは倫理規定のページ。一般規定のページは隠さない
    '「倫理規定」の言葉が入った一般規定のページがある雛形（更新 → 新 → 倫理規定 → 一般規定）': [['cover', 'renew', 'neu', 'ethics', 'weekly', 'policy'],
      ['表紙', '新:見本 新一', '新:見本 新二', '倫理', '更新:見本 更一', '更新:見本 更二', '更新:見本 更三', '倫理', '週', '一般規定']],
    '「倫理規定」の言葉が入った一般規定のページが前にある雛形（一般規定 → 更新 → 新 → 倫理規定）': [['cover', 'policy', 'renew', 'neu', 'ethics', 'weekly'],
      ['表紙', '一般規定', '新:見本 新一', '新:見本 新二', '倫理', '更新:見本 更一', '更新:見本 更二', '更新:見本 更三', '倫理', '週']],
  };
  Object.entries(cases).forEach(([ln, [order, want]]) => {
    const parts = makeParts(order);
    F.applyMemberPages_(parts, LISTS['両方'], { by: {}, seq: 0 });
    ck(J(shown(parts)) === J(want), `${ln}：表示のページ ${J(shown(parts))}（${J(want)} のはず）`);
  });
}
// PowerPoint のセクションがある雛形：入れ替えたら、セクションも並びに合わせる（どのページもどこか1つのセクションに、続けて・並びの順に）
{
  const order = LAYOUTS['Activeチャプターの雛形（更新 → 新 → 倫理規定）'];
  const SECTIONS = {
    'まとまりごと': [['開会', [0]], ['更新メンバー', [1]], ['新メンバー', [2, 3]], ['ウィークリー', [4]]],
    'ひとまとめ': [['前半', [0, 1, 2, 3, 4]]],
  };
  Object.entries(SECTIONS).forEach(([sn, sections]) => Object.entries(LISTS).forEach(([cn, lists]) => {
    const parts = makeParts(order, { sections });
    F.applyMemberPages_(parts, lists, { by: {}, seq: 0 });
    const secs = sectionsOf(parts), flat = secs.reduce((a, s) => a.concat(s.ids), []);
    ck(J(flat) === J(sldIds(parts)), `セクション（${sn}・${cn}）：セクションのページ ${J(flat)} が並び ${J(sldIds(parts))} と合わない`);
    ck(J(secs.map((s) => s.name).sort()) === J(sections.map((s) => s[0]).sort()), `セクション（${sn}・${cn}）：セクションの名前 ${J(secs.map((s) => s.name))}`);
    ck(J(shown(parts)) === J(WANT[cn]), `セクション（${sn}・${cn}）：表示のページ ${J(shown(parts))}`);
    ck(balanced(F.xmlOf_(parts, 'ppt/presentation.xml')), `セクション（${sn}・${cn}）：presentation.xml のタグが対応していない`);
    if (sn === 'まとまりごと' && cn === '両方') {
      ck(J(secs.map((s) => s.name)) === J(['開会', '新メンバー', '更新メンバー', 'ウィークリー']), 'セクション：並び ' + J(secs.map((s) => s.name)));
      const idOf = {};
      F.slideEntries_(parts).forEach((e) => { idOf[e.path] = String(e.id); });
      const secOf = (p) => (secs.find((s) => s.ids.includes(idOf[p])) || {}).name;
      F.slideOrder_(parts).forEach((p) => {
        const k = F.fpMemberKind_(F.xmlOf_(parts, p));
        if (k === 'new' || k === 'renew') ck(secOf(p) === (k === 'new' ? '新メンバー' : '更新メンバー'), `セクション：${k} のページが「${secOf(p)}」にある`);
      });
    }
  }));
  // 入れ替えの無い雛形（新メンバーが先）は、セクションに手を付けない
  const sec = (parts) => F.xmlOf_(parts, 'ppt/presentation.xml').match(/<p14:sectionLst[\s\S]*<\/p14:sectionLst>/)[0];
  const nf = makeParts(LAYOUTS['新メンバーが先の雛形（新 → 更新 → 倫理規定）'], { sections: [['前半', [0, 1, 2, 3, 4]]] }), before = sec(nf);
  F.applyMemberPages_(nf, LISTS['両方'], { by: {}, seq: 0 });
  ck(sec(nf) === before, '入れ替えの無い雛形のセクションが変わった');
}
// syncSections_ そのもの：離れたページ・どこにも無いページ・中身の無いセクション・並びに無い番号・セクションの無いファイル
{
  const deck = (ids, secXml) => {
    const parts = makeParts(['cover']);
    F.putXml_(parts, 'ppt/presentation.xml', `<?xml version="1.0"?><p:presentation ${NS}><p:sldIdLst>`
      + ids.map((id, i) => `<p:sldId id="${id}" r:id="rId${100 + i}"/>`).join('') + '</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/>'
      + (secXml ? '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">'
                  + `<p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">${secXml}</p14:sectionLst></p:ext></p:extLst>` : '')
      + '</p:presentation>');
    return parts;
  };
  const sx = (name, ids) => `<p14:section name="${name}" id="{${name}}">` + (ids ? `<p14:sldIdLst>${ids.map((id) => `<p14:sldId id="${id}"/>`).join('')}</p14:sldIdLst>` : '<p14:sldIdLst/>') + '</p14:section>';
  let parts = deck([256, 259, 257, 258, 260, 261], sx('A', [256, 257, 258]) + sx('B', [259]) + sx('C', null) + sx('D', [260, 999]) + '<p14:section name="E" id="{E}"/>');
  ck(F.syncSections_(parts) === true, 'syncSections_：セクションがあるのに false');
  ck(J(sectionsOf(parts)) === J([{ name: 'A', ids: ['256'] }, { name: 'B', ids: ['259', '257', '258'] }, { name: 'C', ids: [] },
                                { name: 'D', ids: ['260', '261'] }, { name: 'E', ids: [] }]), 'syncSections_：' + J(sectionsOf(parts)));
  ck(balanced(F.xmlOf_(parts, 'ppt/presentation.xml')), 'syncSections_：タグが対応していない');
  parts = deck([300, 256], sx('A', [256]) + sx('B', []));
  F.syncSections_(parts);
  ck(J(sectionsOf(parts)) === J([{ name: 'A', ids: ['300', '256'] }, { name: 'B', ids: [] }]), 'syncSections_（先頭がどこにも無い）：' + J(sectionsOf(parts)));
  parts = deck([256, 257], sx('後', [257]) + sx('前', [256]));
  F.syncSections_(parts);
  ck(J(sectionsOf(parts)) === J([{ name: '前', ids: ['256'] }, { name: '後', ids: ['257'] }]), 'syncSections_（セクションの並び）：' + J(sectionsOf(parts)));
  parts = deck([256, 257], '');
  const x0 = F.xmlOf_(parts, 'ppt/presentation.xml');
  ck(F.syncSections_(parts) === false && F.xmlOf_(parts, 'ppt/presentation.xml') === x0, 'syncSections_：セクションの無いファイルが変わった');
}
// 1人1枚のページの中身：お名前・会社名・【カテゴリー】・更新の年数。写真の無い方は写真の枠を隠し、見本の文字を消す
{
  const parts = makeParts(LAYOUTS['Activeチャプターの雛形（更新 → 新 → 倫理規定）']);
  const res = F.applyMemberPages_(parts, LISTS['両方'], { by: {}, seq: 0 });
  const vis = F.slideOrder_(parts).map((p) => F.xmlOf_(parts, p)).filter((x) => !/<p:sld\b[^>]*\sshow="0"/.test(x));
  const t1 = F.slideText_(vis[1]), t4 = F.slideText_(vis[4]);
  ck(/見本 新一/.test(t1) && /新一商事/.test(t1) && /【税理士】/.test(t1) && !/トリミング/.test(t1) && /<p:cNvPr id="5"[^>]*hidden="1"/.test(vis[1]),
     '新メンバーのページの中身: ' + t1);
  ck(/見本 更一/.test(t4) && /2年更新/.test(t4) && /【工務店】/.test(t4), '更新メンバーのページの中身: ' + t4);
  ck(res.noPhoto.length === 5, '写真の無い方の知らせ: ' + J(res.noPhoto));
}
// 1枚に両方の欄がある雛形（公式ファイル）：そのページのうしろに倫理規定1枚。欄には新メンバー・更新メンバー（年数）
{
  const parts = makeParts(['cover', 'both', 'ethics', 'weekly']);
  F.applyMemberPages_(parts, LISTS['両方'], { by: {}, seq: 0 });
  const t = F.slideText_(F.xmlOf_(parts, F.slideOrder_(parts)[1]));
  ck(J(shown(parts)) === J(['表紙', '両方', '倫理', '週']), '新規および更新メンバーの1枚: ' + J(shown(parts)));
  ck(/見本 新一/.test(t) && /見本 新二/.test(t) && /見本 更一（2年）/.test(t), '新規および更新メンバーの欄: ' + t);
}

console.log(`新メンバー・更新メンバーのページ: 検査 ${checks} 件`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: 新メンバー → 倫理規定 → 更新メンバー → 倫理規定（雛形の並びによらず・ページ番号があっても）・見出しと一言は一緒に動く・'
  + '1人1枚・いないまとまりは非表示・PowerPoint のセクションも合わせる・新規および更新メンバーの1枚');

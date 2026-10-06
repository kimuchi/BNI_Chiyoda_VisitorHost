// 役職のメンバー紹介（role_intro_srv.js）を、作り物のページで確かめる。雛形の実物は使わない。
//
//   node tools/check_role_intro.js
//
// 確かめること
//   ・ふりがな → ローマ字（名・姓の順・長い音・っ・ん）
//   ・枠の見分け方：リーダーシップチームの表（見出しの下にお名前・表の下に写真）／コーディネーターの表（同じ行）／
//     チームの表（見出しの下に1人ずつ・「氏名, 氏名」の1マス）／1人のページ（題が役職、お名前・ローマ字・会社名・（カテゴリー））／
//     メンバーシップ委員会のページ（写真の下にカテゴリーとお名前のグループ）／各サポートチーム（見出し → 写真 → グループ）／
//     題がチームのお名前の表（ビジターホストチーム）。報告のページ・ウィークリープレゼンより後ろのページは見分けない
//   ・入れ方：替わった方の枠だけ（同じ方の枠は写真もそのまま）／居ない方の枠は空にして写真の枠を外す／写真の無い方も枠を外して知らせる／
//     担当者の居ない役職の1人のページは非表示／カテゴリーの（）の中の（）は「」／長いお名前は1行に収まるよう小さく／
//     チームの枠より人数が多いときは知らせる
//   ・差し込み口の雛形：入れる／画面で外したとき（data が null）は差し込み口を空にするだけで写真の枠は残す
//   ・写真の切り抜き：枠より縦長の写真は上をそろえて下だけ切る（頭が切れない）。横長は左右から均等に
//   ・役職もチームも未登録なら、差し込み口の無いページには触らない
//   ・ネットワーキング学習コーナー：「担当：」は20pt・お名前は40ptの太字（狭い枠ではお名前だけ小さく）、自動縮小をやめ、枠の高さを2行ぶんに

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const zlib = require('zlib');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }
const J = (x) => JSON.stringify(x);

// --- 小さなPNG（写真の代わり）---
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
const signed = (buf) => Array.from(buf).map((b) => (b > 127 ? b - 256 : b));
function blob(content, type, name) {
  let nm = name || '';
  const buf = Buffer.isBuffer(content) ? content : Buffer.from(String(content), 'utf8');
  return { _buf: buf, getDataAsString: () => buf.toString('utf8'), getBytes: () => signed(buf), getContentType: () => type,
           setName(n) { nm = n; return this; }, getName: () => nm };
}

// --- 作り物の名簿（架空の方）と写真 ---
const MEMBERS = [
  { name: '見本　一郎', kana: 'みほん いちろう', company: '見本会計', title: '税理士（相続）' },
  { name: '見本　二郎', kana: 'みほん じろう', company: '二郎商事', title: '司法書士' },
  { name: '前任　三郎', kana: '', company: '三郎工務店', title: '工務店' },
  { name: '見本　四郎', kana: '', company: '四郎企画', title: 'イベント企画' },
  { name: '見本　五郎', kana: '', company: '', title: '' },                     // 写真が無い
  { name: '見本　六郎', kana: '', company: '六郎デザイン', title: 'Web制作' },
  { name: '見本　七子', kana: '', company: '七子保険', title: '生命保険' },
  { name: '見本　長い名前の方', kana: '', company: '', title: '行政書士（建設業許可・外国人ビザ・相続手続き）' },
].concat(Array.from({ length: 10 }, (_, i) => ({ name: '案内　' + '甲乙丙丁戊己庚辛壬癸'[i] + '子', kana: '', company: '', title: '' })));
const PHOTO_OF = {};
MEMBERS.forEach((m, i) => { if (m.name !== '見本　五郎') PHOTO_OF[m.name.replace(/[\s　]/g, '')] = 'photo' + i; });
const PHOTO_BLOB = {};
Object.values(PHOTO_OF).forEach((id, i) => { PHOTO_BLOB[id] = png(30 + i, 40, [i * 20 % 256, 100, 200]); });

const sb = {
  console,
  Utilities: { newBlob: (c, t, n) => blob(c, t, n) },
  DriveApp: { getFileById: (id) => ({ getBlob: () => blob(PHOTO_BLOB[id], 'image/png', id) }) },
  normName_: (s) => String(s == null ? '' : s).normalize('NFKC').replace(/[\s　]/g, ''),
  findPhotoIdForName_: (n) => PHOTO_OF[String(n || '').replace(/[\s　]/g, '')] || '',
};
vm.createContext(sb);
for (const f of ['ooxml.js', 'chapter_srv.js', 'member_presen_srv.js', 'referral_srv.js', 'splice_srv.js', 'meeting_slides_srv.js',
                 'role_input_srv.js', 'role_intro_srv.js']) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), sb, { filename: f });
}
const F = sb;

// ===== 1. ローマ字 =====
[['みほん いちろう', 'ICHIRO MIHON'], ['さとう しゅうへい', 'SHUHEI SATO'], ['おおの りょう', 'RYO ONO'],
 ['きっかわ ちえ', 'CHIE KIKKAWA'], ['なんば けんた', 'KENTA NAMBA'], ['ミホン タロウ', 'TARO MIHON'],
 ['まっちゃ', 'MATCHA'], ['', ''], ['見本 一郎', '']].forEach(([k, want]) => {
  ck(F.riRomaji_(k) === want, `ローマ字「${k}」→「${F.riRomaji_(k)}」（${want} のはず）`);
});

// ===== 2. 作り物のページ =====
const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
  + 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const E = 12700;
const xf = (x, y, w, h) => `<a:xfrm><a:off x="${Math.round(x * E)}" y="${Math.round(y * E)}"/><a:ext cx="${Math.round(w * E)}" cy="${Math.round(h * E)}"/></a:xfrm>`;
const para = (t, sz) => `<a:p><a:r><a:rPr lang="ja-JP" sz="${sz}"/><a:t>${t}</a:t></a:r></a:p>`;
function sp(id, x, y, w, h, lines, sz, title) {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Text ${id}"/><p:cNvSpPr txBox="1"/>${title ? '<p:nvPr><p:ph type="title"/></p:nvPr>' : '<p:nvPr/>'}</p:nvSpPr>`
    + `<p:spPr>${xf(x, y, w, h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`
    + `<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>${lines.map((t) => para(t, sz)).join('')}</p:txBody></p:sp>`;
}
function pic(id, x, y, w, h, rid) {
  return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="図 ${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>`
    + `<p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>`
    + `<p:spPr>${xf(x, y, w, h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
}
function tbl(id, x, y, cols, rows) {
  const tc = (c) => {
    const o = typeof c === 'string' ? { t: c } : c;
    return `<a:tc${o.span ? ` gridSpan="${o.span}"` : ''}${o.hm ? ' hMerge="1"' : ''}><a:txBody><a:bodyPr/><a:lstStyle/>${para(o.t || '', 1800)}</a:txBody><a:tcPr/></a:tc>`;
  };
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>`
    + `<p:xfrm><a:off x="${x * E}" y="${y * E}"/><a:ext cx="${cols.reduce((a, b) => a + b, 0) * E}" cy="${rows.length * 40 * E}"/></p:xfrm>`
    + `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/>`
    + `<a:tblGrid>${cols.map((w) => `<a:gridCol w="${w * E}"/>`).join('')}</a:tblGrid>`
    + rows.map((r) => `<a:tr h="${40 * E}">${r.map(tc).join('')}</a:tr>`).join('') + '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>';
}
// グループ（子の座標は 0.8 倍に縮めて置く。スライド上の位置に直せているかを確かめるため）
function grp(id, x, y, w, h, kids) {
  return `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="グループ ${id}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`
    + `<p:grpSpPr><a:xfrm><a:off x="${x * E}" y="${y * E}"/><a:ext cx="${w * E}" cy="${h * E}"/>`
    + `<a:chOff x="0" y="0"/><a:chExt cx="${Math.round(w / 0.8 * E)}" cy="${Math.round(h / 0.8 * E)}"/></a:xfrm></p:grpSpPr>${kids}</p:grpSp>`;
}
// 1人ぶんのまとまり（グループの中の座標：幅 w/0.8 ・上にカテゴリー、下にお名前）
const unit = (id, x, y, w, cat, name, nameSz) => grp(id, x, y, w, 60,
  sp(id + 1, 0, 0, w / 0.8, 30, [cat], 1100) + sp(id + 2, 0, 35, w / 0.8, 40, [name], nameSz || 1400));
const slideXml = (inner) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld><p:spTree>`
  + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'
  + inner + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';

const SLIDES = {
  cover: sp(2, 100, 200, 700, 80, ['ようこそ　BNI 見本チャプター'], 4000),
  // リーダーシップチーム（見出しの下にお名前・表の下に写真）。書記兼会計は替わらない
  leader: sp(9, 0, 21, 960, 70, ['リーダーシップチーム'], 4000)
    + tbl(4, 327, 97, [304], [['チャプタープレジデント'], ['前任　一郎'], ['']]) + tbl(5, 22, 97, [304], [['バイスプレジデント'], ['前任　二郎'], ['']])
    + tbl(6, 632, 97, [304], [['書記兼会計'], ['前任　三郎'], ['']])
    + pic(3, 358, 211, 222, 272, 'rId3') + pic(8, 63, 211, 222, 270, 'rId4') + pic(7, 652, 211, 222, 271, 'rId5'),
  // 1人のページ（プレジデント）
  single: sp(3, 20, 16, 459, 93, ['プレジデント', 'President'], 4000, true)
    + sp(6, 0, 191, 475, 271, ['前任　一郎', 'ICHIRO ZENNIN', '前任商事', '（旧カテゴリー）'], 6600)
    + sp(4, 76, 144, 297, 36, ['チャプターミーティングの議長'], 2800) + pic(7, 506, 88, 353, 433, 'rId3'),
  // 1人のページ（メンターコーディネーター。担当者が居ない → 非表示）
  mentor: sp(3, 0, 56, 462, 93, ['メンターコーディネーター', 'Mentor Coordinator'], 4000, true)
    + sp(6, 38, 178, 414, 239, ['前任　四郎', 'SHIRO ZENNIN', '前任企画', '（旧）'], 6500) + pic(2, 489, 82, 384, 403, 'rId3'),
  // メンバーシップ委員会（写真の下にカテゴリーとお名前）。4枠に2名
  grid: sp(90, 0, 36, 960, 60, ['メンバーシップ委員会'], 4400)
    + [28, 263, 505, 737].map((x, i) => pic(2 + i, x, 143, 194, 194, 'rId' + (3 + i))).join('')
    + [22, 262, 503, 737].map((x, i) => unit(50 + i * 4, x, 367, 200, '旧カテゴリー' + i, '前任　' + '甲乙丙丁'[i] + '雄', 3200)).join(''),
  // ビジターホストチーム（題がチーム・見出しの無いお名前の表 3×3）
  vhTable: sp(17, 94, 32, 753, 83, ['ビジターホストチーム'], 4000, true)
    + tbl(20, 103, 21, [250, 251, 251], [[{ t: '', span: 2 }, { t: '', hm: true }, ''], ['前任　甲', '前任　乙', '前任　丙'],
      ['前任　丁', '前任　戊', '前任　己'], ['前任　庚', '', '']]),
  // 各サポートチーム（見出し → 写真 → カテゴリーとお名前）。BCP委員は居ない（枠を空にして写真を外す）
  labels: sp(2, 0, 15, 960, 64, ['各サポートチーム'], 4000)
    + sp(8, 311, 90, 139, 41, ['WEBマスター'], 1400) + sp(15, 706, 90, 139, 41, ['BCP委員長'], 1400)
    + pic(14, 328, 137, 106, 130, 'rId3') + pic(17, 723, 139, 106, 130, 'rId4')
    + unit(70, 316, 267, 140, 'オフィスプロデュース', '前任　庚') + unit(74, 707, 267, 140, '鍼灸師', '前任　辛'),
  // 公式ファイルの作り（メンバーシップ委員会・コーディネーター・ビジターホストの表）
  support: sp(2, 0, 15, 960, 64, ['サポートチーム'], 4000)
    + tbl(7, 75, 93, [375], [['メンバーシップ委員会'], ['氏名'], ['氏名'], ['氏名'], ['氏名']])
    + tbl(8, 500, 93, [187, 187], [[{ t: 'コーディネーター', span: 2 }, { t: '', hm: true }], ['エデュケーション', '氏名'], ['メンター', '氏名'], ['ビジターホスト', '氏名']])
    + tbl(5, 75, 377, [800], [['ビジターホスト'], ['氏名, 氏名, 氏名, 氏名'], ['']]),
  // 報告のページ（役職の名前が題にあっても見分けない）
  report: sp(2, 0, 20, 960, 60, ['バイスプレジデントによる報告'], 4000) + pic(3, 600, 100, 300, 300, 'rId3')
    + sp(4, 40, 120, 500, 200, ['月間リファーラル数の平均：○件'], 2400),
  weekly: sp(2, 0, 200, 960, 80, ['ウィークリープレゼンテーション'], 5400),
  // ウィークリープレゼンより後ろ（見分けない）
  after: sp(9, 0, 21, 960, 70, ['リーダーシップチーム'], 4000) + tbl(4, 327, 97, [304], [['チャプタープレジデント'], ['前任　一郎']]),
};
const ORDER = ['cover', 'leader', 'single', 'mentor', 'grid', 'vhTable', 'labels', 'support', 'report', 'weekly', 'after'];
function makeParts(slides) {
  const parts = {};
  const put = (p, s) => { parts[p] = blob(s, 'application/xml', p); };
  put('ppt/presentation.xml', `<?xml version="1.0"?><p:presentation ${NS}><p:sldIdLst>`
    + ORDER.map((k, i) => `<p:sldId id="${256 + i}" r:id="rId${100 + i}"/>`).join('')
    + '</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>');
  put('ppt/_rels/presentation.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + ORDER.map((k, i) => `<Relationship Id="rId${100 + i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`).join('')
    + '</Relationships>');
  ORDER.forEach((k, i) => {
    put(`ppt/slides/slide${i + 1}.xml`, slideXml(slides[k]));
    put(`ppt/slides/_rels/slide${i + 1}.xml.rels`, '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + [3, 4, 5, 6].map((n) => `<Relationship Id="rId${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/old${i}_${n}.png"/>`).join('')
      + '</Relationships>');
  });
  return parts;
}
const idx = (k) => ORDER.indexOf(k) + 1;
const sPath = (k) => `ppt/slides/slide${idx(k)}.xml`;
const W = 12192000, H = 6858000;

// ===== 3. 枠の見分け方 =====
const det = (k) => F.riDetectSlots_(slideXml(SLIDES[k]), W, H);
const brief = (d) => d.slots.map((s) => [s.kind, s.key, s.n || 0, s.photo ? 'P' : '-', s.fields.map((f) => f.field).join('+')].join(':'));
let d = det('leader');
ck(J(brief(d)) === J(['holder:president:0:P:氏名', 'holder:vice:0:P:氏名', 'holder:secretary:0:P:氏名'])
   && J(d.slots.map((s) => s.photo)) === J(['3', '8', '7']), 'リーダーシップチームの表: ' + J(brief(d)) + J(d.slots.map((s) => s.photo)));
d = det('single');
ck(d.single === 'president' && J(brief(d)) === J(['holder:president:0:P:氏名+ローマ字+会社名+カテゴリー']) && d.slots[0].photo === '7'
   && J(d.slots[0].fields[3].loc.wrap) === J(['（', '）']), '1人のページ: ' + J(d));
d = det('grid');
ck(J(brief(d)) === J([1, 2, 3, 4].map((n) => `team:membership:${n}:P:氏名+カテゴリー`)) && J(d.slots.map((s) => s.photo)) === J(['2', '3', '4', '5']),
   'メンバーシップ委員会のページ: ' + J(brief(d)) + J(d.slots.map((s) => s.photo)));
d = det('vhTable');
ck(d.slots.length === 9 && d.slots.every((s, i) => s.kind === 'team' && s.key === 'role:vhc' && s.n === i + 1)
   && d.slots[0].current === '前任　甲' && d.slots[8].current === '', 'ビジターホストチームの表: ' + J(brief(d)));
d = det('labels');
ck(J(brief(d)) === J(['holder:web:0:P:氏名+カテゴリー', 'holder:bcp:0:P:氏名+カテゴリー']) && J(d.slots.map((s) => s.photo)) === J(['14', '17']),
   '各サポートチーム: ' + J(brief(d)));
d = det('support');
ck(J(brief(d)) === J(['team:membership:1:-:氏名', 'team:membership:2:-:氏名', 'team:membership:3:-:氏名', 'team:membership:4:-:氏名',
                      'holder:ec:0:-:氏名', 'holder:mentor:0:-:氏名', 'holder:vhc:0:-:氏名', 'teamList:role:vhc:0:-:の一覧']),
   '公式ファイルの作り（サポートチーム）: ' + J(brief(d)));
ck(det('report').slots.length === 0 && det('cover').slots.length === 0 && det('weekly').slots.length === 0, '役職紹介でないページを見分けた');

// ===== 4. その期の方で入れる =====
const byName = {};
MEMBERS.forEach((m) => { byName[m.name] = m; });
const team = (key, names) => ({ key, name: key, members: names.map((n) => ({ name: n })) });
const holders = { president: '見本　一郎', vice: '見本　五郎', secretary: '前任　三郎', ec: '見本　四郎', vhc: '見本　二郎', mentor: '', web: '見本　六郎', bcp: '' };
const vhMembers = MEMBERS.slice(8).map((m) => m.name);                 // 10名（表の枠は9つ）
const data = F.riDataFrom_(24, '2026年10月〜2027年3月', holders, 24,
  [team('membership', ['見本　七子', '見本　長い名前の方']), team('role:vhc', vhMembers)], true, MEMBERS);
ck(data.holdersOk && data.teamsOk && data.roles.president.romaji === 'ICHIRO MIHON' && data.roles.president.category === '税理士（相続）'
   && data.roles.mentor === null, 'その期の方: ' + J(data.roles.president));
const parts = makeParts(SLIDES);
const before = {};
ORDER.forEach((k) => { before[k] = F.xmlOf_(parts, sPath(k)); });
const res = F.applyRoleIntro_(parts, data, { by: {}, seq: 0 });
const X = (k) => F.xmlOf_(parts, sPath(k));
const T = (k) => F.slideText_(X(k));
const picOf = (k, name) => (X(k).match(new RegExp('<p:pic>(?:(?!</p:pic>)[\\s\\S])*?name="' + name.replace(/[{}]/g, '\\$&') + '"[\\s\\S]*?</p:pic>')) || [''])[0];
const targetOf = (k, picXml) => {
  const rid = (picXml.match(/r:embed="(rId\d+)"/) || [])[1];
  const rels = F.xmlOf_(parts, `ppt/slides/_rels/slide${idx(k)}.xml.rels`);
  return ((rels.match(new RegExp('<Relationship\\b[^>]*\\bId="' + rid + '"[^>]*>')) || [''])[0].match(/Target="([^"]+)"/) || [])[1] || '';
};
const photoOf = (k, name) => { const t = targetOf(k, picOf(k, name)).replace('../', 'ppt/'); return parts[t] ? parts[t] : null; };
const isPhotoOf = (b, person) => !!b && Buffer.compare(b._buf, PHOTO_BLOB[PHOTO_OF[person.replace(/[\s　]/g, '')]]) === 0;

// リーダーシップチーム：プレジデント・バイスプレジデントは入れ替え、書記兼会計（同じ方）はそのまま
ck(T('leader').includes('見本　一郎') && T('leader').includes('見本　五郎') && T('leader').includes('前任　三郎') && !T('leader').includes('前任　一郎'),
   'リーダーシップチームのお名前: ' + T('leader'));
ck(isPhotoOf(photoOf('leader', '{{プレジデント写真}}'), '見本　一郎'), 'プレジデントの写真');
ck(!picOf('leader', '{{バイスプレジデント写真}}') && !/ id="8"/.test(X('leader')) && res.noPhoto.includes('見本　五郎'),
   '写真の無い方（見本　五郎）の枠を外して知らせる: ' + J(res.noPhoto));
ck(/<p:cNvPr id="7" name="図 7"\/>/.test(X('leader')) && X('leader').includes('../media/old1_5.png') === false && /rId5/.test(X('leader')),
   '同じ方（書記兼会計）の写真の枠はそのまま');
// 1人のページ
ck(['見本　一郎', 'ICHIRO MIHON', '見本会計', '（税理士「相続」）', 'チャプターミーティングの議長'].every((t) => T('single').includes(t))
   && !T('single').includes('前任') && isPhotoOf(photoOf('single', '{{プレジデント写真}}'), '見本　一郎'),
   'プレジデントのページ: ' + T('single'));
ck(/<p:sld\b[^>]*\sshow="0"/.test(X('mentor')) && res.hidden.includes('メンターコーディネーター') && T('mentor').includes('前任　四郎'),
   '担当者の居ない役職の1人のページを非表示: ' + J(res.hidden));
// メンバーシップ委員会：2名を入れ、残りの枠は空にして写真を外す。長いお名前は小さく
const g = X('grid');
ck(T('grid').includes('見本　七子') && T('grid').includes('生命保険') && T('grid').includes('見本　長い名前の方') && !T('grid').includes('前任')
   && isPhotoOf(photoOf('grid', '{{メンバーシップ委員会1写真}}'), '見本　七子') && !/ id="4" name/.test(g) && !/ id="5" name/.test(g),
   'メンバーシップ委員会のページ: ' + T('grid'));
const longSz = +((g.match(/<p:cNvPr id="56"[\s\S]*?sz="(\d+)"/) || [])[1] || 0);
ck(longSz > 0 && longSz < 3200, '長いお名前を1行に収まるよう小さくする: ' + longSz);
const catSz = +((g.match(/<p:cNvPr id="55"[\s\S]*?sz="(\d+)"/) || [])[1] || 0);
ck(catSz > 0 && catSz < 1100, '長いカテゴリーを枠の高さ（1行）に収まるよう小さくする: ' + catSz);
// 写真の切り抜き：枠より縦長の写真は下だけを切り、上はそろえる（上下から均等に切ると頭のてっぺんが切れていた）
{
  const pic1 = picOf('grid', '{{メンバーシップ委員会1写真}}'), img = photoOf('grid', '{{メンバーシップ委員会1写真}}');
  const pw = img._buf.readUInt32BE(16), ph = img._buf.readUInt32BE(20), want = Math.round((1 - (pw / ph) / (194 / 194)) * 100000);
  const rect = (pic1.match(/<a:srcRect\b[^>]*\/>/) || [''])[0];
  const got = ['l', 't', 'r', 'b'].map((k) => +((rect.match(new RegExp('\\s' + k + '="(-?\\d+)"')) || [])[1] || 0));
  ck(ph > pw && rect && J(got) === J([0, 0, 0, want]),
     `メンバーシップ委員会の縦長の写真は上をそろえて下だけ切る: ${rect || '切り抜きなし'}（l/t/r/b = 0/0/0/${want} のはず。写真 ${pw}×${ph}・枠は正方形）`);
  const rc = (c) => ['l', 't', 'r', 'b'].map((k) => Math.round(c[k]));
  ck(J(rc(F.coverCrop_(300, 400, 100, 100))) === J([0, 0, 0, 25000]) && J(rc(F.coverCrop_(300, 400, 300, 360))) === J([0, 0, 0, 10000])
     && J(rc(F.coverCrop_(400, 300, 100, 100))) === J([12500, 0, 12500, 0]) && J(rc(F.coverCrop_(300, 400, 600, 800))) === J([0, 0, 0, 0]),
     '切り抜きの量（縦長は下だけ・横長は左右均等・同じ形は切らない）: ' + J([rc(F.coverCrop_(300, 400, 100, 100)), rc(F.coverCrop_(400, 300, 100, 100))]));
}
// ビジターホストチームの表：9つの枠に上から、10人目は入らないので知らせる
ck(vhMembers.slice(0, 9).every((n) => T('vhTable').includes(n)) && !T('vhTable').includes(vhMembers[9]) && !T('vhTable').includes('前任')
   && res.overflow.some((t) => /ビジターホスト（10名のうち9名ぶんの枠）/.test(t)), 'ビジターホストチームの表: ' + J(res.overflow));
// 各サポートチーム
ck(T('labels').includes('見本　六郎') && T('labels').includes('Web制作') && !T('labels').includes('前任') && !T('labels').includes('鍼灸師')
   && isPhotoOf(photoOf('labels', '{{webマスター写真}}'), '見本　六郎') && !/ id="17" name/.test(X('labels')),
   '各サポートチーム（BCP委員は空・写真の枠を外す）: ' + T('labels'));
// 公式ファイルの作りの表
ck(T('support').includes('見本　七子') && T('support').includes('見本　四郎') && T('support').includes('見本　二郎')
   && T('support').includes(vhMembers.join('、')) && !T('support').includes('氏名'), 'サポートチームの表: ' + T('support').slice(0, 200));
// 触らないページ
ck(X('report') === before.report && X('cover') === before.cover && X('after') === before.after && X('weekly') === before.weekly,
   '役職紹介でないページ・ウィークリープレゼンより後ろのページが変わった');
ck(!ORDER.some((k) => /\{\{[^{}]+\}\}/.test(T(k))), '差し込み口が残っている');
ck(/24期/.test(res.message) && /メンターコーディネーター/.test(res.message) && /見本　五郎/.test(res.message) && /ビジターホスト/.test(res.message),
   '知らせ: ' + res.message);

// 同じ内容でもう一度：替わった方がいないので、どのページも変わらない
const parts2 = makeParts(SLIDES);
const same = F.riDataFrom_(24, '', { president: '前任　一郎', vice: '前任　二郎', secretary: '前任　三郎', ec: '', mentor: '', vhc: '' }, 24,
  [], false, MEMBERS);
const r2 = F.applyRoleIntro_(parts2, same, { by: {}, seq: 0 });
ck(F.xmlOf_(parts2, sPath('leader')) === slideXml(SLIDES.leader) && F.xmlOf_(parts2, sPath('single')) === slideXml(SLIDES.single)
   && F.xmlOf_(parts2, sPath('grid')) === slideXml(SLIDES.grid), '同じ方のままのページが変わった');
// 担当者が空の役職の「氏名」（公式ファイルの見本の文字）は空にする。チームは未登録なのでそのまま
ck(!/エデュケーション氏名/.test(F.slideText_(F.xmlOf_(parts2, sPath('support'))))
   && (F.slideText_(F.xmlOf_(parts2, sPath('support'))).match(/氏名/g) || []).length === 8,     // メンバーシップ委員会の4つと一覧の4つ
   '担当者が空の役職の見本の文字: ' + F.slideText_(F.xmlOf_(parts2, sPath('support'))));
ck(r2.hidden.includes('メンターコーディネーター'), '担当者の居ない1人のページ（チームが未登録でも）: ' + J(r2.hidden));

// 役職もチームも未登録：差し込み口の無いページには触らない
const parts3 = makeParts(SLIDES);
const none = F.riDataFrom_(24, '', {}, null, [], false, MEMBERS);
const r3 = F.applyRoleIntro_(parts3, none, { by: {}, seq: 0 });
ck(ORDER.every((k) => F.xmlOf_(parts3, sPath(k)) === slideXml(SLIDES[k])) && /未登録のため、テンプレートのまま/.test(r3.message),
   '未登録の期でページが変わった: ' + r3.message);

// ===== 5. 差し込み口の雛形（公式ファイルから作った雛形）=====
const tokXml = F.riTokenizePage_(slideXml(SLIDES.support), W, H);
ck(['{{メンバーシップ委員会1氏名}}', '{{メンバーシップ委員会4氏名}}', '{{エデュケーションコーディネーター氏名}}', '{{メンターコーディネーター氏名}}',
    '{{ビジターホストコーディネーター氏名}}', '{{ビジターホストの一覧}}'].every((t) => F.slideText_(tokXml).includes(t)),
   '差し込み口を入れる: ' + F.slideText_(tokXml));
const tokLeader = F.riTokenizePage_(slideXml(SLIDES.leader), W, H);
ck(/name="\{\{プレジデント写真\}\}"/.test(tokLeader) && F.slideText_(tokLeader).includes('{{書記兼会計氏名}}'), '写真の枠の名前');
const tokSingle = F.riTokenizePage_(slideXml(SLIDES.single), W, H);
ck(F.slideText_(tokSingle).includes('{{プレジデントカテゴリー（）}}') && F.slideText_(tokSingle).includes('{{プレジデントローマ字}}'),
   '1人のページの差し込み口: ' + F.slideText_(tokSingle));
// 差し込み口の入ったページに入れる。画面で外したとき（null）は、差し込み口を空にするだけで写真の枠は残す
const tokSlides = Object.assign({}, SLIDES, {
  leader: tokLeader.replace(/^[\s\S]*?<p:grpSpPr\/>/, '').replace(/<\/p:spTree>[\s\S]*$/, ''),
  support: tokXml.replace(/^[\s\S]*?<p:grpSpPr\/>/, '').replace(/<\/p:spTree>[\s\S]*$/, ''),
});
const parts4 = makeParts(tokSlides);
F.applyRoleIntro_(parts4, null, { by: {}, seq: 0 });
const L4 = F.xmlOf_(parts4, sPath('leader'));
ck(!/\{\{/.test(F.slideText_(L4)) && /name="\{\{プレジデント写真\}\}"/.test(L4) && !F.slideText_(F.xmlOf_(parts4, sPath('support'))).includes('氏名'),
   '外したとき：差し込み口は空・写真の枠は残す');
const parts6 = makeParts(tokSlides);
F.applyRoleIntro_(parts6, none, { by: {}, seq: 0 });
ck(/name="\{\{プレジデント写真\}\}"/.test(F.xmlOf_(parts6, sPath('leader'))) && !/\{\{/.test(F.slideText_(F.xmlOf_(parts6, sPath('leader')))),
   '未登録の期：差し込み口の雛形は、差し込み口を空にして写真の枠は残す');
const parts5 = makeParts(tokSlides);
const r5 = F.applyRoleIntro_(parts5, data, { by: {}, seq: 0 });
const S5 = F.slideText_(F.xmlOf_(parts5, sPath('support')));
ck(S5.includes('見本　七子') && S5.includes(vhMembers.join('、')) && F.slideText_(F.xmlOf_(parts5, sPath('leader'))).includes('見本　一郎')
   && r5.filled.some((f) => f.base === 'プレジデント' && f.name === '見本　一郎'), '差し込み口の雛形に入れる: ' + S5.slice(0, 120));

// ===== 6. ネットワーキング学習コーナー：「担当：」は小さく、お名前は大きく（40pt・太字）。自動縮小はやめる =====
{
  // 雛形の「担当：」の枠（斜体・自動縮小つき・1行ぶんの高さ）
  const learnBox = (id, w) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Text ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`
    + `<p:spPr>${xf(560, 400, w, 50)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`
    + '<p:txBody><a:bodyPr wrap="square"><a:normAutofit fontScale="62500" lnSpcReduction="20000"/></a:bodyPr><a:lstStyle/>'
    + '<a:p><a:r><a:rPr lang="ja-JP" sz="2400" i="1"><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill></a:rPr><a:t>旧カテゴリー</a:t></a:r></a:p>'
    + '<a:p><a:r><a:rPr lang="ja-JP" sz="2400" i="1"><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill></a:rPr><a:t>担当：前任　一郎</a:t></a:r></a:p></p:txBody></p:sp>';
  const learnSlide = (w) => slideXml(sp(2, 0, 20, 960, 60, ['ネットワーキング学習コーナー'], 4000, true)
    + sp(5, 560, 100, 300, 280, ['スピーカーの写真を挿入', '※枠にサイズを合わせてトリミング'], 1400) + pic(3, 560, 100, 300, 280, 'rId3') + learnBox(8, w));
  const runsOf = (xml) => {
    const r = F.findShapeRange_(xml, 8), seg = xml.substring(r.start, r.end);
    return F.findTagRanges_(seg, 'a:p').map((p) => (seg.substring(p.start, p.end).match(/<a:r>[\s\S]*?<\/a:r>/g) || []).map((x) => ({
      t: F.slideText_(x), sz: +(x.match(/\ssz="(\d+)"/) || [])[1], b: (x.match(/\sb="(\d)"/) || [])[1], i: (x.match(/\si="(\d)"/) || [])[1],
      color: /1F3864/.test(x) })));
  };
  const lc = F.riLearningCorner_(learnSlide(300), data, W, H);
  const got = runsOf(lc.xml);
  ck(lc.done && lc.name === '見本　四郎', '学習コーナー：担当（エデュケーションコーディネーター）を入れない: ' + J({ done: lc.done, name: lc.name }));
  ck(J(got) === J([[{ t: 'イベント企画', sz: 2000, b: '0', i: '0', color: true }],
                   [{ t: '担当：', sz: 2000, b: '1', i: '0', color: true }, { t: '見本　四郎', sz: 4000, b: '1', i: '0', color: true }]]),
     '学習コーナー：カテゴリー20pt／「担当：」20pt・お名前40ptの太字（斜体なし・色は雛形のまま）: ' + J(got));
  const seg8 = ((x) => { const r = F.findShapeRange_(x, 8); return x.substring(r.start, r.end); })(lc.xml);
  ck(/<a:noAutofit\/>/.test(seg8) && !/normAutofit|spAutoFit/.test(seg8), '学習コーナー：自動縮小が残っている（Googleスライドなどで字が小さくなる）: ' + (seg8.match(/<a:bodyPr[\s\S]*?(?:\/>|<\/a:bodyPr>)/) || [''])[0]);
  const geo = F.readShapeGeomEmu_(lc.xml, 8);
  ck(geo.cy >= Math.round((20 + 40) * 1.2 * 12700 + 2 * 45720) && geo.y === 400 * 12700, '学習コーナー：枠の高さが2行ぶんに足りない: ' + J(geo));
  // 枠が狭く、お名前が長いときは、お名前だけ小さくする（「担当：」は20ptのまま）
  const lc2 = F.riLearningCorner_(learnSlide(220), Object.assign({}, data, { roles: Object.assign({}, data.roles, { ec: { name: '見本　長い名前の方', category: '行政書士' } }) }), W, H);
  const got2 = runsOf(lc2.xml)[1] || [];
  ck(got2.length === 2 && got2[0].sz === 2000 && got2[1].sz < 4000 && got2[1].sz >= 2000, '学習コーナー：長いお名前を枠の幅に収める: ' + J(got2));
}

// ===== 写真の索引が古い（ファイルを開けない）方は、その方だけ写真なし。スライド全体を止めない =====
{
  const realDrive = sb.DriveApp;
  sb.DriveApp = { getFileById: (id) => { if (id === 'photo1') throw new Error('No item with the given ID could be found, or you do not have permission to access it.'); return realDrive.getFileById(id); } };
  const cache = { by: {}, seq: 0 }, map = {};
  let threw = null, gone = null, ok = null;
  try { gone = F.mpAddPhoto_(map, cache, '見本　二郎'); ok = F.mpAddPhoto_(map, cache, '見本　一郎'); } catch (e) { threw = e; }
  ck(!threw, '写真のファイルを1枚開けないだけで止まった: ' + (threw && threw.message));
  ck(gone === null && ok && ok.path && cache.opened === 1 && JSON.stringify(cache.gone) === JSON.stringify(['見本　二郎']),
     '開けない写真の方だけ写真なしにし、名前を控える: ' + JSON.stringify({ gone, ok: !!ok, opened: cache.opened, list: cache.gone }));
  sb.DriveApp = realDrive;
}

console.log(`役職のメンバー紹介: 検査 ${checks} 件`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: ローマ字・枠の見分け方（表・1人のページ・メンバーシップ委員会・ビジターホストチーム・各サポートチーム）・替わった方だけ入れる・'
  + '居ない方と写真の無い方・非表示・長い文字・人数が多いとき・差し込み口の雛形・外したとき・未登録の期');

// === リファーラル発表（後半スライド）===
//
// 後半スライドの「REFERRAL PRESENTATION」のページを、メンバーの人数ぶんに増やす。
// 作りはメンバープレゼンの個人ページとほぼ同じなので、写真の取り込み・カウントダウン・
// 会社名の収め方は member_presen_srv.js / ooxml.js の処理をそのまま使う。
//
// メンバープレゼンとの違い
//   ・順番は業種区分の巡回ではなく、**名簿のNo.順（1番から）**
//   ・カウントダウンは30秒ではなく**7秒**
//   ・ページは単独のファイルではなく、後半スライドの中で増える

var RF_TITLE_ = 'REFERRAL PRESENTATION';
var RF_SECONDS_ = 7;
var WEEKLY_TITLE_ = 'WEEKLY PRESENTATION';
var WEEKLY_SECONDS_ = 30;

// ひな形のページの図形は、IDではなく「役割」で見分ける。
// 前半のウィークリープレゼン・後半のリファーラル発表・2分30秒の下書きページは
// どれも同じ作りだが、図形IDはページごとに違うため。
//
//   氏名・会社名・カテゴリー … 幅の広い（6,000,000EMU超）文字箱を上から順に3つ
//   NEXT➡               … 一番下に置かれた文字箱。「NEXT」を含む方が見出し
//   写真                 … 背景でない大きな画像
// カウントダウンの数字は幅が狭いので、幅で見るだけで除ける。
function presenterShapes_(xml, slideW) {
  var wide = [], low = [], pics = [], ranges = findTagRanges_(xml, 'p:sp'), i, seg, id, off, ext, t;
  for (i = 0; i < ranges.length; i++) {
    seg = xml.substring(ranges[i].start, ranges[i].end);
    id = (seg.match(/<p:cNvPr id="(\d+)"/) || [])[1];
    off = seg.match(/<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"\s*\/>/);
    ext = seg.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"\s*\/>/);
    if (!id || !off || !ext) continue;
    t = slideText_(seg).trim();
    if (!t) continue;
    var box = { id: id, x: parseInt(off[1], 10), y: parseInt(off[2], 10),
                cx: parseInt(ext[1], 10), cy: parseInt(ext[2], 10), text: t };
    if (box.y > 6000000) low.push(box);
    else if (box.cx > 6000000 && box.y > 1000000) wide.push(box);
  }
  ranges = findTagRanges_(xml, 'p:pic');
  for (i = 0; i < ranges.length; i++) {
    seg = xml.substring(ranges[i].start, ranges[i].end);
    id = (seg.match(/<p:cNvPr id="(\d+)"/) || [])[1];
    ext = seg.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"\s*\/>/);
    if (!id || !ext) continue;
    var pcx = parseInt(ext[1], 10), pcy = parseInt(ext[2], 10);
    if (pcy < 3000000) continue;                       // 小さな飾り
    if (pcx > (slideW || 12192000) * 0.5) continue;    // 背景いっぱいの画像
    pics.push({ id: id, cx: pcx, cy: pcy });
  }
  pics.sort(function (a, b) { return b.cx * b.cy - a.cx * a.cy; });   // 候補が複数あれば大きいもの
  wide.sort(function (a, b) { return a.y - b.y; });
  low.sort(function (a, b) { return a.x - b.x; });
  var label = null, next = null;
  for (i = 0; i < low.length; i++) {
    if (!label && low[i].text.indexOf('NEXT') >= 0) label = low[i];
    else if (!next) next = low[i];
  }
  return {
    nameBox: wide[0] ? wide[0].id : 0,
    company: wide[1] ? wide[1].id : 0,
    category: wide[2] ? wide[2].id : 0,
    nextLabel: label ? label.id : 0,
    nextName: next ? next.id : 0,
    photo: pics[0] ? pics[0].id : 0
  };
}

function slideText_(xml) {
  var t = '', m, re = /<a:t(?=[\s>])[^>]*>([\s\S]*?)<\/a:t>/g;
  while ((m = re.exec(xml)) !== null) t += unescapeXml_(m[1]);
  return t;
}

function relsPathOf_(slidePath) {
  return slidePath.replace(/^(ppt\/slides\/)(slide\d+\.xml)$/, '$1_rels/$2.rels');
}

// ひな形にするページ（複数あればまとめて置き換える）
function rfFindModels_(parts, title) {
  var out = [], p;
  for (p in parts) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(p)) continue;
    var xml = xmlOf_(parts, p);
    if (!xml || slideText_(xml).indexOf(title || RF_TITLE_) < 0) continue;
    out.push({ path: p, xml: xml, no: parseInt(p.replace(/\D+/g, ''), 10) });
  }
  out.sort(function (a, b) { return a.no - b.no; });
  return out;
}

// 1人ぶんのページを作る
function rfBuildSlide_(tplXml, tplRels, item, photo, SH) {
  var xml = tplXml;
  xml = setParagraphsInShape_(xml, SH.nameBox, [item.name || '']);

  xml = setParagraphsInShape_(xml, SH.company, item.companyLines || ['']);
  if (item.companyGeom) xml = setShapeGeomEmu_(xml, SH.company, item.companyGeom);
  if (item.companyPt) xml = setFontSizeInShape_(xml, SH.company, item.companyPt);

  if (item.categoryTop) {
    var cg = readShapeGeomEmu_(xml, SH.category) || {};
    xml = setShapeGeomEmu_(xml, SH.category, { x: cg.x, y: item.categoryTop, cx: cg.cx });
  }
  xml = setBodyAnchorInShape_(xml, SH.category, 't');
  xml = noAutofitInShape_(xml, SH.category);
  xml = setParagraphsInShape_(xml, SH.category, item.categoryLines || ['']);
  if (item.categoryPt) xml = setFontSizeInShape_(xml, SH.category, item.categoryPt);
  if (item.categoryTight) xml = setLineSpacingInShape_(xml, SH.category, 85);

  if (item.nextName && SH.nextName) {
    xml = setParagraphsInShape_(xml, SH.nextName, [item.nextName]);
  } else {
    xml = removeShape_(xml, SH.nextName);
    xml = removeShape_(xml, SH.nextLabel);
  }

  // カウントダウンを作り直す（見本の書式のまま秒数だけ変える）。
  //   auto が true … スライドが出たらすぐ始まり、終わったら自動で次へ
  //   auto が false … クリックで始まり、次のページへもクリックで進む（元のテンプレートと同じ）
  var sec = item.seconds || RF_SECONDS_;
  if (item.auto) {
    xml = mpSetCountdown_(xml, sec, false);
    xml = mpAutoAdvance_(xml, (sec + 1) * 1000);
  } else {
    xml = mpSetCountdown_(xml, sec, true);
    xml = mpNoAutoAdvance_(xml);
  }

  var rels = tplRels;
  if (!SH.photo) {
    // 写真の枠が無いひな形（そのまま）
  } else if (!photo) {
    xml = removeShape_(xml, SH.photo);
  } else {
    var box = readShapeGeomEmu_(xml, SH.photo);
    if (box && photo.width && photo.height) {
      xml = setSrcRectInPic_(xml, SH.photo, coverCrop_(photo.width, photo.height, box.cx, box.cy));
    }
    var set = setPicImage_(xml, rels, SH.photo, '../media/' + photo.path.replace('ppt/media/', ''));
    xml = set.xml; rels = set.rels;
  }
  return { xml: xml, rels: rels };
}

// ひな形のページを、人数ぶんのページに置き換える。
// opts: { title: 見出しの文字, label: 画面に出す名前, seconds: 既定の秒数 }
function expandPresenterSlides_(parts, items, opts) {
  if (!items || !items.length) return null;
  var o = opts || {}, title = o.title || RF_TITLE_, label = o.label || 'リファーラル発表';
  var models = rfFindModels_(parts, title);
  if (!models.length) return { message: label + 'のひな形ページが見つかりませんでした。' };

  var model = models[0];
  var prsXml = xmlOf_(parts, 'ppt/presentation.xml') || '';
  var szm = prsXml.match(/<p:sldSz\s+cx="(\d+)"/);
  var SH = presenterShapes_(model.xml, szm ? parseInt(szm[1], 10) : 12192000);
  if (!SH.nameBox || !SH.company || !SH.category) {
    return { message: label + 'のひな形から、氏名・会社名・カテゴリーの枠を見分けられませんでした。' };
  }
  // ノートはページごとに1つしか持てないので、複製する関係からは外す
  var tplRels = (xmlOf_(parts, relsPathOf_(model.path)) || '')
    .replace(/<Relationship\b[^>]*notesSlides\/[^>]*\/>/g, '');

  var prs = xmlOf_(parts, 'ppt/presentation.xml');
  var prsRels = xmlOf_(parts, 'ppt/_rels/presentation.xml.rels');
  var ct = xmlOf_(parts, '[Content_Types].xml');

  // rId → スライドのパス
  var rid2path = {}, m, re = /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="slides\/(slide\d+\.xml)"[^>]*\/>/g;
  while ((m = re.exec(prsRels)) !== null) rid2path[m[1]] = 'ppt/slides/' + m[2];

  var entries = [], reS = /<p:sldId id="(\d+)" r:id="([^"]+)"\/>/g;
  while ((m = reS.exec(prs)) !== null) entries.push({ id: parseInt(m[1], 10), rid: m[2], path: rid2path[m[2]] || '' });

  var drop = {}, i;
  for (i = 0; i < models.length; i++) drop[models[i].path] = true;
  var at = -1;
  for (i = 0; i < entries.length; i++) if (drop[entries[i].path]) { at = i; break; }
  if (at < 0) return { message: 'リファーラル発表のページが並びの中に見つかりませんでした。' };

  var maxId = 255, maxRid = 0;
  for (i = 0; i < entries.length; i++) maxId = Math.max(maxId, entries[i].id);
  var reR = /Id="rId(\d+)"/g;
  while ((m = reR.exec(prsRels)) !== null) maxRid = Math.max(maxRid, parseInt(m[1], 10));
  var maxSlide = 0;
  for (var p in parts) {
    var sm = p.match(/^ppt\/slides\/slide(\d+)\.xml$/);
    if (sm) maxSlide = Math.max(maxSlide, parseInt(sm[1], 10));
  }

  var cache = o.photoCache || { by: {}, seq: 0 }, made = [], noPhoto = [];
  for (i = 0; i < items.length; i++) {
    var it = items[i], n = maxSlide + 1 + i, rid = 'rId' + (maxRid + 1 + i);
    var photo = mpAddPhoto_(parts, cache, it.photoName || it.name);
    if (!photo && it.name && noPhoto.indexOf(it.name) === -1) noPhoto.push(it.name);
    var built = rfBuildSlide_(model.xml, tplRels, it, photo, SH);
    var path = 'ppt/slides/slide' + n + '.xml';
    putXml_(parts, path, built.xml);
    putXml_(parts, relsPathOf_(path), built.rels);
    made.push({ id: ++maxId, rid: rid, path: path, n: n });
    prsRels = prsRels.replace('</Relationships>',
      '<Relationship Id="' + rid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"'
      + ' Target="slides/slide' + n + '.xml"/></Relationships>');
    ct = ct.replace('</Types>', '<Override PartName="/' + path + '"'
      + ' ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  }

  // 並びを作り直す（ひな形の場所に、作ったページをまとめて入れる）
  var out = '';
  for (i = 0; i < entries.length; i++) {
    if (drop[entries[i].path]) {
      if (i === at) {
        for (var k = 0; k < made.length; k++) {
          out += '<p:sldId id="' + made[k].id + '" r:id="' + made[k].rid + '"/>';
        }
      }
      continue;
    }
    out += '<p:sldId id="' + entries[i].id + '" r:id="' + entries[i].rid + '"/>';
  }
  prs = prs.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, '<p:sldIdLst>' + out + '</p:sldIdLst>');

  // ひな形のページを片付ける
  for (i = 0; i < models.length; i++) {
    var mp = models[i].path;
    prsRels = prsRels.replace(new RegExp('<Relationship\\b[^>]*Target="slides/'
      + mp.replace('ppt/slides/', '') + '"[^>]*/>'), '');
    ct = ct.replace(new RegExp('<Override\\b[^>]*PartName="/' + mp.replace(/\//g, '\\/') + '"[^>]*/>'), '');
    delete parts[mp];
    delete parts[relsPathOf_(mp)];
  }

  putXml_(parts, 'ppt/presentation.xml', prs);
  putXml_(parts, 'ppt/_rels/presentation.xml.rels', prsRels);
  putXml_(parts, '[Content_Types].xml', ct);

  var msg = label + 'のページを ' + made.length + '枚 作りました（'
          + (items[0].seconds || o.seconds || RF_SECONDS_) + '秒）。';
  if (noPhoto.length) msg += '\n写真が見つからない方: ' + noPhoto.join('、');
  var auto = false;
  for (i = 0; i < items.length; i++) if (items[i].auto) auto = true;
  msg = msg.replace(/秒）。$/, '秒・' + (auto ? '自動で次へ' : 'クリックで次へ') + '）。');
  return { message: msg, made: made.length, noPhoto: noPhoto, auto: auto,
           paths: made.map(function (x) { return x.path; }) };
}

// 画面用：ひな形ページの枠の位置を返す（会社名の収め方を画面側で決めるのに使う）
function rfLayoutBoxes_(parts, title) {
  var models = rfFindModels_(parts, title);
  if (!models.length) return null;
  var xml = models[0].xml;
  var prsXml = xmlOf_(parts, 'ppt/presentation.xml') || '';
  var szm = prsXml.match(/<p:sldSz\s+cx="(\d+)"/);
  var SH = presenterShapes_(xml, szm ? parseInt(szm[1], 10) : 12192000);
  if (!SH.company) return null;
  var co = readShapeGeomEmu_(xml, SH.company);
  var ca = SH.category ? readShapeGeomEmu_(xml, SH.category) : null;
  if (!co) return null;
  return { companyTall: co, categoryLow: ca ? ca.y : 0, categoryWidth: ca ? ca.cx : 0,
           hasNext: !!SH.nextName, slides: models.length };
}

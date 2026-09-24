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
// ひな形のページに載っている図形ID。無ければその場で止める。
var RF_SHAPES_ = { photo: 3, nameBox: 56, company: 2, category: 12, nextName: 14, nextLabel: 13 };

function slideText_(xml) {
  var t = '', m, re = /<a:t(?=[\s>])[^>]*>([\s\S]*?)<\/a:t>/g;
  while ((m = re.exec(xml)) !== null) t += unescapeXml_(m[1]);
  return t;
}

function relsPathOf_(slidePath) {
  return slidePath.replace(/^(ppt\/slides\/)(slide\d+\.xml)$/, '$1_rels/$2.rels');
}

// ひな形にするページ（複数あればまとめて置き換える）
function rfFindModels_(parts) {
  var out = [], p;
  for (p in parts) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(p)) continue;
    var xml = xmlOf_(parts, p);
    if (!xml || slideText_(xml).indexOf(RF_TITLE_) < 0) continue;
    out.push({ path: p, xml: xml, no: parseInt(p.replace(/\D+/g, ''), 10) });
  }
  out.sort(function (a, b) { return a.no - b.no; });
  return out;
}

// 1人ぶんのページを作る
function rfBuildSlide_(tplXml, tplRels, item, photo) {
  var xml = tplXml;
  xml = setParagraphsInShape_(xml, RF_SHAPES_.nameBox, [item.name || '']);

  xml = setParagraphsInShape_(xml, RF_SHAPES_.company, item.companyLines || ['']);
  if (item.companyGeom) xml = setShapeGeomEmu_(xml, RF_SHAPES_.company, item.companyGeom);
  if (item.companyPt) xml = setFontSizeInShape_(xml, RF_SHAPES_.company, item.companyPt);

  if (item.categoryTop) {
    var cg = readShapeGeomEmu_(xml, RF_SHAPES_.category) || {};
    xml = setShapeGeomEmu_(xml, RF_SHAPES_.category, { x: cg.x, y: item.categoryTop, cx: cg.cx });
  }
  xml = setBodyAnchorInShape_(xml, RF_SHAPES_.category, 't');
  xml = noAutofitInShape_(xml, RF_SHAPES_.category);
  xml = setParagraphsInShape_(xml, RF_SHAPES_.category, item.categoryLines || ['']);
  if (item.categoryPt) xml = setFontSizeInShape_(xml, RF_SHAPES_.category, item.categoryPt);
  if (item.categoryTight) xml = setLineSpacingInShape_(xml, RF_SHAPES_.category, 85);

  if (item.nextName) {
    xml = setParagraphsInShape_(xml, RF_SHAPES_.nextName, [item.nextName]);
  } else {
    xml = removeShape_(xml, RF_SHAPES_.nextName);
    xml = removeShape_(xml, RF_SHAPES_.nextLabel);
  }

  // カウントダウンを作り直し（見本の書式のまま秒数だけ変える）、自動で次へ進むようにする
  xml = mpSetCountdown_(xml, item.seconds || RF_SECONDS_);
  xml = mpAutoAdvance_(xml, ((item.seconds || RF_SECONDS_) + 1) * 1000);

  var rels = tplRels;
  if (!photo) {
    xml = removeShape_(xml, RF_SHAPES_.photo);
  } else {
    var box = readShapeGeomEmu_(xml, RF_SHAPES_.photo);
    if (box && photo.width && photo.height) {
      xml = setSrcRectInPic_(xml, RF_SHAPES_.photo, coverCrop_(photo.width, photo.height, box.cx, box.cy));
    }
    var rid = (tplXml.substring(findShapeRange_(tplXml, RF_SHAPES_.photo).start,
                                findShapeRange_(tplXml, RF_SHAPES_.photo).end)
                     .match(/<a:blip[^>]*r:embed="([^"]+)"/) || [])[1];
    if (rid) rels = retargetRel_(rels, rid, '../media/' + photo.path.replace('ppt/media/', ''));
  }
  return { xml: xml, rels: rels };
}

// ひな形のページを、人数ぶんのページに置き換える
function expandReferralSlides_(parts, items) {
  if (!items || !items.length) return null;
  var models = rfFindModels_(parts);
  if (!models.length) return { message: 'リファーラル発表のひな形ページが見つかりませんでした。' };

  var model = models[0];
  var need = [RF_SHAPES_.photo, RF_SHAPES_.nameBox, RF_SHAPES_.company, RF_SHAPES_.category];
  var miss = missingShapeIds_(model.xml, need);
  if (miss.length) {
    throw new Error('リファーラル発表のひな形に図形が見つかりません（ID: ' + miss.join('、') + '）。');
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

  var cache = { by: {}, seq: 0 }, made = [], noPhoto = [];
  for (i = 0; i < items.length; i++) {
    var it = items[i], n = maxSlide + 1 + i, rid = 'rId' + (maxRid + 1 + i);
    var photo = mpAddPhoto_(parts, cache, it.photoName || it.name);
    if (!photo && it.name && noPhoto.indexOf(it.name) === -1) noPhoto.push(it.name);
    var built = rfBuildSlide_(model.xml, tplRels, it, photo);
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
  mpUseTimings_(parts);

  var msg = 'リファーラル発表のページを ' + made.length + '枚 作りました（'
          + (items[0].seconds || RF_SECONDS_) + '秒・名簿のNo.順）。';
  if (noPhoto.length) msg += '\n写真が見つからない方: ' + noPhoto.join('、');
  return { message: msg, made: made.length, noPhoto: noPhoto };
}

// 画面用：ひな形ページの枠の位置を返す（会社名の収め方を画面側で決めるのに使う）
function rfLayoutBoxes_(parts) {
  var models = rfFindModels_(parts);
  if (!models.length) return null;
  var xml = models[0].xml;
  var co = readShapeGeomEmu_(xml, RF_SHAPES_.company);
  var ca = readShapeGeomEmu_(xml, RF_SHAPES_.category);
  if (!co) return null;
  return { companyTall: co, categoryLow: ca ? ca.y : 0, categoryWidth: ca ? ca.cx : 0 };
}

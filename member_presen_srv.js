// === メンバープレゼンスライドの自動生成 ===
//
// もとは「メンバープレゼンスライド 自動生成ツール」という単体のHTMLで、
// メンバーリストHTMLを読ませて pptx を作るものだった。それを名簿システムに取り込み、
// 「メンバー名簿」と「メンバー写真」から直接作れるようにしたもの。
// 出来上がりの見た目は元ツールと同じになるようにしてある。
//
// 【役割を2つに分けている理由】
//  ・会社名の折り返しや文字の縮小には文字幅の実測が要るが、GASには文字を測る手段が無い。
//    そこで組版はブラウザ側（canvasで実測）で行い、決まった結果だけを送ってもらう。
//  ・テンプレートは15MB近くある。画面に送ると重いので、サーバーがDriveから直接読む。
//    → 画面とサーバーの間を行き来するのは小さなJSONだけになる。
//
// テンプレートは「⚙️ 設定 ＞ 大きなスライド」の「メンバープレゼン」に登録する。

var MP_TEMPLATE_KIND_ = 'memberPresen';

// テンプレート内の図形ID（<p:cNvPr id="…">）。
// 元ツールは座標の範囲で図形を探していたが、同じテンプレートを使う前提なので
// IDで指定している。テンプレートが差し替わって見つからないときは、その場で止める。
var MP_OVERVIEW_ = {
  slide: 'ppt/slides/slide1.xml', rels: 'ppt/slides/_rels/slide1.xml.rels',
  title: 11,        // 業種区分名
  nextName: 31,     // Next Presenter の氏名
  photo: 2,         // 先頭メンバーの写真
  table: 6,         // 専門分野／メンバーの表
  photoRid: 'rId3'
};
var MP_INDIVIDUAL_ = {
  slide: 'ppt/slides/slide3.xml', rels: 'ppt/slides/_rels/slide3.xml.rels',
  photo: 3, name: 56, company: 2, category: 12,
  nextName: 14,     // NEXT➡ の氏名
  nextLabel: 6,     // NEXT➡ の文字
  photoRid: 'rId2'
};
var MP_ROWS_PER_OVERVIEW_ = 7;     // 一覧の表は7行。8人以上いるとページが増える

// 会社名・カテゴリーの枠（EMU）。元ツールが実物のスライドから採った値。
var MP_COMPANY_DEFAULT_ = { x: 4830858, y: 2849608, cx: 7461517, cy: 769441 };
var MP_COMPANY_TALL_    = { x: 5087424, y: 2776058, cx: 6893161, cy: 1446550 };
var MP_CATEGORY_X_ = 4830481, MP_CATEGORY_CX_ = 7311519;
var MP_CATEGORY_TOP_ = 3456634;        // 会社名が1行のとき
var MP_CATEGORY_TOP_LOW_ = 4071101;    // 会社名が2行のとき（そのぶん下げる）

// 巡回の基準。この週が「プロモーション」始まりで、以降1週ごとに1つずつずれる。
var MP_BASE_DATE_ = '2026/08/19';
var MP_BASE_BLOCK_ = 'プロモーション';

function openMemberPresenDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('member_presen').setWidth(760).setHeight(700),
    'メンバープレゼンスライドの作成');
}

// 業種区分の照合キー。「美容と健康」「美容・健康」のような表記ゆれを吸収する
function mpCatKey_(s) {
  return String(s == null ? '' : s).normalize('NFKC').replace(/[\s・･＆&と]/g, '');
}

// 業種区分マスタ → 巡回順に並んだブロックの一覧
function mpBlocks_() {
  var cats = getCategoryMaster(), out = [];
  for (var i = 0; i < cats.length; i++) {
    var c = cats[i];
    if (!c.key) continue;
    out.push({ key: c.key, block: c.block || c.label || c.key, order: c.order || (i + 1) });
  }
  return out;
}

// 指定の開催日に、どの業種区分から始めるか。基準日からの週数で1つずつずらす。
function mpOrderFor_(blocks, dateStr) {
  var n = blocks.length;
  if (!n) return [];
  var base = new Date(MP_BASE_DATE_ + ' 00:00:00'), target = new Date(dateStr.replace(/-/g, '/') + ' 00:00:00');
  var weeks = Math.round((target.getTime() - base.getTime()) / (7 * 86400000));
  var baseIdx = 0;
  for (var i = 0; i < n; i++) if (mpCatKey_(blocks[i].block) === mpCatKey_(MP_BASE_BLOCK_)) { baseIdx = i; break; }
  var start = ((baseIdx + weeks) % n + n) % n, order = [];
  for (var j = 0; j < n; j++) order.push(blocks[(start + j) % n].key);
  return order;
}

// 画面が必要とする情報を一度に返す
function getMemberPresenContext() {
  try {
    var master = getMemberMaster(), blocks = mpBlocks_(), members = [];
    var byKey = {}, byNorm = {};
    for (var b = 0; b < blocks.length; b++) { byKey[blocks[b].key] = blocks[b]; byNorm[mpCatKey_(blocks[b].key)] = blocks[b]; }

    for (var i = 0; i < (master.members || []).length; i++) {
      var m = master.members[i];
      var hit = byKey[m.cat] || byNorm[mpCatKey_(m.cat)] || null;
      members.push({ name: m.name, company: m.company, title: m.title,
                     cat: m.cat, blockKey: hit ? hit.key : '', hasPhoto: !!findPhotoIdForName_(m.name) });
    }

    var cands = getMeetingCandidates();
    for (var c = 0; c < cands.length; c++) cands[c].order = mpOrderFor_(blocks, cands[c].dateValue);

    var tpl = null, st = getBigTemplateStatus();
    for (var t = 0; t < st.templates.length; t++) if (st.templates[t].kind === MP_TEMPLATE_KIND_) tpl = st.templates[t];

    return { ok: true, members: members, blocks: blocks, candidates: cands,
             template: tpl, rowsPerPage: MP_ROWS_PER_OVERVIEW_ };
  } catch (e) {
    console.error('[MPRESEN] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '読み込みに失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// 写真をテンプレートの中に入れて、そのパスと大きさを返す。
// 同じ人は扉ページと個人ページの2回出てくるので、1枚だけ入れて使い回す。
function mpAddPhoto_(map, cache, name) {
  var key = normName_(name);
  if (key in cache.by) return cache.by[key];
  var id = key ? findPhotoIdForName_(name) : '';
  if (!id) { cache.by[key] = null; return null; }
  var blob = DriveApp.getFileById(id).getBlob();
  var ext = String(blob.getContentType() || '').toLowerCase().indexOf('png') >= 0 ? 'png' : 'jpeg';

  // 写真の縦横比は、切り抜き（srcRect）を決めるのに要る。
  // 大きさはファイル先頭のヘッダーだけで分かるが、それを読むには一度バイト列に
  // 展開することになる。何十人ぶんも抱えるとメモリが足りなくなるので、
  //  ・読み終えたバイト列はすぐ捨て、zipには元のBlobをそのまま入れる
  //  ・一度測った大きさは覚えておき、次からはバイト列を触らない
  var size = null, ck = 'mpsize_' + id, cs = null;
  try { cs = CacheService.getScriptCache(); } catch (e) {}
  if (cs) { try { var hit = cs.get(ck); if (hit) size = JSON.parse(hit); } catch (e) {} }
  if (!size) {
    try { size = imageSizeOf_(blob.getBytes()); } catch (e) {}
    if (size && cs) { try { cs.put(ck, JSON.stringify(size), 21600); } catch (e) {} }
  }
  cache.seq++;
  var path = 'ppt/media/mpphoto' + cache.seq + '.' + ext;
  map[path] = blob.setName(path);
  cache.by[key] = { path: path, width: size ? size.width : 0, height: size ? size.height : 0 };
  return cache.by[key];
}

// 関係ファイルから1件消す
function mpRemoveRel_(relsXml, rId) {
  return relsXml.replace(new RegExp('<Relationship\\b[^>]*\\bId="' + rId + '"[^>]*/>'), '');
}

// 写真を差し替える（無い人は写真の枠ごと消す）
function mpApplyPhoto_(xml, rels, def, photo) {
  if (!photo) return { xml: removeShape_(xml, def.photo), rels: mpRemoveRel_(rels, def.photoRid) };
  var box = readShapeGeomEmu_(xml, def.photo);
  if (box && photo.width && photo.height) {
    xml = setSrcRectInPic_(xml, def.photo, coverCrop_(photo.width, photo.height, box.cx, box.cy));
  }
  return { xml: xml, rels: retargetRel_(rels, def.photoRid, '../media/' + photo.path.replace('ppt/media/', '')) };
}

// 業種区分の扉ページ（一覧の表つき）
function mpOverviewSlide_(tplXml, tplRels, item, photo) {
  var xml = tplXml;
  xml = setParagraphsInShape_(xml, MP_OVERVIEW_.title, [item.block]);
  xml = setParagraphsInShape_(xml, MP_OVERVIEW_.nextName, [item.nextName || '']);
  for (var r = 0; r < MP_ROWS_PER_OVERVIEW_; r++) {
    var row = (item.rows || [])[r] || { title: '', name: '' };
    xml = setTableCellText_(xml, MP_OVERVIEW_.table, r + 1, 0, row.title || '');
    xml = setTableCellText_(xml, MP_OVERVIEW_.table, r + 1, 1, row.name || '');
  }
  return mpApplyPhoto_(xml, tplRels, MP_OVERVIEW_, photo);
}

// 個人ページ
function mpIndividualSlide_(tplXml, tplRels, item, photo) {
  var xml = tplXml;
  xml = setParagraphsInShape_(xml, MP_INDIVIDUAL_.name, [item.name || '']);

  // 会社名：行の分け方と文字の大きさは画面側で決めてある
  xml = setParagraphsInShape_(xml, MP_INDIVIDUAL_.company, item.companyLines || ['']);
  xml = setShapeGeomEmu_(xml, MP_INDIVIDUAL_.company, item.companyTall ? MP_COMPANY_TALL_ : MP_COMPANY_DEFAULT_);
  if (item.companyPt) xml = setFontSizeInShape_(xml, MP_INDIVIDUAL_.company, item.companyPt);

  // カテゴリー：会社名が2行になったぶんだけ下げる
  xml = setShapeGeomEmu_(xml, MP_INDIVIDUAL_.category,
    { x: MP_CATEGORY_X_, y: item.companyTall ? MP_CATEGORY_TOP_LOW_ : MP_CATEGORY_TOP_, cx: MP_CATEGORY_CX_ });
  xml = setBodyAnchorInShape_(xml, MP_INDIVIDUAL_.category, 't');
  xml = noAutofitInShape_(xml, MP_INDIVIDUAL_.category);
  xml = setParagraphsInShape_(xml, MP_INDIVIDUAL_.category, item.categoryLines || ['']);
  if (item.categoryPt) xml = setFontSizeInShape_(xml, MP_INDIVIDUAL_.category, item.categoryPt);
  if (item.categoryTight) xml = setLineSpacingInShape_(xml, MP_INDIVIDUAL_.category, 85);

  if (item.nextName) {
    xml = setParagraphsInShape_(xml, MP_INDIVIDUAL_.nextName, [item.nextName]);
  } else {
    xml = removeShape_(xml, MP_INDIVIDUAL_.nextName);
    xml = removeShape_(xml, MP_INDIVIDUAL_.nextLabel);
  }
  return mpApplyPhoto_(xml, tplRels, MP_INDIVIDUAL_, photo);
}

// テンプレートの中身を、指示どおりのスライドの並びに置き換える
function buildMemberPresenSlides_(map, items) {
  var ovXml = xmlOf_(map, MP_OVERVIEW_.slide), ovRels = xmlOf_(map, MP_OVERVIEW_.rels);
  var ivXml = xmlOf_(map, MP_INDIVIDUAL_.slide), ivRels = xmlOf_(map, MP_INDIVIDUAL_.rels);
  if (!ovXml || !ivXml || !ovRels || !ivRels) {
    throw new Error('テンプレートに slide1（業種区分の扉）と slide3（個人ページ）が見つかりません。'
      + '「⚙️ 設定 ＞ 大きなスライド」に登録したファイルをご確認ください。');
  }
  var miss = missingShapeIds_(ovXml, [MP_OVERVIEW_.title, MP_OVERVIEW_.nextName, MP_OVERVIEW_.photo, MP_OVERVIEW_.table])
    .concat(missingShapeIds_(ivXml, [MP_INDIVIDUAL_.photo, MP_INDIVIDUAL_.name, MP_INDIVIDUAL_.company,
                                     MP_INDIVIDUAL_.category, MP_INDIVIDUAL_.nextName, MP_INDIVIDUAL_.nextLabel]));
  if (miss.length) {
    throw new Error('テンプレートの図形が見つかりません（ID: ' + miss.join('、') + '）。'
      + 'テンプレートを作り直した場合は、図形の並びが変わっている可能性があります。');
  }

  // 元のスライドとノートを全部外す（ひな形2枚の内容はもう読み終えている）
  var paths = [], p;
  for (p in map) paths.push(p);
  for (var d = 0; d < paths.length; d++) {
    if (/^ppt\/slides\/(_rels\/)?slide\d+\.xml(\.rels)?$/.test(paths[d])
     || /^ppt\/notesSlides\//.test(paths[d])) delete map[paths[d]];
  }

  var prs = xmlOf_(map, 'ppt/presentation.xml');
  var prsRels = xmlOf_(map, 'ppt/_rels/presentation.xml.rels');
  var ct = xmlOf_(map, '[Content_Types].xml');

  // スライドの関係だけ外し、残った中で一番大きい rId の続きから振る
  prsRels = prsRels.replace(/<Relationship\b[^>]*relationships\/slide"[^>]*\/>/g, '');
  var maxRid = 0, m, reR = /Id="rId(\d+)"/g;
  while ((m = reR.exec(prsRels)) !== null) maxRid = Math.max(maxRid, parseInt(m[1], 10));

  var cache = { by: {}, seq: 0 }, sldIds = '', relAdd = '', ctAdd = '', noPhoto = [];
  for (var i = 0; i < items.length; i++) {
    var it = items[i], n = i + 1, rid = 'rId' + (maxRid + n);
    var photo = mpAddPhoto_(map, cache, it.photoName);
    if (!photo && it.photoName && noPhoto.indexOf(it.photoName) === -1) noPhoto.push(it.photoName);

    var built = (it.kind === 'overview')
      ? mpOverviewSlide_(ovXml, ovRels, it, photo)
      : mpIndividualSlide_(ivXml, ivRels, it, photo);

    putXml_(map, 'ppt/slides/slide' + n + '.xml', built.xml);
    putXml_(map, 'ppt/slides/_rels/slide' + n + '.xml.rels', built.rels);
    sldIds += '<p:sldId id="' + (255 + n) + '" r:id="' + rid + '"/>';
    relAdd += '<Relationship Id="' + rid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"'
            + ' Target="slides/slide' + n + '.xml"/>';
    ctAdd  += '<Override PartName="/ppt/slides/slide' + n + '.xml"'
            + ' ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>';
  }

  prs = prs.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, '<p:sldIdLst>' + sldIds + '</p:sldIdLst>');
  prsRels = prsRels.replace('</Relationships>', relAdd + '</Relationships>');
  ct = ct.replace(/<Override\b[^>]*PartName="\/ppt\/(slides\/slide|notesSlides\/notesSlide)\d+\.xml"[^>]*\/>/g, '');
  ct = ct.replace('</Types>', ctAdd + '</Types>');
  putXml_(map, 'ppt/presentation.xml', prs);
  putXml_(map, 'ppt/_rels/presentation.xml.rels', prsRels);
  putXml_(map, '[Content_Types].xml', ct);

  var pruned = mpPruneMedia_(map);
  return { slides: items.length, photos: cache.seq, noPhoto: noPhoto, prunedMedia: pruned };
}

// どこからも使われなくなった画像を捨てる。
// 元の61枚ぶんの写真が残ったままだと、出来上がりが無駄に十数MBになる。
function mpPruneMedia_(map) {
  var used = {}, p, m;
  for (p in map) {
    if (!/\.rels$/.test(p)) continue;
    var re = /Target="[^"]*media\/([^"\/]+)"/g, xml = xmlOf_(map, p);
    while ((m = re.exec(xml)) !== null) used[m[1]] = true;
  }
  var drop = [];
  for (p in map) {
    if (p.indexOf('ppt/media/') !== 0) continue;
    if (!used[p.substring('ppt/media/'.length)]) drop.push(p);
  }
  for (var i = 0; i < drop.length; i++) delete map[drop[i]];
  return drop.length;
}

// 画面から呼ぶ本体
function generateMemberPresenSlides(payload) {
  try {
    var items = (payload && payload.items) || [];
    if (!items.length) return { ok: false, message: '作成するスライドがありません。' };
    var label = (payload.dateLabel || '') + (payload.dateLabel ? '_' : '');
    var outName = label + 'BNI_メンバープレゼン.pptx';

    var r = editPptxOnServer_(MP_TEMPLATE_KIND_, outName, function (map) {
      return buildMemberPresenSlides_(map, items);
    });

    var info = r.info || {};
    var msg = '✅ ' + info.slides + '枚のスライドを作りました（' + (r.timing.合計 / 1000).toFixed(1) + '秒）。';
    if (info.noPhoto && info.noPhoto.length) {
      msg += '\n⚠ 写真が見つからなかった方は写真なしで出しています: ' + info.noPhoto.join('、');
    }
    console.log('[MPRESEN] ' + msg.replace(/\n/g, ' / '));
    return { ok: true, message: msg, url: r.saved.url, downloadUrl: r.saved.downloadUrl,
             fileName: outName, slides: info.slides, noPhoto: info.noPhoto || [],
             seconds: r.timing.合計 / 1000 };
  } catch (e) {
    console.error('[MPRESEN] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '作成に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

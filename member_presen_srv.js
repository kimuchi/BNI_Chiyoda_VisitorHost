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
  name: '業種区分の扉ページ',
  title: 11,        // 業種区分名
  nextName: 31,     // Next Presenter の氏名
  photo: 2,         // 先頭メンバーの写真
  table: 6,         // 専門分野／メンバーの表
  photoRid: 'rId3'
};
var MP_INDIVIDUAL_ = {
  name: '個人ページ',
  photo: 3, nameBox: 56, company: 2, category: 12,
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

// 業種区分マスタ → 巡回の並び。
// マスタには昔の区分や表記ゆれの行が残っていることがあるので、
// 「ブロック表示名が同じもの」は1つにまとめる（画面に同じ名前が2つ並ぶのを防ぐ）。
function mpCycle_() {
  var cats = getCategoryMaster(), seen = {}, out = [];
  for (var i = 0; i < cats.length; i++) {
    var c = cats[i];
    if (!c.key) continue;
    var block = c.block || c.label || c.key, g = mpCatKey_(block);
    if (seen[g]) { seen[g].keys.push(c.key); continue; }
    seen[g] = { gkey: g, block: block, order: c.order || (i + 1), keys: [c.key], count: 0, known: true };
    out.push(seen[g]);
  }
  out.sort(function (a, b) { return a.order - b.order; });
  return out;
}

// 巡回の並びに、名簿の人数を乗せる。
// マスタに無い業種区分の人も落とさず、末尾に足して人数を見せる。
function mpBlocks_(members) {
  var cycle = mpCycle_(), byKey = {}, i, j;
  for (i = 0; i < cycle.length; i++) {
    for (j = 0; j < cycle[i].keys.length; j++) byKey[mpCatKey_(cycle[i].keys[j])] = cycle[i];
  }
  var extra = {};
  for (i = 0; i < members.length; i++) {
    var raw = members[i].cat;
    if (!raw) continue;                                  // 業種区分が空の人は数えない
    var nk = mpCatKey_(raw), hit = byKey[nk];
    if (!hit) {
      hit = extra[nk] || (extra[nk] = { gkey: nk, block: raw, order: 9999, keys: [raw], count: 0, known: false });
      if (cycle.indexOf(hit) < 0) cycle.push(hit);
      byKey[nk] = hit;
    }
    hit.count++;
    members[i].blockKey = hit.gkey;
  }
  return cycle;
}

// 指定の開催日に、どの業種区分から始めるか。基準日からの週数で1つずつずらす。
// 人数が0の区分に当たったときは、次の「人がいる区分」まで進める。
function mpStartFor_(cycle, dateStr) {
  var n = cycle.length;
  if (!n) return '';
  var base = new Date(MP_BASE_DATE_ + ' 00:00:00'), target = new Date(dateStr.replace(/-/g, '/') + ' 00:00:00');
  var weeks = Math.round((target.getTime() - base.getTime()) / (7 * 86400000));
  var baseIdx = 0, i;
  for (i = 0; i < n; i++) if (cycle[i].gkey === mpCatKey_(MP_BASE_BLOCK_)) { baseIdx = i; break; }
  var start = ((baseIdx + weeks) % n + n) % n;
  for (i = 0; i < n; i++) {
    var c = cycle[(start + i) % n];
    if (c.count > 0) return c.gkey;
  }
  return cycle[start].gkey;
}

// 画面が必要とする情報を一度に返す
function getMemberPresenContext() {
  try {
    var master = getMemberMaster(), members = [];
    for (var i = 0; i < (master.members || []).length; i++) {
      var m = master.members[i];
      members.push({ name: m.name, company: m.company, title: m.title,
                     cat: m.cat, blockKey: '', hasPhoto: !!findPhotoIdForName_(m.name) });
    }
    var cycle = mpBlocks_(members);          // members[].blockKey がここで決まる

    var cands = getMeetingCandidates();
    for (var c = 0; c < cands.length; c++) cands[c].start = mpStartFor_(cycle, cands[c].dateValue);

    var tpl = null, st = getBigTemplateStatus();
    for (var t = 0; t < st.templates.length; t++) if (st.templates[t].kind === MP_TEMPLATE_KIND_) tpl = st.templates[t];

    return { ok: true, members: members, blocks: cycle, candidates: cands,
             template: tpl, rowsPerPage: MP_ROWS_PER_OVERVIEW_,
             unusedCategories: unusedCategoryRows_(members).map(function (r) { return r.block; }) };
  } catch (e) {
    console.error('[MPRESEN] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '読み込みに失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// 誰も使っていない業種区分マスタの行。
// 昔の区分が残っていると、順番の一覧に人数0の行が並んで分かりにくくなる。
function unusedCategoryRows_(members) {
  var used = {}, i;
  for (i = 0; i < members.length; i++) if (members[i].cat) used[mpCatKey_(members[i].cat)] = true;
  var cats = getCategoryMaster(), out = [];
  for (i = 0; i < cats.length; i++) {
    if (!cats[i].key) continue;
    if (used[mpCatKey_(cats[i].key)]) continue;
    out.push({ key: cats[i].key, block: cats[i].block || cats[i].label || cats[i].key });
  }
  return out;
}

// 誰も使っていない業種区分をマスタから消す。
// apply が false のときは「何を消すか」を返すだけ（画面で確認してもらうため）。
function tidyCategoryMaster(apply) {
  try {
    var master = getMemberMaster(), members = master.members || [];
    var drop = unusedCategoryRows_(members);
    if (!drop.length) return { ok: true, message: '使われていない業種区分はありません。', removed: [] };
    var names = drop.map(function (r) { return r.block + (r.block === r.key ? '' : '（' + r.key + '）'); });
    if (!apply) {
      return { ok: true, preview: true, removed: names,
               message: '次の ' + drop.length + ' 件を消します:\n' + names.join('\n') };
    }
    var dropKey = {};
    for (var d = 0; d < drop.length; d++) dropKey[drop[d].key] = true;
    var cats = getCategoryMaster(), keep = [];
    for (var i = 0; i < cats.length; i++) if (!dropKey[cats[i].key]) keep.push(cats[i]);
    var res = saveCategoryMaster(keep);
    if (!res.ok) return res;
    console.log('[MPRESEN] 業種区分を整理: ' + names.join(' / '));
    return { ok: true, removed: names,
             message: '✅ 使われていない業種区分を ' + drop.length + ' 件消しました:\n' + names.join('\n') };
  } catch (e) {
    console.error('[MPRESEN] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '整理に失敗しました: ' + (e && e.message ? e.message : e) };
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
  xml = setParagraphsInShape_(xml, MP_INDIVIDUAL_.nameBox, [item.name || '']);

  // 会社名：行の分け方・文字の大きさ・枠の大きさは、すべて画面側で決めてある
  // （文字幅を実測できるのは画面側だけなので、寸法もそこで出した方が確実）。
  xml = setParagraphsInShape_(xml, MP_INDIVIDUAL_.company, item.companyLines || ['']);
  xml = setShapeGeomEmu_(xml, MP_INDIVIDUAL_.company,
    item.companyGeom || (item.companyTall ? MP_COMPANY_TALL_ : MP_COMPANY_DEFAULT_));
  if (item.companyPt) xml = setFontSizeInShape_(xml, MP_INDIVIDUAL_.company, item.companyPt);

  // カテゴリー：会社名の枠の真下に置く。
  // 会社名が3行になると、2行ぶんの決め打ちでは文字が重なってしまうため、
  // 画面側が実際の行数と文字の大きさから出した位置を使う。
  xml = setShapeGeomEmu_(xml, MP_INDIVIDUAL_.category,
    { x: MP_CATEGORY_X_, cx: MP_CATEGORY_CX_,
      y: item.categoryTop || (item.companyTall ? MP_CATEGORY_TOP_LOW_ : MP_CATEGORY_TOP_) });
  xml = setBodyAnchorInShape_(xml, MP_INDIVIDUAL_.category, 't');
  xml = noAutofitInShape_(xml, MP_INDIVIDUAL_.category);
  xml = setParagraphsInShape_(xml, MP_INDIVIDUAL_.category, item.categoryLines || ['']);
  if (item.categoryPt) xml = setFontSizeInShape_(xml, MP_INDIVIDUAL_.category, item.categoryPt);
  if (item.categoryTight) xml = setLineSpacingInShape_(xml, MP_INDIVIDUAL_.category, 85);

  if (item.autoAdvanceMs) xml = mpAutoAdvance_(xml, item.autoAdvanceMs);

  if (item.nextName) {
    xml = setParagraphsInShape_(xml, MP_INDIVIDUAL_.nextName, [item.nextName]);
  } else {
    xml = removeShape_(xml, MP_INDIVIDUAL_.nextName);
    xml = removeShape_(xml, MP_INDIVIDUAL_.nextLabel);
  }
  return mpApplyPhoto_(xml, tplRels, MP_INDIVIDUAL_, photo);
}

// 30秒カウントダウンを「スライドが出たら自動で始まり、終わったら次のスライドへ」にする。
//
// テンプレートのカウントダウンは、数字の図形を1秒ごとに1枚ずつ消していく仕掛けで、
// 開始条件が「クリック待ち」(delay="indefinite") になっている。これだと
// 何秒で次に進めばよいかがPowerPointにも決められない。
// ここで開始条件を0秒にし、スライドの「自動で次へ進む」時間(advTm)を
// カウントダウンの長さに合わせる。
function mpAutoAdvance_(xml, ms) {
  // 開始条件（クリック待ち → すぐ開始）。dur="indefinite" には触らないこと。
  var before = xml;
  xml = xml.replace(/<p:cond delay="indefinite"\/><p:cond evt="onBegin" delay="0"><p:tn val="\d+"\/><\/p:cond>/,
                    '<p:cond delay="0"/>');
  if (xml === before) xml = xml.replace('<p:cond delay="indefinite"/>', '<p:cond delay="0"/>');

  // 自動で次へ進む時間
  if (/advTm="\d+"/.test(xml)) {
    xml = xml.replace(/advTm="\d+"/g, 'advTm="' + ms + '"');
  } else if (/<p:transition\b/.test(xml)) {
    xml = xml.replace(/<p:transition\b([^>]*?)(\/?)>/g, '<p:transition$1 advTm="' + ms + '"$2>');
  }
  return xml;
}

// テンプレートの中身を、指示どおりのスライドの並びに置き換える
// ひな形にする2枚を、スライド番号ではなく「載っている図形」で見つける。
// テンプレートを作り直すと何枚目かが変わるため、番号を決め打ちにしない。
function mpFindModelSlide_(map, ids) {
  var nums = [], p;
  for (p in map) {
    var m = p.match(/^ppt\/slides\/slide(\d+)\.xml$/);
    if (m) nums.push(parseInt(m[1], 10));
  }
  nums.sort(function (a, b) { return a - b; });
  for (var i = 0; i < nums.length; i++) {
    var path = 'ppt/slides/slide' + nums[i] + '.xml', xml = xmlOf_(map, path);
    if (!xml || missingShapeIds_(xml, ids).length) continue;
    var rels = xmlOf_(map, 'ppt/slides/_rels/slide' + nums[i] + '.xml.rels');
    if (rels) return { path: path, xml: xml, rels: rels };
  }
  return null;
}

function buildMemberPresenSlides_(map, items) {
  // 個人ページから先に探す。扉ページとIDが一部重なるため、
  // 個人ページにしか無いID（氏名・カテゴリー・NEXT）を先に当てる。
  var iv = mpFindModelSlide_(map, [MP_INDIVIDUAL_.photo, MP_INDIVIDUAL_.nameBox, MP_INDIVIDUAL_.company,
                                   MP_INDIVIDUAL_.category, MP_INDIVIDUAL_.nextName, MP_INDIVIDUAL_.nextLabel]);
  var ov = mpFindModelSlide_(map, [MP_OVERVIEW_.title, MP_OVERVIEW_.nextName,
                                   MP_OVERVIEW_.photo, MP_OVERVIEW_.table]);
  var lack = [];
  if (!ov) lack.push(MP_OVERVIEW_.name);
  if (!iv) lack.push(MP_INDIVIDUAL_.name);
  if (lack.length) {
    throw new Error('テンプレートに「' + lack.join('」「') + '」のひな形が見つかりません。'
      + '「⚙️ 設定 ＞ 大きなスライド」に登録したファイルをご確認ください。'
      + '（図形の並びが変わると見つけられなくなります）');
  }
  var ovXml = ov.xml, ovRels = ov.rels, ivXml = iv.xml, ivRels = iv.rels;

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

// === 公式ファイルから雛形を組み立てる（official_srv.js から呼ぶ）===
//
// 公式ファイルのページを、このシステムが読める形に直して雛形にする。
// ページは番号ではなく「載っている文字」で探す（公式ファイルが新しくなっても、なるべくそのまま使えるように）。
//
//   ビジター紹介・ゲスト紹介・代理紹介 … 「歓迎 本日のビジター」のページ（3人ぶんの 氏名・専門分野・招待者）
//                                         図形の番号を、slides_visitor_srv.js の SLIDE_BLOCKS_ に合わせる
//   ビジタープレゼン … 「ビジターの紹介」のページに、ウィークリープレゼンのカウントダウン（卵時計の動画つき）を足す
//   メンバープレゼン … 業種区分の表のページ（扉）と、ウィークリープレゼンのページ（個人）。
//                      図形の番号を member_presen_srv.js の MP_OVERVIEW_ / MP_INDIVIDUAL_ に合わせる
//   定例会（前半）  … 表紙から「メインプレゼンテーション」まで（見本のウィークリープレゼンのページは除く。
//                      メンバーのページは作るときに差し込む）。メンバーシップ委員会の表とメインプレゼンに差し込み口
//   定例会（後半）  … 「リファーラルと推薦のことば」から「締めの言葉」まで。リファーラル発表のひな形・
//                      推薦のことば（2人のページ）・書記兼会計の更新状況の表に作り直す
//
// 音の出るもの（動画・ベル・卵時計）は消さない。

var OFF_REL_ = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
var OFF_NUM_RE_ = /^\s*\d+(?::[0-5]\d)?\s*$/;

// --- 組み立ての入り口 ---
function offBuildTemplate_(zip, kind, opt) {
  var src = offSource_(zip), plan;
  if (kind === 'intro' || kind === 'guest' || kind === 'dairi') plan = offBuildIntro_(src, opt, kind);
  else if (kind === 'presen') plan = offBuildPresen_(src, opt);
  else if (kind === 'memberPresen') plan = offBuildMemberPresen_(src, opt);
  else if (kind === 'meetingFirst') plan = offBuildFirst_(src, opt);
  else if (kind === 'meetingSecond') plan = offBuildSecond_(src, opt);
  else throw new Error('雛形の種類が不正です: ' + kind);
  var map = offAssemble_(src, plan.slides);
  var media = offAddParts_(src, map);
  return { map: map, note: plan.note || '', media: media };
}

// 公式ファイルの「動画・画像以外」の部品を全部読む（全部で0.3MBほど）→ 組み立ての材料
function offSource_(zip) {
  var want = zip.names.filter(function (n) {
    return !/^ppt\/(media|embeddings)\//.test(n) && !/^ppt\/(changesInfos\/|revisionInfo\.xml$)/.test(n);
  });
  var parts = offZipRead_(zip, want), xml = {};
  var src = { zip: zip, parts: parts, cache: xml };
  src.order = slideOrder_(parts);
  return src;
}
function offXml_(src, path) {
  if (!(path in src.cache)) src.cache[path] = src.parts[path] ? src.parts[path].getDataAsString('UTF-8') : null;
  return src.cache[path];
}
// 並び順で i 番目（0から）のページ → { path, xml, rels }
function offPage_(src, path) {
  return { path: path, xml: offXml_(src, path), rels: offXml_(src, relsPathOf_(path)) || offEmptyRels_() };
}
function offEmptyRels_() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
       + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
}
// 文字（空白を除く）で、条件に合う最初のページを探す
function offFind_(src, label, test, from) {
  for (var i = from || 0; i < src.order.length; i++) {
    var xml = offXml_(src, src.order[i]);
    if (xml && test(slideText_(xml).replace(/[\s　]/g, ''), xml)) return i;
  }
  throw new Error('公式ファイルに「' + label + '」のページが見つかりませんでした。公式ファイルが新しくなって作りが変わった可能性があります。');
}

// --- 新しい pptx の部品一式を組み立てる ---
//   slides … [{ xml, rels, notesOf: 元のページのパス（ノートを持ってくる）, hidden }]
function offAssemble_(src, slides) {
  var map = {}, p;
  for (p in src.parts) {
    if (/^ppt\/(slides|notesSlides)\//.test(p)) continue;
    map[p] = src.parts[p];
  }
  var prs = xmlOf_(map, 'ppt/presentation.xml'), prsRels = xmlOf_(map, 'ppt/_rels/presentation.xml.rels');
  var ct = xmlOf_(map, '[Content_Types].xml');
  prsRels = prsRels.replace(/<Relationship\b[^>]*\/relationships\/slide"[^>]*\/>/g, '');
  var maxRid = 0, m, re = /Id="rId(\d+)"/g;
  while ((m = re.exec(prsRels)) !== null) maxRid = Math.max(maxRid, parseInt(m[1], 10));
  ct = ct.replace(/<Override\b[^>]*PartName="\/ppt\/(slides\/slide|notesSlides\/notesSlide)\d+\.xml"[^>]*\/>/g, '');
  var ids = '', relAdd = '', ctAdd = '', hidden = 0, notes = 0;
  for (var i = 0; i < slides.length; i++) {
    var s = slides[i], n = i + 1, path = 'ppt/slides/slide' + n + '.xml', rid = 'rId' + (maxRid + n);
    var rels = s.rels.replace(/<Relationship\b[^>]*\/notesSlide"[^>]*\/>/g, '');
    var xml = s.hidden === undefined ? s.xml : setSlideShow_(s.xml, !s.hidden);
    if (/<p:sld\b[^>]*\sshow="0"/.test(xml)) hidden++;
    var nt = s.notesOf ? offNotesOf_(src, s.notesOf) : null;
    if (nt) {
      var np = 'ppt/notesSlides/notesSlide' + n + '.xml';
      putXml_(map, np, nt.xml);
      putXml_(map, relsPathOf_(np), nt.rels.replace(/(<Relationship\b[^>]*\/relationships\/slide"[^>]*Target=")[^"]*(")/,
        '$1../slides/slide' + n + '.xml$2'));
      var nr = offRelsAdd_(rels, 'notesSlide', '../notesSlides/notesSlide' + n + '.xml');
      rels = nr.rels;
      ctAdd += '<Override PartName="/' + np + '" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>';
      notes++;
    }
    putXml_(map, path, xml);
    putXml_(map, relsPathOf_(path), rels);
    ids += '<p:sldId id="' + (255 + n) + '" r:id="' + rid + '"/>';
    relAdd += '<Relationship Id="' + rid + '" Type="' + OFF_REL_ + 'slide" Target="slides/slide' + n + '.xml"/>';
    ctAdd += '<Override PartName="/' + path + '" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>';
  }
  prs = prs.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>|<p:sldIdLst\/>/, '<p:sldIdLst>' + ids + '</p:sldIdLst>')
           .replace(/<p:custShowLst>[\s\S]*?<\/p:custShowLst>/, '')
           .replace(/<p:ext uri="\{521415D9-36F7-43E2-AB2F-B90AF26B5E84\}">[\s\S]*?<\/p:ext>/, '');   // セクション
  putXml_(map, 'ppt/presentation.xml', prs);
  putXml_(map, 'ppt/_rels/presentation.xml.rels', prsRels.replace('</Relationships>', relAdd + '</Relationships>'));
  putXml_(map, '[Content_Types].xml', ct.replace('</Types>', ctAdd + '</Types>'));
  // ファイルの情報：ページの数を合わせ、公式ファイルのページの題名の一覧は外す
  var app = xmlOf_(map, 'docProps/app.xml');
  if (app) {
    app = app.replace(/<Slides>\d+<\/Slides>/, '<Slides>' + slides.length + '</Slides>')
             .replace(/<Notes>\d+<\/Notes>/, '<Notes>' + notes + '</Notes>')
             .replace(/<HiddenSlides>\d+<\/HiddenSlides>/, '<HiddenSlides>' + hidden + '</HiddenSlides>')
             .replace(/<HeadingPairs>[\s\S]*?<\/HeadingPairs>/, '').replace(/<TitlesOfParts>[\s\S]*?<\/TitlesOfParts>/, '');
    putXml_(map, 'docProps/app.xml', app);
  }
  return map;
}

// 元のページのノート → { xml, rels }（無ければ null）
function offNotesOf_(src, slidePath) {
  var rels = offXml_(src, relsPathOf_(slidePath)) || '';
  var m = rels.match(/Target="\.\.\/notesSlides\/(notesSlide\d+\.xml)"/);
  if (!m) return null;
  var p = 'ppt/notesSlides/' + m[1];
  var xml = offXml_(src, p), nrels = offXml_(src, relsPathOf_(p));
  return xml && nrels ? { xml: xml, rels: nrels } : null;
}

// 関係ファイルがたどる部品のうち、まだ入れていないもの（画像・動画など）を公式ファイルから読んで足す
function offAddParts_(src, map) {
  var need = {}, p, m;
  for (p in map) {
    if (!/\.rels$/.test(p)) continue;
    var owner = p.replace(/_rels\/([^\/]+)\.rels$/, '$1'), dir = posixDir_(owner);
    if (p === '_rels/.rels') dir = '';
    var re = /<Relationship\b[^>]*>/g, xml = xmlOf_(map, p);
    while ((m = re.exec(xml)) !== null) {
      if (/TargetMode="External"/.test(m[0])) continue;
      var t = (m[0].match(/Target="([^"]+)"/) || [])[1];
      if (!t || t === 'NULL') continue;
      var abs = t.charAt(0) === '/' ? t.substring(1) : normPartPath_(dir, t);
      if (!map[abs]) need[abs] = true;
    }
  }
  var names = Object.keys(need).filter(function (n) { return !!src.zip.entries[n]; });
  var got = offZipRead_(src.zip, names);
  for (p in got) map[p] = got[p];
  return names.length;
}

// --- 図形の小さな道具（文字列のまま扱う。XMLを組み直さない）---

// ページの図形（いちばん外側だけ）→ [{ tag, id, name, xml, text, x, y, cx, cy }]
function offShapes_(xml) {
  var s = xml.indexOf('<p:spTree>'), e = s < 0 ? -1 : mpElemEnd_(xml, s), out = [];
  if (s < 0) return out;
  var i = xml.indexOf('>', s) + 1, end = e - '</p:spTree>'.length;
  while (i < end) {
    var lt = xml.indexOf('<', i);
    if (lt < 0 || lt >= end) break;
    var close = mpElemEnd_(xml, lt), seg = xml.substring(lt, close);
    var tag = (seg.match(/^<([\w:]+)/) || [])[1];
    if (/^p:(sp|pic|graphicFrame|grpSp|cxnSp)$/.test(tag)) {
      var nv = seg.match(/<p:cNvPr\b[^>]*\sid="(\d+)"[^>]*?\sname="([^"]*)"/) || seg.match(/<p:cNvPr\b[^>]*\sid="(\d+)"/) || [];
      var g = seg.match(/<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"\s*\/>\s*<a:ext\s+cx="(\d+)"\s+cy="(\d+)"/);
      out.push({ tag: tag, id: nv[1] || '', name: nv[2] || '', xml: seg, start: lt, end: close,
                 text: slideText_(seg).replace(/^\s+|\s+$/g, ''),
                 x: g ? +g[1] : 0, y: g ? +g[2] : 0, cx: g ? +g[3] : 0, cy: g ? +g[4] : 0 });
    }
    i = close;
  }
  return out;
}
// 番号で図形を探す（ページを書き換えたあとは位置がずれるので、使う直前に探し直すこと）
function offShapeById_(xml, id) {
  var list = offShapes_(xml);
  for (var i = 0; i < list.length; i++) if (String(list[i].id) === String(id)) return list[i];
  return null;
}
function offByText_(shapes, test) {
  return shapes.filter(function (s) { return test(s.text.replace(/[\s　]/g, ''), s); });
}
function offMaxId_(xml) {
  var max = 0, m, re = /<p:cNvPr\b[^>]*\sid="(\d+)"/g;
  while ((m = re.exec(xml)) !== null) max = Math.max(max, parseInt(m[1], 10));
  return max;
}
// 図形の番号を付け替える（アニメーションや線のつながりの参照もいっしょに）。map … { 旧: 新 }
function offRenumber_(xml, map) {
  return xml.replace(/(<p:cNvPr\b[^>]*?\sid=")(\d+)(")|(\bspid=")(\d+)(")|(<a:(?:stCxn|endCxn)\b[^>]*?\sid=")(\d+)(")/g,
    function (all, a1, a2, a3, b1, b2, b3, c1, c2, c3) {
      if (a1) return a1 + (map[a2] || a2) + a3;
      if (b1) return b1 + (map[b2] || b2) + b3;
      return c1 + (map[c2] || c2) + c3;
    });
}
function offSetIdName_(shapeXml, id, name) {
  return shapeXml.replace(/(<p:cNvPr\b[^>]*?\sid=")\d+(")/, '$1' + id + '$2')
                 .replace(/(<p:cNvPr\b[^>]*?\sname=")[^"]*(")/, '$1' + escapeXml_(name) + '$2')
                 .replace(/<a16:creationId\b[^>]*\/>/, '');
}
function offSetGeom_(shapeXml, g) {
  return shapeXml.replace(/<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"\s*\/>(\s*)<a:ext\s+cx="(\d+)"\s+cy="(\d+)"\s*\/>/,
    function (all, x, y, sp, cx, cy) {
      return '<a:off x="' + Math.round(g.x == null ? +x : g.x) + '" y="' + Math.round(g.y == null ? +y : g.y) + '"/>' + sp
           + '<a:ext cx="' + Math.round(g.cx == null ? +cx : g.cx) + '" cy="' + Math.round(g.cy == null ? +cy : g.cy) + '"/>';
    });
}
function offMove_(shapeXml, dx, dy) {
  return shapeXml.replace(/<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"\s*\/>/, function (all, x, y) {
    return '<a:off x="' + (+x + Math.round(dx)) + '" y="' + (+y + Math.round(dy)) + '"/>';
  });
}
// 図形の段落のうち k 番目（0から）だけを残す
function offKeepParagraph_(shapeXml, k) {
  var tb = findTagRanges_(shapeXml, 'p:txBody');
  if (!tb.length) return shapeXml;
  var body = shapeXml.substring(tb[0].start, tb[0].end), ps = findTagRanges_(body, 'a:p');
  if (!ps.length) return shapeXml;
  var keep = body.substring(ps[Math.min(k, ps.length - 1)].start, ps[Math.min(k, ps.length - 1)].end);
  body = body.substring(0, ps[0].start) + keep + body.substring(ps[ps.length - 1].end);
  return shapeXml.substring(0, tb[0].start) + body + shapeXml.substring(tb[0].end);
}
// 図形（または表のセル）の本文の段落を、文字の並びに置き換える。書式は1段落目のまま。
// 文字の無い段落（<a:endParaRPr> だけ）でも、その書式で文字を入れる
function offSetLines_(bodyOwnerXml, lines, tag) {
  var tb = findTagRanges_(bodyOwnerXml, tag || 'p:txBody');
  if (!tb.length) return bodyOwnerXml;
  var body = bodyOwnerXml.substring(tb[0].start, tb[0].end), ps = findTagRanges_(body, 'a:p');
  if (!ps.length) return bodyOwnerXml;
  var tpl = body.substring(ps[0].start, ps[0].end), out = '';
  for (var i = 0; i < lines.length; i++) out += offParagraph_(tpl, lines[i]);
  body = body.substring(0, ps[0].start) + out + body.substring(ps[ps.length - 1].end);
  return bodyOwnerXml.substring(0, tb[0].start) + body + bodyOwnerXml.substring(tb[0].end);
}
function offParagraph_(pXml, text) {
  if (findTagRanges_(pXml, 'a:r').length) return oneRunParagraph_(pXml, text);
  var end = pXml.match(/<a:endParaRPr\b([^>]*?)(\/>|>([\s\S]*?)<\/a:endParaRPr>)/);
  var rPr = end ? '<a:rPr' + end[1].replace(/\sdirty="\d"/, '') + (end[3] ? '>' + end[3] + '</a:rPr>' : '/>') : '<a:rPr lang="ja-JP"/>';
  var run = '<a:r>' + rPr + '<a:t>' + escapeXml_(text) + '</a:t></a:r>';
  return end ? pXml.replace(end[0], run + end[0]) : pXml.replace('</a:p>', run + '</a:p>');
}
// 文字の大きさ（pt）をそろえる（ラン・段落末・既定の書式すべて）
function offSetSize_(xml, pt) {
  return xml.replace(/(<a:(?:rPr|endParaRPr|defRPr)\b[^>]*?\ssz=")\d+(")/g, function (all, a, b) { return a + Math.round(pt * 100) + b; });
}
function offInsertShapes_(xml, shapesXml) { return xml.replace('</p:spTree>', shapesXml + '</p:spTree>'); }
function offRemoveShapes_(xml, ids) {
  for (var i = 0; i < ids.length; i++) xml = removeShape_(xml, ids[i]);
  return xml;
}
// 図形をまるごと差し替える（番号で探す）
function offReplaceShape_(xml, id, newXml) {
  var r = findShapeRange_(xml, id);
  return r ? xml.substring(0, r.start) + newXml + xml.substring(r.end) : xml;
}
// 関係を1つ足す → { rels, rid }
function offRelsAdd_(rels, type, target) {
  var max = 0, m, re = /Id="rId(\d+)"/g;
  while ((m = re.exec(rels)) !== null) max = Math.max(max, parseInt(m[1], 10));
  var rid = 'rId' + (max + 1);
  return { rid: rid, rels: rels.replace('</Relationships>', '<Relationship Id="' + rid + '" Type="' + OFF_REL_ + type
    + '" Target="' + target + '"/></Relationships>') };
}
function offRelTarget_(rels, rid) {
  var m = rels.match(new RegExp('<Relationship\\b[^>]*\\bId="' + rid + '"[^>]*>'));
  return m ? (m[0].match(/Target="([^"]+)"/) || [])[1] : '';
}
function offRelType_(rels, rid) {
  var m = rels.match(new RegExp('<Relationship\\b[^>]*\\bId="' + rid + '"[^>]*>'));
  return m ? (m[0].match(/Type="[^"]*\/([^"\/]+)"/) || [])[1] : '';
}
// ある関係IDを別の番号に付け替える（図形の r:embed などもいっしょに）
function offRelsRename_(xml, rels, from, to) {
  var mark = '__OFFRID__';
  var relsOut = rels.replace(new RegExp('\\bId="' + to + '"'), 'Id="' + mark + '"')
                    .replace(new RegExp('\\bId="' + from + '"'), 'Id="' + to + '"')
                    .replace(new RegExp('\\bId="' + mark + '"'), 'Id="' + from + '"');
  var re = function (id) { return new RegExp('(r:(?:embed|link|id|pict)=")' + id + '(")', 'g'); };
  var xmlOut = xml.replace(re(to), '$1' + mark + '$2').replace(re(from), '$1' + to + '$2').replace(re(mark), '$1' + from + '$2');
  return { xml: xmlOut, rels: relsOut };
}

// 文字の段落を1つの図形に分けたときの高さ（EMU）。Meiryo UI の行の高さ（1.22倍）＋上下の余白
function offLineEmu_(pt, lines) { return Math.round((lines || 1) * pt * 1.22 * 12700 + 91440); }
function offFirstPt_(xml, def) {
  var m = xml.match(/<a:rPr\b[^>]*\ssz="(\d+)"/) || xml.match(/<a:defRPr\b[^>]*\ssz="(\d+)"/) || xml.match(/<a:endParaRPr\b[^>]*\ssz="(\d+)"/);
  return m ? parseInt(m[1], 10) / 100 : def;
}

// 文字の段落が2つ以上ある図形を、段落ごとの図形に分ける。
//   specs … [{ id, name, lines（入れる文字。省略すると元のまま）, gap（前の段落との間 EMU）}]
// 上から順に、元の枠の中に積む。戻り値 { xml, boxes: [{ id, x, y, cx, cy }] }
function offSplitShape_(xml, shape, specs) {
  var out = '', y = shape.y, boxes = [];
  for (var k = 0; k < specs.length; k++) {
    var sp = offKeepParagraph_(shape.xml, k), pt = offFirstPt_(sp, 18);
    var lines = specs[k].lines || null;
    if (lines) sp = offSetLines_(sp, lines);
    var h = specs[k].cy || offLineEmu_(pt, lines ? lines.length : 1);
    y += specs[k].gap || 0;
    sp = offSetIdName_(sp, specs[k].id, specs[k].name);
    sp = offSetGeom_(sp, { x: shape.x, y: y, cx: shape.cx, cy: h });
    boxes.push({ id: specs[k].id, x: shape.x, y: y, cx: shape.cx, cy: h, pt: pt });
    y += h;
    out += sp;
  }
  return { xml: xml.substring(0, shape.start) + out + xml.substring(shape.end), boxes: boxes };
}

// ページの文字の一部を置き換える（段落ごとに連結して探す。書式は当たった最初のランのまま）
function offReplaceText_(xml, from, to) {
  var paras = findTagRanges_(xml, 'a:p');
  for (var p = paras.length - 1; p >= 0; p--) {
    var seg = xml.substring(paras[p].start, paras[p].end);
    var upd = replaceInParagraph_(seg, function (joined) {
      if (from instanceof RegExp) {
        var m = from.exec(joined);
        return m ? [{ start: m.index, end: m.index + m[0].length, value: to }] : null;
      }
      var i = joined.indexOf(from);
      return i < 0 ? null : [{ start: i, end: i + from.length, value: to }];
    });
    if (upd !== seg) xml = xml.substring(0, paras[p].start) + upd + xml.substring(paras[p].end);
  }
  return xml;
}

// --- カウントダウン（数字の箱と卵時計の動画）を別のページへ持ってくる ---
// weekly … ウィークリープレゼンのページ（offPage_）。to … { x, y }（数字の箱の左上。EMU）
// withEgg … 卵時計の動画（音つき）もいっしょに。卵時計が流れてからカウントダウンが始まる
// 戻り値 { xml, rels }。もとのページの動画・音声の登録（右上の動画など）は残す
function offTransplantCountdown_(page, weekly, to, withEgg) {
  var wx = weekly.xml, boxes = mpCountdownShapes_(wx);
  if (boxes.length < 5) throw new Error('公式ファイルのウィークリープレゼンのページに、カウントダウンの数字が見つかりません。');
  var egg = null, i;
  if (withEgg) {
    offShapes_(wx).forEach(function (s) { if (s.tag === 'p:pic' && /<a:videoFile\b/.test(s.xml)) egg = s; });
  }
  var g = readShapeGeomEmu_(wx, boxes[0].id), dx = to.x - g.x, dy = to.y - g.y;
  var next = offMaxId_(page.xml) + 1, map = {}, add = '', rels = page.rels;
  if (egg) map[egg.id] = String(next++);
  boxes.sort(function (a, b) { return a.start - b.start; });
  boxes.forEach(function (b) { map[b.id] = String(next++); });
  if (egg) {
    // 動画の関係（動画ファイル・p14:media・表紙の画像）を持ってくる
    var ex = offMove_(egg.xml, dx, dy), rids = [], m, re = /r:(?:embed|link)="(rId\d+)"/g;
    while ((m = re.exec(ex)) !== null) if (rids.indexOf(m[1]) < 0) rids.push(m[1]);
    rids.forEach(function (rid) {
      var r2 = offRelsAdd_(rels, offRelType_(weekly.rels, rid), offRelTarget_(weekly.rels, rid));
      rels = r2.rels;
      ex = ex.replace(new RegExp('(r:(?:embed|link)=")' + rid + '(")', 'g'), '$1' + r2.rid + '__$2');
    });
    add += offRenumber_(ex.replace(/(r:(?:embed|link)="rId\d+)__"/g, '$1"'), map);
  }
  boxes.forEach(function (b) { add += offRenumber_(offMove_(b.xml, dx, dy), map); });
  var xml = offInsertShapes_(page.xml, add), timing;
  if (egg) {
    timing = offRenumber_(wx.substring(wx.indexOf('<p:timing>'), wx.indexOf('</p:timing>') + 11), map);
  } else {
    var spids = boxes.slice().sort(function (a, b) { return b.num - a.num; })
      .filter(function (b) { return b.num > 0; }).map(function (b) { return map[b.id]; });
    timing = mpCountdownTiming_(spids, true);
  }
  var pt = xml.indexOf('<p:timing>');
  if (pt >= 0) {
    var old = xml.substring(pt, xml.indexOf('</p:timing>', pt) + 11), parsed = mpParseTiming_(old, []);
    if (parsed && parsed.others.length) {
      timing = timing.replace('</p:childTnLst></p:cTn></p:par></p:tnLst>',
        parsed.others.map(mpTokenIds_).join('') + '</p:childTnLst></p:cTn></p:par></p:tnLst>');
      timing = mpRenumberTiming_(timing, {});
    }
    xml = xml.substring(0, pt) + timing + xml.substring(xml.indexOf('</p:timing>', pt) + 11);
  } else if (xml.indexOf('</p:clrMapOvr>') >= 0) {
    xml = xml.replace('</p:clrMapOvr>', '</p:clrMapOvr>' + timing);
  } else {
    xml = xml.replace('</p:sld>', timing + '</p:sld>');
  }
  return { xml: xml, rels: rels };
}

// 関係IDを決まった番号にする（その番号をほかで使っていれば、そちらを空いている番号へ）→ { xml, rels }
function offRelsPut_(xml, rels, rid, type, target) {
  if (new RegExp('\\bId="' + rid + '"').test(rels)) {
    var moved = offRelsAdd_(rels, 'x', 'x').rid;                       // 空いている番号
    var r = offRelsRename_(xml, rels, rid, moved);
    xml = r.xml; rels = r.rels;
  }
  rels = rels.replace('</Relationships>', '<Relationship Id="' + rid + '" Type="' + OFF_REL_ + type + '" Target="' + target + '"/></Relationships>');
  return { xml: xml, rels: rels };
}

// 決まった番号と重なる図形を、空いている番号へ移す（決まった番号に付け替える前に呼ぶ）
function offFreeIds_(xml, ids, keep) {
  var shapes = offShapes_(xml), want = {}, next = Math.max(offMaxId_(xml), 99) + 1, set = {};
  ids.forEach(function (id) { set[String(id)] = true; });
  shapes.forEach(function (s) { if (set[s.id] && !(keep && keep[s.id])) want[s.id] = String(next++); });
  return offRenumber_(xml, want);
}

// === ビジター紹介・ゲスト紹介・代理紹介 ===
// 公式のページは「氏名」「専門分野：」「招待者：氏名（メンバー）」の3人ぶん。
// 上から順に、slides_visitor_srv.js の SLIDE_BLOCKS_ の番号（氏名・専門分野・招待者）に付け替える。
// 「専門分野：」「招待者：」の見出しは値と同じ枠にあり、作るときに残す（buildGroupXml_）。
function offBuildIntro_(src, opt, kind) {
  var i = offFind_(src, '歓迎 本日のビジター', function (t) {
    return t.indexOf('本日のビジター') >= 0 && (t.match(/氏名/g) || []).length >= 3 && t.indexOf('専門分野') >= 0;
  });
  var pg = offPage_(src, src.order[i]), xml = pg.xml, shapes = offShapes_(xml);
  var byY = function (a, b) { return a.y - b.y || a.x - b.x; };
  var names = offByText_(shapes, function (t) { return t === '氏名'; }).sort(byY);
  var cats = offByText_(shapes, function (t) { return /^専門分野/.test(t); }).sort(byY);
  var invs = offByText_(shapes, function (t) { return /^招待者/.test(t); }).sort(byY);
  var title = offByText_(shapes, function (t) { return t.indexOf('本日のビジター') >= 0; })[0];
  if (names.length < 3 || cats.length < 3 || invs.length < 3 || !title) {
    throw new Error('公式ファイルの「歓迎 本日のビジター」のページの作りが変わっていて、3人ぶんの枠を見分けられませんでした。');
  }
  // 決まった番号（氏名・専門分野・招待者・見出し・題名）と重なるほかの図形を、先に空いている番号へ
  var fixed = [INTRO_TITLE_SHAPE_ID_], want = {}, keep = {};
  SLIDE_BLOCKS_.forEach(function (b) { fixed = fixed.concat([b.name, b.category, b.inviter], b.labels || []); });
  [title].concat(names.slice(0, 3), cats.slice(0, 3), invs.slice(0, 3)).forEach(function (s) { keep[s.id] = true; });
  xml = offFreeIds_(xml, fixed, keep);
  want[title.id] = String(INTRO_TITLE_SHAPE_ID_);
  for (var k = 0; k < 3; k++) {
    want[names[k].id] = String(SLIDE_BLOCKS_[k].name);
    want[cats[k].id] = String(SLIDE_BLOCKS_[k].category);
    want[invs[k].id] = String(SLIDE_BLOCKS_[k].inviter);
  }
  xml = offRenumber_(xml, want);
  if (kind === 'guest') xml = offReplaceText_(xml, 'ビジター', 'ゲスト');
  if (kind === 'dairi') xml = offReplaceText_(offReplaceText_(xml, 'ビジター', '代理出席の方々'), '歓迎', '');
  return { slides: [{ xml: xml, rels: pg.rels, notesOf: pg.path }] };
}

// ウィークリープレゼン（カウントダウンと卵時計のある、個人のページ）の見本
function offWeeklyIndex_(src) {
  return offFind_(src, 'ウィークリープレゼンテーション（次のプレゼンター・カウントダウン）', function (t, xml) {
    return t.indexOf('ウィークリープレゼンテーション') >= 0 && t.indexOf('次のプレゼンター') >= 0 && mpCountdownShapes_(xml).length >= 5;
  });
}
function offSecondsLabel_(sec) { return chapterSecondsLabel_(sec); }

// === ビジタープレゼン ===
// 「ビジターの紹介」のページ（お名前・会社名・事業内容）に、ウィークリープレゼンのカウントダウン
// （卵時計の動画つき）を足す。図形の番号は TEMPLATE_KINDS_.presen（25 氏名・27 会社名・29 カテゴリー）
function offBuildPresen_(src, opt) {
  var i = offFind_(src, 'ビジターの紹介（お名前・会社名・事業内容）', function (t) {
    return t.indexOf('ビジターの紹介') >= 0 && t.indexOf('お名前') >= 0 && t.indexOf('会社名') >= 0;
  });
  var pg = offPage_(src, src.order[i]), weekly = offPage_(src, src.order[offWeeklyIndex_(src)]);
  var xml = offFreeIds_(pg.xml, [25, 27, 29]), shapes = offShapes_(xml);
  var nameBox = offByText_(shapes, function (t) { return t.indexOf('お名前') >= 0; })[0];
  var bizBox = offByText_(shapes, function (t) { return t === '事業内容'; })[0];
  var ask = offByText_(shapes, function (t) { return t.indexOf('お手伝い') >= 0; })[0];
  if (!nameBox || !bizBox) throw new Error('公式ファイルの「ビジターの紹介」のページの作りが変わっていて、お名前・事業内容の枠を見分けられませんでした。');
  var sp = offSplitShape_(xml, nameBox, [{ id: 25, name: 'Visitor Name', lines: ['お名前　様'] },
                                           { id: 27, name: 'Visitor Company', lines: ['会社名'] }]);
  xml = sp.xml;
  // 事業内容の枠を【カテゴリー】（専門分野）に
  var biz = offShapes_(xml).filter(function (s) { return s.id === bizBox.id; })[0];
  xml = offReplaceShape_(xml, biz.id, offSetIdName_(offSetLines_(biz.xml, ['【カテゴリー】']), 29, 'Visitor Category'));
  // カウントダウン：【カテゴリー】の下と「お手伝いできますか？」の間の、文字の列の真ん中
  var num = mpCountdownShapes_(weekly.xml)[0], ng = readShapeGeomEmu_(weekly.xml, num.id);
  var top = bizBox.y + bizBox.cy, bottom = ask ? ask.y : top + ng.cy + 254000;
  var center = nameBox.x + nameBox.cx / 2;
  var to = { x: Math.round(center - ng.cx / 2 + 0.4 * ng.cx), y: Math.round(top + Math.max(0, (bottom - top - ng.cy) / 2)) };
  var t = offTransplantCountdown_({ xml: xml, rels: pg.rels }, weekly, to, true);
  xml = mpSetCountdown_(t.xml, opt.seconds.visitor);
  return { slides: [{ xml: xml, rels: t.rels, notesOf: pg.path }],
           note: 'カウントダウン' + offSecondsLabel_(opt.seconds.visitor) + '・卵時計の音つき' };
}

// === メンバープレゼン ===
// 扉 … 業種区分の表のページ。題名（業種区分名）→ 11、表 → 6。左側に先頭の方の写真（2）と「次のプレゼンター」（31）を足す
// 個人 … ウィークリープレゼンのページ。写真 → 3、氏名 → 56、会社名 → 2、連絡先の枠 →【カテゴリー】12、
//        「次のプレゼンター:」→ 6、その下の氏名 → 14。数字の箱と卵時計は、空いている番号へ
function offBuildMemberPresen_(src, opt) {
  var w = offWeeklyIndex_(src), weekly = offPage_(src, src.order[w]);
  var o = offFind_(src, '業種区分の表（専門分野・メンバー）', function (t, xml) {
    return /<a:tbl>/.test(xml) && t.indexOf('専門分野') >= 0 && t.indexOf('メンバー') >= 0 && t.indexOf('氏名') >= 0
        && t.indexOf('委員会') < 0;
  });
  var ov = offPage_(src, src.order[o]);
  var iv = offMemberIndividual_(weekly, opt);
  var ovs = offMemberOverview_(ov, weekly, iv);
  return { slides: [{ xml: ovs.xml, rels: ovs.rels }, { xml: iv.xml, rels: iv.rels }],
           note: 'カウントダウン' + offSecondsLabel_(opt.seconds.weekly) + '・卵時計の音つき' };
}

function offMemberIndividual_(weekly, opt) {
  var fixed = [MP_INDIVIDUAL_.photo, MP_INDIVIDUAL_.nameBox, MP_INDIVIDUAL_.company, MP_INDIVIDUAL_.category,
               MP_INDIVIDUAL_.nextName, MP_INDIVIDUAL_.nextLabel];
  var xml = offFreeIds_(weekly.xml, fixed), rels = weekly.rels, shapes = offShapes_(xml);
  var nc = offByText_(shapes, function (t) { return t.indexOf('メンバー名') >= 0 && t.indexOf('会社名') >= 0 && t.indexOf('次の') < 0; })[0];
  var contact = offByText_(shapes, function (t) { return /^000-/.test(t) || t.indexOf('メールアドレス') >= 0; })[0];
  var next = offByText_(shapes, function (t) { return t.indexOf('次のプレゼンター') >= 0; })[0];
  var photo = shapes.filter(function (s) {
    return s.tag === 'p:pic' && !/<a:videoFile\b/.test(s.xml) && s.cy > 2500000;
  }).sort(function (a, b) { return b.cx * b.cy - a.cx * a.cy; })[0];
  if (!nc || !contact || !next || !photo) {
    throw new Error('公式ファイルのウィークリープレゼンのページの作りが変わっていて、氏名・会社名・次のプレゼンター・写真の枠を見分けられませんでした。');
  }
  var r = offSplitShape_(xml, nc, [{ id: MP_INDIVIDUAL_.nameBox, name: 'Member Name' },
                                   { id: MP_INDIVIDUAL_.company, name: 'Member Company' }]);
  xml = r.xml;
  var ct = offShapes_(xml).filter(function (s) { return s.id === contact.id; })[0];
  // 連絡先（3行）の枠を【カテゴリー】の1行に。高さは1行ぶん、会社名のすぐ下
  var cxml = offSetLines_(offKeepParagraph_(ct.xml, 0), ['【カテゴリー】']);
  var cpt = offFirstPt_(cxml, 24), comp = r.boxes[1];
  cxml = offSetGeom_(offSetIdName_(cxml, MP_INDIVIDUAL_.category, 'Member Category'),
    { x: ct.x, y: comp.y + comp.cy, cx: ct.cx, cy: offLineEmu_(cpt, 1) });
  xml = offReplaceShape_(xml, ct.id, cxml);
  var nx = offShapes_(xml).filter(function (s) { return s.id === next.id; })[0];
  xml = offSplitShape_(xml, nx, [{ id: MP_INDIVIDUAL_.nextLabel, name: 'Next Label' },
                                 { id: MP_INDIVIDUAL_.nextName, name: 'Next Name' }]).xml;
  xml = offRenumber_(xml, (function () { var m = {}; m[photo.id] = String(MP_INDIVIDUAL_.photo); return m; })());
  // 写真の関係IDを rId2 に（member_presen_srv.js がこの番号で差し替える）
  var prid = (photo.xml.match(/<a:blip\b[^>]*r:embed="(rId\d+)"/) || [])[1];
  if (prid && prid !== MP_INDIVIDUAL_.photoRid) {
    var rr = offRelsRename_(xml, rels, prid, MP_INDIVIDUAL_.photoRid);
    xml = rr.xml; rels = rr.rels;
  }
  xml = mpSetCountdown_(xml, opt.seconds.weekly);
  return { xml: xml, rels: rels, photoTarget: offRelTarget_(rels, MP_INDIVIDUAL_.photoRid), photoXml: photo.xml,
           nextXml: nx.xml };
}

// 表の2行目から下の書式を、2行目（1行目は見出し）と同じにする。
// 公式の見本は「募集中」の行だけ赤い字なので、そのままだとその行に入れたお名前が赤くなる
function offUniformRows_(xml, frameId) {
  var r = findShapeRange_(xml, frameId);
  if (!r) return xml;
  var seg = xml.substring(r.start, r.end), trs = findTagRanges_(seg, 'a:tr');
  if (trs.length < 3) return xml;
  var model = seg.substring(trs[1].start, trs[1].end).replace(/<a:extLst>(?:(?!<a:extLst>)[\s\S])*?<\/a:extLst>\s*<\/a:tr>$/, '</a:tr>');
  for (var i = trs.length - 1; i >= 2; i--) {
    var own = seg.substring(trs[i].start, trs[i].end), ext = own.match(/<a:extLst>(?:(?!<a:extLst>)[\s\S])*?<\/a:extLst>\s*<\/a:tr>$/);
    var row = ext ? model.replace(/<\/a:tr>$/, ext[0]) : model;
    seg = seg.substring(0, trs[i].start) + row + seg.substring(trs[i].end);
  }
  return xml.substring(0, r.start) + seg + xml.substring(r.end);
}

function offMemberOverview_(ov, weekly, iv) {
  var fixed = [MP_OVERVIEW_.title, MP_OVERVIEW_.nextName, MP_OVERVIEW_.photo, MP_OVERVIEW_.table];
  var xml = offFreeIds_(ov.xml, fixed.concat([32])), rels = ov.rels, shapes = offShapes_(xml);
  var table = shapes.filter(function (s) { return s.tag === 'p:graphicFrame' && /<a:tbl>/.test(s.xml); })[0];
  var title = shapes.filter(function (s) { return s.tag === 'p:sp' && s.text && s.id !== (table && table.id); })
    .sort(function (a, b) { return b.cx * b.cy - a.cx * a.cy; })[0];
  if (!table || !title) throw new Error('公式ファイルの業種区分の表のページの作りが変わっていて、表と区分名を見分けられませんでした。');
  var map = {};
  map[table.id] = String(MP_OVERVIEW_.table);
  map[title.id] = String(MP_OVERVIEW_.title);
  xml = offRenumber_(xml, map);
  xml = offUniformRows_(xml, MP_OVERVIEW_.table);
  // 左側：区分名を上へ、その下に先頭の方の写真、いちばん下に「次のプレゼンター」
  var colX = title.x, colW = title.cx, emu = 12700;
  xml = setShapeGeomEmu_(xml, MP_OVERVIEW_.title, { y: Math.round(36 * emu) });
  var side = Math.round(190 * emu), px = Math.round(colX + (colW - side) / 2);
  var pic = offSetGeom_(offSetIdName_(iv.photoXml, MP_OVERVIEW_.photo, 'First Member Photo'),
                        { x: px, y: Math.round(150 * emu), cx: side, cy: side });
  var pr = offRelsPut_(xml, rels, MP_OVERVIEW_.photoRid, 'image', iv.photoTarget);
  xml = pr.xml; rels = pr.rels;
  pic = pic.replace(/(<a:blip\b[^>]*r:embed=")rId\d+(")/, '$1' + MP_OVERVIEW_.photoRid + '$2');
  var nextShape = offShapes_('<p:sld><p:cSld><p:spTree><p:nvGrpSpPr/>' + iv.nextXml + '</p:spTree></p:cSld></p:sld>')[0];
  var label = offSetGeom_(offSetIdName_(offKeepParagraph_(nextShape.xml, 0), 32, 'Next Label'),
                          { x: colX, y: Math.round(360 * emu), cx: colW, cy: offLineEmu_(offFirstPt_(offKeepParagraph_(nextShape.xml, 0), 20), 1) });
  var nameP = offKeepParagraph_(nextShape.xml, 1);
  var name = offSetGeom_(offSetIdName_(offSetLines_(nameP, ['メンバー名']), MP_OVERVIEW_.nextName, 'Next Name'),
                         { x: colX, y: Math.round(392 * emu), cx: colW, cy: offLineEmu_(offFirstPt_(nameP, 24), 1) });
  xml = offInsertShapes_(xml, pic + label + name);
  return { xml: xml, rels: rels };
}

// === お2人が並ぶページ（メインプレゼン・推薦のことば）===
// page … 土台のページ（題名・右上の動画などはそのまま）。person … 写真と「メンバー名／会社名／…」の文字のあるページ
// 左右に「写真・氏名・会社名・カテゴリー」を並べ、{{○○1氏名}} などの差し込み口を置く
// （meeting_slides_srv.js の applyTwoPersonPhotos_・fitTwoPersonPage_ が、この形で写真と文字を入れる）
function offTwoPersonPage_(page, person, prefix, title, arrow) {
  var emu = 12700, xml = page.xml, rels = page.rels, shapes = offShapes_(xml);
  var ttl = shapes.filter(function (s) { return /title/i.test(s.name) || /<p:ph\b[^>]*type="(title|ctrTitle)"/.test(s.xml); })[0]
         || shapes.filter(function (s) { return s.tag === 'p:sp' && s.y < 1100000 && s.text; })[0];
  // 題名と右上の小さな動画（音つき）だけ残す
  var drop = shapes.filter(function (s) {
    if (ttl && s.id === ttl.id) return false;
    return !(s.tag === 'p:pic' && /<a:videoFile\b/.test(s.xml) && s.cy < 2000000);
  }).map(function (s) { return s.id; });
  xml = offRemoveShapes_(xml, drop);
  if (ttl && title) xml = offReplaceShape_(xml, ttl.id, offSetLines_(ttl.xml, [title]));
  var ps = offShapes_(person.xml);
  var photo = ps.filter(function (s) { return s.tag === 'p:pic' && !/<a:videoFile\b/.test(s.xml) && s.cy > 2500000; })
    .sort(function (a, b) { return b.cx * b.cy - a.cx * a.cy; })[0];
  var text = offByText_(ps, function (t) { return t.indexOf('メンバー名') >= 0 && t.indexOf('会社名') >= 0; })[0];
  if (!photo || !text) throw new Error('公式ファイルのメインプレゼンテーションのページの作りが変わっていて、写真と文字の枠を見分けられませんでした。');
  var prid = (photo.xml.match(/<a:blip\b[^>]*r:embed="(rId\d+)"/) || [])[1];
  var ra = offRelsAdd_(rels, 'image', offRelTarget_(person.rels, prid));
  rels = ra.rels;
  var id = offMaxId_(xml) + 1, add = '', side = 230 * emu, centers = [240 * emu, 720 * emu];
  var para = function (k, token, pt) { return offSetSize_(offSetLines_(offKeepParagraph_(text.xml, k), [token]), pt); };
  for (var n = 0; n < 2; n++) {
    var cx = centers[n], w = 440 * emu, x = Math.round(cx - w / 2);
    var pic = offSetGeom_(offSetIdName_(photo.xml, id++, prefix + (n + 1) + ' Photo'),
                          { x: Math.round(cx - side / 2), y: 92 * emu, cx: side, cy: side })
      .replace(/(<a:blip\b[^>]*r:embed=")rId\d+(")/, '$1' + ra.rid + '$2');
    add += pic;
    var rows = [['氏名', 0, 30, 332], ['会社名', 1, 20, 380], ['カテゴリー', 2, 18, 414]];
    for (var r = 0; r < rows.length; r++) {
      var sp = para(rows[r][1], '{{' + prefix + (n + 1) + rows[r][0] + '}}', rows[r][2]);
      add += offSetGeom_(offSetIdName_(sp, id++, prefix + (n + 1) + ' ' + rows[r][0]),
                         { x: x, y: rows[r][3] * emu, cx: w, cy: offLineEmu_(rows[r][2], 1) });
    }
  }
  if (arrow) {
    var ar = offSetSize_(offSetLines_(offKeepParagraph_(text.xml, 0), ['➡']), 54);
    add += offSetGeom_(offSetIdName_(ar, id++, 'Arrow'), { x: 420 * emu, y: 170 * emu, cx: 120 * emu, cy: offLineEmu_(54, 1) });
  }
  return { xml: offInsertShapes_(xml, add), rels: rels };
}

// === 定例会（前半）===
// 表紙から「メインプレゼンテーション」まで。見本のウィークリープレゼンのページ（業種区分の表・個人）は入れない
// （メンバーのページは、作るときにメンバープレゼンの雛形から差し込む）。
function offBuildFirst_(src, opt) {
  var cover = offFind_(src, '表紙（ようこそ！）', function (t) { return t.indexOf('ようこそ') >= 0 && t.indexOf('BNI') >= 0; });
  var head = offFind_(src, 'ウィークリープレゼンテーションの見出し', function (t) { return t === 'ウィークリープレゼンテーション'; }, cover);
  var done = offFind_(src, '全員終わりましたか？（ウィークリープレゼンテーション）', function (t) {
    return t.indexOf('終わりましたか') >= 0 && t.indexOf('ウィークリー') >= 0;
  }, head);
  var main = offFind_(src, 'メインプレゼンテーション', function (t) {
    return t.indexOf('メインプレゼンテーション') >= 0 && t.indexOf('メンバー名') >= 0;
  }, done);
  var weekly = offPage_(src, src.order[offWeeklyIndex_(src)]), idx = [], i;
  for (i = cover; i <= head; i++) idx.push(i);
  for (i = done; i < main; i++) idx.push(i);
  var slides = [], notes = [];
  idx.forEach(function (k) {
    var pg = offPage_(src, src.order[k]), t = slideText_(pg.xml).replace(/[\s　]/g, ''), xml = pg.xml, rels = pg.rels, hidden;
    if (k === cover) xml = offCover_(xml, opt);
    else if (t.indexOf('本日のビジター') >= 0 && (t.match(/氏名/g) || []).length >= 3) { hidden = true; notes.push('ビジター紹介のページは非表示'); }
    else if (t.indexOf('リーダーシップチーム') >= 0) { var lp = offPlaceholderPhotos_(xml, rels, weekly); xml = lp.xml; rels = lp.rels; }
    else if (t.indexOf('メンバーシップ委員会による報告') >= 0 && /<a:tbl>/.test(xml)) xml = offMembershipTokens_(xml);
    else if (t.indexOf('スピーカーローテーション') >= 0 && /<a:tbl>/.test(xml)) xml = offRotationArea_(xml);
    slides.push({ xml: xml, rels: rels, notesOf: pg.path, hidden: hidden });
  });
  var mp = offPage_(src, src.order[main]), two = offTwoPersonPage_(mp, mp, 'メインプレゼン', '', false);
  slides.push({ xml: two.xml, rels: two.rels, notesOf: mp.path });
  return { slides: slides, note: slides.length + '枚' };
}

// 表紙：「BNI ◯◯◯」→ チャプター名、「20〇〇年XX月XX日」→ 次回の「第N回 YYYY年MM月DD日」
// （作るときに、差し込み口の無い「第○回」「○年○月○日」を開催回・開催日に書き換える機能で直る）
function offCover_(xml, opt) {
  var d = opt.meetingDate, pad = function (n) { return (n < 10 ? '0' : '') + n; };
  var label = '第' + opt.meetingNo + '回　' + d.getFullYear() + '年' + pad(d.getMonth() + 1) + '月' + pad(d.getDate()) + '日';
  var before = xml;
  xml = offReplaceText_(xml, /20[0-9〇○◯OＯ]{2}\s*年\s*[0-9XＸxｘ]{1,2}\s*月\s*[0-9XＸxｘ]{1,2}\s*日/, label);
  if (xml !== before) {
    // 日付の枠は文字の長さに合わせて縮まない（折り返さない）ので、ページの幅いっぱいに広げて真ん中に置く
    offShapes_(xml).forEach(function (s) {
      if (s.text.indexOf(label.replace(/　/g, '')) >= 0 || s.text.indexOf(label) >= 0) {
        xml = setShapeGeomEmu_(xml, s.id, { x: 0, cx: 12192000 });
      }
    });
  }
  // 日付を直してから（日付の「〇〇」に当たらないように）
  return offReplaceText_(xml, /[◯○〇]{2,}/, opt.chapter + 'チャプター');
}

// リーダーシップチームの見本の人物写真を、写真を入れる枠（Replace With Photograph）に替える
function offPlaceholderPhotos_(xml, rels, weekly) {
  var wp = offShapes_(weekly.xml).filter(function (s) { return s.tag === 'p:pic' && !/<a:videoFile\b/.test(s.xml) && s.cy > 2500000; })[0];
  if (!wp) return { xml: xml, rels: rels };
  var target = offRelTarget_(weekly.rels, (wp.xml.match(/<a:blip\b[^>]*r:embed="(rId\d+)"/) || [])[1]);
  offShapes_(xml).forEach(function (s) {
    if (s.tag !== 'p:pic' || /<a:videoFile\b/.test(s.xml) || s.cy < 2000000) return;
    var rid = (s.xml.match(/<a:blip\b[^>]*r:embed="(rId\d+)"/) || [])[1];
    if (!rid) return;
    rels = retargetRel_(rels, rid, target);
    // 見本の写真に合わせた切り抜きは外す（枠の画像が欠けないように）
    xml = offReplaceShape_(xml, s.id, s.xml.replace(/<a:srcRect\b[^>]*\/>/, '<a:srcRect/>'));
  });
  return { xml: xml, rels: rels };
}

// スピーカーローテーション：作るときに、見本の表の場所へ5回ぶんの表を入れる（speaker_rotation_srv.js）。
// 見本の表の場所は小さいので、題名の下いっぱいに広げておく
function offRotationArea_(xml) {
  var emu = 12700;
  offShapes_(xml).forEach(function (s) {
    if (s.tag !== 'p:graphicFrame' || !/<a:tbl>/.test(s.xml)) return;
    xml = offReplaceShape_(xml, s.id, s.xml.replace(
      /<p:xfrm>\s*<a:off\s+x="-?\d+"\s+y="-?\d+"\s*\/>\s*<a:ext\s+cx="\d+"\s+cy="\d+"\s*\/>\s*<\/p:xfrm>/,
      '<p:xfrm><a:off x="' + 90 * emu + '" y="' + 92 * emu + '"/><a:ext cx="' + 780 * emu + '" cy="' + 372 * emu + '"/></p:xfrm>'));
  });
  return xml;
}

// メンバーシップ委員会の表：「チャプターが求める専門分野」の行を左右2つずつのマス目にして
// {{求める専門分野1}}〜{{求める専門分野12}}（画面が渡すのは12まで。それより下の行は空）、
// 「審査中の申し込み」の1行目に {{審査中カテゴリー}}。見本の「First & last name」は消す
var OFF_WANTED_MAX_ = 12;
function offMembershipTokens_(xml) {
  offShapes_(xml).forEach(function (s) {
    if (s.tag !== 'p:graphicFrame' || !/<a:tbl>/.test(s.xml)) return;
    var t = s.text.replace(/[\s　]/g, ''), seg = s.xml, trs = findTagRanges_(seg, 'a:tr'), out = seg, r;
    var wanted = t.indexOf('求める専門分野') >= 0, review = t.indexOf('審査中') >= 0;
    if (!wanted && !review) return;
    for (r = trs.length - 1; r >= 1; r--) {
      var tr = seg.substring(trs[r].start, trs[r].end), tcs = findTagRanges_(tr, 'a:tc');
      if (!tcs.length) continue;
      var row;
      if (wanted) {
        var first = tr.substring(tcs[0].start, tcs[0].end).replace(/^<a:tc\b[^>]*>/, '<a:tc>');
        // 2つに分けたマス目は狭いので、文字を小さく（太字もやめる）
        var small = offSetSize_(first, 14).replace(/(<a:(?:rPr|endParaRPr)\b[^>]*?)\sb="1"/g, '$1');
        var tok = function (n) { return n <= OFF_WANTED_MAX_ ? '{{求める専門分野' + n + '}}' : ''; };
        row = tr.substring(0, tcs[0].start) + offSetLines_(small, [tok(2 * r - 1)], 'a:txBody')
            + offSetLines_(small, [tok(2 * r)], 'a:txBody') + tr.substring(tcs[tcs.length - 1].end);
      } else {
        var tc = tr.substring(tcs[0].start, tcs[0].end);
        row = tr.substring(0, tcs[0].start) + offSetLines_(offSetSize_(tc, 16), [r === 1 ? '{{審査中カテゴリー}}' : ''], 'a:txBody')
            + tr.substring(tcs[0].end);
      }
      out = out.substring(0, trs[r].start) + row + out.substring(trs[r].end);
    }
    if (out !== seg) xml = xml.replace(seg, out);
  });
  return xml;
}

// === 定例会（後半）===
// 「リファーラルと推薦のことば」の見出しから「締めの言葉」まで。
function offBuildSecond_(src, opt) {
  var head = offFind_(src, 'リファーラルと推薦のことば（見出し）', function (t) { return t === 'リファーラルと推薦のことば'; });
  var closing = offFind_(src, '締めの言葉', function (t) { return t.indexOf('締めの言葉') >= 0; }, head);
  var weekly = offPage_(src, src.order[offWeeklyIndex_(src)]);
  var main = offPage_(src, src.order[offFind_(src, 'メインプレゼンテーション', function (t) {
    return t.indexOf('メインプレゼンテーション') >= 0 && t.indexOf('メンバー名') >= 0;
  })]);
  var slides = [], made = [];
  for (var k = head; k <= closing; k++) {
    var pg = offPage_(src, src.order[k]), t = slideText_(pg.xml).replace(/[\s　]/g, ''), r = null;
    if (t.indexOf('次の発表者') >= 0 && t.indexOf('メンバー名') >= 0) { r = offReferralPage_(pg, weekly, opt); made.push('リファーラル発表'); }
    else if (t.indexOf('推薦のことば') >= 0 && /メンバー[0-9０-９]/.test(t)) { r = offTwoPersonPage_(pg, main, '推薦のことば', '推薦のことば', true); made.push('推薦のことば'); }
    else if (t.indexOf('更新を迎えるメンバー') >= 0) { r = offRenewalTables_(pg); made.push('更新状況の表'); }
    slides.push({ xml: r ? r.xml : pg.xml, rels: r ? r.rels : pg.rels, notesOf: pg.path });
  }
  return { slides: slides, note: slides.length + '枚（' + made.join('・') + '）' };
}

// リファーラル発表のひな形：氏名・会社名・【カテゴリー】（幅の広い文字の枠を上から3つ）、
// 下の段に「次の発表者:」と氏名、写真、カウントダウン（リファーラル発表の秒数。クリックで始まる）。
// referral_srv.js の presenterShapes_ がこの作りで図形を見分ける。題名の図形の名前 REFERRAL PRESENTATION が目印
function offReferralPage_(pg, weekly, opt) {
  var emu = 12700, xml = pg.xml, rels = pg.rels, shapes = offShapes_(xml);
  var ttl = shapes.filter(function (s) { return /title/i.test(s.name); })[0];
  var nc = offByText_(shapes, function (t) { return t.indexOf('メンバー名') >= 0 && t.indexOf('会社名') >= 0 && t.indexOf('次の') < 0; })[0];
  var next = offByText_(shapes, function (t) { return t.indexOf('次の発表者') >= 0; })[0];
  var ask = offByText_(shapes, function (t) { return t.indexOf('貢献') >= 0; })[0];
  if (!ttl || !nc || !next) throw new Error('公式ファイルのリファーラルのページの作りが変わっていて、氏名・次の発表者の枠を見分けられませんでした。');
  xml = offReplaceShape_(xml, ttl.id, ttl.xml.replace(/(<p:cNvPr\b[^>]*?\sname=")[^"]*(")/, '$1' + RF_TITLE_ + '$2'));
  var id = Math.max(offMaxId_(xml), 99) + 1;
  var sp = offSplitShape_(xml, offShapeById_(xml, nc.id), [{ id: id++, name: 'Referral Name' }, { id: id++, name: 'Referral Company' }]);
  xml = sp.xml;
  var comp = sp.boxes[1], cpXml = offShapeById_(xml, comp.id).xml;
  var catPt = 24, cat = offSetSize_(offSetLines_(cpXml, ['【カテゴリー】']), catPt);
  cat = offSetGeom_(offSetIdName_(cat, id++, 'Referral Category'), { x: comp.x, y: comp.y + comp.cy, cx: comp.cx, cy: offLineEmu_(catPt, 1) });
  xml = offInsertShapes_(xml, cat);
  var catBottom = comp.y + comp.cy + offLineEmu_(catPt, 1);
  if (ask) {
    // 「今週のポジティブな貢献は？」は幅を少し狭め（氏名などの枠と見分けるため）、下の段の少し上へ
    // （会社名やカテゴリーが2〜3行になっても重ならない高さ）
    var w = Math.min(ask.cx, 5900000);
    xml = setShapeGeomEmu_(xml, ask.id, { x: Math.round(comp.x + (comp.cx - w) / 2), y: Math.max(catBottom + 6 * emu, 404 * emu), cx: w });
  }
  // 「次の発表者:」と氏名を、下の段（上から472pt より下）へ
  var nsp = offSplitShape_(xml, offShapeById_(xml, next.id), [{ id: id++, name: 'Next Label' }, { id: id++, name: 'Next Name' }]);
  xml = nsp.xml;
  xml = setShapeGeomEmu_(xml, nsp.boxes[0].id, { y: 474 * emu });
  xml = setShapeGeomEmu_(xml, nsp.boxes[1].id, { y: 474 * emu + nsp.boxes[0].cy });
  // カウントダウン（卵時計なし）：写真の下
  var photo = offShapes_(xml).filter(function (s) { return s.tag === 'p:pic' && !/<a:videoFile\b/.test(s.xml) && s.cy > 2500000; })[0];
  var ng = readShapeGeomEmu_(weekly.xml, mpCountdownShapes_(weekly.xml)[0].id);
  var to = photo ? { x: Math.round(photo.x + photo.cx / 2 - ng.cx / 2), y: photo.y + photo.cy + 6 * emu }
                 : { x: ng.x, y: ng.y };
  var t = offTransplantCountdown_({ xml: xml, rels: rels }, weekly, to, false);
  return { xml: mpSetCountdown_(t.xml, opt.seconds.referral, true), rels: t.rels };
}

// 書記兼会計による報告：1つの表（60日以内…・更新日と氏名の7行）を、
// 「90日以内」「60日以内」「30日以内」「期限切れ」の4つの表（見出しとお名前の2行）に作り直す。
// meeting_slides_srv.js の applyRenewalStatus_ が、見出しで表を見分けて2行目にお名前を入れる
function offRenewalTables_(pg) {
  var xml = pg.xml, shapes = offShapes_(xml);
  var tb = shapes.filter(function (s) { return s.tag === 'p:graphicFrame' && /<a:tbl>/.test(s.xml); })[0];
  if (!tb) return { xml: xml, rels: pg.rels };
  var seg = tb.xml, trs = findTagRanges_(seg, 'a:tr');
  var headTr = seg.substring(trs[0].start, trs[0].end), bodyTr = seg.substring(trs[1].start, trs[1].end);
  var headTc = headTr.substring(findTagRanges_(headTr, 'a:tc')[0].start, findTagRanges_(headTr, 'a:tc')[0].end).replace(/^<a:tc\b[^>]*>/, '<a:tc>');
  var bodyTc = bodyTr.substring(findTagRanges_(bodyTr, 'a:tc')[0].start, findTagRanges_(bodyTr, 'a:tc')[0].end).replace(/^<a:tc\b[^>]*>/, '<a:tc>');
  var trOpen = function (tr, h) { return tr.substring(0, tr.indexOf('>') + 1).replace(/\sh="\d+"/, ' h="' + h + '"'); };
  var labels = ['90日以内に更新を迎えるメンバー', '60日以内に更新を迎えるメンバー', '30日以内に更新を迎えるメンバー', '更新期限切れのメンバー'];
  var n = labels.length, gap = 60000, each = Math.floor((tb.cy - gap * (n - 1)) / n), hh = Math.round(each * 0.4), bh = each - hh;
  var out = '', next = Math.max(offMaxId_(xml), 99) + 1;
  for (var i = 0; i < n; i++) {
    var f = seg.replace(/<a:tblGrid>[\s\S]*?<\/a:tblGrid>/, '<a:tblGrid><a:gridCol w="' + tb.cx + '"/></a:tblGrid>');
    var rows = trOpen(headTr, hh) + offSetLines_(headTc, [labels[i]], 'a:txBody') + '</a:tr>'
             + trOpen(bodyTr, bh) + offSetLines_(bodyTc, ['該当者なし'], 'a:txBody') + '</a:tr>';
    var tr0 = findTagRanges_(f, 'a:tr');
    f = f.substring(0, tr0[0].start) + rows + f.substring(tr0[tr0.length - 1].end);
    f = offSetIdName_(f, i === 0 ? tb.id : next++, 'Renewal ' + (i + 1))
      .replace(/<p:nvPr>\s*<p:extLst>[\s\S]*?<\/p:extLst>\s*<\/p:nvPr>/, '<p:nvPr/>');
    f = f.replace(/<p:xfrm>\s*<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"\s*\/>\s*<a:ext\s+cx="(\d+)"\s+cy="(\d+)"\s*\/>\s*<\/p:xfrm>/,
      '<p:xfrm><a:off x="$1" y="' + (tb.y + i * (each + gap)) + '"/><a:ext cx="$3" cy="' + each + '"/></p:xfrm>');
    out += f;
  }
  return { xml: xml.substring(0, tb.start) + out + xml.substring(tb.end), rels: pg.rels };
}

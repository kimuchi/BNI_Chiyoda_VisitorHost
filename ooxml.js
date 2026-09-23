// === OOXML(pptx) 共通ユーティリティ（サーバー側） ===
// pptx は ZIP。Utilities.unzip/zip で展開・再梱包し、slide1.xml の文字だけを差し替えて
// 人数分のスライドに複製する。テンプレートのデザイン・座標・フォントには一切触れない。

// ZIP Blob → { パス: Blob } のマップ。ディレクトリエントリは必ず除外する
function unzipToMap_(zipBlob) {
  var blobs = Utilities.unzip(zipBlob.setContentType('application/zip')), map = {};
  for (var i = 0; i < blobs.length; i++) {
    var name = blobs[i].getName();
    if (!name || name.charAt(name.length - 1) === '/') continue;   // ディレクトリを捨てる
    map[name] = blobs[i];
  }
  return map;
}

// マップ → pptx Blob
function zipFromMap_(map, fileName) {
  var blobs = [];
  for (var path in map) {
    // copyBlob() はメモリを倍使うため、大きなpptxでは行わない。
    // map はこの後破棄するので、元のBlobに名前を付け直すだけでよい。
    blobs.push(map[path].setName(path));
  }
  return Utilities.zip(blobs, fileName)
    .setContentType('application/vnd.openxmlformats-officedocument.presentationml.presentation');
}

function xmlOf_(map, path) { return map[path] ? map[path].getDataAsString('UTF-8') : null; }
function putXml_(map, path, xml) { map[path] = Utilities.newBlob(xml, 'application/xml', path); }

// --- テキスト差し替え -------------------------------------------------
// 元ツール setTextInShape の忠実な移植。XMLの再シリアライズを避けるため文字列操作で行う。
// 指定シェイプID(p:cNvPr id)を含む <p:sp>…</p:sp> を見つけ、
// 最初の <a:r> の最初の <a:t> を newText にし、2本目以降の <a:r> を削除する。
function setTextInShape_(xml, shapeId, newText) {
  var spRanges = findSpRanges_(xml);
  for (var i = 0; i < spRanges.length; i++) {
    var sp = xml.substring(spRanges[i].start, spRanges[i].end);
    var m = sp.match(/<p:cNvPr[^>]*\sid="(\d+)"/);
    if (!m || m[1] !== String(shapeId)) continue;
    var runs = findTagRanges_(sp, 'a:r');
    if (!runs.length) continue;
    var updated = replaceFirstT_(sp.substring(runs[0].start, runs[0].end), newText);
    // 後ろのランから順に削除し、先頭ランを差し替える（位置ずれを防ぐため降順）
    var out = sp;
    for (var j = runs.length - 1; j >= 1; j--) out = out.substring(0, runs[j].start) + out.substring(runs[j].end);
    out = out.substring(0, runs[0].start) + updated + out.substring(runs[0].end);
    return xml.substring(0, spRanges[i].start) + out + xml.substring(spRanges[i].end);
  }
  return xml;   // 見つからなければ何もしない（元ツールと同じく黙って無視）
}

// カテゴリー【】専用。ランが3本以上なら '【' '値' '】' を個別に、少なければ1本にまとめる
function setCategoryInShape_(xml, shapeId, category) {
  var spRanges = findSpRanges_(xml);
  for (var i = 0; i < spRanges.length; i++) {
    var sp = xml.substring(spRanges[i].start, spRanges[i].end);
    var m = sp.match(/<p:cNvPr[^>]*\sid="(\d+)"/);
    if (!m || m[1] !== String(shapeId)) continue;
    var runs = findTagRanges_(sp, 'a:r'), out = sp;
    if (runs.length >= 3) {
      var texts = ['【', category, '】'];
      for (var j = 2; j >= 0; j--) {
        var r = sp.substring(runs[j].start, runs[j].end);
        out = out.substring(0, runs[j].start) + replaceFirstT_(r, texts[j]) + out.substring(runs[j].end);
      }
    } else if (runs.length >= 1) {
      out = out.substring(0, runs[0].start) +
            replaceFirstT_(sp.substring(runs[0].start, runs[0].end), '【' + category + '】') +
            out.substring(runs[0].end);
    } else continue;
    return xml.substring(0, spRanges[i].start) + out + xml.substring(spRanges[i].end);
  }
  return xml;
}

// === 長い文字を枠に収める ===
// 会社名やカテゴリーが長いと、枠からはみ出して下の行に重なってしまう。
// テンプレートから枠の幅と元の文字サイズを読み取り、文字数に応じて小さくする。
// 幅をハードコードせずテンプレートから読むので、デザインを差し替えても付いていける。

// シェイプの内側の幅(pt)・高さ(pt)・元の文字サイズ(pt)を読み取る
function readShapeTextBox_(xml, shapeId) {
  var spRanges = findSpRanges_(xml);
  for (var i = 0; i < spRanges.length; i++) {
    var sp = xml.substring(spRanges[i].start, spRanges[i].end);
    var m = sp.match(/<p:cNvPr[^>]*\sid="(\d+)"/);
    if (!m || m[1] !== String(shapeId)) continue;
    var ext = sp.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"/);
    if (!ext) return null;
    // 左右の余白。既定は 91440 EMU（0.1インチ）
    var lIns = 91440, rIns = 91440, bp = sp.match(/<a:bodyPr[^>]*>/);
    if (bp) {
      var l = bp[0].match(/\slIns="(-?\d+)"/), r = bp[0].match(/\srIns="(-?\d+)"/);
      if (l) lIns = parseInt(l[1], 10);
      if (r) rIns = parseInt(r[1], 10);
    }
    var sz = sp.match(/<a:(?:rPr|defRPr)[^>]*\ssz="(\d+)"/);
    return { widthPt: (parseInt(ext[1], 10) - lIns - rIns) / 12700,   // 12700 EMU = 1pt
             heightPt: parseInt(ext[2], 10) / 12700,
             basePt: sz ? parseInt(sz[1], 10) / 100 : 0 };
  }
  return null;
}

// シェイプ内の文字の大きさをまとめて変える
function setFontSizeInShape_(xml, shapeId, sizePt) {
  var spRanges = findSpRanges_(xml), v = Math.round(sizePt * 100);
  for (var i = 0; i < spRanges.length; i++) {
    var sp = xml.substring(spRanges[i].start, spRanges[i].end);
    var m = sp.match(/<p:cNvPr[^>]*\sid="(\d+)"/);
    if (!m || m[1] !== String(shapeId)) continue;
    var out = sp.replace(/<a:(rPr|endParaRPr|defRPr)\b([^>]*?)(\/?)>/g, function (all, tag, attrs, selfClose) {
      return '<a:' + tag + ' sz="' + v + '"' + attrs.replace(/\ssz="\d+"/g, '') + selfClose + '>';
    });
    return xml.substring(0, spRanges[i].start) + out + xml.substring(spRanges[i].end);
  }
  return xml;
}

// 文字数から、1行に収まる大きさを見積もって適用する。
// 全角は1文字ぶん、半角は0.5文字ぶんとして幅を数える。
// 元の大きさで収まるならそのまま。minPt より小さくはしない（読めなくなるため、
// そこまで長い場合は折り返して2行になる）。
function fitFontToShape_(xml, shapeId, text, minPt) {
  var box = readShapeTextBox_(xml, shapeId);
  if (!box || !box.basePt || box.widthPt <= 0) return xml;
  var w = textWidthUnits_(text);
  if (w <= 0) return xml;
  var size = Math.floor(box.widthPt / w);
  if (size >= box.basePt) return xml;
  return setFontSizeInShape_(xml, shapeId, Math.max(size, minPt || 10));
}

// スライド上の全シェイプの位置と大きさ（pt）を集める。
// p:sp のほか、画像(p:pic)・表など(p:graphicFrame)・グループ(p:grpSp)も見る。
function collectShapeBoxes_(xml) {
  var tags = ['p:sp', 'p:pic', 'p:graphicFrame', 'p:grpSp'], boxes = [];
  for (var t = 0; t < tags.length; t++) {
    var ranges = findTagRanges_(xml, tags[t]);
    for (var i = 0; i < ranges.length; i++) {
      var seg = xml.substring(ranges[i].start, ranges[i].end);
      var off = seg.match(/<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"\s*\/>/);
      var ext = seg.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"\s*\/>/);
      if (!off || !ext) continue;
      var id = seg.match(/<p:cNvPr[^>]*\sid="(\d+)"/);
      boxes.push({ id: id ? id[1] : '',
                   x: parseInt(off[1], 10) / 12700, y: parseInt(off[2], 10) / 12700,
                   w: parseInt(ext[1], 10) / 12700, h: parseInt(ext[2], 10) / 12700 });
    }
  }
  return boxes;
}

// 指定シェイプの下にどれだけ余裕があるか（pt）。
// 横方向に重なっていて、下にある一番近いシェイプまでの距離を返す。
// 下に何も無ければ fallbackPt を返す（スライドの高さは slide XML からは分からないため）。
function roomBelowShape_(xml, shapeId, fallbackPt) {
  var boxes = collectShapeBoxes_(xml), me = null, i;
  for (i = 0; i < boxes.length; i++) if (boxes[i].id === String(shapeId)) { me = boxes[i]; break; }
  if (!me) return 0;
  var myBottom = me.y + me.h, nearest = -1;
  for (i = 0; i < boxes.length; i++) {
    var b = boxes[i];
    if (b.id === String(shapeId)) continue;
    if (b.y < myBottom) continue;                              // 下にない
    if (b.x + b.w <= me.x || b.x >= me.x + me.w) continue;     // 横に重なっていない
    if (nearest < 0 || b.y < nearest) nearest = b.y;
  }
  return nearest < 0 ? (fallbackPt || 72) : Math.max(0, nearest - myBottom);
}

// シェイプを下へずらす（pt）
function moveShapeDown_(xml, shapeId, deltaPt) {
  if (!deltaPt || deltaPt <= 0) return xml;
  var spRanges = findSpRanges_(xml), delta = Math.round(deltaPt * 12700);
  for (var i = 0; i < spRanges.length; i++) {
    var sp = xml.substring(spRanges[i].start, spRanges[i].end);
    var m = sp.match(/<p:cNvPr[^>]*\sid="(\d+)"/);
    if (!m || m[1] !== String(shapeId)) continue;
    var replaced = false;
    var out = sp.replace(/<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"\s*\/>/, function (all, x, y) {
      if (replaced) return all;
      replaced = true;
      return '<a:off x="' + x + '" y="' + (parseInt(y, 10) + delta) + '"/>';
    });
    return xml.substring(0, spRanges[i].start) + out + xml.substring(spRanges[i].end);
  }
  return xml;
}

// 文字の幅を「全角何文字ぶん」で数える
function textWidthUnits_(text) {
  var s = String(text == null ? '' : text), w = 0;
  for (var i = 0; i < s.length; i++) w += s.charCodeAt(i) < 128 ? 0.5 : 1;
  return w;
}

// 長い文字を、1行に縮めるか2行にするかを決める。
// 2行にした方が大きく出せるなら2行にし、そのぶん下のシェイプをずらす。
// 日本語は行送りが大きいので、1行の高さは文字サイズの1.4倍で見積もる。
//
//   shapeId   … 縮める対象（会社名など）
//   belowId   … 2行になったときに下へずらすシェイプ（【カテゴリー】など）
//   minPt     … これ以上は小さくしない
//
// 戻り値は差し替え後のXML。
var LINE_HEIGHT_ = 1.4;
// 2行にするのは、1行のときより文字がこれだけ大きくできる場合だけ。
// わずかな差のために下のシェイプまで動かすと、かえってレイアウトが崩れて見える。
var TWO_LINE_GAIN_ = 1.4;

function fitTextAndPush_(xml, shapeId, belowId, text, minPt) {
  var box = readShapeTextBox_(xml, shapeId);
  if (!box || !box.basePt || box.widthPt <= 0) return xml;
  var w = textWidthUnits_(text);
  if (w <= 0) return xml;

  var one = Math.floor(box.widthPt / w);
  if (one >= box.basePt) return xml;            // 元の大きさで1行に収まる

  var keepOneLine = function () {
    return setFontSizeInShape_(xml, shapeId, Math.max(one, minPt || 10));
  };
  if (!belowId) return keepOneLine();

  // 2行にしたときに使える大きさ
  var two = Math.min(box.basePt, Math.floor(2 * box.widthPt / w));
  if (two < one * TWO_LINE_GAIN_) return keepOneLine();

  // 2行にすると、元の1行ぶんよりどれだけ縦に伸びるか
  var room = roomBelowShape_(xml, belowId, 72);
  var growOf = function (sz) { return Math.ceil((2 * sz - box.basePt) * LINE_HEIGHT_); };
  if (growOf(two) > room) {
    // ずらせる範囲に収まる大きさまで落とす
    two = Math.min(two, Math.floor((room / LINE_HEIGHT_ + box.basePt) / 2));
    if (two < one * TWO_LINE_GAIN_) return keepOneLine();   // 落とした結果、割に合わなくなった
  }

  xml = setFontSizeInShape_(xml, shapeId, two);
  return moveShapeDown_(xml, belowId, Math.max(0, growOf(two)));
}

function replaceFirstT_(runXml, text) {
  var esc = escapeXml_(text);
  // <a:t>…</a:t> / <a:t/>
  if (/<a:t[^>]*\/>/.test(runXml)) return runXml.replace(/<a:t[^>]*\/>/, '<a:t>' + esc + '</a:t>');
  return runXml.replace(/(<a:t[^>]*>)[\s\S]*?(<\/a:t>)/, '$1' + esc + '$2');
}

function escapeXml_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

// <p:sp> … </p:sp> の範囲一覧（入れ子は p:sp 同士では起きない前提。自己終了も考慮）
function findSpRanges_(xml) { return findTagRanges_(xml, 'p:sp'); }

// 指定タグの [start,end) 範囲を、入れ子を数えながら列挙する
function findTagRanges_(xml, tag) {
  var ranges = [], openRe = new RegExp('<' + tag.replace(':', '\\:') + '(?=[\\s>/])', 'g'), m;
  while ((m = openRe.exec(xml)) !== null) {
    var start = m.index;
    var gt = xml.indexOf('>', start);
    if (gt === -1) break;
    if (xml.charAt(gt - 1) === '/') { ranges.push({ start: start, end: gt + 1 }); openRe.lastIndex = gt + 1; continue; }
    var depth = 1, pos = gt + 1;
    var scan = new RegExp('<(\\/?)' + tag.replace(':', '\\:') + '(?=[\\s>/])', 'g');
    scan.lastIndex = pos;
    var s2;
    while ((s2 = scan.exec(xml)) !== null) {
      var g2 = xml.indexOf('>', s2.index);
      if (g2 === -1) break;
      if (s2[1] === '/') { depth--; if (depth === 0) { pos = g2 + 1; break; } }
      else if (xml.charAt(g2 - 1) !== '/') depth++;
      scan.lastIndex = g2 + 1;
    }
    ranges.push({ start: start, end: pos });
    openRe.lastIndex = pos;
  }
  return ranges;
}

// --- スライドの複製 ---------------------------------------------------
// slide1 をひな形に、xmlList の各要素を slide2..N として追加する。
// presentation.xml / rels / [Content_Types].xml を元ツールと同じ方式で更新する。
function addSlidesToMap_(map, xmlList) {
  var prsXml  = xmlOf_(map, 'ppt/presentation.xml');
  var prsRels = xmlOf_(map, 'ppt/_rels/presentation.xml.rels');
  var ctXml   = xmlOf_(map, '[Content_Types].xml');
  var slide1Rels = xmlOf_(map, 'ppt/slides/_rels/slide1.xml.rels');
  var notes1Xml  = xmlOf_(map, 'ppt/notesSlides/notesSlide1.xml');
  var notes1Rels = xmlOf_(map, 'ppt/notesSlides/_rels/notesSlide1.xml.rels');

  var sldIds = [], rIds = [], m;
  var reS = /<p:sldId\s+id="(\d+)"/g;  while ((m = reS.exec(prsXml))  !== null) sldIds.push(parseInt(m[1], 10));
  var reR = /Id="rId(\d+)"/g;          while ((m = reR.exec(prsRels)) !== null) rIds.push(parseInt(m[1], 10));
  var maxSldId = sldIds.length ? Math.max.apply(null, sldIds) : 256;
  var maxRid   = rIds.length   ? Math.max.apply(null, rIds)   : 10;

  var sldIdEntries = '', relEntries = '';
  for (var i = 0; i < xmlList.length; i++) {
    var n = i + 2, newSldId = maxSldId + i + 1, newRid = maxRid + i + 1;
    putXml_(map, 'ppt/slides/slide' + n + '.xml', xmlList[i]);
    if (slide1Rels) putXml_(map, 'ppt/slides/_rels/slide' + n + '.xml.rels',
      slide1Rels.replace(/notesSlide1\.xml/g, 'notesSlide' + n + '.xml'));
    if (notes1Xml)  putXml_(map, 'ppt/notesSlides/notesSlide' + n + '.xml', notes1Xml);
    if (notes1Rels) putXml_(map, 'ppt/notesSlides/_rels/notesSlide' + n + '.xml.rels',
      notes1Rels.replace(/slides\/slide1\.xml/g, 'slides/slide' + n + '.xml'));

    sldIdEntries += '<p:sldId id="' + newSldId + '" r:id="rId' + newRid + '"/>';
    relEntries   += '<Relationship Id="rId' + newRid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide' + n + '.xml"/>';

    var parts = [['/ppt/slides/slide' + n + '.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml']];
    if (notes1Xml) parts.push(['/ppt/notesSlides/notesSlide' + n + '.xml', 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml']);
    for (var p = 0; p < parts.length; p++) {
      var entry = '<Override PartName="' + parts[p][0] + '" ContentType="' + parts[p][1] + '"/>';
      if (ctXml.indexOf(entry) === -1) ctXml = ctXml.replace('</Types>', entry + '</Types>');
    }
  }
  putXml_(map, 'ppt/presentation.xml', prsXml.replace('</p:sldIdLst>', sldIdEntries + '</p:sldIdLst>'));
  putXml_(map, 'ppt/_rels/presentation.xml.rels', prsRels.replace('</Relationships>', relEntries + '</Relationships>'));
  putXml_(map, '[Content_Types].xml', ctXml);
  return map;
}

// テンプレートBlob + データ配列 → pptx Blob
function buildPptxFromTemplate_(templateBlob, dataList, builderFn, fileName) {
  var map = unzipToMap_(templateBlob);
  var slide1 = xmlOf_(map, 'ppt/slides/slide1.xml');
  if (!slide1) throw new Error('テンプレートに ppt/slides/slide1.xml がありません。');
  putXml_(map, 'ppt/slides/slide1.xml', builderFn(slide1, dataList[0]));
  var rest = [];
  for (var i = 1; i < dataList.length; i++) rest.push(builderFn(slide1, dataList[i]));
  addSlidesToMap_(map, rest);
  return zipFromMap_(map, fileName);
}

// --- {{トークン}} の置換 -------------------------------------------------
// PowerPointは1つの文字列を複数の <a:t> に分割して保持することがあるため
// （例: 「{{開催」「回}}」）、段落 <a:p> 単位で全 <a:t> を連結してから探す。
// 置換後は、最初のランに文字を入れ、またがった残りのランを空にする。
function replaceTokensInXml_(xml, map) {
  var paras = findTagRanges_(xml, 'a:p');
  // 後ろの段落から処理して位置ずれを防ぐ
  for (var p = paras.length - 1; p >= 0; p--) {
    var seg = xml.substring(paras[p].start, paras[p].end);
    var updated = replaceTokensInParagraph_(seg, map);
    if (updated !== seg) xml = xml.substring(0, paras[p].start) + updated + xml.substring(paras[p].end);
  }
  return xml;
}

function replaceTokensInParagraph_(seg, map) {
  var ts = findTagRanges_(seg, 'a:t');
  if (!ts.length) return seg;
  var texts = [], spans = [], joined = '';
  for (var i = 0; i < ts.length; i++) {
    var inner = seg.substring(ts[i].start, ts[i].end);
    var m = inner.match(/<a:t[^>]*>([\s\S]*?)<\/a:t>/);
    var t = m ? unescapeXml_(m[1]) : '';
    texts.push(t);
    spans.push({ from: joined.length, to: joined.length + t.length });
    joined += t;
  }
  if (joined.indexOf('{{') === -1) return seg;

  var hits = [], re = /\{\{([^{}]{1,60})\}\}/g, mm;
  while ((mm = re.exec(joined)) !== null) {
    if (Object.prototype.hasOwnProperty.call(map, mm[1])) {
      hits.push({ start: mm.index, end: mm.index + mm[0].length, value: map[mm[1]] == null ? '' : String(map[mm[1]]) });
    }
  }
  if (!hits.length) return seg;

  // 連結文字列の上で後ろから置換し、各ランの新しい文字を決める
  for (var h = hits.length - 1; h >= 0; h--) {
    var hit = hits[h];
    var a = runIndexAt_(spans, hit.start), b = runIndexAt_(spans, hit.end - 1);
    if (a < 0 || b < 0) continue;
    var head = texts[a].substring(0, hit.start - spans[a].from);
    var tail = texts[b].substring(hit.end - spans[b].from);
    texts[a] = head + hit.value + (a === b ? tail : '');
    for (var k = a + 1; k <= b; k++) texts[k] = (k === b && a !== b) ? tail : '';
  }

  // 各 <a:t> を書き戻す（後ろから）
  var out = seg;
  for (var j = ts.length - 1; j >= 0; j--) {
    var innerOld = seg.substring(ts[j].start, ts[j].end);
    var innerNew = innerOld.replace(/(<a:t[^>]*>)[\s\S]*?(<\/a:t>)/, '$1' + escapeXml_(texts[j]) + '$2');
    if (/<a:t[^>]*\/>/.test(innerOld)) innerNew = innerOld.replace(/<a:t[^>]*\/>/, '<a:t>' + escapeXml_(texts[j]) + '</a:t>');
    out = out.substring(0, ts[j].start) + innerNew + out.substring(ts[j].end);
  }
  return out;
}

function runIndexAt_(spans, pos) {
  for (var i = 0; i < spans.length; i++) if (pos >= spans[i].from && pos < spans[i].to) return i;
  return -1;
}

function unescapeXml_(s) {
  return String(s == null ? '' : s)
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

// テンプレートに含まれる {{トークン}} を一覧する（差し込み口の把握用）
function listTokensInXml_(xml) {
  var found = {}, paras = findTagRanges_(xml, 'a:p');
  for (var p = 0; p < paras.length; p++) {
    var seg = xml.substring(paras[p].start, paras[p].end);
    var ts = findTagRanges_(seg, 'a:t'), joined = '';
    for (var i = 0; i < ts.length; i++) {
      var m = seg.substring(ts[i].start, ts[i].end).match(/<a:t[^>]*>([\s\S]*?)<\/a:t>/);
      joined += m ? unescapeXml_(m[1]) : '';
    }
    var re = /\{\{([^{}]{1,60})\}\}/g, mm;
    while ((mm = re.exec(joined)) !== null) found[mm[1]] = (found[mm[1]] || 0) + 1;
  }
  return found;
}

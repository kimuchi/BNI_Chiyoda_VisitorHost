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
  return replaceInParagraph_(seg, function (joined) {
    if (joined.indexOf('{{') === -1) return null;
    var hits = [], re = /\{\{([^{}]{1,60})\}\}/g, mm;
    while ((mm = re.exec(joined)) !== null) {
      if (Object.prototype.hasOwnProperty.call(map, mm[1])) {
        hits.push({ start: mm.index, end: mm.index + mm[0].length,
                    value: map[mm[1]] == null ? '' : String(map[mm[1]]) });
      }
    }
    return hits;
  });
}

// 段落の <a:t> をすべて連結し、findHits(連結文字列) が返した範囲を差し替えて書き戻す。
// PowerPointは1つの文字列を複数のランに割ることがあるため、必ず連結してから探す。
function replaceInParagraph_(seg, findHits) {
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
  var hits = findHits(joined);
  if (!hits || !hits.length) return seg;

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

// --- 決まった形の文字の置換（差し込み口が無いページ向け）-----------------
// テンプレートによっては「第526回」「2026年07月22日」のように、{{ }} を置かずに
// そのまま書いてあるページがある。そういうページも更新できるよう、形で見つけて書き換える。
// rules は [{ re: 正規表現, value: function(一致) { return 置き換える文字; } }]。
// value が null を返した一致は、そのままにする。
function replacePatternsInXml_(xml, rules) {
  var paras = findTagRanges_(xml, 'a:p'), changed = 0;
  for (var p = paras.length - 1; p >= 0; p--) {
    var seg = xml.substring(paras[p].start, paras[p].end);
    var updated = replaceInParagraph_(seg, function (joined) {
      var hits = [], i, m;
      for (i = 0; i < rules.length; i++) {
        var re = new RegExp(rules[i].re.source, 'g');
        while ((m = re.exec(joined)) !== null) {
          var v = rules[i].value(m);
          if (v != null && v !== m[0]) hits.push({ start: m.index, end: m.index + m[0].length, value: String(v) });
          if (m.index === re.lastIndex) re.lastIndex++;      // 空一致で止まらないように
        }
      }
      hits.sort(function (a, b) { return a.start - b.start; });
      var out = [], last = -1;                                // 重なったら先に見つけた方を優先
      for (i = 0; i < hits.length; i++) if (hits[i].start >= last) { out.push(hits[i]); last = hits[i].end; }
      return out;
    });
    if (updated !== seg) {
      xml = xml.substring(0, paras[p].start) + updated + xml.substring(paras[p].end);
      changed++;
    }
  }
  return { xml: xml, changed: changed };
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

// --- 段落・座標・写真の細かい操作 -------------------------------------
// メンバープレゼンのように「1枚のひな形を人数分に増やす」場合は、文字だけでなく
// 段落の数・枠の位置・写真の切り抜きまで触る必要がある。
// どれも <p:cNvPr id="…"> で図形を特定し、その範囲の中だけを書き換える。
// 座標で図形を探す方法もあるが、同じテンプレートを使う前提ならIDの方が確実。

var SHAPE_TAGS_ = ['p:sp', 'p:pic', 'p:graphicFrame'];

// 図形ID から <p:sp>/<p:pic>/<p:graphicFrame> の範囲を得る
function findShapeRange_(xml, shapeId) {
  var want = String(shapeId);
  for (var t = 0; t < SHAPE_TAGS_.length; t++) {
    var ranges = findTagRanges_(xml, SHAPE_TAGS_[t]);
    for (var i = 0; i < ranges.length; i++) {
      var m = xml.substring(ranges[i].start, ranges[i].end).match(/<p:cNvPr[^>]*\sid="(\d+)"/);
      if (m && m[1] === want) return { start: ranges[i].start, end: ranges[i].end, tag: SHAPE_TAGS_[t] };
    }
  }
  return null;
}

// テンプレートに必要な図形が揃っているか確かめる（差し替えで壊れたときに早く気づくため）
function missingShapeIds_(xml, ids) {
  var miss = [];
  for (var i = 0; i < ids.length; i++) if (!findShapeRange_(xml, ids[i])) miss.push(ids[i]);
  return miss;
}

// 段落の雛形から「ランが1本だけの段落」を作る。書式は雛形のまま残る。
function oneRunParagraph_(pXml, text) {
  var runs = findTagRanges_(pXml, 'a:r');
  if (!runs.length) return pXml;                  // 文字を持たない段落は触らない
  var out = pXml;
  for (var j = runs.length - 1; j >= 1; j--) out = out.substring(0, runs[j].start) + out.substring(runs[j].end);
  return out.substring(0, runs[0].start)
       + replaceFirstT_(pXml.substring(runs[0].start, runs[0].end), text)
       + out.substring(runs[0].end);
}

// txBody の中身（<a:p> の並び）を、行ごとの段落に置き換える。
// 1段落目を雛形として複製するので、フォント・色・配置は元のまま。
function setTxBodyLines_(bodyXml, lines) {
  var ps = findTagRanges_(bodyXml, 'a:p');
  if (!ps.length) return bodyXml;
  var tpl = bodyXml.substring(ps[0].start, ps[0].end), out = '';
  var list = (lines && lines.length) ? lines : [''];
  for (var i = 0; i < list.length; i++) out += oneRunParagraph_(tpl, list[i]);
  return bodyXml.substring(0, ps[0].start) + out + bodyXml.substring(ps[ps.length - 1].end);
}

// 図形の本文を、行ごとの段落にして入れ直す
function setParagraphsInShape_(xml, shapeId, lines) {
  var r = findShapeRange_(xml, shapeId);
  if (!r) return xml;
  var seg = xml.substring(r.start, r.end), tb = findTagRanges_(seg, 'p:txBody');
  if (!tb.length) return xml;
  var body = setTxBodyLines_(seg.substring(tb[0].start, tb[0].end), lines);
  return xml.substring(0, r.start)
       + seg.substring(0, tb[0].start) + body + seg.substring(tb[0].end)
       + xml.substring(r.end);
}

// 図形の位置・大きさ（EMU）。指定した項目だけ変える
function setShapeGeomEmu_(xml, shapeId, geom) {
  var r = findShapeRange_(xml, shapeId);
  if (!r) return xml;
  var seg = xml.substring(r.start, r.end);
  seg = seg.replace(/<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"\s*\/>/, function (all, x, y) {
    return '<a:off x="' + (geom.x == null ? x : geom.x) + '" y="' + (geom.y == null ? y : geom.y) + '"/>';
  });
  seg = seg.replace(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"\s*\/>/, function (all, cx, cy) {
    return '<a:ext cx="' + (geom.cx == null ? cx : geom.cx) + '" cy="' + (geom.cy == null ? cy : geom.cy) + '"/>';
  });
  return xml.substring(0, r.start) + seg + xml.substring(r.end);
}

// 図形の位置・大きさ（EMU）を読む
function readShapeGeomEmu_(xml, shapeId) {
  var r = findShapeRange_(xml, shapeId);
  if (!r) return null;
  var seg = xml.substring(r.start, r.end);
  var off = seg.match(/<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"\s*\/>/);
  var ext = seg.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"\s*\/>/);
  if (!off || !ext) return null;
  return { x: parseInt(off[1], 10), y: parseInt(off[2], 10),
           cx: parseInt(ext[1], 10), cy: parseInt(ext[2], 10) };
}

// 文字の上下位置（t=上寄せ / ctr=中央 / b=下寄せ）
function setBodyAnchorInShape_(xml, shapeId, anchor) {
  var r = findShapeRange_(xml, shapeId);
  if (!r) return xml;
  var seg = xml.substring(r.start, r.end).replace(/<a:bodyPr\b([^>]*?)(\/?)>/,
    function (all, attrs, selfClose) {
      return '<a:bodyPr' + attrs.replace(/\sanchor="[^"]*"/g, '') + ' anchor="' + anchor + '"' + selfClose + '>';
    });
  return xml.substring(0, r.start) + seg + xml.substring(r.end);
}

// 「文字に合わせて枠を自動調整」をやめさせる。
// これが効いていると、こちらで決めた枠の高さをPowerPointが勝手に戻してしまう。
function noAutofitInShape_(xml, shapeId) {
  var r = findShapeRange_(xml, shapeId);
  if (!r) return xml;
  var seg = xml.substring(r.start, r.end), bp = findTagRanges_(seg, 'a:bodyPr');
  if (!bp.length) return xml;
  var body = seg.substring(bp[0].start, bp[0].end)
    .replace(/<a:spAutoFit\s*\/>/g, '')
    .replace(/<a:normAutofit\b[^>]*\/>/g, '')
    .replace(/<a:noAutofit\s*\/>/g, '');
  // <a:bodyPr> の子要素は順番が決まっている（prstTxWarp → autofit → …）ので、
  // prstTxWarp があればその直後、無ければ開始タグの直後に入れる。
  if (/\/>\s*$/.test(body)) {
    body = body.replace(/\/>\s*$/, '><a:noAutofit/></a:bodyPr>');
  } else if (/<\/a:prstTxWarp>/.test(body)) {
    body = body.replace('</a:prstTxWarp>', '</a:prstTxWarp><a:noAutofit/>');
  } else {
    body = body.replace(/^(<a:bodyPr\b[^>]*>)/, '$1<a:noAutofit/>');
  }
  var out = seg.substring(0, bp[0].start) + body + seg.substring(bp[0].end);
  return xml.substring(0, r.start) + out + xml.substring(r.end);
}

// 行間（％）。2行に折り返したときに詰めるために使う
function setLineSpacingInShape_(xml, shapeId, pct) {
  var r = findShapeRange_(xml, shapeId);
  if (!r) return xml;
  var lnSpc = '<a:lnSpc><a:spcPct val="' + Math.round(pct * 1000) + '"/></a:lnSpc>';
  var seg = xml.substring(r.start, r.end), ps = findTagRanges_(seg, 'a:p'), out = '';
  var prev = 0;
  for (var i = 0; i < ps.length; i++) {
    var p = seg.substring(ps[i].start, ps[i].end);
    if (/<a:lnSpc>[\s\S]*?<\/a:lnSpc>/.test(p)) p = p.replace(/<a:lnSpc>[\s\S]*?<\/a:lnSpc>/, lnSpc);
    else if (/<a:pPr\b[^>]*\/>/.test(p)) p = p.replace(/<a:pPr\b([^>]*?)\/>/, '<a:pPr$1>' + lnSpc + '</a:pPr>');
    else if (/<a:pPr\b[^>]*>/.test(p)) p = p.replace(/(<a:pPr\b[^>]*>)/, '$1' + lnSpc);
    else p = p.replace(/^(<a:p\b[^>]*>)/, '$1<a:pPr>' + lnSpc + '</a:pPr>');
    out += seg.substring(prev, ps[i].start) + p;
    prev = ps[i].end;
  }
  out += seg.substring(prev);
  return xml.substring(0, r.start) + out + xml.substring(r.end);
}

// 図形をまるごと消す（「次は○○さん」を最後の1人で消す用）
function removeShape_(xml, shapeId) {
  var r = findShapeRange_(xml, shapeId);
  return r ? (xml.substring(0, r.start) + xml.substring(r.end)) : xml;
}

// 写真の切り抜き（srcRect）。CSSの object-fit: cover と同じ考え方で、
// 縦横比を変えずに枠いっぱいに入るよう、はみ出す分を左右または上下から均等に切る。
function setSrcRectInPic_(xml, picId, crop) {
  var r = findShapeRange_(xml, picId);
  if (!r) return xml;
  var seg = xml.substring(r.start, r.end);
  var rect = '<a:srcRect l="' + Math.round(crop.l) + '" t="' + Math.round(crop.t)
           + '" r="' + Math.round(crop.r) + '" b="' + Math.round(crop.b) + '"/>';
  if (/<a:srcRect\b[^>]*\/>/.test(seg)) seg = seg.replace(/<a:srcRect\b[^>]*\/>/, rect);
  else if (/<a:blip\b[^>]*\/>/.test(seg)) seg = seg.replace(/(<a:blip\b[^>]*\/>)/, '$1' + rect);
  else if (/<\/a:blip>/.test(seg)) seg = seg.replace('</a:blip>', '</a:blip>' + rect);
  else return xml;
  return xml.substring(0, r.start) + seg + xml.substring(r.end);
}

// 枠(boxW×boxH)に srcW×srcH の画像を縦横比そのままで敷き詰めるときの切り抜き量。
// 単位は srcRect と同じ「10万分率」。
function coverCrop_(srcW, srcH, boxW, boxH) {
  var c = { l: 0, t: 0, r: 0, b: 0 };
  if (!srcW || !srcH || !boxW || !boxH) return c;
  var sr = srcW / srcH, tr = boxW / boxH;
  if (sr > tr) { c.l = c.r = (1 - tr / sr) / 2 * 100000; }
  else if (sr < tr) { c.t = c.b = (1 - sr / tr) / 2 * 100000; }
  return c;
}

// 表のセルに文字を入れる（行・列は0始まり）
function setTableCellText_(xml, frameId, rowIdx, colIdx, text) {
  var r = findShapeRange_(xml, frameId);
  if (!r) return xml;
  var seg = xml.substring(r.start, r.end), tbls = findTagRanges_(seg, 'a:tbl');
  if (!tbls.length) return xml;
  var tbl = seg.substring(tbls[0].start, tbls[0].end), trs = findTagRanges_(tbl, 'a:tr');
  if (rowIdx >= trs.length) return xml;
  var tr = tbl.substring(trs[rowIdx].start, trs[rowIdx].end), tcs = findTagRanges_(tr, 'a:tc');
  if (colIdx >= tcs.length) return xml;
  var tc = tr.substring(tcs[colIdx].start, tcs[colIdx].end), tb = findTagRanges_(tc, 'a:txBody');
  if (!tb.length) return xml;
  var body = setTxBodyLines_(tc.substring(tb[0].start, tb[0].end), [text]);
  tc  = tc.substring(0, tb[0].start) + body + tc.substring(tb[0].end);
  tr  = tr.substring(0, tcs[colIdx].start) + tc + tr.substring(tcs[colIdx].end);
  tbl = tbl.substring(0, trs[rowIdx].start) + tr + tbl.substring(trs[rowIdx].end);
  seg = seg.substring(0, tbls[0].start) + tbl + seg.substring(tbls[0].end);
  return xml.substring(0, r.start) + seg + xml.substring(r.end);
}

// 関係ファイル(.rels)の差し替え先を変える
function retargetRel_(relsXml, rId, newTarget) {
  return relsXml.replace(new RegExp('(<Relationship\\b[^>]*\\bId="' + rId + '"[^>]*\\bTarget=")[^"]*(")'),
                         '$1' + newTarget + '$2');
}

// 画像の幅・高さを、ファイル先頭のヘッダーだけから読む（PNG / JPEG）。
// GASの getBytes() は符号付きなので、必ず & 0xff してから使う。
function imageSizeOf_(bytes) {
  if (!bytes || bytes.length < 24) return null;
  var b = function (i) { return bytes[i] & 0xff; };
  // PNG: 8バイトの署名のあと IHDR。幅・高さは16〜23バイト目
  if (b(0) === 0x89 && b(1) === 0x50 && b(2) === 0x4E && b(3) === 0x47) {
    return { width:  (b(16) << 24) | (b(17) << 16) | (b(18) << 8) | b(19),
             height: (b(20) << 24) | (b(21) << 16) | (b(22) << 8) | b(23) };
  }
  // JPEG: SOI のあとマーカーを辿り、SOF（サイズが書かれたマーカー）から読む
  if (b(0) === 0xFF && b(1) === 0xD8) {
    var i = 2;
    while (i + 9 < bytes.length) {
      if (b(i) !== 0xFF) { i++; continue; }
      var mk = b(i + 1);
      if (mk === 0xFF) { i++; continue; }                                  // 詰め物
      if (mk === 0x01 || (mk >= 0xD0 && mk <= 0xD9)) { i += 2; continue; } // 中身の無いマーカー
      var len = (b(i + 2) << 8) | b(i + 3);
      if (len < 2) break;
      if (mk >= 0xC0 && mk <= 0xCF && mk !== 0xC4 && mk !== 0xC8 && mk !== 0xCC) {
        return { height: (b(i + 5) << 8) | b(i + 6), width: (b(i + 7) << 8) | b(i + 8) };
      }
      if (mk === 0xDA) break;                                              // 画像データに入った
      i += 2 + len;
    }
  }
  return null;
}

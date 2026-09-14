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
    var b = map[path];
    blobs.push(b.copyBlob ? b.copyBlob().setName(path) : b.setName(path));
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

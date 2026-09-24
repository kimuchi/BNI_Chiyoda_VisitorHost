// === 別のpptxのページを差し込む ===
//
// メンバープレゼンで作ったページを、前半スライドの「ウィークリープレゼンテーション」の
// ところへ差し込むために使う。2つのファイルは同じ土台（レイアウト「2_Blank」・
// テーマ「BNI Colors」）なので、ページ本体と画像を持ってきて、レイアウトを
// 差し込み先の同じ名前のものに付け替えれば、見た目は変わらない。
//
// 持ってくるもの … ページ本体・画像・動画・音声
// 付け替えるもの … レイアウト（名前で探す）
// 持ってこないもの … ノート（ページごとに1つなので複製できない）

var SPLICE_REL_ = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

// レイアウトの名前（<p:cSld name="…">）→ パス
function layoutsByName_(parts) {
  var out = {}, p;
  for (p in parts) {
    if (!/^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(p)) continue;
    var m = (xmlOf_(parts, p) || '').match(/<p:cSld name="([^"]*)"/);
    if (m && !out[m[1]]) out[m[1]] = p;
  }
  return out;
}

// source の slidePaths のページを、target の afterPath の直後に差し込む。
// 戻り値の paths は差し込んだページ（target 側のパス）。
function spliceSlides_(target, source, slidePaths, afterPath) {
  var tLayouts = layoutsByName_(target), sLayoutName = {}, p, m, i;
  for (p in source) {
    if (!/^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(p)) continue;
    m = (xmlOf_(source, p) || '').match(/<p:cSld name="([^"]*)"/);
    if (m) sLayoutName[p] = m[1];
  }

  var prs = xmlOf_(target, 'ppt/presentation.xml');
  var prsRels = xmlOf_(target, 'ppt/_rels/presentation.xml.rels');
  var ct = xmlOf_(target, '[Content_Types].xml');

  var maxSlide = 0, maxId = 255, maxRid = 0;
  for (p in target) {
    m = p.match(/^ppt\/slides\/slide(\d+)\.xml$/);
    if (m) maxSlide = Math.max(maxSlide, parseInt(m[1], 10));
  }
  var re = /<p:sldId id="(\d+)"/g;
  while ((m = re.exec(prs)) !== null) maxId = Math.max(maxId, parseInt(m[1], 10));
  re = /Id="rId(\d+)"/g;
  while ((m = re.exec(prsRels)) !== null) maxRid = Math.max(maxRid, parseInt(m[1], 10));

  var copied = {}, mediaSeq = 0, made = [], missingLayout = [];
  for (i = 0; i < slidePaths.length; i++) {
    var sp = slidePaths[i], xml = xmlOf_(source, sp);
    if (!xml) continue;
    var srels = xmlOf_(source, relsPathOf_(sp)) || '';
    var n = maxSlide + 1 + i, newPath = 'ppt/slides/slide' + n + '.xml';

    // 関係を1つずつ付け替える
    var outRels = srels.replace(/<Relationship\b[^>]*\/>/g, function (rel) {
      if (/TargetMode="External"/.test(rel)) return rel;
      var tgt = (rel.match(/Target="([^"]+)"/) || [])[1] || '';
      if (/notesSlides\//.test(tgt)) return '';                         // ノートは持ってこない
      var abs = normPartPath_(posixDir_(sp), tgt);
      if (/slideLayouts\//.test(tgt)) {
        var want = sLayoutName[abs], hit = want ? tLayouts[want] : null;
        if (!hit) { missingLayout.push(want || abs); hit = firstLayout_(target); }
        return rel.replace(/Target="[^"]+"/, 'Target="../slideLayouts/' + hit.replace(/^.*\//, '') + '"');
      }
      if (/media\//.test(tgt)) {
        if (!copied[abs]) {
          var ext = (abs.match(/\.[A-Za-z0-9]+$/) || ['.bin'])[0];
          var dst = 'ppt/media/spliced' + (++mediaSeq) + ext.toLowerCase();
          target[dst] = source[abs];
          if (target[dst] && target[dst].setName) target[dst].setName(dst);
          copied[abs] = dst;
          ct = ensureDefaultType_(ct, ext);
        }
        return rel.replace(/Target="[^"]+"/, 'Target="../media/' + copied[abs].replace('ppt/media/', '') + '"');
      }
      return rel;
    });

    putXml_(target, newPath, xml);
    putXml_(target, relsPathOf_(newPath), outRels);
    var rid = 'rId' + (maxRid + 1 + i);
    prsRels = prsRels.replace('</Relationships>', '<Relationship Id="' + rid + '" Type="' + SPLICE_REL_
      + 'slide" Target="slides/slide' + n + '.xml"/></Relationships>');
    ct = ct.replace('</Types>', '<Override PartName="/' + newPath + '"'
      + ' ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
    made.push({ id: maxId + 1 + i, rid: rid, path: newPath });
  }

  // 並びの中の差し込み位置（afterPath の直後。見つからなければ末尾）
  var rid2path = {}, reR = /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="slides\/(slide\d+\.xml)"[^>]*\/>/g;
  while ((m = reR.exec(prsRels)) !== null) rid2path[m[1]] = 'ppt/slides/' + m[2];
  var ins = '';
  for (i = 0; i < made.length; i++) ins += '<p:sldId id="' + made[i].id + '" r:id="' + made[i].rid + '"/>';
  var list = prs.match(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/)[1], out = '', placed = false;
  var reE = /<p:sldId id="(\d+)" r:id="([^"]+)"\/>/g;
  while ((m = reE.exec(list)) !== null) {
    out += m[0];
    if (!placed && rid2path[m[2]] === afterPath) { out += ins; placed = true; }
  }
  if (!placed) out += ins;
  prs = prs.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, '<p:sldIdLst>' + out + '</p:sldIdLst>');

  putXml_(target, 'ppt/presentation.xml', prs);
  putXml_(target, 'ppt/_rels/presentation.xml.rels', prsRels);
  putXml_(target, '[Content_Types].xml', ct);
  return { paths: made.map(function (x) { return x.path; }), placed: placed,
           missingLayout: missingLayout };
}

function posixDir_(p) { return p.replace(/\/[^\/]*$/, ''); }
function normPartPath_(dir, rel) {
  var parts = (dir + '/' + rel).split('/'), out = [];
  for (var i = 0; i < parts.length; i++) {
    if (parts[i] === '..') out.pop();
    else if (parts[i] && parts[i] !== '.') out.push(parts[i]);
  }
  return out.join('/');
}
function firstLayout_(parts) {
  var p, best = null;
  for (p in parts) if (/^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(p)) { best = p; break; }
  return best || 'ppt/slideLayouts/slideLayout1.xml';
}
function ensureDefaultType_(ct, ext) {
  var e = String(ext).replace('.', '').toLowerCase();
  if (!e || new RegExp('Extension="' + e + '"', 'i').test(ct)) return ct;
  var mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif',
               mp3: 'audio/mpeg', wav: 'audio/wav', m4a: 'audio/mp4', mp4: 'video/mp4' }[e]
             || 'application/octet-stream';
  return ct.replace('<Default', '<Default Extension="' + e + '" ContentType="' + mime + '"/><Default');
}

// source の並び順のスライドパス一覧
function slideOrder_(parts) {
  var prs = xmlOf_(parts, 'ppt/presentation.xml') || '';
  var prsRels = xmlOf_(parts, 'ppt/_rels/presentation.xml.rels') || '';
  var rid2path = {}, m, re = /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="slides\/(slide\d+\.xml)"[^>]*\/>/g;
  while ((m = re.exec(prsRels)) !== null) rid2path[m[1]] = 'ppt/slides/' + m[2];
  var out = [], reS = /<p:sldId id="\d+" r:id="([^"]+)"\/>/g;
  while ((m = reS.exec(prs)) !== null) if (rid2path[m[1]]) out.push(rid2path[m[1]]);
  return out;
}

// 前半スライドで、メンバープレゼンを差し込む位置（「ウィークリープレゼンテーション」の見出しページ）。
// 「全員終わりましたか？」のページにも同じ言葉が入っているので、それより前の最初の1枚を採る。
function weeklyAnchor_(parts) {
  var order = slideOrder_(parts);
  for (var i = 0; i < order.length; i++) {
    var t = slideText_(xmlOf_(parts, order[i]) || '').replace(/[\s　]/g, '');
    if (t.indexOf('ウィークリー') >= 0 && t.indexOf('プレゼンテーション') >= 0
        && t.indexOf('終わりましたか') < 0) return order[i];
  }
  return null;
}

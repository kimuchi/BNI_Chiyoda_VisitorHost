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
          while (target[dst]) dst = 'ppt/media/spliced' + (++mediaSeq) + ext.toLowerCase();   // 上書きしない
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

// --- スライドの並びを組み直すための小さな道具 ---
// 並び（presentation.xml の sldIdLst）→ [{ id, rid, path }]
function slideEntries_(parts) {
  var prs = xmlOf_(parts, 'ppt/presentation.xml') || '';
  var prsRels = xmlOf_(parts, 'ppt/_rels/presentation.xml.rels') || '';
  var rid2path = {}, m, re = /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="slides\/(slide\d+\.xml)"[^>]*\/>/g;
  while ((m = re.exec(prsRels)) !== null) rid2path[m[1]] = 'ppt/slides/' + m[2];
  var out = [], reS = /<p:sldId id="(\d+)" r:id="([^"]+)"\/>/g;
  while ((m = reS.exec(prs)) !== null) out.push({ id: parseInt(m[1], 10), rid: m[2], path: rid2path[m[2]] || '' });
  return out;
}

// スライドを1枚足す（部品・関係・種類の登録まで。並びには setSlideEntries_ で入れる）
function addSlidePart_(parts, xml, rels) {
  var maxSlide = 0, maxRid = 0, p, m;
  for (p in parts) {
    m = p.match(/^ppt\/slides\/slide(\d+)\.xml$/);
    if (m) maxSlide = Math.max(maxSlide, parseInt(m[1], 10));
  }
  var n = maxSlide + 1, path = 'ppt/slides/slide' + n + '.xml';
  putXml_(parts, path, xml);
  putXml_(parts, relsPathOf_(path), rels);
  var prsRels = xmlOf_(parts, 'ppt/_rels/presentation.xml.rels'), re = /Id="rId(\d+)"/g;
  while ((m = re.exec(prsRels)) !== null) maxRid = Math.max(maxRid, parseInt(m[1], 10));
  var rid = 'rId' + (maxRid + 1);
  putXml_(parts, 'ppt/_rels/presentation.xml.rels', prsRels.replace('</Relationships>',
    '<Relationship Id="' + rid + '" Type="' + SPLICE_REL_ + 'slide" Target="slides/slide' + n + '.xml"/></Relationships>'));
  putXml_(parts, '[Content_Types].xml', xmlOf_(parts, '[Content_Types].xml').replace('</Types>',
    '<Override PartName="/' + path + '" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>'));
  return { id: 0, rid: rid, path: path };
}

// 並びを書き戻す。id が 0 のもの（足したスライド）には、空いている番号を振る
function setSlideEntries_(parts, entries) {
  var maxId = 255, i;
  for (i = 0; i < entries.length; i++) if (entries[i].id) maxId = Math.max(maxId, entries[i].id);
  var out = '';
  for (i = 0; i < entries.length; i++) {
    if (!entries[i].id) entries[i].id = ++maxId;
    out += '<p:sldId id="' + entries[i].id + '" r:id="' + entries[i].rid + '"/>';
  }
  var prs = xmlOf_(parts, 'ppt/presentation.xml');
  putXml_(parts, 'ppt/presentation.xml', prs.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, '<p:sldIdLst>' + out + '</p:sldIdLst>'));
}

// PowerPoint のセクション（presentation.xml の p14:sectionLst）を、いまのスライドの並びに合わせる（ページを並べ替えたあとに呼ぶ）。
// セクションは続いたページのまとまりなので、並びと食い違うと PowerPoint が直そうとする（並びが戻ることもある）。
//   ・どのセクションにも無いページ（写したページ）は、すぐ前のページのセクションへ（先頭なら最初のセクション）
//   ・同じセクションのページが離れたら、離れた方は、その場所のセクションへ
//   ・セクションの並びは最初のページの順。ページの無くなったセクションは、元のすぐ前のセクションのうしろに残す
// 並びに無いページの番号は除く。セクションが無ければ何もしない
function syncSections_(parts) {
  var prs = xmlOf_(parts, 'ppt/presentation.xml') || '';
  var lm = prs.match(/<(\w+):sectionLst\b[^>]*>[\s\S]*?<\/\1:sectionLst>/);
  if (!lm) return false;
  var ns = lm[1], whole = lm[0];
  var secs = (whole.match(new RegExp('<' + ns + ':section\\b[^>]*\\/>|<' + ns + ':section\\b[^>]*>[\\s\\S]*?<\\/' + ns + ':section>', 'g')) || [])
    .map(function (x) {
      var ids = [], m, re = new RegExp('<' + ns + ':sldId\\b[^>]*\\bid="(\\d+)"', 'g');
      while ((m = re.exec(x)) !== null) ids.push(m[1]);
      return { xml: x, ids: ids, got: [] };
    });
  if (!secs.length) return false;
  var home = {}, cur = -1, closed = {}, first = {};
  secs.forEach(function (s, i) { s.ids.forEach(function (id) { if (!(id in home)) home[id] = i; }); });
  slideEntries_(parts).forEach(function (e, pos) {
    var id = String(e.id), h = id in home ? home[id] : (cur >= 0 ? cur : 0);
    if (h !== cur) {
      if (closed[h]) h = cur;                         // 離れた → その場所のセクションへ
      else { if (cur >= 0) closed[cur] = true; cur = h; }
    }
    secs[h].got.push(id);
    if (!(h in first)) first[h] = pos;
  });
  var out = [];
  secs.forEach(function (s, i) { if (s.got.length) out.push(i); });
  out.sort(function (a, b) { return first[a] - first[b]; });
  secs.forEach(function (s, i) {
    if (s.got.length) return;
    var to = 0;
    for (var j = i - 1; j >= 0; j--) { var k = out.indexOf(j); if (k >= 0) { to = k + 1; break; } }
    out.splice(to, 0, i);
  });
  var list = function (ids) {
    return ids.length ? '<' + ns + ':sldIdLst>' + ids.map(function (id) { return '<' + ns + ':sldId id="' + id + '"/>'; }).join('')
                        + '</' + ns + ':sldIdLst>' : '<' + ns + ':sldIdLst/>';
  };
  var lstRe = new RegExp('<' + ns + ':sldIdLst\\s*\\/>|<' + ns + ':sldIdLst\\b[^>]*>[\\s\\S]*?<\\/' + ns + ':sldIdLst>');
  var body = out.map(function (i) {
    var s = secs[i], x = s.xml;
    if (/\/>$/.test(x) && !new RegExp('<\\/' + ns + ':section>$').test(x)) x = x.replace(/\s*\/>$/, '></' + ns + ':section>');
    return lstRe.test(x) ? x.replace(lstRe, function () { return list(s.got); })
                         : x.replace(/^(<[^>]*>)/, function (t) { return t + list(s.got); });
  }).join('');
  var next = whole.match(/^<[^>]*>/)[0] + body + '</' + ns + ':sectionLst>';
  putXml_(parts, 'ppt/presentation.xml', prs.replace(whole, function () { return next; }));
  return true;
}

// その文字（差し込み口など）が載っているスライド。並びの順で最初のもの
function findSlideWithText_(parts, text) {
  var order = slideOrder_(parts);
  for (var i = 0; i < order.length; i++) {
    var x = xmlOf_(parts, order[i]);
    if (x && (x.indexOf(text) >= 0 || slideText_(x).indexOf(text) >= 0)) return order[i];
  }
  return null;
}

// --- 別のpptxのページを、見た目（レイアウト・マスター・テーマ）ごと写す ---
// 熱烈歓迎のページ（別に登録したひな形。welcome_srv.js）を事前MTGのパワポに入れるときに使う。
// source のページのレイアウトと、その元のマスター・テーマが、target に同じ中身で無ければ、
// マスター（そのレイアウト・テーマ・画像も）ごと写す（PowerPoint の「元の書式を保持」で貼り付けるのと同じ）。
// 同じ中身があれば、それを使う（写さない。同じ土台から作ったひな形なら、何も足さずに済む）。
// ページは count 枚写す（同じ画像などは1回だけ写して使い回す）。ノートは写さない。並びには入れない（setSlideEntries_ で入れる）。
//   戻り値 { slides: [{ id: 0, rid, path }], importedMaster: 写したか, sizeDiffers: ページの大きさが違うか }
function importSlideCopies_(target, source, slidePath, count) {
  var ctx = { target: target, source: source, map: {} };
  var srels = xmlOf_(source, partRelsPath_(slidePath)) || '';
  var layoutAbs = relTargetOfType_(source, slidePath, 'slideLayout');
  var useLayout = layoutAbs ? sameLayoutIn_(target, source, layoutAbs) : '';
  var imported = false;
  if (layoutAbs && !useLayout) {
    var masterAbs = relTargetOfType_(source, layoutAbs, 'slideMaster');
    if (masterAbs) {
      var newMaster = copyPartDeep_(ctx, masterAbs);           // テーマ・レイアウト・画像も写る
      registerSlideMaster_(target, newMaster);
      useLayout = ctx.map[layoutAbs] || '';
      imported = true;
    }
  }
  var xml = setSlideShow_(xmlOf_(source, slidePath) || '', true), made = [];
  for (var i = 0; i < count; i++) {
    var rels = rewritePartRels_(ctx, slidePath, 'ppt/slides', srels, function (type, abs) {
      if (/\/notesSlide$/.test(type)) return null;              // ノートは写さない
      if (/\/slideLayout$/.test(type)) return useLayout || firstLayout_(target);
      return copyPartDeep_(ctx, abs);
    });
    made.push(addSlidePart_(target, xml, rels));
  }
  var size = function (parts) { var m = (xmlOf_(parts, 'ppt/presentation.xml') || '').match(/<p:sldSz\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/); return m ? m[1] + 'x' + m[2] : ''; };
  return { slides: made, importedMaster: imported, sizeDiffers: !!size(source) && size(source) !== size(target) };
}

// ひな形（source）のページの大きさを、入れる先のパワポ（target）の大きさに合わせる。
// 縦横の比が同じなら同じ割合で縮める（広げる）。比が違えば、ページに収まる割合にして真ん中に置く。
// source のページ・レイアウト・マスターの図形の位置と大きさ・文字の大きさ・線の太さ・文字の余白などを、同じ割合で変える
// （source の中身を書き換える。importSlideCopies_ の前に呼ぶ）。
// 戻り値 { scaled: 合わせたか, factor: 割合, sameRatio: 縦横の比が同じか }
function fitSourceToTargetSize_(target, source) {
  var size = function (parts) {
    var m = (xmlOf_(parts, 'ppt/presentation.xml') || '').match(/<p:sldSz\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
    return m ? { cx: parseInt(m[1], 10), cy: parseInt(m[2], 10) } : null;
  };
  var t = size(target), s = size(source);
  if (!t || !s || (t.cx === s.cx && t.cy === s.cy)) return { scaled: false, factor: 1, sameRatio: true };
  var f = Math.min(t.cx / s.cx, t.cy / s.cy);
  var dx = Math.round((t.cx - s.cx * f) / 2), dy = Math.round((t.cy - s.cy * f) / 2);
  for (var p in source) {
    if (/^ppt\/(slides|slideLayouts|slideMasters)\/[^\/]+\.xml$/.test(p)) putXml_(source, p, scaleDrawingXml_(xmlOf_(source, p), f, dx, dy));
  }
  var prs = xmlOf_(source, 'ppt/presentation.xml')
    .replace(/(<p:sldSz\b[^>]*\bcx=")\d+("[^>]*\bcy=")\d+(")/, '$1' + t.cx + '$2' + t.cy + '$3');
  putXml_(source, 'ppt/presentation.xml', prs);
  return { scaled: true, factor: f, sameRatio: Math.abs(t.cx / t.cy - s.cx / s.cy) < 0.01 };
}

// 図形の位置・大きさを f 倍し（位置はさらに dx・dy ずらす）、文字の大きさ・線の太さなども f 倍する。
// グループの中の図形は、グループの子の座標（chOff・chExt）も同じ式で変えるので、グループに対する位置は変わらない
function scaleDrawingXml_(xml, f, dx, dy) {
  var n = function (v) { return Math.round(parseInt(v, 10) * f); };
  var attrs = function (tag, names, min) {
    return tag.replace(new RegExp('\\b(' + names + ')="(-?\\d+)"', 'g'), function (all, k, v) {
      return k + '="' + (min ? Math.max(min, n(v)) : n(v)) + '"';
    });
  };
  return String(xml)
    .replace(/<a:(off|chOff)\b[^>]*\/>/g, function (tag) {
      return tag.replace(/\bx="(-?\d+)"/, function (a, v) { return 'x="' + Math.round(parseInt(v, 10) * f + dx) + '"'; })
                .replace(/\by="(-?\d+)"/, function (a, v) { return 'y="' + Math.round(parseInt(v, 10) * f + dy) + '"'; });
    })
    .replace(/<a:(ext|chExt)\b[^>]*\bc[xy]="[^>]*\/>/g, function (tag) { return attrs(tag, 'cx|cy'); })
    // 文字の大きさ（1/100pt。100 より小さくはしない）・文字の間隔・行や段落の間隔（1/100pt）
    .replace(/<a:(rPr|defRPr|endParaRPr)\b[^>]*>/g, function (tag) { return attrs(attrs(tag, 'sz', 100), 'kern|spc'); })
    .replace(/<a:spcPts\b[^>]*>/g, function (tag) { return attrs(tag, 'val'); })
    // 線の太さ・文字の余白・字下げ・タブ位置・表の列の幅と行の高さとセルの余白・影やぼかし（EMU）
    .replace(/<a:(ln|lnL|lnR|lnT|lnB|lnTlToBr|lnBlToTr)\b[^>]*>/g, function (tag) { return attrs(tag, 'w'); })
    .replace(/<a:bodyPr\b[^>]*>/g, function (tag) { return attrs(tag, 'lIns|tIns|rIns|bIns'); })
    .replace(/<a:(pPr|lvl\dpPr)\b[^>]*>/g, function (tag) { return attrs(tag, 'marL|marR|indent'); })
    .replace(/<a:tab\b[^>]*>/g, function (tag) { return attrs(tag, 'pos'); })
    .replace(/<a:gridCol\b[^>]*>/g, function (tag) { return attrs(tag, 'w'); })
    .replace(/<a:tr\b[^>]*>/g, function (tag) { return attrs(tag, 'h'); })
    .replace(/<a:tcPr\b[^>]*>/g, function (tag) { return attrs(tag, 'marL|marR|marT|marB'); })
    .replace(/<a:(outerShdw|innerShdw|prstShdw|glow|softEdge|reflection)\b[^>]*>/g, function (tag) { return attrs(tag, 'dist|blurRad|rad'); });
}

// 部品の関係ファイルのパス（ppt/slides/slide1.xml → ppt/slides/_rels/slide1.xml.rels）
function partRelsPath_(p) { return posixDir_(p) + '/_rels/' + p.replace(/^.*\//, '') + '.rels'; }
// dir から見た abs の相対パス
function relPathFrom_(dir, abs) {
  var a = dir.split('/'), b = abs.split('/'), k = 0;
  while (k < a.length && k < b.length - 1 && a[k] === b[k]) k++;
  var up = [];
  for (var i = k; i < a.length; i++) up.push('..');
  return up.concat(b.slice(k)).join('/');
}
// 部品の関係のうち、種類（slideLayout・slideMaster・theme など）の行き先（パス）
function relTargetOfType_(parts, partPath, type) {
  var rels = xmlOf_(parts, partRelsPath_(partPath)) || '', m = null;
  var re = /<Relationship\b[^>]*\/>/g, r;
  while ((r = re.exec(rels)) !== null) {
    if (new RegExp('Type="[^"]*/' + type + '"').test(r[0]) && !/TargetMode="External"/.test(r[0])) { m = r[0]; break; }
  }
  var t = m ? (m.match(/Target="([^"]+)"/) || [])[1] : '';
  return t ? normPartPath_(posixDir_(partPath), t) : '';
}
// source のレイアウトと同じ中身のレイアウトが target にあれば、そのパス（元のマスター・テーマも同じ中身のときだけ）
function sameLayoutIn_(target, source, layoutAbs) {
  var lx = xmlOf_(source, layoutAbs), sm = relTargetOfType_(source, layoutAbs, 'slideMaster');
  var mx = sm ? xmlOf_(source, sm) : null, st = sm ? relTargetOfType_(source, sm, 'theme') : '';
  var tx = st ? xmlOf_(source, st) : null;
  for (var p in target) {
    if (!/^ppt\/slideLayouts\/[^\/]+\.xml$/.test(p) || xmlOf_(target, p) !== lx) continue;
    var tm = relTargetOfType_(target, p, 'slideMaster');
    if (!tm || xmlOf_(target, tm) !== mx) continue;
    var tt = relTargetOfType_(target, tm, 'theme');
    if ((tt ? xmlOf_(target, tt) : null) !== tx) continue;
    return p;
  }
  return '';
}
// 関係ファイルを書き直す（行き先は decide(種類, 元のパス) が返す target 側のパス。null ならその関係を落とす）
function rewritePartRels_(ctx, srcPath, dstDir, rels, decide) {
  var body = String(rels || '').replace(/<Relationship\b[^>]*\/>/g, function (rel) {
    if (/TargetMode="External"/.test(rel)) return rel;
    var tgt = (rel.match(/Target="([^"]+)"/) || [])[1] || '', type = (rel.match(/Type="([^"]+)"/) || [])[1] || '';
    var dst = decide(type, normPartPath_(posixDir_(srcPath), tgt));
    if (dst === null) return '';
    if (!dst) return rel;
    return rel.replace(/Target="[^"]+"/, 'Target="' + relPathFrom_(dstDir, dst) + '"');
  });
  return body || ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');
}
// source の部品を target に写す（その部品の関係の行き先も、たどって写す。同じ部品は1回だけ）。写した先のパスを返す
function copyPartDeep_(ctx, abs) {
  if (ctx.map[abs]) return ctx.map[abs];
  var blob = ctx.source[abs];
  if (!blob) return '';                                           // 無い部品（壊れた関係）は写さない
  var dir = posixDir_(abs), file = abs.replace(/^.*\//, '');
  var ext = (file.match(/\.[^.\/]+$/) || [''])[0], base = file.slice(0, file.length - ext.length).replace(/\d+$/, '');
  var dst, n = 1;
  if (/^(slideMaster|slideLayout|theme)$/.test(base)) {           // ならわしの名前（番号はいまある最大の次）
    for (var p in ctx.target) { var m = p.match(new RegExp('^' + dir + '/' + base + '(\\d+)' + ext.replace('.', '\\.') + '$')); if (m) n = Math.max(n, parseInt(m[1], 10) + 1); }
    dst = dir + '/' + base + n + ext;
  } else {
    do { dst = dir + '/' + base + 'imp' + (n++) + ext; } while (ctx.target[dst]);
  }
  ctx.map[abs] = dst;
  var isXml = /\.(xml|rels)$/i.test(ext);
  ctx.target[dst] = isXml ? Utilities.newBlob(xmlOf_(ctx.source, abs), 'application/xml', dst) : blob.copyBlob ? blob.copyBlob().setName(dst) : blob.setName(dst);
  ctPartType_(ctx.target, ctx.source, abs, dst);
  var rels = xmlOf_(ctx.source, partRelsPath_(abs));
  if (rels) putXml_(ctx.target, partRelsPath_(dst), rewritePartRels_(ctx, abs, dir, rels, function (type, a) { return copyPartDeep_(ctx, a); }));
  return dst;
}
// 写した部品の種類を [Content_Types].xml に登録する（source の登録と同じ種類で）
function ctPartType_(target, source, abs, dst) {
  var sct = xmlOf_(source, '[Content_Types].xml') || '', ct = xmlOf_(target, '[Content_Types].xml') || '';
  var ov = sct.match(new RegExp('<Override PartName="/' + abs.replace(/[.*+?^${}()|[\]\\\/]/g, '\\$&') + '" ContentType="([^"]+)"\\s*/>'));
  if (ov) {
    ct = ct.replace('</Types>', '<Override PartName="/' + dst + '" ContentType="' + ov[1] + '"/></Types>');
  } else {
    var ext = (dst.match(/\.([^.\/]+)$/) || [])[1] || '';
    if (ext && !new RegExp('<Default Extension="' + ext + '"', 'i').test(ct)) {
      var df = sct.match(new RegExp('<Default Extension="' + ext + '" ContentType="([^"]+)"', 'i'));
      ct = df ? ct.replace('</Types>', '<Default Extension="' + ext + '" ContentType="' + df[1] + '"/></Types>') : ensureDefaultType_(ct, ext);
    }
  }
  putXml_(target, '[Content_Types].xml', ct);
}
// 写したマスターを presentation.xml に登録する。マスター・レイアウトの番号（id）は、ほかと重ならないように振り直す
function registerSlideMaster_(target, masterPath) {
  var prs = xmlOf_(target, 'ppt/presentation.xml') || '', maxId = 2147483647, m, p;
  var re = /<p:sldMasterId\b[^>]*\bid="(\d+)"/g;
  while ((m = re.exec(prs)) !== null) maxId = Math.max(maxId, parseInt(m[1], 10));
  for (p in target) {
    if (p === masterPath || !/^ppt\/slideMasters\/[^\/]+\.xml$/.test(p)) continue;
    var re2 = /<p:sldLayoutId\b[^>]*\bid="(\d+)"/g, mx = xmlOf_(target, p) || '';
    while ((m = re2.exec(mx)) !== null) maxId = Math.max(maxId, parseInt(m[1], 10));
  }
  var masterId = ++maxId;
  putXml_(target, masterPath, (xmlOf_(target, masterPath) || '').replace(/(<p:sldLayoutId\b[^>]*\bid=")(\d+)(")/g, function (all, a, id, b) { return a + (++maxId) + b; }));
  var prsRels = xmlOf_(target, 'ppt/_rels/presentation.xml.rels') || '', maxRid = 0;
  var re3 = /Id="rId(\d+)"/g;
  while ((m = re3.exec(prsRels)) !== null) maxRid = Math.max(maxRid, parseInt(m[1], 10));
  var rid = 'rId' + (maxRid + 1);
  putXml_(target, 'ppt/_rels/presentation.xml.rels', prsRels.replace('</Relationships>',
    '<Relationship Id="' + rid + '" Type="' + SPLICE_REL_ + 'slideMaster" Target="' + masterPath.replace(/^ppt\//, '') + '"/></Relationships>'));
  var entry = '<p:sldMasterId id="' + masterId + '" r:id="' + rid + '"/>';
  prs = /<\/p:sldMasterIdLst>/.test(prs) ? prs.replace('</p:sldMasterIdLst>', entry + '</p:sldMasterIdLst>')
                                          : prs.replace(/(<p:presentation\b[^>]*>)/, '$1<p:sldMasterIdLst>' + entry + '</p:sldMasterIdLst>');
  putXml_(target, 'ppt/presentation.xml', prs);
}

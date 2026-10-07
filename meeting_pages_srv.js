// === 定例会スライド（前半）：新メンバー・更新メンバー／バイスプレジデントによる報告／ネットワーキングリーダー ===
//
// どれもバイスプレジデントがルーティンチェックシートに書く内容。画面で確かめて直したものを受け取る。
// テンプレートのページは、書いてある文字と図形の並びから見分ける（差し込み口 {{ }} は要らない）。
//
//  ・新メンバー／更新メンバー … 1人1枚。人数ぶんページを増やし、いない日はページを非表示にする。
//    お名前（いちばん大きな字）・その下の会社名と【カテゴリー】・写真、更新は「1年更新／2年更新」も
//  ・バイスプレジデントによる報告 … 「月間リファーラル数の平均」などの見出しの次の行の数字と見出しの年月、
//    サンキュー額、「○/○速報」の日付
//  ・ネットワーキングリーダー … 月の最初の定例会だけ表示（ほかの週は非表示）。部門ごとのページに数と受賞者、
//    受賞者がお2人ならお2人のページ、まとめのページ（「○月のネットワーキングリーダーの皆さま」）には全員

// --- 共通の道具 ---
function fpSize_(parts) {
  var prs = xmlOf_(parts, 'ppt/presentation.xml') || '';
  var sz = prs.match(/<p:sldSz\s+cx="(\d+)"\s+cy="(\d+)"/);
  return { W: sz ? +sz[1] : 12192000, H: sz ? +sz[2] : 6858000 };
}
// 名簿（氏名 → 会社名・カテゴリー）
function fpRoster_() {
  var by = {};
  try { (getMemberMaster({ membersOnly: true }).members || []).forEach(function (m) { by[normName_(m.name)] = m; }); } catch (e) {}
  return by;
}
// { name, raw, category } → ページに入れる方 { name, company, category, matched }。名簿に無い方は書いてあったとおり
function fpPerson_(p, by) {
  if (!p) return null;
  var m = by[normName_(p.name || '')];
  // 名簿の文字の中の改行は外す（残っていると、行が増えて枠の外にはみ出し、見えなくなることがある）
  if (m) return { name: m.name, company: slideOneLine_(m.company), category: slideOneLine_(m.title), matched: true };
  var nm = String(p.name || p.raw || '').replace(/(さん|様)$/, '').trim();
  return nm ? { name: nm, company: '', category: slideOneLine_(p.category), matched: false } : null;
}
function fpNonEmpty_(t) { return !!String(t || '').trim(); }

// 図形を見えなくする／見えるようにする（アニメーションがその図形を指していることがあるので、消さずに隠す）
function fpHide_(xml, id, hide) {
  var r = findShapeRange_(xml, id);
  if (!r) return xml;
  var seg = xml.substring(r.start, r.end).replace(/<p:cNvPr\b([^>]*?)(\/?)>/, function (all, attrs, sc) {
    return '<p:cNvPr' + attrs.replace(/\shidden="[^"]*"/g, '') + (hide ? ' hidden="1"' : '') + sc + '>';
  });
  return xml.substring(0, r.start) + seg + xml.substring(r.end);
}
// 写真の枠にその方の写真を入れる（枠いっぱいに切り抜く）。写真の無い方は枠を隠す
function fpSetPhoto_(parts, xml, rels, picId, name, cache) {
  var photo = name ? mpAddPhoto_(parts, cache, name) : null;
  if (!photo) return { xml: fpHide_(xml, picId, true), rels: rels, ok: false };
  xml = fpHide_(xml, picId, false);
  var box = readShapeGeomEmu_(xml, picId);
  if (box && photo.width && photo.height) xml = setSrcRectInPic_(xml, picId, coverCrop_(photo.width, photo.height, box.cx, box.cy));
  var set = setPicImage_(xml, rels, picId, '../media/' + photo.path.replace('ppt/media/', ''));
  return { xml: set.xml, rels: set.rels, ok: true };
}
// 段落を1つ外す
function fpDropPara_(xml, id, idx) {
  var r = findShapeRange_(xml, id);
  if (!r) return xml;
  var seg = xml.substring(r.start, r.end), ps = findTagRanges_(seg, 'a:p');
  if (!ps[idx] || ps.length < 2) return xml;
  seg = seg.substring(0, ps[idx].start) + seg.substring(ps[idx].end);
  return xml.substring(0, r.start) + seg + xml.substring(r.end);
}
// 段落の文字が lines 行に収まるよう、その段落の文字だけ小さくする（元の大きさの半分まで）
function fpFitPara_(xml, id, idx, wEmu, lines) {
  var r = findShapeRange_(xml, id);
  if (!r) return xml;
  var seg = xml.substring(r.start, r.end), ps = findTagRanges_(seg, 'a:p');
  if (!ps[idx]) return xml;
  var p = seg.substring(ps[idx].start, ps[idx].end), sz = 0, m, re = /<a:(?:rPr|endParaRPr)\b[^>]*\ssz="(\d+)"/g;
  while ((m = re.exec(p)) !== null) sz = Math.max(sz, +m[1]);
  var text = slideText_(p).trim();
  if (!sz || !text) return xml;
  var avail = Math.max(1, wEmu - 2 * 91440) / 12700 * 0.95 * (lines || 1), need = riTextEm_(text) * sz / 100;
  if (need <= avail) return xml;
  var to = Math.max(Math.floor(sz / 2 / 50) * 50, Math.floor(sz * avail / need / 50) * 50);
  p = p.replace(/(<a:(?:rPr|endParaRPr)\b[^>]*\ssz=")\d+(")/g, '$1' + to + '$2');
  seg = seg.substring(0, ps[idx].start) + p + seg.substring(ps[idx].end);
  return xml.substring(0, r.start) + seg + xml.substring(r.end);
}
// 図形の段落がそれぞれ1行に収まるよう、図形の字をそろえて小さくする（元の大きさの半分まで）
function fpFitLines_(xml, id, wEmu) {
  var r = findShapeRange_(xml, id);
  if (!r) return xml;
  var seg = xml.substring(r.start, r.end), sz = 0, m, re = /<a:(?:rPr|endParaRPr)\b[^>]*\ssz="(\d+)"/g;
  while ((m = re.exec(seg)) !== null) sz = Math.max(sz, +m[1]);
  if (!sz) return xml;
  var avail = Math.max(1, wEmu - 2 * 91440) / 12700 * 0.95, ratio = 1;
  findTagRanges_(seg, 'a:p').forEach(function (pr) {
    var em = riTextEm_(slideText_(seg.substring(pr.start, pr.end)).trim());
    if (em) ratio = Math.min(ratio, avail / (em * sz / 100));
  });
  if (ratio >= 1) return xml;
  var to = Math.max(Math.floor(sz / 2 / 50) * 50, Math.floor(sz * ratio / 50) * 50);
  seg = seg.replace(/(<a:(?:rPr|endParaRPr)\b[^>]*\ssz=")\d+(")/g, '$1' + to + '$2');
  return xml.substring(0, r.start) + seg + xml.substring(r.end);
}
// 数字の幅（全角・半角）を元の字に合わせる
function fpDigits_(value, like) {
  var s = String(value);
  return /[０-９]/.test(String(like || '')) ? s.replace(/[0-9]/g, function (d) { return String.fromCharCode(d.charCodeAt(0) + 0xFEE0); }) : s;
}
// 段落の中の文字を、正規表現で見つけたところだけ置き換える（ランに分かれていても。書式はそのまま）
//   rules … [{ re, value(m) }]。value が null なら置き換えない
function fpReplaceInPara_(pXml, rules) {
  return replaceInParagraph_(pXml, function (joined) {
    var hits = [];
    rules.forEach(function (r) {
      var m = new RegExp(r.re.source, r.re.flags.replace('g', '')).exec(joined);
      if (!m) return;
      var v = r.value(m);
      if (v != null && v !== m[0]) hits.push({ start: m.index, end: m.index + m[0].length, value: String(v) });
    });
    hits.sort(function (a, b) { return a.start - b.start; });
    var out = [], last = -1;
    hits.forEach(function (h) { if (h.start >= last) { out.push(h); last = h.end; } });
    return out;
  });
}
// 図形の段落ごとに fn(段落のXML, 文字, 番号) → 新しい段落のXML
function fpMapParas_(xml, id, fn) {
  var r = findShapeRange_(xml, id);
  if (!r) return xml;
  var seg = xml.substring(r.start, r.end), ps = findTagRanges_(seg, 'a:p'), out = '', prev = 0;
  for (var i = 0; i < ps.length; i++) {
    var p = seg.substring(ps[i].start, ps[i].end);
    out += seg.substring(prev, ps[i].start) + fn(p, slideText_(p), i);
    prev = ps[i].end;
  }
  out += seg.substring(prev);
  return xml.substring(0, r.start) + out + xml.substring(r.end);
}

// ページの「1人ぶん」の枠：お名前（大きな字の1行）、その下の会社名・【カテゴリー】の行、その上の写真。
// skip(s) … お名前として見ない図形（部門・数・年数など）
function fpUnits_(xml, W, H, skip) {
  var shapes = riShapes_(xml), pics = riPhotos_(shapes, W, H), title = riTitleOf_(shapes, H);
  var texts = shapes.filter(function (s) {
    return s.tag === 'p:sp' && s.text.trim() && s !== title && !(skip && skip(s));
  });
  var names = texts.filter(function (s) {
    return s.sz >= 2000 && s.paras.filter(fpNonEmpty_).length === 1 && riNameLike_(s.text)
        && !/メンバー|写真|トリミング|報告|発表|皆さま|皆様/.test(s.text);
  });
  var big = 0;
  names.forEach(function (s) { big = Math.max(big, s.sz); });
  names = names.filter(function (s) { return s.sz >= big * 0.8; }).sort(function (a, b) { return a.ax - b.ax; });
  var used = {};
  var units = names.map(function (n) {
    var info = texts.filter(function (s) {
      return names.indexOf(s) < 0 && s.ay >= n.ay + n.acy * 0.5 && s.ay <= n.ay + n.acy + H * 0.25 && riOverlap_(s, n) >= 0.3;
    }).sort(function (a, b) { return a.ay - b.ay; });
    var photo = riPicNear_(pics, n, false, used);
    if (photo) used[photo.id] = true;
    return { name: n, info: info, photo: photo };
  });
  return { shapes: shapes, pics: pics, title: title, units: units };
}
// 会社名と【カテゴリー】を入れる。【 】のある行がカテゴリー、ほかの行の1つ目が会社名（残りは空にする）。
// 空になった行は外す（2行の枠で会社名が無いときに、カテゴリーが1行下がらないように）
function fpWriteInfo_(xml, info, p, lines) {
  var company = p ? p.company : '', cat = p && p.category ? '【' + p.category + '】' : '', usedCompany = false;
  info.forEach(function (s) {
    // 【 】が2行にまたがっていたら（「【デジタル広告制作」「(ホームページ・動画)】」）、1行目に入れて残りは外す
    var inCat = false;
    var plan = s.paras.map(function (t) {
      if (!t.trim()) return null;
      if (inCat) { if (/】/.test(t)) inCat = false; return ''; }
      if (/【/.test(t)) { inCat = !/】/.test(t); return cat; }
      if (/】/.test(t)) return cat;
      if (!usedCompany) { usedCompany = true; return company; }
      return '';
    });
    for (var i = plan.length - 1; i >= 0; i--) {
      if (plan[i] === null) continue;
      var others = plan.filter(function (x, k) { return k !== i && x; }).length;
      if (!plan[i] && others) xml = fpDropPara_(xml, s.id, i);
      else xml = riSetPara_(xml, s.id, i, plan[i]);
    }
    var r = findShapeRange_(xml, s.id), count = r ? findTagRanges_(xml.substring(r.start, r.end), 'a:p').length : 0;
    for (var k = 0; k < count; k++) xml = fpFitPara_(xml, s.id, k, s.acx, lines || 1);
  });
  return xml;
}
// 1人ぶんを入れる（お名前・会社名・カテゴリー・写真）。p が null なら空にして写真を隠す
function fpFillUnit_(parts, xml, rels, u, p, cache, res) {
  xml = setParagraphsInShape_(xml, u.name.id, [p ? p.name : '']);
  xml = fpFitPara_(xml, u.name.id, 0, u.name.acx, 1);
  xml = fpWriteInfo_(xml, u.info, p, u.infoLines || 1);
  var ok = false;
  if (u.photo) {
    var ph = fpSetPhoto_(parts, xml, rels, u.photo.id, p ? p.name : '', cache);
    xml = ph.xml; rels = ph.rels; ok = ph.ok;
    if (p && !ph.ok && res && res.noPhoto.indexOf(p.name) < 0) res.noPhoto.push(p.name);
  }
  return { xml: xml, rels: rels, photo: ok };
}
// ページを複製して、並びの after のうしろに入れる（ノートは引き継がない）。戻り値は新しいページのパス
function fpClonePages_(parts, modelPath, count, afterPath, modelXml) {
  var xml = modelXml || xmlOf_(parts, modelPath);
  var rels = (xmlOf_(parts, relsPathOf_(modelPath)) || '').replace(/<Relationship\b[^>]*notesSlides\/[^>]*\/>/g, '');
  var added = [];
  for (var i = 0; i < count; i++) added.push(addSlidePart_(parts, xml, rels));
  if (!added.length) return [];
  var entries = slideEntries_(parts), out = [];
  entries.forEach(function (e) { out.push(e); if (e.path === afterPath) out = out.concat(added); });
  if (out.length === entries.length) out = out.concat(added);
  setSlideEntries_(parts, out);
  return added.map(function (a) { return a.path; });
}

// === 新メンバー・更新メンバー ===
function fpMemberKind_(xml) {
  var t = riNorm_(slideText_(xml));
  var renew = t.indexOf('更新メンバー') >= 0, neu = /(^|[^更])新メンバー/.test(t);
  return renew && neu ? 'both' : renew ? 'renew' : neu ? 'new' : '';
}
var FP_YEARS_RE_ = /^[0-9０-９]\s*年\s*更新$/;
// 1人1枚のページの枠。お名前が1つだけのときだけ
function fpMemberUnit_(xml, W, H) {
  var u = fpUnits_(xml, W, H, function (s) { return FP_YEARS_RE_.test(riNorm_(s.text)); });
  if (u.units.length !== 1) return null;
  var unit = u.units[0];
  unit.years = u.shapes.filter(function (s) { return s.tag === 'p:sp' && FP_YEARS_RE_.test(riNorm_(s.text)); })[0] || null;
  // 写真の後ろの見本の文字（「写真のサイズを枠に合わせてトリミング」）
  unit.holder = unit.photo ? u.shapes.filter(function (s) {
    return s.tag === 'p:sp' && /写真|トリミング/.test(s.text) && riOverlap_(s, unit.photo) > 0.5;
  })[0] || null : null;
  return unit;
}
// lists … { newMembers: [{ name, raw, category }], renewMembers: [{ name, raw, category, years }],
//           unread: { new, renew }（ルーティンチェックシートに書いてあったのに、お名前を読み取れなかった記載。お知らせに出す） }
function applyMemberPages_(parts, lists, cache) {
  var res = { new: [], renew: [], hidden: [], noPhoto: [], unmatched: [], missing: [], pages: [], ethics: [],
              unread: (lists && lists.unread) || {} };
  var sz = fpSize_(parts), by = fpRoster_(), pages = { new: [], renew: [], both: [] };
  slideOrder_(parts).forEach(function (path) {
    var xml = xmlOf_(parts, path), k = xml ? fpMemberKind_(xml) : '';
    if (!k) return;
    if (k === 'both') pages.both.push(path);
    else if (fpMemberUnit_(xml, sz.W, sz.H)) pages[k].push(path);
  });
  var people = {};
  ['new', 'renew'].forEach(function (k) {
    people[k] = ((k === 'new' ? lists.newMembers : lists.renewMembers) || []).map(function (x) {
      var p = fpPerson_(x, by);
      if (p) { p.years = parseInt(x.years, 10) || 0; if (!p.matched) res.unmatched.push(p.name); }
      return p;
    }).filter(Boolean);
  });
  var blocks = [];                                   // まとまり（倫理規定のページをうしろに置く）
  ['new', 'renew'].forEach(function (k) {
    var list = people[k], have = pages[k];
    if (!have.length) { if (list.length && !pages.both.length) res.missing.push(k); return; }
    var models = have.map(function (p) { return xmlOf_(parts, p); });
    var extra = list.length > have.length
      ? fpClonePages_(parts, have[have.length - 1], list.length - have.length, have[have.length - 1], models[models.length - 1]) : [];
    var use = have.concat(extra);
    blocks.push({ kind: k, paths: use, any: list.length > 0 });
    use.forEach(function (path, i) {
      var xml = xmlOf_(parts, path), rp = relsPathOf_(path), rels = xmlOf_(parts, rp);
      var p = list[i] || null, u = fpMemberUnit_(xml, sz.W, sz.H);
      if (!p) { putXml_(parts, path, setSlideShow_(xml, false)); res.hidden.push(path); return; }
      var f = fpFillUnit_(parts, xml, rels, u, p, cache, res);
      xml = f.xml; rels = f.rels;
      if (k === 'renew' && u.years && p.years) xml = setParagraphsInShape_(xml, u.years.id, [p.years + '年更新']);
      if (u.holder && !f.photo) xml = setParagraphsInShape_(xml, u.holder.id, ['']);   // 写真が無い方：見本の文字を消す
      putXml_(parts, path, setSlideShow_(xml, true));
      if (rels) putXml_(parts, rp, rels);
      res[k].push({ path: path, name: p.name, years: k === 'renew' ? p.years : 0 });
      res.pages.push(path);
    });
  });
  // 1枚に「新メンバー」「更新メンバー」の欄がある作り（公式ファイルの「新規および更新メンバー」）
  pages.both.forEach(function (path) {
    var xml = xmlOf_(parts, path), r = fpFillBothPage_(xml, people['new'], people.renew, sz.H);
    if (!r) return;
    xml = r.xml;
    var any = people['new'].length || people.renew.length;
    blocks.push({ kind: 'both', paths: [path], any: !!any });
    putXml_(parts, path, setSlideShow_(xml, !!any));
    if (!any) res.hidden.push(path);
    else res.pages.push(path);
    if (any && !res['new'].length && !res.renew.length) {
      r.placed['new'].forEach(function (p) { res['new'].push({ path: path, name: p.name, years: 0 }); });
      r.placed.renew.forEach(function (p) { res.renew.push({ path: path, name: p.name, years: p.years }); });
      res.crowded = r.placed.crowded;
      res.noSlot = r.placed.noSlot;
    }
  });
  var moved = fpNewBeforeRenew_(parts, blocks, sz.W, sz.H);
  fpEthicsPages_(parts, blocks, res, sz.W, sz.H);
  if (moved) syncSections_(parts);                   // PowerPoint のセクションも、入れ替えた並びに合わせる
  res.message = fpMemberMessage_(res, people);
  return res;
}
// 新メンバーのまとまりを、更新メンバーのまとまりの前にする（定例会の進行・ルーティンチェックシートの「新入会 → 更新式」と同じ順）。
// 雛形が「更新メンバー → 新メンバー」の並びでも入れ替える。一緒に動かすもの：
//   ・まとまりのすぐ前・すぐうしろの、1人1枚でない「新メンバー」のページ（見出し・一言など。fpSideEnd_）
//   ・そのすぐうしろの倫理規定のページ（まとまりごとに倫理規定がある雛形。文字が少し違っても、その場所のものを使う）
// 入れる場所は、更新メンバーのまとまりの前（すぐ前に1人1枚でない「更新メンバー」のページ＝見出しがあれば、その前）。
// 倫理規定のページは、このあと fpEthicsPages_ がそれぞれのうしろに付ける
function fpNewBeforeRenew_(parts, blocks, W, H) {
  var nb = blocks.filter(function (b) { return b.kind === 'new'; })[0], rb = blocks.filter(function (b) { return b.kind === 'renew'; })[0];
  if (!nb || !rb) return false;
  var entries = slideEntries_(parts), order = entries.map(function (e) { return e.path; });
  var ns = nb.paths.map(function (p) { return order.indexOf(p); }), rs = rb.paths.map(function (p) { return order.indexOf(p); });
  var firstNew = Math.min.apply(null, ns), lastNew = Math.max.apply(null, ns), firstRenew = Math.min.apply(null, rs);
  if (firstNew < 0 || firstRenew < 0 || lastNew < firstRenew) return false;   // もう全部が前にある
  var from = fpSideEnd_(parts, order, firstNew, -1, 'new', W, H), to = fpSideEnd_(parts, order, lastNew, 1, 'new', W, H);
  var move = nb.paths.slice(), i;
  for (i = from; i < firstNew; i++) move.push(order[i]);
  for (i = lastNew + 1; i <= to; i++) move.push(order[i]);
  if (to + 1 < order.length && rb.paths.indexOf(order[to + 1]) < 0 && fpIsEthics_(parts, order[to + 1])) move.push(order[to + 1]);
  var headPath = order[fpSideEnd_(parts, order, firstRenew, -1, 'renew', W, H)];
  var moving = entries.filter(function (e) { return move.indexOf(e.path) >= 0; });
  var rest = entries.filter(function (e) { return move.indexOf(e.path) < 0; }), at = 0;
  for (i = 0; i < rest.length; i++) if (rest[i].path === headPath) { at = i; break; }
  setSlideEntries_(parts, rest.slice(0, at).concat(moving, rest.slice(at)));
  return true;
}
// まとまりのすぐ前（step = -1）・すぐうしろ（step = 1）に続く、1人1枚でない同じ種類のページ（見出し・「新メンバーからの一言」など。
// 倫理規定のページは除く）。戻り値は、いちばん端のページの位置（無ければ i のまま）
function fpSideEnd_(parts, order, i, step, kind, W, H) {
  for (var j = i + step; j >= 0 && j < order.length; j += step) {
    var x = xmlOf_(parts, order[j]);
    if (!x || fpSideKind_(x) !== kind || fpMemberUnit_(x, W, H) || fpIsEthics_(parts, order[j])) break;
    i = j;
  }
  return i;
}
// 見出しなどのページの種類。「新メンバー」「更新メンバー」のほか、「入会式」「新入会」「更新式」の見出しも
function fpSideKind_(xml) {
  var k = fpMemberKind_(xml);
  if (k) return k;
  var t = riNorm_(slideText_(xml)), neu = /入会式|新入会/.test(t), renew = /更新式/.test(t);
  return neu && renew ? 'both' : neu ? 'new' : renew ? 'renew' : '';
}
function fpIsEthics_(parts, path) {
  var x = xmlOf_(parts, path);
  return !!x && riNorm_(slideText_(x)).indexOf('倫理規定') >= 0;
}
// 同じ倫理規定のページかを見る文字。ページ番号・日付（a:fld）は、置いた場所で変わるので除く
function fpEthicsKey_(xml) {
  return riNorm_(slideText_(String(xml || '').replace(/<a:fld\b[\s\S]*?<\/a:fld>/g, '')));
}
// 倫理規定のページ：新メンバー・更新メンバーのまとまりの、それぞれすぐうしろに1枚ずつ（新しく入った方・更新した方が読み上げる）。
// すぐうしろに無ければ倫理規定のページを複製して入れる（Activeチャプターの雛形は「更新メンバー → 新メンバー → 倫理規定」の並びで、
// fpNewBeforeRenew_ で「新メンバー → 倫理規定 → 更新メンバー」にしてから、更新メンバーのうしろに写しを入れる）。
// そのまとまりに人がいるときだけ表示（どちらもいない日は、倫理規定のページも非表示）
function fpEthicsPages_(parts, blocks, res, W, H) {
  if (!blocks.length) return;
  var isEthics = function (p) { return fpIsEthics_(parts, p); };
  var order = slideOrder_(parts), lastOf = function (b) {       // まとまりの最後（うしろに続く「一言」などのページも、まとまりのうち）
    var best = -1;
    b.paths.forEach(function (p) { best = Math.max(best, order.indexOf(p)); });
    return b.kind === 'both' || best < 0 ? best : fpSideEnd_(parts, order, best, 1, b.kind, W, H);
  };
  blocks.sort(function (a, b) { return lastOf(a) - lastOf(b); });
  // 複製のもとは、まとまりのすぐうしろにある倫理規定のページ（入れ替えたあとは、新メンバーのうしろ）。
  // 無ければ、まとまりよりうしろの倫理規定のページ、それも無ければ最初に見つかったもの
  // （「倫理規定」の言葉が入った一般規定などのページを、先に拾わないように）
  var model = null, lastAll = lastOf(blocks[blocks.length - 1]);
  blocks.forEach(function (b) { var nx = order[lastOf(b) + 1]; if (!model && nx && isEthics(nx)) model = nx; });
  for (var i = lastAll + 1; i < order.length && !model; i++) if (isEthics(order[i])) model = order[i];
  for (var j = 0; j < order.length && !model; j++) if (isEthics(order[j])) model = order[j];
  if (!model) return;
  var modelXml = setSlideShow_(xmlOf_(parts, model), true), modelText = fpEthicsKey_(modelXml), used = [];
  blocks.forEach(function (b) {
    order = slideOrder_(parts);
    var at = lastOf(b), next = order[at + 1], path = next && isEthics(next) ? next : null;
    if (!path) path = fpClonePages_(parts, model, 1, order[at], modelXml)[0];
    putXml_(parts, path, setSlideShow_(xmlOf_(parts, path), b.any));
    used.push(path);
    res.ethics.push({ kind: b.kind, path: path, shown: b.any, added: path !== next });
  });
  // まとまりのすぐうしろに無い同じ倫理規定のページ（まとまりの前・あいだに別のページがある・2枚目の写し）は隠す。
  // 残すと、倫理規定が余分に出る（誰もいない日にも出る・続けて2回出る）
  slideOrder_(parts).forEach(function (p) {
    if (used.indexOf(p) >= 0) return;
    var x = xmlOf_(parts, p);
    if (x && fpEthicsKey_(x) === modelText) {
      putXml_(parts, p, setSlideShow_(x, false));
      res.ethics.push({ kind: 'extra', path: p, shown: false, added: false });
    }
  });
}
// 「新メンバー」「更新メンバー」の見出しの下の「氏名」の枠に、上から順にお名前を入れる（余った枠は空に）
function fpFillBothPage_(xml, news, renews, H) {
  var shapes = riShapes_(xml).filter(function (s) { return s.tag === 'p:sp'; });
  var head = { new: null, renew: null };
  shapes.forEach(function (s) {
    var t = riNorm_(s.text);
    if (t === '新メンバー') head['new'] = s;
    if (t === '更新メンバー') head.renew = s;
  });
  if (!head['new'] || !head.renew) return null;
  var slots = { new: [], renew: [] };
  shapes.forEach(function (s) {
    if (s === head['new'] || s === head.renew || s.paras.filter(fpNonEmpty_).length > 1) return;
    if (s.ay < Math.max(head['new'].ay, head.renew.ay) || s.sz > 3200) return;
    if (!(riNorm_(s.text) === '氏名' || riNameLike_(s.text) || /年更新|[（(]\d年[）)]/.test(s.text))) return;
    var a = riOverlap_(s, head['new']), b = riOverlap_(s, head.renew);
    if (a > b && a > 0.3) slots['new'].push(s);
    else if (b > a && b > 0.3) slots.renew.push(s);
  });
  // placed … 実際に入れた方。crowded … 枠より多く、最後の枠にまとめて入れた欄。noSlot … 氏名の枠が無くて入れられなかった方
  var placed = { 'new': [], renew: [], crowded: [], noSlot: [] };
  ['new', 'renew'].forEach(function (k) {
    var list = k === 'new' ? news : renews, ss = slots[k].sort(function (a, b) { return a.ay - b.ay; });
    var label = function (p) { return p.name + (k === 'renew' && p.years ? '（' + p.years + '年）' : ''); };   // 欄の見出しが「更新メンバー」なので「年」だけ
    ss.forEach(function (s, i) {
      // 枠より人数が多いときは、最後の枠に残りの方をまとめて入れる（以前は入らない方を黙って落とし、知らせには全員の名前が出ていた）
      var ps = i === ss.length - 1 ? list.slice(i) : (list[i] ? [list[i]] : []);
      if (ps.length > 1) placed.crowded.push(k);
      ps.forEach(function (p) { placed[k].push(p); });
      xml = setParagraphsInShape_(xml, s.id, [ps.map(label).join('、')]);
      xml = fpFitPara_(xml, s.id, 0, s.acx, 1);
    });
    if (!ss.length) list.forEach(function (p) { placed.noSlot.push(p.name); });
  });
  return { xml: xml, placed: placed };
}
function fpMemberMessage_(res, people) {
  var msg = [], nm = function (list) {
    return list.map(function (x) { return x.name + 'さん' + (x.years ? '（' + x.years + '年更新）' : ''); }).join('、');
  };
  [['new', '新メンバー'], ['renew', '更新メンバー']].forEach(function (k) {
    if (res[k[0]].length) msg.push(k[1] + 'のページ：' + nm(res[k[0]]));
    else if (res.missing.indexOf(k[0]) >= 0) msg.push(k[1] + '（' + nm(people[k[0]]) + '）のページがテンプレートにありません');
    else if (res.unread[k[0]]) msg.push(k[1] + '：ルーティンチェックシートの「' + res.unread[k[0]] + '」からお名前を読み取れなかったので、'
      + 'そのページは非表示にしました（入れるときは、画面の「＋ 追加」で選んで作り直してください）');
    else msg.push(k[1] + 'はいないので、そのページは非表示にしました');
  });
  var out = msg.join('。') + '。';
  var eth = res.ethics.filter(function (e) { return e.shown; }).map(function (e) {
    return { renew: '更新メンバー', 'new': '新メンバー', both: '新規および更新メンバー' }[e.kind]; });
  if (res.ethics.length) out += eth.length ? '\n倫理規定のページを、' + eth.join('・') + 'のあとに表示しました。'
                                           : '\n倫理規定のページも非表示にしました（新メンバー・更新メンバーがいないため）。';
  if ((res.crowded || []).length) {
    out += '\n「新規および更新メンバー」のページは氏名の枠より人数が多いため、'
      + res.crowded.map(function (k) { return k === 'new' ? '新メンバー' : '更新メンバー'; }).join('・')
      + 'の最後の枠に、残りの方をまとめて入れました（枠を増やすときはテンプレートを直してください）。';
  }
  if ((res.noSlot || []).length) out += '\n「新規および更新メンバー」のページに氏名の枠が見つからず、入れられなかった方: ' + res.noSlot.join('、');
  if (res.unmatched.length) out += '\n名簿に無い方（書いてあったお名前・カテゴリーで作りました。会社名・写真は空です）: ' + res.unmatched.join('、');
  if (res.noPhoto.length) out += '\n新メンバー・更新メンバーで写真が見つからない方（写真なし）: ' + res.noPhoto.join('、');
  return out;
}

// === バイスプレジデントによる報告 ===
// vp … { avg, month:'2026-08', count, perWeek, from:'2026-03', to:'2026-08', total, thanks, weekCount, weekExt }
//   空の項目はテンプレートのまま。date … 開催日（「○/○速報」の日付）
var FP_NUM_ = '[0-9０-９][0-9０-９,，]*';
function applyVpReport_(parts, vp, date) {
  if (!vp) return null;
  var res = { pages: [], done: {} };
  var ym = function (s) { var m = String(s || '').match(/^(\d{4})-(\d{1,2})$/); return m ? { y: +m[1], m: +m[2] } : null; };
  var month = ym(vp.month), from = ym(vp.from), to = ym(vp.to);
  var num = function (re) { return new RegExp(FP_NUM_ + re); };
  slideOrder_(parts).forEach(function (path) {
    var xml = xmlOf_(parts, path);
    if (!xml) return;
    var t = riNorm_(slideText_(xml));
    if (t.indexOf('バイスプレジデント') < 0 || t.indexOf('報告') < 0) return;
    var before = xml;
    riShapes_(xml).forEach(function (s) {
      if (s.tag !== 'p:sp' || !s.text.trim()) return;
      var want = null;
      xml = fpMapParas_(xml, s.id, function (p, text) {
        var n = riNorm_(text), rules = [], label = null, rest = '';
        if (n.indexOf('月間リファーラル数の平均') >= 0) { label = 'avg'; rest = n.split('月間リファーラル数の平均')[1]; }
        else if (/\d{4}年\d{1,2}月から\d{4}年\d{1,2}月/.test(n) && n.indexOf('リファーラル') >= 0) {
          label = 'total'; rest = n.replace(/^.*リファーラル(?:件)?数(?:の合計)?/, '');
          if (from && to) rules.push({ re: /([0-9０-９]{4})(\s*年\s*)([0-9０-９]{1,2})(\s*月から\s*)([0-9０-９]{4})(\s*年\s*)([0-9０-９]{1,2})(\s*月)/,
            value: function (m) { return fpDigits_(from.y, m[1]) + m[2] + fpDigits_(from.m, m[3]) + m[4] + fpDigits_(to.y, m[5]) + m[6] + fpDigits_(to.m, m[7]) + m[8]; } });
        } else if (/\d{4}年\d{1,2}月の月間リファーラル数/.test(n)) {
          label = 'month'; rest = n.split('月間リファーラル数')[1];
          if (month) rules.push({ re: /([0-9０-９]{4})(\s*年\s*)([0-9０-９]{1,2})(\s*月の)/,
            value: function (m) { return fpDigits_(month.y, m[1]) + m[2] + fpDigits_(month.m, m[3]) + m[4]; } });
        } else if (n.indexOf('サンキュー額') >= 0) { label = 'thanks'; rest = n.split('サンキュー額')[1].replace(/^[（(][^）)]*[）)]/, ''); }
        var valueRules = function (key) {
          if (key === 'avg' && vp.avg) return [{ re: num('(?=\\s*件)'), value: function () { return vp.avg; } }];
          if (key === 'total' && vp.total) return [{ re: num('(?=\\s*件)'), value: function () { return vp.total; } }];
          if (key === 'month' && vp.count) {
            return [{ re: new RegExp(FP_NUM_ + '\\s*件(?:\\s*[(（]\\s*' + FP_NUM_ + '\\s*件\\s*[\\/／]\\s*週\\s*[)）])?'),
                      value: function () { return vp.count + '件' + (vp.perWeek ? '(' + vp.perWeek + '件/週)' : ''); } }];
          }
          if (key === 'thanks' && vp.thanks) {
            return [{ re: new RegExp(FP_NUM_ + '\\s*億\\s*(?:' + FP_NUM_ + '\\s*万\\s*)?円?|' + FP_NUM_ + '\\s*万\\s*円?|' + FP_NUM_ + '\\s*円'),
                      value: function () { return vp.thanks; } }];
          }
          return [];
        };
        if (label) {
          want = label;
          if (/[0-9０-９]/.test(rest)) { rules = rules.concat(valueRules(label)); want = null; }      // 見出しと同じ行に数
          else if (/[:：]$/.test(rest)) {                                                               // 「平均：」のあとに入れる
            var v = { avg: vp.avg && vp.avg + '件', total: vp.total && vp.total + '件', thanks: vp.thanks,
                      month: vp.count && vp.count + '件' + (vp.perWeek ? '(' + vp.perWeek + '件/週)' : '') }[label];
            if (v) rules.push({ re: /[:：](?=[\s　]*$)/, value: function (m) { return m[0] + v; } });
            want = null;
          }
          if (rules.length) res.done[label] = true;
          return rules.length ? fpReplaceInPara_(p, rules) : p;
        }
        if (want && /[0-9０-９]/.test(text)) {
          var vr = valueRules(want);
          if (vr.length) res.done[want] = true;
          want = null;
          return vr.length ? fpReplaceInPara_(p, vr) : p;
        }
        // 「７/２２速報　今週のリファーラル　件（うち外部　件）」
        if (/速報/.test(text)) {
          var fr = [];
          if (date) fr.push({ re: /([0-9０-９]{1,2})(\s*[\/／]\s*)([0-9０-９]{1,2})(?=\s*速報)/, value: function (m) {
            return fpDigits_(date.getMonth() + 1, m[1]) + m[2] + fpDigits_(date.getDate(), m[3]); } });
          if (vp.weekCount) fr.push({ re: /(今週のリファーラル)[\s　]*[0-9０-９,，]*[\s　]*(?=件)/, value: function (m) { return m[1] + ' ' + vp.weekCount + ' '; } });
          if (vp.weekExt) fr.push({ re: /(うち外部)[\s　]*[0-9０-９,，]*[\s　]*(?=件)/, value: function (m) { return m[1] + ' ' + vp.weekExt + ' '; } });
          if (fr.length) res.done.flash = true;
          return fr.length ? fpReplaceInPara_(p, fr) : p;
        }
        return p;
      });
    });
    if (xml !== before) { putXml_(parts, path, xml); res.pages.push(path); }
  });
  var items = [];
  if (res.done.avg) items.push('月間リファーラル数の平均 ' + vp.avg + '件');
  if (res.done.month) items.push((month ? month.m + '月の' : '') + '月間リファーラル数 ' + vp.count + '件');
  if (res.done.total) items.push('累計 ' + vp.total + '件');
  if (res.done.thanks) items.push('サンキュー額 ' + vp.thanks);
  if (res.done.flash && date) items.push('速報の日付 ' + (date.getMonth() + 1) + '/' + date.getDate());
  res.message = items.length ? 'バイスプレジデントによる報告：' + items.join('・') + ' を入れました。'
    : (res.pages.length ? '' : 'バイスプレジデントによる報告のページに、数を入れるところが見つかりませんでした。');
  return res;
}

// === ネットワーキングリーダー ===
// nl … { show: true/false, month: '2026-08', items: [{ key, value, unit, winners: [{ name, raw, category }] }] }
function fpNlKindOf_(text) {
  var t = riNorm_(text);
  if (!t || t.length > 12) return '';
  for (var i = 0; i < NL_KINDS_.length; i++) if (NL_KINDS_[i].re.test(t)) return NL_KINDS_[i].key;
  return '';
}
var FP_NL_VALUE_RE_ = /^[0-9０-９][0-9０-９,，]*(?:億[0-9０-９,，]*)?(?:万)?(?:円|件|回|名|人|pt|p|ポイント)?$/i;
// ページの種類：title（「○年／○月度」）・kind（部門のページ）・summary（まとめ）・other
function fpNlPage_(xml, W, H) {
  var shapes = riShapes_(xml), labels = [];
  shapes.forEach(function (s) { if (s.tag === 'p:sp' && fpNlKindOf_(s.text)) labels.push(s); });
  if (labels.length >= 3) return { type: 'summary', labels: labels, shapes: shapes };
  if (labels.length === 1) {
    var value = shapes.filter(function (s) { return s.tag === 'p:sp' && FP_NL_VALUE_RE_.test(riNorm_(s.text)); })[0] || null;
    // 部門の見出しと数は、お名前・会社名の枠として見ない。fpUnits_ は図形の一覧を作り直すので、図形の番号と字で見分ける。
    // 以前は一覧の中身どうしを比べていていつも外れず、カタカナだけの見出し（「サンキュー」）をお名前の枠と取り違えていた
    // （受賞者がお2人だと、見出しに1人目のお名前・数に1人目の会社名が入り、写真の枠には2人目だけが入った）
    var skip = {};
    skip[labels[0].id] = true;
    if (value) skip[value.id] = true;
    var u = fpUnits_(xml, W, H, function (s) {
      return !!skip[s.id] || !!fpNlKindOf_(s.text) || FP_NL_VALUE_RE_.test(riNorm_(s.text));
    });
    if (u.units.length) return { type: 'kind', kind: fpNlKindOf_(labels[0].text), label: labels[0], value: value, units: u.units };
  }
  var t = riNorm_(slideText_(xml));
  if (/\d{4}年/.test(t) && /月度/.test(t)) return { type: 'title' };
  return { type: 'other' };
}
// 数の図形：数字の部分と単位の部分をそれぞれ置き換える（字の大きさ・色が違っても崩れないように）。枠の幅に収める
function fpSetNlValue_(xml, value, number, unit) {
  if (!value) return xml;
  xml = fpMapParas_(xml, value.id, function (p, text) {
    if (!/[0-9０-９]/.test(text)) return p;
    return replaceInParagraph_(p, function (joined) {
      var m = /([0-9０-９][0-9０-９,，億]*)(\s*)(万円|万|円|件|回|名|人|pt|p|ポイント)?/i.exec(joined);
      if (!m) return [];
      var hits = [{ start: m.index, end: m.index + m[1].length, value: number + (m[3] ? '' : unit) }];
      if (m[3]) hits.push({ start: m.index + m[1].length + m[2].length, end: m.index + m[0].length, value: unit });
      return hits;
    });
  });
  return fpFitPara_(xml, value.id, 0, value.acx, 1);
}
// 数と単位：CEU など数える部門はテンプレートの単位（PT・件・名）、サンキューは金額の単位（万円・円）
function fpNlNumber_(item, templateText) {
  var v = String(item.value || '').trim();
  if (item.key === 'thanks') {
    var m = v.match(/^(.*?)(万円|円)$/);
    return m ? { number: m[1], unit: m[2] } : { number: v, unit: '' };
  }
  var tu = (riNorm_(templateText || '').match(/[^0-9０-９,，]+$/) || [''])[0];
  return { number: v, unit: tu ? String(templateText).trim().match(/[^0-9０-９,，\s]+$/)[0] : (item.unit || '') };
}
function applyNetworkingLeaders_(parts, nl, cache) {
  if (!nl) return null;
  var res = { shown: [], hidden: [], filled: [], noPhoto: [], unmatched: [], pairs: [], missing: [] };
  var sz = fpSize_(parts), by = fpRoster_(), pages = [];
  slideOrder_(parts).forEach(function (path) {
    var xml = xmlOf_(parts, path);
    if (xml && riNorm_(slideText_(xml)).indexOf('ネットワーキングリーダー') >= 0) pages.push({ path: path, xml: xml });
  });
  if (!pages.length) return null;
  var show = function (pg, on) {
    putXml_(parts, pg.path, setSlideShow_(xmlOf_(parts, pg.path), on));
    (on ? res.shown : res.hidden).push(pg.path);
  };
  if (!nl.show) {
    pages.forEach(function (pg) { show(pg, false); });
    res.message = 'ネットワーキングリーダーのページ（' + pages.length + '枚）は非表示にしました。';
    return res;
  }
  var ym = String(nl.month || '').match(/^(\d{4})-(\d{1,2})$/), yy = ym ? +ym[1] : 0, mm = ym ? +ym[2] : 0;
  var items = {};
  (nl.items || []).forEach(function (it) {
    var ps = (it.winners || []).map(function (w) {
      var p = fpPerson_(w, by);
      if (p && !p.matched && res.unmatched.indexOf(p.name) < 0) res.unmatched.push(p.name);
      return p;
    }).filter(Boolean);
    items[it.key] = { key: it.key, value: it.value, unit: it.unit, people: ps };
  });
  pages.forEach(function (pg) { pg.info = fpNlPage_(pg.xml, sz.W, sz.H); });
  // 見出しの年月（「2026年／６月度」「６月のネットワーキングリーダーの皆さま」）
  var monthRules = function () {
    var r = [];
    if (yy) r.push({ re: /([0-9０-９]{4})(\s*年)/, value: function (m) { return fpDigits_(yy, m[1]) + m[2]; } });
    if (mm) r.push({ re: /([0-9０-９]{1,2})(\s*月(?:度|の))/, value: function (m) { return fpDigits_(mm, m[1]) + m[2]; } });
    return r;
  };
  var byKind = {}, pairModel = null;
  pages.forEach(function (pg) {
    if (pg.info.type === 'kind') {
      (byKind[pg.info.kind] = byKind[pg.info.kind] || []).push(pg);
      if (pg.info.units.length === 2 && !pairModel) pairModel = pg;
    }
  });
  var fillKindPage = function (path, xml, info, item) {
    var rp = relsPathOf_(path), rels = xmlOf_(parts, rp);
    var nv = fpNlNumber_(item, info.value ? info.value.text : '');
    xml = fpSetNlValue_(xml, info.value, nv.number, nv.unit);
    info.units.forEach(function (u, i) {
      var f = fpFillUnit_(parts, xml, rels, u, item.people[i] || null, cache, res);
      xml = f.xml; rels = f.rels;
    });
    putXml_(parts, path, setSlideShow_(xml, true));
    if (rels) putXml_(parts, rp, rels);
    res.shown.push(path);
    item.people.slice(0, info.units.length).forEach(function (p) { res.filled.push({ path: path, kind: item.key, name: p.name }); });
  };
  var usedPages = {};
  NL_KINDS_.forEach(function (k) {
    var cands = byKind[k.key] || [], item = items[k.key];
    if (!cands.length) { if (item && item.people.length) res.missing.push(k.label); return; }
    if (!item || !item.people.length) return;                        // 受賞者のいない部門のページは非表示（あとでまとめて）
    var n = item.people.length, exact = cands.filter(function (c) { return c.info.units.length === n; })[0];
    var single = cands.filter(function (c) { return c.info.units.length === 1; })[0] || cands[0];
    if (exact) {
      fillKindPage(exact.path, exact.xml, exact.info, item);
      usedPages[exact.path] = true;
      return;
    }
    if (n === 2 && pairModel) {
      // お2人のページ（ほかの部門のもの）を複製して、この部門の見出しにする
      var xml = setSlideShow_(pairModel.xml, true), inf = pairModel.info;
      var lines = single.info.label.paras.filter(fpNonEmpty_);
      xml = setParagraphsInShape_(xml, inf.label.id, lines.length ? lines : [k.label]);
      var g = readShapeGeomEmu_(single.xml, single.info.label.id);
      if (g && !single.info.label.parent && !inf.label.parent) xml = setShapeGeomEmu_(xml, inf.label.id, g);
      var path = fpClonePages_(parts, pairModel.path, 1, single.path, xml)[0];
      fillKindPage(path, xmlOf_(parts, path), fpNlPage_(xmlOf_(parts, path), sz.W, sz.H), item);
      usedPages[path] = true;
      res.pairs.push(k.label);
      return;
    }
    // 3人以上（またはお2人のページが無い）：1人のページを人数ぶん
    var more = fpClonePages_(parts, single.path, n - 1, single.path, setSlideShow_(single.xml, true));
    [single.path].concat(more).forEach(function (path, i) {
      var one = { key: item.key, value: item.value, unit: item.unit, people: [item.people[i]] };
      fillKindPage(path, xmlOf_(parts, path), fpNlPage_(xmlOf_(parts, path), sz.W, sz.H), one);
      usedPages[path] = true;
    });
  });
  pages.forEach(function (pg) {
    var t = pg.info.type;
    if (t === 'kind') { if (!usedPages[pg.path]) show(pg, false); return; }
    var xml = xmlOf_(parts, pg.path), rules = monthRules();
    if (t === 'title' || t === 'summary') {
      riShapes_(xml).forEach(function (s) {
        if (s.tag === 'p:sp' && /[0-9０-９]/.test(s.text)) xml = fpMapParas_(xml, s.id, function (p) { return rules.length ? fpReplaceInPara_(p, rules) : p; });
      });
    }
    putXml_(parts, pg.path, xml);
    if (t === 'summary') fpNlSummary_(parts, pg.path, items, sz, cache, res);
    show(pg, true);
  });
  var kinds = NL_KINDS_.filter(function (k) { return items[k.key] && items[k.key].people.length && (byKind[k.key] || []).length; });
  res.message = 'ネットワーキングリーダー（' + (yy ? yy + '年' + mm + '月' : '') + '）：'
    + (kinds.length ? kinds.map(function (k) {
        return k.label + ' ' + items[k.key].people.map(function (p) { return p.name; }).join('・');
      }).join('／') : '受賞者の入力がありません') + '。';
  if (res.pairs.length) res.message += '\nお2人のページを作った部門: ' + res.pairs.join('、');
  if (res.emptyKinds && res.emptyKinds.length) {
    res.message += '\n該当者のいない部門（まとめのページでは見出しも出しません' + (res.emptyPacked ? '。ほかの部門を詰めて並べました' : '') + '）: '
      + NL_KINDS_.filter(function (k) { return res.emptyKinds.indexOf(k.key) >= 0; }).map(function (k) { return k.label; }).join('、');
  }
  if (res.missing.length) res.message += '\nテンプレートにページが無い部門: ' + res.missing.join('、');
  if (res.unmatched.length) res.message += '\nネットワーキングリーダーで名簿に無い方: ' + res.unmatched.join('、');
  if (res.noPhoto.length) res.message += '\nネットワーキングリーダーで写真が見つからない方（写真なし）: ' + res.noPhoto.join('、');
  return res;
}

// まとめのページ：部門の見出しの下に、写真・お名前・会社名とカテゴリー。
// 受賞者が2人以上の部門があるときは、全員が並ぶよう列を増やして並べ直す（列の幅は均等に狭める）
function fpNlSummary_(parts, path, items, sz, cache, res) {
  var xml = xmlOf_(parts, path), rp = relsPathOf_(path), rels = xmlOf_(parts, rp);
  var shapes = riShapes_(xml), pics = riPhotos_(shapes, sz.W, sz.H), cols = [], taken = {};
  shapes.forEach(function (s) {
    var k = s.tag === 'p:sp' ? fpNlKindOf_(s.text) : '';
    if (k) cols.push({ kind: k, label: s });
  });
  cols.sort(function (a, b) { return a.label.ax - b.label.ax; });
  cols.forEach(function (c) {
    var names = shapes.filter(function (s) {
      return s.tag === 'p:sp' && !taken[s.id] && s.ay > c.label.ay && riOverlap_(s, c.label) >= 0.4
          && s.paras.filter(fpNonEmpty_).length === 1 && riNameLike_(s.text) && !fpNlKindOf_(s.text);
    }).sort(function (a, b) { return a.ay - b.ay; });
    c.name = names[0] || null;
    if (!c.name) return;
    taken[c.name.id] = true;
    c.info = shapes.filter(function (s) {
      return s.tag === 'p:sp' && !taken[s.id] && s.text.trim() && s.ay >= c.name.ay + c.name.acy * 0.5
          && s.ay <= c.name.ay + c.name.acy + sz.H * 0.2 && riOverlap_(s, c.name) >= 0.3 && !fpNlKindOf_(s.text);
    });
    c.info.forEach(function (s) { taken[s.id] = true; });
    c.photo = pics.filter(function (p) {
      return !taken[p.id] && riOverlap_(p, c.name) >= 0.4 && p.ay + p.acy <= c.name.ay + c.name.acy * 0.5 && p.ay >= c.label.ay;
    })[0] || null;
    if (c.photo) taken[c.photo.id] = true;
  });
  cols = cols.filter(function (c) { return c.name; });
  if (!cols.length) return;
  var peopleOf = function (c) { return (items[c.kind] && items[c.kind].people) || []; };
  // 受賞者のいない部門（該当者なし）は、見出しも出さず、列も取らない（以前は見出しの字だけが浮いて残った）。
  // どの部門にも受賞者がいない（まだ入力が無い）ときは、そのままにする
  var any = cols.some(function (c) { return peopleOf(c).length; });
  var slotsOf = function (c) { return peopleOf(c).length || (any ? 0 : 1); };
  cols.forEach(function (c) { if (!slotsOf(c)) xml = fpHide_(xml, c.label.id, true); });
  var total = 0;
  cols.forEach(function (c) { total += slotsOf(c); });
  var grouped = cols.some(function (c) { return c.label.parent || c.name.parent || (c.photo && c.photo.parent); });
  // お2人以上の部門があるとき・空いた部門があるときは、列を並べ直す（空いたところを詰め、端から端までを等分）
  var relaid = false;
  if (cols.some(function (c) { return slotsOf(c) !== 1; }) && total && !grouped) {
    var re = fpNlRelayout_(xml, cols, slotsOf, total);
    xml = re.xml;
    cols = re.cols;
    relaid = true;
  }
  cols.forEach(function (c) {
    (c.units || [{ name: c.name, info: c.info || [], photo: c.photo }]).forEach(function (u, i) {
      var p = peopleOf(c)[i] || null;
      if (!p && i > 0) return;
      u.infoLines = 2;                                   // 狭い列なので、会社名・カテゴリーは2行まで折り返してよい
      var f = fpFillUnit_(parts, xml, rels, u, p, cache, res);
      xml = f.xml; rels = f.rels;
      if (p) res.filled.push({ path: path, kind: c.kind, name: p.name, summary: true });
    });
    // 受賞者のいない部門は、空にしたお名前・会社名の枠も隠す（色の付いた枠だけが残らないように）。写真の枠は空なら隠れている
    if (!slotsOf(c)) [c.name].concat(c.info || []).forEach(function (s) { xml = fpHide_(xml, s.id, true); });
  });
  if (any) {
    res.emptyKinds = cols.filter(function (c) { return !peopleOf(c).length; }).map(function (c) { return c.kind; });
    res.emptyPacked = relaid;
  }
  putXml_(parts, path, xml);
  if (rels) putXml_(parts, rp, rels);
}
// 列を並べ直す。受賞者の数だけ（写真・お名前・会社名とカテゴリー）の組を複製し、
// 端から端までを人数で等分した位置に置く。部門の見出しは、その部門の列の真ん中に。
// slotsOf(c) が 0 の部門（受賞者なし）は列を取らず、その場で空にするだけ（見出しは呼ぶ側で隠す）。
// 人数が部門の数より少ないときは、枠を広げずに間隔だけ広げる
function fpNlRelayout_(xml, cols, slotsOf, total) {
  var L = Infinity, R = -Infinity, maxId = 0, m, re = /<p:cNvPr\b[^>]*\sid="(\d+)"/g;
  while ((m = re.exec(xml)) !== null) maxId = Math.max(maxId, +m[1]);
  cols.forEach(function (c) {
    [c.label, c.name, c.photo].concat(c.info || []).forEach(function (s) {
      if (!s) return;
      L = Math.min(L, s.ax); R = Math.max(R, s.ax + s.acx);
    });
  });
  var slotW = (R - L) / total, colW = (R - L) / cols.length, f = Math.min(1, slotW / colW), slot = 0;
  var place = function (x, s, center, newCenter) {
    var cx = Math.round(s.acx * f), nx = Math.round(newCenter + (s.ax + s.acx / 2 - center) * f - cx / 2);
    return setShapeGeomEmu_(x, s.id, { x: nx, cx: cx });
  };
  cols.forEach(function (c) {
    var n = slotsOf(c), center = c.name.ax + c.name.acx / 2, members = [c.photo, c.name].concat(c.info || []).filter(Boolean);
    if (!n) { c.units = [{ name: c.name, info: c.info || [], photo: c.photo }]; return; }   // 受賞者なし：動かさずに空にする
    c.units = [];
    for (var i = 0; i < n; i++) {
      var nc = L + (slot + i + 0.5) * slotW, unit = { info: [] };
      members.forEach(function (s) {
        var id = s.id;
        if (i > 0) {
          var r = findShapeRange_(xml, s.id), seg = xml.substring(r.start, r.end);
          id = String(++maxId);
          seg = seg.replace(/(<p:cNvPr\b[^>]*\sid=")\d+(")/, '$1' + id + '$2').replace(/(<p:cNvPr\b[^>]*>)\s*<a:extLst>[\s\S]*?<\/a:extLst>/, '$1');
          var end = findShapeRange_(xml, members[members.length - 1].id).end;
          xml = xml.substring(0, end) + seg + xml.substring(end);
        }
        var copy = { id: id, ax: s.ax, ay: s.ay, acx: s.acx, acy: s.acy, paras: s.paras, parent: null };
        xml = place(xml, copy, center, nc);
        copy.ax = Math.round(nc + (s.ax + s.acx / 2 - center) * f - s.acx * f / 2); copy.acx = Math.round(s.acx * f);
        if (s === c.photo) unit.photo = copy;
        else if (s === c.name) unit.name = copy;
        else unit.info.push(copy);
      });
      c.units.push(unit);
    }
    var lc = L + (slot + n / 2) * slotW, lw = Math.min(c.label.acx, n * slotW);
    xml = setShapeGeomEmu_(xml, c.label.id, { x: Math.round(lc - lw / 2), cx: Math.round(lw) });
    xml = fpFitLines_(xml, c.label.id, lw);                          // 狭くなった見出しは字を小さく（行ごとにそろえて）
    slot += n;
  });
  return { xml: xml, cols: cols };
}

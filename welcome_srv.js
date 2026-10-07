// === 事前MTG（朝イチMTG）の「熱烈歓迎」のページ（新入会メンバーごとに1枚）===
//
// ルーティンチェックシートの「新入会」の欄に書いてある方ごとに、熱烈歓迎のページを作り、
// 事前MTGのパワポの、まとめのページのすぐあとに並べる（まとめのページが無ければ最初）。
//
// ひな形 … 「⚙️ 設定 ＞ 大きなスライド」の「熱烈歓迎（事前MTG・新入会メンバー）」に登録したpptx。
//          登録していなければ、同梱の既定のひな形（welcome_template.html。事前MTGの既定のひな形と同じ土台）で作る。
//          ひな形の中で {{氏名}}（または「お名前」）のあるページ（無ければ最初のページ）を、人数ぶん写して使う。
//          ひな形の土台（レイアウト・マスター・テーマ）が事前MTGのひな形と違えば、土台ごと写すので見た目は変わらない
//          （splice_srv.js の importSlideCopies_）。ページの大きさが違えば、事前MTGのパワポの大きさに合わせて縮める
//          （図形・文字の大きさも同じ割合で。splice_srv.js の fitSourceToTargetSize_）。
// 差し込み口 … {{氏名}} {{名字}} {{よみがな}} {{会社名}} {{カテゴリー}} {{チャプター}} {{月日}} {{開催日}} {{開催回}}
//          文字の「お名前」も {{氏名}} と同じに扱う（画像からひな形を作ったときなど、ふつうの言葉で置けるように）
// 写真 … 名前（または代替テキスト）が「写真」「写真1」の画像。無ければ「お写真」（または「写真」）とだけ書いた図形
//        （その図形の位置・大きさ・形のまま写真の画像にする）。それも無ければ、お名前の文字箱にいちばん近い画像。
//        メンバー写真（02_メンバー写真）に差し替えて枠に合わせて切り抜く。見つからなければ仮の画像のまま
//        （「お写真」の図形は、文字だけ消して枠はそのまま）
// 書体 … 事前MTGのパワポのテーマの書体をメイリオにするので（premtg_srv.js の buildPreMeetingDeck_）、
//        ひな形で書体を指定していない文字はメイリオになる（登録したひな形のテーマが別の書体でも）
// 名簿に無い方（まだ名簿に入っていない新メンバー）は、チェックシートに書いてあったお名前と、
// かっこの中のカテゴリー（「新井 花子さん（エステサロン）」）で作る（会社名は空）。

var WELCOME_KIND_ = 'welcome';
var WELCOME_PHOTO_MARK_ = /^(お写真|写真)$/;              // 写真の場所の目印（図形に書いてある文字がこれだけ）

// 同梱の既定のひな形（tools/build_welcome_template.py で作ったpptxをbase64にしたもの）
function welcomeBuiltinBlob_() {
  var html = HtmlService.createHtmlOutputFromFile('welcome_template').getContent();
  var m = html.match(/PPTX_BASE64_BEGIN([\s\S]*?)PPTX_BASE64_END/);
  var b64 = (m ? m[1] : html).replace(/[^A-Za-z0-9+\/=]/g, '');
  return Utilities.newBlob(Utilities.base64Decode(b64), PPTX_MIME_, 'BNI_テンプレート_熱烈歓迎.pptx');
}

// 使うひな形：登録したもの、無ければ既定のもの
function welcomeTemplateInfo_() {
  var id = PropertiesService.getScriptProperties().getProperty(BIG_TEMPLATE_KINDS_[WELCOME_KIND_].prop) || '';
  if (!id) return { registered: false, id: '', name: '既定のひな形' };
  var f;
  try { f = DriveApp.getFileById(id); }
  catch (e) { return { registered: true, id: id, name: '', error: '登録した熱烈歓迎のひな形を開けませんでした（' + (e && e.message ? e.message : e) + '）' }; }
  if (f.getMimeType && f.getMimeType() === SLIDES_MIME_) {
    return { registered: true, id: id, name: f.getName(), error: '登録した熱烈歓迎のひな形がGoogleスライドです。PowerPoint（.pptx）のファイルを登録してください' };
  }
  return { registered: true, id: id, name: f.getName() };
}
function welcomeTemplateParts_(info) {
  if (info.registered) {
    if (info.error) throw new Error(info.error + '。⚙️ 設定 ＞ 大きなスライド で登録し直してください。');
    return unzipToMap_(DriveApp.getFileById(info.id).getBlob());
  }
  return unzipToMap_(welcomeBuiltinBlob_());
}

// 既定のひな形を書き出す（デザインを変えたいときの元にする）
function exportWelcomeTemplate() {
  try {
    var name = 'BNI_テンプレート_熱烈歓迎.pptx';
    var saved = saveOutputFile_(welcomeBuiltinBlob_(), name);
    return { ok: true, url: saved.url, downloadUrl: saved.downloadUrl, fileName: name,
             message: '既定の熱烈歓迎のひな形を「03_生成物」に書き出しました。PowerPointで直してDriveに置き、'
                    + '⚙️ 設定 ＞ 大きなスライド の「熱烈歓迎（事前MTG・新入会メンバー）」にリンクを登録すると、そのデザインで作ります。' };
  } catch (e) {
    return { ok: false, message: '書き出せませんでした: ' + (e && e.message ? e.message : e) };
  }
}

// 新入会の方（ルーティンチェックシートの「新入会」の欄）→ [{ name, raw, matched, company, category, kana, byCategory }]
// 名簿の方は名簿の会社名・カテゴリー・よみがな。名簿に無い方は、書いてあったお名前とかっこの中（「カテゴリー」のうしろ）のカテゴリー。
// byCategory … かなで書かれたお名前を、カテゴリーで名簿の方に合わせた（routineRosterKana_）
function welcomeMembers_(raw) {
  var list = routineMemberList_(raw), roster = [];
  if (!list.length) return [];
  try { roster = getMemberMaster({ membersOnly: true }).members || []; } catch (e) {}
  var out = [], seen = {};
  for (var i = 0; i < list.length; i++) {
    var x = list[i], m = null;
    if (x.name) {
      var key = normName_(x.name);
      for (var k = 0; k < roster.length; k++) if (normName_(roster[k].name) === key) { m = roster[k]; break; }
    }
    var name = x.name || String(x.raw || '').replace(/[（(].*$/, '').trim().replace(/(さん|様|さま|氏)$/, '').trim();
    if (!name || seen[normName_(name)]) continue;
    seen[normName_(name)] = true;
    out.push({ name: name, raw: x.raw || '', matched: !!x.name,
               company: m ? String(m.company || '').trim() : '',
               category: m ? String(m.title || '').trim() : String(x.category || '').trim(),
               kana: m ? String(m.kana || '').trim() : '', byCategory: !!(x.name && x.byCategory) });
  }
  return out;
}

// ひな形の中の、熱烈歓迎のページ（{{氏名}}・「お名前」のあるページ。無ければ最初のページ）
function welcomeModel_(src) {
  var order = slideOrder_(src);
  for (var i = 0; i < order.length; i++) {
    var t = slideText_(xmlOf_(src, order[i]) || '');
    if (t.indexOf('{{氏名}}') >= 0 || t.indexOf('{{名字}}') >= 0 || t.indexOf('お名前') >= 0) return order[i];
  }
  return order[0] || '';
}

// 文字の「お名前」を {{氏名}} にする（PowerPointが1つの言葉をいくつかの切れ目に分けて持っていても）
function welcomeNameMarks_(xml) {
  return replacePatternsInXml_(xml, [{ re: /お名前/, value: function () { return '{{氏名}}'; } }]).xml;
}

// 写真の場所：名前（または代替テキスト）が「写真」「写真1」の画像 → 「お写真」とだけ書いた図形 → お名前の文字箱にいちばん近い画像。
// 戻り値 { id, mark }（mark … 目印の図形。写真の画像に置き換える）。どれも無ければ null
function welcomePhotoSlot_(xml, anchorId) {
  var pics = premtgShapes_(xml, 'p:pic').filter(function (p) { return !/<a:(audio|video)File\b/.test(p.seg); });
  for (var i = 0; i < pics.length; i++) {
    if (/^写真1?$/.test(pics[i].name) || /^写真1?$/.test(pics[i].descr)) return { id: pics[i].id, mark: false };
  }
  var sps = premtgShapes_(xml, 'p:sp');
  for (var j = 0; j < sps.length; j++) {
    if (WELCOME_PHOTO_MARK_.test(slideText_(sps[j].seg).replace(/[\s　]+/g, ''))) return { id: sps[j].id, mark: true };
  }
  var near = premtgPicFor_(xml, 1, anchorId);
  return near ? { id: near, mark: false } : null;
}

// 目印の図形（「お写真」）を、同じ位置・大きさ・形・枠線の写真の画像にする。target … 写真のファイル（../media/…）
function welcomeMarkToPic_(xml, rels, spId, target) {
  var r = findShapeRange_(xml, spId);
  if (!r || !rels) return { xml: xml, rels: rels };
  var seg = xml.substring(r.start, r.end);
  var spPr = (seg.match(/<p:spPr\b[^>]*>([\s\S]*?)<\/p:spPr>/) || [])[1] || '';
  var xfrm = (spPr.match(/<a:xfrm\b[^>]*>[\s\S]*?<\/a:xfrm>/) || [''])[0];
  var geom = (spPr.match(/<a:prstGeom\b[^>]*\/>|<a:prstGeom\b[^>]*>[\s\S]*?<\/a:prstGeom>|<a:custGeom\b[^>]*>[\s\S]*?<\/a:custGeom>/)
              || ['<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'])[0];
  var ln = (spPr.match(/<a:ln\b[^>]*\/>|<a:ln\b[^>]*>[\s\S]*?<\/a:ln>/) || [''])[0];
  var max = 0, m, re = /Id="rId(\d+)"/g;
  while ((m = re.exec(rels)) !== null) max = Math.max(max, parseInt(m[1], 10));
  var rid = 'rId' + (max + 1);
  rels = rels.replace('</Relationships>', '<Relationship Id="' + rid + '" Type="'
    + 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="' + target + '"/></Relationships>');
  var pic = '<p:pic><p:nvPicPr><p:cNvPr id="' + spId + '" name="写真" descr="写真"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
    + '<p:blipFill><a:blip r:embed="' + rid + '"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
    + '<p:spPr>' + xfrm + geom + ln + '</p:spPr></p:pic>';
  return { xml: xml.substring(0, r.start) + pic + xml.substring(r.end), rels: rels };
}

// 図形の文字を消す（写真が無い方の「お写真」。枠はそのまま）
function welcomeClearText_(xml, id) {
  var r = findShapeRange_(xml, id);
  if (!r) return xml;
  var seg = xml.substring(r.start, r.end).replace(/(<a:t(?=[\s>])[^>]*>)[\s\S]*?(<\/a:t>)/g, '$1$2');
  return xml.substring(0, r.start) + seg + xml.substring(r.end);
}

// ひな形を確かめる（登録したとき・作る前の確かめ）：お名前・お写真を入れる場所があるか
//   戻り値 { name: お名前の場所があるか, photo: 'pic'（写真の画像）| 'mark'（「お写真」の図形）| ''（無い）, notes: [知らせ] }
function welcomeTemplateCheck_(src) {
  var model = welcomeModel_(src);
  if (!model) return { name: false, photo: '', notes: ['ひな形にページがありません。'] };
  var xml = welcomeNameMarks_(xmlOf_(src, model) || '');
  var nameId = premtgShapeWith_(xml, '{{氏名}}') || premtgShapeWith_(xml, '{{名字}}');
  var slot = welcomePhotoSlot_(xml, nameId), notes = [];
  if (!nameId) notes.push('お名前を入れる場所がありません（「お名前」と書いた文字の箱か、{{氏名}} を置いてください）。');
  if (!slot) notes.push('お写真を入れる場所がありません（「お写真」とだけ書いた四角か、名前が「写真」の画像を置いてください）。');
  return { name: !!nameId, photo: slot ? (slot.mark ? 'mark' : 'pic') : '', notes: notes };
}

// 事前MTGのパワポに、新入会の方ごとの熱烈歓迎のページを入れる（まとめのページのすぐあと。無ければ最初）。
//   parts … 事前MTGのパワポ（書き換える）／src … 熱烈歓迎のひな形を展開したもの／members … welcomeMembers_ の結果
//   common … 月日・開催日・開催回など／cache … 写真の控え（mpAddPhoto_）／afterPath … まとめのページ
// 戻り値 { pages: [お名前], paths, noPhoto: [お名前], messages: [] }
function premtgWelcomePages_(parts, src, members, common, cache, afterPath) {
  var res = { pages: [], paths: [], noPhoto: [], messages: [] };
  if (!members || !members.length) return res;
  var model = src ? welcomeModel_(src) : '';
  if (!model) { res.messages.push('熱烈歓迎のひな形にページがありません。'); return res; }
  var fit = fitSourceToTargetSize_(parts, src);                   // ページの大きさが違えば、事前MTGのパワポの大きさに合わせる
  var imp = importSlideCopies_(parts, src, model, members.length);
  if (fit.scaled) {
    res.messages.push(fit.sameRatio
      ? '熱烈歓迎のひな形のページの大きさを、事前MTGのパワポの大きさに合わせました（' + Math.round(fit.factor * 100) + '%）。'
      : '熱烈歓迎のひな形のページの縦横の比が、事前MTGのパワポと違います。ページに収まる大きさ（' + Math.round(fit.factor * 100)
        + '%）にして、真ん中に置きました（パワポで確かめてください）。');
  } else if (imp.sizeDiffers) res.messages.push('熱烈歓迎のひな形のページの大きさが、事前MTGのひな形と違います（そのまま入れました。パワポで確かめてください）。');
  var types = {};
  for (var i = 0; i < members.length; i++) {
    var m = members[i], path = imp.slides[i].path;
    var xml = welcomeNameMarks_(xmlOf_(parts, path)), rels = xmlOf_(parts, partRelsPath_(path)) || '';
    var nameId = premtgShapeWith_(xml, '{{氏名}}') || premtgShapeWith_(xml, '{{名字}}');
    var fitIds = [nameId, premtgShapeWith_(xml, '{{会社名}}'), premtgShapeWith_(xml, '{{カテゴリー}}')];
    var slot = welcomePhotoSlot_(xml, nameId);
    if (slot) {
      var photo = null;
      try { photo = mpAddPhoto_(parts, cache, m.name); }
      catch (e) { console.warn('[WELCOME] 写真を読めませんでした: ' + m.name + ' ' + (e && e.message ? e.message : e)); }
      if (photo) {
        var target = '../media/' + photo.path.replace('ppt/media/', '');
        var set = slot.mark ? welcomeMarkToPic_(xml, rels, slot.id, target) : setPicImage_(xml, rels, slot.id, target);
        xml = set.xml; rels = set.rels;
        types[photo.path.replace(/^.*\./, '')] = true;
        var box = readShapeGeomEmu_(xml, slot.id);
        if (box && photo.width && photo.height) xml = setSrcRectInPic_(xml, slot.id, coverCrop_(photo.width, photo.height, box.cx, box.cy));
      } else {
        if (slot.mark) xml = welcomeClearText_(xml, slot.id);
        res.noPhoto.push(m.name);
      }
    }
    var vals = {};
    for (var k in common) vals[k] = common[k];
    vals['氏名'] = m.name; vals['名字'] = String(m.name).split(/[\s　]+/)[0]; vals['よみがな'] = m.kana || '';
    vals['会社名'] = m.company || ''; vals['カテゴリー'] = m.category || '';
    xml = replaceTokensInXml_(xml, vals);
    for (var f = 0; f < fitIds.length; f++) if (fitIds[f]) xml = premtgFitLine_(xml, fitIds[f], 12);   // 1行に収まる大きさに
    putXml_(parts, path, xml);
    putXml_(parts, partRelsPath_(path), rels);
    res.pages.push(m.name);
    res.paths.push(path);
  }
  var ct = xmlOf_(parts, '[Content_Types].xml');
  for (var t in types) ct = ensureDefaultType_(ct, t);
  putXml_(parts, '[Content_Types].xml', ct);
  // 並び：まとめのページのすぐあと（無ければ最初）。足したページはまだ並びに無いので、写した順にそのまま入れる
  var rest = slideEntries_(parts), list = imp.slides.slice(), out = [], placed = false;
  for (var j = 0; j < rest.length && afterPath; j++) {
    out.push(rest[j]);
    if (rest[j].path === afterPath) { out = out.concat(list, rest.slice(j + 1)); placed = true; break; }
  }
  if (!placed) out = list.concat(rest);
  setSlideEntries_(parts, out);
  return res;
}

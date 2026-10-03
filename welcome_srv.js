// === 事前MTG（朝イチMTG）の「熱烈歓迎」のページ（新入会メンバーごとに1枚）===
//
// ルーティンチェックシートの「新入会」の欄に書いてある方ごとに、熱烈歓迎のページを作り、
// 事前MTGのパワポの、まとめのページのすぐあとに並べる（まとめのページが無ければ最初）。
//
// ひな形 … 「⚙️ 設定 ＞ 大きなスライド」の「熱烈歓迎（事前MTG・新入会メンバー）」に登録したpptx。
//          登録していなければ、同梱の既定のひな形（welcome_template.html。事前MTGの既定のひな形と同じ土台）で作る。
//          ひな形の中で {{氏名}} のあるページ（無ければ最初のページ）を、人数ぶん写して使う。
//          ひな形の土台（レイアウト・マスター・テーマ）が事前MTGのひな形と違えば、土台ごと写すので見た目は変わらない
//          （splice_srv.js の importSlideCopies_）。
// 差し込み口 … {{氏名}} {{名字}} {{よみがな}} {{会社名}} {{カテゴリー}} {{チャプター}} {{月日}} {{開催日}} {{開催回}}
// 写真 … 名前（または代替テキスト）が「写真」「写真1」の画像。無ければ、お名前の文字箱にいちばん近い画像。
//        メンバー写真（02_メンバー写真）に差し替えて枠に合わせて切り抜く。見つからなければ仮の画像のまま
// 名簿に無い方（まだ名簿に入っていない新メンバー）は、チェックシートに書いてあったお名前と、
// かっこの中のカテゴリー（「新井 花子さん（エステサロン）」）で作る（会社名は空）。

var WELCOME_KIND_ = 'welcome';

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

// 新入会の方（ルーティンチェックシートの「新入会」の欄）→ [{ name, raw, matched, company, category, kana }]
// 名簿の方は名簿の会社名・カテゴリー・よみがな。名簿に無い方は、書いてあったお名前とかっこの中のカテゴリー
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
               kana: m ? String(m.kana || '').trim() : '' });
  }
  return out;
}

// ひな形の中の、熱烈歓迎のページ（{{氏名}} のあるページ。無ければ最初のページ）
function welcomeModel_(src) {
  var order = slideOrder_(src);
  for (var i = 0; i < order.length; i++) {
    var t = slideText_(xmlOf_(src, order[i]) || '');
    if (t.indexOf('{{氏名}}') >= 0 || t.indexOf('{{名字}}') >= 0) return order[i];
  }
  return order[0] || '';
}

// 写真の枠：名前（または代替テキスト）が「写真」「写真1」の画像。無ければ、お名前の文字箱にいちばん近い画像
function welcomePicFor_(xml, anchorId) {
  var pics = premtgShapes_(xml, 'p:pic').filter(function (p) { return !/<a:(audio|video)File\b/.test(p.seg); });
  for (var i = 0; i < pics.length; i++) if (pics[i].name === '写真' || pics[i].descr === '写真') return pics[i].id;
  return premtgPicFor_(xml, 1, anchorId);
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
  var imp = importSlideCopies_(parts, src, model, members.length);
  if (imp.sizeDiffers) res.messages.push('熱烈歓迎のひな形のページの大きさが、事前MTGのひな形と違います（そのまま入れました。パワポで確かめてください）。');
  var types = {};
  for (var i = 0; i < members.length; i++) {
    var m = members[i], path = imp.slides[i].path;
    var xml = xmlOf_(parts, path), rels = xmlOf_(parts, partRelsPath_(path)) || '';
    var nameId = premtgShapeWith_(xml, '{{氏名}}') || premtgShapeWith_(xml, '{{名字}}');
    var fitIds = [nameId, premtgShapeWith_(xml, '{{会社名}}'), premtgShapeWith_(xml, '{{カテゴリー}}')];
    var picId = welcomePicFor_(xml, nameId);
    if (picId) {
      var photo = null;
      try { photo = mpAddPhoto_(parts, cache, m.name); }
      catch (e) { console.warn('[WELCOME] 写真を読めませんでした: ' + m.name + ' ' + (e && e.message ? e.message : e)); }
      if (photo) {
        var set = setPicImage_(xml, rels, picId, '../media/' + photo.path.replace('ppt/media/', ''));
        xml = set.xml; rels = set.rels;
        types[photo.path.replace(/^.*\./, '')] = true;
        var box = readShapeGeomEmu_(xml, picId);
        if (box && photo.width && photo.height) xml = setSrcRectInPic_(xml, picId, coverCrop_(photo.width, photo.height, box.cx, box.cy));
      } else {
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

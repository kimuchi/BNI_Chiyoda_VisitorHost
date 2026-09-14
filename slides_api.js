// === 大きなpptxをサーバー側だけで扱う仕組み（Googleスライド経由）===
//
// 【考え方】
// 35MBのpptxをブラウザとサーバーの間でやり取りしようとすると分割送信が必要になるが、
// テンプレートも生成物もDrive上に置けば、巨大なバイト列は一度もスクリプトを通らない。
//   ・テンプレートの複製  … Drive側でコピーされる（makeCopy / Drive.Files.copy）
//   ・pptx→スライド変換   … Drive側で変換される（copy時にmimeTypeを指定）
//   ・本文の書き換え      … SlidesApp（API呼び出し。ファイルサイズに依存しない）
//   ・受け渡し            … URLだけ（数十バイト）
// これでファイルサイズの制約も google.script.run のペイロード上限も無関係になる。

var BIG_TEMPLATE_KINDS_ = {
  meetingFirst:  { prop: 'BNI_TPL_MEETING_FIRST_ID',  label: '定例会スライド（前半）' },
  meetingSecond: { prop: 'BNI_TPL_MEETING_SECOND_ID', label: '定例会スライド（後半）' },
  memberPresen:  { prop: 'BNI_TPL_MEMBER_PRESEN_ID',  label: 'メンバープレゼン' }
};
var SLIDES_MIME_ = 'application/vnd.google-apps.presentation';
var PPTX_MIME_ = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

// 大きなテンプレートは「アップロード」ではなく「Drive上のファイルを指定」で登録する。
// （アップロードだと35MBがscriptを通ってしまうため）
function registerBigTemplate(kind, linkOrId) {
  try {
    var def = BIG_TEMPLATE_KINDS_[kind];
    if (!def) return { ok: false, message: 'テンプレートの種類が不正です。' };
    var id = extractDriveId_(linkOrId);
    if (!id) return { ok: false, message: 'ファイルIDを認識できませんでした。Driveの共有リンクまたはIDを貼り付けてください。' };
    var file;
    try { file = DriveApp.getFileById(id); }
    catch (e) { return { ok: false, message: 'このIDのファイルを開けませんでした。アクセス権をご確認ください。' }; }
    var mime = file.getMimeType();
    if (mime !== SLIDES_MIME_ && mime !== PPTX_MIME_) {
      return { ok: false, message: 'PowerPoint(.pptx)またはGoogleスライドのファイルを指定してください（現在: ' + mime + '）。' };
    }
    PropertiesService.getScriptProperties().setProperty(def.prop, id);
    console.log('[BIGTPL] ' + kind + ' -> ' + id + ' (' + mime + ')');
    return { ok: true, message: '「' + def.label + '」に「' + file.getName() + '」を登録しました。', status: getBigTemplateStatus() };
  } catch (e) {
    console.error('[BIGTPL] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '登録に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

function getBigTemplateStatus() {
  var props = PropertiesService.getScriptProperties(), list = [];
  for (var k in BIG_TEMPLATE_KINDS_) {
    var def = BIG_TEMPLATE_KINDS_[k], id = props.getProperty(def.prop) || '';
    var row = { kind: k, label: def.label, registered: false, fileName: '', url: '', mimeType: '', isSlides: false, sizeMB: 0 };
    if (id) {
      try {
        var f = DriveApp.getFileById(id);
        row.registered = true; row.fileName = f.getName(); row.url = f.getUrl();
        row.mimeType = f.getMimeType(); row.isSlides = (row.mimeType === SLIDES_MIME_);
        row.sizeMB = Math.round((f.getSize() || 0) / 104857.6) / 10;
      } catch (e) {}
    }
    list.push(row);
  }
  return { ok: true, templates: list };
}

// pptx を Googleスライドに変換した複製を作る。変換はDrive側で行われ、
// ファイルの中身はスクリプトを通らない。
function copyAsSlides_(srcFileId, newName, folder) {
  var resource = { name: newName, mimeType: SLIDES_MIME_ };
  if (folder) resource.parents = [folder.getId()];
  var copied = Drive.Files.copy(resource, srcFileId);       // Drive側でコピー＋変換
  return DriveApp.getFileById(copied.id);
}

// Googleスライドをそのまま複製する（変換不要な場合。これもDrive側で完結）
function copySlides_(srcFileId, newName, folder) {
  return DriveApp.getFileById(srcFileId).makeCopy(newName, folder);
}

// 登録テンプレートから「編集用の複製」を1本作る。pptxなら変換もここで行う。
function makeWorkingCopy_(kind, newName) {
  var def = BIG_TEMPLATE_KINDS_[kind];
  if (!def) throw new Error('テンプレートの種類が不正です。');
  var id = PropertiesService.getScriptProperties().getProperty(def.prop);
  if (!id) throw new Error('「' + def.label + '」のテンプレートが未登録です。');
  var src = DriveApp.getFileById(id), folder = getAssetFolder_('output');
  var mime = src.getMimeType();
  var copy = (mime === SLIDES_MIME_) ? copySlides_(id, newName, folder) : copyAsSlides_(id, newName, folder);
  copy.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return copy;
}

// 生成物の受け渡しはURLだけ。pptxが要る場合もブラウザがGoogleから直接落とす
// （スクリプトがバイト列を扱わないので、サイズに関係なく一瞬で終わる）
function slidesUrls_(fileId) {
  return {
    slidesUrl: 'https://docs.google.com/presentation/d/' + fileId + '/edit',
    pptxUrl:   'https://docs.google.com/presentation/d/' + fileId + '/export/pptx',
    pdfUrl:    'https://docs.google.com/presentation/d/' + fileId + '/export/pdf'
  };
}

// --- SlidesApp による本文差し替え（サイズ非依存）---

// {{キー}} を一括置換する。1回のAPI呼び出しで全スライドに効くため大きな資料でも速い。
function replaceTokens_(presentation, map) {
  var n = 0;
  for (var key in map) {
    var v = map[key] == null ? '' : String(map[key]);
    n += presentation.replaceAllText('{{' + key + '}}', v, false);
  }
  return n;
}

// 代替テキスト(alt text)を目印にして図形の文字を差し替える。
// テンプレート側で図形の「代替テキスト（説明）」に名前を入れておく運用。
function setTextByAlt_(presentation, altName, text) {
  var slides = presentation.getSlides(), hit = 0;
  for (var i = 0; i < slides.length; i++) {
    var els = slides[i].getPageElements();
    for (var j = 0; j < els.length; j++) {
      if (els[j].getDescription() !== altName) continue;
      try { els[j].asShape().getText().setText(text == null ? '' : String(text)); hit++; }
      catch (e) {}
    }
  }
  return hit;
}

// 代替テキストを目印に画像を差し替える（写真の入れ替え用）
function replaceImageByAlt_(presentation, altName, blob) {
  var slides = presentation.getSlides(), hit = 0;
  for (var i = 0; i < slides.length; i++) {
    var imgs = slides[i].getImages();
    for (var j = 0; j < imgs.length; j++) {
      if (imgs[j].getDescription() !== altName) continue;
      try { imgs[j].replace(blob, true); hit++; } catch (e) {}
    }
  }
  return hit;
}

// 不要なスライドを消す / 残す（コアバリューの週替わりなど）
function keepOnlySlidesWithAlt_(presentation, altNamePrefix, keepName) {
  var slides = presentation.getSlides(), removed = 0;
  for (var i = slides.length - 1; i >= 0; i--) {
    var els = slides[i].getPageElements(), tag = '';
    for (var j = 0; j < els.length; j++) {
      var d = els[j].getDescription();
      if (d && d.indexOf(altNamePrefix) === 0) { tag = d; break; }
    }
    if (tag && tag !== keepName) { slides[i].remove(); removed++; }
  }
  return removed;
}

// --- 変換の見え方を1回だけ確かめるための機能 ---
// pptx→Googleスライド変換でレイアウトが崩れないかを、実物で確認してもらう。
function previewTemplateConversion(kind) {
  try {
    var def = BIG_TEMPLATE_KINDS_[kind];
    if (!def) return { ok: false, message: 'テンプレートの種類が不正です。' };
    var copy = makeWorkingCopy_(kind, '【変換確認】' + def.label);
    var u = slidesUrls_(copy.getId());
    return { ok: true,
      message: '変換した見本を作りました。レイアウトが崩れていないかご確認ください。問題なければ、この方式で毎週の生成ができます。',
      url: u.slidesUrl, pptxUrl: u.pptxUrl, fileId: copy.getId() };
  } catch (e) {
    console.error('[BIGTPL] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '変換の確認に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// テンプレート内で使える目印（代替テキスト・{{トークン}}）を一覧する。
// テンプレートにどんな差し込み口があるかを把握するための補助。
function inspectTemplatePlaceholders(kind) {
  try {
    var def = BIG_TEMPLATE_KINDS_[kind];
    if (!def) return { ok: false, message: 'テンプレートの種類が不正です。' };
    var id = PropertiesService.getScriptProperties().getProperty(def.prop);
    if (!id) return { ok: false, message: '「' + def.label + '」が未登録です。' };
    var f = DriveApp.getFileById(id), workId = id;
    if (f.getMimeType() !== SLIDES_MIME_) {
      var tmp = copyAsSlides_(id, '【一時】' + def.label, getAssetFolder_('output'));
      workId = tmp.getId();
    }
    var pres = SlidesApp.openById(workId), slides = pres.getSlides();
    var alts = {}, tokens = {}, slideCount = slides.length;
    for (var i = 0; i < slides.length; i++) {
      var els = slides[i].getPageElements();
      for (var j = 0; j < els.length; j++) {
        var d = els[j].getDescription();
        if (d) alts[d] = (alts[d] || 0) + 1;
        try {
          var t = els[j].asShape().getText().asString(), m, re = /\{\{([^}]{1,40})\}\}/g;
          while ((m = re.exec(t)) !== null) tokens[m[1]] = (tokens[m[1]] || 0) + 1;
        } catch (e) {}
      }
    }
    if (workId !== id) { try { DriveApp.getFileById(workId).setTrashed(true); } catch (e) {} }
    var altList = [], tokList = [];
    for (var a in alts) altList.push({ name: a, count: alts[a] });
    for (var k in tokens) tokList.push({ name: k, count: tokens[k] });
    return { ok: true, slideCount: slideCount, altTexts: altList, tokens: tokList,
             message: 'スライド' + slideCount + '枚、代替テキスト' + altList.length + '種、{{トークン}}' + tokList.length + '種を検出しました。' };
  } catch (e) {
    console.error('[BIGTPL] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '解析に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// Googleスライドへのアクセス権限を確実に取得するためのメニュー用関数。
// モーダルダイアログ内の google.script.run は認可画面を出せないため、
// 初回だけメニューから実行してもらい、ここで権限を確定させる。
function authorizeSlidesAccess() {
  var ui = SpreadsheetApp.getUi();
  try {
    var folder = getAssetFolder_('output');
    var tmp = SlidesApp.create('【権限確認】BNI一時ファイル');
    var id = tmp.getId();
    tmp.saveAndClose();
    try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {}
    ui.alert('スライド機能の権限確認',
      '✅ Googleスライドへのアクセスが有効です。\n\n' +
      '保存先フォルダ: ' + folder.getName() + '\n\n' +
      'このまま「⚙️ 大きなスライドの登録」からテンプレートを登録できます。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('スライド機能の権限確認',
      '❌ 権限の取得に失敗しました。\n\n' + (e && e.message ? e.message : e) +
      '\n\n画面に権限の許可を求めるダイアログが出た場合は「許可」を選んでから、もう一度実行してください。', ui.ButtonSet.OK);
  }
}

function openBigTemplateDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('big_templates').setWidth(680).setHeight(680),
    '大きなスライド（定例会・メンバープレゼン）の登録');
}

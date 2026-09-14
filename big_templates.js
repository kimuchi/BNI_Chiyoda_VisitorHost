// === 大きなpptxの登録（Driveのファイルをリンクで指定する）===
//
// 数十MBのテンプレートをブラウザからアップロードさせると、ファイルがscriptを通るため
// 時間がかかり失敗しやすい。Drive上のファイルをIDで指定すれば、加工時に
// サーバーが直接 getBlob() で読むだけで済み、ブラウザとの間でバイト列が動かない。

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

// Googleスライドをそのまま複製する（変換不要な場合。これもDrive側で完結）

// 登録テンプレートから「編集用の複製」を1本作る。pptxなら変換もここで行う。

// 生成物の受け渡しはURLだけ。pptxが要る場合もブラウザがGoogleから直接落とす
// （スクリプトがバイト列を扱わないので、サイズに関係なく一瞬で終わる）

// --- SlidesApp による本文差し替え（サイズ非依存）---

// {{キー}} を一括置換する。1回のAPI呼び出しで全スライドに効くため大きな資料でも速い。

// 代替テキスト(alt text)を目印にして図形の文字を差し替える。
// テンプレート側で図形の「代替テキスト（説明）」に名前を入れておく運用。

// 代替テキストを目印に画像を差し替える（写真の入れ替え用）

// 不要なスライドを消す / 残す（コアバリューの週替わりなど）

// --- 変換の見え方を1回だけ確かめるための機能 ---
// pptx→Googleスライド変換でレイアウトが崩れないかを、実物で確認してもらう。

// テンプレート内で使える目印（代替テキスト・{{トークン}}）を一覧する。
// テンプレートにどんな差し込み口があるかを把握するための補助。

// Googleスライドへのアクセス権限を確実に取得するためのメニュー用関数。
// モーダルダイアログ内の google.script.run は認可画面を出せないため、
// 初回だけメニューから実行してもらい、ここで権限を確定させる。


// === pptxのままサーバー側で加工する ===
//
// 【要点】巨大なバイト列をJavaScriptの配列にしないこと。
// Utilities.unzip() が返すのはBlobの配列で、これをそのまま Utilities.zip() に渡す限り、
// 画像などの中身はApps Script内部で持ち回されるだけでJS側のメモリに展開されない。
// getBytes() / getDataAsString() を呼ぶのは、書き換えが必要な小さなXMLだけに限る。
// これによりブラウザとの分割送信もGoogleスライドへの変換も不要になる。

function getBigTemplateFile_(kind) {
  var def = BIG_TEMPLATE_KINDS_[kind];
  if (!def) throw new Error('テンプレートの種類が不正です。');
  var id = PropertiesService.getScriptProperties().getProperty(def.prop);
  if (!id) throw new Error('「' + def.label + '」のテンプレートが未登録です。');
  return DriveApp.getFileById(id);
}

// テンプレートを読み、editFn で必要なXMLだけ書き換えて、Driveへ保存する。
// editFn(xmlMap, ctx) の中では xmlOf_/putXml_ で個別パートだけを触ること。
function editPptxOnServer_(kind, outName, editFn) {
  var t0 = new Date().getTime();
  var file = getBigTemplateFile_(kind);
  var blob = file.getBlob();                       // Driveから直接。ブラウザを経由しない
  var t1 = new Date().getTime();

  var parts = Utilities.unzip(blob.setContentType('application/zip'));
  var map = {}, count = 0;
  for (var i = 0; i < parts.length; i++) {
    var nm = parts[i].getName();
    if (!nm || nm.charAt(nm.length - 1) === '/') continue;   // ディレクトリは捨てる
    map[nm] = parts[i];                            // Blobのまま保持（getBytesしない）
    count++;
  }
  var t2 = new Date().getTime();

  var info = editFn(map, { fileName: file.getName(), partCount: count });
  var t3 = new Date().getTime();

  var out = zipFromMap_(map, outName);
  var t4 = new Date().getTime();

  var saved = saveOutputFile_(out, outName);
  var t5 = new Date().getTime();

  var timing = { 読込: t1 - t0, 展開: t2 - t1, 書換: t3 - t2, 再梱包: t4 - t3, 保存: t5 - t4, 合計: t5 - t0 };
  console.log('[PPTX] ' + outName + ' parts=' + count + ' timing=' + JSON.stringify(timing));
  return { saved: saved, partCount: count, timing: timing, info: info };
}

// 実測用。テンプレートを「何も変えずに」展開→再梱包→保存して所要時間を測る。
// 35MBのファイルがサーバー側の制限（6分・メモリ）に収まるかを、推測ではなく実物で確認する。
function benchmarkPptxRoundTrip(kind) {
  try {
    var def = BIG_TEMPLATE_KINDS_[kind];
    if (!def) return { ok: false, message: 'テンプレートの種類が不正です。' };
    var file = getBigTemplateFile_(kind);
    var sizeMB = Math.round((file.getSize() || 0) / 104857.6) / 10;
    var r = editPptxOnServer_(kind, '【動作確認】' + def.label + '.pptx', function (map) { return null; });
    var t = r.timing;
    return { ok: true,
      message: '成功しました。' + sizeMB + 'MB / ' + r.partCount + 'パーツを ' + (t.合計 / 1000).toFixed(1) + '秒で往復できました。\n' +
               '内訳: 読込' + (t.読込/1000).toFixed(1) + '秒 / 展開' + (t.展開/1000).toFixed(1) + '秒 / ' +
               '再梱包' + (t.再梱包/1000).toFixed(1) + '秒 / 保存' + (t.保存/1000).toFixed(1) + '秒\n' +
               'この方式で毎週の生成ができます（実行上限は6分）。',
      url: r.saved.url, sizeMB: sizeMB, partCount: r.partCount, timing: t, seconds: t.合計 / 1000 };
  } catch (e) {
    console.error('[PPTX] benchmark ' + (e && e.stack ? e.stack : e));
    return { ok: false,
      message: 'この方式では扱えませんでした: ' + (e && e.message ? e.message : e) +
               '\n\nファイルサイズを小さくする（画像の解像度を下げる）か、前半・後半をさらに分割すると通る場合があります。' };
  }
}

// テンプレートの中身の内訳を調べる。どこに容量を使っているかが分かると、
// 軽量化で往復時間を縮められるかの判断ができる。
function analyzePptxContents(kind) {
  try {
    var def = BIG_TEMPLATE_KINDS_[kind];
    if (!def) return { ok: false, message: 'テンプレートの種類が不正です。' };
    var file = getBigTemplateFile_(kind);
    var parts = Utilities.unzip(file.getBlob().setContentType('application/zip'));
    var groups = {}, slideCount = 0, total = 0, big = [], skipped = 0;
    for (var i = 0; i < parts.length; i++) {
      var nm = parts[i].getName();
      if (!nm || nm.charAt(nm.length - 1) === '/') continue;
      // パート単位でサイズを測る。巨大な画像1枚でJS配列が膨らむのを避けるため、
      // 失敗しても解析全体は止めずに「測定不可」として続行する。
      var sz = 0, unmeasured = false;
      try { sz = parts[i].getBytes().length; } catch (szErr) { unmeasured = true; skipped++; }
      total += sz;
      if (/^ppt\/slides\/slide\d+\.xml$/.test(nm)) slideCount++;
      var g = /^ppt\/media\//.test(nm) ? 'メディア（画像・動画）'
            : /^ppt\/slides\//.test(nm) ? 'スライド本文'
            : /^ppt\/(slideLayouts|slideMasters|theme)\//.test(nm) ? 'レイアウト・テーマ'
            : /^ppt\/notesSlides\//.test(nm) ? 'ノート' : 'その他';
      groups[g] = (groups[g] || 0) + sz;
      if (sz > 1048576) big.push({ name: nm, mb: Math.round(sz / 104857.6) / 10 });
    }
    var list = [];
    for (var k in groups) list.push({ group: k, mb: Math.round(groups[k] / 104857.6) / 10,
                                      pct: Math.round(groups[k] / total * 100) });
    list.sort(function (a, b) { return b.mb - a.mb; });
    big.sort(function (a, b) { return b.mb - a.mb; });
    return { ok: true, slideCount: slideCount, totalMB: Math.round(total / 104857.6) / 10,
             groups: list, bigFiles: big.slice(0, 10),
             message: 'スライド' + slideCount + '枚・展開後 約' + (Math.round(total / 104857.6) / 10) + 'MB です。'
                      + (skipped ? '（' + skipped + '個のパートはサイズ測定できませんでした）' : '') };
  } catch (e) {
    console.error('[PPTX] analyze ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '解析に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

function openBigTemplateDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('big_templates').setWidth(700).setHeight(700),
    '大きなスライド（定例会・メンバープレゼン）の登録');
}

// 初回の権限取得用。モーダル内では認可画面を出せないため、メニューから1回実行する。
function authorizeDriveForBigFiles() {
  var ui = SpreadsheetApp.getUi();
  try {
    var folder = getAssetFolder_('output');
    ui.alert('大きなファイル機能の権限確認',
      '✅ 準備できています。\n\n保存先フォルダ: ' + folder.getName() +
      '\n\n「⚙️ 大きなスライドの登録」でテンプレートを登録し、「動作確認」を実行してください。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('大きなファイル機能の権限確認',
      '❌ 失敗しました: ' + (e && e.message ? e.message : e) +
      '\n\n権限の許可を求める画面が出た場合は「許可」を選んでから、もう一度実行してください。', ui.ButtonSet.OK);
  }
}

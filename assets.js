// === BNI 素材フォルダ（共有Drive）の共通基盤 ===
// 共有フォルダのIDを1つ登録すると、その直下に素材用サブフォルダを自動作成し、
// テンプレート(pptx)・メンバー写真・生成物をすべてそこに保存する。

var ASSET_ROOT_KEY_ = 'BNI_ASSET_ROOT_FOLDER_ID';
var ASSET_SUB_ = { template: '01_テンプレート', photo: '02_メンバー写真', output: '03_生成物' };

// PowerPointテンプレートの種類と、slide1.xml に必須のシェイプID
var TEMPLATE_KINDS_ = {
  presen: { prop: 'BNI_TPL_PRESEN_ID', label: 'ビジタープレゼン（1人1枚）', ids: ['25', '27', '29'] },
  intro:  { prop: 'BNI_TPL_INTRO_ID',  label: 'ビジター紹介（3人1枚）',   ids: ['19','20','21','25','26','28','31','32','33'] },
  dairi:  { prop: 'BNI_TPL_DAIRI_ID',  label: '代理紹介（3人1枚）',       ids: ['19','20','21','25','26','28','31','32','33'] }
};

// 共有リンク or ファイル/フォルダID から ID を取り出す（既存 processMemberListFromDrive と同じ流儀）
function extractDriveId_(s) {
  if (!s) return '';
  s = String(s).trim();
  var m = s.match(/\/folders\/([-\w]{15,})/) || s.match(/\/d\/([-\w]{15,})/) || s.match(/[?&]id=([-\w]{15,})/);
  if (m) return m[1];
  return /^[-\w]{15,}$/.test(s) ? s : '';
}

function getAssetSettings() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(ASSET_ROOT_KEY_) || '';
  var res = { ok: true, folderId: id, folderName: '', folderUrl: '', reachable: false, subFolders: [], message: '' };
  if (!id) { res.message = '未設定です。フォルダIDまたは共有リンクを入力して保存してください。'; return res; }
  try {
    var f = DriveApp.getFolderById(id);
    res.folderName = f.getName();
    res.folderUrl = f.getUrl();
    res.reachable = true;
    for (var k in ASSET_SUB_) {
      var sub = findChildFolder_(f, ASSET_SUB_[k]);
      res.subFolders.push({ kind: k, name: ASSET_SUB_[k], exists: !!sub, url: sub ? sub.getUrl() : '' });
    }
  } catch (e) {
    res.message = 'このIDのフォルダを開けません。IDが正しいか、あなたがアクセスできる共有フォルダかをご確認ください。';
  }
  return res;
}

function saveAssetSettings(linkOrId) {
  try {
    var id = extractDriveId_(linkOrId);
    if (!id) return { ok: false, message: 'フォルダIDを認識できませんでした。共有リンク（.../folders/xxxx）またはIDを入力してください。' };
    var folder;
    try { folder = DriveApp.getFolderById(id); }
    catch (e) { return { ok: false, message: 'このIDのフォルダを開けませんでした。IDが正しいか、アクセス権があるかをご確認ください。' }; }
    PropertiesService.getScriptProperties().setProperty(ASSET_ROOT_KEY_, id);
    // サブフォルダを用意
    for (var k in ASSET_SUB_) ensureChildFolder_(folder, ASSET_SUB_[k]);
    return { ok: true, message: '「' + folder.getName() + '」を素材フォルダに設定し、サブフォルダを用意しました。', status: getAssetSettings() };
  } catch (e) {
    console.error('[ASSET] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '保存中にエラーが発生しました: ' + (e && e.message ? e.message : e) };
  }
}

function findChildFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : null;
}
function ensureChildFolder_(parent, name) {
  return findChildFolder_(parent, name) || parent.createFolder(name);
}

// 素材ルート。未設定ならスプレッドシートの親フォルダにフォールバック（設定しなくても動く）
function getAssetRootFolder_() {
  var id = PropertiesService.getScriptProperties().getProperty(ASSET_ROOT_KEY_);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  return file.getParents().hasNext() ? file.getParents().next() : DriveApp.getRootFolder();
}
// kind: 'template' | 'photo' | 'output'
function getAssetFolder_(kind) {
  return ensureChildFolder_(getAssetRootFolder_(), ASSET_SUB_[kind] || ASSET_SUB_.output);
}

// 生成物を 03_生成物 に保存し、リンク共有を付けてURLを返す
function saveOutputFile_(blob, fileName) {
  var folder = getAssetFolder_('output');
  // 同名ファイルがあれば置き換えではなく更新（URLを変えない）
  var it = folder.getFilesByName(fileName), file;
  if (it.hasNext()) {
    file = it.next();
    try { Drive.Files.update({}, file.getId(), blob); }
    catch (e) { file.setTrashed(true); file = folder.createFile(blob.setName(fileName)); }
  } else {
    file = folder.createFile(blob.setName(fileName));
  }
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { id: file.getId(), url: file.getUrl(), name: fileName };
}

// === PowerPointテンプレートの登録 ===

function getTemplateStatus() {
  var props = PropertiesService.getScriptProperties(), list = [];
  for (var k in TEMPLATE_KINDS_) {
    var def = TEMPLATE_KINDS_[k], id = props.getProperty(def.prop) || '', row = {
      kind: k, label: def.label, requiredIds: def.ids.join(', '), registered: false, fileName: '', url: ''
    };
    if (id) {
      try { var f = DriveApp.getFileById(id); row.registered = true; row.fileName = f.getName(); row.url = f.getUrl(); }
      catch (e) {}
    }
    list.push(row);
  }
  return { ok: true, templates: list, assets: getAssetSettings() };
}

// テンプレートpptxを検証して 01_テンプレート に保存する
function saveTemplateBase64(kind, base64, fileName) {
  try {
    var def = TEMPLATE_KINDS_[kind];
    if (!def) return { ok: false, message: 'テンプレートの種類が不正です。' };
    if (!base64) return { ok: false, message: 'ファイルデータが空です。' };
    var blob = Utilities.newBlob(Utilities.base64Decode(base64), 'application/zip', fileName || (kind + '.pptx'));
    var check = validateTemplate_(blob, def.ids);
    if (!check.ok) return check;

    var folder = getAssetFolder_('template');
    var saveName = kind + '_' + (fileName || 'template.pptx');
    var it = folder.getFilesByName(saveName), file;
    var pptBlob = Utilities.newBlob(Utilities.base64Decode(base64),
      'application/vnd.openxmlformats-officedocument.presentationml.presentation', saveName);
    if (it.hasNext()) { file = it.next(); try { Drive.Files.update({}, file.getId(), pptBlob); } catch (e) { file.setTrashed(true); file = folder.createFile(pptBlob); } }
    else { file = folder.createFile(pptBlob); }
    PropertiesService.getScriptProperties().setProperty(def.prop, file.getId());
    console.log('[TPL] saved ' + kind + ' -> ' + file.getId());
    return { ok: true, message: '「' + def.label + '」のテンプレートを登録しました。', status: getTemplateStatus() };
  } catch (e) {
    console.error('[TPL] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '登録中にエラーが発生しました: ' + (e && e.message ? e.message : e) };
  }
}

// pptxの必須パーツと、slide1.xml 内の必須シェイプIDを確認する
function validateTemplate_(zipBlob, requiredIds) {
  var map;
  try { map = unzipToMap_(zipBlob); }
  catch (e) { return { ok: false, message: 'pptxファイルとして開けませんでした。PowerPointの .pptx をお選びください。' }; }
  var need = ['ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels', '[Content_Types].xml',
              'ppt/slides/slide1.xml', 'ppt/slides/_rels/slide1.xml.rels'];
  for (var i = 0; i < need.length; i++) {
    if (!map[need[i]]) return { ok: false, message: '対応するテンプレートではありません（' + need[i] + ' が見つかりません）。' };
  }
  var xml = map['ppt/slides/slide1.xml'].getDataAsString('UTF-8');
  var found = {}, m, re = /<p:cNvPr[^>]*\sid="(\d+)"/g;
  while ((m = re.exec(xml)) !== null) found[m[1]] = true;
  var missing = [];
  for (var j = 0; j < requiredIds.length; j++) if (!found[requiredIds[j]]) missing.push(requiredIds[j]);
  if (missing.length) {
    return { ok: false, message: '選んだ種類とテンプレートのレイアウトが一致しません（シェイプID ' + missing.join(', ') + ' が見つかりません）。' };
  }
  return { ok: true, hasNotes: !!map['ppt/notesSlides/notesSlide1.xml'] };
}

function getTemplateBlob_(kind) {
  var def = TEMPLATE_KINDS_[kind];
  if (!def) return null;
  var id = PropertiesService.getScriptProperties().getProperty(def.prop);
  if (!id) return null;
  try { return DriveApp.getFileById(id).getBlob(); } catch (e) { return null; }
}

// === メンバー写真の管理 ===

// 氏名の正規化（全角/半角・空白を吸収して照合しやすくする）
function normName_(x) {
  return String(x || '').normalize('NFKC').replace(/[\s　]/g, '').trim();
}

// 写真を 02_メンバー写真 に保存。items = [{name, base64, mimeType}]
function uploadMemberPhotosBase64(items) {
  try {
    if (!items || !items.length) return { ok: false, message: '写真が選択されていません。' };
    var folder = getAssetFolder_('photo'), saved = 0, names = [];
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      if (!it || !it.base64) continue;
      var blob = Utilities.newBlob(Utilities.base64Decode(it.base64), it.mimeType || 'image/jpeg', it.name || ('photo' + i + '.jpg'));
      var ex = folder.getFilesByName(blob.getName());
      if (ex.hasNext()) { var f = ex.next(); try { Drive.Files.update({}, f.getId(), blob); } catch (e) { f.setTrashed(true); folder.createFile(blob); } }
      else folder.createFile(blob);
      saved++; names.push(blob.getName());
    }
    console.log('[PHOTO] saved=' + saved);
    return { ok: true, message: saved + '枚の写真を保存しました。', saved: saved, names: names };
  } catch (e) {
    console.error('[PHOTO] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '写真の保存中にエラーが発生しました: ' + (e && e.message ? e.message : e) };
  }
}

// 写真フォルダを走査して「写真索引」シートを作り直す
function rebuildPhotoIndex() {
  try {
    var folder = getAssetFolder_('photo'), it = folder.getFiles(), rows = [];
    while (it.hasNext()) {
      var f = it.next(), fn = f.getName();
      if (!/\.(jpe?g|png|gif|webp)$/i.test(fn)) continue;
      // 「氏名__ハッシュ.jpg」「氏名.jpg」どちらにも対応。__ 以降とスペースを落として氏名とする
      var base = fn.replace(/\.[^.]+$/, '');
      var nm = base.split('__')[0];
      rows.push([fn, normName_(nm), f.getId()]);
    }
    rows.sort(function (a, b) { return a[0] < b[0] ? -1 : (a[0] > b[0] ? 1 : 0); });
    var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName('写真索引');
    if (!sh) { sh = ss.insertSheet('写真索引'); sh.hideSheet(); }
    sh.clear();
    sh.appendRow(['ファイル名', '氏名(正規化)', 'ファイルID']);
    if (rows.length) sh.getRange(2, 1, rows.length, 3).setValues(rows);
    console.log('[PHOTO] index rebuilt: ' + rows.length);
    return { ok: true, message: '写真索引を作り直しました（' + rows.length + '件）。', total: rows.length };
  } catch (e) {
    console.error('[PHOTO] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '索引の作成中にエラーが発生しました: ' + (e && e.message ? e.message : e) };
  }
}

// 氏名から写真のファイルIDを引く（索引優先、無ければフォルダ走査）
function findPhotoIdForName_(name) {
  var key = normName_(name);
  if (!key) return '';
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName('写真索引');
  if (sh) {
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) if (String(data[i][1]) === key) return String(data[i][2]);
  }
  return '';
}

// 氏名配列 → {氏名: dataURI} のマップ（重いのでクライアントから20件ずつ呼ぶ）
function getMemberPhotosBase64(names) {
  try {
    var map = {}, miss = [];
    for (var i = 0; i < (names || []).length; i++) {
      var id = findPhotoIdForName_(names[i]);
      if (!id) { miss.push(names[i]); continue; }
      try {
        var b = DriveApp.getFileById(id).getBlob();
        map[names[i]] = 'data:' + (b.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(b.getBytes());
      } catch (e) { miss.push(names[i]); }
    }
    return { ok: true, map: map, missing: miss };
  } catch (e) {
    return { ok: false, message: '写真の取得に失敗しました: ' + (e && e.message ? e.message : e), map: {} };
  }
}

function getPhotoOverview() {
  var idx = { total: 0 };
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('写真索引');
    idx.total = sh ? Math.max(0, sh.getLastRow() - 1) : 0;
  } catch (e) {}
  var members = [], unmatched = [];
  try {
    members = getMemberMaster().members || [];
    for (var i = 0; i < members.length; i++) if (!findPhotoIdForName_(members[i].name)) unmatched.push(members[i].name);
  } catch (e) {}
  return { ok: true, indexed: idx.total, memberCount: members.length, unmatched: unmatched, assets: getAssetSettings() };
}

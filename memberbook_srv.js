// === メンバーブック編集・PDF出力 ／ Zoom入室案内 ===
// データはすべて「メンバー名簿」シート。写真はDriveの 02_メンバー写真 から氏名照合で解決する。

function openMemberBookEditorDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createTemplateFromFile('memberbook_editor').evaluate().setWidth(1150).setHeight(780),
    'メンバーブックの編集・PDF出力');
}
function openZoomGuideDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('zoom_guide').setWidth(560).setHeight(500),
    'Zoom入室案内の作成');
}

// 編集画面の初期データ。写真は重いので含めず、あとから分割で取得する
function getMemberBookData() {
  try {
    var m = getMemberMaster();
    if (!m.ok) return m;
    return { ok: true, members: m.members, cover: m.cover, categories: m.categories };
  } catch (e) {
    console.error('[MBOOK] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '読み込みに失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

function saveMemberBookData(members, cover) { return saveMemberMaster(members, cover); }

// 1人ぶんだけ名簿に書く（メンバーブックの編集画面の「反映」。押したらすぐ名簿に残る）。
// 直す前の氏名（origName）で行を探し、編集画面で直せる列だけを書き換える（ほかの列・ほかの方の行には触らない）。
// 見つからない（画面で足したばかりの方）ときは、名簿の最後に足す
var MB_EDIT_FIELDS_ = { cat: '業種区分', name: '氏名', title: 'カテゴリー', company: '会社名', role: '役職',
                        comment: '一言コメント', refer: '紹介してほしい人', collab: '協業したい人', position: '会社での役職' };
function saveMemberBookMember(origName, m) {
  var lock = LockService.getScriptLock();
  try {
    if (!m || !String(m.name || '').trim()) return { ok: false, message: '氏名が空です。' };
    if (!lock.tryLock(15000)) return { ok: false, message: 'ほかの方が保存中です。少し待ってから、もう一度「反映」を押してください。' };
    var sh = ensureMemberSheet_(), n = MEMBER_HEADERS_.length;
    // 以前の名簿には「会社での役職」の列が無い → 見出しを足す
    var head = sh.getRange(1, 1, 1, n).getValues()[0];
    if (String(head[n - 1] || '') !== MEMBER_HEADERS_[n - 1]) {
      sh.getRange(1, 1, 1, n).setValues([MEMBER_HEADERS_]).setFontWeight('bold').setBackground('#f2f6ff');
    }
    var last = sh.getLastRow(), rows = last > 1 ? sh.getRange(2, 1, last - 1, n).getValues() : [];
    var key = normName_(origName || m.name), at = -1;
    for (var i = 0; i < rows.length; i++) if (normName_(rows[i][2]) === key) { at = i; break; }
    var row = at >= 0 ? rows[at] : MEMBER_HEADERS_.map(function () { return ''; });
    Object.keys(MB_EDIT_FIELDS_).forEach(function (k) {
      if (m[k] !== undefined) row[MEMBER_HEADERS_.indexOf(MB_EDIT_FIELDS_[k])] = String(m[k] == null ? '' : m[k]);
    });
    if (at >= 0) sh.getRange(at + 2, 1, 1, n).setValues([row]);
    else sh.getRange(last + 1, 1, 1, n).setValues([row]);
    return { ok: true, added: at < 0, message: '「' + String(m.name).trim() + '」を名簿に保存しました。' };
  } catch (e) {
    console.error('[MBOOK] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '保存に失敗しました: ' + (e && e.message ? e.message : e) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// 配布用の単一HTMLをDriveに書き出す。
// 現行ツールの「自分自身のouterHTMLを書き換えて再ダウンロード」方式は、
// 描画済みのDOMまで保存されてファイルが肥大化するため採用しない。
function exportMemberBookHtml(html, fileName) {
  try {
    if (!html) return { ok: false, message: '出力する内容がありません。' };
    var name = (fileName || 'BNI_memberbook') + '.html';
    var blob = Utilities.newBlob(html, 'text/html', name).setName(name);
    var saved = saveOutputFile_(blob, name);
    return { ok: true, message: '配布用HTMLを保存しました。', url: saved.url, downloadUrl: saved.downloadUrl, fileName: name };
  } catch (e) {
    console.error('[MBOOK] ' + (e && e.stack ? e.stack : e));
    return { ok: false, message: '保存に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// 印刷したPDFをDriveに残したい場合の受け口（ブラウザで保存したPDFを送る）
function saveMemberBookPdfBase64(base64, fileName) {
  try {
    if (!base64) return { ok: false, message: 'ファイルデータが空です。' };
    var name = (fileName || 'BNI_memberbook') + '.pdf';
    var blob = Utilities.newBlob(Utilities.base64Decode(base64), 'application/pdf', name);
    var saved = saveOutputFile_(blob, name);
    return { ok: true, message: 'PDFを保存しました。', url: saved.url };
  } catch (e) {
    return { ok: false, message: '保存に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

// Zoom入室案内。氏名を行順で6列に割る（現行と同じ配分）
// 番号は名簿のNoをそのまま使う。並び順から振り直すと、退会などでNoが飛んでいるときに
// 実際の番号とずれてしまい、「番号抜け／番号違いにご注意ください」と案内する紙自体が
// 間違うことになるため。Noが空のメンバーだけ並び順から補う。
function getZoomGuideContext() {
  try {
    var members = getMemberMaster().members || [];
    var props = PropertiesService.getScriptProperties();
    return { ok: true,
      members: members.map(function (m, i) {
        return { no: String(m.no || (i + 1)), name: m.name, cat: m.title || '' };
      }),
      exampleNum: props.getProperty('BNI_ZOOM_EXAMPLE_NUM') || '4',
      title: props.getProperty('BNI_ZOOM_TITLE') || (members.length + '名') };
  } catch (e) {
    return { ok: false, message: '読み込みに失敗しました: ' + (e && e.message ? e.message : e), members: [] };
  }
}
function saveZoomGuideSettings(data) {
  try {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('BNI_ZOOM_EXAMPLE_NUM', String((data && data.exampleNum) || '4'));
    props.setProperty('BNI_ZOOM_TITLE', String((data && data.title) || ''));
    return { ok: true, message: '設定を保存しました。' };
  } catch (e) {
    return { ok: false, message: '保存に失敗しました: ' + (e && e.message ? e.message : e) };
  }
}

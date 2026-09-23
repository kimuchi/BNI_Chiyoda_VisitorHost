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

// 配布用の単一HTMLをDriveに書き出す。
// 現行ツールの「自分自身のouterHTMLを書き換えて再ダウンロード」方式は、
// 描画済みのDOMまで保存されてファイルが肥大化するため採用しない。
function exportMemberBookHtml(html, fileName) {
  try {
    if (!html) return { ok: false, message: '出力する内容がありません。' };
    var name = (fileName || 'BNI_memberbook') + '.html';
    var blob = Utilities.newBlob(html, 'text/html', name).setName(name);
    var saved = saveOutputFile_(blob, name);
    return { ok: true, message: '配布用HTMLを保存しました。', url: saved.url, fileName: name };
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

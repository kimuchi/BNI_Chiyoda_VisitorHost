// === メニュー画面（ホーム）===
// メニューの項目が増えたため、全機能を一覧できる画面を用意する。
// あわせて「使える状態になっているか」を点検して表示し、つまずきを先回りで防ぐ。

function openHomeDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('menu_home').setWidth(920).setHeight(760),
    'BNI 名簿システム メニュー');
}

// 各機能を「使える状態か」を点検する。失敗しても画面は出したいので個別にtry/catchする
function getHomeStatus() {
  var st = { ok: true, checks: [], latest: {} };
  function add(key, label, ready, detail, fixLabel, fixFn) {
    st.checks.push({ key: key, label: label, ready: !!ready, detail: detail || '',
                     fixLabel: fixLabel || '', fixFn: fixFn || '' });
  }

  // 素材フォルダ
  try {
    var a = getAssetSettings();
    add('assets', '素材フォルダ', a.folderId && a.reachable,
        a.folderId ? (a.reachable ? a.folderName : '開けません（権限をご確認ください）')
                   : '未設定（スプレッドシートと同じ場所に保存されます）',
        '設定する', 'openAssetSettingsDialog');
  } catch (e) { add('assets', '素材フォルダ', false, '確認できませんでした', '設定する', 'openAssetSettingsDialog'); }

  // メンバーリスト（割り振り用）
  try {
    var ml = getMembersList();
    add('memberlist', 'メンバーリスト（割り振り用）', ml.length > 0,
        ml.length ? ml.length + '名' : '未登録', '取り込む', 'openPdfDialog');
  } catch (e) { add('memberlist', 'メンバーリスト（割り振り用）', false, '確認できませんでした', '取り込む', 'openPdfDialog'); }

  // メンバー名簿（メンバーブック・スライド用）
  var memberCount = 0;
  try {
    var mm = getMemberMaster();
    memberCount = (mm.members || []).length;
    add('membermaster', 'メンバー名簿（冊子・スライド用）', memberCount > 0,
        memberCount ? memberCount + '名' : '未登録', '登録する', 'openMemberMasterDialog');
  } catch (e) { add('membermaster', 'メンバー名簿（冊子・スライド用）', false, '確認できませんでした', '登録する', 'openMemberMasterDialog'); }

  // メンバー写真
  try {
    var po = getPhotoOverview();
    var unmatched = (po.unmatched || []).length;
    add('photos', 'メンバー写真', po.indexed > 0 && unmatched === 0,
        po.indexed ? (po.indexed + '枚' + (unmatched ? '（未照合 ' + unmatched + '名）' : '（全員照合済み）'))
                   : '未登録', '管理する', 'openMemberPhotoDialog');
  } catch (e) { add('photos', 'メンバー写真', false, '確認できませんでした', '管理する', 'openMemberPhotoDialog'); }

  // ビジター用テンプレート
  try {
    var ts = getTemplateStatus().templates, done = 0;
    for (var i = 0; i < ts.length; i++) if (ts[i].registered) done++;
    add('tpl', 'ビジター用テンプレート', done === ts.length,
        done + ' / ' + ts.length + ' 件 登録済み', '登録する', 'openTemplateFileDialog');
  } catch (e) { add('tpl', 'ビジター用テンプレート', false, '確認できませんでした', '登録する', 'openTemplateFileDialog'); }

  // 大きなスライド
  try {
    var bs = getBigTemplateStatus().templates, bd = 0;
    for (var j = 0; j < bs.length; j++) if (bs[j].registered) bd++;
    add('bigtpl', '大きなスライド（定例会など）', bd > 0,
        bd + ' / ' + bs.length + ' 件 登録済み', '登録する', 'openBigTemplateDialog');
  } catch (e) { add('bigtpl', '大きなスライド（定例会など）', false, '確認できませんでした', '登録する', 'openBigTemplateDialog'); }

  // Gemini
  try {
    var api = getApiSettings();
    add('gemini', 'Gemini API', !!api.apiKey,
        api.apiKey ? ('設定済み（' + api.modelName + '）') : '未設定', '設定する', 'openApiSettingsDialog');
  } catch (e) { add('gemini', 'Gemini API', false, '確認できませんでした', '設定する', 'openApiSettingsDialog'); }

  try { st.latest = getPdfLinks(); } catch (e) { st.latest = {}; }
  return st;
}

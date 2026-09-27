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
  var st = { ok: true, checks: [], latest: {}, title: '' };
  function add(key, label, ready, detail, fixLabel, fixFn) {
    st.checks.push({ key: key, label: label, ready: !!ready, detail: detail || '',
                     fixLabel: fixLabel || '', fixFn: fixFn || '' });
  }

  // チャプター（名前・期・回数）。空のスプレッドシートから始めたときは、まず設定してもらう
  try {
    var ch = chapterInfo_(), next = (getMeetingCandidates()[0] || {}).display || '';
    st.title = chapterSystemTitle_();
    add('chapter', 'チャプター', ch.saved || !chapterFresh_(),
        chapterLabel_() + '・いまの期 ' + roleTermOf_(new Date()) + '期' + (next ? '・次回 ' + next : '')
        + (ch.saved ? '' : '（初期値）'), '設定する', 'openChapterSettingsDialog');
  } catch (e) { add('chapter', 'チャプター', false, '確認できませんでした', '設定する', 'openChapterSettingsDialog'); }

  // ルーティンチェックシート（次回の開催日の列があるか）。空のスプレッドシートから始めたときは、
  // チャプターの設定を保存したときに作る（setup_srv.js）
  try {
    var nx = getMeetingCandidates()[0], hitR = nx ? findRoutineColumn_(parseDate_(nx.dateValue)) : null;
    var wait = chapterFresh_() && !chapterInfo_().saved;
    add('routine', 'ルーティンチェックシート', !!hitR,
        hitR ? hitR.name : (wait ? 'チャプターの設定を保存すると作ります'
                                 : (nx ? '次回（' + nx.display + '）の列がありません' : '次回の開催日がありません')),
        wait ? '設定する' : '作る', wait ? 'openChapterSettingsDialog' : 'menuCreateMissingSheets');
  } catch (e) { add('routine', 'ルーティンチェックシート', false, '確認できませんでした', '作る', 'menuCreateMissingSheets'); }

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

  // 公式ファイルから作る雛形（ビジター用・メンバープレゼン・定例会）。雛形がそろっていれば準備できている
  try {
    var os = getOfficialTemplateStatus(), odone = (os.kinds || []).filter(function (k) { return k.registered; }).length;
    var ofile = os.file && !os.file.error ? os.file.name : '';
    add('official', '雛形（公式ファイルから）', odone === (os.kinds || []).length,
        (ofile ? '公式ファイル: ' + ofile + '・' : '公式ファイル未設定・') + '雛形 ' + odone + ' / ' + (os.kinds || []).length + ' 件 登録済み',
        '作る', 'openOfficialTemplatesDialog');
  } catch (e) { add('official', '雛形（公式ファイルから）', false, '確認できませんでした', '作る', 'openOfficialTemplatesDialog'); }

  // ビジター用テンプレート
  try {
    var ts = getTemplateStatus().templates, done = 0;
    for (var i = 0; i < ts.length; i++) if (ts[i].registered) done++;
    add('tpl', 'ビジター用テンプレート', done === ts.length,
        done + ' / ' + ts.length + ' 件 登録済み', '登録する', 'openTemplateFileDialog');
  } catch (e) { add('tpl', 'ビジター用テンプレート', false, '確認できませんでした', '登録する', 'openTemplateFileDialog'); }

  // 大きなスライド
  try {
    // 事前MTGのように、登録しなくても同梱の既定のひな形で作れるものは、登録済みと同じに数える
    var bs = getBigTemplateStatus().templates, bd = 0, builtin = [];
    for (var j = 0; j < bs.length; j++) {
      if (bs[j].registered) bd++;
      else if (bs[j].builtin) { bd++; builtin.push(bs[j].label); }
    }
    add('bigtpl', '大きなスライド（定例会など）', bd > builtin.length,
        bd + ' / ' + bs.length + ' 件 登録済み' + (builtin.length ? '（' + builtin.join('・') + 'は既定のひな形）' : ''),
        '登録する', 'openBigTemplateDialog');
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

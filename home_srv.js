// === メニュー画面（ホーム）===
// メニューの項目が増えたため、全機能を一覧できる画面を用意する。
// あわせて「使える状態になっているか」を点検して表示し、つまずきを先回りで防ぐ。

function openHomeDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('menu_home').setWidth(920).setHeight(760),
    'BNI 名簿システム メニュー');
}

// ダイアログとメニューバーからの実行で、同じアカウント・版かを照合する。
// メール送信や外部接続の動作テスト、共有設定の変更はしない。
function homeExecutionInfo_() {
  var info = { version: getSystemVersion(), activeUser: '', effectiveUser: '', authorization: '不明' };
  try { info.activeUser = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  try { info.effectiveUser = Session.getEffectiveUser().getEmail() || ''; } catch (e) {}
  try {
    info.authorization = String(ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL).getAuthorizationStatus());
  } catch (e) {}
  return info;
}

function homeExecutionSummary_(info) {
  return '版: ' + info.version + '\n利用者: ' + (info.activeUser || '取得できません')
       + '\n実行者: ' + (info.effectiveUser || '取得できません') + '\nGoogleの承認状態: ' + info.authorization;
}

function diagnoseHomeStatus() {
  var st = getHomeStatus(), lines = [homeExecutionSummary_(st.execution), '', '【設定の点検】'];
  st.checks.forEach(function(c) {
    lines.push((c.error ? '確認不可' : (c.ready ? '設定済み' : '要確認')) + ': ' + c.label + '\n  ' + c.detail);
  });
  lines.push('', st.latestError ? 'PDF一覧: 確認不可（' + st.latestError + '）' : 'PDF一覧: 取得できました');
  lines.push('', 'これは設定の点検です。PDF・スライドの生成やメール送信の成功を保証するものではありません。',
    'ここでは取得できるのにメニュー画面だけ承認エラーになる場合は、画面を閉じて開き直してください。',
    '改善しなければ、使うGoogleアカウントだけのChromeプロファイルでスプレッドシートを開いて確認してください。');
  var ui = SpreadsheetApp.getUi();
  ui.alert('準備状況と実行アカウント', lines.join('\n'), ui.ButtonSet.OK);
}

// 各機能を「使える状態か」を点検する。失敗しても画面は出したいので個別にtry/catchする
function getHomeStatus() {
  var st = { ok: true, checks: [], latest: {}, title: '', execution: homeExecutionInfo_() };
  function add(key, label, ready, detail, fixLabel, fixFn) {
    st.checks.push({ key: key, label: label, ready: !!ready, detail: detail || '',
                     fixLabel: fixLabel || '', fixFn: fixFn || '' });
  }
  function failed(key, label, e, fixLabel, fixFn) {
    var message = String(e && e.message ? e.message : e);
    add(key, label, false, '確認できませんでした: ' + message, fixLabel, fixFn);
    st.checks[st.checks.length - 1].error = message;
  }
  function templateErrors(rows) {
    return rows.filter(function(t) { return !!t.unreadable; })
      .map(function(t) { return t.label + ': ' + t.unreadable; }).join('\n');
  }

  // チャプター（名前・期・回数）。空のスプレッドシートから始めたときは、まず設定してもらう
  try {
    var ch = chapterInfo_(), next = (getMeetingCandidates()[0] || {}).display || '';
    st.title = chapterSystemTitle_();
    add('chapter', 'チャプター', ch.saved || !chapterFresh_(),
        chapterLabel_() + '・いまの期 ' + roleTermOf_(new Date()) + '期' + (next ? '・次回 ' + next : '')
        + (ch.saved ? '' : '（初期値）'), '設定する', 'openChapterSettingsDialog');
  } catch (e) { failed('chapter', 'チャプター', e, '設定する', 'openChapterSettingsDialog'); }

  // ルーティンチェックシート（次回の開催日の列があるか）。空のスプレッドシートから始めたときは、
  // チャプターの設定を保存したときに作る（setup_srv.js）
  try {
    var nx = getMeetingCandidates()[0], hitR = nx ? findRoutineColumn_(parseDate_(nx.dateValue)) : null;
    var wait = chapterFresh_() && !chapterInfo_().saved;
    add('routine', 'ルーティンチェックシート', !!hitR,
        hitR ? hitR.name : (wait ? 'チャプターの設定を保存すると作ります'
                                 : (nx ? '次回（' + nx.display + '）の列がありません' : '次回の開催日がありません')),
        wait ? '設定する' : '作る', wait ? 'openChapterSettingsDialog' : 'menuCreateMissingSheets');
  } catch (e) { failed('routine', 'ルーティンチェックシート', e, '作る', 'menuCreateMissingSheets'); }

  // 素材フォルダ
  try {
    var a = getAssetSettings();
    add('assets', '素材フォルダ', a.folderId && a.reachable,
        a.folderId ? (a.reachable ? a.folderName : (a.message || '開けません（権限をご確認ください）'))
                   : '未設定（スプレッドシートと同じ場所に保存されます）',
        '設定する', 'openAssetSettingsDialog');
  } catch (e) { failed('assets', '素材フォルダ', e, '設定する', 'openAssetSettingsDialog'); }

  // メンバーリスト（割り振り用）
  try {
    var ml = getMembersList();
    add('memberlist', 'メンバーリスト（割り振り用）', ml.length > 0,
        ml.length ? ml.length + '名' : '未登録', '取り込む', 'openPdfDialog');
  } catch (e) { failed('memberlist', 'メンバーリスト（割り振り用）', e, '取り込む', 'openPdfDialog'); }

  // メンバー名簿（メンバーブック・スライド用）
  var memberCount = 0;
  try {
    var mm = getMemberMaster();
    memberCount = (mm.members || []).length;
    add('membermaster', 'メンバー名簿（冊子・スライド用）', memberCount > 0,
        memberCount ? memberCount + '名' : '未登録', '登録する', 'openMemberMasterDialog');
  } catch (e) { failed('membermaster', 'メンバー名簿（冊子・スライド用）', e, '登録する', 'openMemberMasterDialog'); }

  // メンバー写真
  try {
    var po = getPhotoOverview();
    var unmatched = (po.unmatched || []).length;
    add('photos', 'メンバー写真', po.indexed > 0 && unmatched === 0,
        po.indexed ? (po.indexed + '枚' + (unmatched ? '（未照合 ' + unmatched + '名）' : '（全員照合済み）'))
                   : '未登録', '管理する', 'openMemberPhotoDialog');
  } catch (e) { failed('photos', 'メンバー写真', e, '管理する', 'openMemberPhotoDialog'); }

  // ビジター用テンプレート
  try {
    var ts = getTemplateStatus().templates, done = 0;
    for (var i = 0; i < ts.length; i++) if (ts[i].registered) done++;
    var tplErrors = templateErrors(ts);
    if (tplErrors) failed('tpl', 'ビジター用テンプレート', tplErrors);
    else add('tpl', 'ビジター用テンプレート', done === ts.length,
        done + ' / ' + ts.length + ' 件 登録済み', '登録する', 'openTemplateFileDialog');
  } catch (e) { failed('tpl', 'ビジター用テンプレート', e, '登録する', 'openTemplateFileDialog'); }

  // 大きなスライド
  try {
    // 事前MTGのように、登録しなくても同梱の既定のひな形で作れるものは、登録済みと同じに数える
    var bs = getBigTemplateStatus().templates, bd = 0, builtin = [];
    for (var j = 0; j < bs.length; j++) {
      if (bs[j].registered) bd++;
      else if (bs[j].builtin) { bd++; builtin.push(bs[j].label); }
    }
    var bigErrors = templateErrors(bs);
    if (bigErrors) failed('bigtpl', '大きなスライド（定例会など）', bigErrors);
    else add('bigtpl', '大きなスライド（定例会など）', bd > builtin.length,
        bd + ' / ' + bs.length + ' 件 登録済み' + (builtin.length ? '（' + builtin.join('・') + 'は既定のひな形）' : ''),
        '登録する', 'openBigTemplateDialog');
  } catch (e) { failed('bigtpl', '大きなスライド（定例会など）', e, '登録する', 'openBigTemplateDialog'); }

  // Gemini
  try {
    var api = getApiSettings();
    add('gemini', 'Gemini API', !!api.apiKey,
        api.apiKey ? ('設定済み（' + api.modelName + '）') : '未設定', '設定する', 'openApiSettingsDialog');
  } catch (e) { failed('gemini', 'Gemini API', e, '設定する', 'openApiSettingsDialog'); }

  try { st.latest = getPdfLinks(); } catch (e) { st.latestError = String(e && e.message ? e.message : e); }
  return st;
}

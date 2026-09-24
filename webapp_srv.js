// === ウェブアプリとして開く ===
//
// スプレッドシートのダイアログは googleusercontent.com のiframeで表示されるため、
// 複数のGoogleアカウントにログインしていると中身が真っ白になることがある。
// ウェブアプリなら script.google.com の「普通のタブ」で開くので、その影響を受けない。
//
// 中身は既存のダイアログHTMLをそのまま使う。google.script.run はウェブアプリでも
// 同じように動くので、画面もサーバー関数も作り直す必要はない。
//
// デプロイ手順は showWebAppUrl() の案内を参照。

// このスプレッドシートのID。
// ウェブアプリには「開いているスプレッドシート」が無いので、
// getActiveSpreadsheet() が null になる。IDを覚えておいて openById で開く。
var SS_ID_KEY_ = 'BNI_SPREADSHEET_ID';

function getSS_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) {
    // メニューから使われたときにIDを控えておく（ウェブアプリはこれを頼りにする）
    try {
      var props = PropertiesService.getScriptProperties();
      if (props.getProperty(SS_ID_KEY_) !== ss.getId()) props.setProperty(SS_ID_KEY_, ss.getId());
    } catch (e) {}
    return ss;
  }
  var id = PropertiesService.getScriptProperties().getProperty(SS_ID_KEY_);
  if (!id) {
    throw new Error('対象のスプレッドシートが分かりません。'
      + '一度スプレッドシートを開いて「名簿システム」メニューを表示してから、もう一度お試しください。');
  }
  return SpreadsheetApp.openById(id);
}

// ウェブアプリに出す機能の一覧。キーは表示するHTMLファイル名。
var WEBAPP_PAGES_ = [
  { group: '毎週の作業', items: [
    { key: 'dialog',            label: 'CSVから名簿・PDF作成',   desc: '参加者のCSVを取り込んで名簿とPDFを作ります' },
    { key: 'email',             label: 'メールの確認・一括送信', desc: '案内メールをまとめて送ります' },
    { key: 'allocation',        label: 'ルーム・オリエン割り振り表', desc: 'ビジターごとの担当を決めます' },
    { key: 'pdf_links',         label: '作成済みPDFの確認',      desc: '作ったPDFをもう一度開きます' }
  ]},
  { group: 'スライド・冊子', items: [
    { key: 'slides_visitor',    label: 'ビジター・代理スライド作成', desc: '紹介・プレゼンのPowerPointを作ります' },
    { key: 'slides_meeting',    label: '定例会スライドの自動更新', desc: '更新状況などを差し込みます' },
    { key: 'memberbook_editor', label: 'メンバーブックの編集・PDF出力', desc: '冊子の中身を編集して出力します' },
    { key: 'zoom_guide',        label: 'Zoom入室案内の作成',     desc: '表示名のお願いを作ります' }
  ]},
  { group: 'シートの整理', items: [
    { key: 'archive',           label: 'アーカイブして整理する', desc: '古い開催日のシートを隠します' }
  ]},
  { group: '設定', items: [
    { key: 'asset_settings',    label: 'BNI 素材フォルダ',       desc: '写真やテンプレートの保存先' },
    { key: 'member_master',     label: 'メンバー名簿',           desc: 'すべての機能が参照する正本' },
    { key: 'member_photos',     label: 'メンバー写真',           desc: '氏名で自動照合されます' },
    { key: 'template_files',    label: 'PowerPointテンプレート', desc: 'ビジター用の3種類' },
    { key: 'big_templates',     label: '大きなスライド',         desc: '定例会などDriveリンクで登録' },
    { key: 'spreading',         label: 'Spreadingから名簿を更新', desc: 'MCPで取得した内容を貼り付け' },
    { key: 'pdf',               label: 'メンバーリスト(OCR)の更新', desc: 'PDFから名簿を読み取ります' },
    { key: 'memberbook',        label: 'メンバーブック(PDF)の更新', desc: '配布用PDFを差し替えます' },
    { key: 'ai_documents',      label: 'AI参考資料',             desc: '自動割り振りの参考資料' },
    { key: 'holiday',           label: '休会日',                 desc: '開催しない日' },
    { key: 'template',          label: 'メールテンプレート',     desc: '案内メールの文面' },
    { key: 'allocation_note',   label: '割り振り表の特記事項',   desc: '割り振り表の下に出る文章' },
    { key: 'visitor_host',      label: 'ビジターホスト・優先順位', desc: '割り振りの優先度' },
    { key: 'api_settings',      label: 'Gemini API・モデル',     desc: 'OCRとAI割り振りで使用' }
  ]},
  { group: 'ヘルプ', items: [
    { key: 'manual',            label: '使い方（マニュアル）',   desc: 'すべての操作の説明' }
  ]}
];

function findWebAppPage_(key) {
  for (var g = 0; g < WEBAPP_PAGES_.length; g++) {
    var items = WEBAPP_PAGES_[g].items;
    for (var i = 0; i < items.length; i++) if (items[i].key === key) return items[i];
  }
  return null;
}

// テンプレートから他のHTMLを読み込むための決まり文句
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function doGet(e) {
  try {
    var key = (e && e.parameter && e.parameter.p) ? String(e.parameter.p) : '';
    var page = key ? findWebAppPage_(key) : null;

    if (!page) {
      var t = HtmlService.createTemplateFromFile('webapp_home');
      t.groups = WEBAPP_PAGES_;
      t.status = webAppStatus_();
      // 画面は googleusercontent.com のiframeの中で動くため、
      // href="?p=..." のような相対リンクだとiframeのURLを基準にしてしまい、
      // まったく別の場所へ飛んで真っ白になる。必ず絶対URLを使う。
      t.appUrl = getWebAppUrl_();
      return t.evaluate()
        .setTitle('Activeチャプター 名簿システム')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }

    // 既存のダイアログHTMLをそのまま表示する。
    // ダイアログ用に作られているので、上に「メニューに戻る」の帯だけ足す。
    var content = HtmlService.createTemplateFromFile(page.key).evaluate().getContent();
    var home = getWebAppUrl_() || '?';
    var bar = '<div style="position:sticky;top:0;z-index:9999;background:#16233f;color:#fff;'
            + 'padding:8px 14px;font-family:sans-serif;font-size:13px;display:flex;'
            + 'align-items:center;gap:12px;">'
            + '<a href="' + escapeHtmlText_(home) + '" target="_top" '
            + 'style="color:#ffd200;text-decoration:none;font-weight:bold;">'
            + '← メニューに戻る</a>'
            + '<span style="opacity:.85;">' + escapeHtmlText_(page.label) + '</span></div>';
    content = content.replace(/(<body[^>]*>)/i, '$1' + bar);

    return HtmlService.createHtmlOutput(content)
      .setTitle(page.label + ' | 名簿システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    // 何も出ないと原因が分からないので、画面にそのまま出す
    console.error('[WEBAPP] ' + (err && err.stack ? err.stack : err));
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><meta charset="utf-8"></head>'
      + '<body style="font-family:sans-serif;padding:24px;line-height:1.8;">'
      + '<h2 style="color:#c00;margin:0 0 12px;">画面を表示できませんでした</h2>'
      + '<pre style="white-space:pre-wrap;background:#f6f6f6;padding:12px;border-radius:6px;'
      + 'font-size:12px;">' + escapeHtmlText_(err && err.stack ? err.stack : String(err)) + '</pre>'
      + '<p style="font-size:13px;color:#555;">この内容をそのままお知らせください。</p>'
      + '</body></html>').setTitle('エラー | 名簿システム');
  }
}

// ウェブアプリが「どのスプレッドシートに繋がっているか」を確かめる。
// ここが繋がっていないと、どの画面も動かない。
function webAppStatus_() {
  var st = { ok: false, name: '', url: '', user: '', message: '' };
  try { st.user = Session.getEffectiveUser().getEmail() || ''; } catch (e) {}
  try {
    var ss = getSS_();
    st.ok = true;
    st.name = ss.getName();
    st.url = ss.getUrl();
  } catch (e) {
    st.message = (e && e.message ? e.message : String(e));
  }
  return st;
}

function escapeHtmlText_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// 公開されているウェブアプリのURL。未デプロイなら空。
function getWebAppUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

// メニュー：ウェブアプリのURLを表示する
function showWebAppUrl() {
  var ui = SpreadsheetApp.getUi();
  try { getSS_(); } catch (e) {}          // ここでスプレッドシートIDを控えておく
  var url = getWebAppUrl_();
  if (!url) {
    ui.alert('ウェブアプリのURL',
      'まだデプロイされていません。\n\n'
      + '【手順】\n'
      + '1. 拡張機能 ＞ Apps Script を開く\n'
      + '2. 右上の「デプロイ」＞「新しいデプロイ」\n'
      + '3. 歯車アイコン ＞ 「ウェブアプリ」を選ぶ\n'
      + '4. 「次のユーザーとして実行」… 自分\n'
      + '   「アクセスできるユーザー」… 自分だけ（共有する場合は範囲を広げる）\n'
      + '5. 「デプロイ」を押し、表示されたURLを開く\n\n'
      + 'デプロイ後にもう一度このメニューを実行すると、URLをここに表示します。',
      ui.ButtonSet.OK);
    return;
  }
  ui.alert('ウェブアプリのURL',
    '下のURLをブラウザで開くと、すべての機能が普通のタブで使えます。\n'
    + 'ダイアログが真っ白になる環境でも、こちらなら表示されます。\n\n'
    + url + '\n\n'
    + 'ブックマークしておくと便利です。\n'
    + '※ 機能を追加・変更したあとは、Apps Scriptの「デプロイ」＞「デプロイを管理」から\n'
    + '　 バージョンを更新すると、ウェブアプリにも反映されます。',
    ui.ButtonSet.OK);
}

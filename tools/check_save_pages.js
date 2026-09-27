// 設定の画面（休会日・メールテンプレート・割り振り表の特記事項・ビジターホスト・Gemini API）で、
// 保存したあとも画面が消えないこと（ウェブアプリの「← メニューに戻る」の帯が残る）を確かめる。
//
//   node tools/check_save_pages.js
//
// 以前は保存すると画面を「保存しました／画面を閉じてください」に置き換えていたため、
// メニュー（ウェブアプリ）から開いたときに戻れなくなっていた。
//   ・保存しても、入力欄はそのまま残る
//   ・結果は画面の下（saveNote）に出る
//   ・ウェブアプリでは「← メニューに戻る」のリンクも出る。スプレッドシートの画面では出ない

const { loadPage } = require('./lib_minidom');

const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }

const PAGES = [
  // 休会日の一覧は li を足して描くので（簡易DOMでは描けない）、空で開いてから1日足したことにする
  { file: 'holiday.html', keep: 'holidayList', server: { getHolidays: () => [], saveHolidays: () => '保存しました' },
    before: "holidays.push('2026/05/06')", want: /保存しました（休会日 1日）/ },
  { file: 'allocation_note.html', keep: 'noteText', server: { getAllocationNote: () => '特記事項', saveAllocationNote: () => '特記事項を保存しました。' },
    want: /特記事項を保存しました/ },
  { file: 'api_settings.html', keep: 'apiKey', server: { getApiSettings: () => ({ apiKey: 'k', modelName: 'm' }), saveApiSettings: () => 'API設定を保存しました。' },
    want: /API設定を保存しました/ },
  { file: 'template.html', keep: 'tplVisitorSubj', server: {
      getTemplates: () => ({}), getMailWebAppSettings: () => ({}), getTemplateSettings: () => ({}),
      saveTemplates: () => 'テンプレートを保存しました。', saveMailWebAppSettings: () => '送信の設定を保存しました。' },
    want: /テンプレートを保存しました/ },
  { file: 'visitor_host.html', keep: null, server: {
      getVisitorHostSettings: () => ({ members: [], hosts: [], priorities: {} }), getMembersList: () => [], getVisitorHosts: () => [],
      getMemberPriorities: () => ({}), saveVisitorHosts: () => 'ok', saveMemberPriorities: () => 'ok', getVisitorHostsFromRoles: () => ({ ok: false }) },
    want: /保存しました/ },
];

for (const web of [false, true]) {
  for (const P of PAGES) {
    const page = loadPage(P.file, { server: P.server, fails: [] });
    if (web) page.sandbox.WEBAPP_URL = 'https://script.google.com/macros/s/TEST/exec';
    page.step('開く', () => page.window.onload && page.window.onload());
    if (P.before) page.step('準備', () => page.run(P.before));
    page.step('保存', () => page.run('save()'));
    const note = page.els.saveNote;
    const label = (web ? 'ウェブアプリ' : 'スプレッドシート') + '・' + P.file;
    ck(note && P.want.test(note.innerHTML), label + '：保存の結果が画面の下に出ない: ' + (note && note.innerHTML));
    ck(!P.keep || !!page.els[P.keep], label + '：保存したあと入力欄が消えた');
    ck(web === /← メニューに戻る/.test(note ? note.innerHTML : ''), label + '：メニューに戻るのリンク: ' + (note && note.innerHTML));
    ck(!page.fails.length, label + '：' + page.fails.join(' / '));
  }
}

console.log(`設定の画面の保存: 検査 ${checks} 件`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: 休会日・メールテンプレート・割り振り表の特記事項・Gemini API・ビジターホストは、保存しても画面が残り、ウェブアプリでは「← メニューに戻る」が出る');

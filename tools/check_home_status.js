// メニューバーでの保存確認と画面内の承認エラーを分け、未確認を準備不足と扱わない。
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeEnv } = require('./lib_sheet_fake');
const { launch, openGasPage } = require('./lib_gas_page');
const fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const env = makeEnv();
const srv = Object.assign({}, env.globals);
vm.createContext(srv);
const root = path.join(__dirname, '..');
for (const file of fs.readdirSync(root).filter(f => f.endsWith('.js')).sort()) {
  vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), srv, { filename: file });
}
env.reset([], {});
env.userEmail = 'member@example.com';
srv.ScriptApp.AuthMode = { FULL: 'FULL' };
let authState = 'NOT_REQUIRED', scopeRequests = 0;
srv.ScriptApp.getAuthorizationInfo = () => ({ getAuthorizationStatus: () => authState });
srv.ScriptApp.requireAllScopes = mode => { ck(mode === 'FULL', '権限確認が全スコープを要求しない'); scopeRequests++; };
const alerts = [];
srv.SpreadsheetApp.getUi = () => ({ alert: (...args) => alerts.push(args), ButtonSet: { OK: 'OK' } });

// 実際の保存処理で「保存成功、リンク共有だけ失敗」を再現する。
env.sharingBlocked = true;
const saved = srv.saveOutputFile_(srv.Utilities.newBlob('ok', 'text/plain', 'example.txt'), 'example.txt');
ck(saved.id && saved.url && saved.shared === false, 'リンク共有が禁止されたときに保存まで失敗した');
ck(srv.DriveApp.getFileById(saved.id).getName() === 'example.txt', '共有失敗後にファイルを取得できない');
srv.authorizeDriveAccess();
const warning = alerts.pop()[1];
ck(/ファイルの作成を確認/.test(warning) && !/準備できています/.test(warning), '保存確認を全機能の準備完了と表示した');
ck(/全員が閲覧.*設定.*失敗/.test(warning) && /閲覧権限のない相手/.test(warning), '共有できない場合の影響が分からない');
ck(/別途確認/.test(warning) && /member@example.com/.test(warning), '未確認の操作や実行者が分からない');
ck(scopeRequests === 1, '直接メニューの確認で権限を要求しない');
env.sharingBlocked = false;
srv.authorizeDriveAccess();
ck(/全員が閲覧.*設定.*成功/.test(alerts.pop()[1]), '共有できる場合も失敗と表示した');
const bigKinds = vm.runInContext('BIG_TEMPLATE_KINDS_', srv);
const bigKind = Object.keys(bigKinds)[0];
env.props[bigKinds[bigKind].prop] = 'INACCESSIBLE_TEMPLATE';
const inaccessible = srv.getBigTemplateStatus().templates.find(t => t.kind === bigKind);
ck(!inaccessible.registered && !!inaccessible.unreadable, '登録済みの大きなテンプレートが開けない理由を捨てた');

// 設定の点検だけを再現する。APIキーはサーバー内の設定にのみ存在する。
srv.chapterInfo_ = () => ({ saved: true });
srv.chapterFresh_ = () => false;
srv.chapterLabel_ = () => '見本チャプター';
srv.chapterSystemTitle_ = () => '見本チャプター 名簿システム';
srv.roleTermOf_ = () => 24;
srv.getMeetingCandidates = () => [{ dateValue: '2026/10/21', display: '10月21日' }];
srv.findRoutineColumn_ = () => ({ name: 'ルーティンチェックシート' });
srv.getAssetSettings = () => ({ folderId: 'example', folderName: '素材', reachable: true });
srv.getMembersList = () => [{ name: '見本' }];
srv.getMemberMaster = () => ({ members: [{ name: '見本' }] });
srv.getPhotoOverview = () => ({ indexed: 1, unmatched: [] });
srv.getTemplateStatus = () => ({ templates: [{ registered: true }] });
srv.getBigTemplateStatus = () => ({ templates: [{ registered: true }] });
srv.getApiSettings = () => ({ apiKey: 'PRIVATE_API_KEY_FOR_TEST', modelName: 'gemini-example' });
srv.getPdfLinks = () => ({ visitorList: 'https://drive.google.com/example', date: '2026/10/21' });
const healthy = srv.getHomeStatus();
ck(healthy.checks.length === 9 && healthy.checks.every(c => c.ready), '正常な設定の点検が欠ける');
ck(healthy.execution.version === srv.getSystemVersion() && healthy.execution.effectiveUser === env.userEmail,
  '画面から実行した版・アカウントを取得できない');
ck(healthy.execution.authorization === 'NOT_REQUIRED', 'Googleの承認状態を取得できない');
ck(!JSON.stringify(healthy).includes('PRIVATE_API_KEY_FOR_TEST'), '点検結果にAPIキーが含まれる');
const visitorStatus = srv.getTemplateStatus, bigStatus = srv.getBigTemplateStatus;
srv.getTemplateStatus = () => ({ templates: [{ registered: false, label: 'ビジター', unreadable: 'ファイルへのアクセス権限がありません' }] });
srv.getBigTemplateStatus = () => ({ templates: [inaccessible] });
const denied = srv.getHomeStatus();
ck(denied.checks.find(c => c.key === 'tpl').error.includes('アクセス権限'), '開けないビジターテンプレートを未登録と扱った');
ck(!!denied.checks.find(c => c.key === 'bigtpl').error, '開けない大きなテンプレートを未登録と扱った');
srv.getTemplateStatus = visitorStatus;
srv.getBigTemplateStatus = bigStatus;
srv.getMembersList = () => { throw new Error('名簿を開く権限がありません'); };
srv.getPdfLinks = () => { throw new Error('PDF一覧の取得に失敗しました'); };
const partial = srv.getHomeStatus();
ck(partial.checks.find(c => c.key === 'memberlist').error === '名簿を開く権限がありません', '点検のエラー理由を失う');
ck(partial.checks.find(c => c.key === 'chapter').ready, '一つの点検エラーが他の結果を消した');
ck(partial.latestError === 'PDF一覧の取得に失敗しました', '一覧取得の失敗をファイル未作成と扱った');
srv.diagnoseHomeStatus();
const diagnosis = alerts.pop()[1];
ck(/確認不可: メンバーリスト/.test(diagnosis) && diagnosis.includes(partial.latestError), '直接メニューで失敗した処理や理由が分からない');
ck(diagnosis.includes(srv.getSystemVersion()) && diagnosis.includes(env.userEmail), '直接メニューの診断に版やアカウントが無い');
ck(!diagnosis.includes('PRIVATE_API_KEY_FOR_TEST'), '直接メニューの診断にAPIキーが含まれる');
authState = 'REQUIRED';
ck(srv.homeExecutionInfo_().authorization === 'REQUIRED', '承認が必要な状態を隠した');
const session = srv.Session;
srv.Session = { getActiveUser: () => ({ getEmail: () => 'reader@example.com' }), getEffectiveUser: () => ({ getEmail: () => 'runner@example.com' }) };
const principals = srv.homeExecutionInfo_();
ck(principals.activeUser === 'reader@example.com' && principals.effectiveUser === 'runner@example.com', '利用者と実行者が違う場合に照合できない');
srv.Session = { getActiveUser: () => { throw new Error('取得不可'); }, getEffectiveUser: () => ({ getEmail: () => '' }) };
ck(srv.homeExecutionInfo_().activeUser === '' && srv.homeExecutionInfo_().effectiveUser === '', 'メールアドレスを取得できないだけで点検が止まる');
srv.Session = session;

async function main() {
  const browser = await launch();
  try {
    let state = 'unauthorized';
    const { page } = await openGasPage(browser, 'menu_home.html', {
      getHomeStatus: () => {
        if (state === 'unauthorized') throw new Error('その操作を実行するには承認が必要です。<img src=x onerror="window.injected=true">');
        return state === 'partial' ? partial : healthy;
      },
      openCsvDialog: () => { throw new Error('この画面を開くには承認が必要です。'); }
    }, { fails });
    await page.waitForFunction(() => document.getElementById('checks').textContent.includes('未確認'));
    ck(/承認が必要/.test(await page.locator('#checks').textContent()), '画面の承認エラーの詳細を隠す');
    ck(/未確認/.test(await page.locator('#latest').textContent()) && !/確認中/.test(await page.locator('#latest').textContent()),
      '点検失敗後もPDF一覧が確認中のままになる');
    ck(await page.locator('#checks img').count() === 0 && !(await page.evaluate(() => window.injected)), 'エラー文をHTMLとして実行した');
    ck(await page.locator('#retry').isEnabled(), '承認エラーのあとに再確認できない');
    ck((await page.locator('.note').textContent()).includes('準備状況と実行アカウントを確認'), '画面外の診断を案内していない');
    state = 'healthy';
    await page.locator('#retry').click();
    await page.waitForFunction(() => document.querySelectorAll('#checks .chk').length === 9);
    ck(!(await page.locator('#checks').textContent()).includes('未確認'), '再確認に成功しても前のエラーが残る');
    ck(await page.locator('#latest a').count() === 1, '再確認後にPDF一覧が戻らない');
    ck((await page.locator('#execution').textContent()).includes(env.userEmail), '画面の実行アカウントを照合できない');
    const before = await page.locator('#checks').innerHTML();
    await page.getByRole('button', { name: '1. CSVから名簿・PDF作成' }).click();
    await page.waitForFunction(() => document.getElementById('msg').textContent.includes('開けませんでした'));
    ck(await page.locator('#checks').innerHTML() === before, '画面を開く失敗を準備状況のエラーに書き換えた');
    ck(/承認が必要/.test(await page.locator('#msg').textContent()), '画面を開く失敗の理由が無い');
    state = 'partial';
    await page.locator('#retry').click();
    await page.waitForFunction(() => document.getElementById('checks').textContent.includes('名簿を開く権限'));
    ck(/未確認/.test(await page.locator('#checks .chk').nth(3).textContent()), '失敗した項目を準備不足と扱った');
    ck(await page.locator('#checks .chk').nth(3).locator('a').count() === 0, '確認できない項目から新規登録を勧めた');
    ck((await page.locator('#latest').textContent()).includes(partial.latestError), 'PDF一覧を取得できなかった理由が無い');
  } finally { await browser.close(); }
}
main().then(() => {
  if (fails.length) { console.log('NG ' + fails.length + '件 / ' + checks + '件'); console.log(fails.join('\n')); process.exit(1); }
  console.log('準備状況・実行アカウント・承認エラー・共有制限: 検査 ' + checks + ' 件 OK');
}).catch(e => { console.error(e); process.exit(1); });

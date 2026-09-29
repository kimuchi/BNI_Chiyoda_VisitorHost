// 「⚙️ 設定」＞「メンバー名簿」の画面（member_master.html）を、本物のブラウザ（Chromium）で動かして確かめる。
// サーバー側は本番の *.js を全部、見せかけのスプレッドシート（lib_sheet_fake.js）の上で動かす。
//
//   node tools/check_member_master_dialog.js
//
//   ・最初の読み込みに失敗したまま「＋ 行を追加」「名簿を保存」を押しても、名簿を上書きしない
//     （以前は、名簿全体がその1行だけになり、控えも無かった）
//   ・サーバーも、0名の名簿は保存しない
//   ・読み込めたときは、これまでどおり追加・保存できる
// 名前はすべて架空。

process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeEnv } = require('./lib_sheet_fake');
const { launch, openGasPage } = require('./lib_gas_page');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }

const env = makeEnv({ now: new Date(2026, 8, 29, 10, 0, 0) });
const srv = Object.assign({}, env.globals);
vm.createContext(srv);
for (const f of fs.readdirSync(ROOT).filter((x) => /\.js$/.test(x)).sort()) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), srv, { filename: f });
const HEAD = vm.runInContext('MEMBER_HEADERS_', srv);
const ROWS = [['1', '見本 一郎', '一言その1'], ['2', '試験 花子', '一言その2'], ['3', '架空 三郎', '一言その3']];
env.reset([['メンバー名簿', false, [HEAD].concat(ROWS.map(([no, n, c]) => HEAD.map((h) => (h === 'No' ? no : h === '氏名' ? n : h === '一言コメント' ? c : ''))))]], {});
const names = () => env.values('メンバー名簿').slice(1).map((r) => r[2]).join(',');

(async () => {
  const browser = await launch();
  try {
    // 1) 最初の読み込みに失敗する（通信の切れ・一時的なサーバーのエラー）
    const real = srv.getMemberMaster;
    let first = true;
    srv.getMemberMaster = function () { if (first) { first = false; throw new Error('サーバー エラーが発生しました。しばらくしてからもう一度お試しください。'); } return real.apply(this, arguments); };
    const calls = [];
    const { page } = await openGasPage(browser, 'member_master.html', srv, { fails, calls });
    await page.waitForFunction(() => /読み込めませんでした|エラー/.test(document.getElementById('msg').textContent), null, { timeout: 10000 });
    await page.click('button:has-text("＋ 行を追加")');
    ck(await page.locator('#body tr').count() === 0, '1) 読み込めていないのに行を足せた');
    await page.click('#btnSave');
    await page.waitForTimeout(300);
    ck(!calls.some((c) => c.name === 'saveMemberMaster'), '1) 読み込めていないのに、名簿の保存をサーバーに送った');
    ck(names() === '見本 一郎,試験 花子,架空 三郎', '1) 名簿が書き換わった: ' + names());
    ck(/読み込めていないため/.test(await page.textContent('#msg')), '1) 保存できない理由を知らせない: ' + await page.textContent('#msg'));
    await page.close();

    // 2) サーバーも 0名の名簿は保存しない
    const r0 = srv.saveMemberMaster([], null);
    ck(r0.ok === false && names() === '見本 一郎,試験 花子,架空 三郎', '2) 0名の名簿を保存した');

    // 3) 読み込めたときは、これまでどおり追加・保存できる（ほかの方の一言などは残る）
    const p2 = await openGasPage(browser, 'member_master.html', srv, { fails, calls });
    await p2.page.waitForFunction(() => document.querySelectorAll('#body tr').length === 3, null, { timeout: 10000 });
    await p2.page.click('button:has-text("＋ 行を追加")');
    await p2.page.locator('#body tr').nth(3).locator('input[data-k="name"]').fill('新入 四郎');
    await p2.page.locator('#body tr').nth(3).locator('input[data-k="name"]').dispatchEvent('change');
    await p2.page.click('#btnSave');
    await p2.page.waitForFunction(() => /保存しました/.test(document.getElementById('msg').textContent), null, { timeout: 10000 });
    ck(names() === '見本 一郎,試験 花子,架空 三郎,新入 四郎', '3) 読み込めたのに、追加・保存できない: ' + names());
    const comments = env.values('メンバー名簿').slice(1).map((r) => r[HEAD.indexOf('一言コメント')]);
    ck(comments.slice(0, 3).join(',') === '一言その1,一言その2,一言その3', '3) 保存したら、ほかの方の一言が消えた: ' + comments.join(','));
    await p2.page.close();
  } finally {
    await browser.close();
  }

  if (fails.length) {
    console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
    fails.forEach((f) => console.log('  - ' + f));
    process.exit(1);
  }
  console.log('メンバー名簿の画面（ブラウザ）: 検査 ' + checks + ' 件 OK: 読み込めていないときは追加・保存させない・0名は保存しない・読み込めたら保存できる');
})().catch((e) => {
  console.log('NG 検査が止まった: ' + (e && e.stack ? e.stack.split('\n').slice(0, 2).join(' / ') : e));
  fails.forEach((f) => console.log('  - ' + f));   // 止まる前に見つかったもの
  process.exit(1);
});

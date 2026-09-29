// 「1. CSVから名簿・PDF作成」の画面（dialog.html）を、本物のブラウザ（Chromium）で動かして確かめる。
// サーバー側は本番と同じ コード.js などを、見せかけのスプレッドシート・ドライブ（lib_sheet_fake.js）の上で動かす。
//
//   node tools/check_csv_dialog.js
//
// 確かめること（「きょう」は 2026/10/1。9/30 の定例会は終わっていて、候補は 10/7 から）
//   ・再編集：9/30 の名簿を読み込んで「作成する」→ 9/30 に書き戻す（候補に無くても。10/7 の名簿を上書きしない）。
//     書き込み先が編集画面に出る。前の不具合で隠れて真っ白になった 9/30 のシートとPDFも、これで直る
//   ・CSV：Shift_JIS のCSVを読み込み → 上で選んだ定例会（10/7）に作る
//   ・PDFのみ再作成：開催日が過ぎた回でも使える
//   ・PDFを作れなかったとき：画面に知らせて、はじめの画面に戻る
// 名前・メールアドレスはすべて架空。

process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { execFileSync } = require('child_process');
const { makeEnv } = require('./lib_sheet_fake');
const pw = (() => { try { return require('playwright'); } catch (e) { return require('/opt/node22/lib/node_modules/playwright'); } })();

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }

const env = makeEnv({ now: new Date(2026, 9, 1, 10, 0, 0) });
const srv = Object.assign({}, env.globals);
vm.createContext(srv);
for (const f of ['コード.js', 'chapter_srv.js', 'webapp_srv.js']) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), srv, { filename: f });

const MEMBER_HEAD = ['No', '業種区分', '氏名', 'ふりがな', 'カテゴリー', '会社名', '役職', 'メモ', '写真ファイル名', '一言コメント',
  '紹介してほしい人', '協業したい人', '入会日', '更新日', '更新期限日', '会社での役職'];
const MEMBERS = [['1', '見本 一郎'], ['2', '試験 花子'], ['3', '架空 三郎'], ['4', '仮名 四郎'], ['5', '例示 五月']];
const HEAD = ['No.', '参加者氏名', 'ふりがな', 'カテゴリー', '会社名', '招待者', '備考', '種別', 'メール'];
function start() {
  env.reset([
    ['メンバー名簿', false, [MEMBER_HEAD].concat(MEMBERS.map(([no, n]) => MEMBER_HEAD.map((h, i) => (i === 0 ? no : i === 2 ? n : ''))))],
    ['休会日', true, [['2026/05/06'], ['2026/08/12'], ['2026/12/30']]],
    // 前の不具合のあと：9/30 のシートは隠れていて、ドライブのPDF（F0）は真っ白
    ['20260930参加者', true, [HEAD,
      ['V01', '見本 太郎', 'みほん たろう', '税理士', '見本商事', '試験 花子', '', 'Visitor', 'taro@example.com'],
      ['G01', '仮設 月子', 'かせつ つきこ', '', '', '見本 一郎', '', 'Guest', 'tsuki@example.com']]],
    ['20260930参加者_印刷用', true, [['BNI 見本チャプターの定例会へようこそ'], [''], ['2026/9/30(水) 第535回'], [''], HEAD.slice(0, 7),
      ['V01', '見本 太郎', 'みほん たろう', '税理士', '見本商事', '試験 花子', ''], ['G01', '仮設 月子', 'かせつ つきこ', '', '', '見本 一郎', '']]],
  ], { 'VISITOR_PDF_ID_20260930参加者': 'F0' });
  env.addFile('F0', new env.FakeBlob('%PDF-1.4\n', 'application/pdf', '2026/9/30(水) 第535回 ビジター様リスト.pdf'));
}

const calls = [];
async function openDialog(browser) {
  const page = await browser.newPage();
  const dialogs = [];
  page.on('dialog', async (d) => { dialogs.push(d.type() + ':' + d.message()); await d.accept(); });
  page.on('pageerror', (e) => fails.push('画面のエラー: ' + e.message));
  await page.exposeFunction('__gas', (name, argsJson) => {
    const args = JSON.parse(argsJson);
    calls.push({ name, args });
    try {
      if (typeof srv[name] !== 'function') return JSON.stringify({ ok: false, message: 'サーバーに無い関数: ' + name });
      const r = srv[name](...args);
      return JSON.stringify({ ok: true, r: r === undefined ? null : r });
    } catch (e) { return JSON.stringify({ ok: false, message: String(e && e.message ? e.message : e) }); }
  });
  await page.addInitScript(() => {
    function runner() {
      let ok = null, ng = null;
      const r = new Proxy({}, { get(_, name) {
        if (name === 'withSuccessHandler') return (f) => { ok = f; return r; };
        if (name === 'withFailureHandler') return (f) => { ng = f; return r; };
        if (name === 'withUserObject') return () => r;
        return (...args) => {
          window.__gas(String(name), JSON.stringify(args)).then((s) => {
            const x = JSON.parse(s);
            if (x.ok) { if (ok) ok(x.r); } else if (ng) ng(new Error(x.message));
          });
        };
      } });
      return r;
    }
    window.google = { script: { get run() { return runner(); } } };
  });
  await page.route('https://dialog.test/', (r) => r.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: fs.readFileSync(path.join(ROOT, 'dialog.html'), 'utf8') }));
  await page.goto('https://dialog.test/', { waitUntil: 'load' });
  await page.waitForFunction(() => document.querySelectorAll('#meetingSelect option').length === 4 && document.querySelectorAll('#sheetSelect option').length > 0, null, { timeout: 10000 });
  return { page, dialogs };
}
const visible = (page, sel) => page.evaluate((s) => !document.querySelector(s).classList.contains('hidden'), sel);
const lastCall = (name) => calls.filter((c) => c.name === name).pop();

(async () => {
  const browser = await pw.chromium.launch();
  try {
    // ---- 1) 再編集：過ぎた回（9/30）を読み込んで作る → 9/30 に書き戻す ----
    start();
    {
      const { page } = await openDialog(browser);
      ck(await page.inputValue('#meetingSelect') === '2026/10/07', '1) 定例会の既定が 10/7 でない: ' + await page.inputValue('#meetingSelect'));
      await page.selectOption('#sheetSelect', '20260930参加者');
      await page.click('button:has-text("このシートを読み込んで再編集")');
      await page.waitForFunction(() => !document.getElementById('step2').classList.contains('hidden'), null, { timeout: 10000 });
      const note = await page.evaluate(() => (document.getElementById('targetNote') || {}).textContent || '');
      ck(/2026\/9\/30\(水\) 第535回/.test(note) && /書き戻します/.test(note), '1) 書き込み先が 9/30 と出ない: ' + note);
      ck(await page.inputValue('#meetingSelect') === '2026/09/30', '1) 読み込んだ回（9/30）が選ばれていない: ' + await page.inputValue('#meetingSelect'));
      ck(await page.locator('#dataBody tr').count() === 2, '1) 読み込んだ行の数が違う');
      await page.locator('#dataBody tr').nth(0).locator('input[type=text]').nth(3).fill('直した会社');
      await page.click('button:has-text("この内容でシートとPDFを作成する")');
      await page.waitForFunction(() => /処理が完了しました/.test(document.getElementById('loading').textContent), null, { timeout: 10000 });
      const c = lastCall('createFinalSheet');
      ck(c && c.args[0] === '2026/09/30' && c.args[1] === '2026/9/30(水) 第535回', '1) 9/30 ではなく別の日に作った: ' + JSON.stringify(c && c.args.slice(0, 2)));
      ck(!env.sheet('20261007参加者'), '1) 10/7 の名簿を作ってしまった（上書き）');
      ck(!env.sheet('20260930参加者').isSheetHidden() && !env.sheet('20260930参加者_印刷用').isSheetHidden(), '1) 9/30 のシートが隠れたまま');
      ck(env.fileText('F0').includes('直した会社') && env.fileText('F0').includes('2026/9/30(水) 第535回'), '1) 9/30 のPDF（同じURL）が直らない: ' + JSON.stringify(env.fileText('F0').slice(0, 60)));
      ck(env.drive.created.length === 0, '1) 新しいPDFを作った（URLが変わる）');
      await page.close();
    }

    // ---- 2) CSV（Shift_JIS）から、上で選んだ定例会（10/7）に作る ----
    {
      const { page } = await openDialog(browser);
      const csv = ['Name,Furigana,Company Name,Business Category,Inviter,Email,Type',
        '翌週 来子,よくしゅう らいこ,翌週社,デザイン,例示 五月,raiko@example.com,Visitor',
        '代理 太一,だいり たいち,代理商会,保険,仮名 四郎,taichi@example.com,Substitute'].join('\r\n') + '\r\n';
      const sjis = execFileSync('python3', ['-c', 'import sys; sys.stdout.buffer.write(sys.stdin.buffer.read().decode("utf-8").encode("cp932"))'], { input: Buffer.from(csv, 'utf8') });
      await page.setInputFiles('#csvFile', { name: 'visitors.csv', mimeType: 'text/csv', buffer: sjis });
      await page.click('button:has-text("全データを解析して編集画面へ")');
      await page.waitForFunction(() => !document.getElementById('step2').classList.contains('hidden'), null, { timeout: 10000 });
      const note = await page.evaluate(() => (document.getElementById('targetNote') || {}).textContent || '');
      ck(/2026\/10\/7\(水\) 第536回/.test(note) && !/書き戻/.test(note), '2) CSVの書き込み先が 10/7 と出ない: ' + note);
      const names = await page.locator('#dataBody tr input[type=text]').evaluateAll((els) => els.map((e) => e.value));
      ck(names.includes('翌週 来子') && names.includes('例示 五月'), '2) Shift_JIS のCSVが読めていない（文字化け）: ' + names.slice(0, 6).join(','));
      await page.click('button:has-text("この内容でシートとPDFを作成する")');
      await page.waitForFunction(() => /処理が完了しました/.test(document.getElementById('loading').textContent), null, { timeout: 10000 });
      const c = lastCall('createFinalSheet');
      ck(c && c.args[0] === '2026/10/07', '2) 10/7 に作っていない: ' + JSON.stringify(c && c.args.slice(0, 2)));
      const data = env.values('20261007参加者');
      ck(data && data.some((r) => r[0] === 'V01' && r[1] === '翌週 来子') && data.some((r) => r[0] === '代理4' && r[1] === '代理 太一'), '2) 10/7 の名簿の中身が違う: ' + JSON.stringify(data && data.slice(0, 3)));
      ck(env.sheet('20260930参加者').isSheetHidden() && !env.sheet('20261007参加者').isSheetHidden(), '2) 前の回のアーカイブ・今回の表示が違う');
      const id = env.props['VISITOR_PDF_ID_20261007参加者'];
      ck(id && env.fileText(id).includes('翌週 来子'), '2) 10/7 のPDFに中身が無い');
      await page.close();
    }

    // ---- 3) 過ぎた回（9/30）の「PDFのみ再作成」 ----
    {
      const { page, dialogs } = await openDialog(browser);
      env.sheet('20260930参加者_印刷用').getRange(6, 5).setValue('手で直した会社');
      await page.selectOption('#sheetSelect', '20260930参加者');
      await page.click('button:has-text("PDFのみ再作成")');
      await page.waitForFunction(() => /PDFを再作成しました/.test(document.getElementById('loading').textContent), null, { timeout: 10000 });
      ck(dialogs.some((d) => /^confirm:/.test(d)), '3) PDFのみ再作成の前に確かめていない');
      ck(env.fileText('F0').includes('手で直した会社'), '3) 過ぎた回のPDFのみ再作成で、PDFが直らない');
      ck(!env.sheet('20260930参加者').isSheetHidden(), '3) PDFのみ再作成で、その回のシートが表示に戻らない');
      await page.close();
    }

    // ---- 4) PDFを作れなかったとき：知らせて、はじめの画面に戻る。前のPDFはそのまま ----
    {
      const { page, dialogs } = await openDialog(browser);
      const before = env.fileText('F0');
      env.fetchPlan = [0, 1, 2].map(() => () => ({ code: 500, body: Buffer.from('<html>error</html>'), type: 'text/html' }));
      await page.selectOption('#sheetSelect', '20260930参加者');
      await page.click('button:has-text("このシートを読み込んで再編集")');
      await page.waitForFunction(() => !document.getElementById('step2').classList.contains('hidden'), null, { timeout: 10000 });
      await page.click('button:has-text("この内容でシートとPDFを作成する")');
      await page.waitForFunction(() => !document.getElementById('step1').classList.contains('hidden'), null, { timeout: 10000 });
      ck(dialogs.some((d) => /^alert:.*PDFを作れませんでした/.test(d)), '4) PDFを作れなかったことを知らせない: ' + dialogs.join(' / '));
      ck(env.fileText('F0') === before, '4) 作れなかったのに、前のPDFを書き換えた');
      ck(await visible(page, '#step1'), '4) はじめの画面に戻らない');
      env.fetchPlan = [];
      await page.close();
    }
  } finally {
    await browser.close();
  }

  if (fails.length) {
    console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
    fails.forEach((f) => console.log('  - ' + f));
    process.exit(1);
  }
  console.log('CSVから名簿・PDF作成の画面（ブラウザ）: 検査 ' + checks + ' 件 OK: 過ぎた回の再編集はその回に書き戻す・Shift_JIS のCSV・PDFのみ再作成・PDFを作れなかったとき');
})().catch((e) => { console.log('NG 検査が止まった: ' + (e && e.stack ? e.stack : e)); process.exit(1); });

// 「3. ルーム・オリエン割り振り表」の画面（allocation.html）を、本物のブラウザ（Chromium）で動かして確かめる。
// サーバー側は本番と同じ コード.js などを、見せかけのスプレッドシート・ドライブ（lib_sheet_fake.js）の上で動かす。
//
//   node tools/check_allocation_dialog.js
//
//   ・読み込んだ回に書き込む（読み込んだあとで上の選択を別の回に変えても、その回の割り振り表とPDFを上書きしない）。書き込み先が画面に出る
//   ・Spreadingから来た文字（氏名・会社・メモ・URL）は、HTMLとして動かない（タグ・javascript: のリンク）
//   ・確度とメモは、参加者シートの列名（入会見込み・備考・メモ（メンバー向け））でも出る
//   ・入会確度（☆1〜3）：画面で選ぶ → 割り振り表の「入会確度」の列・PDFに ☆ で入る → 読み込み直すと戻る。
//     「AIに提案させる」のあとも残る。入力欄から☆に移ったとき、☆の欄を描き直さない（1回目のクリックが消えない）
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
for (const f of ['コード.js', 'chapter_srv.js', 'webapp_srv.js']) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), srv, { filename: f });

const MHEAD = ['No', '業種区分', '氏名'];
const HEAD = ['No.', '参加者氏名', 'ふりがな', 'カテゴリー', '会社名', '招待者', '備考', '種別', 'メール', '入会見込み', 'メモ（メンバー向け）', 'ウェブサイト'];
env.reset([
  ['メンバー名簿', false, [MHEAD, ['1', '', '見本 一郎'], ['2', '', '試験 花子'], ['3', '', '架空 三郎']]],
  ['休会日', true, [['2026/12/30']]],
  ['20260930参加者', false, [HEAD,
    ['V01', '<img src=x onerror="window.__xss=1">見本 太郎', 'みほん たろう', '税理士<b>', '見本商事', '見本 一郎', '備考のメモ', 'Visitor', '', '5', '内部メモ', 'javascript:window.__xss=2']]],
  ['20260930割り振り表', false, [['割り振り表（9/30）'], ['すでに共有した中身']]],
  ['20261007参加者', false, [HEAD,
    ['V01', '翌週 来子', 'よくしゅう らいこ', 'デザイン', '翌週社', '試験 花子', '', 'Visitor', '', '3', '', 'https://example.com/"onmouseover="window.__xss=3'],
    // 番号に「'」が入った方・同じ番号（代理??）の方が2人（招待者が名簿に見つからない代理など）
    ["V'02", '引用 符子', 'いんよう ふこ', '見本', '', '', '', 'Visitor', '', '', '', ''],
    ['代理??', '代理 一子', 'だいり いちこ', '', '', '', '', 'Substitute', '', '', '', ''],
    ['代理??', '代理 二子', 'だいり にこ', '', '', '', '', 'Substitute', '', '', '', '']]],
], { 'ALLOC_PDF_ID_20260930割り振り': 'A30' });
env.addFile('A30', new env.FakeBlob('%PDF-1.4\n9/30 の割り振り表（共有済み）', 'application/pdf', '9/30 割り振り表.pdf'));

(async () => {
  const browser = await launch();
  try {
    const calls = [];
    const { page, dialogs } = await openGasPage(browser, 'allocation.html', srv, { fails, calls });
    await page.waitForFunction(() => document.querySelectorAll('#meetingSelect option').length === 4, null, { timeout: 10000 });
    // 10/7 を読み込む
    await page.selectOption('#meetingSelect', '2026/10/07');
    await page.click('button:has-text("データを読み込む")');
    await page.waitForFunction(() => document.getElementById('workspace').style.display === 'flex', null, { timeout: 10000 });
    const note = await page.evaluate(() => (document.getElementById('targetNote') || {}).textContent || '');
    ck(/書き込み先：2026\/10\/7\(水\) 第\d+回/.test(note), '1) 書き込み先が画面に出ない: ' + note);
    // 入会確度：はじめは「－」（決めていない）。☆☆ を選ぶ
    ck(await page.evaluate(() => { const e = document.getElementById('star_V01'); return !!e && e.value === '' && /unset/.test(e.className)
         && Array.from(e.options).map((o) => o.textContent).join(',') === '－,☆,☆☆,☆☆☆'; }), '4) 入会確度の欄が無い・はじめから☆が入っている');
    // 「つなげたいメンバー」の入力欄から☆の欄に移っても、☆の欄を描き直さない（描き直すと1回目のクリックが消える）
    await page.evaluate(() => { document.getElementById('star_V01').__mark = 1; });
    await page.fill('textarea.connect-input', 'IT関係の方');
    await page.focus('#star_V01');
    ck(await page.evaluate(() => document.getElementById('star_V01').__mark === 1), '4) 入力欄から移ったら、☆の欄を描き直した（クリックが消える）');
    await page.selectOption('#star_V01', '2');
    ck(await page.evaluate(() => !/unset/.test(document.getElementById('star_V01').className)), '4) ☆を選んでも色が変わらない');
    // 番号に「'」が入った方でも選べる。同じ番号の方が2人いるときは、2人とも同じ☆になる（どちらの欄も、そうと分かる）
    await page.selectOption('select[data-no="V\'02"]', '1');
    await page.locator('select[data-no="代理??"]').nth(1).selectOption('3');
    ck(await page.evaluate(() => Array.from(document.querySelectorAll('select[data-no="代理??"]')).map((e) => e.value + ':' + (/unset/.test(e.className) ? 'grey' : 'color')).join(',')) === '3:color,3:color',
       '4) 同じ番号の方の☆の欄がそろわない（選んだ欄と違う欄の色が変わる）');
    // 読み込んだあとで、上の選択を 9/30 に変えてから作成する → 10/7 に書く（9/30 の割り振り表とPDFはそのまま）
    await page.selectOption('#meetingSelect', '2026/09/30');
    await page.click('button:has-text("この内容でシートとPDFを作成する")');
    await page.waitForFunction(() => /作成完了/.test(document.getElementById('loading').textContent), null, { timeout: 10000 });
    const save = calls.filter((c) => c.name === 'saveAllocationSheet').pop();
    ck(save && save.args[0] === '2026/10/07' && /^2026\/10\/7\(水\) 第\d+回$/.test(save.args[1]), '1) 読み込んだ回（10/7）ではなく、選び直した回に書いた: ' + JSON.stringify(save && save.args.slice(0, 2)));
    ck(env.values('20260930割り振り表')[1][0] === 'すでに共有した中身' && env.fileText('A30').includes('共有済み'), '1) 9/30 の割り振り表・共有したPDFを上書きした');
    ck(env.sheet('20261007割り振り表') && env.values('20261007割り振り表').some((r) => r[1] === '翌週 来子'), '1) 10/7 の割り振り表ができていない');
    // 入会確度は「入会確度」の列（カテゴリーのあと）とPDFに ☆ で入る
    ck(save && JSON.stringify(save.args[9]) === JSON.stringify({ V01: 2, "V'02": 1, '代理??': 3 }), '4) 入会確度がサーバーに渡らない: ' + JSON.stringify(save && save.args[9]));
    const a7 = env.values('20261007割り振り表') || [], h7 = a7.find((r) => r[0] === 'No.') || [], v7 = a7.find((r) => r[0] === 'V01') || [];
    ck(h7[3] === '入会確度' && v7[3] === '☆☆' && v7[5] === 'IT関係の方', '4) 割り振り表の入会確度の列: ' + JSON.stringify([h7, v7]));
    ck(a7.filter((r) => r[0] === "V'02").map((r) => r[3]).join(',') === '☆' && a7.filter((r) => r[0] === '代理??').map((r) => r[3]).join(',') === '☆☆☆,☆☆☆',
       "4) 番号に「'」が入った方・同じ番号の方の☆: " + JSON.stringify(a7.filter((r) => /V'02|代理/.test(r[0]))));
    ck(/V01\t翌週 来子\tデザイン\t☆☆\t/.test(env.fileText(env.props['ALLOC_PDF_ID_20261007割り振り']) || ''), '4) 割り振り表のPDFに ☆ が入らない');
    await page.close();

    // 読み込み直すと、☆が戻る。「AIに提案させる」のあとも残り、もう一度作ってもそのまま
    const p3 = await openGasPage(browser, 'allocation.html', srv, { fails, calls });
    await p3.page.waitForFunction(() => document.querySelectorAll('#meetingSelect option').length === 4, null, { timeout: 10000 });
    await p3.page.selectOption('#meetingSelect', '2026/10/07');
    await p3.page.click('button:has-text("データを読み込む")');
    await p3.page.waitForFunction(() => document.getElementById('workspace').style.display === 'flex', null, { timeout: 10000 });
    ck(await p3.page.evaluate(() => document.getElementById('star_V01').value) === '2', '4) 読み込み直したら、☆が戻らない');
    env.props.GEMINI_API_KEY = 'test-key';
    env.onFetch = () => ({ code: 200, type: 'application/json', body: Buffer.from(JSON.stringify({ candidates: [{ content: { parts: [{ text: '{"allocations":[{"visitorNo":"V01","room":["3"]}]}' }] } }] })) });
    await p3.page.click('button:has-text("AIに提案させる")');
    await p3.page.waitForFunction(() => document.getElementById('workspace').style.display === 'flex', null, { timeout: 10000 });
    ck(await p3.page.evaluate(() => document.getElementById('star_V01').value) === '2', '4) 「AIに提案させる」のあと、☆が消えた');
    const ai = calls.filter((c) => c.name === 'callGeminiAutoAllocation').pop();
    ck(ai && JSON.stringify(ai.args[0].stars) === JSON.stringify({ V01: 2, "V'02": 1, '代理??': 3 }), '4) AIに☆が渡らない・読み込み直した☆と違う: ' + JSON.stringify(ai && ai.args[0].stars));
    await p3.page.selectOption('#star_V01', '');
    await p3.page.click('button:has-text("この内容でシートとPDFを作成する")');
    await p3.page.waitForFunction(() => /作成完了/.test(document.getElementById('loading').textContent), null, { timeout: 10000 });
    ck(((env.values('20261007割り振り表') || []).find((r) => r[0] === 'V01') || [])[3] === '', '4) ☆を「－」に戻したのに、割り振り表に☆が残った');
    delete env.props.GEMINI_API_KEY;
    env.onFetch = null;
    await p3.page.close();

    // 9/30：Spreadingの文字はHTMLとして動かない。確度・メモは参加者シートの列名でも出る
    const p2 = await openGasPage(browser, 'allocation.html', srv, { fails, calls });
    await p2.page.waitForFunction(() => document.querySelectorAll('#meetingSelect option').length === 4, null, { timeout: 10000 });
    await p2.page.selectOption('#meetingSelect', '2026/09/30');
    await p2.page.click('button:has-text("データを読み込む")');
    await p2.page.waitForFunction(() => document.getElementById('workspace').style.display === 'flex', null, { timeout: 10000 });
    await p2.page.click('#btn_V01');
    await p2.page.waitForTimeout(300);
    const r = await p2.page.evaluate(() => ({
      xss: window.__xss || 0,
      imgs: document.querySelectorAll('#visitorTableBody img').length,
      bolds: document.querySelectorAll('#visitorTableBody b').length,
      jsLinks: Array.from(document.querySelectorAll('#visitorTableBody a')).filter((a) => /^javascript:/i.test(a.getAttribute('href') || '')).length,
      nameText: document.querySelector('#visitorTableBody tr td:nth-child(2)').textContent,
      info: document.querySelector('#visitorTableBody .info-col').textContent.replace(/\s+/g, ' '),
    }));
    ck(r.xss === 0 && r.imgs === 0 && r.bolds === 0 && r.jsLinks === 0, '2) Spreadingの文字がHTMLとして動いた: ' + JSON.stringify(r));
    ck(r.nameText.includes('<img src=x onerror="window.__xss=1">見本 太郎'), '2) 氏名がそのままの文字で出ない: ' + r.nameText);
    ck(/確度: 5/.test(r.info) && /内部メモ/.test(r.info) && /備考のメモ/.test(r.info), '3) 確度・メモが出ない: ' + r.info);
    // 10/7 の方のURL（「"」入り）も、リンクの属性を抜け出さない
    await p2.page.hover('#visitorTableBody');
    ck(await p2.page.evaluate(() => window.__xss || 0) === 0, '2) URLの「"」で属性を抜け出した');
    ck(dialogs.length === 0 && p2.dialogs.length === 0, '画面に思わぬ知らせが出た: ' + dialogs.concat(p2.dialogs).join(' / '));
    await p2.page.close();
  } finally {
    await browser.close();
  }

  if (fails.length) {
    console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
    fails.forEach((f) => console.log('  - ' + f));
    process.exit(1);
  }
  console.log('割り振り表の画面（ブラウザ）: 検査 ' + checks + ' 件 OK: 読み込んだ回に書き込む・Spreadingの文字を動かさない・確度とメモ・入会確度（☆）を選ぶ・シートとPDF・読み込み直す・AIのあと');
})().catch((e) => {
  console.log('NG 検査が止まった: ' + (e && e.stack ? e.stack.split('\n').slice(0, 2).join(' / ') : e));
  fails.forEach((f) => console.log('  - ' + f));   // 止まる前に見つかったもの
  process.exit(1);
});

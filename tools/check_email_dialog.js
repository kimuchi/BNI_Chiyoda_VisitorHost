// 「2. メールの確認・一括送信」の画面（email.html）を、本物のブラウザ（Chromium）で動かして確かめる。
// サーバー側は本番と同じ コード.js などを、見せかけのスプレッドシート・ドライブ・Gmail（lib_sheet_fake.js）の上で動かす。
//
//   node tools/check_email_dialog.js
//
// 送ったメールは取り消せないので、画面の見た目どおりに送られるかを本物の画面で確かめる。
//   ・既定の開催日は次回。本文のリンクはその開催日のPDF。印（キャンセル・送信済み）とチェックの状態
//   ・件名・本文の記号（" < > &）が、画面でも送ったメールでも崩れない
//   ・送ったあと開き直すと、送った方はチェックが外れ「送信済み」と出る。わざと送り直すときは確認で知らせる
//   ・次回のシートが無く、過ぎた回が選ばれたときは赤字で知らせる
//   ・代理送信の Web App が動いていない（ログインの画面などが返る）ときは、1人目で止めて理由を出し、
//     「自分のGmailから送る」で送れる。送れなかった方がいるのに「完了🎉」と出さない
//   ・設定の「代理送信を確かめる」・/dev のURLの知らせ
// 名前・メールアドレスはすべて架空。

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

const HEAD = ['No.', '参加者氏名', 'ふりがな', 'カテゴリー', '会社名', '招待者', '備考', '種別', 'メール', 'ステータス'];
function start(withNext) {
  const sheets = [
    ['メンバー名簿', false, [['No', '業種区分', '氏名'], ['1', '', '見本 一郎']]],
    ['20260923参加者', true, [HEAD, ['V01', '前回 来人', 'ぜんかい くると', '', '', '', '', 'Visitor', 'prev@example.com', '']]],
  ];
  if (withNext) sheets.push(['20260930参加者', false, [HEAD,
    ['V01', '見本 太郎', 'みほん たろう', '税理士', '見本商事', '見本 一郎', '', 'Visitor', 'taro@example.com', '参加予定'],
    ['V02', '取消 次子', 'とりけし つぎこ', '', '', '見本 一郎', '', 'Visitor', 'tori@example.com', 'キャンセル'],
    ['G01', '記号 <花> & "子"', 'きごう はなこ', '', '', '見本 一郎', '', 'Guest', 'ｈａｎａ＠ｅｘａｍｐｌｅ．ｃｏｍ', ''],
    ['V03', '無メール 氏', 'むめーる うじ', '', '', '', '', 'Visitor', '', '']]]);
  env.reset(sheets, {
    'VISITOR_PDF_ID_20260930参加者': 'P30', MEMBER_BOOK_URL: 'https://drive.example/MEMBERBOOK',
    MAIL_TPL_GUEST_SUBJ: '【"見本" <定例会>】{{date}} ご案内 & 資料', MAIL_TPL_GUEST_BODY: '{{name}} 様\n<ビジターリスト> {{visitorlist}}',
  });
  env.addFile('P30', new env.FakeBlob('%PDF-1.4\n中身', 'application/pdf', '9/30 ビジター様リスト.pdf'));
}
const cards = (page) => page.evaluate(() => {
  const out = [];
  for (let i = 0; document.getElementById('chk_' + i); i++) {
    const card = document.getElementById('chk_' + i).closest('.card');
    out.push({ checked: document.getElementById('chk_' + i).checked, head: card.querySelector('.card-header').textContent.replace(/\s+/g, ' ').trim(),
               subject: document.getElementById('subj_' + i).value, body: document.getElementById('body_' + i).value });
  }
  return out;
});
async function open(browser) {
  const r = await openGasPage(browser, 'email.html', srv, { fails });
  await r.page.waitForFunction(() => /読み込みました/.test(document.getElementById('meetNote').textContent) || /見つかりません|ありません/.test(document.getElementById('emailContainer').textContent), null, { timeout: 10000 });
  return r;
}
const LOGIN_PAGE = '<!DOCTYPE html><html><head><title>Google アカウントへのログイン</title></head>'
  + '<body><form action="https://accounts.google.com/ServiceLogin">ログイン</form></body></html>';
const RELAY = 'https://script.google.com/macros/s/RELAYID/exec';
const html = (code, body) => ({ code, body: Buffer.from(body), type: 'text/html' });
const json = (o) => ({ code: 200, body: Buffer.from(JSON.stringify(o)), type: 'application/json' });
const relayFetches = () => env.fetchLog.filter((f) => f.url === RELAY).length;

async function sendAll(page) {
  await page.click('#sendBtn');
  await page.waitForFunction(() => /すべての処理が完了しました/.test(document.body.textContent), null, { timeout: 10000 });
}

// ---- 0) サーバー：代理送信の Web App の返事が、送った結果（JSON）でないとき ----
try {
  start(true);
  env.props.MAIL_WEB_APP_URL = RELAY;
  const draft = { name: '見本 太郎', email: 'taro@example.com', subject: '件名', body: '本文', sheet: '20260930参加者' };
  env.onFetch = () => html(200, LOGIN_PAGE);
  let r = srv.sendSingleEmail(draft, '', '');
  ck(r.success === false && r.relayBroken === true && /送られていません/.test(r.error) && /ログイン/.test(r.error) && /自分のGmailから送る/.test(r.error)
     && !/SyntaxError/.test(r.error), '0) ログインの画面が返ったときの知らせ: ' + JSON.stringify(r));
  ck(env.mail.length === 0 && !env.props['MAIL_SENT_20260930参加者'], '0) 送れていないのに送信済みにした');
  env.onFetch = () => html(404, '<html><title>エラー</title>申し訳ございません。現在ファイルを開くことができません。</html>');
  r = srv.sendSingleEmail(draft, '', '');
  ck(r.relayBroken && /URLが見つかりません/.test(r.error), '0) 見つからないページのときの知らせ: ' + r.error);
  env.onFetch = () => json({ success: false, error: 'Error: アクセス権限がありません' });
  r = srv.sendSingleEmail(draft, '', '');
  ck(r.relayBroken && /合言葉/.test(r.error), '0) 合言葉が違うときの知らせ: ' + r.error);
  // 自分のGmailから送る：代理送信を通らない。送信済みの印も付く
  const n = relayFetches();
  r = srv.sendSingleEmail(draft, '', '', { direct: true });
  ck(r.success && r.direct && env.mail.length === 1 && env.mail[0].to === 'taro@example.com' && relayFetches() === n
     && JSON.parse(env.props['MAIL_SENT_20260930参加者'] || '{}')['taro@example.com'], '0) 自分のGmailから送れない: ' + JSON.stringify(r));
  // 設定の「代理送信を確かめる」
  env.onFetch = () => html(200, LOGIN_PAGE);
  let t = srv.testMailWebApp({ webAppUrl: RELAY, webAppToken: 'x' });
  ck(t.ok === false && /ログイン/.test(t.message), '0) 確かめる（ログインの画面）: ' + JSON.stringify(t));
  env.onFetch = (url, o) => (JSON.parse(o.payload).ping ? json({ success: true, ping: true, sender: 'chapter@example.com' }) : json({ success: false, error: '送った' }));
  t = srv.testMailWebApp({ webAppUrl: RELAY, webAppToken: 'x' });
  ck(t.ok === true && /chapter@example\.com/.test(t.message) && env.mail.length === 1, '0) 確かめる（動いている）: ' + JSON.stringify(t));
  t = srv.testMailWebApp({ webAppUrl: '', webAppToken: '' });
  ck(t.ok === true && /Gmailから送ります/.test(t.message), '0) 確かめる（代理送信なし）: ' + JSON.stringify(t));
  const sv = srv.saveMailWebAppSettings({ webAppUrl: 'https://script.google.com/macros/s/RELAYID/dev', webAppToken: 'x' });
  ck(/\/dev/.test(sv) && /⚠/.test(sv), '0) /dev のURLを保存しても知らせない: ' + sv);
  // doPost：確かめるときは送らずに返事だけ
  env.props.MAIL_WEB_APP_TOKEN = 'tok';
  const out = { text: '' };
  srv.ContentService = { MimeType: { JSON: 'json' }, createTextOutput: (x) => { out.text = x; return { setMimeType: () => out }; } };
  srv.doPost({ postData: { contents: JSON.stringify({ token: 'tok', ping: true }) } });
  ck(/"ping":true/.test(out.text) && env.mail.length === 1, '0) doPost の確かめで、メールを送った・返事が違う: ' + out.text);
} catch (e) {
  fails.push('0) 止まった: ' + (e && e.message ? e.message : e));
} finally {
  env.onFetch = null;
}

(async () => {
  const browser = await launch();
  try {
    // ---- 1) 次回（9/30）の案内 ----
    start(true);
    {
      const { page, dialogs } = await open(browser);
      ck(await page.inputValue('#meetSel') === '20260930参加者', '1) 既定の開催日が次回（9/30）でない: ' + await page.inputValue('#meetSel'));
      const note = await page.textContent('#meetNote');
      ck(/2026年9月30日/.test(note) && !/過ぎた回/.test(note), '1) 開催日の知らせが違う: ' + note);
      ck(/メールアドレスが無いため送れない方：無メール 氏/.test(note), '1) メールの無い方を知らせない: ' + note);
      const c = await cards(page);
      ck(c.length === 3, '1) カードの数が違う（メールの無い方は出さない）: ' + c.length);
      const taro = c.find((x) => /見本 太郎/.test(x.head)), tori = c.find((x) => /取消 次子/.test(x.head)), hana = c.find((x) => /記号/.test(x.head));
      ck(taro && taro.checked && /https:\/\/drive\.example\/P30/.test(taro.body), '1) 9/30 の本文に 9/30 のPDFのリンクが無い・チェックが無い');
      ck(tori && !tori.checked && /キャンセル（Spreading）/.test(tori.head), '1) キャンセルの方にチェックが入っている・印が無い: ' + JSON.stringify(tori));
      ck(hana && hana.head.includes('記号 <花> & "子" 様') && hana.head.includes('hana@example.com'), '1) 名前の記号が崩れた・全角のアドレスが直っていない: ' + (hana && hana.head));
      ck(hana && hana.subject === '【"見本" <定例会>】2026年9月30日 ご案内 & 資料' && hana.body.startsWith('記号 <花> & "子" 様\n<ビジターリスト> https://drive.example/P30'),
         '1) 件名・本文の記号が画面で崩れた: ' + JSON.stringify(hana && [hana.subject, hana.body]));
      await sendAll(page);
      ck(dialogs.length === 1 && /^confirm:/.test(dialogs[0]) && !/すでに送って|キャンセルになっています/.test(dialogs[0]), '1) 送る前の確認が違う: ' + dialogs.join(' / '));
      ck(env.mail.map((m) => m.to).join(',') === 'taro@example.com,hana@example.com', '1) 送った宛先が違う（キャンセルの方には送らない）: ' + env.mail.map((m) => m.to).join(','));
      const h = env.mail.find((m) => m.to === 'hana@example.com');
      ck(h && h.subject === '【"見本" <定例会>】2026年9月30日 ご案内 & 資料' && h.body.includes('<ビジターリスト> https://drive.example/P30'), '1) 送ったメールの件名・本文が画面と違う: ' + JSON.stringify(h));
      await page.close();
    }

    // ---- 2) 開き直す：送った方は「送信済み」でチェックなし。わざと送り直すと確認で知らせる ----
    {
      const { page, dialogs } = await open(browser);
      const c = await cards(page);
      ck(c.every((x) => !x.checked), '2) 開き直したら、送った方・キャンセルの方にチェックが入っている: ' + JSON.stringify(c.map((x) => [x.head.slice(0, 20), x.checked])));
      ck(c.filter((x) => /送信済み \d+\/\d+ \d\d:\d\d/.test(x.head)).length === 2, '2) 送信済みの印が出ない: ' + c.map((x) => x.head).join(' | '));
      const i = c.findIndex((x) => /見本 太郎/.test(x.head));
      await page.check('#chk_' + i);
      const before = env.mail.length;
      await sendAll(page);
      ck(dialogs.length === 1 && /すでに送っています/.test(dialogs[0]), '2) 送り直すときに確認で知らせない: ' + dialogs.join(' / '));
      ck(env.mail.length === before + 1 && env.mail[env.mail.length - 1].to === 'taro@example.com', '2) 送り直した数・宛先が違う');
      await page.close();
    }

    // ---- 2b) 代理送信の Web App が動いていない：1人目で止めて理由を出す。「自分のGmailから送る」で送れる ----
    start(true);
    env.props.MAIL_WEB_APP_URL = RELAY;
    env.onFetch = () => html(200, LOGIN_PAGE);
    {
      const { page, dialogs } = await open(browser);
      await page.click('#sendBtn');
      await page.waitForFunction(() => document.getElementById('relayBox').style.display === 'block' || /すべての処理が完了しました/.test(document.body.textContent), null, { timeout: 10000 });
      ck(relayFetches() === 1, '2b) 代理送信が動かないのに、2人目以降も送ろうとした: ' + relayFetches());
      ck(env.mail.length === 0, '2b) 送れていないはずのメールが送られた');
      ck(!/🎉|すべての処理が完了しました/.test(await page.textContent('#progressArea')), '2b) 送れていないのに「完了🎉」と出た');
      ck(await page.textContent('#relayCount') === '2' && /ログイン/.test(await page.textContent('#relayWhy')) && /tester@example\.com/.test(await page.textContent('#directBtn')),
         '2b) 止めた知らせ: ' + [await page.textContent('#relayCount'), await page.textContent('#directBtn')].join(' / '));
      await page.click('#directBtn');
      await page.waitForFunction(() => /すべての処理が完了しました/.test(document.body.textContent), null, { timeout: 10000 });
      ck(dialogs.some((d) => /自分のGmail|Gmail（tester@example\.com）/.test(d)), '2b) 自分のGmailから送る前に確かめない: ' + dialogs.join(' / '));
      ck(env.mail.map((m) => m.to).join(',') === 'taro@example.com,hana@example.com' && relayFetches() === 1,
         '2b) 自分のGmailから送った宛先が違う・代理送信を通った: ' + env.mail.map((m) => m.to).join(','));
      await page.close();
    }
    env.onFetch = null;
    delete env.props.MAIL_WEB_APP_URL;

    // ---- 2c) 1人だけ送れなかった：「完了🎉」ではなく、送れなかった人数を出す ----
    start(true);
    {
      const real = srv.GmailApp.sendEmail;
      srv.GmailApp.sendEmail = (to, ...a) => { if (to === 'hana@example.com') throw new Error('見本の送信エラー'); return real(to, ...a); };
      const { page } = await open(browser);
      await page.click('#sendBtn');
      await page.waitForFunction(() => /送れなかった方がいます|すべての処理が完了しました/.test(document.getElementById('doneMsg').textContent), null, { timeout: 10000 });
      const done = await page.textContent('#doneMsg');
      ck(/送れなかった方がいます（1名）/.test(done) && !/🎉/.test(done), '2c) 送れなかった方がいるのに「完了🎉」と出た: ' + done);
      srv.GmailApp.sendEmail = real;
      await page.close();
    }

    // ---- 3) 次回のシートが無い：過ぎた回が選ばれたら赤字で知らせ、送る前の確認にも出す ----
    start(false);
    {
      const { page, dialogs } = await open(browser);
      ck(await page.inputValue('#meetSel') === '20260923参加者', '3) 既定が直近の回でない');
      const note = await page.textContent('#meetNote');
      ck(/過ぎた回（2026年9月23日）/.test(note) && /ビジターリストのPDFがありません/.test(note), '3) 過ぎた回・PDFが無いことを知らせない: ' + note);
      ck(await page.evaluate(() => getComputedStyle(document.getElementById('meetNote')).color) === 'rgb(204, 0, 0)', '3) 知らせが赤字でない');
      await sendAll(page);
      ck(dialogs.length === 1 && /過ぎた回/.test(dialogs[0]), '3) 送る前の確認に、過ぎた回であることが出ない: ' + dialogs.join(' / '));
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
  console.log('メールの確認・一括送信の画面（ブラウザ）: 検査 ' + checks + ' 件 OK: 次回のリンク・キャンセル・記号・全角アドレス・送信済みで二重に送らない・過ぎた回の知らせ・'
    + '代理送信が動かないときは止めて理由と「自分のGmailから送る」・送れなかった方がいるのに完了と出さない・代理送信を確かめる');
})().catch((e) => {
  console.log('NG 検査が止まった: ' + (e && e.stack ? e.stack.split('\n').slice(0, 2).join(' / ') : e));
  fails.forEach((f) => console.log('  - ' + f));   // 止まる前に見つかったもの
  process.exit(1);
});

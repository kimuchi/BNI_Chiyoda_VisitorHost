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
async function sendAll(page) {
  await page.click('#sendBtn');
  await page.waitForFunction(() => /すべての処理が完了しました/.test(document.body.textContent), null, { timeout: 10000 });
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
  console.log('メールの確認・一括送信の画面（ブラウザ）: 検査 ' + checks + ' 件 OK: 次回のリンク・キャンセル・記号・全角アドレス・送信済みで二重に送らない・過ぎた回の知らせ');
})().catch((e) => {
  console.log('NG 検査が止まった: ' + (e && e.stack ? e.stack.split('\n').slice(0, 2).join(' / ') : e));
  fails.forEach((f) => console.log('  - ' + f));   // 止まる前に見つかったもの
  process.exit(1);
});

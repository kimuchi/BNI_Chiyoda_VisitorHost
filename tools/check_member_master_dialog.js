// 「⚙️ 設定」＞「メンバー名簿」の画面（member_master.html）を、本物のブラウザ（Chromium）で動かして確かめる。
// サーバー側は本番の *.js を全部、見せかけのスプレッドシート（lib_sheet_fake.js）の上で動かす。
//
//   node tools/check_member_master_dialog.js
//
//   ・最初の読み込みに失敗したまま「＋ 行を追加」「名簿を保存」を押しても、名簿を上書きしない
//     （以前は、名簿全体がその1行だけになり、控えも無かった）
//   ・サーバーも、0名の名簿は保存しない
//   ・読み込めたときは、これまでどおり追加・保存できる
//   ・表計算からの貼り付けで、セルの中の改行（一言コメントの2行目）を別のメンバーにしない
//   ・シートに列を足しても（氏名のうしろに「メール」）、ほかの列をずれて読み書きしない。名簿全体の保存でその列を消さない
//   ・旧「メンバーリスト」からの移行は1回だけ（済んで非表示の古いシートを、もう一度移さない）
//   ・画面を開いたあとで、ほかの方が名簿を変えていたら（取り込み・期の替わり目の役職の反映など）、古い一覧で上書きしない。
//     開き直せば保存でき、自分の保存のあとに続けて保存しても止まらない
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

// 0) 表計算からの貼り付け：セルの中の改行（一言コメントの2行目）で、別のメンバーを作らない・一言を切らない
{
  const row = (no, name, comment, collab) => HEAD.map((h) => (h === 'No' ? no : h === '氏名' ? name : h === '一言コメント' ? comment : h === '協業したい人' ? collab : '')).join('\t');
  const tsv = row('1', '見本 一郎', '"1行目\n2行目に ""引用"" も"', '税理士・弁護士') + '\r\n' + row('2', '試験 花子', '一言その2', '') + '\r\n';
  const r = srv.importMemberTsv(tsv, false);
  const v = env.values('メンバー名簿'), ci = HEAD.indexOf('一言コメント');
  ck(r.ok && names() === '見本 一郎,試験 花子,架空 三郎', '0) 貼り付けのセルの中の改行で、別のメンバーができた: ' + names() + ' / ' + r.message);
  ck(v[1][ci] === '1行目\n2行目に "引用" も', '0) 貼り付けの一言コメントが切れた: ' + JSON.stringify(v[1][ci]));
  // 以降の検査のため、元に戻す
  env.reset([['メンバー名簿', false, [HEAD].concat(ROWS.map(([no, n, c]) => HEAD.map((h) => (h === 'No' ? no : h === '氏名' ? n : h === '一言コメント' ? c : ''))))]], {});
}

// 0b) シートに列を足した（氏名のうしろに「メール」）：ほかの列をずれて読まない・書かない。名簿全体の保存で、その列を消さない
{
  const H2 = HEAD.slice(0, 3).concat(['メール'], HEAD.slice(3));
  const val = (h) => ({ No: '1', 氏名: '見本 一郎', メール: 'ichiro@example.com', ふりがな: 'みほん いちろう', 会社名: '見本社', 役職: 'プレジデント', 一言コメント: '一言その1' }[h] || '');
  env.reset([['メンバー名簿', false, [H2, H2.map(val), H2.map((h) => (h === 'No' ? '2' : h === '氏名' ? '試験 花子' : ''))]]],
            { BNI_ROLE_HOLDERS_TERMS: JSON.stringify({ 23: { president: '試験 花子' } }) });
  const m = srv.getMemberMaster().members[0] || {};
  ck(m.kana === 'みほん いちろう' && m.company === '見本社' && m.role === 'プレジデント' && m.comment === '一言その1',
     '0b) 列を足したシートを、ずれて読んだ: ' + JSON.stringify(m));
  ck(srv.getMembersList().map((x) => x.no + ':' + x.name).join(',') === '1:見本 一郎,2:試験 花子', '0b) 列を足したシートの番号と氏名: ' + JSON.stringify(srv.getMembersList()));
  const r = srv.saveMemberMaster(srv.getMemberMaster().members, null);
  ck(r.ok === false && /D列「メール」/.test(r.message) && env.values('メンバー名簿')[1][3] === 'ichiro@example.com',
     '0b) 名簿全体の保存で、足した列を消した・知らせない: ' + JSON.stringify(r));
  const b = srv.saveMemberBookMember('見本 一郎', { name: '見本 一郎', company: '新しい見本社' });
  const row = env.values('メンバー名簿')[1];
  ck(b.ok && row[H2.indexOf('会社名')] === '新しい見本社' && row[3] === 'ichiro@example.com' && env.values('メンバー名簿')[0].join(',') === H2.join(','),
     '0b) 1人ぶんの保存で、別の列に書いた・見出しを書き換えた: ' + JSON.stringify(row));
  const plan = srv.roleRosterPlan_(23);
  ck(plan.col === H2.indexOf('役職') + 1 && plan.changes.some((c) => c.name === '試験 花子' && c.to === 'プレジデント'),
     '0b) 期の役職の反映で、別の列を「役職」として扱った: ' + JSON.stringify({ col: plan.col, changes: plan.changes }));
  env.reset([['メンバー名簿', false, [HEAD].concat(ROWS.map(([no, n, c]) => HEAD.map((h) => (h === 'No' ? no : h === '氏名' ? n : h === '一言コメント' ? c : ''))))]], {});
}

// 0c) 旧「メンバーリスト」からの移行は1回だけ（済んで非表示の旧シートを、もう一度移さない）
{
  env.reset([['メンバー名簿', false, [HEAD].concat(ROWS.map(([no, n, c]) => HEAD.map((h) => (h === 'No' ? no : h === '氏名' ? n : h === '一言コメント' ? c : ''))))],
             ['メンバーリスト', true, [['No', '氏名'], ['1', '見本 一郎'], ['2', '退会 済子'], ['3', '試験 花子']]]], {});
  const r = srv.migrateMemberListSheet();
  ck(r.ok === false && /もう済んでいます/.test(r.message) && names() === '見本 一郎,試験 花子,架空 三郎',
     '0c) 済んだ移行をもう一度して、退会した方を名簿に戻した: ' + names() + ' / ' + r.message);
  env.reset([['メンバー名簿', false, [HEAD].concat(ROWS.map(([no, n, c]) => HEAD.map((h) => (h === 'No' ? no : h === '氏名' ? n : h === '一言コメント' ? c : ''))))]], {});
}

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

    // 4) 画面を開いたあとで、ほかの方が名簿を変えた（取り込み・期の替わり目の役職の反映など）：古い一覧で上書きしない
    const p3 = await openGasPage(browser, 'member_master.html', srv, { fails, calls });
    await p3.page.waitForFunction(() => document.querySelectorAll('#body tr').length === 4, null, { timeout: 10000 });
    const cur = srv.getMemberMaster().members;
    srv.saveMemberMaster(cur.concat([{ no: '9', name: '取込 五郎' }]), null);          // ほかの方の取り込み（サーバーの中から）
    await p3.page.locator('#body tr').nth(0).locator('input[data-k="comment"]').fill('古い画面で直した一言');
    await p3.page.locator('#body tr').nth(0).locator('input[data-k="comment"]').dispatchEvent('change');
    await p3.page.click('#btnSave');
    await p3.page.waitForFunction(() => /ほかの方|保存しました/.test(document.getElementById('msg').textContent), null, { timeout: 10000 });
    ck(/ほかの方/.test(await p3.page.textContent('#msg')), '4) ほかの方が変えたことを知らせない: ' + await p3.page.textContent('#msg'));
    ck(names() === '見本 一郎,試験 花子,架空 三郎,新入 四郎,取込 五郎', '4) 古い画面の保存で、ほかの方が取り込んだメンバーが消えた: ' + names());
    ck(!env.values('メンバー名簿').some((r) => r.includes('古い画面で直した一言')), '4) 古い画面の内容で上書きした');
    await p3.page.close();

    // 5) 開き直せば保存できる。続けてもう一度保存しても（自分の保存のあとでも）止まらない
    const p4 = await openGasPage(browser, 'member_master.html', srv, { fails, calls });
    await p4.page.waitForFunction(() => document.querySelectorAll('#body tr').length === 5, null, { timeout: 10000 });
    for (const text of ['1回目の一言', '2回目の一言']) {
      await p4.page.locator('#body tr').nth(0).locator('input[data-k="comment"]').fill(text);
      await p4.page.locator('#body tr').nth(0).locator('input[data-k="comment"]').dispatchEvent('change');
      await p4.page.evaluate(() => { document.getElementById('msg').textContent = ''; });
      await p4.page.click('#btnSave');
      await p4.page.waitForFunction(() => /保存しました|ほかの方/.test(document.getElementById('msg').textContent), null, { timeout: 10000 });
      ck(/保存しました/.test(await p4.page.textContent('#msg')) && env.values('メンバー名簿')[1][HEAD.indexOf('一言コメント')] === text,
         '5) 開き直した画面で保存できない（' + text + '）: ' + await p4.page.textContent('#msg'));
    }
    await p4.page.close();
  } finally {
    await browser.close();
  }

  if (fails.length) {
    console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
    fails.forEach((f) => console.log('  - ' + f));
    process.exit(1);
  }
  console.log('メンバー名簿の画面（ブラウザ）: 検査 ' + checks + ' 件 OK: 読み込めていないときは追加・保存させない・0名は保存しない・読み込めたら保存できる・ほかの方の変更を古い一覧で消さない・貼り付けのセルの中の改行・列を足したシート・移行は1回だけ');
})().catch((e) => {
  console.log('NG 検査が止まった: ' + (e && e.stack ? e.stack.split('\n').slice(0, 2).join(' / ') : e));
  fails.forEach((f) => console.log('  - ' + f));   // 止まる前に見つかったもの
  process.exit(1);
});

// チャプターサクセスミーティング（success_srv.js・success_meeting.html）を、実データなしで確かめる。名前はすべて架空。
//
//   node tools/check_success_meeting.js
//
// 確かめること
//   1) 記録：保存・読み直し。前月のアクションアイテムは、前の月に入れた「今月のアクションアイテム」。
//      ほかの方が先に同じ欄を書き換えていたら上書きしない。書記・出席者
//   2) 推定：ビジター数（参加者シートの V・G。キャンセルは数えない）・新メンバー（入会日が前月）・
//      パスポートプログラム（入会から4か月以内）・7ヶ月レビュー・更新の承認（更新期限が今月・来月）
//   3) Excel の読み取り：2025年の書き方（A・C・E）と2026年4月からの書き方（A・B・C・D）。議題の名前で合わせる・書記
//   4) 画面共有用のシート：Excel と同じ並び（議題・前月のアクションアイテム・前月の結果・今月のアクションアイテム・出席者）
//   5) 画面（ブラウザ）：読み込み・記入の状況・推定を入れる・保存で変えた議題だけ送る

process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeEnv } = require('./lib_sheet_fake');
const pw = (() => { try { return require('playwright'); } catch (e) { return require('/opt/node22/lib/node_modules/playwright'); } })();

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const J = (x) => JSON.stringify(x);

const HEAD = ['No', '業種区分', '氏名', 'ふりがな', 'カテゴリー', '会社名', '役職', 'メモ', '写真ファイル名', '一言コメント',
  '紹介してほしい人', '協業したい人', '入会日', '更新日', '更新期限日', '会社での役職'];
const ROSTER = [['見本 一郎', '2025/01/10', '2026/12/20'], ['試験 二郎', '2026/10/05', '2027/10/04'], ['架空 三郎', '2026/04/03', '2027/04/02'],
                ['仮名 四郎', '2024/06/01', '2026/11/30']];
const rows = [HEAD].concat(ROSTER.map(([n, j, x], i) => HEAD.map((h) => ({ No: String(i + 1), 氏名: n, 入会日: j, 更新期限日: x }[h] || ''))));
const visitorSheet = [['', '', ''], ['No.', '参加者氏名', '招待者', 'ステータス'], ['V1', '来訪 A', '見本 一郎', ''], ['V2', '来訪 B', '試験 二郎', 'キャンセル'],
                      ['G1', '来訪 C', '', ''], ['V3', '来訪 D', '', '']];
const env = makeEnv({ now: new Date(2026, 10, 8, 10, 0, 0) });
const F = Object.assign({}, env.globals);
vm.createContext(F);
for (const f of fs.readdirSync(ROOT).filter((x) => /\.js$/.test(x)).sort()) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), F, { filename: f });
env.reset([['メンバー名簿', false, rows], ['20261007参加者', false, visitorSheet], ['20261014参加者', false, visitorSheet], ['20261104参加者', false, visitorSheet]],
          { BNI_CHAPTER: J({ name: '見本', termBase: 23, meetingBaseDate: '2026/03/18', meetingBaseCount: 509 }) });
F.roleHolderTerms_ = () => ({});
F.roleHoldersOfTerm_ = () => ({ holders: { secretary: '見本 一郎' } });

// ===== 1) 記録 =====
let r = F.saveSuccessMeeting('2026-10', [{ key: 'vhc', result: '10月の結果', action: '10月のアクション' }], { secretary: '見本 一郎', attendees: ['見本 一郎'] });
ck(r.ok && /保存しました（1件）/.test(r.message), '1) 10月の保存: ' + J(r.message));
r = F.getSuccessMeeting('2026-11');
const vhc = r.blocks.find((b) => b.key === 'vhc');
ck(r.ok && r.label === '2026年11月' && r.prevLabel === '2026年10月' && vhc.prevAction === '10月のアクション' && vhc.result === '' && vhc.action === '',
   '1) 前月のアクションアイテム: ' + J(vhc));
ck(r.blocks.length === F.SM_BLOCKS_.length && r.blocks[r.blocks.length - 1].key === 'admin', '1) 議題の並び');
ck(r.secretary === '見本 一郎' && J(r.attendees) === '[]' && r.roster.length === 4, '1) 書記の初期値（その期の書記兼会計）・出席者: ' + J([r.secretary, r.attendees]));
r = F.saveSuccessMeeting('2026-11', [{ key: 'vhc', result: '11月の結果', action: '', origResult: '', origAction: '' }], { secretary: '試験 二郎', attendees: ['試験 二郎', '架空 三郎'] });
ck(r.ok && r.blocks.find((b) => b.key === 'vhc').result === '11月の結果' && r.blocks.find((b) => b.key === 'vhc').by === 'tester@example.com'
   && r.secretary === '試験 二郎' && J(r.attendees) === J(['試験 二郎', '架空 三郎']), '1) 11月の保存・書いた方・書記と出席者: ' + J(r.blocks[0]));
// ほかの方が先に書き換えていた
r = F.saveSuccessMeeting('2026-11', [{ key: 'vhc', result: '上書き', action: '新しいアクション', origResult: '', origAction: '' }]);
const v2 = r.blocks.find((b) => b.key === 'vhc');
ck(r.ok && v2.result === '11月の結果' && v2.action === '新しいアクション' && J(r.conflicts) === J(['【ビジターホストコーディネーター】（結果）']),
   '1) 先に書き換えられていた欄を上書きした: ' + J([v2, r.conflicts]));
ck(!F.saveSuccessMeeting('11月', []).ok, '1) 月の書き方が違うのに保存した');
// 月は文字のまま（' を付けて）書く。スプレッドシートが日付に変えてしまっていても読む（前は読めず、取り込んだ月が出なかった）
{
  const sh = env.sheet('サクセスミーティング記録'), v = sh.getDataRange().getValues();
  ck(v.slice(1).every((row) => /^'\d{4}-\d{2}$/.test(String(row[0]))), '1) 月を文字のまま書いていない: ' + J(v.slice(1).map((row) => row[0])));
  v.forEach((row, i) => { if (i && String(row[0]) === "'2026-10") sh.getRange(i + 1, 1).setValue(new Date(2026, 9, 1)); });
  const rr = F.getSuccessMeeting('2026-11');
  ck(rr.blocks.find((b) => b.key === 'vhc').prevAction === '10月のアクション', '1) 日付になった月を読めない');
  // 選べる月：記録のある月より前（1年前）から来月まで。前の月までの記録
  ck(rr.months[0].value === '2026-12' && rr.months.some((m) => m.value === '2025-11') && rr.months.find((m) => m.value === '2026-10').saved,
     '1) 選べる月: ' + J(rr.months.map((m) => m.value)));
  const hv = rr.blocks.find((b) => b.key === 'vhc').history;
  ck(J(hv) === J([{ month: '2026-10', label: '2026年10月', result: '10月の結果', action: '10月のアクション' }]), '1) 前の月までの記録: ' + J(hv));
}

// ===== 2) 推定 =====
{
  const a = F.smAuto_('2026-11');
  ck(a.vhc === '◆10月ビジター数　ビジター4名　ゲスト2名（定例会 2回）', '2) ビジター数: ' + J(a.vhc));
  ck(/◆新メンバー　1名（試験 二郎さん）/.test(a.membership) && /◆純増　\+1名/.test(a.membership), '2) 新メンバー・純増: ' + J(a.membership));
  ck(a.mentor === '□BNIパスポートプログラム（入会から4か月以内）：試験 二郎さん', '2) パスポートプログラム: ' + J(a.mentor));
  ck(/□7ヶ月レビュー：架空 三郎さん/.test(a.admin) && /□更新の承認（更新期限が今月・来月）：見本 一郎さん（12\/20）、仮名 四郎さん（11\/30）/.test(a.admin),
     '2) 7ヶ月レビュー・更新の承認: ' + J(a.admin));
  ck(F.getSuccessMeeting('2026-11').blocks.find((b) => b.key === 'mentor').auto === a.mentor, '2) 画面に推定が渡らない');
}

// ===== 3) Excel の読み取り =====
{
  const old = [['チャプターサクセスミーティングフォーム'], ['チャプター:見本　書記: 　見本一郎　日付:2025.07.02'], ['参加者:'], [], [], [],
    ['6月のアクションアイテム', '', '6月の結果', '', '7月のアクションアイテム'],
    ['□前回のミーティングのアクション', '', ' ', '', ' '],
    ['【ビジターホストコーディネーター】\n◆詳細', '', 'ビジター結果', '', 'ビジターの次'],
    ['【メンバーシップ】\n□レビュー', '', '会員結果1', '', '会員次1'],
    ['【メンバーシップ】\n□レポート', '', '会員結果2', '', '会員次2'],
    ['【1 to 1促進委員】', '', '1to1結果', '', ''],
    ['【見本の新しい委員】', '', '新しい結果', '', '新しい次'],
    ['□ 7 か月レビュー\n氏名：', '', '', '', ''], ['□ 更新の承認', '', '承認の結果', '', '']];
  const g1 = F.smParseGrid_(old);
  ck(g1.secretary === '見本一郎' && g1.blocks.vhc.result === 'ビジター結果' && g1.blocks.vhc.action === 'ビジターの次'
     && g1.blocks.membership.result === '会員結果1\n会員結果2' && g1.blocks.one2one.result === '1to1結果'
     && g1.blocks['x:【見本の新しい委員】'].action === '新しい次' && g1.blocks.admin.result === '承認の結果', '3) 2025年の書き方: ' + J(g1));
  const cur = [['チャプターサクセスミーティングフォーム', '書記：試験'], [], ['', '2026年3月のアクションアイテム', '2026年3月の結果', '2026年4月のアクションアイテム', '', '本日の出席者'],
    ['【ビジターホストコーディネーター】', '前のアクション', '3月の結果', '4月のアクション', '', '1　見本 一郎'],
    ['□ビジター出席レビュー', '', '', '', '', '2　試験 二郎'], ['【WEBマスター】', '', 'ウェブ結果', 'ウェブ次', '', ''],
    ['【メンターコーディネーター】', '', 'メンター結果', 'メンター次'], ['　□BNIパスポートプログラム', '', '', ''],
    ['メンターコーディネーターとの連携（運営マニュアルチェック）', '', '', ''], ['　□7ヶ月レビュー', '', '', '']];
  const g2 = F.smParseGrid_(cur);
  ck(g2.secretary === '試験' && g2.blocks.vhc.result === '3月の結果' && g2.blocks.vhc.action === '4月のアクション' && g2.blocks.web.action === 'ウェブ次'
     && g2.blocks.mentor.result === 'メンター結果' && J(g2.order) === J(['vhc', 'web', 'mentor', 'admin']), '3) 2026年4月からの書き方: ' + J(g2));
  ck(F.smMonthOfSheetName_('2026年11月 ') === '2026-11' && F.smMonthOfSheetName_('出席者一覧') === '', '3) シートの名前から月');
}

// ===== 4) 画面共有用のシート =====
{
  const e = F.exportSuccessMeetingSheet('2026-11');
  const v = env.sheet('サクセスミーティング').getDataRange().getValues();
  ck(e.ok && /2026年11月の内容を書き出しました/.test(e.message), '4) 書き出しの知らせ: ' + J(e.message));
  ck(J(v[2].slice(0, 4)) === J(['議題', '2026年10月のアクションアイテム', '2026年10月の結果', '2026年11月のアクションアイテム']) && v[2][5] === '本日の出席者', '4) 見出し: ' + J(v[2]));
  ck(v[3][0] === '【ビジターホストコーディネーター】' && v[3][1] === '10月のアクション' && v[3][2] === '11月の結果' && v[3][3] === '新しいアクション'
     && v[4][0] === '　□ビジター出席レビュー' && v[3][5] === '1　試験 二郎' && v[4][5] === '2　架空 三郎', '4) 議題の行・出席者: ' + J([v[3], v[4]]));
  ck(/書記：試験 二郎/.test(v[0][1]), '4) 書記: ' + J(v[0]));
}

// ===== 5) 画面 =====
(async () => {
  const D = F.getSuccessMeeting('2026-11');
  const STUB = '<script>(function(){window.__calls=[];var D=' + JSON.stringify(D) + ';'
    + 'function answer(n,a){if(n==="getSuccessMeeting")return D;'
    + 'if(n==="saveSuccessMeeting"){var r=JSON.parse(JSON.stringify(D));r.message="保存しました";r.conflicts=[];return r;}return {ok:true,message:"ok"};}'
    + 'window.google={script:{run:new Proxy({},{get:function(_,n){var ok=null,p=new Proxy({},{get:function(_,m){'
    + 'if(m==="withSuccessHandler")return function(f){ok=f;return p;};if(m==="withFailureHandler")return function(){return p;};'
    + 'return function(){var a=[].slice.call(arguments);window.__calls.push([m,a]);var v=answer(m,a);setTimeout(function(){ok&&ok(JSON.parse(JSON.stringify(v)));},5);};}});'
    + 'return p[n];}})}};})();</script>';
  const html = fs.readFileSync(path.join(ROOT, 'success_meeting.html'), 'utf8').replace(/<head>/i, '<head><meta charset="utf-8">' + STUB);
  const browser = await pw.chromium.launch();
  const page = await browser.newPage();
  page.on('pageerror', (e) => fails.push('5) 画面のエラー: ' + e.message));
  page.on('dialog', (d) => d.accept());
  await page.route('https://sm.test/', (rt) => rt.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: html }));
  await page.goto('https://sm.test/');
  await page.waitForFunction(() => document.querySelectorAll('.blk').length > 0, null, { timeout: 5000 }).catch(() => fails.push('5) 議題が出ない'));
  const txt = (id) => page.evaluate((i) => document.getElementById(i).innerText, id);
  ck(/結果 1\/13・アクションアイテム 1\/13/.test(await txt('prog')), '5) 記入の状況: ' + await txt('prog'));
  ck(/10月のアクション/.test(await txt('b_0')) && /入力済み/.test(await txt('b_0')), '5) 前月のアクションアイテム・入力済み: ' + (await txt('b_0')).slice(0, 200));
  const mi = D.blocks.findIndex((b) => b.key === 'mentor');
  ck(/推定：□BNIパスポートプログラム/.test(await txt('b_' + mi)), '5) 推定が出ない');
  await page.click('#b_' + mi + ' .auto button.chip');
  await page.fill('#a_1', 'メンバーシップの次');
  await page.dispatchEvent('#a_1', 'input');
  ck(/未保存/.test(await txt('b_' + mi)) && /保存していない入力があります/.test(await txt('prog')), '5) 未保存の印');
  await page.click('#btnSave');
  await page.waitForTimeout(100);
  const sent = (await page.evaluate(() => window.__calls)).filter((c) => c[0] === 'saveSuccessMeeting').map((c) => c[1]);
  const items = sent.length ? sent[0][1] : [];
  ck(sent.length === 1 && J(items.map((x) => x.key).sort()) === J(['membership', 'mentor']) && /パスポート/.test(items.find((x) => x.key === 'mentor').result)
     && sent[0][2].secretary === '試験 二郎' && J(sent[0][2].attendees) === J(['試験 二郎', '架空 三郎']), '5) 保存で送るもの: ' + J(sent));
  // 前月のアクションアイテムをコピー・過去の記録を見る・まとめてコピー
  await page.click('#b_0 button:has-text("前月のアクションアイテムをコピー")');
  ck((await page.inputValue('#a_0')) === '新しいアクション\n10月のアクション', '5) 前月のアクションアイテムをコピー（入力のうしろに足す）: ' + J(await page.inputValue('#a_0')));
  await page.click('#b_0 button:has-text("過去の記録を見る")');
  ck(/2026年10月/.test(await txt('b_0')) && /10月の結果/.test(await txt('b_0')), '5) 過去の記録が出ない');
  await page.click('#b_0 button:has-text("結果の欄にコピー")');
  ck((await page.inputValue('#r_0')) === '11月の結果\n10月の結果', '5) 過去の記録をコピー: ' + J(await page.inputValue('#r_0')));
  await browser.close();
  if (fails.length) { console.log('NG ' + fails.length + '件 / ' + checks + '件の検査'); fails.forEach((f) => console.log('  - ' + f)); process.exit(1); }
  console.log('チャプターサクセスミーティング: 検査 ' + checks + ' 件 OK');
})().catch((e) => { console.error(e); process.exit(1); });

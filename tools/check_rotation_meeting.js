// 書記兼会計のスピーカーローテーションの画面（role_input.html）で、上の「定例会」を変えたときの動きを、
// 本物のブラウザ（Chromium）で確かめる。サーバーの返事は作り物（名前はすべて架空）。
//
//   node tools/check_rotation_meeting.js
//
// 確かめること
//   ・ローテーションの画面のまま（書記兼会計の入力・入力状況の一覧に移らない）。
//     ローテーションを直接開いたとき（?p=rotation）も、書記兼会計の入力の「🔁 スピーカーローテーションの管理」から開いたときも
//   ・ご案内する回がその日の回になり、投稿文・メインプレゼンターの画像・ローテーションの表の画像もその回で作り直す
//   ・上の帯（第何回・シートの名前）はその日の分を読む。書記兼会計の初期値の推定は、まだ読まない
//   ・「← 書記兼会計の入力」で、選んだ日の書記兼会計の入力を読む（前の日の画面を出さない）
//   ・並び順を保存して読み直しても、ご案内する回はそのまま（選び直した回も）
//   ・書記兼会計の入力に保存していないものがあれば確かめる（やめたら日を戻す）
//   ・読み込みが遅いとき：ローテーションを読み終わる前に日を変えた・読み終わる前に戻った・
//     続けて2回変えた（ローテーションの画面でも、書記兼会計の入力の画面でも、あとで選んだ日になる）
const fs = require('fs');
const path = require('path');
const pw = (() => { try { return require('playwright'); } catch (e) { return require('/opt/node22/lib/node_modules/playwright'); } })();

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const J = (x) => JSON.stringify(x);

// --- サーバーの返事（作り物）---
const MEETINGS = [
  { dateValue: '2031/10/08', no: '536', display: '10/8(水) 第536回' },
  { dateValue: '2031/10/15', no: '537', display: '10/15(水) 第537回' },
  { dateValue: '2031/10/22', no: '538', display: '10/22(水) 第538回' },
  { dateValue: '2031/10/29', no: '539', display: '10/29(水) 第539回' },
];
const NAMES = ['見本 一郎', '試験 二郎', '架空 三郎', '仮名 四郎', '例題 六郎', '模擬 七郎', '検査 八郎', '標本 九郎', '見本 十郎', '試験 十一郎', '架空 十二郎', '仮名 十三郎'];
const member = (n, i) => ({ name: n, title: '見本の業種' + (i + 1), collab: '見本のつながりたい人' + (i + 1) });
const MEMBERS = NAMES.map(member);
const DATES = ['2031/10/08', '2031/10/15', '2031/10/22', '2031/10/29', '2031/11/05', '2031/11/12'];
const WEEKS = DATES.map((d, i) => {
  const [, m, day] = d.split('/').map(Number);
  return { date: d, no: String(536 + i), md: m + '/' + day + '(水)', label: m + '月' + day + '日', source: 'routine', secretary: '例示 五郎',
           people: [MEMBERS[i * 2], MEMBERS[i * 2 + 1]].map((p) => ({ name: p.name, title: p.title, collab: p.collab, company: '', inMaster: true })) };
});
const ROT = {
  ok: true, order: NAMES.slice(), excluded: [], anchor: { date: DATES[0], pointer: 0 },
  header: 'メインプレゼンテーション（各４分45秒）', notes: ['見本の注意書き'], updated: '2031/09/01 10:00', provisional: false,
  rebased: false, openDate: DATES[0], missing: [], holidayHints: [], fbText: '', secretary: '例示 五郎', holidays: [],
  chapter: '見本チャプター', weeks: WEEKS,
  members: MEMBERS.map((m) => ({ name: m.name, title: m.title, collab: m.collab, inOrder: true })),
};

// google.script.run の代わり。呼ばれた関数と引数を window.__calls に残す。
// 返事までの時間は window.__delay（関数の名前 → ミリ秒）で変えられる
function stub(delay) {
  const data = J({ MEETINGS, ROT, delay: delay || {} }).replace(/</g, '\\u003c');
  return '<script>(function(){var D=' + data + ';window.__calls=[];window.__delay=D.delay;'
    + 'function ctxOf(date,role){var m=D.MEETINGS.filter(function(x){return x.dateValue===date;})[0]||D.MEETINGS[0];'
    + 'return {ok:true,date:m.dateValue,display:m.display,today:"2031/10/01",estimatedFor:role||"",found:true,sheetName:"見本の期",meetingNo:m.no,message:"",'
    + 'meetings:D.MEETINGS,roles:[{key:"secretary",label:"書記兼会計",holder:"例示 五郎",items:["day","memo"],status:{state:"todo",missing:[]}}],'
    + 'items:{day:{id:"day",title:"開催日",value:"開催日 "+m.dateValue,orig:"開催日 "+m.dateValue,readOnly:true,readOnlyWhy:"検査用"},'
    + 'memo:{id:"memo",title:"メモ",value:"",orig:"",group:"",prev:null,history:[]}}};}'
    + 'function answer(n,a){if(n==="getRoleInputContext")return ctxOf(a[0],a[1]);'
    + 'if(n==="getSpeakerRotation"||n==="saveSpeakerRotation")return JSON.parse(JSON.stringify(D.ROT));'
    + 'if(n==="getMemberPhotosBase64")return {ok:true,map:{}};if(n==="getSystemVersion")return "";return null;}'
    + 'window.google={script:{host:{close:function(){}},get run(){var ok=null,p=new Proxy({},{get:function(_,n){'
    + 'if(n==="withSuccessHandler")return function(f){ok=f;return p;};'
    + 'if(n==="withFailureHandler"||n==="withUserObject")return function(){return p;};'
    + 'return function(){var a=[].slice.call(arguments);window.__calls.push([n,a]);var v=answer(n,a);'
    + 'setTimeout(function(){ok&&ok(v);},window.__delay[n]||10);};}});return p;}}};})();</script>';
}

// Apps Script のテンプレート（<? ?>・<?= ?>・<?!= ?>）を展開する（tools/check_mp_image.js と同じ）
function evalTemplate(file, vars) {
  const src = fs.readFileSync(path.join(ROOT, file), 'utf8');
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  let code = 'var __o=[];with(__v){', i = 0, m;
  const re = /<\?(!=|=)?([\s\S]*?)\?>/g;
  while ((m = re.exec(src))) {
    code += '__o.push(' + J(src.slice(i, m.index)) + ');';
    if (m[1] === '=') code += '__o.push(__e(' + m[2] + '));';
    else if (m[1] === '!=') code += '__o.push(String(' + m[2] + '));';
    else code += m[2] + '\n';
    i = re.lastIndex;
  }
  code += '__o.push(' + J(src.slice(i)) + ');}return __o.join("");';
  const include = (n) => fs.readFileSync(path.join(ROOT, n + '.html'), 'utf8');
  return new Function('__v', '__e', code)(Object.assign({ include }, vars), esc);
}

(async () => {
  const browser = await pw.chromium.launch();
  const URL = 'https://role-input.test/';
  // 画面を開く。confirm は page.answer（既定は OK）で答え、文面を page.dialogs に残す
  const start = async (params, delay) => {
    const html = evalTemplate('role_input.html', { params }).replace(/<head>/i, '<head><meta charset="utf-8">' + stub(delay));
    const page = await browser.newPage({ viewport: { width: 1000, height: 900 } });
    page.dialogs = []; page.answer = true;
    page.on('dialog', async (d) => { page.dialogs.push(d.message()); if (page.answer) await d.accept(); else await d.dismiss(); });
    page.on('pageerror', (e) => fails.push('画面のエラー: ' + e.message));
    await page.route(/fonts\.(googleapis|gstatic)\.com/, (r) => r.abort());
    await page.route(URL, (rt) => rt.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: html }));
    await page.goto(URL);
    return page;
  };
  const shown = (page, id) => page.evaluate((i) => { const e = document.getElementById(i); return !!e && e.style.display !== 'none'; }, id);
  const val = (page, id) => page.evaluate((i) => { const e = document.getElementById(i); return e ? e.value : null; }, id);
  const text = (page, id) => page.evaluate((i) => { const e = document.getElementById(i); return e ? e.innerText : null; }, id);
  const ctxCalls = async (page) => (await page.evaluate(() => window.__calls)).filter((c) => c[0] === 'getRoleInputContext').map((c) => c[1]);
  const waitFor = async (page, fn, arg, label, ms) => {
    try { await page.waitForFunction(fn, arg, { timeout: ms || 5000 }); return true; }
    catch (e) { fails.push(label + '（待っても、そうならない）'); return false; }
  };
  const click = async (page, sel, label) => {
    try { await page.click(sel, { timeout: 3000 }); return true; }
    catch (e) { fails.push(label + '（押せない: ' + e.message.split('\n')[0] + '）'); return false; }
  };
  const select = async (page, sel, value, label) => {
    try { await page.selectOption(sel, value, { timeout: 3000 }); return true; }
    catch (e) { fails.push(label + '（選べない: ' + e.message.split('\n')[0] + '）'); return false; }
  };
  const settle = (page, ms) => page.waitForTimeout(ms || 350);
  const rotReady = () => /^data:image\/png/.test((document.getElementById('rotImg') || {}).src || '') && document.getElementById('meeting').options.length === 4;
  const roleShows = (d) => document.getElementById('roleView').style.display !== 'none' && document.getElementById('items').innerText.indexOf('開催日 ' + d) >= 0;
  // ローテーションの画面のままで、ご案内する回が i 番目（その日の回）になっているか
  const rotOn = async (page, i, tag) => {
    const w = WEEKS[i];
    ck(await shown(page, 'rotView'), tag + ' ローテーションの画面でなくなった');
    ck(!(await shown(page, 'roleView')) && !(await shown(page, 'overview')), tag + ' 書記兼会計の入力・入力状況の一覧に移った');
    ck(await val(page, 'rotFbWeek') === String(i), tag + ' ご案内する回がその日の回にならない: ' + await val(page, 'rotFbWeek'));
    ck(String(await val(page, 'rotFb')).indexOf('【' + w.label + '定例会】') >= 0, tag + ' 投稿文がその回にならない: ' + String(await val(page, 'rotFb')).split('\n')[0]);
    ck(await val(page, 'mpName0') === w.people[0].name && await val(page, 'mpName1') === w.people[1].name,
       tag + ' メインプレゼンターの画像のお名前がその回にならない: ' + J([await val(page, 'mpName0'), await val(page, 'mpName1')]));
    ck(await page.evaluate(() => rotImgName) === w.date.replace(/\//g, '') + '_スピーカーローテーション.png',
       tag + ' ローテーションの表の画像がその回からにならない: ' + await page.evaluate(() => rotImgName));
  };
  const NOTE_MEETING = '（上の「定例会」で選んだ回です）';

  // ---- A) ローテーションを直接開いたとき（?p=rotation）----
  let page = await start({ role: 'secretary', view: 'rotation' });
  if (await waitFor(page, rotReady, null, 'A0) ローテーションの画面ができない')) {
    ck(await val(page, 'rotFbWeek') === '0' && await text(page, 'rotFbWeekNote') === '', 'A0) ご案内する回の初め: ' + await val(page, 'rotFbWeek'));
    await select(page, '#meeting', '2031/10/22', 'A1) 定例会を 2031/10/22 に');
    await settle(page);
    await rotOn(page, 2, 'A1)');
    ck(await text(page, 'rotFbWeekNote') === NOTE_MEETING, 'A1) 「定例会」で選んだ回の知らせ: ' + await text(page, 'rotFbWeekNote'));
    let cc = await ctxCalls(page);
    ck(J(cc) === J([['', ''], ['2031/10/22', '']]), 'A1) その日の分（推定なし）だけを読むはず: ' + J(cc));
    ck(await val(page, 'meeting') === '2031/10/22' && /第538回/.test(await text(page, 'sheetNote')), 'A1) 上の帯がその日にならない: ' + await text(page, 'sheetNote'));
    ck(await page.evaluate(() => ctx && ctx.date) === '2031/10/22', 'A1) 読んだ日');

    // 並び順を保存して読み直しても、ご案内する回はそのまま
    await click(page, '#rotSaveBtn', 'A2) 並び順を保存');
    await settle(page);
    ck((await page.evaluate(() => window.__calls)).some((c) => c[0] === 'saveSpeakerRotation'), 'A2) 保存を呼ばない');
    await rotOn(page, 2, 'A2) 保存のあと:');
    ck(await text(page, 'rotFbWeekNote') === NOTE_MEETING, 'A2) 保存のあとの知らせ: ' + await text(page, 'rotFbWeekNote'));

    // ご案内する回を選び直す → 知らせは消える。保存しても、選び直した回のまま
    await select(page, '#rotFbWeek', '1', 'A3) ご案内する回を 1 番目に');
    await settle(page);
    ck(await text(page, 'rotFbWeekNote') === '', 'A3) 選び直しても「定例会」で選んだ回の知らせが残る');
    await click(page, '#rotSaveBtn', 'A3) 並び順を保存');
    await settle(page);
    ck(await val(page, 'rotFbWeek') === '1' && String(await val(page, 'rotFb')).indexOf('【10月15日定例会】') >= 0,
       'A3) 保存したら、選び直したご案内する回が戻った: ' + await val(page, 'rotFbWeek'));

    // 「← 書記兼会計の入力」：上で選んでいる日（10/22）の書記兼会計の入力を読む
    await click(page, '#rotView button[onclick="closeRotation()"]', 'A4) 「← 書記兼会計の入力」');
    if (await waitFor(page, roleShows, '2031/10/22', 'A4) 選んだ日（10/22）の書記兼会計の入力が出ない')) {
      ck(!(await shown(page, 'rotView')), 'A4) ローテーションの画面が残る');
      cc = await ctxCalls(page);
      ck(J(cc[cc.length - 1]) === J(['2031/10/22', 'secretary']), 'A4) 書記兼会計の初期値の推定をその日で読まない: ' + J(cc));
    }
    // 書記兼会計の入力の画面で日を変えると、これまでどおり、その日の書記兼会計の入力
    await select(page, '#meeting', '2031/10/29', 'A5) 定例会を 2031/10/29 に');
    if (await waitFor(page, roleShows, '2031/10/29', 'A5) 書記兼会計の入力の画面で日を変えても、その日の入力にならない')) {
      ck(!(await shown(page, 'rotView')) && !(await shown(page, 'overview')), 'A5) 書記兼会計の入力の画面から移った');
    }
  }
  await page.close();

  // ---- B) 書記兼会計の入力から「🔁 スピーカーローテーションの管理」で開いたとき ----
  page = await start({ role: 'secretary' });
  if (await waitFor(page, roleShows, '2031/10/08', 'B0) 書記兼会計の入力が出ない')) {
    await click(page, '#rotOpenBtn', 'B0) 「🔁 スピーカーローテーションの管理」');
    if (await waitFor(page, rotReady, null, 'B0) ローテーションの画面ができない')) {
      await select(page, '#meeting', '2031/10/15', 'B1) 定例会を 2031/10/15 に');
      await settle(page);
      await rotOn(page, 1, 'B1)');
      await click(page, '#rotView button[onclick="closeRotation()"]', 'B2) 「← 書記兼会計の入力」');
      if (await waitFor(page, roleShows, '2031/10/15', 'B2) 選んだ日（10/15）の書記兼会計の入力が出ない（前の日の画面のまま）')) {
        ck(String(await text(page, 'items')).indexOf('2031/10/08') < 0, 'B2) 前の日の入力が残る');
        ck(!(await shown(page, 'rotView')), 'B2) ローテーションの画面が残る');
      }
      // もう一度開いても、ご案内する回は前に選んだ回のまま
      await click(page, '#rotOpenBtn', 'B3) もう一度「🔁 スピーカーローテーションの管理」');
      await settle(page, 500);
      await rotOn(page, 1, 'B3)');
    }
  }
  await page.close();

  // ---- C) 書記兼会計の入力に保存していないものがあるとき ----
  page = await start({ role: 'secretary' });
  if (await waitFor(page, roleShows, '2031/10/08', 'C0) 書記兼会計の入力が出ない')) {
    await page.fill('#f_1', '見本のメモ');
    await click(page, '#rotOpenBtn', 'C0) 「🔁 スピーカーローテーションの管理」');
    if (await waitFor(page, rotReady, null, 'C0) ローテーションの画面ができない')) {
      page.answer = false;                                   // 「キャンセル」：日を戻す
      await select(page, '#meeting', '2031/10/22', 'C1) 定例会を 2031/10/22 に');
      await settle(page);
      ck(page.dialogs.length === 1 && /保存していない入力/.test(page.dialogs[0]), 'C1) 保存していない入力があるのに確かめない: ' + J(page.dialogs));
      ck(await val(page, 'meeting') === '2031/10/08', 'C1) やめたのに日が戻らない: ' + await val(page, 'meeting'));
      ck(await shown(page, 'rotView') && await val(page, 'rotFbWeek') === '0', 'C1) やめたのに画面・ご案内する回が変わった');
      ck(J(await ctxCalls(page)) === J([['', 'secretary']]), 'C1) やめたのに読み込んだ: ' + J(await ctxCalls(page)));
      page.answer = true;                                    // 「OK」：その日に
      await select(page, '#meeting', '2031/10/22', 'C2) 定例会を 2031/10/22 に');
      await settle(page);
      ck(page.dialogs.length === 2, 'C2) 確かめない');
      await rotOn(page, 2, 'C2)');
      await click(page, '#rotView button[onclick="closeRotation()"]', 'C3) 「← 書記兼会計の入力」');
      if (await waitFor(page, roleShows, '2031/10/22', 'C3) 選んだ日（10/22）の書記兼会計の入力が出ない')) {
        ck(await val(page, 'f_1') === '', 'C3) 前の日に入れかけたメモが残る: ' + await val(page, 'f_1'));
      }
    }
  }
  await page.close();

  // ---- D) ローテーションを読み終わる前に日を変えた ----
  page = await start({ role: 'secretary', view: 'rotation' }, { getSpeakerRotation: 1500 });
  if (await waitFor(page, () => document.getElementById('meeting').options.length === 4, null, 'D0) 定例会の選択肢が出ない')) {
    await select(page, '#meeting', '2031/10/29', 'D1) 定例会を 2031/10/29 に');
    if (await waitFor(page, rotReady, null, 'D1) ローテーションの画面ができない')) {
      await settle(page);
      await rotOn(page, 3, 'D1)');
      ck(await text(page, 'rotFbWeekNote') === NOTE_MEETING, 'D1) 「定例会」で選んだ回の知らせ: ' + await text(page, 'rotFbWeekNote'));
    }
  }
  await page.close();

  // ---- E) 日を変えて、その日の分を読み終わる前に「← 書記兼会計の入力」を押した ----
  page = await start({ role: 'secretary', view: 'rotation' });
  if (await waitFor(page, rotReady, null, 'E0) ローテーションの画面ができない')) {
    await page.evaluate(() => { window.__delay.getRoleInputContext = 800; });
    await select(page, '#meeting', '2031/10/22', 'E1) 定例会を 2031/10/22 に');
    await click(page, '#rotView button[onclick="closeRotation()"]', 'E1) 「← 書記兼会計の入力」');
    if (await waitFor(page, roleShows, '2031/10/22', 'E1) 選んだ日（10/22）の書記兼会計の入力が出ない')) {
      await settle(page, 1200);                              // 遅れて届いた、日を変えたときの読み込みで上書きしない
      ck(await page.evaluate(() => ctx.date) === '2031/10/22', 'E1) 遅れて届いた読み込みで、読んだ日が変わった: ' + await page.evaluate(() => ctx.date));
      ck(await shown(page, 'roleView') && String(await text(page, 'items')).indexOf('開催日 2031/10/22') >= 0, 'E1) 遅れて届いた読み込みで、書記兼会計の入力が変わった');
      ck(await page.evaluate(() => ctx.estimatedFor) === 'secretary', 'E1) 遅れて届いた読み込み（推定なし）で、書記兼会計の推定が消えた');
    }
  }
  await page.close();

  // ---- F) ローテーションの画面で続けて2回変えた（先に選んだ日の読み込みがあとから届く）----
  page = await start({ role: 'secretary', view: 'rotation' });
  if (await waitFor(page, rotReady, null, 'F0) ローテーションの画面ができない')) {
    await page.evaluate(() => { window.__delay.getRoleInputContext = 700; });
    await select(page, '#meeting', '2031/10/15', 'F1) 定例会を 2031/10/15 に');
    await page.evaluate(() => { window.__delay.getRoleInputContext = 30; });
    await select(page, '#meeting', '2031/10/29', 'F1) 定例会を 2031/10/29 に');
    await settle(page, 1100);
    ck(await val(page, 'meeting') === '2031/10/29' && /第539回/.test(await text(page, 'sheetNote')) && await page.evaluate(() => ctx.date) === '2031/10/29',
       'F1) 先に選んだ日の読み込みが、あとから選んだ日を上書きした: ' + J([await val(page, 'meeting'), await text(page, 'sheetNote')]));
    await rotOn(page, 3, 'F1)');
  }
  await page.close();

  // ---- G) 書記兼会計の入力の画面で続けて2回変えた（先に選んだ日の読み込みがあとから届く）----
  page = await start({ role: 'secretary' });
  if (await waitFor(page, roleShows, '2031/10/08', 'G0) 書記兼会計の入力が出ない')) {
    await page.evaluate(() => { window.__delay.getRoleInputContext = 700; });
    await select(page, '#meeting', '2031/10/15', 'G1) 定例会を 2031/10/15 に');
    await page.evaluate(() => { window.__delay.getRoleInputContext = 30; });
    await select(page, '#meeting', '2031/10/29', 'G1) 定例会を 2031/10/29 に');
    await settle(page, 1100);
    ck(await val(page, 'meeting') === '2031/10/29' && String(await text(page, 'items')).indexOf('開催日 2031/10/29') >= 0,
       'G1) 先に選んだ日の読み込みが、あとから選んだ日の入力を上書きした: ' + J([await val(page, 'meeting'), String(await text(page, 'items')).match(/開催日 [\d/]+/)]));
    ck(!(await shown(page, 'loading')) && !(await page.evaluate(() => busy)), 'G1) 読み込み中のまま');
  }
  await page.close();

  await browser.close();
  if (fails.length) {
    console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
    fails.forEach((f) => console.log('  - ' + f));
    process.exit(1);
  }
  console.log('スピーカーローテーションの画面で定例会を変える: 検査 ' + checks + ' 件 OK: 画面はそのまま（直接開いたとき・書記兼会計の入力から開いたとき）・'
    + 'ご案内する回と画像がその日の回に・戻ると選んだ日の書記兼会計の入力・保存しても回はそのまま・保存していない入力の確かめ・読み込みが遅いとき・続けて2回変えたとき');
})().catch((e) => { console.error(e); process.exit(1); });

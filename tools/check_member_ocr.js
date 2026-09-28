// メンバーリスト(OCR)の取り込みと、名簿の「取り込む前に戻す」を確かめる。名簿は作り物。Gemini には送らない。
//
//   node tools/check_member_ocr.js
//
// 確かめること
//   ・名簿の行は氏名で探す（番号では探さない）。Spreadingで新しい方を足したあと、PDFでは番号が1つずつずれていても、
//     写真・一言・紹介・協業・日付・会社での役職は、その方の行に残る（別の方の名前に書き換わらない）。番号はPDFの番号になる
//   ・氏名が1文字違う方（読み取りの誤り・異体字。3文字以上で、名簿に当てはまる方が1人だけ）は、その方として扱う（氏名は名簿のまま）。
//     2文字の氏名・当てはまる方が2人以上のときは、新しい方として足す
//   ・PDFに無い方は消さない。番号が重なる方・氏名を読み取れなかった行は知らせる
//   ・読み取っただけでは名簿を変えない（確認の一覧を返す）。「名簿に反映」で変える。ダイアログを使わない取り込みも、確かめてから
//   ・取り込みの前に名簿を控え、「取り込む前に戻す」で入れ替えられる（もう一度押すと戻す前に戻る）。Spreadingの取り込みも控える
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const J = (x) => JSON.stringify(x);

// ---------- 書き込めるスプレッドシート（必要なところだけ）----------
function makeSheet(name, grid) {
  let g = (grid || []).map((r) => r.slice()), hidden = false, maxCols = 26;
  const width = () => Math.max(0, ...g.map((r) => r.length));
  const sh = {
    _grid: () => g, isHidden: () => hidden,
    getName: () => name,
    getLastRow: () => { for (let r = g.length - 1; r >= 0; r--) if (g[r].some((v) => v !== '' && v != null)) return r + 1; return 0; },
    getLastColumn: () => width(),
    getMaxColumns: () => maxCols,
    insertColumnsAfter(a, n) { maxCols += n; },
    clear() { g = []; return sh; },
    appendRow(row) { g.push(row.slice()); return sh; },
    setFrozenRows() {}, hideSheet() { hidden = true; },
    getDataRange() { return sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), Math.max(width(), 1)); },
    getRange(r, c, nr, nc) {
      nr = nr || 1; nc = nc || 1;
      const rng = {
        getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => { const v = (g[r - 1 + i] || [])[c - 1 + j]; return v == null ? '' : v; })),
        setValues(v) { v.forEach((row, i) => { g[r - 1 + i] = g[r - 1 + i] || []; row.forEach((x, j) => { g[r - 1 + i][c - 1 + j] = x; }); }); return rng; },
        setFontWeight: () => rng, setBackground: () => rng,
      };
      return rng;
    },
  };
  return sh;
}

function makeServer(rosterRows) {
  const props = {};
  const sheets = [];
  const ss = {
    getSheetByName: (n) => sheets.find((s) => s.getName() === n) || null,
    insertSheet: (n) => { const s = makeSheet(n, []); sheets.push(s); return s; },
    getActiveSheet: () => sheets[0] || null, setActiveSheet() {},
  };
  const box = {
    console: { log() {}, warn() {}, error: console.error },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: (k) => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = String(v); },
      getProperties: () => Object.assign({}, props), deleteProperty: (k) => { delete props[k]; } }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, waitLock() {}, releaseLock() {} }) },
    Utilities: {
      formatDate: (d) => d.toISOString().slice(0, 16).replace('T', ' ').replace(/-/g, '/'),
      base64Encode: (a) => Buffer.from(a.map((b) => b & 255)).toString('base64'),
    },
    getSS_: () => ss,
    normName_: (x) => String(x || '').normalize('NFKC').replace(/[\s　]/g, '').trim(),
  };
  vm.createContext(box);
  for (const f of ['コード.js', 'member_master_srv.js', 'spreading_srv.js']) {
    vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), box, { filename: f });
  }
  box.normName_ = (x) => String(x || '').normalize('NFKC').replace(/[\s　]/g, '').trim();
  // 名簿の行だけ読む（表紙・業種区分・役職の期の反映は、この検査では使わない）
  const real = box.getMemberMaster;
  box.getMemberMaster = () => real({ membersOnly: true });
  const HEAD = vm.runInContext('MEMBER_HEADERS_', box).slice();
  sheets.push(makeSheet('メンバー名簿', [HEAD].concat(rosterRows)));
  return { box, props, sheets, HEAD, roster: () => ss.getSheetByName('メンバー名簿') };
}

// 名簿の1行（No・業種区分・氏名・ふりがな・カテゴリー・会社名・役職・メモ・写真・一言・紹介・協業・入会日・更新日・更新期限日・会社での役職）
const row = (no, cat, name, title, company, i) => [no, cat, name, '', title, company, '', '', name.replace(/\s/g, '') + '.jpg',
  '一言' + i, '紹介' + i, '協業' + i, '2020/01/0' + i, '', '2027/01/0' + i, '肩書' + i];
const ROSTER = [
  row('1', '企業サポート', '青木 一郎', '税理士', '青木会計', 1),
  row('2', '企業サポート', '上田 花子', '社会保険労務士', 'うえだ社労士', 2),
  row('3', '企業サポート', '高橋 健太', 'ITコンサルタント', 'タカハシIT', 3),
  row('4', '研修・教育', '野村 大輔', '研修講師', 'ノムラ研修', 4),
  row('5', '不動産関連', '小林 翔太', '不動産売買仲介', '小林不動産', 5),
  row('6', '不動産関連', '加藤 真理', '賃貸管理', 'カトウ住宅', 6),
  row('7', '建築・住まい', '見本 退会', 'リフォーム', '見本工務店', 7),              // 退会した方（PDFに無い）
  // Spreadingで足した新しい方（番号はまだ無い。最後の行）
  ['', '企業サポート', '新井 誠', 'あらい まこと', '司法書士', '新井司法書士事務所', '', '', '', '', '', '', '2026/09/01', '', '', ''],
];
// PDF（公式のメンバーリスト）：新井さんが企業サポートの3番に入り、それより後ろの番号が1つずつずれている。
// 高橋さんは「髙橋」と読み取られている（異体字）。見本 退会さんは載っていない
const PDF = [
  { no: '1', name: '青木 一郎', kana: 'あおき いちろう', cat: '企業サポート', title: '税理士', company: '青木会計', role: '', memo: '' },
  { no: '2', name: '上田 花子', kana: 'うえだ はなこ', cat: '企業サポート', title: '社会保険労務士', company: 'うえだ社労士', role: '', memo: 'ビジターホスト' },
  { no: '3', name: '新井 誠', kana: 'あらい まこと', cat: '企業サポート', title: '司法書士', company: '新井司法書士事務所', role: '', memo: '' },
  { no: '4', name: '髙橋 健太', kana: 'たかはし けんた', cat: '企業サポート', title: 'ITコンサルタント', company: 'タカハシIT', role: '', memo: '' },
  { no: '5', name: '野村 大輔', kana: 'のむら だいすけ', cat: '研修・教育', title: '研修講師', company: 'ノムラ研修', role: '', memo: '' },
  { no: '6', name: '小林 翔太', kana: 'こばやし しょうた', cat: '不動産関連', title: '不動産売買仲介', company: '小林不動産', role: '', memo: '' },
  { no: '7', name: '加藤 真理', kana: 'かとう まり', cat: '不動産関連', title: '賃貸管理', company: 'カトウ住宅', role: '', memo: '' },
];
const byName = (sh, HEAD) => {
  const g = sh._grid(), o = {};
  g.slice(1).forEach((r) => { o[String(r[2])] = Object.fromEntries(HEAD.map((h, i) => [h, String(r[i] == null ? '' : r[i])])); });
  return o;
};

// ---------- 1. 読み取っただけでは名簿を変えない。確認の一覧 ----------
{
  const { box, HEAD, roster } = makeServer(ROSTER);
  const before = J(roster()._grid());
  // Gemini の返事を作り物にして、読み取りの入口（extractMembersFromPdfBlob_）から通す
  box.PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', 'test');
  box.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 200,
    getContentText: () => J({ candidates: [{ content: { parts: [{ text: J({ members: PDF.map((p) => Object.assign({ block: p.cat }, p)) }) }] } }] }) }) };
  const blob = { getBytes: () => [37, 80, 68, 70], getContentType: () => 'application/pdf' };
  const pre = box.extractMembersFromPdfBlob_(blob);
  ck(pre.ok && pre.preview && pre.extracted.length === 7, '読み取りの結果: ' + J(pre).slice(0, 200));
  ck(J(roster()._grid()) === before, '読み取っただけで名簿が変わった');
  const sm = pre.summary;
  ck(sm.adds.length === 0 && sm.near.length === 1 && sm.near[0].pdf === '髙橋 健太' && sm.near[0].roster === '高橋 健太',
     '確認の一覧（新しく足す方・1文字違い）: ' + J({ adds: sm.adds, near: sm.near }));
  ck(J(sm.missing.map((x) => x.name)) === J(['見本 退会']), 'PDFに無い方: ' + J(sm.missing));
  ck(J(sm.dupNos) === J([{ no: '7', names: ['加藤 真理', '見本 退会'] }]), '番号が重なる方: ' + J(sm.dupNos));
  const up = Object.fromEntries(sm.updates.map((u) => [u.name, u.changes.map((c) => c.field + ':' + c.from + '→' + c.to)]));
  ck(J(up['高橋 健太']) === J(['No:3→4', 'ふりがな:→たかはし けんた']) && J(up['新井 誠']) === J(['No:→3']),
     '変わる方の中身: ' + J(up));
  ck(/まだ名簿は変えていません/.test(pre.message), '確認の知らせ: ' + pre.message);

  // ---------- 2. 反映：番号がずれても、別の方の行に入らない ----------
  const res = box.applyMemberListOcr(pre.extracted);
  ck(res.ok, '反映できない: ' + J(res));
  const m = byName(roster(), HEAD);
  const names = roster()._grid().slice(1).map((r) => r[2]);
  ck(J(names) === J(['青木 一郎', '上田 花子', '新井 誠', '高橋 健太', '野村 大輔', '小林 翔太', '加藤 真理', '見本 退会']),
     '名簿の並び（番号順・氏名は名簿のまま・行が増えない）: ' + J(names));
  [['青木 一郎', 1], ['上田 花子', 2], ['高橋 健太', 3], ['野村 大輔', 4], ['小林 翔太', 5], ['加藤 真理', 6], ['見本 退会', 7]].forEach(([nm, i]) => {
    const x = m[nm] || {};
    ck(x['一言コメント'] === '一言' + i && x['紹介してほしい人'] === '紹介' + i && x['協業したい人'] === '協業' + i
       && x['写真ファイル名'] === nm.replace(/\s/g, '') + '.jpg' && x['入会日'] === '2020/01/0' + i && x['会社での役職'] === '肩書' + i,
       `${nm}：その方の写真・一言・紹介・協業・日付・会社での役職が、その方の行に残っていない: ` + J(x));
  });
  ck(J(['青木 一郎', '上田 花子', '新井 誠', '高橋 健太', '野村 大輔', '小林 翔太', '加藤 真理'].map((nm) => m[nm].No)) === J(['1', '2', '3', '4', '5', '6', '7']),
     '番号がPDFの番号になっていない: ' + J(Object.values(m).map((x) => [x['氏名'], x.No])));
  ck(m['新井 誠']['入会日'] === '2026/09/01' && m['上田 花子']['メモ'] === 'ビジターホスト' && m['高橋 健太']['ふりがな'] === 'たかはし けんた',
     'PDFの項目の反映: ' + J([m['新井 誠'], m['上田 花子']['メモ']]));
  ck(m['見本 退会'].No === '7', 'PDFに無い方の行を消した・番号を変えた');
  ck(/番号が重なる方: No7（加藤 真理・見本 退会）/.test(res.message) && /取り込む前に戻す/.test(res.message), '反映の知らせ: ' + res.message);

  // ---------- 3. 取り込む前に戻す（入れ替え。もう一度で戻す前に）----------
  const after = J(roster()._grid());
  const info = box.getMemberBackupInfo_();
  ck(info && info.label === 'メンバーリスト(OCR)の取り込み' && info.rows === 8, '控えの知らせ: ' + J(info));
  const bk = box.getSS_().getSheetByName('メンバー名簿_取り込み前');
  ck(bk && bk.isHidden() && J(bk._grid()) === before, '控えのシート（隠す・取り込む前の中身）');
  let r = box.restoreMemberBackup();
  ck(r.ok && J(roster()._grid()) === before && J(bk._grid()) === after, '取り込む前に戻す: ' + J(r));
  r = box.restoreMemberBackup();
  ck(r.ok && J(roster()._grid()) === after, 'もう一度押すと、戻す前の名簿に戻る: ' + J(r));
}

// ---------- 4. 1文字違いの扱い：2文字の氏名・当てはまる方が2人のときは、新しい方として足す ----------
{
  const { box, HEAD, roster } = makeServer([
    row('1', '企業サポート', '林 彩', 'Web制作', 'はやし工房', 1),
    row('2', '企業サポート', '山田 太郎', '税理士', '山田会計', 2),
    row('3', '企業サポート', '山田 次郎', '弁護士', '山田法律', 3),
  ]);
  const pdf = [
    { no: '1', name: '林 誠', kana: '', cat: '企業サポート', title: 'カメラマン', company: 'はやし写真', role: '', memo: '' },
    { no: '2', name: '山田 三郎', kana: '', cat: '企業サポート', title: '司法書士', company: '山田司法書士', role: '', memo: '' },
    { no: '', name: '', kana: '', cat: '', title: '', company: '', role: '', memo: '' },
    { no: '9', name: '', kana: '', cat: '', title: '', company: '', role: '', memo: '' },
  ];
  const sm = box.ocrPlanSummary_(pdf, box.getMemberMaster().members);
  ck(J(sm.adds.map((x) => x.name)) === J(['林 誠', '山田 三郎']) && sm.near.length === 0, '2文字・2人に当てはまるときは足す: ' + J(sm));
  ck(J(sm.noName) === J([{ no: '' }, { no: '9' }]), '氏名を読み取れなかった行: ' + J(sm.noName));
  box.applyMemberListOcr(pdf);
  const m = byName(roster(), HEAD);
  ck(m['林 彩']['カテゴリー'] === 'Web制作' && m['山田 太郎']['カテゴリー'] === '税理士' && m['林 誠'] && m['山田 三郎'] && roster()._grid().length === 6,
     '別の方を書き換えた: ' + J(Object.keys(m)));
}

// ---------- 5. ダイアログを使わない取り込みも、確かめてから（キャンセルなら名簿は変わらない）----------
{
  const { box, roster } = makeServer(ROSTER);
  const alerts = [];
  let answer = 'CANCEL';
  box.SpreadsheetApp = { getUi: () => ({
    ButtonSet: { OK: 'OK', OK_CANCEL: 'OK_CANCEL' }, Button: { OK: 'OK', CANCEL: 'CANCEL' },
    prompt: () => ({ getSelectedButton: () => 'OK', getResponseText: () => 'https://drive.google.com/file/d/ABCDEFGHIJKLMNOPQRSTUVWXYZ012345/view' }),
    alert: (title, msg, btns) => { alerts.push(msg); return btns === 'OK_CANCEL' ? answer : 'OK'; },
  }) };
  box.processMemberListFromDrive = () => box.extractMembersFromPdfBlob_({ getBytes: () => [1], getContentType: () => 'application/pdf' });
  box.PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', 'test');
  box.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 200,
    getContentText: () => J({ candidates: [{ content: { parts: [{ text: J({ members: PDF.map((p) => Object.assign({ block: p.cat }, p)) }) }] } }] }) }) };
  const before = J(roster()._grid());
  box.importMemberListSimple();
  ck(J(roster()._grid()) === before && alerts.some((a) => /反映しますか/.test(a)) && /変えていません/.test(alerts[alerts.length - 1]),
     'ダイアログを使わない取り込みで、キャンセルしたのに名簿が変わった: ' + J(alerts));
  answer = 'OK';
  box.importMemberListSimple();
  ck(J(roster()._grid()) !== before && /反映しました/.test(alerts[alerts.length - 1]), 'ダイアログを使わない取り込みで、OKしても反映しない: ' + alerts[alerts.length - 1]);
}

// ---------- 6. Spreadingの取り込みも、前の名簿を控える ----------
{
  const { box } = makeServer(ROSTER);
  const json = J({ ok: true, data: { items: [{ name: '青木 一郎', furigana: 'あおき', category: '税理士', company_name: '青木会計', position: '', date_join: '',
    status: 0, category_groups: [{ category_name: '企業サポート', sort: 1, member_sort: 1 }] }] } });
  const r = box.importSpreadingMembers(json, false);
  const info = box.getMemberBackupInfo_();
  ck(r.ok && info && info.label === 'Spreadingからの取り込み' && /取り込む前に戻す/.test(r.message), 'Spreadingの取り込みの控え: ' + J([r.ok, info]));
}

// ---------- 7. 画面（Chromium）：読み取りの確認の一覧 →「名簿に反映」／「やめる」、名簿の画面の「取り込む前に戻す」----------
const pw = (() => { try { return require('playwright'); } catch (e) { return require('/opt/node22/lib/node_modules/playwright'); } })();
// google.script.run の代わり。answers … 関数名 → 返す値。呼ばれた関数と引数は window.__calls に残す
const STUB = (answers) => `<script>window.__calls=[];window.__ans=${J(answers).replace(/</g, '\\u003c')};
  window.google={script:{get run(){var ok=null,p=new Proxy({},{get:function(_,n){
    if(n==='withSuccessHandler')return function(f){ok=f;return p;};
    if(n==='withFailureHandler'||n==='withUserObject')return function(){return p;};
    return function(){var a=[].slice.call(arguments);window.__calls.push({fn:n,args:JSON.parse(JSON.stringify(a))});
      setTimeout(function(){ok&&ok(JSON.parse(JSON.stringify(window.__ans[n]||{ok:true,message:'OK'})));},10);};}});return p;}}};</script>`;
(async () => {
  const { box } = makeServer(ROSTER);
  box.PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', 'test');
  box.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 200,
    getContentText: () => J({ candidates: [{ content: { parts: [{ text: J({ members: PDF.map((p) => Object.assign({ block: p.cat }, p)) }) }] } }] }) }) };
  const pre = box.extractMembersFromPdfBlob_({ getBytes: () => [1], getContentType: () => 'application/pdf' });
  const browser = await pw.chromium.launch();
  const page = await browser.newPage();
  const dialogs = [];
  page.on('dialog', (d) => { dialogs.push(d.message()); d.accept(); });
  const serve = async (file, answers) => {
    const html = fs.readFileSync(path.join(ROOT, file), 'utf8').replace(/<head>/i, '<head><meta charset="utf-8">' + STUB(answers));
    await page.route('https://ocr.test/' + file, (r) => r.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: html }));
    await page.goto('https://ocr.test/' + file, { waitUntil: 'load' });
  };
  // OCRの画面：読み取ったら一覧が出るだけ（反映しない）。「名簿に反映」で読み取った内容を送る
  await serve('pdf.html', { processMemberListFromDrive: pre, applyMemberListOcr: { ok: true, message: '「メンバー名簿」に反映しました（名簿は計 8名）。' } });
  await page.fill('#driveLink', 'https://drive.google.com/file/d/ABCDEFGHIJKLMNOPQRSTUVWXYZ012345/view');
  await page.click('#btn2');
  await page.waitForSelector('#plan', { state: 'visible' });
  const planText = await page.textContent('#planBody');
  ck(/氏名が1文字違う方/.test(planText) && /髙橋 健太/.test(planText) && /PDFに無い方/.test(planText) && /見本 退会/.test(planText)
     && /番号が重なる方/.test(planText) && /No:? ?3 → 4|No 3 → 4/.test(planText), '確認の一覧の中身: ' + planText.slice(0, 300));
  ck(!(await page.evaluate(() => window.__calls.some((c) => c.fn === 'applyMemberListOcr'))), '読み取っただけで反映した');
  await page.click('#btnApply');
  await page.waitForFunction(() => /反映しました/.test(document.getElementById('msg').textContent));
  const call = await page.evaluate(() => window.__calls.find((c) => c.fn === 'applyMemberListOcr'));
  ck(call && call.args[0].length === 7 && call.args[0][2].name === '新井 誠', '「名簿に反映」で送った内容: ' + J(call && call.args[0].length));
  ck(!(await page.isVisible('#plan')), '反映したあとも一覧が出ている');
  // やめる：送らない
  await serve('pdf.html', { processMemberListFromDrive: pre });
  await page.fill('#driveLink', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ012345');
  await page.click('#btn2');
  await page.waitForSelector('#plan', { state: 'visible' });
  await page.click('button:has-text("やめる")');
  ck(!(await page.isVisible('#plan')) && /名簿は変えていません/.test(await page.textContent('#msg'))
     && !(await page.evaluate(() => window.__calls.some((c) => c.fn === 'applyMemberListOcr'))), '「やめる」で送った・一覧が残った');
  // 名簿の画面：控えがあれば「取り込む前に戻す」が出る。押すと確かめてから戻す
  const members = box.getMemberMaster().members;
  await serve('member_master.html', { getMemberMaster: { ok: true, members, categories: [],
    backup: { at: '2026/09/28 10:15', label: 'メンバーリスト(OCR)の取り込み', rows: 8 } }, restoreMemberBackup: { ok: true, message: '戻しました' } });
  await page.evaluate(() => window.onload && window.onload());
  await page.waitForSelector('#backup', { state: 'visible' });
  ck(/2026\/09\/28 10:15（メンバーリスト\(OCR\)の取り込みの前・8名）/.test(await page.textContent('#backupText')), '控えの知らせ: ' + await page.textContent('#backupText'));
  const n0 = dialogs.length;
  await page.click('button:has-text("取り込む前に戻す")');
  await page.waitForFunction(() => window.__calls.some((c) => c.fn === 'restoreMemberBackup'));
  ck(dialogs.length === n0 + 1 && /入れ替えます/.test(dialogs[dialogs.length - 1]), '戻す前に確かめていない: ' + J(dialogs.slice(n0)));
  // 控えが無ければ出ない
  await serve('member_master.html', { getMemberMaster: { ok: true, members, categories: [], backup: null } });
  await page.evaluate(() => window.onload && window.onload());
  await page.waitForTimeout(200);
  ck(!(await page.isVisible('#backup')), '控えが無いのに「取り込む前に戻す」が出る');
  await browser.close();

  console.log(`メンバーリスト(OCR)の取り込み: 検査 ${checks} 件`);
  if (fails.length) {
    console.log(`NG: ${fails.length} 件`);
    fails.forEach((f) => console.log('   ' + f));
    process.exit(1);
  }
  console.log('OK: 氏名で突き合わせ（番号がずれても別の方に入らない・1文字違い・新しい方・PDFに無い方・番号の重なり）・確かめてから反映（画面も）・取り込む前に戻す（画面も）');
})().catch((e) => { console.error(e); process.exit(1); });

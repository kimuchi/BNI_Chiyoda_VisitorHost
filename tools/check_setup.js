// 初回の準備（setup_srv.js）を、偽のスプレッドシートで確かめる。
//
//   node tools/check_setup.js <routine.json> <members.json>
//
// 確かめること
//   ・空のスプレッドシートで開くと（onOpen）、メンバー名簿・業種区分マスタ（非表示）・休会日（非表示）を作り、
//     最初からある空のシートを消して、「初回の準備」の記録（fresh）を残す。2回目からは何もしない
//   ・チャプターの設定を保存するまでは、ルーティンチェックシートを作らない（期と開催日が決まらないため）。
//     メニュー画面は「チャプターの設定」を促す
//   ・チャプターの設定を保存すると、次の定例会の期のルーティンチェックシートを作る
//     （開催日はその期の定例会の曜日の毎週・回数は休会日を除いて数える・曜日目安は定例会の曜日から）。
//     作ったシートを、スライド（getRoutineInfo）・役職ごとの入力（読み込みと保存）・トークスクリプトが読める
//   ・期の番号を付け直すと、作ったシートの名前の期もずれる
//   ・今まで使ってきたスプレッドシート（existing）では、開いても何も作らない。
//     「足りないシートを作る」では、次の定例会に列の無い期のシートを、前の期のシートを写して作る
//     （項目・担当・期日などはそのまま、開催日の列だけ新しく。1列おきの表も1列おきのまま）

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeRoleServer } = require('./lib_role_fixture');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }
const J = (x) => JSON.stringify(x);

const S = makeRoleServer(process.argv[2], process.argv[3]);
const F = S.F, props = S.props;
for (const f of ['home_srv.js', 'talk_script_default.js', 'talk_script_srv.js']) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), S.sandbox, { filename: f });
}
const ROUTINE = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const resetAll = () => { vm.runInContext('CHAPTER_CACHE_ = null;', S.sandbox); F.routineResetCache_(); };

// ===== 偽のスプレッドシート（シートを足す・消す・写す・名前を変える・列を足す・書式）=====
function makeBook(initial) {
  const book = { sheets: [], toasts: [], alerts: [], active: null };
  let nextId = 1;
  function makeSheet(name, grid, opts) {
    const g = (grid || []).map((r) => (r || []).slice());
    let maxRows = Math.max(1000, g.length), maxCols = Math.max(26, ...g.map((r) => r.length), 0);
    const id = nextId++;
    const cellsOk = (r, c, nr, nc) => {
      if (r < 1 || c < 1 || r + nr - 1 > maxRows || c + nc - 1 > maxCols) {
        throw new Error('範囲の座標がシートの範囲外です（' + name + ' ' + r + ',' + c + ' ' + nr + '×' + nc + ' / ' + maxRows + '×' + maxCols + '）');
      }
    };
    const put = (r, c, v) => {
      while (g.length < r) g.push([]);
      const row = g[r - 1];
      while (row.length < c) row.push('');
      row[c - 1] = v;
    };
    const sh = {
      _grid: g, _fmt: [], _hidden: !!(opts && opts.hidden), _frozen: {}, _widths: {},
      getName: () => name,
      setName: (n) => { if (book.sheets.some((s) => s !== sh && s.getName() === n)) throw new Error('同じ名前のシートがあります: ' + n); name = n; return sh; },
      getSheetId: () => id,
      getIndex: () => book.sheets.indexOf(sh) + 1,
      getLastRow: () => { for (let r = g.length - 1; r >= 0; r--) if ((g[r] || []).some((v) => v !== '' && v != null)) return r + 1; return 0; },
      getLastColumn: () => {
        let w = 0;
        g.forEach((row) => { for (let c = (row || []).length - 1; c >= w; c--) if (row[c] !== '' && row[c] != null) { w = c + 1; break; } });
        return w;
      },
      getMaxRows: () => maxRows,
      getMaxColumns: () => maxCols,
      insertColumnsAfter: (after, n) => { maxCols += n; g.forEach((row) => { if (row.length > after) row.splice(after, 0, ...new Array(n).fill('')); }); },
      hideSheet: () => { sh._hidden = true; return sh; },
      isSheetHidden: () => sh._hidden,
      appendRow: (row) => { const r = sh.getLastRow() + 1; row.forEach((v, j) => put(r, j + 1, v)); return sh; },
      getDataRange: () => sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), Math.max(sh.getLastColumn(), 1)),
      setColumnWidth: (c, w) => { sh._widths[c] = w; return sh; },
      setFrozenRows: (n) => { sh._frozen.rows = n; return sh; },
      setFrozenColumns: (n) => { sh._frozen.cols = n; return sh; },
      clear: () => { g.length = 0; return sh; },
      activate: () => { book.active = sh; return sh; },
      copyTo: () => {
        const c = makeSheet('Copy of ' + name, g, {});
        c._widths = Object.assign({}, sh._widths);
        book.sheets.push(c);
        return c;
      },
      getRange(r, c, nr, nc) {
        nr = nr || 1; nc = nc || 1;
        cellsOk(r, c, nr, nc);
        const fmt = (m) => (v) => { sh._fmt.push([m, r, c, nr, nc, v]); return rg; };
        const rg = {
          getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => {
            const v = (g[r - 1 + i] || [])[c - 1 + j];
            return v == null ? '' : v;
          })),
          getFormulas: () => Array.from({ length: nr }, () => new Array(nc).fill('')),
          setValues(vals) {
            if (vals.length !== nr || vals.some((row) => row.length !== nc)) throw new Error('データの行数・列数が範囲と合いません');
            vals.forEach((row, i) => row.forEach((v, j) => put(r + i, c + j, v)));
            return rg;
          },
          setValue(v) { return rg.setValues([[v]]); },
          clearContent() {
            for (let i = 0; i < nr; i++) {
              const row = g[r - 1 + i];
              if (!row) continue;
              for (let j = 0; j < nc; j++) if (c - 1 + j < row.length) row[c - 1 + j] = '';
            }
            return rg;
          },
          setNumberFormat: fmt('numberFormat'), setFontWeight: fmt('fontWeight'), setBackground: fmt('background'),
          setHorizontalAlignment: fmt('hAlign'), setVerticalAlignment: fmt('vAlign'), setWrap: fmt('wrap'),
          setFontColor: fmt('fontColor'), setFontSize: fmt('fontSize'),
        };
        return rg;
      },
    };
    return sh;
  }
  Object.assign(book, {
    makeSheet,
    getSheets: () => book.sheets.slice(),
    getSheetByName: (n) => book.sheets.find((s) => s.getName() === n) || null,
    insertSheet: (n, index) => {
      if (book.getSheetByName(n)) throw new Error('同じ名前のシートがあります: ' + n);
      const sh = makeSheet(n, [], {});
      if (index == null) book.sheets.push(sh); else book.sheets.splice(index, 0, sh);
      return sh;
    },
    deleteSheet: (sh) => {
      if (book.sheets.filter((s) => s !== sh && !s._hidden).length === 0) throw new Error('すべてのシートを削除することはできません');
      book.sheets.splice(book.sheets.indexOf(sh), 1);
    },
    setActiveSheet: (sh) => { book.active = sh; return sh; },
    moveActiveSheet: (pos) => {
      const sh = book.active, i = book.sheets.indexOf(sh);
      book.sheets.splice(i, 1);
      book.sheets.splice(pos - 1, 0, sh);
    },
    toast: (msg, title) => { book.toasts.push({ msg, title }); },
    getUrl: () => 'https://docs.google.com/spreadsheets/d/TEST/edit',
    getId: () => 'TEST',
  });
  (initial || []).forEach((s) => book.sheets.push(makeSheet(s.name, s.grid, s)));
  return book;
}
// メニュー（onOpen）と alert
const menus = [];
function makeUi() {
  const mk = (name) => {
    const m = { name, items: [] };
    m.addItem = (label, fn) => { m.items.push({ label, fn }); return m; };
    m.addSeparator = () => m;
    m.addSubMenu = (sub) => { m.items.push({ sub }); return m; };
    m.addToUi = () => { menus.push(m); return m; };
    return m;
  };
  return { createMenu: mk, alert: (t, msg) => { book.alerts.push({ t, msg }); return 'OK'; }, ButtonSet: { OK: 'OK' } };
}
let book = null;
S.sandbox.SpreadsheetApp = { getActiveSpreadsheet: () => book, getUi: () => makeUi() };
F.getSS_ = () => book;
F.requireSheetAccess_ = () => {};   // 設定を読み書きできる方か（本番は webapp_srv.js。ここでは読み込まない）
const names = () => book.getSheets().map((s) => s.getName());
const valuesOf = (sh) => sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), Math.max(sh.getLastColumn(), 1)).getValues();
// 期日から曜日目安（検査の側でも数える）
const WEEK = ['日', '月', '火', '水', '木', '金', '土'];

// ===== 1. 空のスプレッドシートで開く =====
for (const k of Object.keys(props)) delete props[k];
resetAll();
book = makeBook([{ name: 'シート1', grid: [] }]);
F.onOpen();
const allItems = (m) => m.items.flatMap((it) => (it.sub ? allItems(it.sub) : [it]));
const menuItems = menus.length ? allItems(menus[menus.length - 1]) : [];
ck(menuItems.some((it) => it.fn === 'menuCreateMissingSheets' && /足りないシートを作る/.test(it.label)), 'メニューに「足りないシートを作る」が無い');
ck(J(names().slice().sort()) === J(['メンバー名簿', '休会日', '業種区分マスタ'].sort()), '初回に作ったシート: ' + J(names()));
const setup1 = JSON.parse(props.BNI_SETUP || 'null');
ck(setup1 && setup1.mode === 'fresh' && setup1.made.length === 3, '初回の準備の記録: ' + props.BNI_SETUP);
{
  const mem = book.getSheetByName('メンバー名簿'), cat = book.getSheetByName('業種区分マスタ'), hol = book.getSheetByName('休会日');
  ck(mem && !mem._hidden && J(valuesOf(mem)[0]) === J(vm.runInContext('MEMBER_HEADERS_', S.sandbox)), 'メンバー名簿の見出し: ' + (mem && J(valuesOf(mem)[0])));
  ck(cat && cat._hidden && cat.getLastRow() === 1 + vm.runInContext('DEFAULT_CATEGORIES_', S.sandbox).length, '業種区分マスタ（非表示・既定の区分）');
  ck(hol && hol._hidden && hol.getLastRow() === 0, '休会日（非表示・空）');
}
ck(book.toasts.length === 1 && /チャプター/.test(book.toasts[0].msg) && /初回の準備/.test(book.toasts[0].title), 'お知らせ（toast）: ' + J(book.toasts));
F.onOpen();
ck(names().length === 3 && book.toasts.length === 1, '2回目に開いても何もしない: ' + J(names()));
ck(F.chapterFresh_() === true, '新しく始めたスプレッドシートとして扱われない');

// チャプターの設定の前：ルーティンチェックシートは作らない。メニュー画面は設定を促す
let home = F.getHomeStatus(), row = (k) => (home.checks || []).find((c) => c.key === k);
ck(row('chapter') && !row('chapter').ready, 'チャプターの設定を促さない（初回）');
ck(row('routine') && !row('routine').ready && row('routine').fixFn === 'openChapterSettingsDialog' && /チャプターの設定/.test(row('routine').detail),
   'チェックシート（設定の前）: ' + J(row('routine')));
let res = F.createMissingSheets();
ck(res.ok && res.made.length === 0 && /足りないシートはありません/.test(res.message) && /チャプターの設定/.test(res.message)
   && !names().some((n) => /ルーティンチェックシート/.test(n)), '設定の前に「足りないシートを作る」: ' + J(res));

// ===== 2. チャプターの設定を保存すると、ルーティンチェックシートを作る（火曜開催・12期、10/6 が第41回）=====
book.getSheetByName('休会日').getRange(1, 1, 2, 1).setValues([['2026/12/29'], ['2027/03/16']]);
res = F.saveChapterSettings({ name: 'サンプル', region: 'BNI東京サンプルリージョン', term: '12', meetingBaseDate: '2026/10/06', meetingBaseCount: '41' });
const R13 = '【13期】ルーティンチェックシート';
ck(res.ok && /ルーティンチェックシートを作りました/.test(res.message) && res.message.includes(R13), '設定を保存したときのお知らせ: ' + (res.message || ''));
let sh = book.getSheetByName(R13);
ck(sh && sh.getIndex() === 1 && !names().includes('【12期】ルーティンチェックシート'), '作ったシート（次の定例会の期だけ・いちばん前）: ' + J(names()));
let grid = sh ? valuesOf(sh) : [[]];
{
  // 開催日：2026/10/1〜2027/3/31 の火曜。回数は 10/6 を41回として、休会日の週を除いて数える
  const want = [], hol = ['2026/12/29', '2027/03/16'];
  const fmt = (d) => d.getFullYear() + '/' + String(d.getMonth() + 1).padStart(2, '0') + '/' + String(d.getDate()).padStart(2, '0');
  let n = 41;
  for (let d = new Date(2026, 9, 6); d <= new Date(2027, 2, 31); d.setDate(d.getDate() + 7)) {
    const k = fmt(d);
    want.push([k, hol.includes(k) ? '' : n]);
    if (!hol.includes(k)) n++;
  }
  const got = grid[0].slice(9).map((v, i) => [v, grid[1][9 + i]]).filter((x) => x[0] !== '');
  ck(J(got) === J(want), '開催日と回数: ' + J(got.slice(10, 14)) + ' / ' + J(want.slice(10, 14)) + '（' + got.length + '/' + want.length + '）');
  ck(grid[0][1] === '定例会開催日' && grid[1][1] === '定例会回数' && grid[2][1] === '担当割り振り'
     && J(grid[3].slice(1, 9)) === J(['No.', '内容', '', '', '担当', '期日', '曜日目安', '備考']), '1〜4行目: ' + J(grid[3]));
  // 曜日目安（火曜開催）
  const byLabel = (l) => grid.find((r) => [r[2], r[3], r[4]].some((v) => String(v).replace(/\s/g, '') === l)) || [];
  const dueOf = (l) => byLabel(l).slice(6, 8).join('／');
  ck(dueOf('スピーカーローテーション用スライド画像') === '5日前／木曜まで' && dueOf('メンバーシップから報告') === '2日前／日曜まで'
     && dueOf('リージョン参加者') === '3日前／土曜まで' && dueOf('割振表') === '1日前／月曜' && dueOf('真正度確認') === '当日／火曜'
     && dueOf('スタートアッププレゼン') === '／', '期日と曜日目安（火曜開催）: ' + ['スピーカーローテーション用スライド画像', 'メンバーシップから報告', '真正度確認'].map(dueOf).join(' | '));
  const hs = sh._fmt.filter((f) => f[0] === 'background' && f[5] === '#e0e0e0').map((f) => grid[0][f[2] - 1]);
  ck(J(hs) === J(hol), '休会日の列に色: ' + J(hs));
  ck(sh._frozen.rows === 4 && sh._widths[10] === 150, '書式（見出しの固定・列の幅）: ' + J(sh._frozen));
}
// 読む側：スライド・役職ごとの入力・トークスクリプト
{
  const labels = ['定例会回数', 'BNI目的と概要', 'スタートアッププレゼン', 'メインプレゼン', '募集カテゴリー', '開放カテゴリー', '審査中カテゴリー',
                  '一般規定', '推薦の言葉', 'リージョン参加者', 'ウィークリープレゼン', '真正度確認', '更新対象者30日前', '更新対象者60日前',
                  '更新対象者90日前', '担当割り振り', '代理', '欠席', '医療欠席', '体験談', 'エデュケーション', '新入会'];
  const lost = labels.filter((l) => F.routineFindRow_(grid, [l]) < 0);
  ck(!lost.length, 'スライドなどが読む項目が無い: ' + J(lost));
  const r = F.routineFindRow_(grid, ['メインプレゼン']);
  ck(r >= 0 && grid[r][3] === 'メインプレゼン', '「メインプレゼン」が画像の行に当たる: ' + J(grid[r]));
  const DEF = vm.runInContext('TALK_DEFAULT_ROWS_', S.sandbox);
  const keys = [...new Set(DEF.flat().join('\n').match(/\{ルーティン[:：][^}]+\}/g) || [])].map((k) => k.replace(/^\{ルーティン[:：]\s*|\}$/g, ''));
  const noRow = keys.filter((k) => F.routineFindRow_(grid, [k.replace(/[\s　]/g, '').replace(/[※＊].*$/, '')]) < 0);
  ck(keys.length > 10 && !noRow.length, 'トークスクリプトの既定のひな形が読む項目が無い: ' + J(noRow));
}
resetAll();
let info = F.getRoutineInfo('2026/10/06');
ck(info.ok && info.found && info.sheetName === R13 && info.meetingNo === '41', 'スライドが読む（第41回）: ' + J(info).slice(0, 160));
let ctx = F.getRoleInputContext('2026/10/06', 'vice');
const vice = (ctx.roles || []).find((x) => x.key === 'vice') || { items: [] };
ck(ctx.ok && ctx.found && vice.items.length >= 10 && vice.status.required >= 8, '役職ごとの入力が読む（バイスプレジデント）: ' + J({ found: ctx.found, n: vice.items.length, st: vice.status }));
{
  const id = vice.items.find((i) => ctx.items[i].title === 'メンバーシップから報告');
  const sv = id ? F.saveRoleInput('2026/10/06', 'vice', [{ id, value: '見本の報告', orig: '' }]) : { ok: false };
  const g2 = valuesOf(sh), rr = F.routineFindRow_(g2, ['メンバーシップから報告']);
  ck(sv.ok && g2[rr][9] === '見本の報告', '役職ごとの入力を保存（10/6 の列）: ' + J(sv).slice(0, 120) + ' / ' + (g2[rr] || [])[9]);
}
{
  // スタートアッププレゼンの行（新しく作ったシートの呼び名）に書いた方を、スライド・トークスクリプトが読む
  const who = F.getMemberMaster().members[0].name, rr = F.routineFindRow_(valuesOf(sh), ['スタートアッププレゼン']);
  sh.getRange(rr + 1, 10, 1, 1).setValues([[who.replace(/[\s　].*$/, '') + 'さん']]);
  resetAll();
  const ri = F.getRoutineInfo('2026/10/06');
  ck(ri.found && ri.longPresenter === who, 'スタートアッププレゼンの方を読む: ' + J({ raw: ri.longPresenterRaw, name: ri.longPresenter }));
}
{
  const pv = F.previewTalkScript('2026/10/06');
  ck(pv.ok && pv.unknown.length === 0 && /第41回/.test(pv.label) && pv.rows.some((x) => /サンプルチャプター/.test(x.talk)),
     'トークスクリプト（作ったシートで）: ' + J({ ok: pv.ok, label: pv.label, unknown: pv.unknown }));
}
home = F.getHomeStatus();
ck(row('routine') && row('routine').ready && row('routine').detail === R13 && row('chapter').ready, 'メニュー画面（設定のあと）: ' + J(row('routine')));

// ===== 3. 期の番号を付け直すと、作ったシートの名前もずれる =====
res = F.saveChapterSettings({ name: 'サンプル', region: 'BNI東京サンプルリージョン', term: '13', meetingBaseDate: '2026/10/06', meetingBaseCount: '41' });
ck(res.ok && names().includes('【14期】ルーティンチェックシート') && !names().includes(R13) && /名前の期もずらしました/.test(res.message)
   && !/作りました:/.test(res.message), '期を付け直す（13→14）: ' + J(names()) + ' ' + res.message);
res = F.saveChapterSettings({ name: 'サンプル', region: 'BNI東京サンプルリージョン', term: '12', meetingBaseDate: '2026/10/06', meetingBaseCount: '41' });
ck(res.ok && names().includes(R13) && !names().includes('【14期】ルーティンチェックシート'), '期を戻す: ' + J(names()));

// ===== 4. 消えたシートは「足りないシートを作る」で作り直せる =====
book.deleteSheet(book.getSheetByName(R13));
resetAll();
home = F.getHomeStatus();
ck(row('routine') && !row('routine').ready && row('routine').fixFn === 'menuCreateMissingSheets' && /列がありません/.test(row('routine').detail),
   'メニュー画面（シートが無いとき）: ' + J(row('routine')));
res = F.menuCreateMissingSheets();
ck(res.ok && J(res.made) === J([R13]) && book.alerts.length === 1 && /26回ぶんの開催日/.test(book.alerts[0].msg), '足りないシートを作る（メニュー）: ' + J(res));
res = F.createMissingSheets();
ck(res.ok && res.made.length === 0 && /足りないシートはありません/.test(res.message) && !/チャプターの設定/.test(res.message), 'もう一度: ' + res.message);

// ===== 5. 今まで使ってきたスプレッドシート =====
for (const k of Object.keys(props)) delete props[k];
const holGrid = S.sheets.find((s) => s.getName() === '休会日')._grid;
const prevTerms = Object.keys(ROUTINE).filter((n) => !/【24期】/.test(n));      // 24期のシートを、まだ作っていないつもり
book = makeBook(prevTerms.map((n) => ({ name: n, grid: ROUTINE[n] })).concat([
  { name: 'メンバー名簿', grid: [vm.runInContext('MEMBER_HEADERS_', S.sandbox)] },
  { name: '業種区分マスタ', grid: [['キー']], hidden: true }, { name: '休会日', grid: holGrid, hidden: true },
  { name: '20260930参加者', grid: [['No.', '参加者氏名']] }]));
resetAll();
const before = names();
F.onOpen();
ck(JSON.parse(props.BNI_SETUP || 'null').mode === 'existing' && J(names()) === J(before) && book.toasts.length === 0, '今までのスプレッドシートで開く: ' + props.BNI_SETUP);
ck(F.chapterFresh_() === false, '今までのスプレッドシートを「新しく始めた」と扱う');
res = F.createMissingSheets();
const R23 = '【23期】ルーティンチェックシート', R24 = '【24期】ルーティンチェックシート';
sh = book.getSheetByName(R24);
ck(res.ok && J(res.made) === J([R24]) && res.message.includes('「' + R23 + '」から写しました'), '次の期のシートを作る: ' + J(res));
ck(sh && sh.getIndex() === book.getSheetByName(R23).getIndex() - 1, '写したシートの置き場所（前の期の前）: ' + J(names()));
if (sh) {
  const src = ROUTINE[R23], g = valuesOf(sh);
  const labelsSame = src.every((r, i) => J((r || []).slice(0, 9).map((v) => (v == null ? '' : v))) === J(((g[i] || []).slice(0, 9)).concat(new Array(9).fill('')).slice(0, 9)));
  ck(labelsSame, '項目・担当・期日・備考（A〜I列）がそのまま写っていない');
  const dates = [], counts = [];
  g[0].forEach((v, c) => { if (/^\d{4}\/\d{2}\/\d{2}$/.test(String(v))) { dates.push([c, v]); counts.push(g[1][c]); } });
  ck(dates.length === 26 && dates[0][0] === 9 && dates[1][0] === 11 && dates[0][1] === '2026/10/07' && dates[25][1] === '2027/03/31',
     '開催日（J列から1列おき・10/7〜3/31の水曜）: ' + J(dates.slice(0, 3)) + ' … ' + J(dates.slice(-1)));
  const wantCounts = dates.map(([, v]) => F.meetingCountOf_(new Date(v + ' 00:00:00')) || '');
  ck(J(counts) === J(wantCounts) && counts[0] === 536, '回数（第536回から）: ' + J(counts.slice(0, 4)));
  const dirty = [];
  g.forEach((r, i) => { if (i >= 2) r.slice(9).forEach((v, j) => { if (v !== '' && v != null) dirty.push([i + 1, j + 10, v]); }); });
  ck(!dirty.length, '開催日の列に前の期の中身が残っている: ' + J(dirty.slice(0, 3)));
}
resetAll();
info = F.getRoutineInfo('2026/10/07');
ck(info.ok && info.found && info.sheetName === R24 && info.meetingNo === '536', '写したシートをスライドが読む: ' + J(info).slice(0, 140));
res = F.createMissingSheets();
ck(res.ok && res.made.length === 0, 'もう一度作ろうとしても作らない: ' + J(res.made));

// ===== 6. 回数の数え方（基準より前の日・休会日・曜日の違う日）=====
{
  props.BNI_CHAPTER = J({ name: 'サンプル', region: '', termBase: 12, meetingBaseDate: '2026/10/06', meetingBaseCount: 41 });
  resetAll();
  const c = (y, m, d, hol) => F.setupCountOf_(new Date(y, m - 1, d), hol || []);
  ck(c(2026, 10, 6) === 41 && c(2026, 10, 20) === 43 && c(2026, 9, 29) === 40 && c(2026, 4, 7) === 15, '回数（前後）: ' + [c(2026, 10, 20), c(2026, 9, 29), c(2026, 4, 7)]);
  ck(c(2026, 10, 20, ['2026/10/13']) === 42 && c(2026, 9, 29, ['2026/09/22']) === 40 && c(2026, 9, 15, ['2026/09/22']) === 39
     && c(2026, 10, 13, ['2026/10/13']) === '', '回数（休会日）: ' + [c(2026, 10, 20, ['2026/10/13']), c(2026, 9, 15, ['2026/09/22'])]);
  ck(c(2026, 10, 7) === '' && c(2025, 1, 7) === '', '曜日の違う日・1回より前は空: ' + [c(2026, 10, 7), c(2025, 1, 7)]);
  ck(F.setupWeekdayText_(5, 3) === '金曜まで' && F.setupWeekdayText_(2, 3) === '月曜まで' && F.setupWeekdayText_(1, 3) === '火曜'
     && F.setupWeekdayText_(0, 3) === '水曜' && F.setupWeekdayText_(3, 0) === '木曜まで' && F.setupWeekdayText_('', 3) === '',
     '曜日目安（水曜・日曜開催）');
  const tr = F.setupTermRange_(13);
  ck(tr.from.getFullYear() === 2026 && tr.from.getMonth() === 9 && tr.to.getFullYear() === 2027 && tr.to.getMonth() === 2 && tr.to.getDate() === 31,
     '期の範囲（13期＝2026年10月〜2027年3月）');
  void WEEK;
}

console.log(`初回の準備: 検査 ${checks} 件`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: 空のスプレッドシート（開いたとき・設定のあと・期の付け直し・作り直し）・今までのスプレッドシート（何もしない・前の期を写す）・回数と曜日目安');

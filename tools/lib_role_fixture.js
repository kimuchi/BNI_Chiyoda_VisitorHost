// 「役職ごとの入力」の検査で使う準備：ルーティンチェックシートの写しを書き込めるシートにし、
// 本番と同じファイル（role_input_srv.js ほか）をNodeで動かす。
// check_role_input.js（サーバー）と check_role_input_dialog.js（画面）で使う。
//
//   const { makeRoleServer } = require('./lib_role_fixture');
//   const { F, sheets, members, SUB_FOR } = makeRoleServer(routineJsonPath, membersJsonPath);
//
// 参加者シート（20260930参加者）は、架空のビジター・ゲストと、名簿の方の代理で作る。
// 「今日」は 2026/09/26（土）に固定する（次回の定例会＝9/30）。

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { readZip, writeZip } = require('./lib_zip');

const ROOT = path.join(__dirname, '..');

// GAS の Blob を最小限だけ真似る（中身は Buffer）
function makeBlob(buf, type, name) {
  let nm = name || '', ct = type || '';
  return {
    _buf: buf,
    getBytes: () => Array.from(buf).map((b) => (b > 127 ? b - 256 : b)),
    getDataAsString: () => buf.toString('utf8'),
    getContentType: () => ct,
    setContentType(t) { ct = t; return this; },
    setName(n) { nm = n; return this; },
    getName: () => nm,
  };
}

function makeRoleServer(routinePath, membersPath) {
  const ROUTINE = JSON.parse(fs.readFileSync(routinePath, 'utf8'));
  const MEMBERS = JSON.parse(fs.readFileSync(membersPath, 'utf8'));
  // --- 書き込めるシート ---
  function makeSheet(name, grid) {
    const g = grid.map((r) => r.slice());
    const formulas = {};                                     // 'r,c' → '=…'
    let width = Math.max(0, ...g.map((r) => r.length));
    const pad = () => { g.forEach((r) => { while (r.length < width) r.push(''); }); };
    pad();
    const sh = {
      _grid: g, _formulas: formulas, _writes: [],
      getName: () => name,
      getLastRow: () => {
        for (let r = g.length - 1; r >= 0; r--) if (g[r].some((v) => v !== '' && v != null)) return r + 1;
        return 0;
      },
      getLastColumn: () => width,
      getMaxRows: () => g.length,
      insertRowsAfter(after, n) {
        for (let i = 0; i < n; i++) g.splice(after, 0, new Array(width).fill(''));
      },
      getDataRange() { return sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), Math.max(width, 1)); },
      getRange(r, c, nr, nc) {
        nr = nr || 1; nc = nc || 1;
        return {
          getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => {
            const row = g[r - 1 + i] || [];
            const v = row[c - 1 + j];
            return v == null ? '' : v;
          })),
          getFormulas: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) =>
            formulas[(r - 1 + i) + ',' + (c - 1 + j)] || '')),
          setValue(v) { return this.setValues([[v]]); },
          setValues(vals) {
            vals.forEach((row, i) => row.forEach((v, j) => {
              while (g.length < r + i) g.push(new Array(width).fill(''));
              if (c + j > width) { width = c + j; pad(); }
              g[r - 1 + i][c - 1 + j] = v;
              sh._writes.push({ row: r + i, col: c + j, value: v });
            }));
            return this;
          },
          setNumberFormat() { return this; },
        };
      },
    };
    return sh;
  }

  const sheets = Object.keys(ROUTINE).map((n) => makeSheet(n, ROUTINE[n]));

  // 休会日：列はあっても「定例会回数」が空の週（check_weekly_start.js と同じ見当のつけ方）
  const fmt = (d) => d.getFullYear() + '/' + ('0' + (d.getMonth() + 1)).slice(-2) + '/' + ('0' + d.getDate()).slice(-2);
  const HOLIDAYS = [];
  for (const n of Object.keys(ROUTINE)) {
    const g = ROUTINE[n];
    (g[0] || []).forEach((v, c) => {
      if (/^\d{4}\/\d{2}\/\d{2}$/.test(String(v)) && String(v) < '2026/10/01' && !String((g[1] || [])[c] || '').trim()) {
        HOLIDAYS.push(String(v));
      }
    });
  }
  sheets.push(makeSheet('休会日', HOLIDAYS.map((d) => [new Date(d + ' 00:00:00')])));

  // 参加者シート（9/30）：架空のビジター2名・ゲスト1名・キャンセル1名と、名簿の方の代理1名
  const SUB_FOR = MEMBERS.find((m) => /^船木/.test(m.name)) || MEMBERS[5];
  sheets.push(makeSheet('20260930参加者', [
    ['No.', '参加者氏名', 'ふりがな', 'カテゴリー', '会社名', '招待者', '備考', '種別', 'ステータス'],
    ['', '見本 一郎', 'みほん いちろう', '税理士', '見本会計', MEMBERS[0].name, '', 'Visitor', ''],
    ['', '見本 二郎', 'みほん じろう', '行政書士', '見本事務所', MEMBERS[1].name, '', 'Visitor', ''],
    ['', '見本 三郎', 'みほん さぶろう', '保険', '見本保険', MEMBERS[2].name, '', 'Visitor', 'キャンセル'],
    ['', '見本 花子', 'みほん はなこ', '司会', '見本企画', MEMBERS[3].name, '', 'Guest', ''],
    ['', '見本 代理', 'みほん だいり', '', '', SUB_FOR.name, '', 'Substitute', ''],
  ]));

  // 名簿：更新期限日・入会日を何人かに入れる（推定の確かめ用。9/30 を基準に）
  const members = MEMBERS.map((m) => Object.assign({ joinDate: '', renewDate: '', expireDate: '' }, m));
  members[10].expireDate = '2026/10/20';                    // 残り20日 → 30日以内
  members[11].expireDate = '2026/11/20';                    // 残り51日 → 60日以内
  members[12].expireDate = '2026/12/20';                    // 残り81日 → 90日以内
  members[13].expireDate = '2027/06/01';
  members[14].joinDate = '2026/09/28';                      // 前回（9/23）より後に入会
  members[15].renewDate = '2026/09/25';                     // 前回より後に更新
  members[16].joinDate = '2025/01/10';

  const props = {};
  const sandbox = {
    console,
    SpreadsheetApp: {},
    PropertiesService: { getScriptProperties: () => ({
      getProperty: (k) => (k in props ? props[k] : null),
      setProperty: (k, v) => { props[k] = String(v); },
      deleteProperty: (k) => { delete props[k]; },
    }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Utilities: {
      formatDate(d, tz, f) {
        const p = (n) => ('0' + n).slice(-2);
        return f.replace('yyyy', d.getFullYear()).replace('MM', p(d.getMonth() + 1)).replace('dd', p(d.getDate()))
          .replace(/(^|[^M])M(?!M)/, '$1' + (d.getMonth() + 1)).replace(/(^|[^d])d(?!d)/, '$1' + d.getDate());
      },
      // pptx の展開・再梱包（事前MTGのパワポで使う）
      newBlob: (data, type, name) => makeBlob(Buffer.isBuffer(data) ? data
        : Array.isArray(data) ? Buffer.from(data.map((b) => b & 255)) : Buffer.from(String(data), 'utf8'), type, name),
      base64Decode: (s) => Buffer.from(String(s), 'base64'),
      unzip: (blob) => Object.entries(readZip(blob._buf)).map(([n, b]) => makeBlob(b, '', n)),
      zip: (blobs, name) => makeBlob(writeZip(Object.fromEntries(blobs.map((b) => [b.getName(), b._buf]))), 'application/zip', name),
    },
    // 同梱のファイル（premtg_template.html など）を読む
    HtmlService: { createHtmlOutputFromFile: (n) => ({ getContent: () => fs.readFileSync(path.join(ROOT, n + '.html'), 'utf8') }) },
    getSS_: () => ({ getSheets: () => sheets, getSheetByName: (n) => sheets.find((s) => s.getName() === n) || null }),
    getMemberMaster: () => ({ ok: true, members: members.map((m) => Object.assign({}, m)) }),
    getMembersList: () => MEMBERS.map((m) => ({ no: m.no, name: m.name })),
    normName_: (s) => String(s == null ? '' : s).normalize('NFKC').replace(/[\s　]/g, ''),
    findPhotoIdForName_: () => '',
    SYSTEM_VERSION_: 'test',
  };
  sandbox.global = sandbox;
  vm.createContext(sandbox);
  for (const f of ['コード.js', 'ooxml.js', 'splice_srv.js', 'referral_srv.js', 'big_templates_srv.js',
                   'member_master_srv.js', 'routine_srv.js', 'member_presen_srv.js', 'meeting_slides_srv.js',
                   'visitor_post_srv.js', 'role_input_srv.js', 'speaker_rotation_srv.js', 'premtg_srv.js']) {
    vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), sandbox, { filename: f });
  }
  // メンバー名簿のシート（担当者を「役職」に反映する検査用）。名簿の方の役職は空欄にし、
  // 役職の書き方を確かめる行は架空の方で足す
  //   見本 前任 … 13の役職の名前だけ（前の期の担当者）→ 担当者でなければ空欄になる
  //   見本 兼任 … 13の役職の名前を2つ → 空欄になる
  //   見本 ウェブ … 全角で書いた役職の名前 → 空欄になる
  //   見本 経歴・見本 ホスト … 13の役職の名前以外も入っている → そのまま
  const HEAD = vm.runInContext('MEMBER_HEADERS_', sandbox).slice(), RC = HEAD.indexOf('役職');
  const rosterRow = (no, name, role) => { const r = HEAD.map(() => ''); r[0] = no; r[2] = name; r[RC] = role; return r; };
  const ROSTER_EXTRA = [['見本 前任', 'バイスプレジデント'], ['見本 兼任', 'BCP委員・スプレディング委員'], ['見本 ウェブ', 'Ｗｅｂマスター'],
                        ['見本 経歴', 'ビジホス　過去の経験は、書記兼会計、トレーニング委員'], ['見本 ホスト', 'ビジターホスト']];
  sheets.push(makeSheet('メンバー名簿', [HEAD].concat(
    members.map((m) => rosterRow(m.no, m.name, '')),
    ROSTER_EXTRA.map((x, i) => rosterRow(String(90 + i), x[0], x[1])))));

  // 業種区分マスタは初期値（member_master_srv.js の DEFAULT_CATEGORIES_）を使う
  sandbox.getCategoryMaster = () => vm.runInContext('DEFAULT_CATEGORIES_', sandbox)
    .map((r) => ({ key: r[0], label: r[1], block: r[4], order: r[5] }));
  // 画面から呼ぶ関数と同じく、名簿はシートではなく上の members を使う
  sandbox.getMemberMaster = () => ({ ok: true, members: members.map((m) => Object.assign({}, m)) });
  sandbox.getMembersList = () => MEMBERS.map((m) => ({ no: m.no, name: m.name }));
  // 「今日」は 2026/09/26（土）に固定する（次回＝9/30）
  const RealDate = Date;
  const TODAY = new RealDate('2026-09-26T09:00:00');
  sandbox.Date = class extends RealDate {
    constructor(...a) { if (a.length) super(...a); else super(TODAY.getTime()); }
    static now() { return TODAY.getTime(); }
  };
  vm.runInContext('Date = this.Date;', sandbox);

  const F = sandbox;
  return { F, sandbox, sheets, members, MEMBERS, SUB_FOR, props };
}

module.exports = { makeRoleServer, makeBlob };

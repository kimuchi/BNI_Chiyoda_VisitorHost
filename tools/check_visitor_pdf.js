// ビジターリスト・割り振り表のシートとPDF（コード.js の createFinalSheet / regeneratePdfOnly / saveAllocationSheet）を、
// Node上の見せかけのスプレッドシートで動かして確かめる。
//
//   node tools/check_visitor_pdf.js
//
// 確かめること
//   ・作ったばかりの開催日のシートが、自動アーカイブで隠れないこと（ほかの開催日だけが隠れる）
//   ・アーカイブ（非表示）してあった開催日を作り直す／PDFだけ作り直すと、その開催日のシートが表示に戻ること
//   ・隠れたシートからPDFを作らないこと（本物は真っ白のPDFになる。ここでは中身の無いPDFとして見分ける）
//   ・PDFが返ってこないときは、ドライブの前のPDFを上書きせずに止めること（混み合いはやり直す）
// 名前はすべて架空。

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }

// ---- 見せかけのスプレッドシート ----
// 表示・非表示は flush() で「保存」され、PDFの書き出しは保存された状態を見る（本物と同じく、flush の前の変更は書き出しに出ない前提）
let nextId = 100;
function makeSheet(ss, name, hidden) {
  const sh = {
    _name: name, _id: nextId++, _hidden: !!hidden, _saved: !!hidden, _values: [],
    getName: () => sh._name,
    getSheetId: () => sh._id,
    isSheetHidden: () => sh._hidden,
    showSheet: () => { sh._hidden = false; return sh; },
    hideSheet: () => {
      const others = ss._sheets.filter((x) => x !== sh && !x._hidden);
      if (!others.length) throw new Error("You can't hide all the sheets in a document.");
      if (ss._active === sh) ss._active = others[0];
      sh._hidden = true; return sh;
    },
    activate: () => { ss._active = sh; if (sh._hidden) sh._hidden = false; return sh; },
    clear: () => { sh._values = []; return sh; },
    getLastRow: () => sh._values.length,
    getDataRange: () => rangeOf(sh, 1, 1, sh._values.length, (sh._values[0] || []).length),
    getRange: (a, b, c, d) => {
      if (typeof a !== 'string') return rangeOf(sh, a, b, c || 1, d || 1);
      const m = a.match(/^([A-Z])(\d+)(?::([A-Z])(\d+))?$/);   // 「A3」「A2:H2」
      const r0 = Number(m[2]), c0 = m[1].charCodeAt(0) - 64;
      return m[3] ? rangeOf(sh, r0, c0, Number(m[4]) - r0 + 1, m[3].charCodeAt(0) - 64 - c0 + 1) : rangeOf(sh, r0, c0, 1, 1);
    },
  };
  return new Proxy(sh, { get: (t, k) => (k in t ? t[k] : () => t) });   // 列幅・行の高さなどは何もしない
}
function rangeOf(sh, r, c, nr, nc) {
  const rg = {
    setValues: (vals) => {
      vals.forEach((row, i) => {
        while (sh._values.length < r + i) sh._values.push([]);
        row.forEach((v, j) => { sh._values[r - 1 + i][c - 1 + j] = v; });
      });
      return proxy;
    },
    getValues: () => sh._values.slice(r - 1, r - 1 + nr).map((row) => row.slice(c - 1, c - 1 + nc)),
    getValue: () => ((sh._values[r - 1] || [])[c - 1] == null ? '' : sh._values[r - 1][c - 1]),
  };
  const proxy = new Proxy(rg, { get: (t, k) => (k in t ? t[k] : () => proxy) });   // 書式の指定は何もしない
  return proxy;
}
function makeSS(names) {
  const ss = { _sheets: [], _active: null };
  Object.assign(ss, {
    getId: () => 'SSID',
    getSheets: () => ss._sheets.slice(),
    getSheetByName: (n) => ss._sheets.find((s) => s._name === n) || null,
    insertSheet: (n) => { const s = makeSheet(ss, n, false); ss._sheets.push(s); ss._active = s; return s; },
    setActiveSheet: (s) => { ss._active = s; if (s._hidden) s._hidden = false; return s; },
    getActiveSheet: () => ss._active,
  });
  names.forEach(([n, hidden, rows]) => { const s = makeSheet(ss, n, hidden); if (rows) s._values = rows; ss._sheets.push(s); });
  ss._active = ss._sheets.find((s) => !s._hidden) || null;
  return ss;
}
const visibleNames = (ss) => ss._sheets.filter((s) => !s._hidden).map((s) => s._name);
const hiddenNames = (ss) => ss._sheets.filter((s) => s._hidden).map((s) => s._name);

// ---- ドライブ・書き出し ----
let ss, props, drive, fetchPlan, fetchLog, sleeps;
function reset(names, preset) {
  ss = makeSS(names);
  props = Object.assign({}, preset || {});
  drive = { updated: [], created: [] };
  fetchPlan = [];      // 先頭から順に使う応答の指定（空なら本物どおり）
  fetchLog = [];
  sleeps = [];
}
function signed(buf) { return Array.from(buf).map((b) => (b > 127 ? b - 256 : b)); }
function blobOf(buf, type) {
  let name = '';
  const b = { getBytes: () => signed(buf), getDataAsString: () => buf.toString('utf8'), getContentType: () => type,
              setName: (n) => { name = n; return b; }, getName: () => name };
  return b;
}
function exportPdf(url) {
  const gid = Number((url.match(/[?&]gid=(\d+)/) || [])[1]);
  const sh = ss._sheets.find((s) => s._id === gid);
  if (!sh) return { code: 400, body: Buffer.from('<html>no sheet</html>'), type: 'text/html' };
  // 隠れた（保存済みの状態で非表示の）シートは、本物だと真っ白のPDFになる
  const text = sh._saved ? '' : sh._values.map((r) => r.join('\t')).join('\n');
  return { code: 200, body: Buffer.from('%PDF-1.4\n' + text), type: 'application/pdf' };
}
function pdfText(blob) { return blob.getDataAsString().replace(/^%PDF-1\.4\n/, ''); }

const srv = {
  console: { log() {}, warn() {}, error() {} },
  Utilities: {
    formatDate: (d, tz, f) => {
      const p = (n) => ('0' + n).slice(-2);
      return f.replace('yyyy', d.getFullYear()).replace('MM', p(d.getMonth() + 1)).replace('dd', p(d.getDate()));
    },
    sleep: (ms) => { sleeps.push(ms); },
  },
  PropertiesService: { getScriptProperties: () => ({
    getProperty: (k) => (k in props ? props[k] : null),
    setProperty: (k, v) => { props[k] = String(v); },
  }) },
  SpreadsheetApp: { flush: () => { ss._sheets.forEach((s) => { s._saved = s._hidden; }); } },
  ScriptApp: { getOAuthToken: () => 'TOKEN' },
  UrlFetchApp: { fetch: (url, opt) => {
    fetchLog.push(url);
    const plan = fetchPlan.shift();
    const r = plan ? plan(url) : exportPdf(url);
    return { getResponseCode: () => r.code, getBlob: () => blobOf(r.body, r.type), getContent: () => signed(r.body) };
  } },
  Drive: { Files: {
    get: (id) => {
      if (!drive.files || !drive.files[id]) throw new Error('File not found: ' + id);
      return { id, trashed: !!drive.files[id].trashed };
    },
    update: (meta, id, blob) => {
      if (!drive.files || !drive.files[id]) throw new Error('File not found: ' + id);
      if (meta && meta.trashed === false) drive.files[id].trashed = false;
      if (blob) { drive.updated.push({ id, meta, blob }); drive.files[id].blob = blob; }
      return {};
    },
  } },
  DriveApp: {
    Access: { ANYONE_WITH_LINK: 'ANYONE_WITH_LINK' }, Permission: { VIEW: 'VIEW' },
    getFileById: (id) => ({
      getUrl: () => 'https://drive.example/' + id,
      getParents: () => ({ hasNext: () => true, next: () => folder }),
    }),
    getRootFolder: () => folder,
  },
};
const folder = { createFile: (blob) => {
  const id = 'F' + (drive.created.length + 1);
  drive.files = drive.files || {};
  drive.files[id] = { blob };
  drive.created.push({ id, blob });
  return { getId: () => id, getUrl: () => 'https://drive.example/' + id, setSharing: () => {} };
} };
vm.createContext(srv);
vm.runInContext(fs.readFileSync(path.join(ROOT, 'コード.js'), 'utf8'), srv, { filename: 'コード.js' });
srv.getSS_ = () => ss;
srv.chapterLabel_ = () => 'BNI 見本チャプター';
srv.getAllocationNote = () => '注意書き';

const HEADER = ['_No', '参加者氏名', 'ふりがな', 'カテゴリー', '会社名', '招待者', 'メモ（ビジターリストに表示）', '種別', 'メール'];
const ROWS = [
  { _No: 'V01', 参加者氏名: '見本 太郎', ふりがな: 'みほん たろう', カテゴリー: '税理士', 会社名: '見本商事', 招待者: '試験 花子', 種別: 'Visitor' },
  { _No: 'G01', 参加者氏名: '例示 次郎', ふりがな: 'れいじ じろう', カテゴリー: '', 会社名: '', 招待者: '', 種別: 'Guest' },
];
const OLD = (key, hidden) => [[key + '参加者', hidden], [key + '参加者_印刷用', hidden], [key + '割り振り表', hidden], [key + 'オリエン', hidden]];
const DISPLAY = '2026年9月30日(水) 第535回';

// ---- 開催日のキー ----
[['20260930', true], ['0930', true], ['1231', true], ['2026', false], ['1999', false], ['0000', false], ['20261332', false],
 ['', false], [null, false], ['20260930参加者', false]].forEach(([k, ok]) => {
  ck(srv.isMeetingKey_(k) === ok, '開催日のキーの判定が違う: ' + k + ' → ' + srv.isMeetingKey_(k));
});

// ---- 1) はじめて作る：その日のシートは見えたまま、ほかの開催日だけ隠れる。PDFに中身がある ----
reset([['メンバー名簿', false]].concat(OLD('20260923', false), OLD('0916', false)));
srv.createFinalSheet('2026-09-30T00:00:00', DISPLAY, ROWS, HEADER);
ck(ss.getSheetByName('20260930参加者') && !ss.getSheetByName('20260930参加者').isSheetHidden(), '1) 作った参加者シートが隠れた: 表示=' + visibleNames(ss).join(','));
ck(ss.getSheetByName('20260930参加者_印刷用') && !ss.getSheetByName('20260930参加者_印刷用').isSheetHidden(), '1) 作った印刷用シートが隠れた');
ck(OLD('20260923').concat(OLD('0916')).every(([n]) => ss.getSheetByName(n).isSheetHidden()), '1) 前の開催日が隠れていない: 表示=' + visibleNames(ss).join(','));
ck(!ss.getSheetByName('メンバー名簿').isSheetHidden(), '1) 開催日でないシートまで隠れた');
ck(drive.created.length === 1 && /見本 太郎/.test(pdfText(drive.created[0].blob)), '1) PDFに中身が無い（真っ白）: ' + (drive.created[0] && pdfText(drive.created[0].blob)));
ck(drive.created.length === 1 && drive.created[0].blob.getName() === DISPLAY + ' ビジター様リスト.pdf', '1) PDFの名前が違う');
ck(props['VISITOR_PDF_ID_20260930参加者'] === 'F1', '1) PDFのIDを控えていない');
ck(props.LATEST_VISITOR_LIST_URL === 'https://drive.example/F1', '1) 最新のPDFのURLを控えていない');
ck(ss.getActiveSheet() && ss.getActiveSheet().getName() === '20260930参加者', '1) 作った参加者シートが開いていない: ' + (ss.getActiveSheet() && ss.getActiveSheet().getName()));

// ---- 2) 以前の不具合で、その日のシートが隠れてしまっている状態から作り直す → 表示に戻り、PDFに中身がある。URLは同じ ----
reset([['メンバー名簿', false]].concat(OLD('20260930', true), [['20260930オープンネット', true]], OLD('20260923', false)),
      { 'VISITOR_PDF_ID_20260930参加者': 'F0' });
drive.files = { F0: { blob: null } };
srv.createFinalSheet('2026-09-30T00:00:00', DISPLAY, ROWS, HEADER);
ck(OLD('20260930').concat([['20260930オープンネット']]).every(([n]) => !ss.getSheetByName(n).isSheetHidden()),
   '2) 隠れていたその日のシートが表示に戻らない: 非表示=' + hiddenNames(ss).join(','));
ck(OLD('20260923').every(([n]) => ss.getSheetByName(n).isSheetHidden()), '2) 前の開催日が隠れていない');
ck(drive.updated.length === 1 && drive.updated[0].id === 'F0' && /見本 太郎/.test(pdfText(drive.updated[0].blob)),
   '2) 同じファイル（URLそのまま）に中身のあるPDFで差し替えていない: ' + JSON.stringify(drive.updated.map((u) => [u.id, pdfText(u.blob)])));
ck(drive.created.length === 0, '2) 新しいPDFを作ってしまった（URLが変わる）');

// ---- 2b) 前のPDFが消えていて差し替えられない → 新しく作り、URLが変わったことを知らせる ----
reset([['メンバー名簿', false]].concat(OLD('20260930', false)), { 'VISITOR_PDF_ID_20260930参加者': 'GONE' });
const html2b = srv.createFinalSheet('2026-09-30T00:00:00', DISPLAY, ROWS, HEADER);
ck(drive.created.length === 1 && /新しいPDFを作りました/.test(html2b) && props['VISITOR_PDF_ID_20260930参加者'] === 'F1', '2b) 新しく作ったのに、URLが変わったことを知らせない');

// ---- 3) アーカイブ済みの開催日で「PDFのみ再作成」→ その開催日が表示に戻り、PDFに中身がある ----
reset([['メンバー名簿', false], ['20260930参加者', true, [['No.'], ['V01']]],
       ['20260930参加者_印刷用', true, [['BNI 見本チャプターの定例会へようこそ'], [''], [DISPLAY], [''], ['No.', '参加者氏名'], ['V01', '見本 太郎']]],
       ['20260930割り振り表', true]], { 'VISITOR_PDF_ID_20260930参加者': 'F0' });
drive.files = { F0: { blob: null } };
srv.regeneratePdfOnly('20260930参加者');
ck(hiddenNames(ss).length === 0, '3) PDFのみ再作成で、その開催日のシートが表示に戻らない: ' + hiddenNames(ss).join(','));
ck(drive.updated.length === 1 && /見本 太郎/.test(pdfText(drive.updated[0].blob)), '3) PDFのみ再作成のPDFに中身が無い');
ck(drive.updated.length === 1 && drive.updated[0].meta.name === DISPLAY + ' ビジター様リスト.pdf', '3) PDFの名前が印刷用シートのA3から付いていない');

// ---- 4) 開催日の名前でない隠れたシートも、書き出す間だけ表示にして中身のあるPDFにし、終わったら隠し直す ----
reset([['メンバー名簿', false], ['見本_印刷用', true, [['見本 太郎']]]]);
const url4 = srv.exportSheetToPdf(ss.getSheetByName('見本_印刷用'), '見本.pdf', 'X_KEY');
ck(url4 === 'https://drive.example/F1' && /見本 太郎/.test(pdfText(drive.created[0].blob)), '4) 隠れたシートのPDFに中身が無い');
ck(ss.getSheetByName('見本_印刷用').isSheetHidden(), '4) 書き出したあと、隠れていたシートを隠し直していない');

// ---- 5) 書き出しが失敗し続ける → 止める。前のPDFは上書きしない。隠れていたシートは隠し直す ----
reset([['メンバー名簿', false], ['見本_印刷用', true, [['見本 太郎']]]], { X_KEY: 'F0' });
drive.files = { F0: { blob: 'OLD' } };
fetchPlan = [0, 1, 2].map(() => () => ({ code: 500, body: Buffer.from('<html>error</html>'), type: 'text/html' }));
let err5 = null;
try { srv.exportSheetToPdf(ss.getSheetByName('見本_印刷用'), '見本.pdf', 'X_KEY'); } catch (e) { err5 = e; }
ck(err5 && /PDFを作れませんでした/.test(err5.message) && /前のPDFはそのまま/.test(err5.message), '5) 失敗を知らせていない: ' + (err5 && err5.message));
ck(drive.updated.length === 0 && drive.created.length === 0 && drive.files.F0.blob === 'OLD', '5) 失敗したのにドライブのPDFを書き換えた');
ck(fetchLog.length === 3 && sleeps.length === 2, '5) 一時的な不調のやり直しの回数が違う: fetch=' + fetchLog.length + ' sleep=' + sleeps.length);
ck(ss.getSheetByName('見本_印刷用').isSheetHidden(), '5) 失敗したあと、隠れていたシートを隠し直していない');

// ---- 6) 混み合い（429）は待ってやり直し、うまくいけばそのまま進む ----
reset([['メンバー名簿', false], ['見本_印刷用', false, [['見本 太郎']]]]);
fetchPlan = [() => ({ code: 429, body: Buffer.from('<html>busy</html>'), type: 'text/html' })];
srv.exportSheetToPdf(ss.getSheetByName('見本_印刷用'), '見本.pdf', 'X_KEY');
ck(fetchLog.length === 2 && sleeps.length === 1 && drive.created.length === 1 && /見本 太郎/.test(pdfText(drive.created[0].blob)), '6) 混み合いのあと、やり直して作れていない');

// ---- 7) 200 でもPDFでない（ログイン画面など）→ 止める。やり直さない（権限などは待っても直らない） ----
reset([['メンバー名簿', false], ['見本_印刷用', false, [['見本 太郎']]]]);
fetchPlan = [() => ({ code: 200, body: Buffer.from('<!DOCTYPE html><html>sign in</html>'), type: 'text/html' })];
let err7 = null;
try { srv.exportSheetToPdf(ss.getSheetByName('見本_印刷用'), '見本.pdf', 'X_KEY'); } catch (e) { err7 = e; }
ck(err7 && /PDFを作れませんでした/.test(err7.message), '7) PDFでない応答を通してしまった');
ck(fetchLog.length === 1 && drive.created.length === 0, '7) PDFでない応答でやり直した／保存した');
// 403（権限）もやり直さない
reset([['メンバー名簿', false], ['見本_印刷用', false, [['見本 太郎']]]]);
fetchPlan = [() => ({ code: 403, body: Buffer.from('<html>denied</html>'), type: 'text/html' })];
let err7b = null;
try { srv.exportSheetToPdf(ss.getSheetByName('見本_印刷用'), '見本.pdf', 'X_KEY'); } catch (e) { err7b = e; }
ck(err7b && /応答 403/.test(err7b.message) && fetchLog.length === 1, '7) 403 でやり直した／知らせていない');

// ---- 8) 自動アーカイブ：開催日でないキー（年だけ等）では何も隠さない。開催日のキーなら、その日以外を隠す ----
reset([['メンバー名簿', false]].concat(OLD('20260930', false), OLD('20260923', false)));
srv.autoArchiveOtherDates_('2026');
ck(hiddenNames(ss).length === 0, '8) 年だけのキーで隠してしまった: ' + hiddenNames(ss).join(','));
srv.autoArchiveOtherDates_('20260930参加者');
ck(hiddenNames(ss).length === 0, '8) シート名そのままのキーで隠してしまった');
srv.autoArchiveOtherDates_('20260930');
ck(OLD('20260930').every(([n]) => !ss.getSheetByName(n).isSheetHidden()), '8) その日のシートまで隠した');
ck(OLD('20260923').every(([n]) => ss.getSheetByName(n).isSheetHidden()), '8) ほかの開催日を隠していない');
props.AUTO_ARCHIVE_ENABLED = 'false';
reset([['メンバー名簿', false]].concat(OLD('20260930', false), OLD('20260923', false)), { AUTO_ARCHIVE_ENABLED: 'false' });
srv.autoArchiveOtherDates_('20260930');
ck(hiddenNames(ss).length === 0, '8) 自動アーカイブを切っているのに隠した');

// ---- 9) 割り振り表：隠れた割り振り表を作り直しても、その日のシートは見え、PDFに中身がある ----
reset([['メンバー名簿', false]].concat(OLD('20260930', true), OLD('20260923', false)), { 'ALLOC_PDF_ID_20260930割り振り': 'F0' });
drive.files = { F0: { blob: null } };
const visitors = [{ no: 'V01', name: '見本 太郎', cat: '税理士', inviter: '試験 花子' }];
const pool = [{ no: '3', name: '架空 三郎' }];
srv.saveAllocationSheet('2026-09-30T00:00:00', DISPLAY, visitors, pool, { V01: '3' }, {}, { V01: [] }, {}, {});
ck(OLD('20260930').every(([n]) => !ss.getSheetByName(n).isSheetHidden()) && !ss.getSheetByName('20260930オープンネット').isSheetHidden(),
   '9) 割り振り表を作ったのに、その日のシートが隠れている: ' + hiddenNames(ss).join(','));
ck(OLD('20260923').every(([n]) => ss.getSheetByName(n).isSheetHidden()), '9) 割り振り表を作っても、前の開催日が隠れない');
ck(drive.updated.length === 1 && drive.updated[0].id === 'F0' && /見本 太郎/.test(pdfText(drive.updated[0].blob)) && /架空 三郎/.test(pdfText(drive.updated[0].blob)),
   '9) 割り振り表のPDFに中身が無い: ' + JSON.stringify(drive.updated.map((u) => pdfText(u.blob).slice(0, 80))));

// ---- 10) 書き出しの前に、表示に戻した状態を保存している（flush してから書き出す） ----
// 上の 2)〜4)・9) は flush をしないと、書き出しから見て隠れたままで中身の無いPDFになる。ここでは、作る処理の最後の書き出しより前に表示へ戻していることを確かめる
reset([['メンバー名簿', false]].concat(OLD('20260930', true)));
const realFlush = srv.SpreadsheetApp.flush;
let flushed = 0;
srv.SpreadsheetApp.flush = () => { flushed++; realFlush(); };
srv.createFinalSheet('2026-09-30T00:00:00', DISPLAY, ROWS, HEADER);
srv.SpreadsheetApp.flush = realFlush;
ck(flushed >= 1 && drive.created.length === 1 && /見本 太郎/.test(pdfText(drive.created[0].blob)), '10) 表示に戻したことが書き出しに届いていない');

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('OK ' + checks + '件の検査');

// 毎週の流れ（名簿・PDF・メール・割り振り表）の検査で使う、書き込める見せかけのスプレッドシート・ドライブ・Gmail。
//
//   const { makeEnv } = require('./lib_sheet_fake');
//   const env = makeEnv({ now: new Date(2026, 8, 29, 10, 0) });
//   env.reset([['メンバー名簿', false, [[見出し], [行], ...]], ...], { プロパティ: 値 });
//   vm.createContext(Object.assign(sandbox, env.globals));
//
// 本物に合わせてあるところ
//   ・シートを隠す／表示に戻すのは flush() で保存され、PDFの書き出しは保存された状態を見る。
//     隠れたシートを書き出すと、本物と同じく中身の無い（真っ白の）PDFになる
//   ・最後の1枚は隠せない。同じ名前のシートは作れない。setValues は範囲と値の大きさが合わないと止まる
//   ・書き出しの r1/c1/r2/c2（0から数える。r2・c2 はその手前まで）の範囲だけがPDFに入る
//   ・new Date()（引数なし）は now を返す（「今日」を固定する）
// PDFの中身は「%PDF-1.4」の次の行から、セルをタブ・行を改行でつないだ文字にしてある（検査で読みやすいように）。

class FakeBlob {
  constructor(buf, type, name) { this._buf = Buffer.isBuffer(buf) ? buf : Buffer.from(String(buf == null ? '' : buf), 'utf8'); this._type = type || ''; this._name = name || ''; }
  getBytes() { return Array.from(this._buf).map((b) => (b > 127 ? b - 256 : b)); }
  getDataAsString() { return this._buf.toString('utf8'); }
  getContentType() { return this._type; }
  setContentType(t) { this._type = t; return this; }
  getName() { return this._name; }
  setName(n) { this._name = n; return this; }
  copyBlob() { return new FakeBlob(Buffer.from(this._buf), this._type, this._name); }
}

function makeEnv(opt) {
  opt = opt || {};
  const NOW = (opt.now || new Date(2026, 8, 29, 10, 0, 0)).getTime();
  class FakeDate extends Date {
    constructor(...a) { if (a.length === 0) super(NOW); else super(...a); }
    static now() { return NOW; }
  }

  const env = {
    ss: null, props: {}, drive: null, mail: [], fetchLog: [], fetchPlan: [], sleeps: [], errors: [], nextSheetId: 100,
    Date: FakeDate, FakeBlob,
  };

  // ---- シート ----
  const cellsOf = (sh) => sh._values;
  const lastRowOf = (sh) => {
    for (let r = sh._values.length - 1; r >= 0; r--) if ((sh._values[r] || []).some((v) => v !== '' && v != null)) return r + 1;
    return 0;
  };
  const lastColOf = (sh) => {
    let c = 0;
    sh._values.forEach((row) => { for (let j = (row || []).length - 1; j >= 0; j--) if (row[j] !== '' && row[j] != null) { c = Math.max(c, j + 1); break; } });
    return c;
  };
  function makeRange(sh, r, c, nr, nc) {
    if (!(r >= 1 && c >= 1 && nr >= 1 && nc >= 1)) throw new Error('範囲の指定が正しくありません: ' + [r, c, nr, nc].join(','));
    const rg = {
      getRow: () => r, getColumn: () => c, getNumRows: () => nr, getNumColumns: () => nc,
      getLastRow: () => r + nr - 1, getLastColumn: () => c + nc - 1,
      setValues: (vals) => {
        if (!Array.isArray(vals) || vals.length !== nr) throw new Error('The number of rows in the data does not match the number of rows in the range. The data has ' + (vals && vals.length) + ' but the range has ' + nr + '.');
        vals.forEach((row, i) => {
          if (row.length !== nc) throw new Error('The number of columns in the data does not match the number of columns in the range. The data has ' + row.length + ' but the range has ' + nc + '.');
          while (sh._values.length < r + i) sh._values.push([]);
          row.forEach((v, j) => { sh._values[r - 1 + i][c - 1 + j] = v == null ? '' : v; });
        });
        return proxy;
      },
      setValue: (v) => rg.setValues([[v]]),
      getValues: () => {
        const out = [];
        for (let i = 0; i < nr; i++) {
          const row = [];
          for (let j = 0; j < nc; j++) { const v = (sh._values[r - 1 + i] || [])[c - 1 + j]; row.push(v == null ? '' : v); }
          out.push(row);
        }
        return out;
      },
      getValue: () => rg.getValues()[0][0],
      getDisplayValues: () => rg.getValues().map((row) => row.map((v) => (v instanceof Date ? v.toISOString() : String(v)))),
      getDisplayValue: () => rg.getDisplayValues()[0][0],
      clearContent: () => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) if (sh._values[r - 1 + i]) sh._values[r - 1 + i][c - 1 + j] = ''; return proxy; },
      getSheet: () => sh._proxy,
    };
    const proxy = new Proxy(rg, { get: (t, k) => (k in t ? t[k] : () => proxy) });   // 書式の指定などは何もしない
    return proxy;
  }
  const colNo = (s) => s.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  function a1(sh, s) {
    const m = String(s).match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
    if (!m) throw new Error('A1 の書き方が分かりません: ' + s);
    const r0 = Number(m[2]), c0 = colNo(m[1]);
    return m[3] ? makeRange(sh, r0, c0, Number(m[4]) - r0 + 1, colNo(m[3]) - c0 + 1) : makeRange(sh, r0, c0, 1, 1);
  }
  function makeSheet(ss, name, hidden, rows) {
    const sh = {
      _name: name, _id: env.nextSheetId++, _hidden: !!hidden, _saved: !!hidden, _values: (rows || []).map((r) => r.slice()), _maxCols: 26,
      getName: () => sh._name,
      setName: (n) => { sh._name = n; return sh._proxy; },
      getSheetId: () => sh._id,
      getParent: () => ss,
      isSheetHidden: () => sh._hidden,
      showSheet: () => { sh._hidden = false; return sh._proxy; },
      hideSheet: () => {
        const others = ss._sheets.filter((x) => x !== sh && !x._hidden);
        if (!others.length) throw new Error("You can't hide all the sheets in a document.");
        if (ss._active === sh) ss._active = others[0];
        sh._hidden = true; return sh._proxy;
      },
      activate: () => { ss._active = sh; sh._hidden = false; return sh._proxy; },
      clear: () => { sh._values = []; return sh._proxy; },
      clearContents: () => { sh._values = []; return sh._proxy; },
      getLastRow: () => lastRowOf(sh),
      getLastColumn: () => lastColOf(sh),
      getMaxRows: () => Math.max(1000, sh._values.length),
      getMaxColumns: () => Math.max(sh._maxCols, lastColOf(sh)),
      insertColumnsAfter: (after, n) => { sh._maxCols = Math.max(sh._maxCols, after + n); return sh._proxy; },
      getDataRange: () => makeRange(sh, 1, 1, Math.max(1, lastRowOf(sh)), Math.max(1, lastColOf(sh))),
      getRange: (a, b, c, d) => (typeof a === 'string' ? a1(sh, a) : makeRange(sh, a, b, c || 1, d || 1)),
      appendRow: (row) => { const at = lastRowOf(sh); while (sh._values.length < at) sh._values.push([]); sh._values[at] = row.slice(); return sh._proxy; },
    };
    sh._proxy = new Proxy(sh, { get: (t, k) => (k in t ? t[k] : () => t._proxy) });   // 列幅・行の高さ・枠の固定などは何もしない
    return sh._proxy;
  }
  function makeSS(list) {
    const ss = { _sheets: [], _active: null };
    Object.assign(ss, {
      getId: () => 'SSID', getName: () => '名簿システム（検査）', getUrl: () => 'https://docs.example/SSID',
      getSheets: () => ss._sheets.slice(),
      getSheetByName: (n) => ss._sheets.find((s) => s._name === n) || null,
      insertSheet: (n) => {
        if (ss._sheets.some((s) => s._name === n)) throw new Error('A sheet with the name "' + n + '" already exists. Please enter another name.');
        const s = makeSheet(ss, n, false); ss._sheets.push(s); ss._active = s; return s;
      },
      deleteSheet: (s) => { ss._sheets = ss._sheets.filter((x) => x !== s && x._proxy !== s); return ss; },
      setActiveSheet: (s) => { ss._active = s; s._hidden = false; return s; },
      getActiveSheet: () => ss._active,
      toast: () => {},
    });
    (list || []).forEach(([n, hidden, rows]) => ss._sheets.push(makeSheet(ss, n, hidden, rows)));
    ss._active = ss._sheets.find((s) => !s._hidden) || null;
    return ss;
  }

  env.reset = (sheets, props) => {
    env.ss = makeSS(sheets);
    env.props = Object.assign({}, props || {});
    env.drive = { files: {}, created: [], updated: [] };
    env.addFile('SSID', null, '名簿システム（検査）');   // スプレッドシートそのもの（PDFを同じフォルダに作るときに親を引く）
    env.mail = []; env.fetchLog = []; env.fetchPlan = []; env.sleeps = []; env.errors = [];
    env.denied = new Set();   // このIDのファイルは開けない（作った方以外で権限が無い）
    env.readOnly = new Set(); // このフォルダには作れない（閲覧だけの共有）
    env.sharingBlocked = false;   // リンクでの共有が組織の設定で禁止されている
    return env;
  };
  env.sheet = (n) => env.ss.getSheetByName(n);
  env.visibleNames = () => env.ss._sheets.filter((s) => !s._hidden).map((s) => s._name);
  env.hiddenNames = () => env.ss._sheets.filter((s) => s._hidden).map((s) => s._name);
  env.values = (n) => { const s = env.sheet(n); return s ? s._values.map((r) => (r || []).slice()) : null; };
  env.addFile = (id, blob, name) => { env.drive.files[id] = { id, blob, name: name || '', trashed: false }; return env.drive.files[id]; };
  env.pdfText = (blob) => String(blob && blob.getDataAsString ? blob.getDataAsString() : blob).replace(/^%PDF-1\.4\n/, '');
  env.fileText = (id) => env.pdfText(env.drive.files[id] && env.drive.files[id].blob);
  env.idOfUrl = (url) => String(url || '').replace('https://drive.example/', '');

  // ---- PDFの書き出し ----
  function exportPdf(url) {
    const q = {};
    String(url).replace(/[?&]([^=&]+)=([^&]*)/g, (m, k, v) => { q[k] = v; });
    const sh = env.ss._sheets.find((s) => s._id === Number(q.gid));
    if (!sh) return { code: 400, body: Buffer.from('<html>no such sheet</html>'), type: 'text/html' };
    if (sh._saved) return { code: 200, body: Buffer.from('%PDF-1.4\n'), type: 'application/pdf' };   // 隠れたシート → 真っ白
    const r1 = q.r1 != null ? Number(q.r1) : 0, c1 = q.c1 != null ? Number(q.c1) : 0;
    const r2 = q.r2 != null ? Number(q.r2) : sh._values.length, c2 = q.c2 != null ? Number(q.c2) : 99;
    const text = sh._values.slice(r1, r2).map((row) => (row || []).slice(c1, c2).map((v) => (v == null ? '' : v)).join('\t')).join('\n');
    return { code: 200, body: Buffer.from('%PDF-1.4\n' + text), type: 'application/pdf' };
  }
  const response = (r) => ({
    getResponseCode: () => r.code,
    getBlob: () => new FakeBlob(r.body, r.type, ''),
    getContent: () => Array.from(r.body).map((b) => (b > 127 ? b - 256 : b)),
    getContentText: () => r.body.toString('utf8'),
    getHeaders: () => ({ 'Content-Type': r.type }),
  });

  // ---- 日付の字 ----
  const pad = (n, w) => String(n).padStart(w, '0');
  function formatDate(d, tz, fmt) {
    return String(fmt).replace(/yyyy|MM|M|dd|d|HH|H|mm|ss|E/g, (t) => ({
      yyyy: d.getFullYear(), MM: pad(d.getMonth() + 1, 2), M: d.getMonth() + 1, dd: pad(d.getDate(), 2), d: d.getDate(),
      HH: pad(d.getHours(), 2), H: d.getHours(), mm: pad(d.getMinutes(), 2), ss: pad(d.getSeconds(), 2),
      E: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][d.getDay()],
    })[t]);
  }
  function parseCsv(text) {
    const rows = []; let row = [], cell = '', q = false;
    const s = String(text);
    for (let i = 0; i < s.length; i++) {
      const ch = s[i];
      if (q) {
        if (ch === '"') { if (s[i + 1] === '"') { cell += '"'; i++; } else q = false; }
        else cell += ch;
      } else if (ch === '"') q = true;
      else if (ch === ',') { row.push(cell); cell = ''; }
      else if (ch === '\n' || ch === '\r') { if (ch === '\r' && s[i + 1] === '\n') i++; row.push(cell); rows.push(row); row = []; cell = ''; }
      else cell += ch;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows;
  }

  // ---- ドライブのファイル ----
  const fileObj = (id) => {
    const f = env.drive.files[id];
    if (!f || env.denied.has(id)) throw new Error('No item with the given ID could be found, or you do not have permission to access it. (' + id + ')');
    return {
      getId: () => id, getUrl: () => 'https://drive.example/' + id, getName: () => f.name,
      getBlob: () => f.blob, isTrashed: () => !!f.trashed,
      setTrashed: (t) => { f.trashed = !!t; },
      setSharing: (a, p) => {
        if (env.sharingBlocked) throw new Error('Access denied: DriveApp.（リンクの共有が組織の設定で禁止されています）');
        f.sharing = a + '/' + p;
      },
      getParents: () => { let used = false; return { hasNext: () => !used, next: () => { used = true; return folder; } }; },
    };
  };
  // フォルダ：スプレッドシートのある「親フォルダ」（FOLDER）と、マイドライブ（ROOT）。env.readOnly に入れたフォルダには作れない
  const makeFolder = (fid, fname) => {
    const children = {};
    const fo = {
      getId: () => fid, getName: () => fname, getUrl: () => 'https://drive.example/folders/' + fid,
      createFile: (blob) => {
        if (env.readOnly.has(fid)) throw new Error('Access denied: DriveApp.（このフォルダは閲覧だけの共有です）');
        const id = 'F' + (env.drive.created.length + 1);
        env.addFile(id, blob, blob.getName());
        env.drive.files[id].parent = fid;
        env.drive.created.push({ id, blob, parent: fid });
        return fileObj(id);
      },
      getFoldersByName: (n) => { let used = !children[n]; return { hasNext: () => !used, next: () => { used = true; return children[n]; } }; },
      createFolder: (n) => { children[n] = makeFolder(fid + '/' + n, n); return children[n]; },
      getFiles: () => ({ hasNext: () => false, next: () => null }),
    };
    return fo;
  };
  const folder = makeFolder('FOLDER', '親フォルダ'), rootFolder = makeFolder('ROOT', 'マイドライブ');
  env.drive = { files: {}, created: [], updated: [] };

  const store = () => ({
    getProperty: (k) => (Object.prototype.hasOwnProperty.call(env.props, k) ? env.props[k] : null),
    setProperty: (k, v) => { env.props[k] = String(v); },
    deleteProperty: (k) => { delete env.props[k]; },
    getProperties: () => Object.assign({}, env.props),
    setProperties: (o) => { Object.keys(o).forEach((k) => { env.props[k] = String(o[k]); }); },
    getKeys: () => Object.keys(env.props),
  });

  env.globals = {
    Date: FakeDate,
    console: { log() {}, info() {}, warn() {}, error: (...a) => { env.errors.push(a.join(' ')); } },
    Utilities: {
      formatDate, parseCsv,
      sleep: (ms) => { env.sleeps.push(ms); },
      newBlob: (data, type, name) => new FakeBlob(Array.isArray(data) ? Buffer.from(data.map((b) => b & 255)) : data, type, name),
      base64Encode: (b) => Buffer.from(Array.isArray(b) ? b.map((x) => x & 255) : String(b)).toString('base64'),
      base64Decode: (s) => Array.from(Buffer.from(String(s), 'base64')).map((b) => (b > 127 ? b - 256 : b)),
      getUuid: () => 'uuid-' + env.fetchLog.length,
    },
    PropertiesService: { getScriptProperties: store, getDocumentProperties: store, getUserProperties: store },
    LockService: { getScriptLock: () => ({ tryLock: () => true, waitLock: () => {}, releaseLock: () => {}, hasLock: () => true }),
                   getDocumentLock: () => ({ tryLock: () => true, waitLock: () => {}, releaseLock: () => {} }) },
    SpreadsheetApp: {
      flush: () => { env.ss._sheets.forEach((s) => { s._saved = s._hidden; }); },
      getActiveSpreadsheet: () => env.ss,
      openById: () => env.ss,
      getUi: () => ({ alert: () => {}, ButtonSet: { OK: 'OK' } }),
    },
    ScriptApp: { getOAuthToken: () => 'TOKEN' },
    Session: { getActiveUser: () => ({ getEmail: () => 'tester@example.com' }), getEffectiveUser: () => ({ getEmail: () => 'tester@example.com' }) },
    UrlFetchApp: { fetch: (url, o) => {
      env.fetchLog.push({ url, options: o });
      const plan = env.fetchPlan.shift();
      if (plan) return response(plan(url, o));
      if (/docs\.google\.com\/spreadsheets\/d\/[^/]+\/export/.test(url)) return response(exportPdf(url));
      if (env.onFetch) return response(env.onFetch(url, o));
      throw new Error('検査では外に出ません: ' + url);
    } },
    Drive: { Files: {
      get: (id) => { const f = env.drive.files[id]; if (!f) throw new Error('File not found: ' + id); return { id, name: f.name, trashed: !!f.trashed }; },
      update: (meta, id, blob) => {
        const f = env.drive.files[id];
        if (!f) throw new Error('File not found: ' + id);
        if (meta && meta.trashed === false) f.trashed = false;
        if (meta && meta.name) f.name = meta.name;
        if (blob) { f.blob = blob; env.drive.updated.push({ id, meta, blob }); }
        return { id };
      },
    } },
    DriveApp: {
      Access: { ANYONE_WITH_LINK: 'ANYONE_WITH_LINK' }, Permission: { VIEW: 'VIEW' },
      getFileById: fileObj,
      getRootFolder: () => rootFolder,
      getFolderById: (id) => {
        if (env.denied.has(id)) throw new Error('No item with the given ID could be found, or you do not have permission to access it. (' + id + ')');
        return folder;
      },
    },
    GmailApp: { sendEmail: (to, subject, body, options) => { env.mail.push({ to, subject, body, options: options || {} }); } },
  };
  env.reset([], {});
  return env;
}

module.exports = { makeEnv, FakeBlob };

// ウェブアプリの画面一覧（webapp_srv.js の WEBAPP_PAGES_）とトップページのリンクを確かめる。
//
//   node tools/check_webapp_pages.js
//
// 確かめること
//   ・一覧のキーがURLにそのまま書ける文字だけで、重なっていない。開く画面のHTMLがある
//   ・?p=キー で開いたとき、一覧どおりの画面（file があればそちら）を、一覧の params つきで開く
//     （URLに書いた値が優先。?p=role_input&role=vice の role など）
//   ・トップページのリンクは ?p=キー だけ。
//     リンクの ? より後ろに <?= ?> で「&…」を足すと、Apps Script が & や = をURL用に置き換えて（%26 %3D）
//     開く画面が分からなくなる（事前MTGのリンクがトップページから開けなかった原因）

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }

let opened = null;
const sandbox = {
  console,
  HtmlService: {
    createTemplateFromFile: (name) => {
      const t = { _name: name };
      t.evaluate = () => {
        opened = { name, params: t.params, groups: t.groups };
        const out = { getContent: () => '<html><body>' + name + '</body></html>' };
        out.setTitle = () => out; out.addMetaTag = () => out;
        return out;
      };
      return t;
    },
    createHtmlOutput: (c) => { const o = { _content: c }; o.setTitle = () => o; o.addMetaTag = () => o; return o; },
    createHtmlOutputFromFile: (n) => ({ getContent: () => fs.readFileSync(path.join(ROOT, n + '.html'), 'utf8') }),
  },
  ScriptApp: { getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/TEST/exec' }) },
  Session: { getEffectiveUser: () => ({ getEmail: () => '' }) },
  SpreadsheetApp: { getActiveSpreadsheet: () => null },
  PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty() {} }) },
};
sandbox.global = sandbox;
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(ROOT, 'webapp_srv.js'), 'utf8'), sandbox, { filename: 'webapp_srv.js' });
const F = sandbox;
const PAGES = vm.runInContext('WEBAPP_PAGES_', sandbox);

// --- 一覧 ---
const items = PAGES.flatMap((g) => g.items);
const keys = items.map((it) => it.key);
ck(new Set(keys).size === keys.length, 'キーが重なっている: ' + keys.filter((k, i) => keys.indexOf(k) !== i).join(','));
for (const it of items) {
  ck(/^[a-z_]+$/.test(it.key), `キー「${it.key}」にURLで困る文字がある（英小文字と _ だけにする）`);
  ck(fs.existsSync(path.join(ROOT, (it.file || it.key) + '.html')), `「${it.label}」の画面 ${(it.file || it.key)}.html が無い`);
  ck(!it.query, `「${it.label}」に query がある（リンクが崩れる。file と params を使う）`);
}

// --- ?p=キー で開く ---
for (const it of items) {
  opened = null;
  F.doGet({ parameter: { p: it.key } });
  ck(opened && opened.name === (it.file || it.key), `?p=${it.key} で開いた画面: ${opened && opened.name}`);
  for (const [k, v] of Object.entries(it.params || {})) {
    ck(opened && opened.params && opened.params[k] === v, `?p=${it.key} で画面に ${k}=${v} が渡らない: ${JSON.stringify(opened && opened.params)}`);
  }
}
// 事前MTG：役職ごとの入力の画面を、事前MTGの欄を開いた状態で
opened = null;
F.doGet({ parameter: { p: 'premtg' } });
ck(opened && opened.name === 'role_input' && opened.params.view === 'premtg', '?p=premtg: ' + JSON.stringify(opened));
// URLに書いた値が優先・これまでのURLもそのまま使える
opened = null;
F.doGet({ parameter: { p: 'role_input', role: 'vice' } });
ck(opened && opened.name === 'role_input' && opened.params.role === 'vice' && !opened.params.view, '?p=role_input&role=vice: ' + JSON.stringify(opened));
opened = null;
F.doGet({ parameter: { p: 'role_input', view: 'premtg' } });
ck(opened && opened.name === 'role_input' && opened.params.view === 'premtg', '?p=role_input&view=premtg: ' + JSON.stringify(opened));
// 知らないキーはトップページ
opened = null;
F.doGet({ parameter: { p: 'role_input&view=premtg' } });
ck(opened && opened.name === 'webapp_home', '知らないキー（崩れたリンク）でトップページに戻らない: ' + JSON.stringify(opened && opened.name));
ck(opened && opened.groups && opened.groups.flatMap((g) => g.items).some((x) => x.key === 'premtg'), 'トップページに事前MTGが出ていない');

// --- トップページのリンク ---
const home = fs.readFileSync(path.join(ROOT, 'webapp_home.html'), 'utf8');
const hrefs = [...home.matchAll(/href="([^"]*<\?[\s\S]*?)"/g)].map((m) => m[1]);
for (const h of hrefs) {
  const end = h.indexOf('?>');                                 // 最初の <?= appUrl ?> のあとの「?」
  const q = end < 0 ? -1 : h.indexOf('?', end + 2);
  if (q < 0) continue;
  const after = h.slice(q);
  ck(/^\?p=<\?= items\[i\]\.key \?>$/.test(after), 'トップページのリンクの ? より後ろ: ' + after);
}

console.log(`ウェブアプリの画面一覧: 検査 ${checks} 件`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: キー・画面のHTML・?p= での開き方（事前MTG・役職の指定）・トップページのリンク');

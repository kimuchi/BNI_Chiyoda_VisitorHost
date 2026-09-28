// ウェブアプリの画面一覧（webapp_srv.js の WEBAPP_PAGES_）とトップページのリンクを確かめる。
//
//   node tools/check_webapp_pages.js
//
// 確かめること
//   ・一覧のキーがURLにそのまま書ける文字だけで、重なっていない。開く画面のHTMLがある
//   ・?p=キー で開いたとき、一覧どおりの画面（file があればそちら）を、一覧の params つきで開く
//     （URLに書いた値が優先。?p=role_input&role=vice の role など）
//   ・事前MTGのパワポは「役職ごとの入力」と入口を1つにまとめた（トップページには出さない）。
//     前のリンク ?p=premtg は、トップページに出さない入口（WEBAPP_ALIASES_）として開ける
//   ・公式ファイルから雛形を作る（初期設定のときだけ）も、トップページには出さず ?p=official_templates で開く。
//     画面には WEBAPP_URL を渡し、BNI 素材フォルダの画面のリンクからそのURLで開く
//   ・ウェブアプリで開くと、画面の先頭に「← メニューに戻る」の帯が入る。body を横並び（display:flex で
//     flex-direction が column でない）にしている画面は、帯が左の列になってしまうので、そうなっていない
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
for (const f of ['chapter_srv.js', 'webapp_srv.js']) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), sandbox, { filename: f });
const F = sandbox;
const PAGES = vm.runInContext('WEBAPP_PAGES_', sandbox);
const ALIASES = vm.runInContext('WEBAPP_ALIASES_', sandbox);

// --- 一覧 ---
const items = PAGES.flatMap((g) => g.items);
const keys = items.concat(ALIASES).map((it) => it.key);
ck(new Set(keys).size === keys.length, 'キーが重なっている: ' + keys.filter((k, i) => keys.indexOf(k) !== i).join(','));
for (const it of items.concat(ALIASES)) {
  ck(/^[a-z_]+$/.test(it.key), `キー「${it.key}」にURLで困る文字がある（英小文字と _ だけにする）`);
  ck(fs.existsSync(path.join(ROOT, (it.file || it.key) + '.html')), `「${it.label}」の画面 ${(it.file || it.key)}.html が無い`);
  ck(!it.query, `「${it.label}」に query がある（リンクが崩れる。file と params を使う）`);
  // 帯（先頭に入る「← メニューに戻る」）が左の列にならない：body を横並びにしていない
  const html = fs.existsSync(path.join(ROOT, (it.file || it.key) + '.html')) ? fs.readFileSync(path.join(ROOT, (it.file || it.key) + '.html'), 'utf8') : '';
  const bodyCss = (html.match(/(^|[\s}])body\s*\{[^}]*\}/g) || []).join(' ') + ' ' + ((html.match(/<body[^>]*style="([^"]*)"/i) || [])[1] || '');
  ck(!/display\s*:\s*(inline-)?flex/.test(bodyCss) || /flex-direction\s*:\s*column/.test(bodyCss),
     `「${it.label}」（${it.file || it.key}.html）は body を横並びにしている（ウェブアプリの「← メニューに戻る」が左の列になる）`);
}

// --- ?p=キー で開く ---
for (const it of items.concat(ALIASES)) {
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
// スピーカーローテーション：書記兼会計の画面から、ローテーションの管理を開く
opened = null;
F.doGet({ parameter: { p: 'rotation' } });
ck(opened && opened.name === 'role_input' && opened.params.role === 'secretary' && opened.params.view === 'rotation',
   '?p=rotation: ' + JSON.stringify(opened && { name: opened.name, params: opened.params }));
// URLに書いた値が優先・これまでのURLもそのまま使える
opened = null;
F.doGet({ parameter: { p: 'role_input', role: 'vice' } });
ck(opened && opened.name === 'role_input' && opened.params.role === 'vice' && !opened.params.view, '?p=role_input&role=vice: ' + JSON.stringify(opened));
opened = null;
F.doGet({ parameter: { p: 'role_input', view: 'premtg' } });
ck(opened && opened.name === 'role_input' && opened.params.view === 'premtg', '?p=role_input&view=premtg: ' + JSON.stringify(opened));
// 画面には、ウェブアプリのURL（WEBAPP_URL）を渡す。画面から別の画面を開くときに使う
// （BNI 素材フォルダの画面の「公式ファイルから雛形を作る」。トップページには出さない入口）
{
  const out = F.doGet({ parameter: { p: 'asset_settings' } });
  ck(out && /<script>var WEBAPP_URL = "https:\/\/script\.google\.com\/macros\/s\/TEST\/exec";<\/script>/.test(out._content || ''),
     '画面に WEBAPP_URL が渡らない: ' + String(out && out._content).slice(0, 300));
  ck(!items.some((x) => x.key === 'official_templates') && ALIASES.some((x) => x.key === 'official_templates'),
     '公式ファイルから雛形を作る：トップページに出ている／URLで開けない');
}
// 知らないキーはトップページ
opened = null;
F.doGet({ parameter: { p: 'role_input&view=premtg' } });
ck(opened && opened.name === 'webapp_home', '知らないキー（崩れたリンク）でトップページに戻らない: ' + JSON.stringify(opened && opened.name));
const homeItems = (opened && opened.groups ? opened.groups : []).flatMap((g) => g.items);
const roleItem = homeItems.find((x) => x.key === 'role_input');
ck(roleItem && /役職ごとの入力/.test(roleItem.label) && /事前MTG/.test(roleItem.label) && !homeItems.some((x) => x.file === 'role_input' && x.params && x.params.view === 'premtg'),
   'トップページの入口（役職ごとの入力・事前MTGのパワポを1つに）: ' + homeItems.filter((x) => (x.file || x.key) === 'role_input').map((x) => x.label).join(' / '));

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
console.log('OK: キー・画面のHTML・?p= での開き方（事前MTG・スピーカーローテーション・役職の指定・公式ファイルから雛形）・WEBAPP_URL・トップページの入口とリンク');

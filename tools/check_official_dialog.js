// 「公式ファイルから雛形を作る」の画面（official_templates.html）と、設定の保存（official_srv.js）を確かめる。
// 公式ファイルそのものは使わない（雛形を実際に作る検査は tools/check_official.js）。
//
//   node tools/check_official_dialog.js
//
// 確かめること
//   ・公式ファイルのリンク：共有リンク・IDから読む。Googleスライドに変換されたもの・pptxでないものは断る
//   ・画面：まだ登録していない雛形にだけチェックが入る。登録してあるものを作り直すときは確かめてから。
//     チェックした雛形を1つずつ順に作り、結果を行ごとに出す。途中で失敗しても残りを続ける
//   ・入口：初期設定のときだけ使うので、メニュー・ホーム画面・ウェブアプリのトップには出さない。
//     BNI 素材フォルダの画面の小さなリンクから開く（スプレッドシートではダイアログ、ウェブアプリでは ?p=official_templates）

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { loadPage } = require('./lib_minidom');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }
const J = (x) => JSON.stringify(x);

// --- サーバー（official_srv.js を、Drive とプロパティだけ差し替えて動かす）---
const props = {};
const FILES = {
  OFFICIALPPTX0000001: { name: 'BNI公式.pptx', mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', size: 97194303 },
  SLIDESCONVERTED0001: { name: 'BNI公式', mime: 'application/vnd.google-apps.presentation', size: 0 },
  PDFFILE000000000001: { name: '資料.pdf', mime: 'application/pdf', size: 1000 },
  REGISTEREDINTRO0001: { name: 'intro_自分で直した.pptx', mime: 'x', size: 1 },
};
const sb = {
  console,
  PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = String(v); } }) },
  DriveApp: {
    getFileById: (id) => {
      const f = FILES[id];
      if (!f) throw new Error('not found');
      return { getName: () => f.name, getMimeType: () => f.mime, getSize: () => f.size, getUrl: () => 'https://drive/' + id, getId: () => id };
    },
  },
};
vm.createContext(sb);
for (const f of ['chapter_srv.js', 'assets.js', 'big_templates_srv.js', 'official_srv.js']) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), sb, { filename: f });
}
sb.getMeetingCandidates = () => [];
sb.roleTermOf_ = () => 23;
const F = sb;

let st = F.getOfficialTemplateStatus();
ck(st.ok && st.file === null && st.kinds.length === 7 && st.kinds.every((k) => !k.registered), '初めの状態: ' + J(st).slice(0, 200));
let r = F.saveOfficialFileLink('https://docs.google.com/presentation/d/SLIDESCONVERTED0001/edit');
ck(!r.ok && /Googleスライドに変換/.test(r.message) && !props.BNI_OFFICIAL_FILE_ID, 'Googleスライドに変換されたファイルを断らない: ' + J(r));
r = F.saveOfficialFileLink('https://drive.google.com/file/d/PDFFILE000000000001/view');
ck(!r.ok && /PowerPoint/.test(r.message), 'pptxでないファイルを断らない: ' + J(r));
r = F.saveOfficialFileLink('https://drive.google.com/file/d/NOSUCHFILE000000001/view');
ck(!r.ok && /開けませんでした/.test(r.message), '開けないファイル: ' + J(r));
r = F.saveOfficialFileLink('https://docs.google.com/presentation/d/OFFICIALPPTX0000001/edit?usp=drive_link&rtpof=true&sd=true');
ck(r.ok && props.BNI_OFFICIAL_FILE_ID === 'OFFICIALPPTX0000001' && r.status.file.name === 'BNI公式.pptx' && r.status.file.sizeMB === 92.7,
   '共有リンクから設定する: ' + J(r).slice(0, 200));
props.BNI_TPL_INTRO_ID = 'REGISTEREDINTRO0001';
st = F.getOfficialTemplateStatus();
const intro = st.kinds.find((k) => k.kind === 'intro');
ck(intro.registered && !intro.fromOfficial && intro.fileName === 'intro_自分で直した.pptx', '登録してある雛形: ' + J(intro));
ck(J(st.seconds) === J({ weekly: 30, startup: 150, visitor: 20, referral: 7 }) && st.chapter === 'Activeチャプター', '秒数・チャプター: ' + J([st.seconds, st.chapter]));
r = F.buildOfficialTemplate('nokind');
ck(!r.ok && /種類が不正/.test(r.message), '種類が不正: ' + J(r));

// --- 画面 ---
const calls = [];
let buildResult = (kind) => (kind === 'meetingFirst' ? { ok: false, kind, message: '作れませんでした: テスト' }
  : { ok: true, kind, url: 'https://drive/' + kind, sizeMB: 1.5, message: '「' + kind + '」を作って登録しました（1.5MB）。' });
const server = {
  getOfficialTemplateStatus: () => { calls.push(['status']); return F.getOfficialTemplateStatus(); },
  saveOfficialFileLink: (v) => { calls.push(['save', v]); return F.saveOfficialFileLink(v); },
  buildOfficialTemplate: (kind) => { calls.push(['build', kind]); return buildResult(kind); },
};
const page = loadPage('official_templates.html', { server, fails });
const { els, run, step } = page;
step('開く', () => page.window.onload());
ck(/設定済み/.test(els.fileState.innerHTML) && /BNI公式\.pptx/.test(els.fileState.innerHTML), '公式ファイルの状態: ' + els.fileState.innerHTML);
ck(els.c_intro && !els.c_intro.checked && els.c_guest.checked && els.c_meetingFirst.checked, 'チェック（登録済みは外す）: '
   + ['intro', 'guest', 'meetingFirst'].map((k) => k + '=' + (els['c_' + k] && els['c_' + k].checked)).join(' '));
ck(/ウィークリー 30秒・スタートアッププレゼン 2分30秒・ビジタープレゼン 20秒・リファーラル発表 7秒/.test(els.secNote.innerHTML)
   && /BNI Activeチャプター/.test(els.secNote.innerHTML), '秒数の案内: ' + els.secNote.innerHTML);
step('リンクが空のまま設定', () => run('saveLink()'));
ck(/貼り付けてください/.test(els.msg.innerText) && !calls.some((c) => c[0] === 'save'), '空のリンク');
step('作る（ビジター紹介は外したまま）', () => { els.c_intro.checked = false; run('go()'); });
page.flush && page.flush();
const builds = calls.filter((c) => c[0] === 'build').map((c) => c[1]);
ck(J(builds) === J(['guest', 'dairi', 'presen', 'memberPresen', 'meetingFirst', 'meetingSecond']), '順に作る: ' + J(builds));
// 終わったあと一覧を読み直しても、行ごとの結果は残る
const cell = (k) => (els.kinds.innerHTML.match(new RegExp('id="r_' + k + '">(.*?)</td>')) || [])[1] || '';
ck(/登録しました/.test(cell('guest')) && /作れませんでした/.test(cell('meetingFirst')) && /登録しました/.test(cell('meetingSecond')) && cell('intro') === '',
   '行ごとの結果: ' + ['guest', 'meetingFirst', 'meetingSecond', 'intro'].map(cell).join(' | '));
ck(/作って登録しました（5件）/.test(els.msg.innerText) && /作れなかったもの/.test(els.msg.innerText) && page.log.confirms.length === 0,
   '終わりの知らせ: ' + els.msg.innerText.slice(0, 120));
step('登録してあるものも作り直す', () => { els.c_intro.checked = true; run('go()'); });
ck(page.log.confirms.length === 1 && /いま登録してある雛形の代わり/.test(page.log.confirms[0] || ''), '作り直す前に確かめない: ' + J(page.log.confirms));

// --- 入口：メニューには出さず、BNI 素材フォルダの画面の小さなリンクから ---
const menuSrc = fs.readFileSync(path.join(ROOT, 'コード.js'), 'utf8');
ck(!/addItem\([^)]*openOfficialTemplatesDialog/.test(menuSrc), 'スプレッドシートのメニューに「公式ファイルから雛形を作る」がある');
ck(!/openOfficialTemplatesDialog/.test(fs.readFileSync(path.join(ROOT, 'home_srv.js'), 'utf8')), 'ホーム画面に「公式ファイルから雛形を作る」がある');
{
  const wsb = { console, ScriptApp: { getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/TEST/exec' }) } };
  vm.createContext(wsb);
  vm.runInContext(fs.readFileSync(path.join(ROOT, 'webapp_srv.js'), 'utf8'), wsb, { filename: 'webapp_srv.js' });
  const top = vm.runInContext('WEBAPP_PAGES_', wsb).flatMap((g) => g.items).map((x) => x.key);
  const alias = vm.runInContext('WEBAPP_ALIASES_', wsb).map((x) => x.key);
  ck(top.indexOf('official_templates') < 0 && alias.indexOf('official_templates') >= 0,
     'ウェブアプリ：トップページには出さず ?p=official_templates で開く: ' + J({ top: top.includes('official_templates'), alias }));
}
const assetCalls = [];
const asset = loadPage('asset_settings.html', { fails, server: {
  getAssetSettings: () => ({ folderId: 'F1', reachable: true, folderUrl: '#', folderName: '素材', subFolders: [] }),
  openOfficialTemplatesDialog: () => { assetCalls.push('dialog'); },
} });
asset.step('素材フォルダの画面を開く', () => asset.window.onload());
const linkHtml = fs.readFileSync(path.join(ROOT, 'asset_settings.html'), 'utf8');
ck(/公式ファイルから雛形を作る<\/a>\s*（初期設定のときに使います。ふだんは使いません）/.test(linkHtml), '素材フォルダの画面に小さなリンクと「初期設定のとき」の案内が無い');
asset.step('スプレッドシートでリンクを押す', () => asset.run('openOfficial()'));
ck(J(assetCalls) === J(['dialog']), 'スプレッドシート：リンクで公式ファイルの画面（ダイアログ）を開かない: ' + J(assetCalls));
const opens = [];
asset.window.open = (u, t) => { opens.push([u, t]); };
asset.sandbox.WEBAPP_URL = 'https://script.google.com/macros/s/TEST/exec';
asset.step('ウェブアプリでリンクを押す', () => asset.run('openOfficial()'));
ck(J(opens) === J([['https://script.google.com/macros/s/TEST/exec?p=official_templates', '_top']]) && assetCalls.length === 1,
   'ウェブアプリ：リンクで ?p=official_templates を開かない: ' + J(opens));

console.log(`公式ファイルから雛形の画面: 検査 ${checks} 件`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: 公式ファイルのリンク（変換済み・pptx以外・開けない）・登録状況・チェックの初期値・順に作る・失敗しても続ける・作り直す前の確認・入口（メニューに出さず、素材フォルダの小さなリンクから）');

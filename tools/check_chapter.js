// チャプターの設定（chapter_srv.js）を確かめる。
//
//   node tools/check_chapter.js <routine.json> <members.json>
//
// 確かめること
//   ・何も保存していなければ Activeチャプターの値（名前・2026年4月〜9月が23期・2026/3/18(水)が第509回）で動く
//   ・名前を変えると、ウェブアプリの題名・下の行、メニュー画面の見出し、メールの件名と送り主の初期値、
//     メンバーブックの表紙の初期値が変わる。コードに「Activeチャプター」と書いた所が残っていない
//   ・定例会の基準（日と回数）を変えると、開催日の候補（曜日・回数）と回数の数え方・次の開催日が変わる
//   ・いまの期を変えると、期の番号と、期ごとに保存してある担当者・チーム・名簿へ反映した期の記録・
//     メンバーブックのプレジデント設定が同じだけずれる（同じ半期の担当者・挨拶文がそのまま出る）。戻すと元どおり
//   ・入力の誤り（名前が空・期が数字でない・日付が読めない・回数が0）は保存しない
//   ・役職の担当者の初期値はコードに持たない（何も保存していなければ空欄）
//   ・空のスプレッドシートから始めたとき（初回の準備の記録が fresh）は、メニュー画面の「準備の状況」でチャプターの設定を促す
//   ・プレゼンの秒数（ウィークリー30秒・スタートアップ2分30秒・ビジター20秒・リファーラル7秒が既定）を変えられる。
//     「3:00」「３０秒」の書き方も読む。5秒〜10分の外は保存しない。画面は分・秒の欄

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeRoleServer } = require('./lib_role_fixture');
const { loadPage } = require('./lib_minidom');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }
const J = (x) => JSON.stringify(x);

const { F, sandbox, props } = makeRoleServer(process.argv[2], process.argv[3]);
vm.runInContext(fs.readFileSync(path.join(ROOT, 'home_srv.js'), 'utf8'), sandbox, { filename: 'home_srv.js' });
const resetCache = () => vm.runInContext('CHAPTER_CACHE_ = null;', sandbox);
const D = (y, m, d) => new sandbox.Date(y, m - 1, d);

// ウェブアプリのトップページ（getSS_ を置き換えないよう、別の場所で動かす）
function webHome() {
  let got = null;
  const W = {
    console,
    HtmlService: {
      createTemplateFromFile: (name) => {
        const t = {};
        t.evaluate = () => {
          got = { name, title: t.title, footer: t.footer };
          const o = { getContent: () => '' }; o.setTitle = (x) => { got.tabTitle = x; return o; }; o.addMetaTag = () => o;
          return o;
        };
        return t;
      },
      createHtmlOutput: () => { const o = {}; o.setTitle = () => o; o.addMetaTag = () => o; return o; },
    },
    ScriptApp: { getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/TEST/exec' }) },
    Session: { getEffectiveUser: () => ({ getEmail: () => '' }) },
    SpreadsheetApp: { getActiveSpreadsheet: () => null },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => (k in props ? props[k] : null), setProperty() {} }) },
  };
  vm.createContext(W);
  for (const f of ['chapter_srv.js', 'webapp_srv.js']) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), W, { filename: f });
  W.doGet({ parameter: {} });
  return got;
}

// ===== 1. 何も保存していないとき（Activeチャプター）=====
ck(F.chapterLabel_() === 'Activeチャプター', '初期値の名前: ' + F.chapterLabel_());
ck(F.chapterSystemTitle_() === 'Activeチャプター 名簿システム', '初期値の題名: ' + F.chapterSystemTitle_());
ck(F.chapterFooter_() === 'BNI東京千代田リージョン ｜ Activeチャプター', '初期値の下の行: ' + F.chapterFooter_());
ck(F.roleTermOf_(D(2026, 9, 26)) === 23 && F.roleTermOf_(D(2026, 10, 7)) === 24 && F.roleTermOf_(D(2027, 3, 31)) === 24
   && F.roleTermOf_(D(2027, 4, 1)) === 25, '初期値の期（2026年4月〜9月が23期）');
ck(F.roleTermLabel_(24) === '2026年10月〜2027年3月', '期の月: ' + F.roleTermLabel_(24));
let cand = F.getMeetingCandidates();
ck(cand.length === 4 && /^2026\/9\/30\(水\) 第\d+回$/.test(cand[0].display), '初期値の開催日の候補: ' + J(cand.map((c) => c.display)));
const no930 = F.meetingCountOf_(D(2026, 9, 30));
ck(no930 > 509 && cand[0].display.endsWith('第' + no930 + '回'), '候補の回数と数え方が合わない: ' + no930 + ' / ' + cand[0].display);
ck(F.meetingCountOf_(D(2026, 3, 18)) === 509, '2026/3/18 が第509回でない: ' + F.meetingCountOf_(D(2026, 3, 18)));
ck(F.getTemplates().visitorSubj === '【Activeチャプター】{{date}} 定例会のご案内', 'メールの件名の初期値: ' + F.getTemplates().visitorSubj);
let cover = F.defaultCover_();
ck(cover.title === 'BNI Active chapter Member Book' && cover.aboutTitle === 'BNI Activeチャプターとは？'
   && cover.prole === 'Activeチャプター\n第23期プレジデント' && cover.term === '23期', '表紙の初期値: ' + J([cover.title, cover.aboutTitle, cover.prole, cover.term]));
let st = F.getChapterSettings();
ck(st.ok && st.name === 'Active' && st.term === 23 && st.weekday === '水' && st.meetingBaseDate === '2026/03/18'
   && st.meetingBaseCount === 509 && !st.saved && st.termLabel === '2026年4月〜9月', '画面に出す設定（初期値）: ' + J(st));
let web = webHome();
ck(web && web.title === 'Activeチャプター 名簿システム' && web.tabTitle === web.title
   && web.footer === 'BNI東京千代田リージョン ｜ Activeチャプター', 'ウェブアプリのトップ（初期値）: ' + J(web));
let home = F.getHomeStatus();
let chk = (home.checks || []).find((c) => c.key === 'chapter');
ck(home.title === 'Activeチャプター 名簿システム' && chk && chk.ready, 'メニュー画面（初期値）: ' + J({ title: home.title, chk }));

// ===== 2. 入力の誤りは保存しない =====
for (const bad of [{ name: '　', term: 23, meetingBaseDate: '2026/10/07', meetingBaseCount: 536 },
                   { name: 'Active', term: 'にじゅう', meetingBaseDate: '2026/10/07', meetingBaseCount: 536 },
                   { name: 'Active', term: 23, meetingBaseDate: '来週', meetingBaseCount: 536 },
                   { name: 'Active', term: 23, meetingBaseDate: '2026/02/30', meetingBaseCount: 536 },
                   { name: 'Active', term: 23, meetingBaseDate: '2026/10/07', meetingBaseCount: 0 }]) {
  const r = F.saveChapterSettings(bad);
  ck(!r.ok && r.message && !('BNI_CHAPTER' in props), '誤りを保存した: ' + J(bad) + ' → ' + J(r));
}

// ===== 3. 期ごとに保存してあるもの（担当者・チーム・名簿への反映の記録）=====
const MEM = F.getMemberMaster().members.map((m) => m.name);
const holders = (off) => Object.fromEntries(vm.runInContext('ROLE_DEFS_', sandbox).map((r, i) => [r.key, MEM[(i + off) % MEM.length]]));
F.saveRoleHolders(holders(0), 23, '2026/09/30');
F.saveRoleHolders(holders(3), 24, '2026/10/07',
  [{ key: 'role:vhc', name: 'ビジターホスト', members: [{ name: MEM[20], note: '' }, { name: MEM[21], note: 'サブリーダー' }] },
   { key: 'membership', name: 'メンバーシップ委員会', leader: MEM[1], members: [{ name: MEM[22], note: '' }] }]);
props.BNI_ROLE_ROSTER_STATE = J({ applied: 24, at: '2026/10/01', seen: 23 });
const teams24 = props.BNI_ROLE_TEAMS_24;
const before = { h930: F.roleHolders_(D(2026, 9, 30)), h1007: F.roleHolders_(D(2026, 10, 7)), t1007: F.roleTeams_(D(2026, 10, 7)) };
ck(teams24 && Object.keys(JSON.parse(props.BNI_ROLE_HOLDERS_TERMS)).sort().join() === '23,24', '準備（23期・24期の担当者と24期のチーム）: ' + props.BNI_ROLE_HOLDERS_TERMS);
// メンバーブックのプレジデント設定（期ごと）：期ごとにする前の1件（23期）と、24期の挨拶
props.BNI_MB_COVER = J({ term: '23期', pname: '見本 前期', ptext: '23期の挨拶' });
F.saveMemberBookCover({ ptext: '24期の挨拶' }, 24);
ck(Object.keys(JSON.parse(props.BNI_MB_PRESIDENTS)).sort().join() === '23,24', '準備（メンバーブックのプレジデント設定 23期・24期）: ' + props.BNI_MB_PRESIDENTS);

// ===== 4. 別のチャプター（名前・期・曜日を変える）=====
let r = F.saveChapterSettings({ name: 'BNI Brave チャプター', region: '', term: '５', meetingBaseDate: '2026-10-06', meetingBaseCount: '１００' });
ck(r.ok, '保存できない: ' + J(r));
const saved = JSON.parse(props.BNI_CHAPTER || '{}');
ck(saved.name === 'Brave' && saved.region === '' && saved.termBase === 5 && saved.meetingBaseDate === '2026/10/06' && saved.meetingBaseCount === 100,
   '保存した中身（「BNI」「チャプター」を外す・全角数字）: ' + props.BNI_CHAPTER);
ck(F.chapterLabel_() === 'Braveチャプター' && F.chapterFooter_() === 'Braveチャプター', '名前・下の行（リージョン空欄）: ' + F.chapterFooter_());
ck(F.roleTermOf_(D(2026, 9, 26)) === 5 && F.roleTermOf_(D(2026, 10, 7)) === 6 && F.roleTermLabel_(5) === '2026年4月〜9月'
   && F.roleTermLabel_(6) === '2026年10月〜2027年3月', '期の番号の付け直し: ' + F.roleTermOf_(D(2026, 9, 26)));
ck(Object.keys(JSON.parse(props.BNI_ROLE_HOLDERS_TERMS)).sort().join() === '5,6', '担当者の期がずれていない: ' + Object.keys(JSON.parse(props.BNI_ROLE_HOLDERS_TERMS)));
ck(props.BNI_ROLE_TEAMS_6 === teams24 && !('BNI_ROLE_TEAMS_24' in props), 'チームの期がずれていない: ' + Object.keys(props).filter((k) => /TEAMS/.test(k)));
ck(Object.keys(JSON.parse(props.BNI_MB_PRESIDENTS)).sort().join() === '5,6' && F.getCoverInfo_(5).ptext === '23期の挨拶' && F.getCoverInfo_(5).pname === '見本 前期'
   && F.getCoverInfo_(6).ptext === '24期の挨拶' && F.getCoverInfo_(6).prole === 'Braveチャプター\n第6期プレジデント' && F.getCoverInfo_().termNo === 5,
   'メンバーブックのプレジデント設定の期がずれていない: ' + props.BNI_MB_PRESIDENTS);
ck(/プレジデント設定の期も同じだけずらしました/.test(r.message), '期をずらした知らせ: ' + r.message);
const rs = JSON.parse(props.BNI_ROLE_ROSTER_STATE);
ck(rs.applied === 6 && rs.seen === 5 && rs.at === '2026/10/01', '名簿へ反映した期の記録: ' + props.BNI_ROLE_ROSTER_STATE);
ck(J(F.roleHolders_(D(2026, 9, 30))) === J(before.h930) && J(F.roleHolders_(D(2026, 10, 7))) === J(before.h1007)
   && J(F.roleTeams_(D(2026, 10, 7))) === J(before.t1007), '同じ半期の担当者・チームが出ない（期をずらしたあと）');
cand = F.getMeetingCandidates();
ck(cand[0].display === '2026/10/6(火) 第100回' && /^2026\/10\/13\(火\) 第101回$/.test(cand[1].display), '開催日の候補（火曜・第100回から）: ' + J(cand.map((c) => c.display)));
ck(F.meetingCountOf_(D(2026, 10, 20)) === 102 && F.meetingCountOf_(D(2026, 10, 21)) === 0, '回数の数え方（火曜）: ' + F.meetingCountOf_(D(2026, 10, 20)));
const nx = F.rotNextMeeting_([]);
ck(nx && nx.getDay() === 2 && F.fmtDate_(nx) === '2026/10/06', 'スピーカーローテーションの次の開催日: ' + (nx && F.fmtDate_(nx)));
ck(F.getTemplates().visitorSubj === '【Braveチャプター】{{date}} 定例会のご案内' && F.getTemplates().substituteSubj.indexOf('【Braveチャプター】') === 0,
   'メールの件名の初期値（名前を変えたあと）: ' + F.getTemplates().visitorSubj);
cover = F.defaultCover_();
ck(cover.title === 'BNI Brave chapter Member Book' && cover.aboutTitle === 'BNI Braveチャプターとは？'
   && cover.prole === 'Braveチャプター\n第5期プレジデント' && cover.term === '5期', '表紙の初期値（名前を変えたあと）: ' + J([cover.title, cover.prole]));
st = F.getChapterSettings();
ck(st.saved && st.name === 'Brave' && st.term === 5 && st.weekday === '火' && st.next === '2026/10/6(火) 第100回', '画面に出す設定（保存後）: ' + J(st));
web = webHome();
ck(web.title === 'Braveチャプター 名簿システム' && web.footer === 'Braveチャプター', 'ウェブアプリのトップ（名前を変えたあと）: ' + J(web));
home = F.getHomeStatus();
chk = (home.checks || []).find((c) => c.key === 'chapter');
ck(home.title === 'Braveチャプター 名簿システム' && chk && chk.ready && /Braveチャプター・いまの期 5期・次回 2026\/10\/6\(火\) 第100回/.test(chk.detail),
   'メニュー画面（名前を変えたあと）: ' + J(chk));

// ===== 5. 期を元に戻す（5期 → 23期）と、元どおり =====
r = F.saveChapterSettings({ name: 'Active chapter', region: 'BNI東京千代田リージョン', term: 23, meetingBaseDate: '2026/03/18', meetingBaseCount: 509 });
ck(r.ok && F.chapterLabel_() === 'Activeチャプター', '戻せない: ' + J(r));
ck(Object.keys(JSON.parse(props.BNI_ROLE_HOLDERS_TERMS)).sort().join() === '23,24' && props.BNI_ROLE_TEAMS_24 === teams24 && !('BNI_ROLE_TEAMS_6' in props)
   && JSON.parse(props.BNI_ROLE_ROSTER_STATE).applied === 24, '期を戻したあとの保存先: ' + Object.keys(props).filter((k) => /ROLE/.test(k)));
ck(Object.keys(JSON.parse(props.BNI_MB_PRESIDENTS)).sort().join() === '23,24' && F.getCoverInfo_(24).ptext === '24期の挨拶',
   '期を戻したあとのプレジデント設定: ' + props.BNI_MB_PRESIDENTS);
ck(J(F.roleHolders_(D(2026, 10, 7))) === J(before.h1007) && F.getMeetingCandidates()[0].display === '2026/9/30(水) 第' + no930 + '回',
   '期・曜日を戻したあとの担当者・候補');
ck(!/ずらしました/.test(F.saveChapterSettings({ name: 'Active', region: 'BNI東京千代田リージョン', term: 23, meetingBaseDate: '2026/03/18', meetingBaseCount: 509 }).message),
   '期を変えていないのに「ずらしました」と出る');

// ===== 5b. プレゼンの秒数（ウィークリー・スタートアップ・ビジター・リファーラル）=====
{
  const DEF = { weekly: 30, startup: 150, visitor: 20, referral: 7 };
  resetCache();
  ck(J(F.chapterPresenSeconds_()) === J(DEF) && J(F.getChapterSettings().seconds) === J(DEF), '秒数の初期値: ' + J(F.chapterPresenSeconds_()));
  ck(F.chapterSecondsLabel_(150) === '2分30秒' && F.chapterSecondsLabel_(30) === '30秒' && F.chapterSecondsLabel_(180) === '3分',
     '秒数の書き方: ' + [150, 30, 180].map(F.chapterSecondsLabel_).join(' / '));
  const base = { name: 'Active', region: 'BNI東京千代田リージョン', term: 23, meetingBaseDate: '2026/03/18', meetingBaseCount: 509 };
  const keep = props.BNI_CHAPTER;
  for (const bad of [{ weekly: 4 }, { startup: 601 }, { visitor: 'あいう' }, { referral: '' + 0 }]) {
    const rb = F.saveChapterSettings(Object.assign({}, base, { seconds: bad }));
    ck(!rb.ok && /秒数は、5秒〜10分/.test(rb.message) && props.BNI_CHAPTER === keep, '範囲の外の秒数を保存した: ' + J(bad) + ' → ' + J(rb));
  }
  let rs = F.saveChapterSettings(Object.assign({}, base, { seconds: { weekly: '45', startup: '3:00', visitor: '３０秒', referral: '0分10秒' } }));
  ck(rs.ok && J(JSON.parse(props.BNI_CHAPTER).seconds) === J({ weekly: 45, startup: 180, visitor: 30, referral: 10 })
     && J(F.chapterPresenSeconds_()) === J({ weekly: 45, startup: 180, visitor: 30, referral: 10 }),
     '秒数（「3:00」「３０秒」「0分10秒」の書き方）: ' + props.BNI_CHAPTER);
  ck(/ウィークリープレゼンテーション 45秒・スタートアッププレゼン 3分・ビジタープレゼン 30秒・リファーラル発表 10秒/.test(rs.message),
     '保存したときの知らせに秒数が出ない: ' + rs.message);
  // 秒数を渡さない保存（名前だけ直すなど）では、秒数はそのまま
  rs = F.saveChapterSettings(Object.assign({}, base));
  ck(rs.ok && F.chapterPresenSeconds_().startup === 180, '秒数を渡さない保存で、秒数が変わった: ' + J(F.chapterPresenSeconds_()));
  // 範囲の外の値が保存されていても（手で書き換えたなど）、その項目は既定の値で動く
  props.BNI_CHAPTER = J(Object.assign(JSON.parse(props.BNI_CHAPTER), { seconds: { weekly: 2, startup: 9999, visitor: 25 } }));
  resetCache();
  ck(J(F.chapterPresenSeconds_()) === J({ weekly: 30, startup: 150, visitor: 25, referral: 7 }), '壊れた秒数の読み方: ' + J(F.chapterPresenSeconds_()));

  // 画面：分・秒の欄に出して、直して保存する
  const sent = [];
  const page = loadPage('chapter_settings.html', { server: {
    getChapterSettings: () => F.getChapterSettings(),
    saveChapterSettings: (d) => { sent.push(d); return F.saveChapterSettings(d); } }, fails });
  const { els, run, step } = page;
  step('チャプターの設定を開く', () => page.window.onload());
  ck(String(els.sec_weekly_m.value) === '0' && String(els.sec_weekly_s.value) === '30' && String(els.sec_startup_m.value) === '2'
     && String(els.sec_startup_s.value) === '30' && String(els.sec_visitor_s.value) === '25',
     '分・秒の欄: ' + ['weekly', 'startup', 'visitor', 'referral'].map((k) => els['sec_' + k + '_m'].value + ':' + els['sec_' + k + '_s'].value).join(' '));
  ck(/5秒〜10分/.test(els.secNote.textContent) && /30秒・2分30秒・20秒・7秒/.test(els.secNote.textContent), '秒数の説明: ' + els.secNote.textContent);
  step('スタートアッププレゼンを3分に・ビジターを20秒に', () => {
    els.sec_startup_m.value = '3'; els.sec_startup_s.value = '0'; els.sec_visitor_m.value = ''; els.sec_visitor_s.value = '20';
    run('save()');
  });
  ck(sent.length === 1 && J(sent[0].seconds) === J({ weekly: 30, startup: 180, visitor: 20, referral: 7 })
     && J(F.chapterPresenSeconds_()) === J({ weekly: 30, startup: 180, visitor: 20, referral: 7 }), '画面から保存した秒数: ' + J(sent[0] && sent[0].seconds));
  // 元に戻す
  F.saveChapterSettings(Object.assign({}, base, { seconds: DEF }));
  ck(J(F.chapterPresenSeconds_()) === J(DEF), '秒数を戻せない');
}

// ===== 6. 空のスプレッドシートから始めたとき =====
delete props.BNI_CHAPTER; delete props.BNI_ROLE_HOLDERS_TERMS; delete props.BNI_ROLE_HOLDERS; resetCache();
let all = F.roleHolderTerms_();
ck(all[24] && Object.values(all[24]).every((v) => v === '') && vm.runInContext('ROLE_DEFS_', sandbox).every((d) => !('holder' in d)),
   '何も保存していないのに担当者が入る（担当者の初期値はコードに持たない）: ' + J(all[24]));
props.BNI_SETUP = J({ mode: 'fresh', at: '2026/09/26' }); resetCache();
home = F.getHomeStatus();
chk = (home.checks || []).find((c) => c.key === 'chapter');
ck(chk && !chk.ready && chk.fixFn === 'openChapterSettingsDialog', 'メニュー画面でチャプターの設定を促さない（初回）: ' + J(chk));
ck(F.getChapterSettings().fresh === true, '画面に「新しく始めた」が渡らない');
delete props.BNI_SETUP;

// ===== 7. コードに「Activeチャプター」を書いた所が残っていない（コメントは除く）=====
// メール送信用Web Appの合言葉の初期値（SECRET_TOKEN）は画面に出る名前ではないので除く
// （変えると、送信用のWeb App側と合わなくなってメールが送れなくなる）
const left = [];
for (const f of fs.readdirSync(ROOT)) {
  if (!/\.(js|html)$/.test(f) || f === 'manual.html') continue;
  fs.readFileSync(path.join(ROOT, f), 'utf8').split('\n').forEach((line, i) => {
    const code = line.replace(/^\s*\/\/.*$/, '').replace(/\s\/\/\s.*$/, '');
    if (/Active\s*(チャプター|chapter)|アクティブチャプター|Tokyo Chiyoda/i.test(code) && !/SECRET_TOKEN/.test(code)) left.push(f + ':' + (i + 1) + ' ' + line.trim().slice(0, 60));
  });
}
ck(!left.length, 'コードに「Activeチャプター」が残っている:\n    ' + left.join('\n    '));
const homeHtml = fs.readFileSync(path.join(ROOT, 'webapp_home.html'), 'utf8');
ck(/<h1><\?= title \?><\/h1>/.test(homeHtml) && /<footer><\?= footer \?>/.test(homeHtml), 'ウェブアプリのトップの見出し・下の行がチャプターの設定から出ない');

console.log(`チャプターの設定: 検査 ${checks} 件`);
if (fails.length) {
  console.log(`NG: ${fails.length} 件`);
  fails.forEach((f) => console.log('   ' + f));
  process.exit(1);
}
console.log('OK: 初期値（Activeチャプター）・名前（題名・メール・冊子）・定例会の曜日と回数・期の付け直し（担当者・チーム・記録）・入力の誤り・プレゼンの秒数・初回の準備');

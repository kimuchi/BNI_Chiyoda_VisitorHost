// パワーチームが求める専門分野のページ（powerteam_srv.js）を、同梱のひな形から作った作り物で確かめる。名前はすべて架空。
//
//   node tools/check_powerteam.js
//
// 確かめること
//   ・チームごとに1枚、「チャプターが求める専門分野」のページのすぐあとに入る（無ければ最後）
//   ・「〇〇パワーチーム」にチーム名、「カテゴリー」の欄に上から順に専門分野。足りない欄は空・多すぎるときは知らせる
//   ・セクションも並びに合わせる・pptxとして壊れていない
//   ・登録とローテーション：これからの予定・休会日を飛ばす・起点より前の回・1回に2チーム・前の形を読む
//   ・作成（generateMeetingSlides）：選んだチームだけ。画面で選び直しても登録は変えない。ひな形が未登録なら知らせる
process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const os = require('os');
const path = require('path');
const vm = require('vm');
const { spawnSync } = require('child_process');
const { makeEnv } = require('./lib_sheet_fake');
const { readZip, writeZip } = require('./lib_zip');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const J = (x) => JSON.stringify(x);

const env = makeEnv({ now: new Date(2026, 10, 17, 10, 0, 0) });
const F = Object.assign({}, env.globals);
vm.createContext(F);
for (const f of fs.readdirSync(ROOT).filter((x) => /\.js$/.test(x)).sort()) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), F, { filename: f });
env.reset([['メンバー名簿', false, [['No', '氏名'], ['1', '見本 一郎']]]], { BNI_CHAPTER: J({ name: '見本', termBase: 23, meetingBaseDate: '2026/03/18', meetingBaseCount: 509 }) });
let SAVED = null;
F.saveOutputFile_ = (blob, name) => { SAVED = { blob, name }; return { id: 'out', url: 'https://example/' + name, downloadUrl: 'https://example/dl/' + name }; };
F.findPhotoIdForName_ = () => '';

const tpl = (f) => fs.readFileSync(path.join(ROOT, 'docs', 'templates', f));
const para = (t) => '<a:p><a:r><a:rPr lang="ja-JP" sz="2000"/><a:t>' + t + '</a:t></a:r></a:p>';
// ひな形：熱烈歓迎のページの文字を、タイトル「〇〇パワーチームが求める専門分野」と「カテゴリー」の欄5つにする
function teamTemplate() {
  const z = readZip(tpl('BNI_テンプレート_熱烈歓迎.pptx')), p = 'ppt/slides/slide1.xml';
  let x = z[p].toString('utf8').replace('🎉 熱烈歓迎 🎉', '〇〇パワーチームが求める専門分野');
  const box = '<p:sp><p:nvSpPr><p:cNvPr id="901" name="表"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="5000000" cy="3000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>'
    + para('カテゴリー').repeat(5) + '</p:txBody></p:sp>';
  z[p] = Buffer.from(x.replace('</p:spTree>', box + '</p:spTree>'), 'utf8');
  return writeZip(z);
}
// 前半：事前MTGのひな形（3枚）の2枚目を「チャプターが求める専門分野」（{{求める専門分野1}}）にし、セクションを2つ付ける
function firstHalf({ wanted = true } = {}) {
  const z = readZip(tpl('BNI_テンプレート_事前MTG.pptx')), get = (p) => z[p].toString('utf8'), put = (p, s) => { z[p] = Buffer.from(s, 'utf8'); };
  let prs = get('ppt/presentation.xml');
  const ids = [...prs.matchAll(/<p:sldId id="(\d+)"/g)].map((m) => m[1]);
  const sec = '<p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">'
    + `<p14:section name="はじめ" id="{00000000-0000-0000-0000-000000000001}"><p14:sldIdLst><p14:sldId id="${ids[0]}"/><p14:sldId id="${ids[1]}"/></p14:sldIdLst></p14:section>`
    + `<p14:section name="あと" id="{00000000-0000-0000-0000-000000000002}"><p14:sldIdLst><p14:sldId id="${ids[2]}"/></p14:sldIdLst></p14:section>`
    + '</p14:sectionLst></p:ext>';
  prs = /<p:extLst>/.test(prs) ? prs.replace('<p:extLst>', '<p:extLst>' + sec) : prs.replace('</p:presentation>', '<p:extLst>' + sec + '</p:extLst></p:presentation>');
  put('ppt/presentation.xml', prs);
  if (wanted) put('ppt/slides/slide2.xml', get('ppt/slides/slide2.xml').replace(/<a:t>[^<]*<\/a:t>/, '<a:t>チャプターが求める専門分野 {{求める専門分野1}}</a:t>'));
  return writeZip(z);
}
const partsOf = (buf) => { const z = readZip(buf), m = {}; Object.keys(z).forEach((k) => { m[k] = new env.FakeBlob(z[k], '', k); }); return m; };
const zipOf = (parts) => { const o = {}; Object.keys(parts).forEach((k) => { o[k] = parts[k]._buf || Buffer.from(parts[k].getDataAsString(), 'utf8'); }); return o; };
const integrity = (z, label) => {
  const f = path.join(os.tmpdir(), 'check_powerteam_' + process.pid + '_' + label + '.pptx');
  fs.writeFileSync(f, writeZip(z));
  const r = spawnSync('python3', [path.join(ROOT, 'tools', 'pptx_integrity.py'), f], { encoding: 'utf8' });
  fs.unlinkSync(f);
  return { ok: r.status === 0, out: (r.stdout || '') + (r.stderr || '') };
};
const texts = (parts, p) => F.slideText_(F.xmlOf_(parts, p) || '');

// ===== 1) ページを入れる =====
{
  const parts = partsOf(firstHalf()), before = F.slideEntries_(parts).map((e) => e.path), wanted = before[1];
  const r = F.insertPowerTeamPages_(parts, partsOf(teamTemplate()), [{ name: '見本', cats: ['税理士', 'webデザイナー'] }, { name: '試験パワーチーム', cats: ['A', 'B', 'C', 'D', 'E', 'F'] }]);
  const order = F.slideEntries_(parts).map((e) => e.path);
  ck(r.pages === 2 && order.length === 5 && order[1] === wanted && J(order.slice(2, 4)) === J(r.paths) && order[4] === before[2],
     '1) チャプターが求める専門分野のすぐあとに2枚: ' + J({ order, paths: r.paths }));
  const t1 = texts(parts, r.paths[0]), t2 = texts(parts, r.paths[1]);
  ck(/見本パワーチームが求める専門分野/.test(t1) && /税理士/.test(t1) && /webデザイナー/.test(t1) && !/〇〇/.test(t1), '1) 1枚目の中身: ' + t1);
  ck(/試験パワーチームが求める専門分野/.test(t2) && !/試験パワーチームパワーチーム/.test(t2) && /A/.test(t2) && /E/.test(t2) && !/F/.test(t2), '1) 2枚目の中身: ' + t2);
  const x1 = F.xmlOf_(parts, r.paths[0]), cells = (x1.match(/<p:cNvPr id="901"[\s\S]*?<\/p:sp>/) || [''])[0].match(/<a:t>[^<]*<\/a:t>/g) || [];
  ck(J(cells) === J(['<a:t>税理士</a:t>', '<a:t>webデザイナー</a:t>', '<a:t></a:t>', '<a:t></a:t>', '<a:t></a:t>']), '1) 欄の順番・空の欄: ' + J(cells));
  ck(/入りきらなかったチーム.*試験（6件のうち5件）/.test(r.message) && /見本パワーチーム、試験パワーチーム/.test(r.message), '1) 知らせ: ' + r.message);
  const secIds = [...F.xmlOf_(parts, 'ppt/presentation.xml').matchAll(/<p14:sldId id="(\d+)"/g)].map((m) => m[1]);
  ck(F.slideEntries_(parts).every((e) => secIds.indexOf(String(e.id)) >= 0), '1) セクションに入っていないページがある');
  const it = integrity(zipOf(parts), 'ins');
  ck(it.ok, '1) pptxとして壊れている: ' + it.out);
  ck(F.insertPowerTeamPages_(parts, partsOf(teamTemplate()), [{ name: '', cats: [] }]) === null, '1) 空のチームでページを作った');
  ck(F.insertPowerTeamPages_(parts, partsOf(teamTemplate()), [{ name: '見本', cats: ['税理士'], on: false }]) === null, '1) 選んでいないチームでページを作った');
}
// チャプターが求める専門分野のページが無い：最後に入れて知らせる
{
  const parts = partsOf(firstHalf({ wanted: false }));
  const r = F.insertPowerTeamPages_(parts, partsOf(teamTemplate()), [{ name: '見本', cats: ['税理士'] }]);
  const order = F.slideEntries_(parts).map((e) => e.path);
  ck(order[order.length - 1] === r.paths[0] && /見つからないので、最後に入れました/.test(r.message), '1) ページが無いとき: ' + r.message);
}

// ===== 2) 登録とローテーション =====
{
  env.reset([['休会日', true, [['2026/12/30']]]], Object.assign({}, env.props, { BNI_CHAPTER: J({ name: '見本', termBase: 23, meetingBaseDate: '2026/03/18', meetingBaseCount: 509 }) }));
  F.requireSheetAccess_ = () => {};
  F.rotNextMeeting_ = () => new Date(2026, 10, 18);
  let r = F.savePowerTeamSettings({ teams: [{ name: '見本', cats: ['税理士'] }, { name: '試験', cats: ['保険'] }, { name: '架空パワーチーム', cats: ['印刷'] }],
                                    perMeeting: 1, anchor: { date: '2026/11/18', index: 1 } });
  ck(r.ok && J(r.teams.map((t) => t.name)) === J(['見本', '試験', '架空']), '2) 保存: ' + J(r));
  ck(J(r.schedule.slice(0, 4).map((x) => x.date + ':' + x.teams.join(''))) === J(['2026/11/18:試験', '2026/11/25:架空', '2026/12/02:見本', '2026/12/09:試験']),
     '2) これからの予定（起点の回から順に）: ' + J(r.schedule.slice(0, 4)));
  // 休会日（12/30）は飛ばす。起点より前の回は、さかのぼって数える
  const on = (d) => F.powerTeamsForDate_(new Date(d[0], d[1] - 1, d[2])).filter((t) => t.on).map((t) => t.name).join('・');
  ck(on([2026, 12, 23]) === '見本' && on([2026, 12, 30]) === '' && on([2027, 1, 6]) === '試験', '2) 休会日を飛ばす: ' + J([on([2026, 12, 23]), on([2026, 12, 30]), on([2027, 1, 6])]));
  ck(on([2026, 11, 11]) === '見本' && on([2026, 11, 4]) === '架空', '2) 起点より前の回: ' + J([on([2026, 11, 11]), on([2026, 11, 4])]));
  // 1回に2チーム
  r = F.savePowerTeamSettings({ teams: r.teams, perMeeting: 2, anchor: { date: '2026/11/18', index: 0 } });
  ck(J(r.schedule.slice(0, 3).map((x) => x.teams.join('・'))) === J(['見本・試験', '架空・見本', '試験・架空']), '2) 1回に2チーム: ' + J(r.schedule.slice(0, 3)));
  ck(!F.savePowerTeamSettings({ teams: [{ name: '', cats: ['税理士'] }] }).ok, '2) チーム名が空なのに保存した');
  // 前の形（チームの並びだけ）も読む
  env.props.BNI_POWER_TEAMS = J([{ name: '旧', cats: ['A'], on: true }]);
  ck(J(F.ptLoad_().teams) === J([{ name: '旧', cats: ['A'] }]), '2) 前の形を読めない: ' + J(F.ptLoad_()));
}

// ===== 3) 作成：選んだチームだけ・未登録の知らせ =====
{
  env.addFile('TPLF', new env.FakeBlob(firstHalf(), 'application/zip', 'first.pptx'), 'first.pptx');
  env.addFile('TPLT', new env.FakeBlob(teamTemplate(), 'application/zip', 'team.pptx'), 'team.pptx');
  env.props.BNI_TPL_MEETING_FIRST_ID = 'TPLF';
  env.props.BNI_TPL_POWERTEAM_ID = 'TPLT';
  const list = [{ name: '見本', cats: ['税理士'], on: true }, { name: '試験', cats: ['保険'], on: false }, { name: '架空', cats: ['印刷'], on: true }];
  const before = env.props.BNI_POWER_TEAMS;
  let res = F.generateMeetingSlides('meetingFirst', { 開催回: '540' }, '2026/11/18', { roleIntro: false, powerTeams: list });
  ck(res.ok && /パワーチームが求める専門分野のページを入れました（見本パワーチーム、架空パワーチーム）/.test(res.message), '3) 作成の知らせ（選んだチームだけ）: ' + J(res.message));
  ck(env.props.BNI_POWER_TEAMS === before, '3) 画面で選び直しても、登録は変えない');
  res = F.generateMeetingSlides('meetingFirst', { 開催回: '540' }, '2026/11/18', { roleIntro: false, powerTeams: list.map((t) => Object.assign({}, t, { on: false })) });
  ck(res.ok && !/パワーチーム/.test(res.message), '3) 0チームのとき: ' + J(res.message));
  delete env.props.BNI_TPL_POWERTEAM_ID;
  res = F.generateMeetingSlides('meetingFirst', { 開催回: '540' }, '2026/11/18', { roleIntro: false, powerTeams: list });
  ck(res.ok && /パワーチームが求める専門分野（前半）」のひな形を登録してください/.test(res.message), '3) 未登録の知らせ: ' + J(res.message));
}

// ===== 4) 設定の画面（ブラウザ）=====
(async () => {
  const pw = (() => { try { return require('playwright'); } catch (e) { return require('/opt/node22/lib/node_modules/playwright'); } })();
  env.props.BNI_POWER_TEAMS = J({ teams: [{ name: '見本', cats: ['税理士'] }, { name: '試験', cats: ['保険'] }], perMeeting: 1, anchor: { date: '2026/11/18', index: 0 } });
  const D = F.getPowerTeamSettings();
  const STUB = '<script>window.__calls=[];var D=' + J(D) + ';window.google={script:{run:new Proxy({},{get:function(_,n){var ok=null,p=new Proxy({},{get:function(_,m){'
    + 'if(m==="withSuccessHandler")return function(f){ok=f;return p;};if(m==="withFailureHandler")return function(){return p;};'
    + 'return function(){var a=[].slice.call(arguments);window.__calls.push([m,a]);var v=m==="savePowerTeamSettings"?Object.assign({},D,{message:"保存しました"}):D;setTimeout(function(){ok&&ok(JSON.parse(JSON.stringify(v)));},5);};}});return p[n];}})}};</script>';
  const html = fs.readFileSync(path.join(ROOT, 'powerteam_settings.html'), 'utf8').replace(/<head>/i, '<head><meta charset="utf-8">' + STUB);
  const browser = await pw.chromium.launch(), page = await browser.newPage();
  page.on('pageerror', (e) => fails.push('4) 画面のエラー: ' + e.message));
  page.on('dialog', (d) => d.accept());
  await page.route('https://pt.test/', (rt) => rt.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: html }));
  await page.goto('https://pt.test/');
  await page.waitForFunction(() => document.getElementById('n_0'), null, { timeout: 5000 }).catch(() => fails.push('4) チームが出ない'));
  const sch = () => page.evaluate(() => document.getElementById('sch').innerText);
  ck(/11\/18.*見本パワーチーム/.test(await sch()) && /11\/25.*試験パワーチーム/.test(await sch()), '4) これからの予定: ' + await sch());
  await page.click('button:has-text("＋ チームを足す")');
  await page.fill('#n_2', '架空');
  await page.dispatchEvent('#n_2', 'input');
  await page.fill('#c_2', '印刷、 広告');
  await page.dispatchEvent('#c_2', 'input');
  await page.selectOption('#aTeam', '2');
  ck(/11\/18.*架空パワーチーム/.test(await sch()) && /11\/25.*見本パワーチーム/.test(await sch()), '4) 変えたときの予定: ' + await sch());
  await page.click('#btnSave');
  await page.waitForTimeout(100);
  const sv = (await page.evaluate(() => window.__calls)).filter((c) => c[0] === 'savePowerTeamSettings').map((c) => c[1][0]);
  ck(sv.length === 1 && J(sv[0].teams) === J([{ name: '見本', cats: ['税理士'] }, { name: '試験', cats: ['保険'] }, { name: '架空', cats: ['印刷', '広告'] }])
     && sv[0].perMeeting === '1' && sv[0].anchor.index === '2' && sv[0].anchor.date === '2026/11/18', '4) 保存で送るもの: ' + J(sv));
  await browser.close();
  if (fails.length) { console.log('NG ' + fails.length + '件 / ' + checks + '件の検査'); fails.forEach((f) => console.log('  - ' + f)); process.exit(1); }
  console.log('パワーチームが求める専門分野: 検査 ' + checks + ' 件 OK');
})().catch((e) => { console.error(e); process.exit(1); });

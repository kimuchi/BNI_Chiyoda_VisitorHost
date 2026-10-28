// パスポートプログラム卒業コメントのページ（passport_srv.js）を、実データなしで確かめる。名前はすべて架空。
//
//   node tools/check_passport.js
//
// 確かめること
//   ・書記兼会計の入力に「パスポートプログラム卒業コメント」（名簿のメンバーを選ぶ欄）がある
//   ・入力の値（「、」区切り）から名簿の方を引く（会社名・【カテゴリー】）。「なし」・空なら作らない
//   ・ひな形の「苗字 名前」「【カテゴリー】」にお名前・カテゴリーが入る。「メンバー写真」の四角は、写真が無ければ文字だけ消す
//   ・お1人1枚、指定したページのあと（無ければ最後）に入る
//   ・ひな形が登録されていないときは、後半スライドは作り、登録してほしいことを知らせる

process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeEnv } = require('./lib_sheet_fake');
const { readZip, writeZip } = require('./lib_zip');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
function ck(ok, msg) { checks++; if (!ok) fails.push(msg); }
const J = (x) => JSON.stringify(x);

const env = makeEnv({ now: new Date(2026, 9, 5, 10, 0, 0) });
const F = Object.assign({}, env.globals);
vm.createContext(F);
for (const f of fs.readdirSync(ROOT).filter((x) => /\.js$/.test(x)).sort()) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), F, { filename: f });
F.HtmlService = { createHtmlOutputFromFile: (n) => ({ getContent: () => fs.readFileSync(path.join(ROOT, n + '.html'), 'utf8') }) };

// ===== 1) 入力の項目 =====
{
  const it = F.ROLE_EXTRA_ITEMS_.find((x) => x.title === 'パスポートプログラム卒業コメント');
  ck(it && it.role === 'secretary' && it.kind === 'members', '1) 書記兼会計の「パスポートプログラム卒業コメント」の項目: ' + J(it));
  const html = fs.readFileSync(path.join(ROOT, 'role_input.html'), 'utf8');
  ck(/it\.kind==='members'/.test(html) && /function memberPicks\(/.test(html), '1) 入力の画面に、メンバーを選ぶ欄が無い');
}

// 選べるのは入会日から4か月以内の方（入会日が無い方・先の入会日の方は選べない）
{
  const r = F.roleRecentMembers_([{ name: 'A', joinDate: '2026/06/08' }, { name: 'B', joinDate: '2026/06/06' }, { name: 'C', joinDate: '' },
                                  { name: 'D', joinDate: '2026/10/07' }, { name: 'E', joinDate: '2026/10/08' }], new Date(2026, 9, 7));
  ck(J(r) === J(['A', 'D']), '1) 入会日から4か月以内の方: ' + J(r));
  ck(/ctx\.recentMembers/.test(fs.readFileSync(path.join(ROOT, 'role_input.html'), 'utf8')), '1) 画面が4か月以内の方から選ばない');
}

// ===== 2) 入力の値から名簿の方 =====
F.getMemberMaster = () => ({ members: [{ name: '見本 一郎', company: '見本会計事務所', title: '税理士' },
                                       { name: '試験 二郎', company: '株式会社試験', title: '【工務店】' }] });
{
  const ms = F.passportMembersOf_('見本 一郎、試験二郎さん、架空 三子');
  ck(J(ms.map((m) => [m.name, m.company, m.category])) === J([['見本 一郎', '見本会計事務所', '【税理士】'], ['試験 二郎', '株式会社試験', '【工務店】'], ['架空 三子', '', '']]),
     '2) 名簿の方を引く: ' + J(ms));
  ck(F.passportMembersOf_('なし').length === 0 && F.passportMembersOf_('').length === 0, '2) 「なし」・空でページを作る');
}

// ===== 3) ページを作る =====
const builtin = () => { const h = fs.readFileSync(path.join(ROOT, 'welcome_template.html'), 'utf8'); return Buffer.from(h.match(/PPTX_BASE64_BEGIN([\s\S]*?)PPTX_BASE64_END/)[1].replace(/\s/g, ''), 'base64'); };
const partsOf = (buf) => F.unzipToMap_(new env.FakeBlob(buf, 'application/zip', 'x.pptx'));
function template() {
  const z = readZip(builtin()), p = 'ppt/slides/slide1.xml';
  z[p] = Buffer.from(z[p].toString('utf8').replace('{{氏名}}', '苗字　 名前').replace('{{カテゴリー}}', '【カテゴリー】'), 'utf8');
  return writeZip(z);
}
const texts = (parts, p) => (F.xmlOf_(parts, p).match(/<a:t>[^<]*<\/a:t>/g) || []).map((t) => t.replace(/<\/?a:t>/g, '')).join('|');
F.mpAddPhoto_ = () => null;
{
  const tgt = partsOf(builtin()), first = F.slideEntries_(tgt)[0].path;
  const ms = F.passportMembersOf_('見本 一郎、試験 二郎');
  const r = F.passportPages_(tgt, partsOf(template()), ms, { by: {}, seq: 0 }, '');
  const order = F.slideEntries_(tgt).map((e) => e.path);
  ck(r.pages.length === 2 && order.length === 3 && order[0] === first && J(order.slice(1)) === J(r.paths), '3) 最後にお1人1枚: ' + J(order) + J(r.paths));
  const t1 = texts(tgt, r.paths[0]), t2 = texts(tgt, r.paths[1]);
  ck(/見本 一郎/.test(t1) && /【税理士】/.test(t1) && !/苗字|【カテゴリー】|\{\{氏名|\{\{カテゴリー/.test(t1), '3) お名前・カテゴリー: ' + t1);
  ck(/試験 二郎/.test(t2) && /【工務店】/.test(t2), '3) 2人目: ' + t2);
  ck(/パスポートプログラム卒業コメントのページを入れました（見本 一郎、試験 二郎）/.test(r.message), '3) お知らせ: ' + r.message);
  // 指定したページのあと
  const tgt2 = partsOf(builtin()), e0 = F.slideEntries_(tgt2);
  const r2 = F.passportPages_(tgt2, partsOf(template()), ms.slice(0, 1), { by: {}, seq: 0 }, e0[0].path);
  ck(F.slideEntries_(tgt2)[1].path === r2.paths[0], '3) 指定したページのすぐあと');
  ck(F.passportPages_(tgt2, partsOf(template()), [], {}, '') === null, '3) 選んだ方がいないのに作る');
}
// 「メンバー写真」の四角：写真が無ければ文字だけ消す（四角はそのまま）
{
  const src = partsOf(template()), sp = F.slideOrder_(src)[0];
  F.putXml_(src, sp, F.xmlOf_(src, sp).replace('</p:spTree>', '<p:sp><p:nvSpPr><p:cNvPr id="900" name="四角"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>メンバー写真</a:t></a:r></a:p></p:txBody></p:sp></p:spTree>'));
  F.passportPrepare_(src);
  const x = F.xmlOf_(src, sp), slot = F.welcomePhotoSlot_(x, null);
  ck(/<a:t>お写真<\/a:t>/.test(x) && slot, '3) 「メンバー写真」を写真の場所にする: ' + J(slot));
}

if (fails.length) { console.error('NG ' + fails.length + '/' + checks + '\n' + fails.join('\n')); process.exit(1); }
console.log('OK ' + checks + ' checks');

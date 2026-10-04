// メンバー写真（⚙️ 設定 ＞ メンバー写真）：アップロードした写真が、すぐ氏名で見つかるか（写真索引）を確かめる。
// サーバーは本番の *.js を見せかけのスプレッドシート（lib_sheet_fake.js）の上で動かし、写真のフォルダ（02_メンバー写真）は
// この検査の中の作り物にする。画面（member_photos.html）は本物のブラウザ（Chromium）で、サーバーの返事は作り物。名前はすべて架空。
//
//   node tools/check_member_photos.js
//
// 確かめること
//   1) サーバー
//      ・アップロードすると写真索引も作り直し、新しい方の写真がすぐ見つかる（「写真が見つからないメンバー」から消える）
//      ・同じファイル名で入れ直すと、同じファイルの中身を差し替える（索引はそのまま）
//      ・同じ方の写真を別のファイル名で入れると、新しい方を使う（使わない方を知らせる）
//      ・ファイル名の空白の有無は問わない。HEIC など表示できない形式は保存しない（そのときは索引も作り直さない）
//      ・素材フォルダを開けない方は保存しない
//   2) 画面
//      ・アップロードのあと、保存した枚数・保存しなかったファイル（HEIC など）・写真索引の結果を出し、写真の状況を読み直す
//      ・20枚ごとに分けて送り、最後の写真索引の結果を出す。全部が表示できない形式なら赤で知らせる
//      ・写真索引を作り直せなかったときは「索引を作り直す」を押すよう知らせる

process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeEnv } = require('./lib_sheet_fake');
const pw = (() => { try { return require('playwright'); } catch (e) { return require('/opt/node22/lib/node_modules/playwright'); } })();

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const J = (x) => JSON.stringify(x);

// ===== 1) サーバー =====
const HEAD = ['No', '業種区分', '氏名', 'ふりがな', 'カテゴリー', '会社名', '役職', 'メモ', '写真ファイル名', '一言コメント',
  '紹介してほしい人', '協業したい人', '入会日', '更新日', '更新期限日', '会社での役職'];
const NAMES = ['見本 一郎', '試験 花子', '架空 三郎'];
const rosterRows = [HEAD].concat(NAMES.map((n, i) => HEAD.map((h) => (h === 'No' ? String(i + 1) : h === '氏名' ? n
  : h === '業種区分' ? '企業サポート' : h === 'カテゴリー' ? 'カテゴリー' + (i + 1) : h === '会社名' ? '見本会社' + (i + 1) : ''))));
const env = makeEnv({ now: new Date(2026, 9, 4, 10, 0, 0) });
const F = Object.assign({}, env.globals);
vm.createContext(F);
for (const f of fs.readdirSync(ROOT).filter((x) => /\.js$/.test(x)).sort()) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), F, { filename: f });
}
env.reset([['メンバー名簿', false, rosterRows]], {});

// 写真のフォルダ（02_メンバー写真）の代わり。ファイルごとに更新日時を持つ（同じ方の写真が複数あれば新しい方を使うため）
const photos = {};
let seq = 0, clock = Date.UTC(2026, 9, 1, 0, 0, 0);
const pfile = (id) => {
  const f = photos[id];
  if (!f || f.trashed) throw new Error('No item with the given ID could be found, or you do not have permission to access it. (' + id + ')');
  return { getId: () => id, getName: () => f.name, getBlob: () => f.blob, getMimeType: () => f.blob.getContentType(),
           getLastUpdated: () => new Date(f.updated), setTrashed: (t) => { f.trashed = !!t; } };
};
const iter = (ids) => { let i = 0; return { hasNext: () => i < ids.length, next: () => pfile(ids[i++]) }; };
const live = () => Object.keys(photos).filter((id) => !photos[id].trashed);
const photoFolder = {
  getFiles: () => iter(live()),
  getFilesByName: (n) => iter(live().filter((id) => photos[id].name === n)),
  createFile: (blob) => { const id = 'P' + (++seq); photos[id] = { name: blob.getName(), blob, updated: (clock += 60000) }; return pfile(id); },
};
const put = (name, type) => photoFolder.createFile(F.Utilities.newBlob('見本の写真:' + name, type, name)).getId();
F.getAssetFolder_ = () => photoFolder;
F.assetRootUnreachable_ = () => '';
F.Drive = { Files: { update: (meta, id, blob) => { photos[id].blob = blob; photos[id].updated = (clock += 60000); return { id }; } } };
F.DriveApp.getFileById = (id) => pfile(id);
const item = (name, type, text) => ({ name, mimeType: type, base64: Buffer.from(text || ('見本の写真:' + name)).toString('base64') });

const idOld = put('見本 一郎.jpg', 'image/jpeg');
put('試験花子__a1.png', 'image/png');                       // 空白なし・識別子つき
put('説明.txt', 'text/plain');                               // 写真でないものは数えない・知らせない
put('iPhone写真.HEIC', 'image/heic');                        // 表示できない形式は知らせる
let r = F.rebuildPhotoIndex();
ck(r.ok && r.total === 2 && J(r.unsupported) === J(['iPhone写真.HEIC']), '1) はじめの写真索引: ' + J(r));
let ov = F.getPhotoOverview();
ck(ov.indexed === 2 && ov.memberCount === 3 && J(ov.unmatched) === J(['架空 三郎']), '1) はじめの写真の状況: ' + J(ov));

// 新しい方の写真をアップロード → 索引を作り直さなくても見つかる
r = F.uploadMemberPhotosBase64([item('架空 三郎.jpg', 'image/jpeg')]);
const idNew = live().find((id) => photos[id].name === '架空 三郎.jpg');
ck(r.ok && r.saved === 1 && idNew, '1) アップロード: ' + J(r));
ck(r.index && r.index.ok && r.index.total === 3, '1) アップロードで写真索引を作り直さない: ' + J(r.index));
ck(/写真索引を作り直しました（3件）/.test(r.message), '1) アップロードの知らせに写真索引の結果が無い: ' + J(r.message));
ck(F.findPhotoIdForName_('架空 三郎') === idNew, '1) アップロードした新しい方の写真が見つからない（「索引を作り直す」を押すまで）: ' + J(F.findPhotoIdForName_('架空 三郎')));
ov = F.getPhotoOverview();
ck(ov.indexed === 3 && J(ov.unmatched) === J([]), '1) アップロードしても「写真が見つからないメンバー」に残る: ' + J(ov));
ck(F.findPhotoIdForName_('試験 花子') && F.findPhotoIdForName_('試験花子') === F.findPhotoIdForName_('試験 花子'), '1) ファイル名の空白の有無');

// 同じファイル名で入れ直す → 同じファイルの中身を差し替える（索引はそのまま）
r = F.uploadMemberPhotosBase64([item('架空 三郎.jpg', 'image/jpeg', '見本の写真:撮り直し')]);
ck(r.ok && r.saved === 1 && F.findPhotoIdForName_('架空 三郎') === idNew && photos[idNew].blob.getDataAsString() === '見本の写真:撮り直し'
   && live().filter((id) => photos[id].name === '架空 三郎.jpg').length === 1, '1) 同じファイル名で入れ直す: ' + J({ r, id: F.findPhotoIdForName_('架空 三郎') }));

// 同じ方の写真を別のファイル名で入れる → 新しい方を使う
r = F.uploadMemberPhotosBase64([item('見本一郎__2026.png', 'image/png')]);
const idNewer = live().find((id) => photos[id].name === '見本一郎__2026.png');
ck(r.ok && F.findPhotoIdForName_('見本 一郎') === idNewer && idNewer !== idOld, '1) 同じ方の新しい写真を使わない: ' + J(F.findPhotoIdForName_('見本 一郎')));
ck(/新しい方を採用/.test(r.message) && /見本 一郎\.jpg/.test(r.message), '1) 使わない写真の知らせ: ' + J(r.message));

// 表示できない形式だけ → 保存しない・索引も作り直さない
const before = live().length;
r = F.uploadMemberPhotosBase64([item('新しい方.HEIC', 'image/heic')]);
ck(r.ok && r.saved === 0 && J(r.rejected) === J(['新しい方.HEIC']) && r.index === null && live().length === before && /表示できない形式/.test(r.message),
   '1) HEIC: ' + J(r));

// 素材フォルダを開けない方は保存しない
F.assetRootUnreachable_ = () => '設定した「BNI 素材フォルダ」を、このアカウントでは開けません（見本）';
r = F.uploadMemberPhotosBase64([item('見本 二郎.jpg', 'image/jpeg')]);
ck(!r.ok && /開けません/.test(r.message) && live().length === before, '1) 素材フォルダを開けない方: ' + J(r));
F.assetRootUnreachable_ = () => '';

// ===== 2) 画面（Chromium）=====
// google.script.run の代わり。アップロードは HEIC を保存しない・写真索引の件数は保存した枚数ぶん増やす（window.__indexFail で失敗にする）
const STUB = '<script>(function(){window.__calls=[];window.__indexed=2;window.__unmatched=["架空 三郎"];window.__indexFail="";'
  + 'function answer(n,a){'
  + 'if(n==="getPhotoOverview")return {ok:true,indexed:window.__indexed,memberCount:3,unmatched:window.__unmatched.slice(),assets:{}};'
  + 'if(n==="uploadMemberPhotosBase64"){var b=a[0]||[],ok=b.filter(function(x){return !/\\.heic$/i.test(x.name);}),'
  + 'rej=b.filter(function(x){return /\\.heic$/i.test(x.name);}).map(function(x){return x.name;});'
  + 'window.__indexed+=ok.length;if(ok.length)window.__unmatched=[];'
  + 'var idx=ok.length?(window.__indexFail?{ok:false,message:window.__indexFail}:{ok:true,message:"写真索引を作り直しました（"+window.__indexed+"件）。",total:window.__indexed}):null;'
  + 'return {ok:true,saved:ok.length,rejected:rej,index:idx,message:""};}'
  + 'return null;}'
  + 'window.google={script:{host:{close:function(){}},get run(){var ok=null,p=new Proxy({},{get:function(_,n){'
  + 'if(n==="withSuccessHandler")return function(f){ok=f;return p;};'
  + 'if(n==="withFailureHandler"||n==="withUserObject")return function(){return p;};'
  + 'return function(){var a=[].slice.call(arguments);window.__calls.push([n,a.map(function(x){return Array.isArray(x)?x.map(function(y){return y.name;}):x;})]);'
  + 'var v=answer(n,a);setTimeout(function(){ok&&ok(JSON.parse(JSON.stringify(v)));},10);};}});return p;}}};})();</script>';

(async () => {
  const browser = await pw.chromium.launch();
  const URL = 'https://member-photos.test/';
  const html = fs.readFileSync(path.join(ROOT, 'member_photos.html'), 'utf8').replace(/<head>/i, '<head><meta charset="utf-8">' + STUB);
  const start = async () => {
    const page = await browser.newPage();
    page.on('pageerror', (e) => fails.push('画面のエラー: ' + e.message));
    page.on('dialog', (d) => d.accept());
    await page.route(URL, (rt) => rt.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: html }));
    await page.goto(URL);
    await page.waitForFunction(() => /索引済みの写真/.test(document.getElementById('stat').innerText), null, { timeout: 5000 }).catch(() => fails.push('2) 写真の状況が出ない'));
    return page;
  };
  const text = (page, id) => page.evaluate((i) => document.getElementById(i).innerText, id);
  const file = (name, type) => ({ name, mimeType: type, buffer: Buffer.from('見本の写真:' + name) });
  const upload = async (page, files, label) => {
    await page.setInputFiles('#files', files);
    await page.click('#btnUp');
    try { await page.waitForFunction(() => /枚の写真を保存しました/.test(document.getElementById('msg').innerText) && !document.getElementById('btnUp').disabled, null, { timeout: 8000 }); }
    catch (e) { fails.push(label + ' 終わりの知らせが出ない: ' + J(await text(page, 'msg'))); }
    await page.waitForTimeout(150);
    return text(page, 'msg');
  };
  const uploads = async (page) => (await page.evaluate(() => window.__calls)).filter((c) => c[0] === 'uploadMemberPhotosBase64').map((c) => c[1][0].length);

  // 新しい方の写真と、HEIC を一緒に選んだ
  let page = await start();
  ck(/写真が見つからないメンバー \(1名\)/.test(await text(page, 'stat')) && /架空 三郎/.test(await text(page, 'stat')), '2) はじめの写真の状況: ' + J(await text(page, 'stat')));
  let msg = await upload(page, [file('架空 三郎.jpg', 'image/jpeg'), file('x.heic', 'image/heic')], '2)');
  ck(/^1枚の写真を保存しました。/.test(msg), '2) 保存した枚数: ' + J(msg));
  ck(/表示できない形式のため保存しませんでした（1件: x\.heic）/.test(msg), '2) 保存しなかった HEIC を知らせない: ' + J(msg));
  ck(/写真索引を作り直しました（3件）/.test(msg), '2) 写真索引の結果を出さない: ' + J(msg));
  ck(/全員の写真が見つかりました/.test(await text(page, 'stat')), '2) 写真の状況を読み直さない: ' + J(await text(page, 'stat')));
  ck((await page.evaluate(() => getComputedStyle(document.getElementById('msg')).color)) === 'rgb(0, 128, 0)', '2) 知らせの色（緑）');
  await page.close();

  // 22枚 → 20枚と2枚に分けて送り、最後の写真索引の結果を出す
  page = await start();
  msg = await upload(page, Array.from({ length: 22 }, (_, i) => file('見本 ' + (i + 1) + '.jpg', 'image/jpeg')), '2) 22枚:');
  ck(J(await uploads(page)) === J([20, 2]), '2) 20枚ごとに分けて送らない: ' + J(await uploads(page)));
  ck(/^22枚の写真を保存しました。/.test(msg) && /写真索引を作り直しました（24件）/.test(msg), '2) 22枚の知らせ: ' + J(msg));
  await page.close();

  // 全部が表示できない形式 → 赤で知らせる（写真索引は作り直していない）
  page = await start();
  msg = await upload(page, [file('a.HEIC', 'image/heic')], '2) HEICだけ:');
  ck(/^0枚の写真を保存しました。/.test(msg) && /a\.HEIC/.test(msg) && !/写真索引/.test(msg), '2) HEICだけの知らせ: ' + J(msg));
  ck((await page.evaluate(() => getComputedStyle(document.getElementById('msg')).color)) === 'rgb(221, 0, 0)', '2) HEICだけのとき赤で知らせない');
  await page.close();

  // 写真索引を作り直せなかった → 「索引を作り直す」を押すよう知らせる
  page = await start();
  await page.evaluate(() => { window.__indexFail = '見本のエラー'; });
  msg = await upload(page, [file('架空 三郎.jpg', 'image/jpeg')], '2) 索引の失敗:');
  ck(/写真索引を作り直せませんでした（見本のエラー）。「索引を作り直す」を押してください。/.test(msg), '2) 写真索引を作り直せなかったときの知らせ: ' + J(msg));
  await page.close();

  await browser.close();
  if (fails.length) {
    console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
    fails.forEach((f) => console.log('  - ' + f));
    process.exit(1);
  }
  console.log('メンバー写真: 検査 ' + checks + ' 件 OK: アップロードで写真索引を作り直す（新しい方がすぐ見つかる）・同じファイル名の入れ直し・'
    + '同じ方の新しい写真・空白の有無・HEIC・素材フォルダを開けない方・画面の知らせ（枚数・HEIC・索引の結果・20枚ごと・索引の失敗）');
})().catch((e) => { console.error(e); process.exit(1); });

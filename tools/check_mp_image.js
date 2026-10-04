// 書記兼会計のスピーカーローテーションの画面（role_input.html）で作る「メインプレゼンターの画像」
// （今週のメインプレゼンテーション：写真・お名前・ご紹介して欲しいカテゴリー）を、本物のブラウザ（Chromium）で確かめる。
// サーバーの返事は作り物（名前はすべて架空）。書体は Meiryo UI（検査の環境に無いときは似た書体で描く）。
//
//   node tools/check_mp_image.js
//
// 確かめること
//   ・ご案内する回のお2人のお名前と、メンバー名簿の「協業したい人」を「・」で分けたカテゴリー（4つ以上は3つ目にまとめる）が欄に入る
//   ・メンバー写真を取りに行き、写真の枠に入る。写真の無い方は「写真なし」の枠。パソコンの写真を選ぶとそれを使う
//   ・1920×1080 の画像：左右のパネルの色・下の赤い帯・お名前・カテゴリーの数だけ枠（カテゴリーの無い方は見出しも出さない）
//   ・欄を直すと描き直す。長いカテゴリーも描ける。ご案内する回を替えると、その回のお2人になる
//   ・保存（ダウンロード）・Driveに保存のファイル名は「開催日_メインプレゼンター.png」
//   ・スピーカーローテーションの表の画像も、これまでどおり作る
//   ・書体は2枚とも Meiryo UI（Googleフォントは読みに行かない）
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');
const pw = (() => { try { return require('playwright'); } catch (e) { return require('/opt/node22/lib/node_modules/playwright'); } })();

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const J = (x) => JSON.stringify(x);

// --- 作り物の写真（単色のPNG）---
function png(w, h, rgb) {
  const chunk = (type, data) => {
    const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
    const td = Buffer.concat([Buffer.from(type), data]), crc = Buffer.alloc(4);
    crc.writeUInt32BE(zlib.crc32(td) >>> 0);
    return Buffer.concat([len, td, crc]);
  };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2;
  const row = Buffer.alloc(1 + w * 3); for (let x = 0; x < w; x++) row.set(rgb, 1 + x * 3);
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr),
                        chunk('IDAT', zlib.deflateSync(Buffer.concat(Array.from({ length: h }, () => row)))), chunk('IEND', Buffer.alloc(0))]);
}
const BLUE = png(300, 400, [0x33, 0x66, 0xcc]), GREEN = png(400, 300, [0x22, 0xaa, 0x44]);

// --- サーバーの返事（作り物）---
const M = [
  { name: '見本 一郎', title: '税理士', collab: '起業支援' },
  { name: '試験 二郎', title: '弁護士（企業法務）', collab: '司法書士・税理士・外国人就労支援事業者' },
  { name: '架空 三郎', title: '工務店', collab: '不動産賃貸管理・内装業・インテリアデザイン会社・建築設計' },
  { name: '仮名 四郎', title: 'Web制作', collab: '' },
];
const person = (m) => ({ name: m.name, title: m.title, collab: m.collab, company: '', inMaster: true });
const ROT = {
  ok: true, order: M.map((m) => m.name), excluded: [], anchor: { date: '2031/10/08', pointer: 0 },
  header: 'メインプレゼンテーション（各４分45秒）', notes: ['見本の注意書き'], updated: '2031/09/01 10:00', provisional: false,
  rebased: false, openDate: '2031/10/08', missing: [], holidayHints: [], fbText: '', secretary: '例示 五郎', holidays: [],
  chapter: '見本チャプター',
  weeks: [
    { date: '2031/10/08', no: '536', md: '10/8(水)', label: '10月8日', source: 'routine', people: [person(M[0]), person(M[1])], secretary: '例示 五郎' },
    { date: '2031/10/15', no: '537', md: '10/15(水)', label: '10月15日', source: 'routine', people: [person(M[2]), person(M[3])], secretary: '例示 五郎' },
  ],
  members: M.map((m) => ({ name: m.name, title: m.title, collab: m.collab, inOrder: true })),
};
const ANSWERS = {
  getSystemVersion: '',
  getRoleInputContext: { ok: false, message: '検査では読まない' },
  getSpeakerRotation: ROT,
  getMemberPhotosBase64: { ok: true, map: { '見本 一郎': 'data:image/png;base64,' + BLUE.toString('base64') } },
  saveSpeakerRotationImage: { ok: true, url: '#drive', downloadUrl: '#dl', fileName: '', message: '「03_生成物」に保存しました。' },
};
// google.script.run の代わり。呼ばれた関数と引数を window.__calls に残す
const STUB = '<script>window.__calls=[];var __R=' + J(ANSWERS).replace(/</g, '\\u003c') + ';'
  + 'var google={script:{host:{close:function(){}},get run(){var ok=null,p=new Proxy({},{get:function(_,n){'
  + 'if(n==="withSuccessHandler")return function(f){ok=f;return p;};'
  + 'if(n==="withFailureHandler"||n==="withUserObject")return function(){return p;};'
  + 'return function(){var a=[].slice.call(arguments);window.__calls.push([n,a]);var v=__R[n]===undefined?null:JSON.parse(JSON.stringify(__R[n]));'
  + 'if(n==="saveSpeakerRotationImage"&&v)v.fileName=a[1];setTimeout(function(){ok&&ok(v);},10);};}});return p;}}};</script>';

// Apps Script のテンプレート（<? ?>・<?= ?>・<?!= ?>）を展開する（tools/make_manual_shots.js と同じ）
function evalTemplate(file, vars) {
  const src = fs.readFileSync(path.join(ROOT, file), 'utf8');
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  let code = 'var __o=[];with(__v){', i = 0, m;
  const re = /<\?(!=|=)?([\s\S]*?)\?>/g;
  while ((m = re.exec(src))) {
    code += '__o.push(' + J(src.slice(i, m.index)) + ');';
    if (m[1] === '=') code += '__o.push(__e(' + m[2] + '));';
    else if (m[1] === '!=') code += '__o.push(String(' + m[2] + '));';
    else code += m[2] + '\n';
    i = re.lastIndex;
  }
  code += '__o.push(' + J(src.slice(i)) + ');}return __o.join("");';
  const include = (n) => fs.readFileSync(path.join(ROOT, n + '.html'), 'utf8');
  return new Function('__v', '__e', code)(Object.assign({ include }, vars), esc);
}

// 画像に使った書体（canvas の font に入れた文字）を window.__fonts に残す
const FONT_SPY = '<script>window.__fonts=[];(function(){var d=Object.getOwnPropertyDescriptor(CanvasRenderingContext2D.prototype,"font");'
  + 'Object.defineProperty(CanvasRenderingContext2D.prototype,"font",{configurable:true,get:function(){return d.get.call(this);},'
  + 'set:function(v){window.__fonts.push(String(v));d.set.call(this,v);}});})();</script>';

(async () => {
  const html = evalTemplate('role_input.html', { params: { role: 'secretary', view: 'rotation' } })
    .replace(/<head>/i, '<head><meta charset="utf-8">' + STUB + FONT_SPY);
  // 日本語のファイル名で保存できるよう、UTF-8 の言語設定で開く（言語設定の無い環境では「download」になる）
  const browser = await pw.chromium.launch({ env: Object.assign({}, process.env, { LANG: 'C.UTF-8', LC_ALL: 'C.UTF-8' }) });
  const page = await browser.newPage({ viewport: { width: 1000, height: 900 }, acceptDownloads: true });
  page.on('pageerror', (e) => fails.push('画面のエラー: ' + e.message));
  let fontRequests = 0;
  await page.route(/fonts\.(googleapis|gstatic)\.com/, (r) => { fontRequests++; r.abort(); });
  // 本物と同じく https の画面として開く（file:// だと、保存のファイル名の指定が効かない）
  const URL = 'https://role-input.test/';
  await page.route(URL, (rt) => rt.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: html }));
  await page.goto(URL);
  const calls = () => page.evaluate(() => window.__calls);
  const val = (id) => page.evaluate((i) => { const e = document.getElementById(i); return e ? e.value : null; }, id);
  // 画像の点の色と、範囲の中の黒っぽい点の数
  const probe = (spec) => page.evaluate((sp) => new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const c = document.createElement('canvas'); c.width = img.naturalWidth; c.height = img.naturalHeight;
      const g = c.getContext('2d'); g.drawImage(img, 0, 0);
      const hex = (x, y) => { const d = g.getImageData(x, y, 1, 1).data; return '#' + [d[0], d[1], d[2]].map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase(); };
      const count = (r, test) => { const d = g.getImageData(r[0], r[1], r[2] - r[0], r[3] - r[1]).data; let n = 0;
        for (let i = 0; i < d.length; i += 4) if (test(d[i], d[i + 1], d[i + 2])) n++; return n; };
      const dark = (r, gg, b) => r < 90 && gg < 90 && b < 90, red = (r, gg, b) => r > 200 && gg < 90 && b < 90;
      resolve({ w: img.naturalWidth, h: img.naturalHeight, pts: Object.fromEntries(Object.entries(sp.pts || {}).map(([k, p]) => [k, hex(p[0], p[1])])),
                dark: Object.fromEntries(Object.entries(sp.dark || {}).map(([k, r]) => [k, count(r, dark)])),
                red: Object.fromEntries(Object.entries(sp.red || {}).map(([k, r]) => [k, count(r, red)])) });
    };
    img.onerror = () => resolve(null);
    img.src = document.getElementById('mpImg').src;
  }), spec);
  const near = (a, b) => { const p = (h) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16)); const x = p(a), y = p(b); return x.every((v, i) => Math.abs(v - y[i]) <= 6); };
  const settle = async () => { await page.waitForTimeout(450); };

  try {
    await page.waitForFunction(() => /^data:image\/png/.test((document.getElementById('mpImg') || {}).src || ''), null, { timeout: 8000 });
  } catch (e) { fails.push('メインプレゼンターの画像ができない（欄・画像が無い）: ' + e.message.split('\n')[0]); }
  await settle();

  // ---- 1) 欄の中身：お名前・カテゴリー（協業したい人を「・」で分ける）----
  const cats = async (i) => Promise.all([0, 1, 2].map((k) => val('mpCat' + i + '_' + k)));
  ck(await val('mpName0') === '見本 一郎' && await val('mpName1') === '試験 二郎', '1) お名前の欄: ' + J([await val('mpName0'), await val('mpName1')]));
  ck(J(await cats(0)) === J(['起業支援', '', '']), '1) 1人目のカテゴリー: ' + J(await cats(0)));
  ck(J(await cats(1)) === J(['司法書士', '税理士', '外国人就労支援事業者']), '1) 2人目のカテゴリー（「・」で分ける）: ' + J(await cats(1)));
  const photoCalls = (await calls()).filter((c) => c[0] === 'getMemberPhotosBase64');
  ck(photoCalls.length === 1 && J(photoCalls[0][1][0]) === J(['見本 一郎', '試験 二郎']), '1) メンバー写真を取りに行かない: ' + J(photoCalls));
  ck(/^data:image\/png/.test(await page.evaluate(() => document.getElementById('rotImg').src || '')), '1) スピーカーローテーションの表の画像ができない');
  // 書体：2枚とも Meiryo UI（太さ 400・700・900 も）。Googleフォントは読みに行かない
  const fonts = [...new Set(await page.evaluate(() => window.__fonts))];
  ck(fonts.length && fonts.every((f) => /^(?:bold |\d{3} )?[\d.]+px "Meiryo UI",/.test(f)) && fonts.some((f) => /^900 /.test(f)),
     '1) 画像の書体が Meiryo UI でない: ' + J(fonts.filter((f) => !/px "Meiryo UI",/.test(f)).slice(0, 3)));
  ck(fontRequests === 0, '1) Googleフォントを読みに行った: ' + fontRequests + '回');

  // ---- 2) 描いた画像 ----
  const DX = 950;
  let r = await probe({
    pts: { panelL: [470, 250], panelR: [1420, 250], footer: [400, 1060], photo1: [265, 580], photo2: [265 + DX, 600], tag: [200, 180], bg: [30, 30] },
    dark: { name1: [485, 280, 878, 350], name2: [485 + DX, 280, 878 + DX, 350], cat1box1: [574, 515, 870, 570], cat1box2: [574, 647, 870, 702],
            cat2box3: [574 + DX, 779, 870 + DX, 834], title: [560, 50, 1360, 110] },
    red: { head1: [485, 412, 880, 445], head2: [485 + DX, 412, 880 + DX, 445], ring1: [480, 500, 560, 585], ring1b: [480, 632, 560, 717] },
  });
  ck(r && r.w === 1920 && r.h === 1080, '2) 画像の大きさが 1920×1080 でない: ' + J(r && [r.w, r.h]));
  if (r) {
    ck(near(r.pts.panelL, '#FFEFF9') && near(r.pts.panelR, '#F2F2F2') && near(r.pts.footer, '#CD1925') && near(r.pts.bg, '#FFFFFF') && near(r.pts.tag, '#FB3431'),
       '2) パネル・帯・PRESENTER の色: ' + J(r.pts));
    ck(near(r.pts.photo1, '#3366CC'), '2) 1人目の写真が枠に入らない: ' + r.pts.photo1);
    ck(near(r.pts.photo2, '#DDE3EA'), '2) 写真の無い方が「写真なし」の枠でない: ' + r.pts.photo2);
    ck(r.dark.title > 300 && r.dark.name1 > 200 && r.dark.name2 > 200, '2) 見出し・お名前の文字が無い: ' + J(r.dark));
    ck(r.dark.cat1box1 > 100 && r.dark.cat1box2 === 0 && r.dark.cat2box3 > 100, '2) カテゴリーの数だけ枠に文字が入っていない: ' + J(r.dark));
    ck(r.red.head1 > 50 && r.red.head2 > 50 && r.red.ring1 > 50 && r.red.ring1b === 0, '2) 見出し「ご紹介して欲しいカテゴリー」・チェックの丸: ' + J(r.red));
  }

  // ---- 3) 欄を直すと描き直す。長いカテゴリーも描ける ----
  await page.fill('#mpCat0_1', '見本の追加カテゴリー');
  await page.fill('#mpCat0_2', '見本のとても長いカテゴリーの名前・二行にしないと入らない説明の続き');
  await settle();
  r = await probe({ dark: { box2: [574, 647, 870, 702], box3top: [574, 768, 870, 806], box3bottom: [574, 808, 870, 846] } });
  ck(r && r.dark.box2 > 100, '3) 直したカテゴリーが描かれない: ' + J(r && r.dark));
  ck(r && r.dark.box3top > 50 && r.dark.box3bottom > 50, '3) 長いカテゴリーを2行で描かない: ' + J(r && r.dark));
  await page.fill('#mpName0', '見本 一郎（代理）');
  await settle();
  r = await probe({ dark: { name1: [485, 280, 878, 350] } });
  ck(r && r.dark.name1 > 200, '3) 長いお名前が描かれない: ' + J(r && r.dark));

  // ---- 4) ご案内する回を替える：その回のお2人。4つ以上のカテゴリーは3つ目にまとめる。カテゴリーの無い方は見出しも出さない ----
  await page.selectOption('#rotFbWeek', '1');
  await settle();
  ck(await val('mpName0') === '架空 三郎' && await val('mpName1') === '仮名 四郎', '4) 回を替えてもお名前が替わらない: ' + J([await val('mpName0'), await val('mpName1')]));
  ck(J(await cats(0)) === J(['不動産賃貸管理', '内装業', 'インテリアデザイン会社・建築設計']), '4) 4つ以上のカテゴリー: ' + J(await cats(0)));
  const pc2 = (await calls()).filter((c) => c[0] === 'getMemberPhotosBase64').pop();
  ck(pc2 && J(pc2[1][0]) === J(['架空 三郎', '仮名 四郎']), '4) 替えた回の方の写真を取りに行かない: ' + J(pc2));
  r = await probe({ red: { head2: [485 + DX, 412, 880 + DX, 445] }, pts: { photo2: [265 + DX, 600] } });
  ck(r && r.red.head2 === 0, '4) カテゴリーの無い方に見出しを出した: ' + J(r && r.red));

  // ---- 5) パソコンの写真を選ぶ ----
  await page.setInputFiles('#mpFile1', { name: 'mihon.png', mimeType: 'image/png', buffer: GREEN });
  await page.waitForTimeout(700);
  r = await probe({ pts: { photo2: [265 + DX, 600] } });
  ck(r && near(r.pts.photo2, '#22AA44'), '5) 選んだ写真が枠に入らない: ' + J(r && r.pts));
  ck(/選んだ写真を使います/.test(await page.evaluate(() => document.getElementById('mpPhotoNote1').innerText)), '5) 写真の知らせ');

  // ---- 6) 保存のファイル名・Driveに保存 ----
  const [dl] = await Promise.all([page.waitForEvent('download', { timeout: 5000 }).catch(() => null), page.evaluate(() => mpImgDownload())]);
  ck(dl && dl.suggestedFilename() === '20311015_メインプレゼンター.png', '6) 保存のファイル名: ' + (dl && dl.suggestedFilename()));
  await page.evaluate(() => mpImgDrive());
  await page.waitForTimeout(300);
  const save = (await calls()).filter((c) => c[0] === 'saveSpeakerRotationImage').pop();
  ck(save && /^iVBORw0KGgo/.test(save[1][0]) && save[1][1] === '20311015_メインプレゼンター.png', '6) Driveに保存の呼び出し: ' + J(save && [String(save[1][0]).slice(0, 12), save[1][1]]));
  ck(/03_生成物/.test(await page.evaluate(() => document.getElementById('mpImgMsg').innerText)), '6) Driveに保存の知らせ');

  await browser.close();
  if (fails.length) {
    console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
    fails.forEach((f) => console.log('  - ' + f));
    process.exit(1);
  }
  console.log('メインプレゼンターの画像: 検査 ' + checks + ' 件 OK: お名前・カテゴリー（協業したい人を分ける）・写真（メンバー写真／選んだ写真／写真なし）・'
    + '1920×1080 の色と文字・直すと描き直す・回を替える・保存とDriveのファイル名');
})().catch((e) => { console.error(e); process.exit(1); });

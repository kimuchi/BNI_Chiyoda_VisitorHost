// メンバーブックのPDFを編集画面の中で作る（memberbook_pdf.html）ことと、ドライブのメンバーブック（メールで送るPDF）を
// 差し替える（memberbook_srv.js の saveMemberBookPdfToDrive）ことを確かめる。名簿・写真は作り物。
//
//   node tools/check_memberbook_pdf.js [Googleフォントを写したディレクトリ]
//   （css.txt と gstatic/ のディレクトリを渡すと、書体（Noto Serif JP）を埋め込むところも確かめる。渡さなければパソコンの書体で描く）
//
// 確かめること
//   1. サーバー：登録してあるファイルの中身だけを差し替える（URL・名前はそのまま）。ゴミ箱にあれば戻してから。
//      開けない・差し替えられないときは新しく作らない（URLを変えない）。「作り直す」ときだけ新しく作る。
//      まだ無いときは作って登録する（リンクを知っている全員が見られるように）。ほかの方が更新中・空のデータ。
//      「メンバーブック(PDF)の更新」（アップロード）も、差し替えた日時を残す
//   2. 画面の中で作ったPDF：PDFとして読める（相互参照表・ページ・A4）。ページの画像（JPEG）の大きさ。
//      見えない文字（検索・コピー用）：お名前・会社名・表紙の題が入る。枠で切って見えていない文字は入れない。
//      見た目がふつうに開いた冊子（印刷と同じ）と同じ。写真は縮めてから使う。大きくなりすぎたら小さくして作り直す
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const pw = (() => { try { return require('playwright'); } catch (e) { return require('/opt/node22/lib/node_modules/playwright'); } })();

const ROOT = path.join(__dirname, '..');
const FONT_DIR = process.argv[2] || '';
const fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const J = (x) => JSON.stringify(x);

// ===================== 1. サーバー（ドライブの差し替え）=====================
function server(opts) {
  const o = opts || {};
  const props = Object.assign({}, o.props || {});
  const files = Object.assign({}, o.files || {});   // id → { trashed, content, name, shared }
  const log = [];
  let seq = 0;
  const box = {
    console: { log() {}, warn() {}, error() {} },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: (k) => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = String(v); } }) },
    LockService: { getScriptLock: () => ({ tryLock: () => !o.busy, releaseLock() {} }) },
    Utilities: {
      base64Decode: (s) => Array.from(Buffer.from(s, 'base64')).map((b) => (b > 127 ? b - 256 : b)),
      newBlob: (bytes, type, name) => {
        let nm = name;
        return { _bytes: Buffer.from(bytes.map((b) => b & 255)), getContentType: () => type, getName: () => nm, setName(n) { nm = n; return this; } };
      },
    },
    Drive: { Files: {
      get: (id) => {
        log.push(['get', id]);
        if (o.denied && o.denied.includes(id)) throw new Error('The user does not have sufficient permissions for file ' + id + '.');
        if (!files[id]) throw new Error('File not found: ' + id + '.');
        return { id, trashed: !!files[id].trashed };
      },
      update: (res, id, blob) => {
        log.push(['update', id, J(res), blob ? blob._bytes.toString('latin1').slice(0, 8) : null]);
        if (o.readOnly && o.readOnly.includes(id)) throw new Error('The user does not have sufficient permissions for file ' + id + '.');
        if (!files[id]) throw new Error('File not found: ' + id + '.');
        if (res && res.trashed === false) files[id].trashed = false;
        if (blob) files[id].content = blob._bytes;
        return { id };
      },
      create: (res, blob) => {
        const id = 'new' + (++seq);
        log.push(['create', id, res.name]);
        files[id] = { name: res.name, content: blob._bytes, trashed: false };
        return { id };
      },
    } },
    DriveApp: {
      Access: { ANYONE_WITH_LINK: 'ANYONE_WITH_LINK' }, Permission: { VIEW: 'VIEW' },
      getFileById: (id) => ({
        getUrl: () => 'https://drive.test/file/d/' + id + '/view',
        setSharing: (a, p) => { if (o.noShare) throw new Error('共有は組織の設定で禁止'); log.push(['share', id, a, p]); files[id].shared = true; },
      }),
    },
  };
  vm.createContext(box);
  for (const f of ['コード.js', 'assets.js', 'memberbook_srv.js']) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), box, { filename: f });
  return { box, props, files, log };
}
const PDF64 = Buffer.from('%PDF-1.4 見本のPDF').toString('base64');
{
  // まだ無いとき：新しく作って登録（リンク共有）
  let S = server();
  let r = S.box.saveMemberBookPdfToDrive(PDF64, 'MemberBook.pdf');
  ck(r.ok && r.created && S.props.MEMBER_BOOK_ID === 'new1' && S.props.MEMBER_BOOK_URL === 'https://drive.test/file/d/new1/view' && r.url === S.props.MEMBER_BOOK_URL,
     'まだ無いとき: ' + J(r) + ' ' + J(S.props));
  ck(S.files.new1.shared && S.files.new1.content.toString('latin1').startsWith('%PDF') && /新しく作って登録/.test(r.message), 'まだ無いとき（共有・中身・知らせ）: ' + J(S.log));
  ck(!!S.props.MEMBER_BOOK_UPDATED && /^\d{4}-\d{2}-\d{2}T/.test(S.props.MEMBER_BOOK_UPDATED), '差し替えた日時が残らない: ' + S.props.MEMBER_BOOK_UPDATED);
  ck(r.downloadUrl === 'https://drive.google.com/uc?export=download&id=new1', 'ダウンロードのURL: ' + r.downloadUrl);

  // 登録してあるとき：中身だけ差し替える（新しく作らない・URLそのまま）
  const reg = { MEMBER_BOOK_ID: 'mb', MEMBER_BOOK_URL: 'https://drive.test/file/d/mb/view' };
  S = server({ props: reg, files: { mb: { name: '配布用.pdf', content: Buffer.from('old'), trashed: false } } });
  r = S.box.saveMemberBookPdfToDrive(PDF64, 'MemberBook.pdf');
  ck(r.ok && !r.created && r.url === reg.MEMBER_BOOK_URL && S.props.MEMBER_BOOK_ID === 'mb' && S.props.MEMBER_BOOK_URL === reg.MEMBER_BOOK_URL,
     '登録してあるとき: ' + J(r) + ' ' + J(S.props));
  ck(!S.log.some((x) => x[0] === 'create') && S.files.mb.content.toString('latin1').startsWith('%PDF') && S.files.mb.name === '配布用.pdf',
     '中身だけ差し替えていない（新しく作った・名前が変わった）: ' + J(S.log));
  ck(S.log.filter((x) => x[0] === 'update').every((x) => x[2] === '{}'), '差し替えで名前などを書き換えた: ' + J(S.log));
  ck(/URLはそのまま/.test(r.message) && !!S.props.MEMBER_BOOK_UPDATED, '差し替えの知らせ・日時: ' + r.message);

  // ゴミ箱にあるとき：戻してから差し替える
  S = server({ props: reg, files: { mb: { content: Buffer.from('old'), trashed: true } } });
  r = S.box.saveMemberBookPdfToDrive(PDF64);
  ck(r.ok && !r.created && S.files.mb.trashed === false && S.files.mb.content.toString('latin1').startsWith('%PDF'), 'ゴミ箱にあるとき: ' + J(S.log));

  // 開けない（削除した）・編集の権限が無いとき：新しく作らない。登録（URL）はそのまま。作り直せることを知らせる
  for (const [label, o] of [['削除した', { files: {} }], ['見る権限も無い', { files: { mb: {} }, denied: ['mb'] }], ['編集の権限が無い', { files: { mb: {} }, readOnly: ['mb'] }]]) {
    S = server(Object.assign({ props: reg }, o));
    r = S.box.saveMemberBookPdfToDrive(PDF64);
    ck(!r.ok && r.canRecreate && !S.log.some((x) => x[0] === 'create') && S.props.MEMBER_BOOK_ID === 'mb' && S.props.MEMBER_BOOK_URL === reg.MEMBER_BOOK_URL
       && !S.props.MEMBER_BOOK_UPDATED, `${label}: 新しく作った・登録が変わった: ` + J(r) + ' ' + J(S.log));
    ck(/差し替えられませんでした/.test(r.message) && /新しく作り直す/.test(r.message), `${label}: 知らせ: ` + r.message);
  }
  // 「新しく作り直す」：新しく作って、登録を替える
  S = server({ props: reg, files: {} });
  r = S.box.saveMemberBookPdfToDrive(PDF64, 'MemberBook.pdf', true);
  ck(r.ok && r.created && S.props.MEMBER_BOOK_ID === 'new1' && S.props.MEMBER_BOOK_URL !== reg.MEMBER_BOOK_URL && /作り直して/.test(r.message), '作り直す: ' + J(r));
  // 共有を設定できない（組織の設定）：作ったことは知らせて、共有を確かめるよう添える
  S = server({ noShare: true });
  r = S.box.saveMemberBookPdfToDrive(PDF64);
  ck(r.ok && r.created && /共有を設定できませんでした/.test(r.message), '共有できないとき: ' + J(r));
  // ほかの方が更新中・空のデータ
  S = server({ props: reg, files: { mb: {} }, busy: true });
  r = S.box.saveMemberBookPdfToDrive(PDF64);
  ck(!r.ok && /更新中/.test(r.message) && !S.log.length, 'ほかの方が更新中: ' + J(r));
  S = server({ props: reg, files: { mb: {} } });
  r = S.box.saveMemberBookPdfToDrive('');
  ck(!r.ok && !S.log.length, '空のデータ: ' + J(r));
  // 編集画面に出す様子
  S = server({ props: Object.assign({ MEMBER_BOOK_UPDATED: '2026-09-28T06:00:00.000Z' }, reg) });
  ck(J(S.box.memberBookDriveInfo_()) === J({ id: 'mb', url: reg.MEMBER_BOOK_URL, updated: '2026-09-28T06:00:00.000Z' }), 'ドライブの様子: ' + J(S.box.memberBookDriveInfo_()));
  // 「メンバーブック(PDF)の更新」（アップロード）：差し替えも、はじめて作るときも日時を残す
  S = server({ props: reg, files: { mb: {} } });
  let u = S.box.uploadMemberBookBlob_(S.box.Utilities.newBlob(Array.from(Buffer.from('%PDF-1.4')), 'application/pdf', 'a.pdf'));
  ck(u.msg === '更新しました。' && u.url === reg.MEMBER_BOOK_URL && !!S.props.MEMBER_BOOK_UPDATED, 'アップロードで差し替え: ' + J(u));
  S = server();
  u = S.box.uploadMemberBookBlob_(S.box.Utilities.newBlob(Array.from(Buffer.from('%PDF-1.4')), 'application/pdf', 'a.pdf'));
  ck(u.msg === '新規登録しました。' && S.props.MEMBER_BOOK_ID === 'new1' && S.files.new1.shared && !!S.props.MEMBER_BOOK_UPDATED, 'アップロードで新しく登録: ' + J(u));
}

// ===================== 2. 画面の中で作るPDF =====================
// PDF を読む（相互参照表の位置・オブジェクト・ストリーム）
function readPdf(buf) {
  const s = buf.toString('latin1'), out = { errors: [], objs: {} };
  if (!s.startsWith('%PDF-1.')) out.errors.push('先頭が %PDF- でない');
  const m = s.match(/startxref\n(\d+)\n%%EOF\n?$/);
  if (!m) { out.errors.push('startxref が無い'); return out; }
  const x = +m[1], hm = s.slice(x).match(/^xref\n0 (\d+)\n0000000000 65535 f \n/);
  if (!hm) { out.errors.push('相互参照表が無い'); return out; }
  const n = +hm[1];
  for (let i = 1; i < n; i++) {
    const e = s.substr(x + hm[0].length + (i - 1) * 20, 20);
    if (!/^\d{10} 00000 n \n$/.test(e)) { out.errors.push(i + ' 番の相互参照の行: ' + J(e)); continue; }
    const off = parseInt(e, 10), head = i + ' 0 obj\n';
    if (s.substr(off, head.length) !== head) { out.errors.push(i + ' 番の位置が違う'); continue; }
    const at = off + head.length, dm = s.slice(at).match(/^(<<[\s\S]*?>>)\n(stream\n)?/);
    if (!dm) { out.errors.push(i + ' 番の中身が読めない'); continue; }
    const obj = { dict: dm[1] };
    if (dm[2]) {
      const len = +(dm[1].match(/\/Length (\d+)/) || [])[1], from = at + dm[0].length;
      obj.data = buf.subarray(from, from + len);
      if (s.substr(from + len, 18) !== '\nendstream\nendobj\n') out.errors.push(i + ' 番のストリームの長さが違う');
    } else if (s.substr(at + dm[1].length, 8) !== '\nendobj\n') out.errors.push(i + ' 番の終わりが違う');
    out.objs[i] = obj;
  }
  const tr = s.slice(x).match(/trailer\n<<\/Size (\d+)\/Root (\d+) 0 R\/Info (\d+) 0 R>>/);
  if (!tr || +tr[1] !== n) out.errors.push('trailer: ' + (tr && tr[0]));
  out.root = tr && out.objs[+tr[2]]; out.info = tr && out.objs[+tr[3]];
  return out;
}
const ucs2 = (h) => { let t = ''; for (let i = 0; i + 4 <= h.length; i += 4) t += String.fromCharCode(parseInt(h.substr(i, 4), 16)); return t; };
const jpegSize = (b) => {
  if (b[0] !== 0xFF || b[1] !== 0xD8) return null;
  for (let i = 2; i + 9 < b.length;) {
    if (b[i] !== 0xFF) return null;
    const mk = b[i + 1];
    if ((mk >= 0xC0 && mk <= 0xC3) || (mk >= 0xC5 && mk <= 0xC7) || (mk >= 0xC9 && mk <= 0xCB) || (mk >= 0xCD && mk <= 0xCF)) return { h: b.readUInt16BE(i + 5), w: b.readUInt16BE(i + 7) };
    i += 2 + b.readUInt16BE(i + 2);
  }
  return null;
};
// PDF のページ → { box, image: { w, h, jpeg }, text: 行の文字の並び, lines }
function pagesOf(pdf) {
  const pagesObj = pdf.objs[+((pdf.root && pdf.root.dict.match(/\/Pages (\d+) 0 R/)) || [])[1]];
  const kids = pagesObj ? (pagesObj.dict.match(/\/Kids\[([^\]]*)\]/)[1].match(/\d+(?= 0 R)/g) || []).map(Number) : [];
  return kids.map((k) => {
    const d = pdf.objs[k].dict, im = pdf.objs[+d.match(/\/Im0 (\d+) 0 R/)[1]], ct = pdf.objs[+d.match(/\/Contents (\d+) 0 R/)[1]];
    const c = ct.data.toString('latin1'), lines = [];
    const re = /([\d.-]+) 0 0 ([\d.-]+) ([\d.-]+) ([\d.-]+) Tm <([0-9A-F]*)> Tj/g;
    let mm;
    while ((mm = re.exec(c)) !== null) lines.push({ sx: +mm[1], size: +mm[2], x: +mm[3], y: +mm[4], s: ucs2(mm[5]) });
    return { dict: d, box: (d.match(/\/MediaBox\[([^\]]*)\]/) || [])[1], content: c, lines, text: lines.map((l) => l.s).join('\n'),
             image: { dict: im.dict, jpeg: im.data, size: jpegSize(im.data) } };
  });
}

// 作り物の冊子（22名：表紙＋2ページ）
const CATS = [['企業サポート', '#FFDE58'], ['研修・教育', '#C1FF72'], ['不動産関連', '#37B5FF'], ['建築・住まい', '#AAB5D9']].map(([key, bg]) => ({ key, label: key, bg }));
const LONG = '見本の長い文です。';
const MEMBERS = [];
for (let i = 0; i < 22; i++) {
  const k = i % 5;
  MEMBERS.push({ no: String(i + 1), name: ['見本 一郎', '見本 花子', '見本 太郎左衛門', '見本 三郎', 'Mihon Taro'][k] + (i >= 5 ? i : ''), cat: CATS[i % 4].key,
    title: ['税理士', '不動産売買仲介（相続・事業承継）', '内装工事', '生命保険（法人）', '行政書士（建設業許可・相続）'][k],
    company: ['見本株式会社', '一般社団法人見本フォーラム', '(株)見本', '見本生命保険株式会社 東京本店', 'MIHON Holdings Co., Ltd.'][k],
    position: ['代表取締役', '', '取締役 兼 営業本部長', '支店長', ''][k], role: ['ビジターホスト', 'Webチーム', '', 'プレジデント', ''][k],
    comment: ['見本の一言です', LONG.repeat(2), LONG.repeat(3), '切れて見えない始めの文。' + LONG.repeat(40) + '切れて見えない終わりの文', ''][k],
    refer: ['税理士・弁護士', '新しく店舗を出す飲食店のオーナー', LONG.repeat(3), LONG.repeat(9), '司法書士'][k],
    collab: ['司法書士', '保険・不動産・士業', LONG.repeat(2), LONG.repeat(9), ''][k] });
}
const COVER = { title: 'BNI 見本 chapter Member Book', term: '24期', pname: '見本 会長', prole: '見本チャプター\n第24期プレジデント',
  ptext: 'あいさつの文です。'.repeat(40), philosophyTitle: 'BNIの理念', philosophy: '理念の本文です。'.repeat(6), aboutTitle: 'チャプターとは？',
  about: '紹介の本文です。'.repeat(5), benefitsTitle: 'メリット', benefits: '1.メリット　2.メリット\n3.メリット',
  scheduleFrom: '7:15', scheduleTo: '9:15', schedule: Array.from({ length: 12 }, (_, i) => '項目' + (i + 1)).join('\n'),
  termsTitle: 'BNIの用語の説明', terms: '①チャプター\n説明の文です。\n②リファーラル\n説明の文です。' };

async function routeFonts(ctx) {
  if (!FONT_DIR) { await ctx.route(/fonts\.(googleapis|gstatic)\.com/, (r) => r.abort()); return; }
  const css = fs.readFileSync(path.join(FONT_DIR, 'css.txt'), 'utf8'), hdr = { 'access-control-allow-origin': '*' };
  await ctx.route('https://fonts.googleapis.com/**', (r) => r.fulfill({ status: 200, contentType: 'text/css', body: css, headers: hdr }));
  await ctx.route('https://fonts.gstatic.com/**', (r) => {
    const f = path.join(FONT_DIR, 'gstatic', r.request().url().replace('https://fonts.gstatic.com/', '').replace(/\//g, '_'));
    if (!fs.existsSync(f)) return r.fulfill({ status: 404, body: '' });
    r.fulfill({ status: 200, contentType: 'font/woff2', body: fs.readFileSync(f), headers: hdr });
  });
}
const strip = (f) => fs.readFileSync(path.join(ROOT, f), 'utf8').replace(/^<script>|<\/script>\s*$/g, '');

(async () => {
  // 文字の色つきのにじみ（LCD）は、ふつうの画面にだけ付き、SVG の画像には付かない。比べるときに差にならないよう、はじめから外す
  const browser = await pw.chromium.launch({ args: ['--disable-lcd-text'] });
  const ctx = await browser.newContext();
  await routeFonts(ctx);
  const page = await ctx.newPage();
  await page.route('https://mb.test/', (r) => r.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: '<!DOCTYPE html><html><body></body></html>' }));
  await page.goto('https://mb.test/');
  const R = await page.evaluate(async ({ render, pdfjs, members, cats, cover }) => {
    window.members = members; window.cats = cats; window.cover = cover; window.photosB64 = {};
    window.catOf = (k) => cats.find((c) => c.key === k) || null;
    (0, eval)(render); (0, eval)(pdfjs);
    // 写真：大きなもの（元の写真の代わり）と、小さなもの
    const photo = (hue, w, h) => { const c = document.createElement('canvas'); c.width = w; c.height = h; const g = c.getContext('2d');
      g.fillStyle = 'hsl(' + hue + ',60%,55%)'; g.fillRect(0, 0, w, h); g.fillStyle = '#fff'; g.beginPath(); g.arc(w / 2, h / 3, w / 4, 0, 7); g.fill();
      return c.toDataURL('image/jpeg', 0.92); };
    const photos = {};
    members.forEach((m, i) => { if (i % 3 !== 2) photos[m.name] = photo(i * 37, i % 2 ? 2400 : 300, i % 2 ? 3200 : 400); });
    photos[cover.pname] = photo(200, 3000, 2000);
    const size = (d) => new Promise((ok) => { const im = new Image(); im.onload = () => ok([im.naturalWidth, im.naturalHeight]); im.src = d; });
    const small = await mbpSmallPhotos(photos);
    const sizes = {};
    for (const n of Object.keys(small)) sizes[n] = await size(small[n]);
    const again = await mbpSmallPhotos(photos);
    const steps = [];
    const html = bookHtml({ photos: small });
    const t0 = performance.now();
    const pdf = await mbpBuildPdf(html, cover.title, (t) => steps.push(t));
    const ms = performance.now() - t0;
    // 書体の埋め込み（Googleフォントを読めたときだけ）
    const fr = await mbpOpenBook(html), loaded = mbpFontsLoaded(fr.contentDocument);
    const css = loaded ? await mbpFontCss(fr.contentDocument, fr.contentDocument.querySelectorAll('.page')[1].textContent, {}) : '';
    fr.parentNode.removeChild(fr);
    // 枠で切ってある文字は入れない（高さ20pxの枠に3行ぶんの文字）
    const tdoc = document.implementation.createHTMLDocument('t');
    tdoc.body.innerHTML = '<div class="page" style="position:relative;width:793px;height:1122px;">'
      + '<div style="position:absolute;left:10px;top:10px;width:160px;height:20px;overflow:hidden;font-size:16px;line-height:20px;">'
      + 'あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめも</div>'
      + '<div style="position:absolute;left:10px;top:100px;">見える文字<br>二行め</div></div>';
    const holder = document.createElement('div'); holder.appendChild(document.adoptNode(tdoc.body.firstChild)); document.body.appendChild(holder);
    const clipLines = mbpPageText(holder.querySelector('.page'));
    holder.remove();
    // 大きくなりすぎたとき：小さくして作り直す
    const keep = MBP_MAX_BYTES; MBP_MAX_BYTES = 400 * 1024;
    const steps2 = [];
    const pdf2 = await mbpBuildPdf(html, cover.title, (t) => steps2.push(t));
    MBP_MAX_BYTES = keep;
    return { pdf, ms, steps, sizes, sameAgain: Object.keys(small).every((n) => again[n] === small[n]), small: Object.keys(small).length,
             unchangedSmall: small[members[4].name] === photos[members[4].name], loaded, css: css.length, cssRules: (css.match(/@font-face/g) || []).length,
             clipLines, pdf2: { base64: pdf2.base64, bytes: pdf2.bytes }, steps2, html };
  }, { render: strip('memberbook_render.html'), pdfjs: strip('memberbook_pdf.html'), members: MEMBERS, cats: CATS, cover: COVER });

  // 写真を縮める（長い辺 480px まで。小さな写真はそのまま。2回目は覚えたものを使う）
  ck(R.small === 16, '写真の数: ' + R.small);
  ck(Object.values(R.sizes).every(([w, h]) => Math.max(w, h) <= 480), '写真が縮んでいない: ' + J(R.sizes));
  ck(J(R.sizes[MEMBERS[1].name]) === J([360, 480]) && J(R.sizes[COVER.pname]) === J([480, 320]), '縮めた写真の大きさ（縦横の比はそのまま）: ' + J(R.sizes));
  ck(R.unchangedSmall && R.sameAgain, '小さな写真はそのまま・2回目は覚えたもの: ' + J([R.unchangedSmall, R.sameAgain]));

  // PDF の作り
  const buf = Buffer.from(R.pdf.base64, 'base64'), pdf = readPdf(buf), pages = pagesOf(pdf);
  ck(!pdf.errors.length, 'PDFの作りが正しくない: ' + J(pdf.errors.slice(0, 5)));
  ck(buf.length === R.pdf.bytes && R.pdf.pages === 3 && pages.length === 3, 'ページ数: ' + R.pdf.pages + ' / ' + pages.length);
  ck(pages.every((p) => p.box === '0 0 595.28 841.89'), 'ページの大きさ（A4）: ' + J(pages.map((p) => p.box)));
  ck(pages.every((p) => p.image.size && Math.abs(p.image.size.w - 1984) <= 2 && Math.abs(p.image.size.h - 2806) <= 3
     && p.image.dict.includes('/Width ' + p.image.size.w + '/') && p.image.dict.includes('/Height ' + p.image.size.h + '/')
     && /\/Filter\/DCTDecode/.test(p.image.dict)), 'ページの画像（JPEG・約240dpi）: ' + J(pages.map((p) => [p.image.size, p.image.dict])));
  ck(pages.every((p) => /^q 595\.28 0 0 841\.89 0 0 cm \/Im0 Do Q\n/.test(p.content) && /BT 3 Tr \/F1 1 Tf\n/.test(p.content)), '画像を全面に・文字は描かない（3 Tr）');
  const font = Object.values(pdf.objs).find((o) => /\/Subtype\/Type0/.test(o.dict));
  ck(font && /\/Encoding\/UniJIS-UCS2-H/.test(font.dict) && /\/ToUnicode \d+ 0 R/.test(font.dict), '文字の書体（UniJIS-UCS2-H・ToUnicode）: ' + (font && font.dict));
  const info = pdf.info ? pdf.info.dict : '', tt = (info.match(/\/Title<FEFF([0-9A-F]*)>/) || [])[1];
  ck(tt && ucs2(tt) === COVER.title, 'PDFの題: ' + info);
  ck(R.pdf.bytes < 6 * 1024 * 1024, 'PDFが大きすぎる: ' + R.pdf.bytes);
  ck(R.steps.length === 3 && /3\/3ページ/.test(R.steps[2]), '進み具合の知らせ: ' + J(R.steps));

  // 見えない文字（検索・コピー用）
  const flat = (t) => t.replace(/\s/g, '');
  ck(flat(pages[0].text).includes(flat(COVER.title)) && flat(pages[0].text).includes('タイムスケジュール') && flat(pages[0].text).includes('見本会長'),
     '表紙の文字: ' + pages[0].text.slice(0, 120));
  MEMBERS.forEach((m, i) => {
    const p = pages[1 + Math.floor(i / 18)], t = flat(p.text);
    ck(t.includes(flat(m.name)) && t.includes(flat(m.company)), `${m.no} ${m.name}: お名前・会社名が文字に無い`);
  });
  ck(pages[1].lines.some((l) => l.s === '見本 花子'), 'お名前が1行の文字として入っていない（検索できない）');
  // とても長い一言（入りきらず、枠で上下が切れる）：見えている行だけ入る
  ck(!flat(pages[1].text).includes('切れて見えない終わりの文') && !flat(pages[1].text).includes('切れて見えない始めの文')
     && flat(pages[1].text).includes('見本の長い文です。見本の長い文です。'), '枠で切ってある文字も入っている（または一言が無い）');
  if (process.env.DUMP) fs.writeFileSync(path.join(process.env.DUMP, 'text2.txt'), pages[1].lines.map((l) => J(l)).join('\n'));
  ck(J(R.clipLines.map((l) => l.s)) === J(['あいうえおかきくけこ', '見える文字', '二行め']), '枠で切ってある文字を入れない: ' + J(R.clipLines.map((l) => l.s)));
  ck(R.clipLines.every((l) => l.h > 10 && l.h < 20 && l.w > 20), '文字の位置と大きさ（pt）: ' + J(R.clipLines));
  // 文字の位置：お名前の行が、画像の中のお名前のところ（カードの中）にある
  const nm = pages[1].lines.find((l) => l.s === '見本 一郎');
  ck(nm && nm.x > 10 && nm.x < 150 && nm.y > 700 && nm.y < 800 && nm.size > 9 && nm.size < 18, 'お名前の文字の位置: ' + J(nm));

  // 書体の埋め込み
  if (FONT_DIR) ck(R.loaded && R.cssRules > 0 && R.css > 10000, '書体を埋め込んでいない: ' + J([R.loaded, R.cssRules, R.css]));
  else ck(!R.loaded && R.css === 0, '書体を読めないのに埋め込もうとした: ' + J([R.loaded, R.css]));

  // 大きくなりすぎたとき：小さくして作り直す
  const pages2 = pagesOf(readPdf(Buffer.from(R.pdf2.base64, 'base64')));
  ck(R.steps2.some((t) => /小さくして作り直し/.test(t)) && pages2.length === 3 && pages2[1].image.size.w < 1984 && R.pdf2.bytes < R.pdf.bytes,
     '大きすぎるときに作り直していない: ' + J([R.steps2.slice(-1), R.pdf2.bytes, pages2.map((p) => p.image.size)]));

  // 見た目：ふつうに開いた冊子（印刷と同じ）の画面と比べる（1/5 に縮めて、色の差の平均と、大きく違う点の割合）
  const ctx2 = await browser.newContext({ deviceScaleFactor: 2.5 });
  await routeFonts(ctx2);
  const p2 = await ctx2.newPage();
  await p2.route('https://mb.test/book', (r) => r.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: R.html }));
  await p2.goto('https://mb.test/book');
  await p2.waitForFunction(() => document.documentElement.getAttribute('data-fitted') === '1', null, { timeout: 30000 });
  await p2.addStyleTag({ content: '.page{margin:0!important;box-shadow:none!important;} body{padding:0!important;background:#fff!important;}' });
  const els = await p2.$$('.page');
  for (let i = 0; i < els.length; i++) {
    // そのページだけを出して撮る（ページの上端が小数の位置にあると、文字の置き方が1pxほどずれて比べにくい）
    await p2.evaluate((k) => document.querySelectorAll('.page').forEach((p, j) => { p.style.display = j === k ? '' : 'none'; }), i);
    const shotBuf = await els[i].screenshot(), shot = shotBuf.toString('base64'), jpg = pages[i].image.jpeg.toString('base64');
    if (process.env.DUMP) { fs.writeFileSync(path.join(process.env.DUMP, `shot${i + 1}.png`), shotBuf); fs.writeFileSync(path.join(process.env.DUMP, `pdf${i + 1}.jpg`), pages[i].image.jpeg); }
    const d = await p2.evaluate(async ({ a, b }) => {
      const load = (src) => new Promise((ok) => { const im = new Image(); im.onload = () => ok(im); im.src = src; });
      const [ia, ib] = await Promise.all([load('data:image/png;base64,' + a), load('data:image/jpeg;base64,' + b)]);
      // 1/10 に縮めて（JPEG の文字の輪郭のにじみを消す）、明るさの差が 24 を超える点の割合。
      // 書体を埋め込まずにパソコンの書体で描くと 2〜11%、正しく描けていれば 0.4% 以下
      const W = 198, H = 280, px = (im) => {                // 10×10 の点ごとの明るさの平均
        const c = document.createElement('canvas'); c.width = W * 10; c.height = H * 10;
        const g = c.getContext('2d'); g.drawImage(im, 0, 0);
        const d = g.getImageData(0, 0, W * 10, H * 10).data, out = new Float32Array(W * H);
        for (let y = 0; y < H * 10; y++) for (let x = 0; x < W * 10; x++) {
          const k = (y * W * 10 + x) * 4;
          out[Math.floor(y / 10) * W + Math.floor(x / 10)] += (0.299 * d[k] + 0.587 * d[k + 1] + 0.114 * d[k + 2]) / 100;
        }
        return out;
      };
      const x = px(ia), y = px(ib);
      let big = 0;
      for (let k = 0; k < x.length; k++) if (Math.abs(x[k] - y[k]) > 24) big++;
      return { big: big / (W * H) };
    }, { a: shot, b: jpg });
    // 書体を渡さないとき（パソコンの書体）：この検査の環境には冊子の明朝体が無く、総称の serif になる。総称の書体は SVG の画像の中では
    // 別の書体に当たることがあり、文字が1px足らずずれるので、大きく壊れていない（白いページ・写真が無いなど）ことだけを見る
    const limit = FONT_DIR ? 0.01 : 0.1;
    ck(d.big < limit, `${i + 1}ページめの見た目が、ふつうに開いた冊子と違う: 違う点 ${(d.big * 100).toFixed(2)}%`);
    if (process.env.V) console.log(`  ${i + 1}ページめ: 違う点 ${(d.big * 100).toFixed(3)}%`);
  }
  if (process.env.V) console.log('  PDF: ' + R.pdf.bytes + ' bytes, ' + Math.round(R.ms) + 'ms');
  await browser.close();

  console.log(`メンバーブックのPDF（画面の中で作る・ドライブの差し替え）: 検査 ${checks} 件` + (FONT_DIR ? '（書体を埋め込む）' : '（パソコンの書体）'));
  if (fails.length) {
    console.log(`NG: ${fails.length} 件`);
    fails.forEach((f) => console.log('   ' + f));
    process.exit(1);
  }
  console.log('OK: ドライブのメンバーブックは中身だけ差し替え（URLそのまま・差し替えられないときは新しく作らない）・PDFの作り（A4・画像・検索できる文字・枠で切った文字は入れない）・'
    + '見た目は印刷と同じ・写真を縮める・大きすぎるときは小さくして作り直す');
})().catch((e) => { console.error(e); process.exit(1); });

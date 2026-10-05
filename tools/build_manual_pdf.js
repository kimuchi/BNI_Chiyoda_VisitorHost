// 使い方ガイド（manual.html。もとは MANUAL.md）を、1つのPDFにする。
// 表紙（版・作った日）→ 目次（押すとその章へ飛ぶ）→ 本文（章ごとにページを替える）。PDFのしおり・ページ番号つき。
//
//   node tools/build_manual_pdf.js <出力.pdf> [Googleフォントを写したディレクトリ]
//
// 書体：Googleフォントを写したディレクトリ（css.txt と gstatic/。Noto Sans JP の 400・700）を渡すとそれで組む。
// 渡さなければ、パソコンの日本語の書体（IPAゴシックなど）で組む。写し方は tools/check_memberbook.js と同じ:
//   curl -A "（Chrome の User-Agent）" "https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;700&display=swap" -o css.txt
//   css.txt の https://fonts.gstatic.com/… を gstatic/ に、パスの「/」を「_」にした名前で置く
// Noto Sans JP は太さを変えられる書体（可変フォント）で、そのままだとPDFの中で文字が絵（Type3）として入り、大きくなる。
// tools/instance_webfont.py で太さごとの書体に切り出したもの（同じ置き方）を渡すと、ふつうの書体として入る。
// 図（画面の写真）は JPEG にしてから入れる（そのままだと大きくなるため）。
// 作ったPDFはリポジトリには入れない（マニュアルを直すたびに作り直す）。
const fs = require('fs');
const path = require('path');
const pw = (() => { try { return require('playwright'); } catch (e) { return require('/opt/node22/lib/node_modules/playwright'); } })();

const ROOT = path.join(__dirname, '..');
const OUT = process.argv[2];
const FONT_DIR = process.argv[3] || '';
if (!OUT) { console.error('使い方: node tools/build_manual_pdf.js <出力.pdf> [Googleフォントを写したディレクトリ]'); process.exit(2); }

const version = (fs.readFileSync(path.join(ROOT, 'コード.js'), 'utf8').match(/SYSTEM_VERSION_\s*=\s*'([^']+)'/) || [])[1] || '';
const now = new Date(Date.now() + 9 * 3600 * 1000);                   // 日本時間の日付
const today = now.getUTCFullYear() + '年' + (now.getUTCMonth() + 1) + '月' + now.getUTCDate() + '日';
const FONT = FONT_DIR ? '"Noto Sans JP","IPAPGothic","IPAGothic",sans-serif' : '"IPAPGothic","IPAGothic",sans-serif';

// 印刷用の見た目（画面用の manual.html に足す）
const PRINT_CSS = `
  @page { size: A4; }
  html, body { background: #fff; }
  body { font-family: ${FONT}; font-size: 10.5pt; line-height: 1.7; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  .wrap { padding: 0 !important; }
  .top { display: none !important; }
  .cover { height: 230mm; display: flex; flex-direction: column; justify-content: center; text-align: center; break-after: page; }
  .cover .t { font-size: 30pt; font-weight: 700; color: #0055ff; }
  .cover .s { font-size: 16pt; margin-top: 10mm; color: #333; }
  .cover .v { font-size: 11pt; margin-top: 24mm; color: #666; line-height: 2; }
  .toc { break-after: page; border: none; background: none; padding: 0; margin: 0; }
  .toc b { font-size: 16pt; margin-bottom: 4mm; }
  .toc ol { margin: 0 0 0 10mm; }                                         /* 2けたの番号が欠けないように */
  .toc li { margin: 2.2mm 0; font-size: 11.5pt; }
  h1 { font-size: 19pt; }
  h2 { break-before: page; font-size: 15pt; margin-top: 0; }
  h2, h3, h4 { break-after: avoid; }
  figure.shot, pre, blockquote, tr { break-inside: avoid; }
  figure.shot img { box-shadow: none; max-height: 190mm; width: auto; max-width: 100%; }
  pre { white-space: pre-wrap; word-break: break-all; overflow: visible; }
  code { font-family: inherit; overflow-wrap: anywhere; }
  pre, pre code { font-family: "IPAGothic", monospace; }                  /* 図（├─ など）は桁がそろう書体で */
  td, th { overflow-wrap: anywhere; }
  a { color: inherit; text-decoration: none; }
  .toc a { color: #0f3d91; }
`;
const COVER = '<div class="cover"><div class="t">BNI 名簿システム</div><div class="s">ご利用マニュアル（使い方ガイド）</div>'
  + '<div class="v">' + (version ? '版 ' + version + '<br>' : '') + today + ' 作成</div></div>';
// ページの下：ページ番号（見出しや目次の書体は、ページ番号の欄では使えないので数字だけ）
const FOOTER = '<div style="width:100%;font-size:8pt;color:#888;text-align:center;font-family:sans-serif;"><span class="pageNumber"></span> / <span class="totalPages"></span></div>';

(async () => {
  let html = fs.readFileSync(path.join(ROOT, 'manual.html'), 'utf8');
  html = html.replace('</head>', (FONT_DIR ? '<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;700&display=swap">' : '')
    + '<style>' + PRINT_CSS + '</style></head>')
    .replace('<div class="wrap" id="top">', '<div class="wrap" id="top">' + COVER)
    .replace(/<script>[\s\S]*?<\/script>/, '');                          // 画面の中で目次から飛ぶための処理は要らない
  const browser = await pw.chromium.launch();
  const page = await browser.newPage();
  const URL = 'https://manual.test/';
  await page.route(URL, (r) => r.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: html }));
  if (FONT_DIR) {
    const css = fs.readFileSync(path.join(FONT_DIR, 'css.txt'), 'utf8');
    await page.route('https://fonts.googleapis.com/**', (r) => r.fulfill({ status: 200, contentType: 'text/css', body: css }));
    await page.route('https://fonts.gstatic.com/**', (r) => {
      const f = path.join(FONT_DIR, 'gstatic', r.request().url().replace('https://fonts.gstatic.com/', '').replace(/\//g, '_'));
      if (!fs.existsSync(f)) return r.abort();
      return r.fulfill({ status: 200, contentType: 'font/woff2', headers: { 'Access-Control-Allow-Origin': '*' }, body: fs.readFileSync(f) });
    });
  } else {
    await page.route(/fonts\.(googleapis|gstatic)\.com/, (r) => r.abort());
  }
  await page.goto(URL, { waitUntil: 'load' });
  await page.emulateMedia({ media: 'print' });
  // 本文の文字の書体を読み終わるまで待つ（字の形ごとにファイルが分かれているので、使う字のファイルをすべて）
  const fonts = await page.evaluate(async (useNoto) => {
    if (!useNoto) return null;
    const text = document.body.innerText;
    await Promise.all(['400', '700'].map((w) => document.fonts.load(w + ' 16px "Noto Sans JP"', text)));
    await document.fonts.ready;
    return { loaded: [...document.fonts].filter((f) => f.family.replace(/"/g, '') === 'Noto Sans JP' && f.status === 'loaded').length };
  }, !!FONT_DIR);
  // 図（画面の写真）を JPEG にする（透明なところは白に）
  const shots = await page.evaluate(async () => {
    const imgs = [...document.querySelectorAll('figure.shot img')];
    for (const img of imgs) {
      if (!img.naturalWidth) continue;
      const cv = document.createElement('canvas');
      cv.width = img.naturalWidth; cv.height = img.naturalHeight;
      const g = cv.getContext('2d');
      g.fillStyle = '#FFFFFF'; g.fillRect(0, 0, cv.width, cv.height); g.drawImage(img, 0, 0);
      img.src = cv.toDataURL('image/jpeg', 0.85);
      await img.decode();
    }
    return imgs.length;
  });
  await page.pdf({
    path: OUT, format: 'A4', printBackground: true, outline: true, tagged: true,
    displayHeaderFooter: true, headerTemplate: '<div></div>', footerTemplate: FOOTER,
    margin: { top: '16mm', bottom: '18mm', left: '15mm', right: '15mm' },
  });
  await browser.close();
  console.log('PDF: ' + OUT + '（' + Math.round(fs.statSync(OUT).size / 1024) + 'KB・版 ' + version + '・'
    + (fonts ? 'Noto Sans JP（' + fonts.loaded + 'ファイル）' : 'パソコンの書体') + '・図 ' + shots + '枚）');
})().catch((e) => { console.error(e); process.exit(1); });

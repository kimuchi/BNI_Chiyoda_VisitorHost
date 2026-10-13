// 会社名の改行（company_break.js の CB_）を、メンバーブックとビジター様リストのPDFで使っているかを確かめる。会社名・お名前はすべて架空。
//
//   node tools/check_company_lines_html_pdf.js
//
// 確かめること
//   1. メンバーブック（memberbook_render.html。Chromium で実際に描いて測る）
//      ・会社名は名簿の書き方にそろえる（セルの中の改行・㈱）。1行のまま小さくして入る名前は1行のまま
//      ・入らない名前は、CB_ の切れ目の候補（文字の data-l2）で2行にする。いちばん良い候補が入るなら、それを使う
//        （以前はブラウザに折らせていて「…ソリュー／ションズ株式会社」のように語の途中で折れた）。描いた行も同じ（ブラウザが折り直さない）
//      ・どの候補も入らないとても長い名前は、ブラウザに折らせるが、ー・小さいかなで行を始めない（line-break:strict）
//      ・冊子のHTML（配布用・印刷・PDF）には CB_ が無い。それでも属性の候補で同じように収める。何度収め直しても同じ
//      ・候補は、サーバーの CB_ と同じ（画面に渡す決め方が同じ）。編集画面は company_break を組版より先に読む。
//        CB_ が無いところ（読み込み前など）でも冊子は作れる（以前と同じくブラウザに折らせる）
//   2. ビジター様リストのPDF（コード.js の createFinalSheet）
//      ・印刷用シートの会社名は、列の内側の幅（約115pt）・10pt で CB_ が決めた切れ目に改行を入れる（4行まで・小さくしない）
//      ・名簿のシート（再編集で読み戻す）は元の書き方のまま。カテゴリーなどほかの欄は変えない
process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeEnv } = require('./lib_sheet_fake');
const pw = (() => { try { return require('playwright'); } catch (e) { return require('/opt/node22/lib/node_modules/playwright'); } })();

const ROOT = path.join(__dirname, '..');
const read = (f) => fs.readFileSync(path.join(ROOT, f), 'utf8');
const fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const J = (x) => JSON.stringify(x);
const flat = (s) => String(s).replace(/[ 　]/g, '');

const S = { console };
vm.createContext(S);
vm.runInContext(read('company_break.js'), S, { filename: 'company_break.js' });
const CB = S.CB_;
// 行の頭に来てはいけない字（company_break.js の NO_START と同じ）
const NO_START = /^[ー‐゠－~〜・･、。，．,.：:；;！？!?）)」』】〕〉》］\]｝}々ゝゞヽヾぁぃぅぇぉっゃゅょゎゕゖァィゥェォッャュョヮヵヶ％%]/;
const CORPS = ['株式会社', '合同会社', '一般社団法人', '特定非営利活動法人'];
const brokenCorp = (lines) => {
  const s = lines.join('\n');
  return CORPS.some((w) => { for (let k = 1; k < w.length; k++) if (s.includes(w.slice(0, k) + '\n' + w.slice(k))) return true; return false; });
};

// ===================== 1. メンバーブック =====================
// [会社名, どうなるはずか] one … 1行 / best … いちばん良い候補で2行 / two … どれかの候補で2行（英字の幅は書体による）/ css … ブラウザに折らせる
const BOOK = [
  ['見本商事', 'one'],
  ['一般社団法人見本人材育成フォーラム', 'one'],
  ['㈱見本\nエンジニアリングサービス', 'one'],                              // 名簿のセルの中の改行・㈱
  ['株式会社見本不動産コンサルティングオフィス', 'best'],
  ['見本マーケティングソリューションズ株式会社', 'best'],
  ['一般社団法人日本見本まちづくり推進協会東京支部', 'best'],
  ['特定非営利活動法人見本子ども未来支援センター', 'best'],
  ['株式会社見本ビューティー＆ウェルネスラボラトリー', 'best'],
  ['見本建築設計事務所一級建築士事務所 東京事務所', 'best'],
  ['Mihon Global Consulting Partners株式会社 東京支社', 'two'],
  ['見本&Sample "Quote" <Lab>合同会社 東京第一営業部', 'best'],               // 属性に入れるときの & " < >
  ['Mihon Shiken Kakuu ミホンキャッシュフローショッピングモールウェーブシャトー', 'css'],   // どの候補も1行が長すぎる
];
const CATS = [{ key: '企業サポート', label: '企業サポート', bg: '#FFDE58' }];
const MEMBERS = BOOK.map(([company], i) => ({ no: String(i + 1), name: '見本 ' + (i + 1) + '郎', cat: '企業サポート', title: '見本の業種', company,
  position: '', role: '', comment: '見本の一言です', refer: '税理士', collab: '司法書士' }));
const COVER = { title: 'BNI 見本 chapter Member Book', term: '24期', pname: '見本 会長', prole: '見本チャプター', ptext: 'あいさつ',
  philosophyTitle: '理念', philosophy: '理念', aboutTitle: '紹介', about: '紹介', benefitsTitle: 'メリット', benefits: 'メリット',
  scheduleFrom: '7:15', scheduleTo: '9:15', schedule: '開会', termsTitle: '用語', terms: '用語\n説明' };

// 編集画面：company_break を、組版（memberbook_render）より先に読む
{
  const ed = read('memberbook_editor.html');
  const a = ed.search(/<\?!=\s*include\('company_break'\)\s*\?>/), b = ed.indexOf("createHtmlOutputFromFile('memberbook_render')");
  ck(a >= 0 && b > a, '1) 編集画面が company_break を組版より先に読んでいない');
}

// 冊子の会社名の欄を、描いた行（字の位置で分ける）まで読む
const SNAP = () => Array.from(document.querySelectorAll('.page:not(.cover) .c')).filter((c) => c.querySelector('.co')).map((c) => {
  const el = c.querySelector('.co'), t = el.firstElementChild, rows = [], rg = document.createRange();
  const w = document.createTreeWalker(t, NodeFilter.SHOW_TEXT);
  let n, cur = null;
  while ((n = w.nextNode())) {
    for (let i = 0; i < n.nodeValue.length; i++) {
      rg.setStart(n, i); rg.setEnd(n, i + 1);
      const r = rg.getClientRects()[0];
      if (!r || !r.width) continue;                                   // 行の終わりの空白
      if (!cur || Math.abs(r.top - cur.top) > r.height * 0.5) rows.push(cur = { top: r.top, s: '' });
      cur.s += n.nodeValue.charAt(i);
    }
  }
  const er = el.getBoundingClientRect(), tr = t.getBoundingClientRect();
  return { wrap: el.classList.contains('wrap'), lines: el.classList.contains('lines'), size: parseFloat(el.style.fontSize), over: el.hasAttribute('data-over'),
    text: t.textContent, parts: Array.from(t.childNodes).filter((x) => x.nodeType === 3).map((x) => x.nodeValue), br: t.querySelectorAll('br').length,
    l2: t.getAttribute('data-l2'), rows: rows.map((r) => r.s), lineBreak: getComputedStyle(t).lineBreak,
    inside: tr.left >= er.left - 0.5 && tr.right <= er.right + 0.5 && tr.top >= er.top - 0.5 && tr.bottom <= er.bottom + 0.5 };
});

async function checkBook(browser) {
  const ctx = await browser.newContext();
  await ctx.route(/fonts\.(googleapis|gstatic)\.com/, (r) => r.abort());       // Googleフォントは読まない（手元の書体で描く）
  const page = await ctx.newPage();
  await page.setContent('<html><body></body></html>');
  const render = read('memberbook_render.html').replace(/^<script>|<\/script>\s*$/g, '');
  // 編集画面と同じ並び（company_break → memberbook_render）で冊子のHTMLを作る
  const make = (withCb) => page.evaluate(({ js, members, cats, cover }) => {
    window.members = members; window.cats = cats; window.cover = cover; window.photosB64 = {};
    window.catOf = (k) => cats.find((c) => c.key === k) || null;
    (0, eval)(js);
    return bookHtml({ photos: {} });
  }, { js: (withCb ? read('company_break.js') + '\n' : '') + render, members: MEMBERS, cats: CATS, cover: COVER });

  // CB_ が無いところでも冊子は作れる（候補は付かない。以前と同じくブラウザに折らせる）
  const plain = await make(false);
  ck(!/data-l2=/.test(plain) && /株式会社見本不動産コンサルティングオフィス/.test(plain), '1) CB_ が無いところで冊子を作れない・候補が付いた');

  await page.setContent('<html><body></body></html>');
  const html = await make(true);
  const script = (html.match(/<script>([\s\S]*)<\/script>/) || [])[1] || '';
  ck(/data-l2=/.test(html) && script && !/CB_|companyBreakLib_/.test(script), '1) 冊子のHTMLに候補が無い・中のスクリプトが CB_ を使う');

  // 冊子のHTMLを、別の新しいページで開く（配布用・印刷・PDF と同じ。この中には CB_ が無い。
  // 同じページの setContent では、前に読んだ CB_ が window に残る）
  await page.close();
  const book = await ctx.newPage();
  await book.route('https://book.test/', (r) => r.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: html }));
  await book.goto('https://book.test/', { waitUntil: 'load' });
  await book.waitForFunction(() => document.documentElement.getAttribute('data-fitted') === '1', null, { timeout: 30000 });
  ck(await book.evaluate(() => typeof CB_ === 'undefined' && typeof companyBreakLib_ === 'undefined' && typeof bookHtml === 'undefined'),
     '1) 冊子の中に CB_ がある（配布用HTMLで確かめられない）');
  const A = await book.evaluate(SNAP);
  ck(A.length === BOOK.length, '1) カードの数: ' + A.length);
  BOOK.forEach(([company, want], i) => {
    const x = A[i] || {}, tag = `1) 「${company.replace(/\n/g, '\\n')}」`, norm = CB.normalize(company);
    const opts = CB.twoLineOptions(norm, 3);
    let got = null;
    try { got = JSON.parse(x.l2); } catch (e) { got = null; }
    ck(J(got) === J(opts), `${tag}: 候補がサーバーの CB_ と違う: ${x.l2} / ${J(opts)}`);
    ck(flat(x.text) === flat(norm) && !/[\n㈱]/.test(x.text), `${tag}: 名簿の書き方にそろっていない: ${J(x.text)}`);
    ck(x.inside && !x.over, `${tag}: 枠からはみ出す（${x.size}pt）: ${J(x.rows)}`);
    ck(x.rows.slice(1).every((r) => !NO_START.test(r)), `${tag}: 行の頭に来てはいけない字で行を始めた: ${J(x.rows)}`);
    if (want === 'one') ck(!x.wrap && x.rows.length === 1 && x.size <= 7 && x.size >= 4.5, `${tag}: 1行のまま小さくしていない: ${J(x)}`);
    if (want === 'best' || want === 'two') {
      const pick = opts.findIndex((o) => J(o) === J(x.parts));
      ck(x.wrap && x.lines && x.br === 1 && (want === 'best' ? pick === 0 : pick >= 0), `${tag}: CB_ の候補で2行にしていない: ${J(x.parts)} / ${J(opts)}`);
      ck(J(x.rows.map(flat)) === J(x.parts.map(flat)), `${tag}: 決めた2行を、ブラウザが折り直した: ${J(x.rows)}`);
      ck(x.size <= 4.5 && x.size >= 3.8, `${tag}: 2行の大きさ（4.5〜3.8pt）: ${x.size}`);
      ck(!brokenCorp(x.parts), `${tag}: 法人の種類の途中で折った: ${J(x.parts)}`);
    }
    if (want === 'css') {
      ck(x.wrap && !x.lines && x.br === 0 && x.rows.length === 2 && x.lineBreak === 'strict', `${tag}: 入らない名前をブラウザに折らせていない（line-break:strict）: ${J(x)}`);
    }
  });
  // 以前の折れ方（語の途中）になっていない
  const joined = A.map((x) => x.parts.join('|'));
  ck(!joined.some((s) => /ソリュー\||コン\|サル|まちづ\||見\|本/.test(s)), '1) 語の途中で折った: ' + J(joined));

  // 何度収め直しても同じ（書体が届いたあとの収め直しなど）
  await book.evaluate(() => { mbFitAll(document); mbFitAll(document); });
  const B = await book.evaluate(SNAP);
  ck(J(B) === J(A), '1) 収め直すと会社名の行が変わる: ' + J(B.map((x, i) => (J(x) === J(A[i]) ? '' : x.parts)).filter(Boolean)));
  await ctx.close();
}

// ===================== 2. ビジター様リストのPDF（印刷用シート）=====================
function checkVisitorList() {
  const env = makeEnv({ now: new Date(2026, 8, 29, 10, 0, 0) });
  const srv = Object.assign({}, env.globals);
  vm.createContext(srv);
  for (const f of fs.readdirSync(ROOT).filter((x) => /\.js$/.test(x)).sort()) {   // 本番と同じく全部のファイルを1つの場所に
    vm.runInContext(read(f), srv, { filename: f });
  }
  const HEAD = ['No', '業種区分', '氏名', 'ふりがな', 'カテゴリー', '会社名', '役職', 'メモ', '写真ファイル名', '一言コメント',
    '紹介してほしい人', '協業したい人', '入会日', '更新日', '更新期限日', '会社での役職'];
  env.reset([
    ['メンバー名簿', false, [HEAD, HEAD.map((h, i) => (i === 0 ? '1' : i === 2 ? '試験 花子' : ''))]],
    ['休会日', true, [['2026/12/30']]],
  ], {});

  // [会社名（名簿に入っている書き方）, 印刷用シートの欄]。以前はシートの折り返し（1行11字ほど）で「株式会社見本不動産コン／サルティング…」と折れた
  const VISITORS = [
    ['株式会社見本不動産コンサルティングオフィス', '株式会社見本不動産\nコンサルティング\nオフィス'],
    ['一般社団法人日本見本まちづくり推進協会', '一般社団法人\n日本見本まちづくり\n推進協会'],
    ['見本建築設計事務所一級建築士事務所', '見本建築設計事務所\n一級建築士事務所'],
    ['見本マーケティングソリューションズ株式会社', '見本マーケティング\nソリューションズ\n株式会社'],
    ['Mihon Global Consulting Partners株式会社', 'Mihon Global\nConsulting Partners\n株式会社'],
    ['㈱見本\nエンジニアリング', '株式会社\n見本エンジニアリング'],
    ['特定非営利活動法人見本子ども支援センター', null],
    ['株式会社見本ビューティー＆ウェルネスラボ', null],
    ['見本商事', '見本商事'],
    ['', ''],
  ];
  const CAT = '不動産コンサルティング（売買・賃貸・相続・事業承継）';
  const HEADER = ['_No', '参加者氏名', 'ふりがな', 'カテゴリー', '会社名', '招待者', 'メモ（ビジターリストに表示）', '種別', 'メール'];
  const rows = VISITORS.map(([c], i) => ({ _No: 'V' + ('0' + (i + 1)).slice(-2), 参加者氏名: '見本 来' + (i + 1), ふりがな: 'みほん らい', カテゴリー: CAT,
    会社名: c, 招待者: '試験 花子', 'メモ（ビジターリストに表示）': '', 種別: 'Visitor', メール: '' }));

  // 欄の内側の幅：列 160px ＝ 120pt から、セルの左右の余白（3pxずつ）と罫線（1px）を除く。字はシートの既定の 10pt
  const box = srv.VISITOR_COMPANY_BOX_ || {};
  ck(Math.abs(box.width - (160 - 3 * 2 - 1) * 0.75) <= 0.5 && box.pt === 10 && box.minPt === 10 && box.maxLines === 4, '2) 会社名の欄の幅・字の大きさ: ' + J(box));

  srv.createFinalSheet('2026/09/30', '2026/9/30(水) 第535回', rows.map((r) => Object.assign({}, r)), HEADER);
  const print = env.values('20260930参加者_印刷用') || [], data = env.values('20260930参加者') || [];
  const pi = print.findIndex((r) => r[0] === 'No.');
  ck(pi === 4 && print[pi][4] === '会社名', '2) 印刷用シートの見出しの位置: ' + pi);
  VISITORS.forEach(([raw, want], i) => {
    const tag = `2) 「${raw.replace(/\n/g, '\\n')}」`, cell = String((print[pi + 1 + i] || [])[4]), lines = cell.split('\n');
    const expect = CB.layout(raw, { width: 115, pt: 10, minPt: 10, maxLines: 4 });
    if (want != null) ck(cell === want, `${tag}: 印刷用の欄 ${J(cell)}（${J(want)} のはず）`);
    ck(cell === expect.lines.join('\n'), `${tag}: CB_ の切れ目と違う: ${J(cell)} / ${J(expect.lines)}`);
    ck(flat(lines.join('')) === flat(CB.normalize(raw)), `${tag}: 字が欠けた・増えた: ${J(cell)}`);
    ck(lines.length <= 4 && lines.every((l) => CB.fits(l, 10, 115)), `${tag}: 欄に入らない行（シートが折り直す）: ${J(lines)}`);
    ck(lines.slice(1).every((l) => !NO_START.test(l)) && !brokenCorp(lines), `${tag}: 行の頭の字・法人の種類の途中: ${J(lines)}`);
    ck(expect.natural, `${tag}: 語の途中で折った: ${J(lines)}`);
    // 名簿のシート（再編集で読み戻す）は元の書き方のまま。ほかの欄は変えない
    ck((data[1 + i] || [])[4] === raw, `${tag}: 名簿のシートの会社名が変わった: ${J((data[1 + i] || [])[4])}`);
    ck((print[pi + 1 + i] || [])[3] === CAT && (print[pi + 1 + i] || [])[1] === rows[i]['参加者氏名'], `${tag}: 会社名のほかの欄が変わった`);
  });
  const id = env.props['VISITOR_PDF_ID_20260930参加者'];
  ck(id && env.fileText(id).includes('株式会社見本不動産\nコンサルティング\nオフィス'), '2) PDFの会社名に改行が入っていない');
  // 再編集：読み戻すのは元の書き方。作り直しても同じ欄（改行を重ねない）
  const loaded = srv.loadSheetData('20260930参加者');
  ck(J(loaded.rows.map((r) => r['会社名'])) === J(VISITORS.map((v) => v[0])), '2) 再編集で読み戻した会社名: ' + J(loaded.rows.map((r) => r['会社名'])));
  srv.createFinalSheet(loaded.meeting.dateValue, loaded.meeting.display, loaded.rows, loaded.header);
  const again = env.values('20260930参加者_印刷用') || [];
  ck(J(again.map((r) => r[4])) === J(print.map((r) => r[4])), '2) 作り直したら印刷用の会社名が変わった');
  ck(J((env.values('20260930参加者') || []).map((r) => r[4])) === J(data.map((r) => r[4])), '2) 作り直したら名簿のシートの会社名が変わった');
}

(async () => {
  try { checkVisitorList(); } catch (e) { fails.push('2) 止まった: ' + (e && e.stack ? e.stack.split('\n').slice(0, 3).join(' / ') : e)); }
  const browser = await pw.chromium.launch();
  try { await checkBook(browser); } catch (e) { fails.push('1) 止まった: ' + (e && e.stack ? e.stack.split('\n').slice(0, 3).join(' / ') : e)); }
  await browser.close();
  if (fails.length) {
    console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
    fails.slice(0, 40).forEach((f) => console.log('  - ' + f));
    process.exit(1);
  }
  console.log('会社名の改行（メンバーブック・ビジター様リスト）: 検査 ' + checks + ' 件 OK: メンバーブックは CB_ の候補で2行・入らなければ行頭禁則つきで折る・'
    + '冊子のHTMLは CB_ なしで収める・収め直しても同じ／ビジター様リストの印刷用は列の幅で改行・名簿のシートは元のまま');
})();

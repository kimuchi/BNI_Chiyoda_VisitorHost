// メンバーブック（memberbook_render.html・memberbook_editor.html・memberbook_srv.js）を確かめる。名簿は作り物。
//   1. 名簿への1人ぶんの保存（saveMemberBookMember）… 直せる列だけ書き換える・居なければ足す・古い名簿に列を足す
//   2. 組版（Chromium で実際に描いて測る）… 1ページ18名（3列×6行・ページいっぱい）、長い文字を枠に収める
//      （はみ出さない・ほかの欄に重ならない）、会社での役職とBNIの役職、表紙、業種区分の説明は使っている区分だけ
//   3. 編集画面 … 「反映」ですぐ名簿に保存・閉じるときに反映していない変更を聞く・並べ替えは自動で保存・プレビューのカードを押すと編集
//
//   node tools/check_memberbook.js [Googleフォントを写したディレクトリ（css.txt と gstatic/）]
//   フォントのディレクトリを渡さなければ、パソコンにある書体で描く（収め方の確かめはどちらでもできる）
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

// ===================== 1. 名簿への1人ぶんの保存 =====================
{
  const HEAD15 = ['No', '業種区分', '氏名', 'ふりがな', 'カテゴリー', '会社名', '役職', 'メモ', '写真ファイル名', '一言コメント',
                  '紹介してほしい人', '協業したい人', '入会日', '更新日', '更新期限日'];
  const grid = [HEAD15.slice(),
    ['1', '企業サポート', '見本 一郎', 'みほん', '税理士', '見本会計', 'ビジターホスト', 'メモ1', 'p1.jpg', '古い一言', '古い紹介', '古い協業', '2024/01/01', '', '2027/01/01'],
    ['2', '不動産関連', '見本 花子', 'みほん', '売買仲介', '花子不動産', '', '', '', '', '', '', '', '', '']];
  const sheet = {
    getLastRow: () => grid.length,
    getRange(r, c, nr, nc) {
      return {
        getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => ((grid[r - 1 + i] || [])[c - 1 + j] ?? ''))),
        setValues(v) { v.forEach((row, i) => { grid[r - 1 + i] = grid[r - 1 + i] || []; row.forEach((x, j) => { grid[r - 1 + i][c - 1 + j] = x; }); }); return this; },
        setFontWeight() { return this; }, setBackground() { return this; },
      };
    },
  };
  const box = { console, normName_: (s) => String(s == null ? '' : s).normalize('NFKC').replace(/[\s　]/g, ''),
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) } };
  vm.createContext(box);
  vm.runInContext(fs.readFileSync(path.join(ROOT, 'member_master_srv.js'), 'utf8'), box);
  vm.runInContext(fs.readFileSync(path.join(ROOT, 'memberbook_srv.js'), 'utf8'), box);
  box.ensureMemberSheet_ = () => sheet;
  // 直す前の氏名で行を探し、直せる列だけ書き換える（ふりがな・メモ・写真・日付はそのまま）
  let r = box.saveMemberBookMember('見本 一郎', { cat: '企業サポート', name: '見本　一郎', title: '税理士（相続）', company: '見本会計事務所',
    position: '代表取締役', role: 'ビジターホスト', comment: '新しい一言', refer: '税理士', collab: '協業したい人・弁護士' });
  ck(r.ok && !r.added, '1人ぶんの保存: ' + J(r));
  ck(grid[0].length === 16 && grid[0][15] === '会社での役職', '古い名簿に「会社での役職」の列が足されていない: ' + J(grid[0]));
  ck(J(grid[1]) === J(['1', '企業サポート', '見本　一郎', 'みほん', '税理士（相続）', '見本会計事務所', 'ビジターホスト', 'メモ1', 'p1.jpg',
    '新しい一言', '税理士', '協業したい人・弁護士', '2024/01/01', '', '2027/01/01', '代表取締役']), '書き換えた行: ' + J(grid[1]));
  ck(J(grid[2].slice(0, 6)) === J(['2', '不動産関連', '見本 花子', 'みほん', '売買仲介', '花子不動産']), 'ほかの方の行が変わった: ' + J(grid[2]));
  // 氏名を直した方も、直す前の氏名で探す
  r = box.saveMemberBookMember('見本 花子', { name: '見本 華子', collab: '司法書士' });
  ck(r.ok && grid[2][2] === '見本 華子' && grid[2][11] === '司法書士' && grid[2][4] === '売買仲介', '氏名を直した方: ' + J(grid[2]));
  // 居ない方は最後に足す
  r = box.saveMemberBookMember('', { name: '見本 三郎', company: '三郎商店', position: '店長' });
  ck(r.ok && r.added && grid.length === 4 && grid[3][2] === '見本 三郎' && grid[3][15] === '店長', '足した方: ' + J(grid[3]));
  ck(!box.saveMemberBookMember('見本 三郎', { name: ' ' }).ok, '氏名が空でも保存された');
  // 名簿を読むと「会社での役職」が入る。列の足りない古い行は空
  ck(box.MEMBER_HEADERS_.length === 16 && box.MEMBER_HEADERS_[15] === '会社での役職' && box.MEMBER_HEADERS_[6] === '役職',
     '名簿の列: ' + J(box.MEMBER_HEADERS_));
}

// ===================== 2・3. 組版と編集画面（Chromium）=====================
// 作り物の名簿。長さの違う文をわざと混ぜる（1行で収まる・小さくすれば1行・2行にする・どうしても入らない）
const LONG = 'とても長い文章がここに入ります。';
const MEMBERS = [];
const CATS = [['企業サポート', '#FFDE58'], ['研修・教育', '#C1FF72'], ['不動産関連', '#37B5FF'], ['建築・住まい', '#AAB5D9'],
  ['プロモーション', '#FF66C3'], ['暮らし・生活', '#5CE1E6'], ['美容と健康', '#7DD957'], ['飲食・エンタメ', '#CB6BE6'], ['金融保険', '#1f3864']]
  .map(([key, bg]) => ({ key, label: key, bg }));
for (let i = 0; i < 40; i++) {
  const k = i % 5;
  MEMBERS.push({
    no: String(i + 1), name: ['見本 一郎', '見本 花子', '長い名前の見本 太郎左衛門', '見本 三郎', 'Mihon Taro'][k] + (i >= 5 ? i : ''),
    cat: CATS[i % 8].key,
    title: ['税理士', '不動産売買仲介（相続・事業承継・海外資産）', '内装工事', '生命保険（法人）', '行政書士（建設業許可・外国人ビザ・相続・遺言）'][k],
    company: ['見本株式会社', '一般社団法人見本人材育成フォーラム', '(株)見本', '見本生命保険株式会社 東京本店 法人営業第一部', 'MIHON Holdings Co., Ltd.'][k],
    position: ['代表取締役', '', '取締役 兼 営業本部長 兼 経営企画室長', '支店長', ''][k],
    role: ['ビジターホスト', 'エデュケーションコーディネーター（サポート）・Webチーム', '', 'プレジデント', ''][k],
    comment: ['見本の一言です', LONG.repeat(2), LONG.repeat(3), LONG.repeat(8), ''][k],
    refer: ['税理士・弁護士', '事業を広げたい中小企業の社長、新しく店舗を出す飲食店のオーナー', LONG.repeat(3), LONG.repeat(9), '司法書士'][k],
    collab: ['司法書士', '保険・不動産・士業', LONG.repeat(2), LONG.repeat(9), ''][k],
  });
}
const OVER_OK = new Set(['cm', 'rf', 'cb']);      // 4番目の方（とても長い文）だけ、入りきらないのが正しい
const COVER = { title: 'BNI 見本 chapter Member Book', term: '24期', pname: '見本 一郎', prole: '見本チャプター\n第24期プレジデント',
  ptext: 'あいさつの文です。'.repeat(60), philosophyTitle: 'BNIの理念', philosophy: '理念の本文です。'.repeat(6), aboutTitle: 'チャプターとは？',
  about: '紹介の本文です。'.repeat(5), benefitsTitle: 'メリット', benefits: '1.メリット　2.メリット\n3.メリット',
  scheduleFrom: '7:15', scheduleTo: '9:15', schedule: Array.from({ length: 20 }, (_, i) => '項目' + (i + 1) + (i % 4 === 1 ? 'とても長い項目の名前が入るところです' : '')).join('\n'),
  termsTitle: 'BNIの用語の説明', terms: '①チャプター\n説明の文です。説明の文です。\n②リファーラル\n説明の文です。' };

async function routeFonts(ctx) {
  if (!FONT_DIR) {                                   // 写したフォントが無いときは、Googleフォントを読まない（手元の書体で描く）
    await ctx.route(/fonts\.(googleapis|gstatic)\.com/, (r) => r.abort());
    return;
  }
  const css = fs.readFileSync(path.join(FONT_DIR, 'css.txt'), 'utf8');
  await ctx.route('https://fonts.googleapis.com/**', (r) => r.fulfill({ status: 200, contentType: 'text/css', body: css }));
  await ctx.route('https://fonts.gstatic.com/**', (r) => {
    const f = path.join(FONT_DIR, 'gstatic', r.request().url().replace('https://fonts.gstatic.com/', '').replace(/\//g, '_'));
    if (!fs.existsSync(f)) return r.fulfill({ status: 404, body: '' });
    r.fulfill({ status: 200, contentType: 'font/woff2', body: fs.readFileSync(f), headers: { 'access-control-allow-origin': '*' } });
  });
}

// 編集画面のHTML（テンプレートの include を中身に置き換える）
function editorHtml() {
  const render = fs.readFileSync(path.join(ROOT, 'memberbook_render.html'), 'utf8');
  return fs.readFileSync(path.join(ROOT, 'memberbook_editor.html'), 'utf8')
    .replace(/<\?!=\s*HtmlService\.createHtmlOutputFromFile\('memberbook_render'\)\.getContent\(\);\s*\?>/, () => render);
}
// google.script.run の代わり（呼ばれた関数と引数を window.__calls に残す）
const STUB = (members, cats, cover) => `
  window.__calls = [];
  window.__data = ${J({ members, cats, cover })};
  (function(){
    function runner(){
      var ok = null, ng = null, r = {};
      r.withSuccessHandler = function(f){ ok = f; return r; };
      r.withFailureHandler = function(f){ ng = f; return r; };
      var answer = {
        getMemberBookData: function(){ return { ok: true, members: JSON.parse(JSON.stringify(window.__data.members)), categories: window.__data.cats, cover: window.__data.cover }; },
        getMemberPhotoThumbs: function(){ return { ok: true, map: {} }; },
        getMemberPhotosBase64: function(){ return { ok: true, map: {} }; },
        saveMemberBookMember: function(o, m){ return { ok: true, message: '「' + m.name + '」を名簿に保存しました。' }; },
        saveMemberBookData: function(){ return { ok: true, message: 'メンバー名簿を保存しました。' }; },
        saveMemberBookCover: function(c){ return { ok: true, cover: c, message: '保存しました' }; },
        exportMemberBookHtml: function(){ return { ok: true, message: '保存しました', url: '#' }; }
      };
      Object.keys(answer).forEach(function(k){
        r[k] = function(){ var args = Array.prototype.slice.call(arguments);
          window.__calls.push({ fn: k, args: JSON.parse(JSON.stringify(args)) });
          setTimeout(function(){ ok && ok(answer[k].apply(null, args)); }, 5); };
      });
      return r;
    }
    window.google = { script: { get run(){ return runner(); } } };
  })();`;

(async () => {
  const browser = await pw.chromium.launch();
  const ctx = await browser.newContext();
  await routeFonts(ctx);

  // ---------- 2. 組版 ----------
  {
    const page = await ctx.newPage();
    await page.setContent('<html><body></body></html>');
    const render = fs.readFileSync(path.join(ROOT, 'memberbook_render.html'), 'utf8').replace(/^<script>|<\/script>\s*$/g, '');
    const html = await page.evaluate(({ render, members, cats, cover }) => {
      window.members = members; window.cats = cats; window.cover = cover; window.photosB64 = {};
      window.catOf = (k) => cats.find((c) => c.key === k) || null;
      (0, eval)(render);
      return bookHtml({ photos: {} });
    }, { render, members: MEMBERS, cats: CATS, cover: COVER });
    await page.setContent(html, { waitUntil: 'load' });
    await page.waitForFunction(() => document.documentElement.getAttribute('data-fitted') === '1', null, { timeout: 30000 });
    const R = await page.evaluate(() => {
      const pt = (px) => px * 72 / 96;
      const rect = (e) => { const b = e.getBoundingClientRect(); return { x: b.left, y: b.top, w: b.width, h: b.height, r: b.right, b: b.bottom }; };
      const pages = Array.from(document.querySelectorAll('.page'));
      const cards = Array.from(document.querySelectorAll('.page .c'));
      const out = { pages: pages.length, cover: !!document.querySelector('.page.cover'), cards: [], pageSize: [pt(pages[1].offsetWidth), pt(pages[1].offsetHeight)] };
      out.gridCells = pages.slice(1).map((p) => p.querySelectorAll('.c').length);
      cards.forEach((c) => {
        if (!c.querySelector('.nm')) { out.cards.push(null); return; }
        const cr = rect(c), f = {};
        c.querySelectorAll('.fit').forEach((e) => {
          const k = e.className.split(' ')[0], t = e.firstElementChild, er = rect(e), tr = rect(t);
          f[k] = { size: parseFloat(e.style.fontSize), over: e.hasAttribute('data-over'), wrap: e.classList.contains('wrap'),
                   text: t.textContent, box: er, tbox: tr,
                   inside: e.hasAttribute('data-over') || (tr.w <= er.w + 1 && tr.h <= er.h + 1) };
        });
        out.cards.push({ cell: cr, f, bg: getComputedStyle(c.querySelector('.hd')).backgroundColor });
      });
      const lg = Array.from(document.querySelectorAll('.cv-lgi span')).map((s) => s.textContent);
      out.legend = lg;
      out.blocks = Array.from(document.querySelectorAll('.fitblock')).map((e) => ({ c: e.className, z: e.firstElementChild.style.zoom, over: e.hasAttribute('data-over') }));
      out.title = document.querySelector('.cv-ttl span').textContent;
      out.font = getComputedStyle(document.body).fontFamily;
      out.loaded = ['400', '500', '700'].map((w) => document.fonts.check(w + ' 12pt "Noto Serif JP"', '見本'));
      return out;
    });
    ck(R.cover && R.pages === 1 + Math.ceil(MEMBERS.length / 18), 'ページ数: ' + R.pages);
    ck(Math.abs(R.pageSize[0] - 595.3) < 1.5 && Math.abs(R.pageSize[1] - 841.9) < 1.5, 'ページの大きさ（A4）: ' + J(R.pageSize));
    ck(R.gridCells.every((n) => n === 18), '1ページ18枠（最後のページも罫線の枠はそのまま）: ' + J(R.gridCells));
    ck(/Noto Serif JP/.test(R.font), '書体の指定: ' + R.font);
    if (FONT_DIR) ck(R.loaded.every(Boolean), '書体（Noto Serif JP）を読み込んでから収めていない: ' + J(R.loaded));
    ck(J(R.legend) === J(CATS.slice(0, 8).map((c) => c.label)), '業種区分の説明（使っている区分だけ）: ' + J(R.legend));
    ck(R.blocks.filter((b) => /cv-ptx/.test(b.c)).every((b) => +b.z < 1), '長い挨拶文が縮んでいない: ' + J(R.blocks));
    ck(R.title === COVER.title, '表紙の題: ' + R.title);
    const real = R.cards.filter(Boolean);
    ck(real.length === MEMBERS.length, 'カードの数: ' + real.length);
    real.forEach((c, i) => {
      const m = MEMBERS[i], k = i % 5, tag = `${m.no} ${m.name}`;
      // 寸法：1枚 70mm×49.5mm（198.4pt×140.3pt）
      ck(Math.abs(c.cell.w * 0.75 - 198.4) < 1.5 && Math.abs(c.cell.h * 0.75 - 140.3) < 1.5, tag + ': カードの大きさ ' + (c.cell.w * 0.75).toFixed(1) + '×' + (c.cell.h * 0.75).toFixed(1));
      // はみ出さない（入りきらないと分かっている欄を除く）
      Object.keys(c.f).forEach((key) => {
        const x = c.f[key];
        if (k === 3 && OVER_OK.has(key)) { ck(x.over, tag + ': とても長い ' + key + ' が「入りきらない」扱いになっていない'); return; }
        ck(x.inside && !x.over, `${tag}: ${key}「${x.text.slice(0, 16)}」が枠からはみ出す（${x.size}pt）`);
        ck(x.box.x >= c.cell.x - 0.5 && x.box.r <= c.cell.r + 0.5 && x.box.y >= c.cell.y - 0.5 && x.box.b <= c.cell.b + 0.5, `${tag}: ${key} の枠がカードの外`);
      });
      // 大きさの決まり：短い文はいつもの大きさ（名前14pt・会社名7pt・一言8pt・紹介8pt）
      if (k === 0) ck(c.f.nm.size === 14 && c.f.co.size === 7 && c.f.cat.size === 7 && c.f.cm.size === 8 && c.f.rf.size === 8 && c.f.cb.size === 8 && c.f.ps.size === 7,
                      tag + ': いつもの大きさでない ' + J(Object.fromEntries(Object.entries(c.f).map(([a, b]) => [a, b.size]))));
      // 長い文は小さく（1行のまま）、もっと長い文は2行に
      if (k === 1) ck(c.f.co.size < 7 && !c.f.co.wrap && c.f.rf.size < 8, tag + ': 長い会社名・紹介が小さくなっていない');
      if (k === 2) ck(c.f.cm.wrap && c.f.rf.wrap, tag + ': とても長い一言・紹介が2行になっていない');
      // 業種・会社名も、1行に入らなければ2行に（会社での役職・BNIの役職・お名前の3行と重ならない）
      if (k === 1 || k === 4) ck(c.f.cat.wrap, tag + ': 長い業種が2行になっていない');
      if (k === 3) ck(c.f.co.wrap, tag + ': 長い会社名が2行になっていない');
      // 会社での役職・BNIの役職は、あるときだけ。どちらもお名前の上で、重ならない
      ck(!!c.f.ps === !!m.position && !!c.f.bn === !!m.role, tag + ': 会社での役職・BNIの役職の出し方');
      const stack = ['ps', 'bn', 'nm'].filter((x) => c.f[x]).map((x) => c.f[x].tbox);
      for (let s = 1; s < stack.length; s++) ck(stack[s].y >= stack[s - 1].b - 1, tag + ': 役職とお名前が重なる');
      ck(c.f.co.tbox.b <= (c.f.ps || c.f.bn || c.f.nm).tbox.y + 0.5, tag + ': 会社名とその下の行が重なる');
      ck(c.f.cat.tbox.b <= c.f.co.tbox.y + 0.5, tag + ': 業種と会社名が重なる');
      ck(c.bg !== 'rgba(0, 0, 0, 0)', tag + ': 業種区分の色が無い');
    });
    await page.close();
  }

  // ---------- 3. 編集画面 ----------
  {
    const page = await ctx.newPage();
    const dialogs = [];
    page.on('dialog', async (d) => { dialogs.push(d.message()); await d.accept(); });
    await page.addInitScript(STUB(MEMBERS.slice(0, 5), CATS, COVER));
    await page.route('https://memberbook.test/', (r) => r.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: editorHtml() }));
    await page.goto('https://memberbook.test/', { waitUntil: 'load' });
    await page.waitForFunction(() => /5名を読み込みました/.test(document.getElementById('msg').textContent), null, { timeout: 10000 });
    const pv = page.frameLocator('#pv');
    await pv.locator('html[data-fitted="1"]').waitFor({ timeout: 20000 });
    ck(await pv.locator('.c.pick').count() === 5, 'プレビューのカードの数');
    // プレビューのカードを押すと、その方の編集が開く
    await pv.locator('.c.pick').nth(1).click();
    await page.waitForFunction(() => document.getElementById('modal').style.display === 'flex');
    ck(await page.inputValue('#e_name') === MEMBERS[1].name && await page.inputValue('#e_position') === MEMBERS[1].position
       && await page.inputValue('#e_role') === MEMBERS[1].role, 'カードを押して開いた編集の中身');
    // 反映：その方だけ、直す前の氏名で名簿に保存する
    await page.fill('#e_collab', '新しい協業カテゴリー'); await page.fill('#e_position', '代表');
    await page.click('button:has-text("反映")');
    await page.waitForFunction(() => window.__calls.some((c) => c.fn === 'saveMemberBookMember'));
    let call = await page.evaluate(() => window.__calls.filter((c) => c.fn === 'saveMemberBookMember').pop());
    ck(call.args[0] === MEMBERS[1].name && call.args[1].collab === '新しい協業カテゴリー' && call.args[1].position === '代表',
       '反映で保存した中身: ' + J(call.args));
    await page.waitForFunction(() => /名簿に保存済み/.test(document.getElementById('saveState').textContent));
    // 閉じる：反映していない変更があれば聞く（OK＝保存して閉じる）
    await pv.locator('.c.pick').nth(2).click();
    await page.fill('#e_comment', '閉じる前に直した一言');
    const before = await page.evaluate(() => window.__calls.length);
    await page.click('button:has-text("閉じる")');
    await page.waitForFunction((n) => window.__calls.length > n, before);
    call = await page.evaluate(() => window.__calls[window.__calls.length - 1]);
    ck(dialogs.some((m) => /反映していない変更があります/.test(m)) && call.fn === 'saveMemberBookMember' && call.args[1].comment === '閉じる前に直した一言',
       '閉じるときに変更を保存しない: ' + J(dialogs) + J(call));
    // 変えずに閉じるときは聞かない
    const nd = dialogs.length;
    await pv.locator('.c.pick').nth(0).click();
    await page.click('button:has-text("閉じる")');
    ck(dialogs.length === nd && await page.evaluate(() => document.getElementById('modal').style.display) === 'none', '変えていないのに聞かれた');
    // 並べ替え：少し待ってから、一覧をまとめて保存
    await page.evaluate(() => mv(0, 1));
    ck(await page.evaluate(() => /保存しています/.test(document.getElementById('saveState').textContent)), '並べ替えのあと、保存中の表示が無い');
    await page.waitForFunction(() => window.__calls.some((c) => c.fn === 'saveMemberBookData'), null, { timeout: 5000 });
    call = await page.evaluate(() => window.__calls.filter((c) => c.fn === 'saveMemberBookData').pop());
    ck(call.args[0][0].name === MEMBERS[1].name && call.args[0][1].name === MEMBERS[0].name, '並べ替えの保存の並び: ' + call.args[0].slice(0, 2).map((m) => m.name));
    // 追加：反映すると、並びごと保存
    await page.evaluate(() => addM());
    await page.fill('#e_name', '見本 追加'); await page.fill('#e_company', '追加株式会社');
    const b2 = await page.evaluate(() => window.__calls.length);
    await page.click('button:has-text("反映")');
    await page.waitForFunction((n) => window.__calls.slice(n).some((c) => c.fn === 'saveMemberBookData'), b2);
    call = await page.evaluate(() => window.__calls.filter((c) => c.fn === 'saveMemberBookData').pop());
    ck(call.args[0].length === 6 && call.args[0][5].name === '見本 追加' && call.args[0][5].company === '追加株式会社', '追加した方の保存: ' + J(call.args[0][5]));
    // 追加して、何も入れずに閉じると行は残らない
    await page.evaluate(() => addM());
    await page.click('button:has-text("閉じる")');
    ck(await page.evaluate(() => members.length) === 6, '空のまま閉じた追加の行が残った');
    // 入りきらない欄のお知らせ（とても長い文の方）
    await page.evaluate(() => { members[0].refer = 'とても長い文章がここに入ります。'.repeat(9); render(); });
    await page.waitForFunction(() => /入りきらない欄/.test(document.getElementById('overNote').textContent), null, { timeout: 20000 });
    ck(/最も紹介して欲しいカテゴリー/.test(await page.textContent('#overNote')), '入りきらない欄のお知らせ: ' + await page.textContent('#overNote'));
    // 印刷：冊子を開くと、収めたあとに印刷の画面を開く（ここでは開いたページの中身を確かめる）
    const popup = page.waitForEvent('popup');
    await page.evaluate(() => { window.print = () => {}; printBook(); });
    const w = await popup;
    await w.waitForFunction(() => document.documentElement.getAttribute('data-fitted') === '1', null, { timeout: 30000 });
    ck(await w.locator('.page').count() === 2 && await w.locator('.c.pick').count() === 0, '印刷の冊子（表紙＋1ページ・押せるカードは無し）');
    await page.close();
  }
  await browser.close();

  console.log(`メンバーブック: 検査 ${checks} 件`);
  if (fails.length) {
    console.log(`NG: ${fails.length} 件`);
    fails.slice(0, 40).forEach((f) => console.log('   ' + f));
    process.exit(1);
  }
  console.log('OK: 名簿への1人ぶんの保存・組版（3列×6行・長い文字を枠に収める・会社での役職とBNIの役職・表紙）・編集画面（反映ですぐ保存・閉じるときの確認・並べ替えの自動保存・プレビュー）');
})().catch((e) => { console.error(e); process.exit(1); });

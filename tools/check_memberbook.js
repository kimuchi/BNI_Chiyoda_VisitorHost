// メンバーブック（memberbook_render.html・memberbook_editor.html・memberbook_srv.js）を確かめる。名簿は作り物。
//   1. 名簿への1人ぶんの保存（saveMemberBookMember）… 直せる列だけ書き換える・居なければ足す・古い名簿に列を足す
//   2. 組版（Chromium で実際に描いて測る）… 1ページ18名（3列×6行・ページいっぱい）、長い文字を枠に収める
//      （はみ出さない・ほかの欄に重ならない）、会社での役職とBNIの役職、表紙、業種区分の説明は使っている区分だけ
//   3. 編集画面 … 「反映」ですぐ名簿に保存・閉じるときに反映していない変更を聞く・並べ替えは自動で保存・プレビューのカードを押すと編集・
//      「PDFを作ってドライブを更新」（PDFを画面の中で作って送る・差し替えられないときは確かめてから作り直す・作れないときの案内）
//      （PDFの中身と、サーバーの差し替えは check_memberbook_pdf.js）
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
  let maxCols = 15;                                   // 列を詰めた古い名簿（15列しかない）
  const sheet = {
    getLastRow: () => grid.length,
    getMaxColumns: () => maxCols,
    insertColumnsAfter(after, n) { maxCols += n; },
    getRange(r, c, nr, nc) {
      if (c - 1 + nc > maxCols) throw new Error('The coordinates of the range are outside the dimensions of the sheet.');
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
  ck(grid[0].length === 16 && grid[0][15] === '会社での役職' && maxCols === 16, '古い名簿に「会社での役職」の列が足されていない: ' + J(grid[0]) + ' ' + maxCols);
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

// ===================== 1b. プレジデント設定（期ごと）=====================
// 期ごとに保存する。以前の1件はその期のものとして読む。既定（氏名＝その期のプレジデントの担当者・肩書き）は保存しない。
// 表紙のほかの文章は期によらず1件。期の番号を付け直すとずれる
{
  const props = {};
  const RealDate = Date;
  let today = new RealDate('2026-09-28T09:00:00');                   // 23期（2026年4月〜9月）の終わり
  const box = { console, normName_: (s) => String(s == null ? '' : s).normalize('NFKC').replace(/[\s　]/g, ''),
    PropertiesService: { getScriptProperties: () => ({
      getProperty: (k) => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = String(v); },
      getProperties: () => Object.assign({}, props), deleteProperty: (k) => { delete props[k]; } }) } };
  box.Date = class extends RealDate { constructor(...a) { if (a.length) super(...a); else super(today.getTime()); } static now() { return today.getTime(); } };
  vm.createContext(box);
  vm.runInContext('Date = this.Date;', box);
  for (const f of ['chapter_srv.js', 'role_input_srv.js', 'member_master_srv.js', 'memberbook_srv.js']) {
    vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), box, { filename: f });
  }
  // 期ごとのプロパティ（BNI_MB_PRESIDENT_24 など）をまとめて読む
  const PRES = () => {
    const o = {};
    Object.keys(props).filter((k) => /^BNI_MB_PRESIDENT_\d+$/.test(k)).forEach((k) => { o[k.replace('BNI_MB_PRESIDENT_', '')] = JSON.parse(props[k]); });
    return o;
  };
  const PRESJ = () => J(PRES());
  // 以前の1件（23期のプレジデントの設定と、表紙のほかの文章）。24期の担当者は登録済み
  props.BNI_MB_COVER = J({ title: '見本の題', term: '23期', pname: '見本 一郎', prole: 'Activeチャプター\n第23期プレジデント',
    ptext: '23期の挨拶', philosophy: '見本の理念' });
  props.BNI_ROLE_HOLDERS_TERMS = J({ 24: { president: '見本 二郎' } });
  let c = box.getCoverInfo_();
  ck(c.termNo === 23 && c.term === '23期' && c.pname === '見本 一郎' && c.ptext === '23期の挨拶' && c.title === '見本の題' && c.philosophy === '見本の理念',
     '以前の1件を、いまの期（23期）の設定として読む: ' + J(c));
  let list = box.coverPresidentList_();
  ck(J(list.map((p) => [p.term, p.saved, p.current])) === J([[22, false, false], [23, true, true], [24, false, false]]), '選べる期: ' + J(list.map((p) => [p.term, p.saved, p.current])));
  // 24期（まだ保存していない）：氏名は24期のプレジデントの担当者、肩書きは既定、挨拶文は空
  c = box.getCoverInfo_(24);
  ck(c.term === '24期' && c.pname === '見本 二郎' && c.prole === 'Activeチャプター\n第24期プレジデント' && c.ptext === '', '保存していない期: ' + J(c));
  ck(box.getCoverInfo_(25).pname === '', '担当者を登録していない期は、前の期のプレジデントを出さない: ' + box.getCoverInfo_(25).pname);
  // 24期の挨拶を保存（氏名・肩書きは既定のまま）→ 既定は保存しない。23期はそのまま。ほかの文章は1件に
  let r = box.saveMemberBookCover({ title: '新しい題', pname: '見本 二郎', prole: 'Activeチャプター\n第24期プレジデント', ptext: '24期の挨拶',
    philosophy: '見本の理念' }, 24);
  ck(r.ok && r.cover.termNo === 24 && r.cover.ptext === '24期の挨拶' && r.presidents.some((p) => p.term === 24 && p.saved), '24期の保存: ' + J(r));
  ck(J(PRES()) === J({ 23: { pname: '見本 一郎', ptext: '23期の挨拶' }, 24: { ptext: '24期の挨拶' } }), '期ごとの保存の中身: ' + PRESJ());
  const shared = JSON.parse(props.BNI_MB_COVER);
  ck(shared.title === '新しい題' && !('pname' in shared) && !('ptext' in shared) && !('prole' in shared) && !('term' in shared), '表紙の1件からプレジデントの項目を外す: ' + props.BNI_MB_COVER);
  c = box.getCoverInfo_(23);
  ck(c.ptext === '23期の挨拶' && c.pname === '見本 一郎' && c.title === '新しい題', '24期を保存しても23期はそのまま: ' + J(c));
  // 10月になると、いまの期は24期
  today = new RealDate('2026-10-01T09:00:00');
  c = box.getCoverInfo_();
  ck(c.termNo === 24 && c.ptext === '24期の挨拶' && c.pname === '見本 二郎', '期が替わると、その期の設定: ' + J(c));
  // 担当者を直すと、既定のままの氏名も変わる。既定と違う氏名は保存する
  props.BNI_ROLE_HOLDERS_TERMS = J({ 24: { president: '見本 次郎' } });
  ck(box.getCoverInfo_(24).pname === '見本 次郎', '担当者を直したあとの氏名: ' + box.getCoverInfo_(24).pname);
  box.saveMemberBookCover({ pname: '見本 別名' }, 24);
  ck(box.getCoverInfo_(24).pname === '見本 別名' && PRES()[24].pname === '見本 別名', '既定と違う氏名: ' + PRESJ());
  // 定型文を初期値に戻しても、期ごとのプレジデント設定は残る
  r = box.resetMemberBookCoverText(24);
  ck(r.ok && r.cover.title === 'BNI Active chapter Member Book' && r.cover.ptext === '24期の挨拶' && box.getCoverInfo_(23).ptext === '23期の挨拶',
     '定型文を戻したあと: ' + J([r.cover.title, r.cover.ptext]));
  // メンバーブックHTMLの取り込み（表紙の「期」「氏名」「挨拶文」）は、その期に入る
  box.saveCoverInfo_({ term: 22, pname: '見本 三郎', ptext: '22期の挨拶' });
  ck(box.getCoverInfo_(22).ptext === '22期の挨拶' && box.getCoverInfo_(22).pname === '見本 三郎' && box.getCoverInfo_(24).ptext === '24期の挨拶',
     '取り込んだ表紙の期: ' + PRESJ());
  // 期の番号を付け直すと（+2）、期ごとの設定もずれる。既定の肩書きは新しい番号で
  box.coverShiftTerms_(2);
  ck(J(Object.keys(PRES()).sort()) === J(['24', '25', '26']) && box.coverPresidentOf_(26).ptext === '24期の挨拶'
     && box.coverPresidentOf_(26).prole === 'Activeチャプター\n第26期プレジデント' && !('BNI_MB_PRESIDENT_22' in props), '期の番号の付け直し: ' + PRESJ());
  // 期ごとにする前の1件のまま、期の番号を付け直したとき（前の番号の期として読んでからずらす）
  for (const k of Object.keys(props)) delete props[k];
  props.BNI_MB_COVER = J({ term: '23期', pname: '見本 一郎', ptext: '23期の挨拶' });
  box.coverShiftTerms_(1);
  ck(J(PRES()) === J({ 24: { pname: '見本 一郎', ptext: '23期の挨拶' } }), '以前の1件のまま付け直したとき: ' + PRESJ());
  // 期ごとに別のプロパティ（1つ 9KB の上限にかからない）。長い挨拶文を20期ぶん保存しても、どれも保存できる
  let big = true;
  for (let t = 30; t < 50; t++) big = big && box.saveMemberBookCover({ ptext: t + '期の挨拶。'.repeat(1) + 'あ'.repeat(600) }, t).ok;
  const sizes = Object.keys(props).filter((k) => /^BNI_MB_PRESIDENT_\d+$/.test(k)).map((k) => Buffer.byteLength(props[k], 'utf8'));
  ck(big && sizes.length === 21 && Math.max(...sizes) < 9000 && box.getCoverInfo_(49).ptext.indexOf('49期の挨拶。') === 0,
     '長い挨拶文を何期ぶんも保存したとき: ' + J({ big, n: sizes.length, max: Math.max(...sizes) }));
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
  return fs.readFileSync(path.join(ROOT, 'memberbook_editor.html'), 'utf8')
    .replace(/<\?!=\s*HtmlService\.createHtmlOutputFromFile\('([\w]+)'\)\.getContent\(\);\s*\?>/g,
      (m, name) => fs.readFileSync(path.join(ROOT, name + '.html'), 'utf8'));
}
// 期ごとのプレジデント設定（サーバーの coverPresidentList_ と同じ形）。24期が、いまの期
const PRESIDENTS = [
  { term: 24, label: '24期', range: '2026年10月〜2027年3月', current: true, saved: true, pname: '見本 一郎', prole: '見本チャプター\n第24期プレジデント', ptext: '24期の挨拶です。' },
  { term: 25, label: '25期', range: '2027年4月〜9月', current: false, saved: false, pname: '見本 二郎', prole: '見本チャプター\n第25期プレジデント', ptext: '' },
];
// google.script.run の代わり（呼ばれた関数と引数を window.__calls に残す）
const STUB = (members, cats, cover, drive) => `
  window.__calls = [];
  window.__failNext = window.__failNext || {};        // 関数名 → 何回失敗させるか（ok:false を返す）
  window.__pdfAnswers = [];                            // saveMemberBookPdfToDrive の返事（空ならうまくいったことにする）
  window.__data = ${J({ members, cats, presidents: PRESIDENTS, drive: drive || {},
    cover: Object.assign({}, cover, { termNo: 24, term: '24期', pname: PRESIDENTS[0].pname, prole: PRESIDENTS[0].prole, ptext: PRESIDENTS[0].ptext }) })};
  (function(){
    function runner(){
      var ok = null, ng = null, r = {};
      r.withSuccessHandler = function(f){ ok = f; return r; };
      r.withFailureHandler = function(f){ ng = f; return r; };
      var answer = {
        getMemberBookData: function(){ return { ok: true, members: JSON.parse(JSON.stringify(window.__data.members)), categories: window.__data.cats,
          cover: Object.assign({}, window.__data.cover), presidents: JSON.parse(JSON.stringify(window.__data.presidents)), drive: window.__data.drive }; },
        saveMemberBookPdfToDrive: function(){ return window.__pdfAnswers.shift() || { ok: true, created: false,
          url: 'https://drive.test/file/d/MB/view', downloadUrl: 'https://drive.google.com/uc?export=download&id=MB',
          message: 'ドライブのメンバーブック（メールで送るPDF）を差し替えました。URLはそのままです。' }; },
        getMemberPhotoThumbs: function(){ return { ok: true, map: {} }; },
        // 写真の実体：window.__photos（氏名 → 画像）にある方だけ。window.__photoFail の方は「読めなかった」
        getMemberPhotosBase64: function(names){
          var map = {}, missing = [], failed = [], gone = [], ph = window.__photos || {}, bad = window.__photoFail || [], lost = window.__photoGone || [];
          (names || []).forEach(function(n){ if (bad.indexOf(n) >= 0) failed.push(n); else if (lost.indexOf(n) >= 0) gone.push(n); else if (ph[n]) map[n] = ph[n]; else missing.push(n); });
          return { ok: true, map: map, missing: missing, gone: gone, failed: failed }; },
        saveMemberBookMember: function(o, m){ return { ok: true, message: '「' + m.name + '」を名簿に保存しました。' }; },
        saveMemberBookData: function(){ return { ok: true, message: 'メンバー名簿を保存しました。' }; },
        // プレジデントの項目は、渡した期（t）の設定として保存したことにする
        saveMemberBookCover: function(c, t){
          window.__data.presidents = window.__data.presidents.map(function(p){
            return p.term === t ? Object.assign({}, p, { saved: true, pname: c.pname, prole: c.prole, ptext: c.ptext }) : p; });
          return { ok: true, cover: Object.assign({}, c, { termNo: t, term: t + '期' }), presidents: window.__data.presidents, message: '保存しました' }; },
        exportMemberBookHtml: function(){ return { ok: true, message: '保存しました', url: '#' }; }
      };
      Object.keys(answer).forEach(function(k){
        r[k] = function(){ var args = Array.prototype.slice.call(arguments);
          window.__calls.push({ fn: k, args: JSON.parse(JSON.stringify(args)) });
          var bad = window.__failNext[k] > 0 && window.__failNext[k]--;
          setTimeout(function(){ ok && ok(bad ? { ok: false, message: '見本の失敗' } : answer[k].apply(null, args)); }, (window.__delay || {})[k] || 5); };
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
        out.cards.push({ cell: cr, f, no: rect(c.querySelector('.no')), bg: getComputedStyle(c.querySelector('.hd')).backgroundColor });
      });
      const lg = Array.from(document.querySelectorAll('.cv-lgi .it')).map((s) => s.textContent);
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
      // 会社での役職・BNIの役職の文字は、左上の番号の四角に重ならない
      const cross = (a, b) => a.x < b.r - 0.5 && a.r > b.x + 0.5 && a.y < b.b - 0.5 && a.b > b.y + 0.5;
      ['ps', 'bn', 'nm'].forEach((x) => { if (c.f[x]) ck(!cross(c.f[x].tbox, c.no), `${tag}: ${x} が番号の四角に重なる`); });
      const stack = ['ps', 'bn', 'nm'].filter((x) => c.f[x]).map((x) => c.f[x].tbox);
      for (let s = 1; s < stack.length; s++) ck(stack[s].y >= stack[s - 1].b - 1, tag + ': 役職とお名前が重なる');
      ck(c.f.co.tbox.b <= (c.f.ps || c.f.bn || c.f.nm).tbox.y + 0.5, tag + ': 会社名とその下の行が重なる');
      ck(c.f.cat.tbox.b <= c.f.co.tbox.y + 0.5, tag + ': 業種と会社名が重なる');
      ck(c.bg !== 'rgba(0, 0, 0, 0)', tag + ': 業種区分の色が無い');
    });
    await page.close();
  }

  // ---------- 2b. 表紙の長い肩書き・業種区分が11以上・一言の改行 ----------
  {
    const page = await ctx.newPage();
    await page.setContent('<html><body></body></html>');
    const render = fs.readFileSync(path.join(ROOT, 'memberbook_render.html'), 'utf8').replace(/^<script>|<\/script>\s*$/g, '');
    const cats11 = CATS.concat([['暮らしサービス', '#255e5e'], ['士業・コンサルティング・その他の専門サービス', '#7030a0']].map(([key, bg]) => ({ key, label: key, bg })));
    const mem = cats11.map((c, i) => Object.assign({}, MEMBERS[0], { no: String(i + 1), name: '見本 ' + (i + 1), cat: c.key,
      comment: i === 0 ? '一行目の一言\n二行目の一言' : '見本の一言です' }));
    const cov = Object.assign({}, COVER, { prole: '株式会社見本コンサルティンググループ 代表取締役\n見本チャプター\n第24期プレジデント\n（2026年4月〜9月）',
      pname: 'Mihon Taro Alexander Christopher Wellington' });
    const html = await page.evaluate(({ render, members, cats, cover }) => {
      window.members = members; window.cats = cats; window.cover = cover; window.photosB64 = {};
      window.catOf = (k) => cats.find((c) => c.key === k) || null;
      (0, eval)(render);
      return bookHtml({ photos: {} });
    }, { render, members: mem, cats: cats11, cover: cov });
    await page.setContent(html, { waitUntil: 'load' });
    await page.waitForFunction(() => document.documentElement.getAttribute('data-fitted') === '1', null, { timeout: 30000 });
    const R = await page.evaluate(() => {
      const r = (e) => e.getBoundingClientRect(), pg = r(document.querySelector('.page.cover'));
      const prs = document.querySelector('.cv-prs .r'), zi = prs.firstElementChild;
      const box = prs.closest('.cv-box'), lbl = box.querySelector('.cv-lbl'), ptx = box.querySelector('.cv-ptx');
      const txt = Array.from(zi.querySelectorAll('.rl,.nm')).map(r);
      return {
        prsZoom: +zi.style.zoom, prsOver: prs.hasAttribute('data-over'),
        prsTop: Math.min(...txt.map((x) => x.top)), prsBottom: Math.max(...txt.map((x) => x.bottom)),
        lblBottom: r(lbl).bottom, ptxTop: r(ptx).top,
        legend: Array.from(document.querySelectorAll('.cv-lgi .it')).map((e) => ({ b: r(e).bottom, over: !!e.querySelector('[data-over]'), text: e.textContent })),
        pageBottom: pg.bottom,
        cm: (() => { const e = document.querySelector('.page:not(.cover) .c .cm'); return { br: e.querySelectorAll('br').length, over: e.hasAttribute('data-over'), wrap: e.classList.contains('wrap') }; })(),
      };
    });
    ck(R.prsZoom < 1 && !R.prsOver && R.prsTop >= R.lblBottom - 0.5 && R.prsBottom <= R.ptxTop + 0.5,
       '表紙：長いプレジデントの肩書き・お名前が、見出しや挨拶文に重なる: ' + J(R));
    ck(R.legend.length === 11 && R.legend.every((x) => x.b <= R.pageBottom + 0.5 && !x.over), '業種区分が11のとき、説明がページからはみ出す・入らない: ' + J(R.legend));
    ck(R.cm.br === 1 && !R.cm.over, '一言の改行が残っていない・入らない: ' + J(R.cm));
    await page.close();
  }

  // ---------- 3. 編集画面 ----------
  {
    const page = await ctx.newPage();
    const dialogs = [];
    page.on('dialog', async (d) => { dialogs.push(d.message()); await d.accept(); });
    const EM = MEMBERS.slice(0, 5).map((m, i) => (i === 4 ? Object.assign({}, m, { cat: '研修教育' }) : m));   // 5人目は古い業種区分
    await page.addInitScript(STUB(EM, CATS, COVER));
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
    // 表紙・プレジデント設定（期ごと）：開いたときは、いまの期。期を切り替えると、その期の内容・冊子も切り替わる
    {
      ck(/24期/.test(await page.textContent('#coverNote')), '左の表紙の知らせに、冊子に出す期が無い: ' + await page.textContent('#coverNote'));
      await page.check('#pvCover');
      await pv.locator('html[data-fitted="1"]').waitFor({ timeout: 20000 });
      await page.evaluate(() => openCover());
      ck(await page.inputValue('#c_termNo') === '24' && await page.inputValue('#c_ptext') === '24期の挨拶です。'
         && await page.locator('#c_termNo option').count() === 2, '開いたときの期・挨拶文');
      await page.selectOption('#c_termNo', '25');
      ck(await page.inputValue('#c_pname') === '見本 二郎' && await page.inputValue('#c_ptext') === '', '25期に切り替えたときの中身');
      await page.waitForFunction(() => { try { return /25期 プレジデント挨拶/.test(document.getElementById('pv').contentDocument.body.textContent); } catch (e) { return false; } }, null, { timeout: 20000 });
      ck(/25期のプレジデント挨拶がまだありません/.test(await page.textContent('#coverNote')), '挨拶文の無い期の知らせ: ' + await page.textContent('#coverNote'));
      // 変えて保存していないまま期を切り替えると聞く（ここでは OK で切り替える）
      const n1 = dialogs.length;
      await page.fill('#c_ptext', '捨てる挨拶');
      await page.selectOption('#c_termNo', '24');
      ck(dialogs.length === n1 + 1 && /保存していません/.test(dialogs[dialogs.length - 1]) && await page.inputValue('#c_ptext') === '24期の挨拶です。',
         '保存していない挨拶文のまま期を切り替えたとき: ' + J(dialogs.slice(n1)));
      // 25期の挨拶を入れて保存 → 25期として保存する
      await page.selectOption('#c_termNo', '25');
      await page.fill('#c_ptext', '25期の挨拶です。');
      const b5 = await page.evaluate(() => window.__calls.length);
      await page.click('#coverModal button:has-text("保存")');
      await page.waitForFunction((n) => window.__calls.slice(n).some((c) => c.fn === 'saveMemberBookCover'), b5);
      call = await page.evaluate(() => window.__calls.filter((c) => c.fn === 'saveMemberBookCover').pop());
      ck(call.args[1] === 25 && call.args[0].ptext === '25期の挨拶です。' && call.args[0].pname === '見本 二郎', '25期の保存: ' + J(call.args));
      await page.waitForFunction(() => document.getElementById('coverModal').style.display === 'none');
      ck(!/まだありません/.test(await page.textContent('#coverNote')), '保存したあとも「挨拶がまだ」と出る: ' + await page.textContent('#coverNote'));
      // 閉じるとき、保存していない変更があれば聞く（OK＝保存して閉じる）
      await page.evaluate(() => openCover());
      await page.fill('#c_title', '閉じる前に直した題');
      const n2 = dialogs.length, b6 = await page.evaluate(() => window.__calls.length);
      await page.click('#coverModal button:has-text("閉じる")');
      await page.waitForFunction((n) => window.__calls.slice(n).some((c) => c.fn === 'saveMemberBookCover'), b6);
      ck(dialogs.length === n2 + 1 && /保存していない変更があります/.test(dialogs[dialogs.length - 1]), '表紙を閉じるときの確認: ' + J(dialogs.slice(n2)));
      await page.uncheck('#pvCover');
    }
    // 業種区分マスタに無い区分の方：開いても区分は空にならず、何も変えずに閉じても聞かれない
    {
      const n0 = dialogs.length;
      await page.evaluate(() => openM(4));
      ck(await page.inputValue('#e_cat') === '研修教育', '業種区分マスタに無い区分が空になる: ' + await page.inputValue('#e_cat'));
      await page.click('button:has-text("閉じる")');
      ck(dialogs.length === n0, '業種区分マスタに無い区分の方で、変えていないのに聞かれた');
    }
    // 1人ぶんの保存に失敗 → ほかの方の保存が通っても「保存できませんでした」は消えない（名簿に保存で消える）
    // 氏名を直して失敗した方は、次の「反映」でも直す前の氏名で探す（行が2つにならない）
    {
      const nameOf = (i) => page.evaluate((j) => members[j].name, i);
      const orig2 = await nameOf(2);
      await page.evaluate(() => { window.__failNext.saveMemberBookMember = 1; });
      await page.evaluate(() => openM(2));
      await page.fill('#e_name', '見本 改名');
      await page.click('button:has-text("反映")');
      await page.waitForFunction(() => /保存できませんでした/.test(document.getElementById('saveState').textContent));
      await page.evaluate(() => openM(3));
      await page.fill('#e_comment', 'ほかの方の一言');
      const b3 = await page.evaluate(() => window.__calls.length);
      await page.click('button:has-text("反映")');
      await page.waitForFunction((n) => window.__calls.length > n, b3);
      await page.waitForTimeout(100);
      ck(/保存できませんでした/.test(await page.textContent('#saveState')), '前の保存の失敗が、ほかの方の保存で消えた: ' + await page.textContent('#saveState'));
      await page.evaluate(() => openM(2));
      await page.fill('#e_comment', '改名した方の一言');
      const b4 = await page.evaluate(() => window.__calls.length);
      await page.click('button:has-text("反映")');
      await page.waitForFunction((n) => window.__calls.length > n, b4);
      call = await page.evaluate(() => window.__calls[window.__calls.length - 1]);
      ck(call.fn === 'saveMemberBookMember' && call.args[0] === orig2 && call.args[1].name === '見本 改名',
         '氏名を直して失敗した方の次の保存が、直す前の氏名で探していない: ' + J(call.args[0]) + ' ' + orig2);
      await page.evaluate(() => save());
      await page.waitForFunction(() => /名簿に保存済み/.test(document.getElementById('saveState').textContent));
    }
    // 印刷：冊子を開くと、収めたあとに印刷の画面を開く（ここでは開いたページの中身を確かめる）
    const popup = page.waitForEvent('popup');
    await page.evaluate(() => { window.print = () => {}; printBook(); });
    const w = await popup;
    await w.waitForFunction(() => document.documentElement.getAttribute('data-fitted') === '1', null, { timeout: 30000 });
    ck(await w.locator('.page').count() === 2 && await w.locator('.c.pick').count() === 0, '印刷の冊子（表紙＋1ページ・押せるカードは無し）');
    await page.close();
  }
  // ---------- 3b. 名簿を読み込めなかったとき：追加・保存をさせない（名簿が空で上書きされるため）----------
  {
    const page = await ctx.newPage();
    page.on('dialog', (d) => d.accept());
    await page.addInitScript('window.__failNext = { getMemberBookData: 1 };');
    await page.addInitScript(STUB(MEMBERS.slice(0, 5), CATS, COVER));
    await page.route('https://memberbook.test/', (r) => r.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: editorHtml() }));
    await page.goto('https://memberbook.test/', { waitUntil: 'load' });
    await page.waitForFunction(() => /見本の失敗/.test(document.getElementById('msg').textContent), null, { timeout: 10000 });
    await page.evaluate(() => { addM(); save(); openCover(); saveCover(); });
    await page.waitForTimeout(1500);
    const st = await page.evaluate(() => ({ n: members.length, modal: document.getElementById('modal').style.display,
      cover: document.getElementById('coverModal').style.display,
      saves: window.__calls.filter((c) => /^save|^reset/.test(c.fn)).length }));
    ck(st.n === 0 && st.modal !== 'flex' && st.cover !== 'flex' && st.saves === 0, '名簿を読み込めなかったのに、追加・保存・表紙の保存ができた: ' + J(st));
    await page.close();
  }
  // ---------- 3c. PDFを作ってドライブを更新（アップロードしない）----------
  {
    const page = await ctx.newPage();
    const dialogs = [];
    page.on('dialog', async (d) => { dialogs.push(d.message()); await d.accept(); });
    await page.addInitScript(STUB(MEMBERS.slice(0, 5), CATS, COVER, { id: 'MB', url: 'https://drive.test/file/d/MB/view', updated: '2026-09-20T01:02:00.000Z' }));
    await page.route('https://memberbook.test/', (r) => r.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: editorHtml() }));
    await page.goto('https://memberbook.test/', { waitUntil: 'load' });
    await page.waitForFunction(() => /5名を読み込みました/.test(document.getElementById('msg').textContent), null, { timeout: 10000 });
    const note = () => page.evaluate(() => document.getElementById('driveNote').innerHTML);
    let n = await note();
    ck(/href="https:\/\/drive\.test\/file\/d\/MB\/view"/.test(n) && /最後に差し替え：9\/20 /.test(n), 'ドライブのメンバーブックの様子: ' + n);
    const lastPdfCall = () => page.evaluate(() => window.__calls.filter((c) => c.fn === 'saveMemberBookPdfToDrive').map((c) => [c.args[0].length, atob(c.args[0].slice(0, 12)), c.args[1], c.args[2]]));
    await page.click('#btnPdf');
    await page.waitForFunction(() => /差し替えました/.test(document.getElementById('msg').textContent), null, { timeout: 60000 });
    let calls = await lastPdfCall();
    ck(calls.length === 1 && calls[0][1].startsWith('%PDF-1.4') && calls[0][0] > 50000 && calls[0][2] === 'MemberBook.pdf' && calls[0][3] === false,
       'PDFを作って送っていない（アップロードなしで差し替え）: ' + J(calls));
    const msg = await page.evaluate(() => document.getElementById('msg').innerHTML);
    ck(/ドライブで開く/.test(msg) && /uc\?export=download&amp;id=MB/.test(msg), '差し替えたあとのリンク: ' + msg);
    n = await note();
    ck(/最後に差し替え/.test(n) && !/9\/20 /.test(n), '差し替えた日時が変わっていない: ' + n);
    ck(await page.evaluate(() => Array.from(document.querySelectorAll('.side button')).every((b) => !b.disabled)), '終わったのにボタンが押せないまま');
    // 差し替えられない（削除した・権限が無い）：「新しく作り直す」を押し、確かめてから作り直す
    await page.evaluate(() => { window.__pdfAnswers.push({ ok: false, canRecreate: true, message: 'ドライブのメンバーブック（メールで送るPDF）を差し替えられませんでした。' }); });
    await page.click('#btnPdf');
    await page.waitForFunction(() => /新しく作り直す/.test(document.getElementById('msg').textContent), null, { timeout: 60000 });
    ck((await lastPdfCall()).length === 2, '差し替えられないときの呼び出し');
    // 作り直す前に内容が変わった：作り直しは、いまの内容でPDFを作り直してから送る
    await page.evaluate(() => { members[0].company = '見本 作り直し前に直した会社'; });
    await page.click('#msg button');
    await page.waitForFunction(() => window.__calls.filter((c) => c.fn === 'saveMemberBookPdfToDrive').length === 3, null, { timeout: 60000 });
    calls = await lastPdfCall();
    const b64 = await page.evaluate(() => window.__calls.filter((c) => c.fn === 'saveMemberBookPdfToDrive').map((c) => c.args[0]));
    ck(calls[2][3] === true && b64[2] !== b64[1] && dialogs.some((m) => /URL|リンク/.test(m) && /古いまま/.test(m) && /いまの内容/.test(m)),
       '作り直す（いまの内容で作り直して・確かめてから）: ' + J(calls.map((c) => c.slice(1))) + ' ' + J(dialogs));
    // PDFを作れない（このブラウザーでは描けない など）：知らせて、印刷からアップロードする道を案内する。ボタンは押せるまま
    await page.evaluate(() => { window.mbpBuildPdf = () => Promise.reject(new Error('見本の失敗')); });
    await page.click('#btnPdf');
    await page.waitForFunction(() => /PDFを作れませんでした/.test(document.getElementById('msg').textContent), null, { timeout: 30000 });
    const t = await page.evaluate(() => document.getElementById('msg').textContent);
    ck(/見本の失敗/.test(t) && /メンバーブック\(PDF\)の更新/.test(t) && (await lastPdfCall()).length === 3, 'PDFを作れないときの知らせ: ' + t);
    ck(await page.evaluate(() => Array.from(document.querySelectorAll('.side button')).every((b) => !b.disabled)), '失敗のあと、ボタンが押せないまま');
    await page.close();
    // 写真：一覧のサムネイルの読み込みを待たずに押しても、全員分の写真（元の写真）をそろえてから作る。
    // 作っているあいだは画面を覆って操作を止める。あとから届いたサムネイルの知らせで、結果を消さない
    {
      const p3 = await ctx.newPage();
      const RED = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==';
      const names = MEMBERS.slice(0, 5).map((m) => m.name);
      await p3.addInitScript(`window.__delay = { getMemberPhotoThumbs: 2500, saveMemberBookPdfToDrive: 400 }; window.__photos = ${J(Object.fromEntries(names.map((n) => [n, RED])))};`);
      await p3.addInitScript(STUB(MEMBERS.slice(0, 5), CATS, COVER, { id: 'MB', url: 'https://drive.test/file/d/MB/view' }));
      await p3.route('https://memberbook.test/', (r) => r.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: editorHtml() }));
      await p3.goto('https://memberbook.test/', { waitUntil: 'load' });
      await p3.waitForFunction(() => /5名を読み込みました/.test(document.getElementById('msg').textContent), null, { timeout: 10000 });
      ck(await p3.evaluate(() => Object.keys(photos).length) === 0, '（前提）サムネイルがまだ届いていない');
      await p3.click('#btnPdf');
      ck(await p3.evaluate(() => document.getElementById('pdfWait').style.display) === 'flex', '作っているあいだ画面を覆っていない');
      await p3.waitForFunction(() => /差し替えました/.test(document.getElementById('msg').textContent), null, { timeout: 60000 });
      const got = await p3.evaluate(() => ({ asked: [].concat.apply([], window.__calls.filter((c) => c.fn === 'getMemberPhotosBase64').map((c) => c.args[0])),
        msg: document.getElementById('msg').textContent, wait: document.getElementById('pdfWait').style.display }));
      ck(J(got.asked.slice().sort()) === J(names.slice().sort()), 'サムネイルを待たずに、全員分の写真をそろえていない: ' + J(got.asked));
      ck(/写真 5\/5名/.test(got.msg) && got.wait === 'none', '写真の数の知らせ・覆いが消えない: ' + J(got));
      await p3.waitForTimeout(3000);                    // サムネイル（2.5秒）が届いたあと
      ck(await p3.evaluate(() => document.getElementById('photoState').textContent === '' && window.__calls.some((c) => c.fn === 'getMemberPhotoThumbs')),
         '（前提）サムネイルの読み込みが終わっていない');
      ck(/差し替えました/.test(await p3.evaluate(() => document.getElementById('msg').textContent)), 'あとから届いたサムネイルの知らせで、結果が消えた');
      // 写真を読めなかった方がいる：ドライブは変えずに止める（写真の欠けたPDFを出さない）
      await p3.evaluate((n) => { photosB64 = {}; window.__photoFail = [n]; }, names[1]);
      const before = await p3.evaluate(() => window.__calls.filter((c) => c.fn === 'saveMemberBookPdfToDrive').length);
      await p3.click('#btnPdf');
      await p3.waitForFunction(() => /PDFを作れませんでした/.test(document.getElementById('msg').textContent), null, { timeout: 30000 });
      const t3 = await p3.evaluate(() => ({ msg: document.getElementById('msg').textContent, n: window.__calls.filter((c) => c.fn === 'saveMemberBookPdfToDrive').length,
        wait: document.getElementById('pdfWait').style.display }));
      ck(t3.n === before && /変えていません/.test(t3.msg) && t3.msg.includes(names[1]) && t3.wait === 'none', '写真を読めなかったのに差し替えた・知らせ: ' + J(t3));
      const uploads = () => p3.evaluate(() => window.__calls.filter((c) => c.fn === 'saveMemberBookPdfToDrive').length);
      // 知らせを消してから押し、作り終わる（覆いが消える）まで待つ。keys … 押したあとに打つキー
      const run = async (keys) => {
        await p3.waitForFunction(() => !pdfRunning, null, { timeout: 60000 });
        await p3.evaluate(() => { document.getElementById('msg').textContent = ''; });
        await p3.click('#btnPdf');
        let during = null;
        if (keys) {
          for (const k of keys) await p3.keyboard.press(k);
          during = await p3.evaluate(() => ({ disabled: Array.from(document.querySelectorAll('.side button')).every((b) => b.disabled),
            cover: document.getElementById('coverModal').style.display, running: pdfRunning }));
        }
        await p3.waitForFunction(() => !pdfRunning && document.getElementById('msg').textContent !== '', null, { timeout: 60000 });
        return { msg: await p3.evaluate(() => document.getElementById('msg').textContent), during };
      };
      // ブラウザーで開けない写真（HEIC を .jpg にしたもの）：止めて、その方を知らせる（印刷からのアップロードは案内しない）
      await p3.evaluate((n) => { photosB64 = {}; window.__photoFail = []; window.__photos[n] = 'data:image/jpeg;base64,AAAAGGZ0eXBoZWljAAAAAG1pZjE='; }, names[2]);
      let n0 = await uploads();
      let t = (await run()).msg;
      ck(await uploads() === n0 && t.includes(names[2]) && /開けなかった/.test(t) && /メンバー写真/.test(t) && !/冊子を印刷/.test(t), '開けない写真: ' + t);
      // 写真の一覧にあるのにファイルが無い方（削除した）：写真なしで作り、知らせる
      await p3.evaluate((n) => { photosB64 = {}; window.__photos[n[2]] = window.__photos[n[0]]; window.__photoGone = [n[3]]; }, names);
      n0 = await uploads();
      t = (await run()).msg;
      ck(await uploads() === n0 + 1 && /差し替えました/.test(t) && /写真 4\/5名/.test(t) && /見つからなかった方/.test(t) && t.includes(names[3]), '写真のファイルが無い方: ' + t);
      // 全員の写真のファイルが無い（写真のフォルダを見る権限が無い）：止める
      await p3.evaluate((ns) => { photosB64 = {}; window.__photoGone = ns.slice(); }, names);
      n0 = await uploads();
      t = (await run()).msg;
      ck(await uploads() === n0 && /1枚も開けませんでした/.test(t) && /権限/.test(t), '全員の写真のファイルが無いとき: ' + t);
      // 作っているあいだに Enter・Space・Tab を押しても、2回目は作らない・ほかの操作もできない（ボタンは押せない）
      await p3.evaluate(() => { photosB64 = {}; window.__photoGone = []; window.__delay.saveMemberBookPdfToDrive = 1500; });
      n0 = await uploads();
      const photo0 = await p3.evaluate(() => window.__calls.filter((c) => c.fn === 'getMemberPhotosBase64').length);
      const k = await run(['Enter', 'Space', 'Tab', 'Enter', 'Tab', 'Space']);
      await p3.waitForTimeout(2500);
      const photo1 = await p3.evaluate(() => window.__calls.filter((c) => c.fn === 'getMemberPhotosBase64').length);
      ck(await uploads() === n0 + 1 && photo1 - photo0 === 1 && k.during.disabled && k.during.running && k.during.cover !== 'flex' && /差し替えました/.test(k.msg),
         '作っているあいだのキー操作で、2回作った・ほかの操作ができた: ' + J({ uploads: (await uploads()) - n0, photoCalls: photo1 - photo0, during: k.during }));
      await p3.close();
    }
    // サムネイルの最後の回が読めなかった：「写真を読み込んでいます」を残さない
    {
      const p5 = await ctx.newPage();
      await p5.addInitScript('window.__failNext = { getMemberPhotoThumbs: 1 };');
      await p5.addInitScript(STUB(MEMBERS.slice(0, 3), CATS, COVER, {}));
      await p5.route('https://memberbook.test/', (r) => r.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: editorHtml() }));
      await p5.goto('https://memberbook.test/', { waitUntil: 'load' });
      await p5.waitForFunction(() => /3名を読み込みました/.test(document.getElementById('msg').textContent), null, { timeout: 10000 });
      await p5.waitForTimeout(500);
      ck(await p5.evaluate(() => document.getElementById('photoState').textContent) === '', 'サムネイルが読めなかったのに「写真を読み込んでいます」が残る');
      await p5.close();
    }
    // まだ登録していないとき
    const p2 = await ctx.newPage();
    await p2.addInitScript(STUB(MEMBERS.slice(0, 2), CATS, COVER, {}));
    await p2.route('https://memberbook.test/', (r) => r.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: editorHtml() }));
    await p2.goto('https://memberbook.test/', { waitUntil: 'load' });
    await p2.waitForFunction(() => /2名を読み込みました/.test(document.getElementById('msg').textContent), null, { timeout: 10000 });
    ck(/まだありません/.test(await p2.evaluate(() => document.getElementById('driveNote').textContent)), 'まだ登録していないときの様子');
    await p2.close();
  }
  await browser.close();

  console.log(`メンバーブック: 検査 ${checks} 件`);
  if (fails.length) {
    console.log(`NG: ${fails.length} 件`);
    fails.slice(0, 40).forEach((f) => console.log('   ' + f));
    process.exit(1);
  }
  console.log('OK: 名簿への1人ぶんの保存・組版（3列×6行・長い文字を枠に収める・会社での役職とBNIの役職・表紙）・編集画面（反映ですぐ保存・閉じるときの確認・並べ替えの自動保存・プレビュー・'
    + 'PDFを作ってドライブを更新）');
})().catch((e) => { console.error(e); process.exit(1); });

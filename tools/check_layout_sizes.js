// 会社名・カテゴリーの組版（slides_layout.html）の文字の大きさを確かめる。
// メンバーのページ（前半のウィークリープレゼン）と、後半のリファーラル発表で共通に使う。名前はすべて架空。
//
//   node tools/check_layout_sizes.js
//
// 確かめること
//   ・会社名44pt・カテゴリー32ptにそろえる（以前は1行に入るまで1人ずつ小さくしていて、大きさがバラバラだった）
//   ・1行に入らなければ、同じ大きさのまま2行（「社名／株式会社」「株式会社／社名」で分けられればそこで）。
//     2行にも入らないほど長いときだけ小さくする
//   ・2行目が1〜2文字だけにならない。語の途中（「ホームページ」）で折らず、「・」のあとで折る。「・」「）」を行の頭にしない
//   ・公式ファイルから作った雛形でも、雛形の字の大きさ（companyPt・categoryPt）には合わせない。
//     雛形の会社名の枠が44ptの1行より低いときは、枠を伸ばしてカテゴリーもそのぶん下げる
//   ・サーバーへ渡す大きさはいつも書いてある（0 だとテンプレートの大きさのままになり、そろわない）
// canvas の代わりに、全角1文字・半角0.5文字で測る（tools/mp_plan.js と同じ）
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const J = (x) => JSON.stringify(x);

const html = fs.readFileSync(path.join(__dirname, '..', 'slides_layout.html'), 'utf8');
const js = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map((m) => m[1]).join('\n');
const ctx = () => { let size = 44; return { set font(v) { const m = /(\d+(?:\.\d+)?)px/.exec(v); size = m ? +m[1] : 44; },
  measureText(t) { let w = 0; for (const ch of String(t)) w += ch.charCodeAt(0) < 128 ? 0.5 : 1; return { width: w * size }; } }; };
const L = { console, document: { createElement: () => ({ getContext: ctx }) } };
vm.createContext(L);
vm.runInContext(js, L, { filename: 'slides_layout.html' });
const co = (s) => { const r = L.layoutCompany(s); return { lines: r.lines, pt: r.fontPt, tall: r.geom !== L.COMPANY_DEFAULT }; };
const ca = (s) => { const r = L.layoutCategory(s); return { lines: r.lines, pt: r.fontPt, tight: r.tight }; };

// ===== 1. 会社名 =====
// よくある長さの会社名は、みな44pt
const typical = ['見本商事', '試験デザイン事務所', '架空法律事務所', '株式会社仮名', '例示会計事務所', '模擬ホールディングス', 'MIHON DESIGN', '見本工務店株式会社'];
const got = typical.map(co);
ck(got.every((r) => r.pt === 44), '1) よくある長さの会社名が44ptでない: ' + J(typical.map((s, i) => [s, got[i].pt, got[i].lines])));
ck(J(co('見本商事')) === J({ lines: ['見本商事'], pt: 44, tall: false }), '1) 短い会社名: ' + J(co('見本商事')));
// 1行に入らない会社名：44ptのまま「社名／株式会社」
ck(J(co('見本コンサルティング株式会社')) === J({ lines: ['見本コンサルティング', '株式会社'], pt: 44, tall: true }),
   '1) 「社名／株式会社」の2行（44pt）: ' + J(co('見本コンサルティング株式会社')));
ck(J(co('株式会社見本コンサルティング')) === J({ lines: ['株式会社', '見本コンサルティング'], pt: 44, tall: true }),
   '1) 「株式会社／社名」の2行（44pt）: ' + J(co('株式会社見本コンサルティング')));
// ㈱の略記は展開して2行（元ツールと同じ）
ck(J(co('見本生命保険㈱')) === J({ lines: ['見本生命保険', '株式会社'], pt: 44, tall: true }), '1) ㈱の略記: ' + J(co('見本生命保険㈱')));
// 法人の言葉が無くても、44ptのまま2行
const noCorp = co('見本ホールディングスグループ東京');
ck(noCorp.pt === 44 && noCorp.lines.length === 2 && noCorp.tall, '1) 法人の言葉の無い長い会社名: ' + J(noCorp));
// 2行にも入らないほど長いときだけ、小さくする
const huge = co('一般社団法人見本とても長い名前の協会連合会の東京支部と関東地区本部');
ck(huge.pt < 44 && huge.pt >= 20 && huge.lines.length >= 2 && huge.lines.length <= 3, '1) とても長い会社名: ' + J(huge));
// 2行目が1〜2文字だけにならない
[11, 12].forEach((n) => {
  const s = '見本の' + 'ながいなまえのかいしゃめい'.slice(0, n - 3) + 'です';
  const r = co(s);
  ck(r.pt === 44 && (r.lines.length === 1 || r.lines[r.lines.length - 1].length >= 3), '1) 2行目が短すぎる: ' + J([s, r]));
});

// ===== 2. カテゴリー =====
ck(J(ca('税理士')) === J({ lines: ['【税理士】'], pt: 32, tight: false }), '2) 短いカテゴリー: ' + J(ca('税理士')));
const typicalCat = ['司法書士', '生命保険(法人)', '不動産売買仲介(相続)', 'デジタル広告制作(ホームページ・動画)', '行政書士（建設業許可・外国人ビザ）'];
ck(typicalCat.map(ca).every((r) => r.pt === 32), '2) よくある長さのカテゴリーが32ptでない: ' + J(typicalCat.map((s) => [s, ca(s)])));
const twoCat = ca('行政書士（建設業許可・外国人ビザ・相続手続き）');
ck(twoCat.pt === 32 && twoCat.lines.length === 2 && twoCat.tight && twoCat.lines[1].length >= 3, '2) 2行のカテゴリー（32pt）: ' + J(twoCat));
// 2行目が「務】」のように短くならない
for (let n = 13; n <= 17; n++) {
  const r = ca('見本'.repeat(20).slice(0, n));
  ck(r.lines.length === 1 || r.lines[r.lines.length - 1].length >= 3, '2) カテゴリーの2行目が短すぎる: ' + J([n, r]));
}
// 語の途中（「ホームページ」）で折らない。「・」「）」を行の頭にしない（「・」のあとでは折ってよい）
ck(J(ca('デジタル広告制作(ホームページ・動画)').lines) === J(['【デジタル広告制作(', 'ホームページ・動画)】']), '2) 語の途中で折った: ' + J(ca('デジタル広告制作(ホームページ・動画)')));
ck(J(twoCat.lines) === J(['【行政書士(建設業許可・', '外国人ビザ・相続手続き)】']), '2) 「・」のあとで折らない: ' + J(twoCat));
const allLines = typicalCat.concat(['行政書士（建設業許可・外国人ビザ・相続手続き・遺言書作成）', '税理士(個人事業主・スタートアップ企業)'])
  .map((s) => ca(s).lines).concat(typical.concat(['見本ホールディングスグループ東京']).map((s) => co(s).lines));
ck(allLines.every((ls) => ls.every((l) => !/^[・ー、。）)」】]/.test(l))), '2) 行の頭に「・」「）」などが来た: ' + J(allLines));
const hugeCat = ca('見本のとても長いカテゴリーの名前で、二行にも入らないほど長い専門分野の説明と、そのほかの細かい説明の続き');
ck(hugeCat.pt < 32 && hugeCat.pt >= 20, '2) とても長いカテゴリー: ' + J(hugeCat));

// ===== 3. 雛形の枠（setLayoutBoxes / withLayoutBoxes）=====
// 公式ファイルから作った雛形：雛形の字の大きさ（28pt・24pt）には合わせない。会社名の枠が44ptの1行より低いので伸ばし、カテゴリーも下げる
const official = { companyDefault: { x: 500000, y: 3000000, cx: 7000000, cy: 500000 }, categoryTop: 3600000, categoryWidth: 7000000,
                   companyPt: 28, categoryPt: 24 };
const inOfficial = JSON.parse(J(L.withLayoutBoxes(official, () => ({
  co: L.layoutCompany('見本商事'), co2: L.layoutCompany('見本コンサルティング株式会社'), ca: L.layoutCategory('税理士'),
  def: L.COMPANY_DEFAULT, top: L.CATEGORY_TOP, low: L.CATEGORY_TOP_LOW }))));
const oneLine = Math.round(91440 + 44 * 12700 * 1.213), grow = oneLine - 500000;
ck(inOfficial.co.fontPt === 44 && inOfficial.ca.fontPt === 32 && inOfficial.co2.fontPt === 44,
   '3) 公式ファイルの雛形で、雛形の字の大きさに合わせた: ' + J([inOfficial.co.fontPt, inOfficial.co2.fontPt, inOfficial.ca.fontPt]));
ck(inOfficial.def.cy === oneLine && inOfficial.top === 3600000 + grow && inOfficial.co.categoryTop === 3600000 + grow,
   '3) 会社名の枠を44ptの1行ぶんに伸ばし、カテゴリーを下げる: ' + J({ cy: inOfficial.def.cy, want: oneLine, top: inOfficial.top, catTop: inOfficial.co.categoryTop }));
ck(inOfficial.low > inOfficial.top && inOfficial.co2.categoryTop === inOfficial.low, '3) 2行のときのカテゴリーの位置: ' + J([inOfficial.top, inOfficial.low, inOfficial.co2.categoryTop]));
// リファーラル発表（元ツールの寸法のひな形）
const referral = { companyTall: { x: 600000, y: 2500000, cx: 6500000, cy: 1400000 }, categoryLow: 4000000, categoryWidth: 6800000 };
const inRef = JSON.parse(J(L.withLayoutBoxes(referral, () => [L.layoutCompany('見本商事'), L.layoutCategory('生命保険(法人)')])));
ck(inRef[0].fontPt === 44 && inRef[1].fontPt === 32, '3) リファーラル発表の大きさ: ' + J([inRef[0].fontPt, inRef[1].fontPt]));
// 切り替えたあとは元に戻る
ck(L.COMPANY_PT === 44 && L.CATEGORY_PT === 32 && L.COMPANY_DEFAULT.cy === 769441 && L.CATEGORY_TOP === 3456634, '3) 枠を元に戻さない');

// ===== 4. メンバーのページの並び（memberPresenItems）：サーバーへ渡す大きさは、いつも44・32 =====
const mctx = { rowsPerPage: 7, seconds: { weekly: 30, startup: 150 },
  blocks: [{ gkey: 'A', block: '見本の区分', count: 3 }],
  members: [{ name: '見本 一郎', company: '見本商事', title: '税理士', blockKey: 'A' },
            { name: '試験 二郎', company: '見本コンサルティング株式会社', title: '行政書士（建設業許可・外国人ビザ・相続手続き）', blockKey: 'A' },
            { name: '架空 三郎', company: '', title: '', blockKey: 'A' }],
  layoutBoxes: official };
const items = JSON.parse(J(L.memberPresenItems(mctx, 'A', '', false))).filter((x) => x.kind === 'individual');
ck(items.length === 3 && items.every((x) => x.companyPt === 44 && x.categoryPt === 32),
   '4) メンバーのページへ渡す大きさ: ' + J(items.map((x) => [x.name, x.companyPt, x.categoryPt])));

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('会社名・カテゴリーの組版: 検査 ' + checks + ' 件 OK: 会社名44pt・カテゴリー32ptにそろえる・入らなければ同じ大きさで2行'
  + '（社名／株式会社）・2行目が短すぎない・とても長いときだけ小さく・雛形の字の大きさに合わせない・枠を戻す');

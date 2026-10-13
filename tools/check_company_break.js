// 会社名の改行の決め方（company_break.js の CB_）を、作り物の会社名で確かめる。会社名はすべて架空。
//
//   node tools/check_company_break.js
//
// 確かめること
//   ・法人の種類の前後・空白・かっこ・よく使う語の切れ目で折る（「一般社団法人日本見／本協会」のように語の途中で折らない）
//   ・行の頭に来てはいけない字（ー・ッ・）・、など）で行を始めない。法人の種類・英単語の途中で折らない。どの行も枠に入る
//   ・字を少し小さくすれば良い切れ目で入るときは、語の途中で折るより小さくする
//   ・同じ会社名は、幅が近ければ同じ切れ目になる（メンバーのページとリファーラル発表など）
//   ・名簿のセルの中の改行・㈱ をそろえる
//   ・画面に渡す中身（webapp_srv.js の include('company_break')）が、サーバーと同じ計算になる
//   ・ooxml.js の lfToLineBreaks_：文字の中の改行を、同じ書式のランと <a:br> にする
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const J = (x) => JSON.stringify(x);

const S = { console };
vm.createContext(S);
for (const f of ['ooxml.js', 'company_break.js']) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), S, { filename: f });
const CB = S.CB_;
const L = (t, o) => CB.layout(t, o).lines.join('／');

// ===== 1) メンバーのページの大きさ（44pt・枠の内側 528pt・2行は28ptまで・3行は20ptまで）=====
const MEMBER = { width: 528, pt: 44, minPt: 28, hardMinPt: 20, maxLines: 3, linePenalty: 10 };
[
  ['株式会社見本コンサルティング', '株式会社／見本コンサルティング'],
  ['見本テクノロジーズ株式会社', '見本テクノロジーズ／株式会社'],
  ['一般社団法人日本見本協会', '一般社団法人／日本見本協会'],          // 以前は「一般社団法人日本見／本協会」
  ['税理士法人見本会計事務所', '税理士法人／見本会計事務所'],          // 以前は「税理士法人見本会計／事務所」
  ['見本総合法律事務所', '見本総合法律事務所'],
  ['Mihon Design Studio合同会社', 'Mihon Design Studio／合同会社'],
  ['株式会社見本不動産コンサルティングオフィス', '株式会社見本不動産／コンサルティングオフィス'],   // 以前は「…コン／サルティング…」
  ['社会保険労務士法人見本人事サポートセンター', '社会保険労務士法人／見本人事サポートセンター'],
  ['特定非営利活動法人見本子ども支援センター', '特定非営利活動法人／見本子ども支援センター'],     // 以前は「…法人見／本…」
  ['見本建築設計事務所一級建築士事務所', '見本建築設計事務所／一級建築士事務所'],                 // 以前は「…事務所一／級…」
  ['見本マーケティングソリューションズ株式会社', '見本マーケティング／ソリューションズ株式会社'],
  ['㈱見本エンジニアリング', '株式会社／見本エンジニアリング'],
  ['架空ライフデザイン株式会社', '架空ライフデザイン／株式会社'],
  ['株式会社SAMPLE Creative Works', '株式会社SAMPLE／Creative Works'],
].forEach(([t, want]) => ck(L(t, MEMBER) === want, `1) 「${t}」→「${L(t, MEMBER)}」（${want} のはず）`));
{
  // 小さくすれば良い切れ目で入るとき：語の途中で折らず、少し小さくする
  const r = CB.layout('株式会社見本ビューティー＆ウェルネスラボ', MEMBER);
  ck(r.natural && !/ビューテ$|^ィ/.test(r.lines.join('\n').split('\n').join('|')) && r.lines.length === 2, '1) ビューティーの途中で折った: ' + J(r));
}

// ===== 2) 決まりごと（作り物の名前を組み合わせて、たくさん試す）=====
const parts = {
  pre: ['', '', '株式会社', '有限会社', '一般社団法人', '税理士法人', 'NPO法人', '合同会社'],
  body: ['見本', '試験', '架空', 'ミホン', 'シケン', 'SAMPLE', 'Mihon', '見本まちづくり', '日本見本'],
  mid: ['', 'コンサルティング', 'テクノロジーズ', '不動産', 'デザインオフィス', 'ビューティー＆ウェルネス', ' Creative Works',
        '総合法律事務所', 'インターナショナルトレーディング', '建設工業', '（ミホン）', '・アンド・パートナーズ', 'ホールディングス'],
  post: ['', '', '株式会社', '合同会社', ' 東京支社', '研究所', 'センター', 'ラボ'],
};
const NO_START = /^[ー‐・･、。，．,.：:；;！？!?）)」』】〕〉》］\]々ゝゞヽヾぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶ％%]/;
const CORPS = ['株式会社', '有限会社', '合同会社', '一般社団法人', '税理士法人', 'NPO法人'];
let n = 0;
const sizes = [MEMBER, { width: 128.8, pt: 14, minPt: 7, maxLines: 2 }, { width: 115, pt: 10, minPt: 10, maxLines: 4 },
               { width: 605.6, pt: 44, minPt: 16, maxLines: 2 }];
for (const a of parts.pre) for (const b of parts.body) for (const c of parts.mid) for (const d of parts.post) {
  const t = (a + b + c + d).trim();
  if (!t) continue;
  for (const o of sizes) {
    n++;
    const r = CB.layout(t, o);
    if (r.fallback) continue;                                          // 入りきらない（とても長い）ときは詰めて折る
    const lines = r.lines;
    if (lines.join('').replace(/[ 　]/g, '') !== CB.normalize(t).replace(/[ 　]/g, '')) { ck(false, `2) 字が欠けた・増えた: ${t} → ${J(lines)}`); continue; }
    if (lines.some((x, i) => i > 0 && NO_START.test(x))) ck(false, `2) 行の頭に来てはいけない字: ${t} → ${J(lines)}`);
    if (lines.some((x) => !CB.fits(x, r.pt, o.width1 && lines.length === 1 ? o.width1 : o.width))) ck(false, `2) 枠に入らない行: ${t} @${r.pt} → ${J(lines)}`);
    for (let i = 1; i < lines.length; i++) {
      const joint = lines[i - 1] + '|' + lines[i];
      if (CORPS.some((w) => { for (let k = 1; k < w.length; k++) if (joint.includes(w.slice(0, k) + '|' + w.slice(k))) return true; return false; }))
        ck(false, `2) 法人の種類の途中で折った: ${t} → ${J(lines)}`);
    }
    // 英単語・数字の途中で折っていない（空白の位置で折ったときは、空白が消える）
    let rest = CB.normalize(t);
    lines.forEach((x, i) => {
      rest = rest.replace(/^[ 　]+/, '');
      const gap = rest.slice(x.length).charAt(0);
      if (i < lines.length - 1 && !/[ 　]/.test(gap) && /[A-Za-z0-9]$/.test(x) && /^[A-Za-z0-9]/.test(lines[i + 1]))
        ck(false, `2) 英単語の途中で折った: ${t} → ${J(lines)}`);
      rest = rest.slice(x.length);
    });
  }
}
checks++;
ck(n > 1000, '2) 試した数が少ない: ' + n);

// ===== 3) 同じ会社名は、幅が近ければ同じ切れ目（メンバーのページ 528pt・リファーラル発表 528pt・ビジター 605pt）=====
['株式会社見本コンサルティング', '一般社団法人日本見本協会', '見本マーケティングソリューションズ株式会社', '税理士法人見本会計事務所'].forEach((t) => {
  const a = CB.layout(t, MEMBER).lines, b = CB.layout(t, Object.assign({}, MEMBER, { width: 540 })).lines;
  ck(J(a) === J(b), `3) 幅が少し違うだけで切れ目が変わる: ${t} ${J(a)} / ${J(b)}`);
});

// ===== 4) 名簿の文字をそろえる =====
ck(CB.normalize('株式会社見本\nコンサルティング') === '株式会社見本コンサルティング', '4) セルの中の改行: ' + J(CB.normalize('株式会社見本\nコンサルティング')));
ck(CB.normalize('SAMPLE\nWORKS株式会社') === 'SAMPLE WORKS株式会社', '4) 英字のあいだの改行は空白に');
ck(CB.normalize('  ㈱見本   エンジニアリング ') === '株式会社見本 エンジニアリング', '4) ㈱・空白: ' + J(CB.normalize('  ㈱見本   エンジニアリング ')));
ck(L('', MEMBER) === '' && L('見本', MEMBER) === '見本', '4) 空・短い会社名');

// ===== 5) 画面に渡す中身（include('company_break')）も、同じ計算 =====
{
  const W = { console };
  vm.createContext(W);
  vm.runInContext(fs.readFileSync(path.join(ROOT, 'webapp_srv.js'), 'utf8'), W, { filename: 'webapp_srv.js' });
  W.companyBreakLib_ = S.companyBreakLib_;
  const html = W.include('company_break');
  const C = { console };
  vm.createContext(C);
  vm.runInContext(html.replace(/^<script>|<\/script>$/g, ''), C);
  const same = ['株式会社見本不動産コンサルティングオフィス', 'Mihon Global Consulting Partners株式会社', '見本建築設計事務所一級建築士事務所']
    .every((t) => J(C.CB_.layout(t, MEMBER)) === J(CB.layout(t, MEMBER)));
  ck(/^<script>\n/.test(html) && same, '5) 画面に渡す改行の決め方が、サーバーと違う');
}

// ===== 6) lfToLineBreaks_（文字の中の改行 → <a:br>）=====
{
  const x = '<a:p><a:r><a:rPr lang="ja-JP" sz="2000"/><a:t>株式会社\n見本</a:t></a:r></a:p>';
  ck(S.lfToLineBreaks_(x) === '<a:p><a:r><a:rPr lang="ja-JP" sz="2000"/><a:t>株式会社</a:t></a:r><a:br><a:rPr lang="ja-JP" sz="2000"/></a:br>'
     + '<a:r><a:rPr lang="ja-JP" sz="2000"/><a:t>見本</a:t></a:r></a:p>', '6) 改行を <a:br> にしない: ' + S.lfToLineBreaks_(x));
  const y = '<a:p><a:r><a:rPr lang="ja-JP"/><a:t>改行なし</a:t></a:r></a:p>';
  ck(S.lfToLineBreaks_(y) === y, '6) 改行の無い文字を変えた');
}

if (fails.length) {
  console.log('NG ' + fails.length + '件 / ' + checks + '件の検査');
  fails.slice(0, 40).forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('会社名の改行: 検査 ' + checks + ' 件 OK（組み合わせ ' + n + ' 通り）: 法人の種類・空白・語の切れ目で折る・行頭禁則・英単語と法人の種類の途中で折らない・'
  + '枠に入る・同じ会社は同じ切れ目・名簿の文字をそろえる・画面とサーバーが同じ計算・改行を <a:br> に');

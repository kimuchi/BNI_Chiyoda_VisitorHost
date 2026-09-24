// 画面(member_presen.html)の組版ロジックを、そのままNodeで動かして items.json を作る。
// canvasはNodeに無いので、全角1文字ぶん・半角0.5文字ぶんで測る簡易版に差し替える。
// 折り返しの「位置」までは本物と一致しないが、1行/2行の分岐や社名の分割は同じ経路を通る。
//
//   node tools/mp_plan.js <出力先ディレクトリ>

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const OUT = process.argv[2];
const html = fs.readFileSync(path.join(__dirname, '..', 'member_presen.html'), 'utf8');
const js = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(m => m[1]).join('\n');

function fakeCtx() {
  let size = 44, bold = false;
  return {
    set font(v) { const m = /(\d+(?:\.\d+)?)px/.exec(v); size = m ? parseFloat(m[1]) : 44; bold = /bold/.test(v); },
    get font() { return `${bold ? 'bold ' : ''}${size}px`; },
    measureText(t) {
      let w = 0;
      for (const ch of String(t)) w += ch.charCodeAt(0) < 128 ? 0.5 : 1.0;
      return { width: w * size };
    },
  };
}
const els = {};
const sandbox = {
  console,
  document: {
    createElement: (tag) => (tag === 'canvas' ? { getContext: () => fakeCtx() } : { set textContent(v) { this._t = v; }, get innerHTML() { return String(this._t == null ? '' : this._t); } }),
    // checked は既定でオン（画面の <input type="checkbox" checked> と同じ）。
    // long は「2分30秒プレゼンの方」の選択。検証のため1人選んだ状態にしておく。
    getElementById: (id) => (els[id] = els[id] || {
      style: {}, innerHTML: '', textContent: '', innerText: '', checked: true,
      value: (id === 'long' ? '佐藤　祐之' : '0'),
    }),
  },
  google: { script: { run: { withSuccessHandler: () => ({ withFailureHandler: () => ({ getMemberPresenContext() {} }) }) } } },
  window: {},
  alert: () => {},
};
sandbox.window = sandbox;
vm.createContext(sandbox);
vm.runInContext(js, sandbox, { filename: 'member_presen.html' });

// --- 検証用のメンバー（組版のあらゆる分岐を通す）---
const gk = (x) => x.replace(/[\s・･＆&と]/g, '');
const B = (block, order) => ({ gkey: gk(block), block, order, keys: [block], count: 0, known: true });
const blocks = [B('企業サポート', 1), B('研修・教育', 2), B('不動産関連', 3), B('美容・健康', 4)];
const M = (name, company, title, block, hasPhoto) =>
  ({ name, company, title, cat: block, blockKey: gk(block), hasPhoto: hasPhoto !== false });
const members = [
  // 企業サポート … 8名（＝扉ページが2枚になる）
  M('岡安　秀明', 'プルデンシャル生命保険㈱', '生命保険(法人)', '企業サポート'),          // ㈱ → 展開して2行
  M('田中　秀一', 'ABC総研', '中小企業診断士', '企業サポート'),                           // 短い → 1行44pt
  M('佐藤　祐之', '損害保険ジャパン株式会社', '損害保険', '企業サポート'),                 // 長い → 縮めて1行
  M('竹田　明日翔', 'ジブラルタ生命保険株式会社 新宿支社 第十一営業所', '生命保険(法人営業)', '企業サポート'), // 3行になっていた例
  M('合川　周平', '一般社団法人日本オフィスプロデュース協会連合会', 'オフィスプロデュース', '企業サポート'),
  M('分銅　雅一', '分銅税理士事務所', '税理士（事業承継と相続対策の専門家として）', '企業サポート'), // カテゴリーが長い
  M('岡本　翔太', '岡本法律事務所', '弁護士（企業法務）', '企業サポート'),
  M('写真　無子', 'テスト商会', 'テスト', '企業サポート', false),                          // 写真なし
  // 研修・教育 … 1名（＝NEXTが消える）
  M('研修　太郎', '研修ラボ合同会社', '企業研修', '研修・教育'),
  // 不動産関連 … 2名
  M('不動　産一', '不動産一番館', '売買仲介', '不動産関連'),
  M('宅建　次郎', '有限会社宅建サポート', '賃貸管理', '不動産関連'),
  // 美容・健康 … 0名（＝スライドが作られない）
];
members.forEach((m) => { const b = blocks.find((x) => x.gkey === m.blockKey); if (b) b.count++; });
sandbox.ctx = { ok: true, blocks, members, rowsPerPage: 7, candidates: [], template: { registered: true } };
sandbox.startKey = blocks[0].gkey;

const items = sandbox.buildItems();
fs.mkdirSync(OUT, { recursive: true });
fs.writeFileSync(path.join(OUT, 'items.json'), JSON.stringify(items, null, 1));

console.log('スライド枚数:', items.length);
for (const it of items) {
  if (it.kind === 'overview') {
    console.log(`  [扉] ${it.block}  先頭=${it.photoName}  表${it.rows.length}行`);
  } else {
    console.log(`  [個] ${it.name}  会社=${JSON.stringify(it.companyLines)} ${it.companyPt || 44}pt`
      + ` 枠cy=${it.companyGeom.cy} 区分の上端=${it.categoryTop}`
      + `  区分=${JSON.stringify(it.categoryLines)} ${it.categoryPt || 44}pt${it.categoryTight ? ' 行間詰' : ''}`
      + `  次=${it.nextName || '(なし)'}`);
  }
}

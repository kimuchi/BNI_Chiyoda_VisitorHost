// 業種区分の巡回が、元の生成ツールと同じ並びになるかを確かめる。
// テンプレートは要らないので単体で走る。
//
//   node tools/mp_check_rotation.js

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const s = { console };
s.global = s;
vm.createContext(s);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'member_presen_srv.js'), 'utf8'), s);

// 業種区分マスタの初期値（巡回順つき）
const blocks = [['企業サポート', 1], ['研修・教育', 2], ['不動産関連', 3], ['建築＆住まい', 4],
                ['プロモーション', 5], ['暮らし・生活', 6], ['美容・健康', 7], ['飲食・エンタメ', 8]]
  .map(([b, o]) => ({ key: b, block: b, order: o }));

// 元ツールの計算を、そのまま素直に書いたもの
const ORIG = blocks.map(b => b.block);
function original(iso) {
  const w = Math.round((new Date(iso + 'T00:00:00') - new Date('2026-08-19T00:00:00')) / (7 * 86400000));
  const st = ((ORIG.indexOf('プロモーション') + w) % 8 + 8) % 8;
  return Array.from({ length: 8 }, (_, i) => ORIG[(st + i) % 8]);
}

let bad = 0, n = 0;
for (let k = -52; k <= 104; k++) {          // 基準日の前後2年ぶん
  const d = new Date('2026-08-19T00:00:00');
  d.setDate(d.getDate() + 7 * k);
  const iso = d.toISOString().slice(0, 10);
  const mine = s.mpOrderFor_(blocks, iso.replace(/-/g, '/'));
  n++;
  if (JSON.stringify(mine) !== JSON.stringify(original(iso))) {
    bad++;
    if (bad <= 5) console.log('  不一致 %s  こちら=%s  元ツール=%s', iso, mine[0], original(iso)[0]);
  }
}
if (bad) { console.log('NG: %d / %d 週が元ツールと違います', bad, n); process.exit(1); }
console.log('OK: 業種区分の巡回は %d 週ぶん元ツールと一致（基準 2026-08-19 = プロモーション）', n);

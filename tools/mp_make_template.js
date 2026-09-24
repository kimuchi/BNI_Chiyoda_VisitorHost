// テンプレートづくりの、スライド本文の書き換えだけを担当する。
// 本番と同じ ooxml.js を使うので、生成時とまったく同じやり方で文字を入れ替えられる。
//
//   node tools/mp_make_template.js <パーツのディレクトリ>

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const DIR = process.argv[2];
const sandbox = { console };
sandbox.global = sandbox;
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'ooxml.js'), 'utf8'), sandbox, { filename: 'ooxml.js' });
const F = sandbox;

// 会社名・カテゴリーの枠は「1行のとき」の位置に戻しておく
const COMPANY_DEFAULT = { x: 4830858, y: 2849608, cx: 7461517, cy: 769441 };
const CATEGORY = { x: 4830481, y: 3456634, cx: 7311519 };

function edit(file, fn) {
  const p = path.join(DIR, file);
  fs.writeFileSync(p, fn(fs.readFileSync(p, 'utf8')), 'utf8');
}

// --- 扉ページ ---
edit('ppt/slides/slide1.xml', (xml) => {
  xml = F.setParagraphsInShape_(xml, 11, ['業種区分']);
  xml = F.setParagraphsInShape_(xml, 31, ['氏名']);
  for (let r = 1; r <= 7; r++) {
    xml = F.setTableCellText_(xml, 6, r, 0, '専門分野');
    xml = F.setTableCellText_(xml, 6, r, 1, '氏名');
  }
  return xml;
});

// --- 個人ページ ---
edit('ppt/slides/slide2.xml', (xml) => {
  xml = F.setParagraphsInShape_(xml, 56, ['氏名']);
  xml = F.setParagraphsInShape_(xml, 2, ['会社名']);
  xml = F.setShapeGeomEmu_(xml, 2, COMPANY_DEFAULT);
  xml = F.setParagraphsInShape_(xml, 12, ['【カテゴリー】']);
  xml = F.setShapeGeomEmu_(xml, 12, CATEGORY);
  xml = F.setParagraphsInShape_(xml, 14, ['氏名']);

  // 30秒カウントダウンを、スライドが出たら自動で始まるようにする。
  // 元は「クリックで開始」（delay="indefinite"）だったため、
  // 何秒で次へ進めばよいかがPowerPointにも分からなかった。
  const before = xml;
  xml = xml.replace(
    '<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>',
    '<p:cond delay="0"/>');
  if (xml === before) throw new Error('カウントダウンの開始条件が見つかりません');

  // カウントダウンが終わったら自動で次のスライドへ（30秒＋余韻1秒）。
  // 元の値は 400（0.4秒）で、読み終わる前に進んでしまっていた。
  const n = (xml.match(/advTm="\d+"/g) || []).length;
  if (!n) throw new Error('advTm（自動で次へ進む時間）が見つかりません');
  xml = xml.replace(/advTm="\d+"/g, 'advTm="31000"');
  return xml;
});

console.log('スライドの文字と動きを書き換えました');

// 後半スライドの出力を、差し込み口（{{ }}）付きのテンプレートに変える。
//
//   node tools/make_meeting_second_template.js <パーツのディレクトリ>

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { textOf, tokenizeTwoPerson } = require('./lib_meeting_tokens.js');

const DIR = process.argv[2];
const sandbox = { console };
sandbox.global = sandbox;
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'ooxml.js'), 'utf8'), sandbox, { filename: 'ooxml.js' });
const F = sandbox;

const slideDir = path.join(DIR, 'ppt/slides');
const slides = fs.readdirSync(slideDir).filter((n) => /^slide\d+\.xml$/.test(n))
  .sort((a, b) => parseInt(a.match(/\d+/)[0], 10) - parseInt(b.match(/\d+/)[0], 10));
const read = (n) => fs.readFileSync(path.join(slideDir, n), 'utf8');
const write = (n, x) => fs.writeFileSync(path.join(slideDir, n), x, 'utf8');

// 見出しの文字でページを見分ける（スライド番号は当てにしない）
const PAGES = [
  { title: 'Words of Recommendation', prefix: '推薦のことば' },
  { title: '賞品の抽選', prefix: '抽選' },
];
const report = [];
for (const pg of PAGES) {
  const name = slides.find((n) => textOf(read(n)).indexOf(pg.title) >= 0);
  if (!name) throw new Error(`「${pg.title}」のページが見つかりません`);
  const r = tokenizeTwoPerson(F, read(name), pg.prefix);
  write(name, r.xml);
  r.report.forEach((x) => report.push(`  ${name} ${x}`));
}

console.log('差し込み口を入れました:');
report.forEach((r) => console.log(r));

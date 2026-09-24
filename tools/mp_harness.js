// メンバープレゼン生成の検証用ハーネス。
// 実際の ooxml.js / member_presen_srv.js を、GASのAPIだけ差し替えてNodeで動かす。
// 本番と同じコードを、実物のテンプレートに対して走らせるのが目的。
//
//   node tools/mp_harness.js <展開先> <出力先>

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const WORK = process.argv[2];
const OUT = process.argv[3];
const PARTS = path.join(WORK, 'parts');
const manifest = JSON.parse(fs.readFileSync(path.join(WORK, 'manifest.json'), 'utf8'));

// --- GASの Blob を最小限だけ真似る ---
function fileBlob(partPath, type) {
  let name = partPath;
  return {
    _src: path.join(PARTS, partPath),
    getDataAsString() { return fs.readFileSync(this._src, 'utf8'); },
    getBytes() { return Array.from(fs.readFileSync(this._src)).map(b => (b > 127 ? b - 256 : b)); },
    getContentType() { return type || ''; },
    setName(n) { name = n; return this; },
    getName() { return name; },
  };
}
function memBlob(content, type, n) {
  let name = n;
  return {
    _mem: content,
    getDataAsString() { return this._mem; },
    getBytes() { return Array.from(Buffer.from(this._mem)); },
    getContentType() { return type || ''; },
    setName(x) { name = x; return this; },
    getName() { return name; },
  };
}

// --- テスト用の写真（テンプレート内の実物を流用する）---
const PHOTOS = {
  '岡安秀明': 'ppt/media/image12.jpeg',
  '田中秀一': 'ppt/media/image16.jpeg',
  '佐藤祐之': 'ppt/media/image14.png',
};

const sandbox = {
  console,
  Utilities: { newBlob: (c, t, n) => memBlob(c, t, n) },
  DriveApp: { getFileById: (id) => ({ getBlob: () => fileBlob(id, id.endsWith('.png') ? 'image/png' : 'image/jpeg') }) },
  normName_: (s) => String(s == null ? '' : s).replace(/[\s　]/g, ''),
  findPhotoIdForName_: (name) => PHOTOS[String(name).replace(/[\s　]/g, '')] || '',
};
sandbox.global = sandbox;
vm.createContext(sandbox);
for (const f of ['ooxml.js', 'member_presen_srv.js']) {
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', f), 'utf8'), sandbox, { filename: f });
}

// --- テンプレートを map にする ---
const map = {};
for (const p of manifest.parts) map[p] = fileBlob(p);

// --- 入力（画面が作るのと同じ形）---
const items = JSON.parse(fs.readFileSync(path.join(WORK, 'items.json'), 'utf8'));
const info = sandbox.buildMemberPresenSlides_(map, items);

// --- 結果を書き出す（Python側でzipに固める）---
fs.rmSync(OUT, { recursive: true, force: true });
fs.mkdirSync(OUT, { recursive: true });
const plan = {};
for (const p of Object.keys(map)) {
  const b = map[p];
  if (b._mem !== undefined) {
    const dst = path.join(OUT, 'gen', p);
    fs.mkdirSync(path.dirname(dst), { recursive: true });
    fs.writeFileSync(dst, b._mem, 'utf8');
    plan[p] = { from: 'gen' };
  } else {
    plan[p] = { from: 'parts', src: path.relative(PARTS, b._src) };
  }
}
fs.writeFileSync(path.join(OUT, 'plan.json'), JSON.stringify({ plan, info }, null, 1));
console.log('生成:', JSON.stringify(info));
console.log('パーツ数:', Object.keys(map).length);

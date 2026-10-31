// スーパーサポーター（supporters_srv.js）と、その方の写真（assets.js）を、実データなしで確かめる。名前はすべて架空。
//
//   node tools/check_supporters.js
//
// 確かめること
//   ・登録・読み直し（空のお名前・同じお名前は入れない・「さん」を外す）。名簿のメンバーと同じお名前は止める
//   ・写真：アップロード前の照らし合わせで、スーパーサポーターのファイル名も合う。合わない写真の一覧に出さない。
//     写真の無いスーパーサポーターを知らせる

process.env.TZ = 'Asia/Tokyo';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeEnv } = require('./lib_sheet_fake');

const ROOT = path.join(__dirname, '..');
const fails = [];
let checks = 0;
const ck = (ok, msg) => { checks++; if (!ok) fails.push(msg); };
const J = (x) => JSON.stringify(x);

const env = makeEnv({ now: new Date(2026, 9, 31, 10, 0, 0) });
const F = Object.assign({}, env.globals);
vm.createContext(F);
for (const f of fs.readdirSync(ROOT).filter((x) => /\.js$/.test(x)).sort()) vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), F, { filename: f });
env.reset([], {});
F.getMemberMaster = () => ({ members: [{ name: '見本 一郎' }, { name: '試験 花子' }] });
let index = { '見本一郎': 'P1', '他所一子': 'P2', '名簿外': 'P3' };
F.findPhotoIdForName_ = (n) => index[F.photoKey_(n)] || '';

// ===== 1) 登録 =====
let r = F.saveSuperSupporters([{ name: '他所 一子さん', chapter: '見本チャプター', company: '他所印刷', category: '印刷' },
                               { name: '' }, { name: '他所　一子', chapter: '重複' }, { name: '遠方 二郎', category: '保険' }]);
ck(r.ok && J(r.supporters.map((s) => [s.name, s.chapter, s.hasPhoto])) === J([['他所 一子', '見本チャプター', true], ['遠方 二郎', '', false]]),
   '1) 登録: ' + J(r));
ck(J(F.getSuperSupporters().supporters.map((s) => s.name)) === J(['他所 一子', '遠方 二郎']), '1) 読み直し');
r = F.saveSuperSupporters([{ name: '試験花子' }]);
ck(!r.ok && /メンバー名簿に同じお名前/.test(r.message) && F.getSuperSupporters().supporters.length === 2, '1) 名簿のメンバーと同じお名前を止めない: ' + J(r));

// ===== 2) 写真 =====
{
  const m = F.matchPhotoFileNames(['他所一子_261031.jpg', '遠方 二郎.png', 'IMG_1.jpg']);
  ck(J(m.items.map((x) => x.member)) === J(['他所 一子', '遠方 二郎', '']) && m.members.indexOf('他所 一子') >= 0, '2) 照らし合わせ: ' + J(m));
  env.reset([['写真索引', true, [['ファイル名', '氏名(正規化)', 'ファイルID', '版'], ['見本 一郎.jpg', '見本一郎', 'P1', '1'],
                                  ['他所 一子.jpg', '他所一子', 'P2', '1'], ['名簿外.jpg', '名簿外', 'P3', '1']]]], { BNI_SUPER_SUPPORTERS: J([{ name: '他所 一子' }, { name: '遠方 二郎' }]) });
  F.getAssetFolder_ = () => ({ getUrl: () => 'u' });
  F.getAssetSettings = () => ({});
  const ov = F.getPhotoOverview();
  ck(J(ov.strays) === J(['名簿外.jpg']) && J(ov.supportersNoPhoto) === J(['遠方 二郎']), '2) 写真の状況: ' + J(ov));
}

if (fails.length) { console.error('NG ' + fails.length + '/' + checks + '\n' + fails.join('\n')); process.exit(1); }
console.log('OK ' + checks + ' checks');
